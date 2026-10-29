import marimo

__generated_with = "0.25.2"
app = marimo.App(width="medium", app_title="Radar de Leads · AvIA Analytics")


@app.cell
def _():
    # =====================================================================
    # BIBLIOTECAS
    # =====================================================================
    import base64
    import json
    import os
    import re
    import tempfile
    import time
    import uuid
    from functools import lru_cache
    from pathlib import Path
    from types import SimpleNamespace
    from urllib.parse import urlencode

    import marimo as mo
    import msal
    import pandas as pd
    import requests
    return (
        Path,
        SimpleNamespace,
        base64,
        json,
        lru_cache,
        mo,
        msal,
        os,
        pd,
        re,
        requests,
        tempfile,
        time,
        urlencode,
        uuid,
    )


@app.cell
def _(Path, SimpleNamespace, os, uuid):
    # =====================================================================
    # CONFIGURAÇÃO — credenciais vêm do .env (na pasta radar_leads/ ou na pasta acima)
    # =====================================================================
    PASTA = Path(__file__).resolve().parent            # pasta radar_leads/
    _env_arquivo = next((p for p in (PASTA / ".env", PASTA.parent / ".env") if p.exists()), PASTA / ".env")

    def _ler_env(arquivo):
        """Lê o .env: aceita espaços, aspas e arquivos UTF-8/UTF-16."""
        if not arquivo.exists():
            return {}
        bruto = arquivo.read_bytes()
        for codificacao in ("utf-8-sig", "utf-16", "latin-1"):
            try:
                texto = bruto.decode(codificacao)
                break
            except UnicodeDecodeError:
                continue
        valores = {}
        for linha in texto.splitlines():
            linha = linha.strip()
            if linha and not linha.startswith("#") and "=" in linha:
                chave, valor = linha.split("=", 1)
                valores[chave.strip()] = valor.strip().strip('"').strip("'").strip()
        return valores

    _env = _ler_env(_env_arquivo)

    def _valor(nome, padrao=""):
        return (os.getenv(nome) or _env.get(nome) or padrao).strip()

    def _problemas():
        if not _env_arquivo.exists():
            return [f"Arquivo .env não encontrado. Crie em: {_env_arquivo}"]
        erros = []
        for nome in ("AVIA_TENANT_ID", "AVIA_CLIENT_ID"):
            try:
                uuid.UUID(_valor(nome))
            except ValueError:
                erros.append(f"{nome}: valor ausente ou inválido em {_env_arquivo} (lido: {_valor(nome)!r}).")
        if not _valor("AVIA_CLIENT_SECRET"):
            erros.append(f"AVIA_CLIENT_SECRET: valor ausente em {_env_arquivo}.")
        return erros

    config = SimpleNamespace(
        BASES=PASTA / "bases",
        # Microsoft Entra ID
        TENANT_ID=_valor("AVIA_TENANT_ID"),
        CLIENT_ID=_valor("AVIA_CLIENT_ID"),
        CLIENT_SECRET=_valor("AVIA_CLIENT_SECRET"),
        REDIRECT_URI=_valor("RADAR_REDIRECT_URI", "http://localhost:8000/radar-leads/"),
        # Power BI: relatório de origem
        WORKSPACE_ID=_valor("RADAR_WORKSPACE_ID", "fcab744b-4f01-4f26-a3bb-e382505d162b"),
        REPORT_ID=_valor("RADAR_REPORT_ID", "ad1c7f8c-e011-4810-9845-0c0d999256a1"),
        PAGE_NAME=_valor("RADAR_PAGE_NAME", "ReportSection5418c63652acbe57a240"),
        # Colunas das bases e campos do modelo
        COL_UF="dm_Localizacao[UF]",
        COL_MUN="dm_Localizacao[NM_MUN]",
        CAMPO_UF="dm_Localizacao/UF",
        CAMPO_MUN="dm_Localizacao/NM_MUN",
        # (chave, rótulo, coluna no CSV, campo no Power BI) — na ordem da hierarquia
        CAMPOS_CNAE=[
            ("setor", "Grande setor", "dm_Cnae[NM_GRANDES_SETORES]", "dm_Cnae/NM_GRANDES_SETORES"),
            ("secao", "Seção", "dm_Cnae[COD_CNAE_SECAO]", "dm_Cnae/COD_CNAE_SECAO"),
            ("divisao", "Divisão", "dm_Cnae[NM_DIVISAO]", "dm_Cnae/NM_DIVISAO"),
            ("grupo", "Grupo", "dm_Cnae[NM_GRUPO]", "dm_Cnae/NM_GRUPO"),
            ("classe", "Classe", "dm_Cnae[NM_CLASSE]", "dm_Cnae/NM_CLASSE"),
            ("subclasse", "Subclasse", "dm_Cnae[NM_SUBCLASSE]", "dm_Cnae/NM_SUBCLASSE"),
        ],
        problemas=_problemas,
    )
    CNAE = {_k: (_rotulo, _col, _campo) for _k, _rotulo, _col, _campo in config.CAMPOS_CNAE}
    return CNAE, PASTA, config


@app.cell
def _(Path, SimpleNamespace, config, json, msal, re, tempfile, time):
    # =====================================================================
    # AUTENTICAÇÃO MICROSOFT (MSAL)
    # 1. link_login() gera o endereço da Microsoft
    # 2. a Microsoft volta para REDIRECT_URI com ?code=...&state=...
    # 3. concluir_login() troca o código por tokens
    # Os logins pendentes ficam em arquivos temporários (a página recarrega no retorno).
    # =====================================================================
    _ESCOPOS = [
        "https://analysis.windows.net/powerbi/api/Content.Create",
        "https://analysis.windows.net/powerbi/api/Report.ReadWrite.All",
        "https://analysis.windows.net/powerbi/api/Dataset.ReadWrite.All",
    ]
    _PENDENTES = Path(tempfile.gettempdir()) / "avia_radar_login"

    def _cliente_msal(cache=None):
        return msal.ConfidentialClientApplication(
            config.CLIENT_ID, authority=f"https://login.microsoftonline.com/{config.TENANT_ID}",
            client_credential=config.CLIENT_SECRET, token_cache=cache)

    class SessaoMicrosoft:
        """Usuário conectado: guarda o cache de tokens e renova silenciosamente."""

        def __init__(self, cache, resultado):
            self._cache = cache
            claims = resultado.get("id_token_claims", {}) or {}
            self.nome = claims.get("name") or claims.get("preferred_username") or "Conta Microsoft"
            self._token = resultado.get("access_token", "")
            self._expira = time.time() + int(resultado.get("expires_in", 3600)) - 60

        def obter_token(self, forcar=False):
            if not forcar and self._token and self._expira > time.time():
                return self._token
            cliente = _cliente_msal(self._cache)
            contas = cliente.get_accounts()
            try:
                r = cliente.acquire_token_silent(_ESCOPOS, account=contas[0], force_refresh=forcar) if contas else None
            except Exception:  # noqa: BLE001
                r = None
            if r and "access_token" in r:
                self._token, self._expira = r["access_token"], time.time() + int(r.get("expires_in", 3600)) - 60
                return self._token
            return None

    def _diagnosticar(resultado):
        orientacoes = {
            "90002": "AVIA_TENANT_ID: diretório não encontrado.",
            "700016": "AVIA_CLIENT_ID / AVIA_TENANT_ID: aplicativo não encontrado nesse diretório.",
            "7000215": "AVIA_CLIENT_SECRET: segredo inválido. Use o campo Valor, não o ID do segredo.",
            "7000222": "AVIA_CLIENT_SECRET: segredo expirado. Gere outro no Entra.",
            "65001": "Permissões: falta consentimento para as permissões do Power BI.",
            "50011": "URI de redirecionamento: cadastre no Entra exatamente o valor de RADAR_REDIRECT_URI.",
        }
        descricao = str(resultado.get("error_description", "")).replace(config.CLIENT_SECRET or "\0", "[OCULTO]")
        codigos = {str(c) for c in (resultado.get("error_codes") or [])} | set(re.findall(r"AADSTS(\d+)", descricao))
        dicas = [orientacoes[c] for c in sorted(codigos) if c in orientacoes]
        return " ".join(dicas or ["Não foi possível concluir o login."]) + f"\n\nDetalhe da Microsoft: {descricao[:1500]}"

    def _link_login():
        flow = _cliente_msal().initiate_auth_code_flow(_ESCOPOS, redirect_uri=config.REDIRECT_URI, response_mode="query")
        if "auth_uri" not in flow:
            raise RuntimeError(_diagnosticar(flow))
        _PENDENTES.mkdir(exist_ok=True)
        for antigo in _PENDENTES.glob("*.json"):
            if time.time() - antigo.stat().st_mtime > 600:
                antigo.unlink(missing_ok=True)
        (_PENDENTES / f"{re.sub(r'[^A-Za-z0-9_-]', '', flow['state'])}.json").write_text(json.dumps(flow))
        return flow["auth_uri"]

    def _retorno_de_login(parametros):
        return "state" in parametros and ("code" in parametros or "error" in parametros)

    def _concluir_login(parametros):
        """Devolve (SessaoMicrosoft, None) ou (None, mensagem de erro)."""
        params = {k: (v[0] if isinstance(v, list) else v) for k, v in parametros.items()}
        arquivo = _PENDENTES / f"{re.sub(r'[^A-Za-z0-9_-]', '', params.get('state', ''))}.json"
        if not arquivo.exists() or time.time() - arquivo.stat().st_mtime > 600:
            return None, "O link de login expirou ou já foi usado. Clique em Entrar com Microsoft novamente."
        flow = json.loads(arquivo.read_text())
        arquivo.unlink(missing_ok=True)
        cache = msal.SerializableTokenCache()
        try:
            resultado = _cliente_msal(cache).acquire_token_by_auth_code_flow(flow, params)
        except ValueError:
            return None, "O retorno da Microsoft não passou pela validação de segurança. Tente entrar novamente."
        if "access_token" not in resultado:
            return None, _diagnosticar(resultado)
        return SessaoMicrosoft(cache, resultado), None

    autenticacao = SimpleNamespace(link_login=_link_login, retorno_de_login=_retorno_de_login,
                                   concluir_login=_concluir_login)
    return SessaoMicrosoft, autenticacao


@app.cell
def _(SimpleNamespace, base64, json, requests, time, urlencode):
    # =====================================================================
    # POWER BI — REST API, clonagem, filtros e HTML de incorporação
    # =====================================================================
    _API = "https://api.powerbi.com/v1.0/myorg"
    _SDK = "https://cdn.jsdelivr.net/npm/powerbi-client@2.23.1/dist/powerbi.min.js"

    class ErroPowerBI(RuntimeError):
        """Mensagem para o usuário; `detalhe` guarda o diagnóstico técnico (sem token)."""

        def __init__(self, mensagem, detalhe=""):
            super().__init__(mensagem)
            self.detalhe = detalhe

    def _resumo_token(token):
        try:
            parte = token.split(".")[1]
            d = json.loads(base64.urlsafe_b64decode(parte + "=" * (-len(parte) % 4)))
            return {"aud": d.get("aud"), "scp": d.get("scp", ""),
                    "expirado": bool(d.get("exp") and float(d["exp"]) <= time.time())}
        except (IndexError, ValueError, TypeError, AttributeError):
            return {}

    class ClientePowerBI:
        def __init__(self, obter_token):
            self._obter_token = obter_token

        @property
        def token(self):
            token = self._obter_token()
            if not token:
                raise ErroPowerBI("Sua sessão Microsoft expirou. Entre novamente.")
            return token

        def chamar(self, metodo, caminho, **kwargs):
            token = self.token
            resp = requests.request(metodo, _API + caminho, timeout=45, headers={"Authorization": "Bearer " + token}, **kwargs)
            if resp.status_code == 401 and (novo := self._obter_token(forcar=True)):   # tenta uma vez renovado
                token = novo
                resp = requests.request(metodo, _API + caminho, timeout=45, headers={"Authorization": "Bearer " + token}, **kwargs)
            if resp.ok:
                return resp.json() if resp.content else {}
            detalhe = f"HTTP {resp.status_code} em {metodo} {caminho}\n" + resp.text[:1200].replace(token, "[TOKEN OCULTO]")
            if resp.status_code in (401, 403):
                meta = _resumo_token(token)
                detalhe += f"\naud={meta.get('aud')} scp={meta.get('scp')} expirado={meta.get('expirado')}"
                if meta.get("expirado"):
                    raise ErroPowerBI("Sua sessão Microsoft expirou. Entre novamente.", detalhe)
                raise ErroPowerBI("Sua conta não tem autorização para criar conteúdo neste workspace. Solicite acesso "
                                  "de Colaborador (ou superior) e permissão no modelo semântico.", detalhe)
            raise ErroPowerBI(f"Não foi possível concluir a operação no Power BI (HTTP {resp.status_code}).", detalhe)

    def _detalhes(cliente, workspace, relatorio):
        return cliente.chamar("GET", f"/groups/{workspace}/reports/{relatorio}")

    def _clonar(cliente, workspace, original, nome):
        """Cria uma cópia no mesmo workspace, ligada ao mesmo modelo semântico."""
        escopos = set(_resumo_token(cliente.token).get("scp", "").split())
        if escopos and "Content.Create" not in escopos:
            raise ErroPowerBI("O token não contém a permissão Content.Create, exigida para criar a cópia. "
                              "Peça a permissão e entre novamente.")
        dataset = _detalhes(cliente, workspace, original).get("datasetId")
        if not dataset:
            raise ErroPowerBI("O relatório de origem não possui modelo semântico.")
        existentes = cliente.chamar("GET", f"/groups/{workspace}/reports").get("value", [])
        if any(r.get("name", "").casefold() == nome.casefold() for r in existentes):
            raise ErroPowerBI("Já existe um relatório com esse nome no workspace. Escolha outro nome.")
        novo = cliente.chamar("POST", f"/groups/{workspace}/reports/{original}/Clone",
                              json={"name": nome, "targetWorkspaceId": workspace, "targetModelId": dataset})
        return {**novo, "datasetIdEsperado": dataset}

    # Filtro da URL do Service e filtro do JavaScript SDK são formatos diferentes.
    def _filtro_url(campo, valores):
        v = [str(x).replace("'", "''") for x in valores]
        return f"{campo} eq '{v[0]}'" if len(v) == 1 else f"{campo} in (" + ",".join(f"'{x}'" for x in v) + ")"

    def _filtros_embedded(campos):
        return [{"$schema": "http://powerbi.com/product/schema#basic",
                 "target": {"table": campo.split("/")[0], "column": campo.split("/", 1)[1]},
                 "operator": "In", "values": list(valores), "filterType": 1}
                for campo, valores in campos if valores]

    def _url_service(workspace, relatorio, pagina, tenant, campos):
        filtro = " and ".join(_filtro_url(c, v) for c, v in campos if v)
        return (f"https://app.powerbi.com/groups/{workspace}/reports/{relatorio}/{pagina}?"
                + urlencode({"experience": "power-bi", "filter": filtro, "ctid": tenant}))

    def _config_embedded(workspace, relatorio, pagina, lista_filtros, url=None, painel_filtros=False, **extra):
        """Configuração para o Power BI JavaScript SDK (sem token)."""
        return {"type": "report", "id": relatorio, "groupId": workspace,
                "embedUrl": url or "https://app.powerbi.com/reportEmbed?" + urlencode({"reportId": relatorio, "groupId": workspace}),
                "pageName": pagina, "filters": lista_filtros,
                "settings": {"panes": {"filters": {"visible": painel_filtros}, "pageNavigation": {"visible": True}}},
                **extra}

    def _editor_html(cfg, token, modo="previa"):
        """modo "previa" (leitura) ou "salvar" (aplica e salva os filtros na cópia, em modo edição). Exibir com mo.iframe."""
        payload = json.dumps({**cfg, "accessToken": token, "modo": modo}, ensure_ascii=False).replace("<", "\\u003c")
        return f"""<!doctype html><html lang="pt-BR"><head><meta charset="utf-8"><style>
    body{{font:14px/1.5 "Segoe UI",system-ui,sans-serif;color:#1B2B45;margin:0;background:transparent}}
    #status{{padding:12px 14px;border-left:4px solid #154388;background:#E8EEF7;border-radius:8px;margin-bottom:10px}}
    #status.erro{{border-left-color:#B42318;background:#FEF3F2}} #status.ok{{border-left-color:#15803D;background:#F0FDF4}}
    #report{{height:680px;border-radius:10px;overflow:hidden;border:1px solid #C9D6EA}} #report iframe{{border:0}}
    </style></head><body>
    <div id="status" role="status">Carregando o relatório…</div><div id="report"></div>
    <script src="{_SDK}"></script>
    <script>
    const cfg = {payload};
    const models = window['powerbi-client'].models;
    const status = document.getElementById('status');
    const leitura = cfg.modo === 'previa';
    const msg = (t, c) => {{ status.textContent = t; status.className = c || ''; }};
    // Ao salvar, o relatório abre SEM filtros: assim o setFilters abaixo é uma alteração real
    // e o save() tem o que gravar. Se abrisse já filtrado, não haveria mudança a salvar.
    const report = powerbi.embed(document.getElementById('report'), {{
      ...cfg, filters: leitura ? cfg.filters : [], tokenType: models.TokenType.Aad,
      permissions: leitura ? models.Permissions.Read : models.Permissions.All,
      viewMode: leitura ? models.ViewMode.View : models.ViewMode.Edit }});
    function falhou(e) {{
      const t = String((e?.detail?.errorCode || '') + ' ' + (e?.detail?.message || e?.message || ''));
      msg(/unauthoriz|forbidden|permission|accessdenied|401|403/i.test(t)
        ? 'Sua conta não tem autorização para este conteúdo. Solicite acesso de Colaborador e permissão no modelo semântico.'
        : (leitura ? 'Não foi possível carregar o relatório.' : 'Não foi possível salvar os filtros. A cópia já existe; abra-a e salve manualmente.'), 'erro');
    }}
    report.on('error', falhou);
    let aguardando = null;
    report.on('saved', () => {{
      if (!aguardando) return;
      clearTimeout(aguardando);
      aguardando = null;
      msg('Filtros salvos no novo relatório. Pode abri-lo no workspace.', 'ok');
    }});
    report.on('loaded', async () => {{
      report.off('loaded');   // aplica e salva uma vez só
      try {{
        // Troca só os campos escolhidos e preserva os filtros de negócio da origem.
        const atuais = await report.getFilters();
        const chave = f => JSON.stringify(f.target);
        const escolhidos = new Set(cfg.filters.map(chave));
        await report.setFilters(atuais.filter(f => !escolhidos.has(chave(f))).concat(cfg.filters));
        if (cfg.pageName) await report.setPage(cfg.pageName);
        if (leitura) {{ msg('Prévia com os filtros selecionados. Confira antes de criar sua cópia.'); return; }}
        msg('Salvando os filtros na sua cópia… Não feche esta página nem abra o relatório ainda.');
        // save() só envia o pedido; a confirmação de verdade é o evento 'saved'.
        aguardando = setTimeout(() => {{
          aguardando = null;
          msg('O Power BI não confirmou o salvamento. Abra a cópia no workspace, confira os filtros e salve manualmente.', 'erro');
        }}, 60000);
        await report.save();
      }} catch (e) {{
        if (aguardando) {{ clearTimeout(aguardando); aguardando = null; }}
        falhou(e);
      }}
    }});
    </script></body></html>"""

    powerbi = SimpleNamespace(ErroPowerBI=ErroPowerBI, Cliente=ClientePowerBI, detalhes=_detalhes, clonar=_clonar,
                              filtros_embedded=_filtros_embedded, url_service=_url_service,
                              config_embedded=_config_embedded, editor_html=_editor_html)
    return ClientePowerBI, ErroPowerBI, powerbi


@app.cell
def _(SimpleNamespace, config, lru_cache, pd):
    # =====================================================================
    # DADOS — bases CSV que alimentam os filtros
    # =====================================================================
    @lru_cache(maxsize=1)
    def _carregar():
        def ler(nome):
            df = pd.read_csv(config.BASES / nome, dtype=str, keep_default_na=False, encoding="utf-8")
            return df.apply(lambda c: c.str.strip().replace("", pd.NA))
        localidade, cnae = ler("localidade.csv"), ler("cnae.csv")
        faltando = [c for c in (config.COL_UF, config.COL_MUN) if c not in localidade.columns]
        faltando += [c for _, _, c, _ in config.CAMPOS_CNAE if c not in cnae.columns]
        if faltando:
            raise ValueError("Colunas não encontradas nas bases: " + ", ".join(faltando))
        return localidade, cnae

    dados = SimpleNamespace(
        carregar=_carregar,
        opcoes=lambda df, coluna: sorted(df[coluna].dropna().astype(str).unique().tolist()),
        filtrar=lambda df, coluna, valores: df[df[coluna].isin(valores)] if valores else df,
    )
    return (dados,)


@app.cell
def _(json, mo):
    # =====================================================================
    # VISUAL — mesma identidade da home: branco + azul #154388
    # Fontes do sistema e ícones em SVG desenhados aqui (nada carregado de fora).
    # =====================================================================
    _TRACOS = {
        "voltar": '<path d="M19 12H5M11 6l-6 6 6 6"/>',
        "usuario": '<circle cx="12" cy="8" r="4"/><path d="M4 21c0-4.4 3.6-8 8-8s8 3.6 8 8"/>',
        "entrar": '<path d="M10 17l5-5-5-5"/><path d="M15 12H3"/><path d="M14 3h5a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2h-5"/>',
        "externo": '<path d="M14 4h6v6"/><path d="M20 4l-9 9"/><path d="M18 14v5a1 1 0 0 1-1 1H5'
                   'a1 1 0 0 1-1-1V7a1 1 0 0 1 1-1h5"/>',
        "ok": '<circle cx="12" cy="12" r="9"/><path d="M8 12.5l2.8 2.8L16 10"/>',
    }

    def icone(nome, tam=18):
        return (f'<svg width="{tam}" height="{tam}" viewBox="0 0 24 24" fill="none" '
                f'stroke="currentColor" stroke-width="1.9" stroke-linecap="round" '
                f'stroke-linejoin="round" aria-hidden="true">{_TRACOS[nome]}</svg>')

    LOGO = (
        '<svg width="34" height="34" viewBox="0 0 40 40" aria-hidden="true">'
        '<rect width="40" height="40" rx="10" fill="#FFFFFF"/>'
        '<path d="M8 31 L17 13 L22 22 L32 8" fill="none" stroke="#154388" stroke-width="3.4" '
        'stroke-linecap="round" stroke-linejoin="round"/>'
        '<path d="M13 24 H24" stroke="#154388" stroke-opacity=".4" stroke-width="2.6" stroke-linecap="round"/>'
        '<circle cx="32" cy="8" r="2.6" fill="#154388"/></svg>'
    )
    LINHA = (
        '<svg class="avia-linha" viewBox="0 0 1200 140" preserveAspectRatio="none" aria-hidden="true">'
        '<path d="M0 132 L180 112 L300 120 L460 80 L590 92 L760 50 L900 62 L1040 20 L1200 6 V140 H0 Z" '
        'fill="rgba(255,255,255,.07)"/>'
        '<path d="M0 132 L180 112 L300 120 L460 80 L590 92 L760 50 L900 62 L1040 20 L1200 6" '
        'fill="none" stroke="#FFFFFF" stroke-opacity=".75" stroke-width="2.5" '
        'vector-effect="non-scaling-stroke"/></svg>'
    )

    def pilha(itens, gap=1):
        """vstack que ignora itens vazios (None)."""
        return mo.vstack([_i for _i in itens if _i is not None], gap=gap)

    def json_fmt(obj):
        return json.dumps(obj, ensure_ascii=False, indent=2)

    def secao(numero, titulo, texto, itens):
        """Cartão de uma etapa: número, título, explicação e o conteúdo (elementos marimo)."""
        return mo.Html(
            f'<section class="avia-secao"><header class="avia-secao-topo">'
            f'<span class="avia-num">{numero}</span><div><h2>{titulo}</h2><p>{texto}</p></div></header>'
            f'<div class="avia-secao-corpo">{pilha(itens)}</div></section>')

    def subtitulo(texto):
        return mo.Html(f'<h3 class="avia-subtitulo">{texto}</h3>')

    def nota(texto):
        return mo.Html(f'<p class="avia-nota">{texto}</p>')

    def indicadores(pares):
        return mo.Html('<div class="avia-indicadores">' + "".join(
            f'<div class="avia-indicador"><strong>{valor}</strong><span>{rotulo}</span></div>'
            for rotulo, valor in pares) + "</div>")

    def erro_pbi(exc, prefixo=""):
        _partes = [mo.md(f"**{prefixo}{exc}**")]
        if getattr(exc, "detalhe", ""):
            _partes.append(mo.accordion({"Detalhes técnicos": mo.md(f"```text\n{exc.detalhe}\n```")}))
        return mo.callout(mo.vstack(_partes), kind="danger")

    _ESTILO = """<style>
    :root {
      --avia-azul: #154388;
      --avia-azul-escuro: #0E2F61;
      --avia-azul-claro: #E8EEF7;
      --avia-azul-borda: #C9D6EA;
      --avia-texto: #1B2B45;
      --avia-suave: #5A6B86;
      --avia-fonte: "Segoe UI", system-ui, -apple-system, Roboto, "Helvetica Neue", Arial, sans-serif;
      --marimo-heading-font: var(--avia-fonte);
      --marimo-text-font: var(--avia-fonte);
    }

    /* ---------- faixa azul do topo ---------- */
    .avia-faixa { position: relative; overflow: hidden; color: #fff; border-radius: 16px;
      font-family: var(--avia-fonte);
      background: linear-gradient(135deg, var(--avia-azul) 0%, var(--avia-azul-escuro) 100%);
      padding: 24px 32px 96px; }
    .avia-faixa::before { content: ""; position: absolute; inset: 0;
      background-image: radial-gradient(rgba(255,255,255,.09) 1px, transparent 1px);
      background-size: 24px 24px; pointer-events: none; }
    .avia-topo { position: relative; display: flex; justify-content: space-between; align-items: center; gap: 12px; }
    .avia-marca { display: flex; align-items: center; gap: 10px; color: #fff !important;
      text-decoration: none !important; font: 600 17px var(--avia-fonte); }
    .avia-marca b { font-weight: 400; opacity: .8; }
    .avia-voltar { display: inline-flex; align-items: center; gap: 6px; color: #D4DEEF !important;
      font-size: 14px; text-decoration: none !important; padding: 6px 12px; border-radius: 999px;
      border: 1px solid rgba(255,255,255,.25); transition: background .15s; }
    .avia-voltar:hover { background: rgba(255,255,255,.12); color: #fff !important; }
    .avia-faixa h1 { position: relative; font: 700 clamp(28px, 4vw, 40px)/1.1 var(--avia-fonte) !important;
      letter-spacing: -1px; color: #fff !important; margin: 36px 0 10px !important; border: 0 !important; }
    .avia-faixa p { position: relative; color: #D4DEEF; font-size: 16.5px; max-width: 58ch; margin: 0; }
    .avia-linha { position: absolute; left: 0; bottom: 0; width: 100%; height: 96px; }

    /* ---------- conta Microsoft ---------- */
    .avia-conta { display: flex; align-items: center; justify-content: space-between; gap: 16px;
      flex-wrap: wrap; padding: 16px 20px; border: 1px solid var(--avia-azul-borda); border-radius: 14px;
      font-family: var(--avia-fonte); color: var(--avia-texto); background: #fff; }
    .avia-conta-info { display: flex; align-items: center; gap: 12px; }
    .avia-conta-info p { margin: 2px 0 0; color: var(--avia-suave); font-size: 14px; }
    .avia-avatar { width: 40px; height: 40px; border-radius: 50%; display: grid; place-items: center;
      background: var(--avia-azul-claro); color: var(--avia-azul); flex: 0 0 auto; }

    /* ---------- etapas ---------- */
    .avia-secao { border: 1px solid var(--avia-azul-borda); border-radius: 16px; background: #fff;
      font-family: var(--avia-fonte); color: var(--avia-texto); overflow: hidden; }
    .avia-secao-topo { display: flex; gap: 14px; align-items: flex-start; padding: 18px 22px;
      background: var(--avia-azul-claro); border-bottom: 1px solid var(--avia-azul-borda); }
    .avia-num { flex: 0 0 34px; height: 34px; border-radius: 50%; background: var(--avia-azul);
      color: #fff; display: grid; place-items: center; font: 700 15px var(--avia-fonte); }
    .avia-secao-topo h2 { font: 600 20px var(--avia-fonte) !important; margin: 3px 0 2px !important;
      color: var(--avia-azul) !important; border: 0 !important; padding: 0 !important; }
    .avia-secao-topo p { margin: 0; color: var(--avia-suave); font-size: 14.5px; max-width: 70ch; }
    .avia-secao-corpo { padding: 20px 22px 22px; }

    .avia-subtitulo { font: 600 15px var(--avia-fonte) !important; color: var(--avia-azul) !important;
      margin: 4px 0 0 !important; border: 0 !important; }
    .avia-nota { margin: 0; color: var(--avia-suave); font: italic 14.5px var(--avia-fonte); }

    .avia-indicadores { display: grid; grid-template-columns: repeat(auto-fit, minmax(150px, 1fr)); gap: 12px; }
    .avia-indicador { border: 1px solid var(--avia-azul-borda); border-radius: 12px; padding: 14px 16px;
      display: flex; flex-direction: column; gap: 2px; font-family: var(--avia-fonte); }
    .avia-indicador strong { font-size: 28px; line-height: 1.1; color: var(--avia-azul); }
    .avia-indicador span { color: var(--avia-suave); font-size: 13.5px; }

    /* ---------- botões e links ---------- */
    .avia-botao { display: inline-flex; align-items: center; justify-content: center; gap: 8px;
      white-space: nowrap; border-radius: 10px; padding: 11px 18px; font: 600 15px var(--avia-fonte);
      text-decoration: none !important; background: var(--avia-azul); color: #fff !important;
      transition: background .15s; }
    .avia-botao:hover { background: var(--avia-azul-escuro); }
    .avia-botao:focus-visible, .avia-voltar:focus-visible { outline: 3px solid var(--avia-azul-borda); outline-offset: 3px; }
    a.avia-link { display: inline-flex; align-items: center; gap: 6px; color: var(--avia-azul) !important;
      font-weight: 600; text-decoration: none !important; border-bottom: 2px solid var(--avia-azul-borda); }
    a.avia-link:hover { border-bottom-color: var(--avia-azul); }

    @media (max-width: 720px) {
      .avia-faixa { padding: 18px 18px 80px; }
      .avia-secao-topo, .avia-secao-corpo { padding-left: 16px; padding-right: 16px; }
    }
    </style>"""

    # Traduz os textos fixos em inglês dos componentes do marimo (ex.: "Select..." do seletor múltiplo).
    # O marimo não executa <script> em mo.Html; por isso o código entra na página por um mo.iframe invisível.
    _TRADUTOR = """
    (() => {
      if (window.__aviaPt) return;
      window.__aviaPt = true;
      const textos = {
        "Select...": "Selecione...", "Nothing found.": "Nenhum resultado encontrado.",
        "Select all": "Selecionar todos", "Deselect all": "Desmarcar todos", "Clear": "Limpar",
        "Clear all": "Limpar tudo", "Search...": "Buscar...", "Download": "Baixar",
        "Copy to clipboard": "Copiar", "Copied!": "Copiado!", "Copy": "Copiar",
      };
      const padroes = [
        [/^Select (\\d+) matching$/, "Selecionar $1 correspondentes"],
        [/^Deselect (\\d+) matching$/, "Desmarcar $1 correspondentes"],
        [/^\\+(\\d+) more$/, "+$1 outros"],
      ];
      const traduzir = (t) => {
        const s = (t || "").trim();
        if (!s) return null;
        if (textos[s]) return t.replace(s, textos[s]);
        for (const [re, novo] of padroes) if (re.test(s)) return t.replace(s, s.replace(re, novo));
        return null;
      };
      const raizes = new WeakSet();
      function varrer(raiz) {
        const w = document.createTreeWalker(raiz, NodeFilter.SHOW_TEXT);
        for (let n; (n = w.nextNode());) {
          const t = traduzir(n.nodeValue);
          if (t !== null && t !== n.nodeValue) n.nodeValue = t;
        }
        raiz.querySelectorAll("*").forEach((el) => {
          for (const attr of ["placeholder", "aria-label", "title"]) {
            const t = traduzir(el.getAttribute(attr));
            if (t !== null) el.setAttribute(attr, t);
          }
          if (el.shadowRoot) observar(el.shadowRoot);
        });
      }
      function observar(raiz) {
        if (raizes.has(raiz)) return;
        raizes.add(raiz);
        let agendado = false;
        const rodar = () => { agendado = false; varrer(raiz); };
        varrer(raiz);
        new MutationObserver(() => { if (!agendado) { agendado = true; setTimeout(rodar, 30); } })
          .observe(raiz, { childList: true, subtree: true, characterData: true });
      }
      observar(document.body);
    })();
    """
    _injetor = ("<script>const s = parent.document.createElement('script');"
                f"s.textContent = {json.dumps(_TRADUTOR)}; parent.document.head.appendChild(s);</script>")
    mo.Html(_ESTILO + mo.iframe(_injetor, height="0px").text)
    return (
        LINHA,
        LOGO,
        erro_pbi,
        icone,
        indicadores,
        json_fmt,
        nota,
        pilha,
        secao,
        subtitulo,
    )


@app.cell
def _(mo):
    # Sessão Microsoft do usuário nesta aba
    get_sessao, set_sessao = mo.state(None)
    return get_sessao, set_sessao


@app.cell
def _(autenticacao, mo, set_sessao):
    # ---------------- AUTENTICAÇÃO: retorno da Microsoft (?code=...&state=...) ----------------
    _params = mo.query_params()
    _erro = None
    if autenticacao.retorno_de_login(_params):
        _sessao, _erro = autenticacao.concluir_login(_params.to_dict())
        if _sessao:
            set_sessao(_sessao)
        _params.clear()   # remove o código da URL
    mo.callout(mo.md(f"**Não foi possível entrar.**\n\n{_erro}"), kind="danger") if _erro else None
    return


@app.cell
def _(get_sessao, powerbi):
    sessao = get_sessao()
    cliente_pbi = powerbi.Cliente(sessao.obter_token) if sessao else None
    return cliente_pbi, sessao


@app.cell
def _(mo):
    botao_sair = mo.ui.run_button(label="Sair", kind="neutral")
    return (botao_sair,)


@app.cell
def _(botao_sair, set_sessao):
    if botao_sair.value:
        set_sessao(None)
    return


@app.cell
def _(LINHA, LOGO, autenticacao, botao_sair, config, icone, mo, pilha, sessao):
    # ---------------- INTERFACE: cabeçalho e conta ----------------
    _faixa = mo.Html(f"""
    <section class="avia-faixa">
      <div class="avia-topo">
        <a class="avia-marca" href="/" target="_top">{LOGO}<span>AvIA <b>Analytics</b></span></a>
        <a class="avia-voltar" href="/" target="_top">{icone("voltar", 16)} Voltar para o início</a>
      </div>
      <h1>Radar de Leads</h1>
      <p>Identifique oportunidades por localização e atividade econômica, confira o dashboard
         e salve sua própria cópia no Power BI.</p>
      {LINHA}
    </section>""")

    if sessao:
        _conta = mo.Html(f"""
        <div class="avia-conta">
          <div class="avia-conta-info"><span class="avia-avatar">{icone("usuario", 20)}</span>
            <div>Conectado como <b>{sessao.nome}</b><p>Sua conta Microsoft está pronta para criar relatórios.</p></div></div>
          <div>{botao_sair}</div>
        </div>""")
    elif config.problemas():
        _conta = mo.callout(mo.md("**O login Microsoft não está configurado.**\n\n" + "\n".join(f"- {p}".replace("\\", "\\\\") for p in config.problemas())), kind="warn")
    else:
        try:
            _link = autenticacao.link_login()
            _conta = mo.Html(f"""
            <div class="avia-conta">
              <div class="avia-conta-info"><span class="avia-avatar">{icone("usuario", 20)}</span>
                <div><b>Entre com sua conta Microsoft</b>
                <p>Entre antes de montar a análise: a página é recarregada ao voltar da Microsoft.</p></div></div>
              <a class="avia-botao" href="{_link}" target="_top">{icone("entrar")} Entrar com Microsoft</a>
            </div>""")
        except Exception as _exc:  # noqa: BLE001
            _conta = mo.callout(mo.md(f"**Não foi possível iniciar o login.** {_exc}"), kind="danger")
    pilha([_faixa, _conta])
    return


@app.cell
def _(CNAE, dados, mo):
    # ---------------- Bases (CSV) ----------------
    try:
        df_loc, df_cnae = dados.carregar()
    except Exception as _exc:  # noqa: BLE001
        mo.stop(True, mo.callout(mo.md(f"**Não foi possível carregar as bases.** {_exc}"), kind="danger"))

    # Memória não reativa: mantém as escolhas ainda válidas quando um nível anterior muda.
    memoria = {}

    def seletor(chave, rotulo, opcoes, desabilitado=False):
        _validos = [v for v in memoria.get(chave, []) if v in set(opcoes)]
        memoria[chave] = _validos
        return mo.ui.multiselect(options=opcoes, value=_validos, label=rotulo, full_width=True,
                                 disabled=desabilitado, on_change=lambda v: memoria.__setitem__(chave, list(v)))

    def nivel(chave, df):
        _rotulo, _coluna, _ = CNAE[chave]
        return seletor(chave, _rotulo, dados.opcoes(df, _coluna))

    return df_cnae, df_loc, memoria, nivel, seletor


@app.cell
def _(config, dados, df_loc, seletor):
    sel_uf = seletor("uf", "Estado (UF) *", dados.opcoes(df_loc, config.COL_UF))
    return (sel_uf,)


@app.cell
def _(config, dados, df_loc, sel_uf, seletor):
    _municipios = (dados.opcoes(dados.filtrar(df_loc, config.COL_UF, sel_uf.value), config.COL_MUN)
                   if sel_uf.value else [])
    sel_mun = seletor("mun", "Município", _municipios, desabilitado=not sel_uf.value)
    return (sel_mun,)


@app.cell
def _(df_cnae, nivel):
    sel_setor = nivel("setor", df_cnae)
    return (sel_setor,)


@app.cell
def _(CNAE, dados, df_cnae, nivel, sel_setor):
    df_n1 = dados.filtrar(df_cnae, CNAE["setor"][1], sel_setor.value)
    sel_secao = nivel("secao", df_n1)
    return df_n1, sel_secao


@app.cell
def _(CNAE, dados, df_n1, nivel, sel_secao):
    df_n2 = dados.filtrar(df_n1, CNAE["secao"][1], sel_secao.value)
    sel_divisao = nivel("divisao", df_n2)
    return df_n2, sel_divisao


@app.cell
def _(CNAE, dados, df_n2, nivel, sel_divisao):
    df_n3 = dados.filtrar(df_n2, CNAE["divisao"][1], sel_divisao.value)
    sel_grupo = nivel("grupo", df_n3)
    return df_n3, sel_grupo


@app.cell
def _(CNAE, dados, df_n3, nivel, sel_grupo):
    df_n4 = dados.filtrar(df_n3, CNAE["grupo"][1], sel_grupo.value)
    sel_classe = nivel("classe", df_n4)
    return df_n4, sel_classe


@app.cell
def _(CNAE, dados, df_n4, nivel, sel_classe):
    _df_n5 = dados.filtrar(df_n4, CNAE["classe"][1], sel_classe.value)
    sel_subclasse = nivel("subclasse", _df_n5)
    return (sel_subclasse,)


@app.cell
def _(
    mo,
    secao,
    sel_classe,
    sel_divisao,
    sel_grupo,
    sel_mun,
    sel_secao,
    sel_setor,
    sel_subclasse,
    sel_uf,
):
    # ---------------- INTERFACE: etapas 1 e 2 ----------------
    mo.vstack([
        secao(1, "Localização",
              "Selecione ao menos um estado. Sem município, todos os municípios dos estados escolhidos entram na análise.",
              [mo.hstack([sel_uf, sel_mun], widths="equal", gap=1.5)]),
        secao(2, "Atividade econômica",
              "Refine pela CNAE, do grande setor à subclasse. Todos os níveis são opcionais e cada um limita as opções do seguinte.",
              [mo.hstack([sel_setor, sel_secao], widths="equal", gap=1.5),
               mo.hstack([sel_divisao, sel_grupo], widths="equal", gap=1.5),
               mo.hstack([sel_classe, sel_subclasse], widths="equal", gap=1.5)]),
    ], gap=1.25)
    return


@app.cell
def _(
    CNAE,
    config,
    json,
    powerbi,
    sel_classe,
    sel_divisao,
    sel_grupo,
    sel_mun,
    sel_secao,
    sel_setor,
    sel_subclasse,
    sel_uf,
):
    # Seleção atual -> filtros do Power BI
    _niveis = {"setor": sel_setor, "secao": sel_secao, "divisao": sel_divisao,
               "grupo": sel_grupo, "classe": sel_classe, "subclasse": sel_subclasse}
    selecoes_cnae = {CNAE[_k][0]: list(_el.value) for _k, _el in _niveis.items()}
    selecoes = {"UF": list(sel_uf.value), "Município": list(sel_mun.value), **selecoes_cnae}
    _campos = [(config.CAMPO_UF, selecoes["UF"]), (config.CAMPO_MUN, selecoes["Município"])]
    _campos += [(CNAE[_k][2], list(_el.value)) for _k, _el in _niveis.items()]

    tem_uf = bool(sel_uf.value)
    filtros = powerbi.filtros_embedded(_campos)
    assinatura = json.dumps(selecoes, ensure_ascii=False, sort_keys=True)
    link_service = powerbi.url_service(config.WORKSPACE_ID, config.REPORT_ID, config.PAGE_NAME, config.TENANT_ID, _campos)
    cfg_origem = powerbi.config_embedded(config.WORKSPACE_ID, config.REPORT_ID, config.PAGE_NAME, filtros)
    return (
        assinatura,
        cfg_origem,
        filtros,
        link_service,
        selecoes,
        selecoes_cnae,
        tem_uf,
    )


@app.cell
def _(indicadores, mo, pilha, secao, selecoes, selecoes_cnae, subtitulo, tem_uf):
    # ---------------- INTERFACE: etapa 3 (resumo) ----------------
    def _linhas(itens):
        return "\n".join(f"- **{rotulo}:** {', '.join(v)}" for rotulo, v in itens if v)

    _uf, _mun = selecoes["UF"], selecoes["Município"]
    _geo = (_linhas([("UF", _uf), ("Município", _mun)])
            + ("" if _mun else "\n\n_Todos os municípios dos estados selecionados._")) if _uf else "_Nenhum estado selecionado._"
    _cnae = _linhas(selecoes_cnae.items()) or "_Todas as atividades econômicas, sem restrição de CNAE._"

    secao(3, "Resumo da análise", "Confira os critérios que serão usados no dashboard e na incorporação.", [
        indicadores([
            ("Estados selecionados", len(_uf)),
            ("Critérios CNAE", sum(bool(v) for v in selecoes_cnae.values())),
            ("Valores selecionados", sum(len(v) for v in selecoes.values())),
        ]),
        mo.hstack([
            pilha([subtitulo("Abrangência geográfica"), mo.md(_geo)], gap=0.5),
            pilha([subtitulo("Perfil econômico"), mo.md(_cnae)], gap=0.5),
        ], widths="equal", gap=2),
        None if tem_uf else mo.callout(mo.md("Selecione pelo menos um estado para gerar o dashboard."), kind="warn"),
    ])
    return


@app.cell
def _(mo, tem_uf):
    botao_previa = mo.ui.run_button(label="Gerar dashboard", disabled=not tem_uf, full_width=True)
    return (botao_previa,)


@app.cell
def _(
    botao_previa,
    cfg_origem,
    icone,
    link_service,
    mo,
    nota,
    powerbi,
    secao,
    sessao,
    tem_uf,
):
    # ---------------- INTERFACE: etapa 4 (prévia) ----------------
    _link = (mo.Html(f'<a class="avia-link" href="{link_service}" target="_blank" rel="noopener">'
                     f'Abrir no Power BI Service {icone("externo", 15)}</a>') if tem_uf else None)
    if not botao_previa.value:
        _conteudo = nota("Gere o dashboard com a seleção atual para conferir os resultados antes de criar sua cópia.")
    elif not sessao:
        _conteudo = mo.callout(mo.md("Entre com sua conta Microsoft para ver a prévia aqui. Você ainda pode abrir o relatório no Power BI Service."), kind="info")
    else:
        _token = sessao.obter_token()
        _conteudo = (mo.iframe(powerbi.editor_html(cfg_origem, _token, "previa"), height="760px") if _token
                     else mo.callout(mo.md("Sua sessão Microsoft expirou. Entre novamente."), kind="warn"))
    secao(4, "Gerar dashboard", "Confira sua análise no Power BI e salve uma cópia com o nome que desejar.",
          [botao_previa, _conteudo, _link])
    return


@app.cell
def _(mo):
    get_copia, set_copia = mo.state(None)
    return get_copia, set_copia


@app.cell
def _(mo, pilha, subtitulo):
    nome_relatorio = mo.ui.text(label="Nome do novo relatório", placeholder="Radar de Leads | CE | Indústria", full_width=True)
    pilha([subtitulo("Criar relatório no workspace"), nome_relatorio], gap=0.5)
    return (nome_relatorio,)


@app.cell
def _(assinatura, get_copia, mo, nota, pilha, sessao, tem_uf):
    # Recriados quando a seleção muda (referência a `assinatura`): a conferência precisa ser refeita.
    _ = assinatura
    conferido = mo.ui.checkbox(label="Conferi o dashboard com a seleção atual")
    botao_criar = mo.ui.run_button(label="Criar relatório no workspace", full_width=True,
                                   disabled=not (sessao and tem_uf) or get_copia() is not None)
    pilha([conferido, botao_criar,
           None if sessao else nota("Entre com sua conta Microsoft para criar relatórios.")], gap=0.75)
    return botao_criar, conferido


@app.cell
def _(
    assinatura,
    botao_criar,
    cliente_pbi,
    conferido,
    config,
    erro_pbi,
    filtros,
    get_copia,
    mo,
    nome_relatorio,
    powerbi,
    set_copia,
):
    # Ação do botão "Criar relatório no workspace"
    mo.stop(not botao_criar.value)
    _nome = nome_relatorio.value.strip()
    if get_copia() is not None:
        _saida = mo.callout(mo.md("Uma cópia já foi criada. Use **Preparar outra criação** abaixo."), kind="warn")
    elif not conferido.value:
        _saida = mo.callout(mo.md("Marque **Conferi o dashboard** antes de criar a cópia."), kind="warn")
    elif not _nome:
        _saida = mo.callout(mo.md("Informe um nome para o novo relatório."), kind="warn")
    else:
        try:
            with mo.status.spinner(title="Criando a cópia do relatório…"):
                _novo = powerbi.clonar(cliente_pbi, config.WORKSPACE_ID, config.REPORT_ID, _nome)
            set_copia({"relatorio": _novo, "filtros": filtros, "assinatura": assinatura})
            _saida = None
        except powerbi.ErroPowerBI as _exc:
            _saida = erro_pbi(_exc)
    _saida
    return


@app.cell
def _(cliente_pbi, config, erro_pbi, get_copia, icone, mo, pilha, powerbi, sessao):
    # Depois da clonagem: confere o modelo semântico e salva os filtros na cópia
    copia = get_copia()
    export_copia = None
    _itens = []
    if copia:
        _rel = copia["relatorio"]
        _abrir = f"https://app.powerbi.com/groups/{config.WORKSPACE_ID}/reports/{_rel['id']}?ctid={config.TENANT_ID}"
        _itens.append(mo.Html(
            f'<div class="avia-conta"><div class="avia-conta-info"><span class="avia-avatar">{icone("ok", 20)}</span>'
            f'<div><b>Cópia criada:</b> {_rel.get("name", "Novo relatório")}</div></div>'
            f'<a class="avia-link" href="{_abrir}" target="_blank" rel="noopener">'
            f'Abrir o novo relatório no workspace {icone("externo", 15)}</a></div>'))
        try:
            _det = powerbi.detalhes(cliente_pbi, config.WORKSPACE_ID, _rel["id"])
            if _det.get("datasetId") != _rel["datasetIdEsperado"]:
                raise powerbi.ErroPowerBI("Não foi possível confirmar o mesmo modelo semântico. Verifique a cópia no workspace.")
            export_copia = powerbi.config_embedded(config.WORKSPACE_ID, _rel["id"], config.PAGE_NAME, copia["filtros"],
                                                   url=_det["embedUrl"], painel_filtros=True, datasetId=_det["datasetId"])
            _itens.append(mo.iframe(powerbi.editor_html(export_copia, sessao.obter_token() or "", "salvar"), height="760px"))
        except powerbi.ErroPowerBI as _exc:
            _itens.append(erro_pbi(_exc, "A cópia já foi criada — não crie outra. "))
            export_copia = powerbi.config_embedded(config.WORKSPACE_ID, _rel["id"], config.PAGE_NAME, copia["filtros"])
    pilha(_itens, gap=0.75) if _itens else None
    return copia, export_copia


@app.cell
def _(assinatura, copia, mo):
    (mo.callout(mo.md("A seleção mudou depois da criação. A cópia mantém os critérios usados quando foi criada."), kind="info")
     if copia and copia["assinatura"] != assinatura else None)
    return


@app.cell
def _(mo):
    botao_nova = mo.ui.run_button(label="Preparar outra criação")
    return (botao_nova,)


@app.cell
def _(botao_nova, get_copia):
    botao_nova if get_copia() is not None else None
    return


@app.cell
def _(botao_nova, set_copia):
    if botao_nova.value:
        set_copia(None)
    return


@app.cell
def _(
    cfg_origem,
    config,
    copia,
    export_copia,
    filtros,
    json_fmt,
    mo,
    nota,
    secao,
    subtitulo,
    tem_uf,
):
    # ---------------- INTERFACE: etapa 5 (Power BI Embedded) ----------------
    _cfg = export_copia or cfg_origem
    _filtros = copia["filtros"] if copia else filtros
    _origem = (
        mo.callout(mo.md(f"Dados da nova cópia **{copia['relatorio'].get('name', '')}**. Aguarde a confirmação de salvamento acima antes de entregar a integração."), kind="success")
        if copia else nota("Estes são os dados do relatório original. Depois de criar a cópia, eles passam a ser os dela."))
    secao(5, "Integração com o Power BI Embedded",
          "Identificadores do relatório e filtros em JSON, prontos para o SDK JavaScript do Power BI. "
          "No Embedded, os filtros vão na configuração, e não no endereço.", [
        _origem,
        mo.hstack([mo.md(f"**ID do workspace**\n\n```text\n{config.WORKSPACE_ID}\n```"),
                   mo.md(f"**ID do relatório**\n\n```text\n{_cfg['id']}\n```")], widths="equal", gap=1),
        mo.md(f"**Página do relatório**\n\n```text\n{_cfg['pageName']}\n```"),
        mo.md(f"**Endereço de incorporação (sem filtros)**\n\n```text\n{_cfg['embedUrl']}\n```"),
        subtitulo("Filtros selecionados"),
        nota("Use esta lista como a propriedade <code>filters</code> da configuração do SDK."),
        mo.md(f"```json\n{json_fmt(_filtros)}\n```"),
        mo.download(data=json_fmt(_filtros).encode("utf-8"), filename="radar_leads_filtros.json",
                    mimetype="application/json", label="Baixar JSON dos filtros", disabled=not tem_uf),
        subtitulo("Configuração completa"),
        nota("A aplicação integradora acrescenta o token de acesso e o tipo de autenticação."),
        mo.md(f"```json\n{json_fmt(_cfg)}\n```"),
        mo.download(data=json_fmt(_cfg).encode("utf-8"),
                    filename="radar_novo_relatorio.json" if copia else "radar_leads_embedded.json",
                    mimetype="application/json", label="Baixar configuração do Embedded", disabled=not tem_uf),
    ])
    return


if __name__ == "__main__":
    app.run()
