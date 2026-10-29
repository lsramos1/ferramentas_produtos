import marimo

__generated_with = "0.25.2"
app = marimo.App(width="medium", app_title="AvIA Analytics")


@app.cell
def _():
    import marimo as mo
    return (mo,)


@app.cell
def _():
    # ------------------------------------------------------------------
    # APLICAÇÕES: uma entrada por aplicação.
    # url = endereço onde a app é servida (ver executar.py)
    # Para liberar GovMonitor/Sindicatos, troque disponivel para True.
    # ------------------------------------------------------------------
    APLICACOES = [
        {
            "nome": "Radar de Leads",
            "icone": "radar",
            "descricao": "Monte análises de mercado por localização e atividade econômica (CNAE) "
                         "e salve sua própria cópia do dashboard no Power BI.",
            "url": "/radar-leads/",
            "disponivel": True,
        },
        {
            "nome": "GovMonitor",
            "icone": "gov",
            "descricao": "Explore indicadores de saúde, educação, economia, emprego e infraestrutura "
                         "por estado, mesorregião e município e salve sua cópia no Power BI.",
            "url": "/govmonitor/",
            "disponivel": True,
        },
        {
            "nome": "Sindicatos",
            "icone": "pessoas",
            "descricao": "Aplicação em preparação.",
            "url": "/sindicatos/",
            "disponivel": False,
        },
    ]
    return (APLICACOES,)


@app.cell
def _():
    # ------------------------------------------------------------------
    # ÍCONES: todos desenhados aqui em SVG, sem nenhum recurso externo
    # ------------------------------------------------------------------
    _TRACOS = {
        "radar": '<circle cx="12" cy="12" r="9"/><circle cx="12" cy="12" r="5"/>'
                 '<path d="M12 12l6-6"/><circle cx="12" cy="12" r="1.4" fill="currentColor"/>',
        "gov": '<path d="M2 10h20L12 3z"/><path d="M5 10v10M9.7 10v10M14.3 10v10M19 10v10"/>'
               '<path d="M3 21h18"/>',
        "pessoas": '<circle cx="9" cy="8" r="3"/><path d="M3 20c0-3.3 2.7-6 6-6s6 2.7 6 6"/>'
                   '<circle cx="17" cy="9" r="2.5"/><path d="M16 14.2c2.9.3 5 2.7 5 5.8"/>',
        "seta": '<path d="M5 12h14M13 6l6 6-6 6"/>',
    }

    def icone(nome, tam=20):
        return (f'<svg width="{tam}" height="{tam}" viewBox="0 0 24 24" fill="none" '
                f'stroke="currentColor" stroke-width="1.8" stroke-linecap="round" '
                f'stroke-linejoin="round" aria-hidden="true">{_TRACOS[nome]}</svg>')

    LOGO = (
        '<svg width="36" height="36" viewBox="0 0 40 40" aria-hidden="true">'
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
    return LINHA, LOGO, icone


@app.cell
def _():
    # ------------------------------------------------------------------
    # VISUAL: paleta branco + azul #154388 (tons derivados do mesmo azul)
    # Fontes do próprio sistema, nada carregado de fora.
    # ------------------------------------------------------------------
    ESTILO = """
    :root {
      --avia-azul: #154388;
      --avia-azul-escuro: #0E2F61;
      --avia-azul-claro: #E8EEF7;
      --avia-azul-borda: #C9D6EA;
      --avia-texto: #1B2B45;
      --avia-suave: #5A6B86;
      --avia-branco: #FFFFFF;
      --avia-fonte: "Segoe UI", system-ui, -apple-system, Roboto, "Helvetica Neue", Arial, sans-serif;
      --marimo-heading-font: var(--avia-fonte);
      --marimo-text-font: var(--avia-fonte);
    }

    .avia-shell { font-family: var(--avia-fonte); color: var(--avia-texto); background: var(--avia-branco); }
    .avia-marca { position: relative; display: flex; align-items: center; gap: 10px;
      font: 600 17px var(--avia-fonte); margin-bottom: 40px; }
    .avia-marca b { font-weight: 400; opacity: .8; }

    /* ---------- faixa azul ---------- */
    .avia-faixa { position: relative; overflow: hidden; color: #fff; border-radius: 16px;
      background: linear-gradient(135deg, var(--avia-azul) 0%, var(--avia-azul-escuro) 100%);
      padding: 30px 32px 110px; }
    .avia-faixa::before { content: ""; position: absolute; inset: 0;
      background-image: radial-gradient(rgba(255,255,255,.09) 1px, transparent 1px);
      background-size: 24px 24px; pointer-events: none; }
    .avia-faixa h1 { font: 700 clamp(28px, 4vw, 44px)/1.1 var(--avia-fonte) !important;
      letter-spacing: -1px; color: #fff !important; margin: 0 0 14px !important;
      max-width: 20ch; border: 0 !important; position: relative; }
    .avia-faixa p { color: #D4DEEF; font-size: 17px; max-width: 54ch; margin: 0; position: relative; }
    .avia-linha { position: absolute; left: 0; bottom: 0; width: 100%; height: 110px; }

    .avia-titulo-secao { font: 600 13px var(--avia-fonte); letter-spacing: 1.2px;
      text-transform: uppercase; color: var(--avia-suave); margin: 32px 0 14px; }

    /* ---------- cartões das aplicações ---------- */
    .avia-cartoes { display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 16px; }
    .avia-cartao { border: 1px solid var(--avia-azul-borda); border-radius: 14px; padding: 20px;
      display: flex; flex-direction: column; gap: 12px; background: #fff;
      transition: border-color .15s, box-shadow .15s, transform .15s; }
    .avia-cartao:hover { border-color: var(--avia-azul); transform: translateY(-2px);
      box-shadow: 0 10px 24px -14px rgba(21,67,136,.55); }
    .avia-cartao.em-breve:hover { border-color: var(--avia-azul-borda); transform: none; box-shadow: none; }
    .avia-cartao h3 { font: 600 19px var(--avia-fonte) !important; margin: 0 !important;
      color: var(--avia-azul) !important; }
    .avia-cartao p { margin: 0; color: var(--avia-suave); font-size: 14.5px; line-height: 1.5; flex: 1; }
    .avia-cartao.em-breve h3 { color: var(--avia-suave) !important; }
    .avia-icone { width: 42px; height: 42px; border-radius: 12px; display: grid; place-items: center;
      background: var(--avia-azul-claro); color: var(--avia-azul); }
    .avia-cartao-topo { display: flex; justify-content: space-between; align-items: center; }

    .avia-selo { font-size: 12px; font-weight: 600; padding: 4px 10px; border-radius: 999px;
      background: var(--avia-azul); color: #fff; }
    .avia-selo.em-breve { background: var(--avia-azul-claro); color: var(--avia-suave); }

    .avia-botao { display: inline-flex; align-items: center; justify-content: center; gap: 8px;
      white-space: nowrap; border-radius: 10px; padding: 11px 18px; font: 600 15px var(--avia-fonte);
      text-decoration: none !important; background: var(--avia-azul); color: #fff !important;
      transition: background .15s; }
    .avia-botao:hover { background: var(--avia-azul-escuro); }
    .avia-botao:focus-visible { outline: 3px solid var(--avia-azul-borda); outline-offset: 3px; }
    .avia-botao.inativo { background: var(--avia-azul-claro); color: var(--avia-suave) !important;
      pointer-events: none; }

    @media (max-width: 760px) {
      .avia-faixa { padding: 22px 20px 90px; }
    }
    """
    return (ESTILO,)


@app.cell
def _(APLICACOES, ESTILO, LINHA, LOGO, icone, mo):
    # INTERFACE: faixa principal + um cartão por aplicação
    def _cartao(a):
        if a["disponivel"]:
            selo = '<span class="avia-selo">Disponível</span>'
            botao = (f'<a class="avia-botao" href="{a["url"]}" target="_top">'
                     f'Abrir {a["nome"]} {icone("seta", 18)}</a>')
            classe = "avia-cartao"
        else:
            selo = '<span class="avia-selo em-breve">Em breve</span>'
            botao = '<span class="avia-botao inativo">Em breve</span>'
            classe = "avia-cartao em-breve"
        return f"""
        <div class="{classe}">
          <div class="avia-cartao-topo"><span class="avia-icone">{icone(a["icone"], 22)}</span>{selo}</div>
          <h3>{a["nome"]}</h3>
          <p>{a["descricao"]}</p>
          {botao}
        </div>"""

    mo.Html(f"""
    <style>{ESTILO}</style>
    <div class="avia-shell">
      <section class="avia-faixa">
        <div class="avia-marca">{LOGO}<span>AvIA <b>Analytics</b></span></div>
        <h1>Dados de mercado, prontos para a sua decisão.</h1>
        <p>Escolha uma aplicação, defina seus critérios e gere dashboards no Power BI
           com a sua conta Microsoft.</p>
        {LINHA}
      </section>
      <div class="avia-titulo-secao">Aplicações</div>
      <div class="avia-cartoes">{"".join(_cartao(a) for a in APLICACOES)}</div>
    </div>
    """)
    return


if __name__ == "__main__":
    app.run()
