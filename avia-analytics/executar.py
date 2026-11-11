"""AvIA Analytics — sobe a home e as aplicações juntas.

Como rodar (no terminal do VS Code, dentro desta pasta):
    pip install -r requirements.txt
    python executar.py                 ->  http://localhost:8000

Credenciais do Radar: arquivo radar_leads/.env com
    AVIA_TENANT_ID=...
    AVIA_CLIENT_ID=...
    AVIA_CLIENT_SECRET=...
    RADAR_REDIRECT_URI=http://localhost:8000/radar-leads/   (cadastre igual no Entra)

Credenciais do GovMonitor: arquivo govmonitor/.env com
    GOVMONITOR_TENANT_ID=...
    GOVMONITOR_CLIENT_ID=...
    GOVMONITOR_CLIENT_SECRET=...
    GOVMONITOR_REDIRECT_URI=http://localhost:8000/govmonitor/   (cadastre igual no Entra)

Cada aplicação também roda sozinha:
    marimo edit home/home.py
    marimo run  radar_leads/radar_leads.py     (aí o endereço é http://localhost:2718/)
    marimo run  govmonitor/govmonitor.py
"""

import os
from pathlib import Path

import marimo
import uvicorn

PASTA = Path(__file__).resolve().parent


class SemCacheHtml:
    """Middleware ASGI: manda Cache-Control: no-store nas páginas HTML.

    Sem isso, o botão voltar do navegador reaproveita a página antiga (cache HTTP / bfcache)
    com a conexão do marimo já morta. Com no-store, o navegador busca uma página nova.
    Só toca em respostas HTTP de HTML: websockets e arquivos estáticos (JS/CSS) não mudam.
    """

    def __init__(self, app):
        self.app = app

    async def __call__(self, scope, receive, send):
        if scope["type"] != "http":
            await self.app(scope, receive, send)
            return

        async def enviar(mensagem):
            if mensagem["type"] == "http.response.start":
                cabecalhos = list(mensagem.get("headers", []))
                tipo = next((v for k, v in cabecalhos if k.lower() == b"content-type"), b"")
                if b"text/html" in tipo.lower():
                    cabecalhos = [(k, v) for k, v in cabecalhos
                                  if k.lower() not in (b"cache-control", b"etag", b"last-modified")]
                    cabecalhos.append((b"cache-control", b"no-store"))
                    mensagem = {**mensagem, "headers": cabecalhos}
            await send(mensagem)

        await self.app(scope, receive, enviar)


servidor = (
    marimo.create_asgi_app(quiet=True)
    .with_app(path="/radar-leads", root=str(PASTA / "radar_leads" / "radar_leads.py"))
    .with_app(path="/govmonitor", root=str(PASTA / "govmonitor" / "govmonitor.py"))
    # .with_app(path="/sindicatos", root=str(PASTA / "sindicatos" / "sindicatos.py"))
    .with_app(path="/", root=str(PASTA / "home" / "home.py"))
)
app = servidor.build()
app.add_middleware(SemCacheHtml)

if __name__ == "__main__":
    uvicorn.run(app, host=os.getenv("HOST", "127.0.0.1"), port=int(os.getenv("PORT", "8000")))