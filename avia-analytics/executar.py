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

servidor = (
    marimo.create_asgi_app(quiet=True)
    .with_app(path="/radar-leads", root=str(PASTA / "radar_leads" / "radar_leads.py"))
    .with_app(path="/govmonitor", root=str(PASTA / "govmonitor" / "govmonitor.py"))
    # .with_app(path="/sindicatos", root=str(PASTA / "sindicatos" / "sindicatos.py"))
    .with_app(path="/", root=str(PASTA / "home" / "home.py"))
)
app = servidor.build()

if __name__ == "__main__":
    uvicorn.run(app, host=os.getenv("HOST", "127.0.0.1"), port=int(os.getenv("PORT", "8000")))