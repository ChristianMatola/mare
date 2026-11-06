from __future__ import annotations

import html
import io
import json
import mimetypes
import os
import re
import shutil
import sqlite3
import subprocess
import tempfile
import zipfile
from contextlib import contextmanager
from datetime import datetime
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, quote, quote_plus, urlparse


ROOT = Path(__file__).resolve().parent
DB_PATH = ROOT / "dados_relatorios.sqlite3"
OUTPUT_DIR = ROOT / "gerados"
TEMPLATE_PATH = ROOT / "modelo_pagamento_grifado.docx"

FIELDS = [
    ("tipo_processo", "Tipo de Processo", "text", True),
    ("num_relatorio", "Nº Relatório", "text", True),
    ("data_relatorio", "Data Relatório", "date", False),
    ("num_processo", "Nº Processo", "text", True),
    ("data_inicio", "Data Início", "date", True),
    ("setor", "Setor", "text", False),
    ("nome_evento", "Nome Evento", "text", False),
    ("deao_artista", "De(a/o) Artista", "text", False),
    ("razao_social", "Razão Social", "text", True),
    ("cnpj", "CNPJ", "text", True),
    ("nota_fiscal", "Nota Fiscal", "text", False),
    ("data_emissao_nf", "Data Emissão NF", "date", False),
    ("valor", "Valor", "text", False),
    ("valor_extenso", "Valor Extenso", "text", False),
    ("volumes", "Volume(s)", "text", False),
    ("num_contrato", "Nº Contrato", "text", False),
    ("data_ass_contrato", "Data Ass Contrato", "date", False),
    ("pgs_copia_contrato", "Pgs Cópia Contrato", "text", False),
    ("num_paginas", "Nº Páginas", "text", False),
    ("num_processo_contratacao", "Nº Processo Contratação", "text", False),
    ("num_ordem_servico", "Nº Ordem de Serviço", "text", False),
]

YELLOW_FIELDS = [
    ("data_relato", "Data do relato", "date"),
    ("num_empenho", "Nota de empenho nº", "text"),
    ("folhas_empenho", "Folhas da nota de empenho", "text"),
    ("folhas_termo_referencia", "Folhas do termo de referência", "text"),
    ("folha_memoria_calculo", "Folha da memória de cálculo", "text"),
    ("folha_autorizo", "Folha da publicação do autorizo", "text"),
    ("folhas_extrato_contrato", "Folhas do extrato do contrato", "text"),
    ("folhas_nota_fiscal", "Folhas da nota fiscal", "text"),
    ("folha_requisicao_pagamento", "Folha da requisição de pagamento", "text"),
    ("folhas_relatorio_fiscalizacao", "Folhas do relatório de fiscalização", "text"),
    ("folha_ordem_servico", "Folha da ordem de serviço", "text"),
    ("folha_relatorio_entrada", "Folha do relatório de entrada", "text"),
    ("folha_mapa_controle", "Folha do mapa de controle", "text"),
    ("folha_portaria", "Folha da portaria", "text"),
    ("num_portaria", "Número da portaria", "text"),
    ("folhas_tce", "Folhas do encaminhamento ao TCE/RJ", "text"),
    ("cert_trabalhista_emissao", "Certidão trabalhista — emissão", "date"),
    ("cert_trabalhista_validade", "Certidão trabalhista — validade", "date"),
    ("cert_trabalhista_folha", "Certidão trabalhista — folha", "text"),
    ("cert_federal_emissao", "Certidão federal — emissão", "date"),
    ("cert_federal_validade", "Certidão federal — validade", "date"),
    ("cert_federal_folha", "Certidão federal — folha", "text"),
    ("cert_fgts_emissao", "CRF/FGTS — emissão", "date"),
    ("cert_fgts_validade", "CRF/FGTS — validade", "date"),
    ("cert_fgts_folha", "CRF/FGTS — folha", "text"),
]

ALL_KEYS = [x[0] for x in FIELDS] + [x[0] for x in YELLOW_FIELDS]

APONTAMENTO_FIELDS = [
    ("sequencia", "Sequência"),
    ("pagina", "Página"),
    ("apontamento", "Apontamento"),
    ("status", "Status"),
]

STATUS_APONTAMENTO_PADRAO = ["ABERTO", "EM ANÁLISE", "PENDENTE", "RESOLVIDO", "DISPENSADO"]


class PostgresQueries:
    def __init__(self, connection):
        self.connection = connection

    def execute(self, query, params=()):
        return self.connection.execute(query.replace("?", "%s"), params)


@contextmanager
def db():
    if database_url := os.environ.get("MARE_DATABASE_URL"):
        import psycopg
        from psycopg.rows import dict_row

        with psycopg.connect(database_url, row_factory=dict_row, prepare_threshold=None,
                             connect_timeout=5, sslmode="require") as conn:
            conn.execute("SET LOCAL search_path TO mare")
            security = conn.execute("""
                SELECT current_user = 'mare_aplicacao'
                   AND row_security_active('mare.processos')
                   AND row_security_active('mare.apontamentos')
                   AND row_security_active('mare.status_apontamentos')
                   AND row_security_active('mare.modelos') AS protegida
            """).fetchone()
            if not security["protegida"]:
                raise RuntimeError("A conexão web exige a função mare_aplicacao sujeita a RLS.")
            yield PostgresQueries(conn)
        return
    conn = sqlite3.connect(DB_PATH)
    try:
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA journal_mode=WAL")
        with conn:
            yield conn
    finally:
        conn.close()


def init_db() -> None:
    with db() as conn:
        columns = ",\n".join(f"{key} TEXT NOT NULL DEFAULT ''" for key in ALL_KEYS)
        conn.execute(f"""
            CREATE TABLE IF NOT EXISTS processos (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                {columns},
                status TEXT NOT NULL DEFAULT 'RASCUNHO',
                criado_em TEXT NOT NULL,
                atualizado_em TEXT NOT NULL,
                arquivo_gerado TEXT NOT NULL DEFAULT ''
            )
        """)
        existing = {row[1] for row in conn.execute("PRAGMA table_info(processos)").fetchall()}
        for key in ALL_KEYS:
            if key not in existing:
                conn.execute(f"ALTER TABLE processos ADD COLUMN {key} TEXT NOT NULL DEFAULT ''")
        conn.execute("""
            CREATE TABLE IF NOT EXISTS apontamentos (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                processo_id INTEGER NOT NULL,
                sequencia TEXT NOT NULL DEFAULT '',
                pagina TEXT NOT NULL DEFAULT '',
                apontamento TEXT NOT NULL DEFAULT '',
                status TEXT NOT NULL DEFAULT '',
                criado_em TEXT NOT NULL,
                atualizado_em TEXT NOT NULL,
                FOREIGN KEY (processo_id) REFERENCES processos(id)
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS status_apontamentos (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nome TEXT NOT NULL UNIQUE,
                criado_em TEXT NOT NULL
            )
        """)
        now = datetime.now().isoformat(timespec="seconds")
        for status in STATUS_APONTAMENTO_PADRAO:
            conn.execute("INSERT INTO status_apontamentos (nome, criado_em) VALUES (?, ?) ON CONFLICT (nome) DO NOTHING", (status, now))


def esc(value: object) -> str:
    return html.escape(str(value or ""), quote=True)


def numero_sequencia(value: str) -> int:
    match = re.match(r"\s*([+-]?[0-9]{1,18})", value)
    return int(match.group(1)) if match else 0


def row_values(row: sqlite3.Row | None) -> dict[str, str]:
    return {key: str(row[key] or "") if row else "" for key in ALL_KEYS}


def validacao(data: dict[str, str]) -> list[str]:
    erros = []
    for key, label, _tipo, _obrigatorio in FIELDS:
        if not data.get(key, "").strip():
            erros.append(f"Preencha: {label}.")
    for key, label, _tipo in YELLOW_FIELDS:
        if not data.get(key, "").strip():
            erros.append(f"Preencha: {label}.")
    cnpj = re.sub(r"\D", "", data.get("cnpj", ""))
    if cnpj and len(cnpj) != 14:
        erros.append("O CNPJ deve conter 14 dígitos.")
    return erros


def layout(title: str, body: str, notice: str = "") -> bytes:
    notice_html = f'<div class="notice">{esc(notice)}</div>' if notice else ""
    page = f"""<!doctype html><html lang="pt-BR"><head><meta charset="utf-8">
    <meta name="viewport" content="width=device-width,initial-scale=1"><meta name="theme-color" content="#123d5a"><link rel="manifest" href="/manifest.json"><title>{esc(title)}</title>
    <style>
    :root{{font-family:Segoe UI,Arial,sans-serif;color:#17324d;background:#f4f7fb}}*{{box-sizing:border-box}}body{{margin:0;background:#f4f7fb}}
    .top{{background:#fff;border-top:5px solid #0d4d6c;box-shadow:0 2px 8px #0002;position:relative;overflow:hidden}}
    .top::after{{content:"";position:absolute;inset:0;background:url('/static/marica-background.png') center/cover no-repeat;opacity:.22;pointer-events:none}}
    .top-inner{{max-width:1280px;margin:auto;min-height:112px;padding:14px 28px;display:flex;align-items:center;gap:22px;position:relative;z-index:1}}
    .brand-logo{{width:205px;max-height:88px;object-fit:contain;object-position:left center}}
    .brand-copy{{border-left:1px solid #d6dfe5;padding-left:20px}}.brand-copy strong{{display:block;color:#154f70;font-size:20px;letter-spacing:.3px}}.brand-copy span{{display:block;color:#6a7a83;font-size:13px;margin-top:4px}}
    .top-actions{{margin-left:auto;color:#154f70;font-size:13px;text-align:right}}.top-actions a{{color:#154f70;text-decoration:none;margin-left:16px}}
    .menu{{background:#c43d3d;color:#fff;box-shadow:0 2px 4px #0002}}.menu-inner{{max-width:1280px;margin:auto;display:flex;align-items:center;gap:4px;padding:0 22px;overflow:auto}}.menu a{{color:#fff;text-decoration:none;font-weight:700;font-size:13px;padding:14px 17px;white-space:nowrap}}.menu a:hover{{background:#9f2d32}}
    main{{max-width:1280px;margin:0 auto;padding:0 22px 36px;background-image:linear-gradient(#fff8,#fff8),url('/static/marica-background.png');background-size:cover;background-attachment:fixed}}
    .page-title{{font-size:28px;color:#154f70;font-weight:750;padding:25px 0 18px;border-bottom:3px solid #c43d3d;margin-bottom:20px}}
    .card{{background:#fff;border-radius:12px;padding:22px;margin-bottom:18px;box-shadow:0 2px 10px #0002}}
    .grid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(230px,1fr));gap:14px}}
    label{{font-weight:600;font-size:13px;display:flex;flex-direction:column;gap:6px}}.field-label{{display:block;line-height:1.25}}.field-label:first-letter{{color:#a33a3a}}input,textarea{{font:inherit;padding:9px;border:1px solid #b6c3cf;border-radius:7px}}textarea{{min-height:80px}}.compact-input{{max-width:150px}}.description-field{{grid-column:span 2}}.description-field textarea{{min-height:150px;resize:vertical}}small{{font-weight:400;color:#647582;font-size:11px}}
    button,.button{{background:#0b6b52;color:#fff;border:0;border-radius:7px;padding:10px 14px;text-decoration:none;cursor:pointer;display:inline-block}}
    .secondary{{background:#557086}}.danger{{background:#a33a3a}}.notice{{background:#fff2c2;border-left:5px solid #d39e00;padding:12px;margin-bottom:18px}}
    table{{width:100%;border-collapse:collapse}}th,td{{padding:10px;border-bottom:1px solid #dde5eb;text-align:left;font-size:13px}}th{{background:#e9f0f5}}
    .actions{{display:flex;gap:8px;flex-wrap:wrap;margin-top:20px}}.compact-actions{{margin-top:0;margin-bottom:14px}}.form-tabs{{display:flex;gap:8px;margin:0 0 18px;border-bottom:3px solid #c43d3d;padding-bottom:8px}}.form-tabs a{{color:#154f70;font-weight:700;text-decoration:none;padding:8px 12px;border-radius:7px;background:#e9f0f5}}.form-tabs a.active{{background:#c43d3d;color:#fff}}.form-tabs a.disabled{{color:#71808a;cursor:not-allowed}}.form-tabs a:hover{{background:#c43d3d;color:#fff}}.record-summary{{display:flex;flex-wrap:wrap;gap:8px 22px;background:#f4f7fb;border-left:4px solid #c43d3d;padding:12px;margin-bottom:14px;font-size:13px}}.inline-form{{display:inline;margin:0}}.alert-link{{color:#a33a3a;font-weight:700}}.warning{{border-left-color:#a33a3a}}h1{{display:none}}h2{{font-size:18px;border-bottom:1px solid #dde5eb;padding-bottom:8px;color:#154f70}}
    @media (max-width: 900px){{.top-inner{{padding:12px 18px;min-height:92px}}.brand-logo{{width:165px}}.brand-copy strong{{font-size:17px}}.brand-copy span{{font-size:11px}}main{{padding-left:16px;padding-right:16px}}}}
    @media (max-width: 680px){{.top-inner{{padding:10px 14px;gap:12px;min-height:76px}}.brand-logo{{width:145px;max-height:62px}}.brand-copy{{display:none}}.top-actions{{font-size:11px}}.top-actions a{{margin-left:6px}}.menu-inner{{padding:0 8px;gap:0}}.menu a{{font-size:12px;padding:12px 11px}}main{{padding:0 12px 24px}}.page-title{{font-size:23px;padding:20px 0 14px}}.card{{padding:15px;border-radius:9px}}.grid{{grid-template-columns:1fr;gap:12px}}.description-field{{grid-column:span 1}}.form-tabs{{overflow:auto;white-space:nowrap}}table{{display:block;width:100%;white-space:normal}}table thead{{display:none}}table tbody,table tr{{display:block;width:100%}}table tr{{background:#fff;border:1px solid #dbe3e8;border-radius:9px;padding:9px 11px;margin-bottom:12px;box-shadow:0 2px 7px #0001}}table td{{display:flex;align-items:flex-start;justify-content:space-between;gap:12px;width:100%;padding:8px 0;border:0;white-space:normal;overflow-wrap:anywhere}}table td::before{{font-weight:700;color:#154f70;flex:0 0 84px}}table td:nth-child(1)::before{{content:'Relatório'}}table td:nth-child(2)::before{{content:'Processo'}}table td:nth-child(3)::before{{content:'Contratada'}}table td:nth-child(4)::before{{content:'Status'}}table td:nth-child(5)::before{{content:'Atualizado'}}table td:nth-child(6)::before{{content:'Apontamentos'}}table td:nth-child(7){{display:block;padding-top:10px;border-top:1px solid #e3e9ed;margin-top:4px}}table td:nth-child(7)::before{{content:'Ações';display:block;margin-bottom:7px}}table td:nth-child(7) .button{{width:auto;margin:0 6px 0 0}}.actions{{display:grid;grid-template-columns:1fr}}button,.button{{text-align:center;width:100%}}}}
    @media (max-width: 380px){{.brand-logo{{width:125px}}.menu a{{padding-left:9px;padding-right:9px;font-size:11px}}.page-title{{font-size:21px}}}}
    </style></head><body><header class="top"><div class="top-inner"><img class="brand-logo" src="/static/mare-logo.png" alt="Maricá Cultura e Turismo"><div class="brand-copy"><strong>Gestão de Processos</strong><span>CTMAR / MARÉ · Assessoria Especial de Controle Interno</span></div><div class="top-actions"><span>Ambiente local</span><br><a href="https://www.marica.rj.gov.br/" target="_blank" rel="noopener">Portal de Maricá ↗</a></div></div></header><nav class="menu"><div class="menu-inner"><a href="/">Início</a><a href="/">Processos</a><a href="/novo">Novo processo</a><a href="/">Relatórios</a><a href="/">Modelos</a><a href="/">Ajuda</a></div></nav><main><div class="page-title">{esc(title)}</div>{notice_html}{body}</main><script>if('serviceWorker' in navigator) navigator.serviceWorker.register('/sw.js');</script></body></html>"""
    return page.encode("utf-8")


def input_html(key: str, label: str, tipo: str, value: str, required: bool = False) -> str:
    req = " required" if required else ""
    title = f'<span class="field-label">{("* " if required else "")}{esc(label)}</span>'
    if tipo == "textarea":
        return f'<label>{title}<textarea name="{esc(key)}"{req}>{esc(value)}</textarea></label>'
    return f'<label>{title}<input name="{esc(key)}" type="{esc(tipo)}" value="{esc(value)}"{req}></label>'


def apontamentos_section(processo_id: int | None, row: sqlite3.Row | None) -> str:
    if not processo_id:
        return """<section class='card' id='apontamentos'><h2>APONTAMENTOS</h2>
        <p>Salve o processo primeiro para cadastrar apontamentos.</p></section>"""
    with db() as conn:
        items = conn.execute("SELECT * FROM apontamentos WHERE processo_id=? ORDER BY id", (processo_id,)).fetchall()
    items.sort(key=lambda item: numero_sequencia(item["sequencia"]))
    header = f"<div class='record-summary'><b>Nº do Relatório:</b> {esc(row['num_relatorio'])}<b>Nº do Processo:</b> {esc(row['num_processo'])}<b>Contratada:</b> {esc(row['razao_social'])}</div>"
    rows = []
    for item in items:
        rows.append(f"<tr><td>{esc(item['sequencia'])}</td><td>{esc(item['pagina'])}</td><td>{esc(item['apontamento'])}</td><td>{esc(item['status'])}</td><td><a class='button secondary' href='/editar-apontamento?id={item['id']}'>Editar Apontamento</a> <form class='inline-form' method='post' action='/excluir-apontamento'><input type='hidden' name='id' value='{item['id']}'><input type='hidden' name='processo_id' value='{processo_id}'><button class='danger' type='submit'>Excluir Apontamento</button></form></td></tr>")
    table = "<p>Nenhum apontamento cadastrado.</p>" if not rows else "<table class='apontamentos-table'><thead><tr><th>Sequência</th><th>Página</th><th>Apontamento</th><th>Status</th><th>Ações</th></tr></thead><tbody>" + "".join(rows) + "</tbody></table>"
    return f"<section class='card' id='apontamentos'><h2>APONTAMENTOS</h2>{header}<div class='actions compact-actions'><a class='button' href='/novo-apontamento?processo_id={processo_id}'>Novo Apontamento</a></div>{table}</section>"


def form_page(row: sqlite3.Row | None, notice: str = "") -> bytes:
    values = row_values(row)
    ident = row["id"] if row else ""
    primary = "Dados principais"
    main_fields = "".join(input_html(*field[:3], values[field[0]], True) for field in FIELDS)
    yellow_fields = "".join(input_html(field[0], field[1], field[2], values[field[0]], True) for field in YELLOW_FIELDS)
    tabs = f"<nav class='form-tabs'><a class='active' href='/editar?id={ident}'>DADOS DO PROCESSO</a><a href='/apontamentos?processo_id={ident}'>APONTAMENTOS</a></nav>" if ident else "<nav class='form-tabs'><a class='active' href='#dados'>DADOS DO PROCESSO</a><a class='disabled' href='#apontamentos'>APONTAMENTOS (salve o processo primeiro)</a></nav>"
    body = f"""<h1>{'Editar processo' if row else 'Novo processo'}</h1>
    {tabs}
    <form method="post" action="/salvar"><input type="hidden" name="id" value="{esc(ident)}">
    <section class="card" id="dados">
    <section class="card"><h2>{primary}</h2><div class="grid">{main_fields}</div></section>
    <section class="card"><h2>Checklist e campos documentais</h2><p>Preencha os dados que hoje aparecem como marcações amarelas no modelo. A aplicação salvará o checklist mesmo antes da geração.</p><div class="grid">{yellow_fields}</div></section>
    <div class="actions"><button type="submit">Salvar processo</button><a class="button secondary" href="/">Cancelar</a></div></section></form>"""
    return layout("Formulário", body, notice)


def apontamentos_page(processo: sqlite3.Row | None, notice: str = "") -> bytes:
    if processo is None:
        return layout("Apontamentos", "<section class='card'><p>Processo não encontrado.</p><a class='button secondary' href='/'>Voltar</a></section>", notice)
    tabs = f"<nav class='form-tabs'><a href='/editar?id={processo['id']}'>DADOS DO PROCESSO</a><a class='active' href='/apontamentos?processo_id={processo['id']}'>APONTAMENTOS</a></nav>"
    body = f"<h1>Apontamentos</h1>{tabs}{apontamentos_section(int(processo['id']), processo)}"
    return layout("Apontamentos", body, notice)


def apontamento_status_options() -> list[str]:
    with db() as conn:
        rows = conn.execute("SELECT nome FROM status_apontamentos ORDER BY nome").fetchall()
    return [str(row["nome"]) for row in rows]


def proxima_sequencia_apontamento(processo_id: int) -> str:
    with db() as conn:
        rows = conn.execute("SELECT sequencia FROM apontamentos WHERE processo_id=?", (processo_id,)).fetchall()
    return str(max((numero_sequencia(row["sequencia"]) for row in rows), default=0) + 1)


def apontamento_form_page(item: sqlite3.Row | None, processo: sqlite3.Row | None, notice: str = "") -> bytes:
    if processo is None:
        return layout("Apontamento", "<section class='card'><p>Processo não encontrado.</p><a class='button secondary' href='/'>Voltar</a></section>", notice)
    values = {key: str(item[key] or "") if item else "" for key, _label in APONTAMENTO_FIELDS}
    if not item:
        values["sequencia"] = proxima_sequencia_apontamento(int(processo["id"]))
    ident = item["id"] if item else ""
    status_options = "".join(f"<option value='{esc(status)}'></option>" for status in apontamento_status_options())
    fields = f"""<label class='compact-field'><span class='field-label'>* Sequência</span><input class='compact-input' name='sequencia' type='text' inputmode='numeric' value='{esc(values['sequencia'])}' required></label>
    <label class='compact-field'><span class='field-label'>* Página</span><input class='compact-input' name='pagina' type='text' inputmode='numeric' value='{esc(values['pagina'])}' placeholder='Ex.: 12 ou 50/56' required></label>
    <label class='description-field'><span class='field-label'>* Apontamento (Descrição)</span><textarea name='apontamento' rows='6' required>{esc(values['apontamento'])}</textarea></label>
    <label class='compact-field'><span class='field-label'>* Status</span><input name='status' list='status-options' value='{esc(values['status'])}' placeholder='Selecione ou digite um novo status' required><datalist id='status-options'>{status_options}</datalist><small>Você pode selecionar uma opção ou digitar uma nova; ela ficará disponível nos próximos apontamentos.</small></label>"""
    body = f"""<h1>{'Editar apontamento' if item else 'Novo apontamento'}</h1>
    <section class='card'><h2>APONTAMENTOS</h2><div class='record-summary'><b>Nº do Relatório:</b> {esc(processo['num_relatorio'])}<b>Nº do Processo:</b> {esc(processo['num_processo'])}<b>Contratada:</b> {esc(processo['razao_social'])}</div>
    <form method='post' action='/salvar-apontamento'><input type='hidden' name='id' value='{esc(ident)}'><input type='hidden' name='processo_id' value='{processo['id']}'><div class='grid'>{fields}</div><div class='actions'><button type='submit'>Salvar Apontamento</button><a class='button secondary' href='/apontamentos?processo_id={processo['id']}'>Cancelar</a></div></form></section>"""
    return layout("Apontamentos", body, notice)


def index_page(notice: str = "") -> bytes:
    with db() as conn:
        rows = conn.execute("SELECT p.*, (SELECT COUNT(*) FROM apontamentos a WHERE a.processo_id=p.id) AS apontamentos_count FROM processos p ORDER BY atualizado_em DESC").fetchall()
    trs = []
    alert_count = sum(int(row["apontamentos_count"] or 0) for row in rows)
    for row in rows:
        count = int(row["apontamentos_count"] or 0)
        badge = f"<a class='alert-link' href='/apontamentos?processo_id={row['id']}'>⚠ {count} apontamento(s)</a>" if count else "Nenhum"
        trs.append(f"<tr><td>{esc(row['num_relatorio'])}</td><td>{esc(row['num_processo'])}</td><td>{esc(row['razao_social'])}</td><td>{esc(row['status'])}</td><td>{esc(row['atualizado_em'])}</td><td>{badge}</td><td><a class='button secondary' href='/editar?id={row['id']}'>Editar</a> <a class='button' href='/gerar?id={row['id']}'>Gerar DOCX</a></td></tr>")
    global_alert = f"<div class='notice warning'>Existem {alert_count} apontamento(s) cadastrados nos processos. Consulte a coluna Apontamentos.</div>" if alert_count else ""
    table = "<p>Nenhum processo cadastrado.</p>" if not trs else "<table><thead><tr><th>Relatório</th><th>Processo</th><th>Contratada</th><th>Status</th><th>Atualizado</th><th>Apontamentos</th><th>Ações</th></tr></thead><tbody>" + "".join(trs) + "</tbody></table>"
    body = f"<h1>Processos</h1>{global_alert}<section class='card'>{table}</section>"
    return layout("Processos", body, notice)


def xml_escape(value: str) -> str:
    return html.escape(value or "", quote=False)


def model_font_run(run: str) -> str:
    """Fixa a fonte do corpo quando o modelo deixou a fonte apenas herdada."""
    if re.search(r"<w:rFonts\b", run):
        return run
    font = '<w:rFonts w:ascii="Times New Roman" w:eastAsia="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>'
    if re.search(r"<w:rPr\b[^>]*>", run):
        return re.sub(r"(<w:rPr\b[^>]*>)", r"\1" + font, run, count=1)
    return re.sub(r"(<w:r\b[^>]*>)", r"\1<w:rPr>" + font + "</w:rPr>", run, count=1)


def bold_run(run: str) -> str:
    """Adiciona negrito ao run OOXML sem alterar o restante da formatação."""
    if re.search(r"<w:b\b", run):
        return run
    if re.search(r"<w:rPr\b[^>]*>", run):
        return re.sub(r"(<w:rPr\b[^>]*>)", r"\1<w:b/>", run, count=1)
    return re.sub(r"(<w:r\b[^>]*>)", r"\1<w:rPr><w:b/></w:rPr>", run, count=1)


def replace_merge_placeholders(text: str, field_map: dict[str, str]) -> str:
    run_pattern = re.compile(r"<w:r\b[^>]*>.*?</w:r>", re.DOTALL)

    def replace_run(match: re.Match[str]) -> str:
        run = match.group(0)
        changed = False
        for field, value in field_map.items():
            marker = f"«{field}»"
            if marker in run:
                run = run.replace(marker, xml_escape(str(value or "")))
                changed = True
        return bold_run(model_font_run(run)) if changed else run

    return run_pattern.sub(replace_run, text)


def flatten_merge_fields(text: str) -> str:
    """Remove o código MERGEFIELD depois que seu resultado já foi preenchido.

    Os campos PAGE permanecem intactos para o Word continuar numerando as folhas.
    """
    run_pattern = re.compile(r"<w:r\b[^>]*>.*?</w:r>", re.DOTALL)
    parts: list[str] = []
    inside_merge = False
    cursor = 0
    for match in run_pattern.finditer(text):
        parts.append(text[cursor:match.start()])
        run = match.group(0)
        is_merge_instruction = bool(re.search(r"<w:instrText\b[^>]*>.*?MERGEFIELD", run, re.DOTALL))
        has_begin = bool(re.search(r"<w:fldChar\b[^>]*w:fldCharType=\"begin\"", run))
        has_separate = bool(re.search(r"<w:fldChar\b[^>]*w:fldCharType=\"separate\"", run))
        has_end = bool(re.search(r"<w:fldChar\b[^>]*w:fldCharType=\"end\"", run))
        if is_merge_instruction:
            while parts and parts[-1] == "":
                parts.pop()
            if parts and re.search(r"<w:fldChar\b[^>]*w:fldCharType=\"begin\"", parts[-1]):
                parts.pop()
            inside_merge = True
            cursor = match.end()
            continue
        if inside_merge and (has_separate or has_end):
            if has_end:
                inside_merge = False
            cursor = match.end()
            continue
        if has_begin:
            parts.append(run)
        else:
            parts.append(run)
        cursor = match.end()
    parts.append(text[cursor:])
    return "".join(parts)


def replace_yellow_segments(text: str, values: list[str]) -> str:
    """Substitui, em ordem, grupos consecutivos de runs amarelos.

    O modelo grifado não possui nomes técnicos nos campos amarelos. A ordem
    dos grupos foi registrada no mapeamento do modelo Pagamento; segmentos sem
    valor continuam com o texto original, mas perdem o realce na saída.
    """
    run_pattern = re.compile(r"<w:r\b[^>]*>.*?</w:r>", re.DOTALL)
    runs = list(run_pattern.finditer(text))
    if not runs:
        return text
    pieces = []
    cursor = 0
    yellow_group: list[str] = []
    yellow_indexes: list[int] = []
    group_index = 0

    def flush_group() -> None:
        nonlocal group_index, yellow_group, yellow_indexes
        if not yellow_indexes:
            return
        replacement = values[group_index] if group_index < len(values) else ""
        group_index += 1
        if replacement:
            for pos, run_text in enumerate(yellow_group):
                if pos == 0:
                    run_text = re.sub(r"(<w:t\b[^>]*>).*?(</w:t>)", lambda m: m.group(1) + xml_escape(replacement) + m.group(2), run_text, count=1, flags=re.DOTALL)
                    run_text = bold_run(model_font_run(run_text))
                else:
                    run_text = re.sub(r"(<w:t\b[^>]*>).*?(</w:t>)", r"\1\2", run_text, count=1, flags=re.DOTALL)
                pieces[yellow_indexes[pos]] = run_text
        yellow_group = []
        yellow_indexes = []

    for index, match in enumerate(runs):
        between = text[cursor:match.start()]
        if yellow_indexes and re.search(r"</w:(?:p|tc|tr)>", between):
            flush_group()
        pieces.append(between)
        run = match.group(0)
        is_yellow = bool(re.search(r"<w:highlight\b[^>]*w:val=\"yellow\"", run))
        if is_yellow:
            yellow_group.append(run)
            yellow_indexes.append(len(pieces))
            pieces.append(run)
        else:
            flush_group()
            pieces.append(run)
        cursor = match.end()
    flush_group()
    pieces.append(text[cursor:])
    result = "".join(pieces)
    return re.sub(r"<w:highlight\b[^>]*/>", "", result)


def set_page_numbering(text: str, start_page: int) -> str:
    """Configura a numeração Word para começar após as páginas já existentes."""
    page_type = f'<w:pgNumType w:start="{start_page}"/>'
    if re.search(r"<w:pgNumType\b[^>]*/>", text):
        text = re.sub(r"<w:pgNumType\b[^>]*/>", page_type, text)
    elif "</w:sectPr>" in text:
        text = text.replace("</w:sectPr>", page_type + "</w:sectPr>", 1)
    return text


def set_cached_page_field(text: str, start_page: int) -> str:
    pattern = re.compile(
        r'(<w:instrText\b[^>]*>\s*PAGE\b.*?</w:instrText>.*?'
        r'<w:fldChar\b[^>]*w:fldCharType="separate"[^>]*/>.*?'
        r'<w:t\b[^>]*>)[0-9]+(</w:t>)',
        re.DOTALL,
    )
    return pattern.sub(rf"\g<1>{start_page}\g<2>", text)


def add_header_bottom_border(text: str) -> str:
    # A moldura superior e a lateral direita pertencem à página. No header,
    # manter somente a divisão antes do quadro do processo e a linha inferior
    # contínua, conforme o modelo de referência.
    cells = list(re.finditer(r"<w:tc\b[^>]*>.*?</w:tc>", text, flags=re.DOTALL))
    if len(cells) < 3:
        return text
    desired = {
        0: ["bottom"],
        1: ["right", "bottom"],
        2: ["bottom"],
    }
    output = []
    cursor = 0
    for index, match in enumerate(cells[:3]):
        block = match.group(0)
        tcpr = re.search(r"<w:tcPr\b[^>]*>.*?</w:tcPr>", block, flags=re.DOTALL)
        if not tcpr:
            output.append(text[cursor:match.start()] + block)
            cursor = match.end()
            continue
        props = tcpr.group(0)
        border = '<w:tcBorders>' + "".join(
            f'<w:{side} w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
            for side in desired[index]
        ) + '</w:tcBorders>'
        if re.search(r"<w:tcBorders\b[^>]*>.*?</w:tcBorders>", props, flags=re.DOTALL):
            props = re.sub(r"<w:tcBorders\b[^>]*>.*?</w:tcBorders>", border, props, count=1, flags=re.DOTALL)
        else:
            props = props.replace("</w:tcPr>", border + "</w:tcPr>", 1)
        block = block[:tcpr.start()] + props + block[tcpr.end():]
        output.append(text[cursor:match.start()] + block)
        cursor = match.end()
    output.append(text[cursor:])
    return "".join(output)


def force_heading_page_break(text: str, heading: str) -> str:
    pattern = re.compile(r"<w:p\b[^>]*>(?:(?!<w:p\b).)*?" + re.escape(heading) + r"(?:(?!</w:p>).)*?</w:p>", re.DOTALL)
    def add_break(match: re.Match[str]) -> str:
        paragraph = match.group(0)
        if "w:pageBreakBefore" in paragraph:
            return paragraph
        if "<w:pPr" in paragraph:
            return re.sub(r"(<w:pPr\b[^>]*>)", r"\1<w:pageBreakBefore/>", paragraph, count=1)
        return paragraph.replace("<w:r", '<w:pPr><w:pageBreakBefore/></w:pPr><w:r', 1)
    return pattern.sub(add_break, text, count=1)


def generate_docx(row: sqlite3.Row, output_dir: Path = OUTPUT_DIR,
                  template: Path | io.BytesIO = TEMPLATE_PATH) -> Path:
    if (row["tipo_processo"] or "").strip().casefold() not in {"pagamento"}:
        raise ValueError("Este tipo de processo ainda não possui um modelo Word configurado.")
    if isinstance(template, Path) and not template.is_file():
        raise FileNotFoundError(f"Modelo não encontrado: {template}")
    output_dir.mkdir(exist_ok=True)
    safe_name = re.sub(r'[\\/:*?"<>|]+', "-", f"Processo {row['num_processo']} - Relatório {row['num_relatorio']} - {row['razao_social']}").strip()
    target = output_dir / f"{safe_name}.docx"
    field_map = {
        "Nº_Processo": row["num_processo"], "Data_Início": row["data_inicio"], "Volumes": row["volumes"],
        "Deao_Artista": row["deao_artista"], "Nome_Evento": row["nome_evento"], "Razão_Social": row["razao_social"],
        "CNPJ": row["cnpj"], "Nº_Relatório": row["num_relatorio"], "Nº_Contrato": row["num_contrato"],
        "Nota_Fiscal": row["nota_fiscal"], "Valor": row["valor"], "Valor_Extenso": row["valor_extenso"],
        "Nº_Páginas": row["num_paginas"], "Data_Ass_Contrato": row["data_ass_contrato"],
        "Pgs_Cópia_Contrato": row["pgs_copia_contrato"], "Nº_Processo_Contratação": row["num_processo_contratacao"],
        "Data_Emissão_NF": row["data_emissao_nf"], "Nº_Ordem_de_Serviço": row["num_ordem_servico"],
    }
    with zipfile.ZipFile(template, "r") as source, zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as dest:
        for info in source.infolist():
            data = source.read(info.filename)
            if info.filename.endswith(".xml"):
                text = data.decode("utf-8")
                if info.filename == "[Content_Types].xml":
                    text = text.replace(
                        "application/vnd.ms-word.template.macroEnabledTemplate.main+xml",
                        "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
                    )
                text = replace_merge_placeholders(text, field_map)
                if info.filename == "word/document.xml":
                    yellow_values = [
                        str(row["data_relato"] or ""), str(row["num_empenho"] or ""), str(row["folhas_empenho"] or ""),
                        str(row["folhas_termo_referencia"] or ""), "13", str(row["folha_memoria_calculo"] or ""),
                        str(row["folha_autorizo"] or ""), str(row["folhas_extrato_contrato"] or ""),
                        str(row["folhas_nota_fiscal"] or ""), str(row["folha_requisicao_pagamento"] or ""),
                        str(row["folhas_relatorio_fiscalizacao"] or ""),
                        str(row["cert_trabalhista_emissao"] or ""), str(row["cert_trabalhista_validade"] or ""), str(row["cert_trabalhista_folha"] or ""),
                        str(row["cert_federal_emissao"] or ""), str(row["cert_federal_validade"] or ""), str(row["cert_federal_folha"] or ""),
                        str(row["cert_fgts_emissao"] or ""), str(row["cert_fgts_validade"] or ""), str(row["cert_fgts_folha"] or ""),
                        str(row["folha_ordem_servico"] or ""), str(row["folha_relatorio_entrada"] or ""),
                        str(row["folha_mapa_controle"] or ""), str(row["folha_portaria"] or ""), str(row["num_portaria"] or ""),
                        str(row["folhas_tce"] or ""),
                    ]
                    text = replace_yellow_segments(text, yellow_values)
                elif info.filename in {"word/header1.xml", "word/header2.xml"}:
                    start_page = int(re.sub(r"\D", "", str(row["num_paginas"] or "0")) or "0") + 1
                    text = set_cached_page_field(text, start_page)
                    text = add_header_bottom_border(text)
                if info.filename == "word/document.xml":
                    start_page = int(re.sub(r"\D", "", str(row["num_paginas"] or "0")) or "0") + 1
                    text = set_page_numbering(text, start_page)
                    text = force_heading_page_break(text, "XVI – ENCAMINHAMENTO")
                text = flatten_merge_fields(text)
                text = re.sub(r"<w:highlight\b[^>]*/>", "", text)
                data = text.encode("utf-8")
            dest.writestr(info, data)
    return target


def convert_to_pdf(docx_path: Path) -> Path | None:
    """Converte localmente usando o Word instalado; retorna None se indisponível."""
    pdf_path = docx_path.with_suffix(".pdf")
    doc_json = json.dumps(str(docx_path), ensure_ascii=False)
    pdf_json = json.dumps(str(pdf_path), ensure_ascii=False)
    script = f"""
$doc = {doc_json}
$pdf = {pdf_json}
$word = $null
$document = $null
try {{
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $document = $word.Documents.Open($doc, $false, $true)
    $document.ExportAsFixedFormat($pdf, 17)
    $document.Close($false)
    $word.Quit()
}} catch {{
    if ($document) {{ $document.Close($false) }}
    if ($word) {{ $word.Quit() }}
    exit 1
}}
"""
    try:
        result = subprocess.run(
            ["powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script],
            capture_output=True,
            timeout=90,
            check=False,
        )
        return pdf_path if result.returncode == 0 and pdf_path.exists() else None
    except (OSError, subprocess.TimeoutExpired):
        return None


def download_header(filename: str) -> str:
    return "attachment; filename*=UTF-8''" + quote(filename, safe="")


class Routes:
    hosted = False

    def send_error_page(self, exc: Exception) -> None:
        if self.hosted:
            self.send_bytes(layout("Erro", "<section class='card'>Não foi possível concluir a operação.</section>"), status=500)
        else:
            self.send_bytes(layout("Erro", f"<section class='card'><pre>{esc(exc)}</pre></section>"), status=500)

    def generate_report(self, row) -> None:
        if self.hosted:
            with db() as conn:
                model = conn.execute("SELECT conteudo FROM modelos WHERE tipo_processo=?", ("pagamento",)).fetchone()
            if model is None:
                raise FileNotFoundError("Modelo Pagamento indisponível no banco.")
            with tempfile.TemporaryDirectory(prefix="mare-") as folder:
                target = generate_docx(row, Path(folder), io.BytesIO(model["conteudo"]))
                self.send_report(target, target.name, row["id"])
        else:
            target = generate_docx(row)
            self.send_report(convert_to_pdf(target) or target, str(target), row["id"])

    def send_report(self, report: Path, stored_name: str, ident: int) -> None:
        data = report.read_bytes()
        with db() as conn:
            conn.execute("UPDATE processos SET status='DOCUMENTO_GERADO', arquivo_gerado=?, atualizado_em=? WHERE id=?",
                         (stored_name, datetime.now().isoformat(timespec="seconds"), ident))
        content_type = ("application/pdf" if report.suffix == ".pdf" else
                        "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        self.send_bytes(data, content_type, download=report.name)

    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        query = parse_qs(parsed.query)
        try:
            if parsed.path == "/":
                self.send_bytes(index_page(query.get("notice", [""])[0]))
            elif parsed.path in {"/static/mare-logo.png", "/static/marica-background.png"}:
                asset = ROOT / "static" / parsed.path.removeprefix("/static/")
                if not asset.exists():
                    self.send_bytes(b"Not found", "text/plain; charset=utf-8", 404)
                else:
                    self.send_bytes(asset.read_bytes(), "image/png")
            elif parsed.path == "/manifest.json":
                manifest = {
                    "name": "Relatórios de Pagamento",
                    "short_name": "Relatórios",
                    "start_url": "/",
                    "display": "standalone",
                    "background_color": "#f4f7fb",
                    "theme_color": "#123d5a",
                    "description": "Cadastro e geração local de relatórios administrativos.",
                }
                self.send_bytes(json.dumps(manifest, ensure_ascii=False).encode("utf-8"), "application/manifest+json; charset=utf-8")
            elif parsed.path == "/sw.js":
                sw = """self.addEventListener('install', event => event.waitUntil(self.skipWaiting()));
self.addEventListener('activate', event => event.waitUntil(
  caches.delete('relatorios-mvp-v1').then(() => self.clients.claim())
));
"""
                self.send_bytes(sw.encode("utf-8"), "application/javascript; charset=utf-8")
            elif parsed.path == "/novo":
                self.send_bytes(form_page(None))
            elif parsed.path == "/editar":
                ident = int(query.get("id", [0])[0])
                with db() as conn:
                    row = conn.execute("SELECT * FROM processos WHERE id=?", (ident,)).fetchone()
                self.send_bytes(form_page(row, "Processo não encontrado." if row is None else ""))
            elif parsed.path == "/apontamentos":
                processo_id = int(query.get("processo_id", [0])[0])
                with db() as conn:
                    processo = conn.execute("SELECT * FROM processos WHERE id=?", (processo_id,)).fetchone()
                self.send_bytes(apontamentos_page(processo, "Processo não encontrado." if processo is None else ""))
            elif parsed.path == "/novo-apontamento":
                processo_id = int(query.get("processo_id", [0])[0])
                with db() as conn:
                    processo = conn.execute("SELECT * FROM processos WHERE id=?", (processo_id,)).fetchone()
                self.send_bytes(apontamento_form_page(None, processo))
            elif parsed.path == "/editar-apontamento":
                ident = int(query.get("id", [0])[0])
                with db() as conn:
                    item = conn.execute("SELECT * FROM apontamentos WHERE id=?", (ident,)).fetchone()
                    processo = conn.execute("SELECT p.* FROM processos p JOIN apontamentos a ON a.processo_id=p.id WHERE a.id=?", (ident,)).fetchone()
                self.send_bytes(apontamento_form_page(item, processo, "Apontamento não encontrado." if item is None else ""))
            elif parsed.path == "/gerar":
                ident = int(query.get("id", [0])[0])
                with db() as conn:
                    row = conn.execute("SELECT * FROM processos WHERE id=?", (ident,)).fetchone()
                if row is None:
                    self.send_bytes(index_page("Processo não encontrado."), status=404)
                    return
                self.generate_report(row)
            else:
                self.send_bytes(b"Not found", "text/plain; charset=utf-8", 404)
        except Exception as exc:
            self.send_error_page(exc)

    def do_POST(self) -> None:
        parsed = urlparse(self.path)
        length = int(self.headers.get("Content-Length", "0"))
        raw = self.rfile.read(length).decode("utf-8")
        submitted = {key: values[0] for key, values in parse_qs(raw, keep_blank_values=True).items()}
        try:
            if parsed.path == "/salvar":
                erros = validacao(submitted)
                if erros:
                    self.send_bytes(form_page(None, " ".join(erros)))
                    return
                now = datetime.now().isoformat(timespec="seconds")
                ident = submitted.get("id", "").strip()
                with db() as conn:
                    if ident:
                        assignments = ", ".join(f"{key}=?" for key in ALL_KEYS) + ", atualizado_em=?"
                        conn.execute(f"UPDATE processos SET {assignments} WHERE id=?", [submitted.get(key, "").strip() for key in ALL_KEYS] + [now, int(ident)])
                    else:
                        columns = ", ".join(ALL_KEYS + ["criado_em", "atualizado_em"])
                        placeholders = ", ".join("?" for _ in ALL_KEYS) + ", ?, ?"
                        conn.execute(f"INSERT INTO processos ({columns}) VALUES ({placeholders})", [submitted.get(key, "").strip() for key in ALL_KEYS] + [now, now])
                self.redirect("/?notice=" + quote_plus("Processo salvo com sucesso."))
            elif parsed.path == "/salvar-apontamento":
                processo_id = int(submitted.get("processo_id", "0"))
                erros = [f"Preencha: {label}." for key, label in APONTAMENTO_FIELDS if not submitted.get(key, "").strip()]
                with db() as conn:
                    processo = conn.execute("SELECT * FROM processos WHERE id=?", (processo_id,)).fetchone()
                    item = None
                    if submitted.get("id", "").strip():
                        item = conn.execute("SELECT * FROM apontamentos WHERE id=?", (int(submitted["id"]),)).fetchone()
                if processo is None:
                    erros.append("Processo não encontrado.")
                if erros:
                    self.send_bytes(apontamento_form_page(item, processo, " ".join(erros)))
                    return
                now = datetime.now().isoformat(timespec="seconds")
                with db() as conn:
                    status = submitted.get("status", "").strip()
                    conn.execute("INSERT INTO status_apontamentos (nome, criado_em) VALUES (?, ?) ON CONFLICT (nome) DO NOTHING", (status, now))
                    if submitted.get("id", "").strip():
                        conn.execute("UPDATE apontamentos SET sequencia=?, pagina=?, apontamento=?, status=?, atualizado_em=? WHERE id=?", [submitted.get(key, "").strip() for key, _label in APONTAMENTO_FIELDS] + [now, int(submitted["id"])])
                    else:
                        conn.execute("INSERT INTO apontamentos (processo_id, sequencia, pagina, apontamento, status, criado_em, atualizado_em) VALUES (?, ?, ?, ?, ?, ?, ?)", [processo_id] + [submitted.get(key, "").strip() for key, _label in APONTAMENTO_FIELDS] + [now, now])
                self.redirect(f"/apontamentos?processo_id={processo_id}")
            elif parsed.path == "/excluir-apontamento":
                ident = int(submitted.get("id", "0"))
                processo_id = int(submitted.get("processo_id", "0"))
                with db() as conn:
                    conn.execute("DELETE FROM apontamentos WHERE id=?", (ident,))
                self.redirect(f"/apontamentos?processo_id={processo_id}")
            elif parsed.path == "/excluir":
                ident = int(submitted.get("id", "0"))
                with db() as conn:
                    conn.execute("UPDATE processos SET status='INATIVO', atualizado_em=? WHERE id=?", (datetime.now().isoformat(timespec="seconds"), ident))
                self.redirect("/")
            else:
                self.send_bytes(b"Not found", "text/plain; charset=utf-8", 404)
        except Exception as exc:
            self.send_error_page(exc)


class Handler(Routes, BaseHTTPRequestHandler):
    def send_bytes(self, data: bytes, content_type: str = "text/html; charset=utf-8", status: int = 200, download: str = "") -> None:
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Cache-Control", "no-store")
        if download:
            self.send_header("Content-Disposition", download_header(download))
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def redirect(self, path: str) -> None:
        self.send_response(303)
        self.send_header("Cache-Control", "no-store")
        self.send_header("Location", path)
        self.end_headers()


if __name__ == "__main__":
    init_db()
    OUTPUT_DIR.mkdir(exist_ok=True)
    server = ThreadingHTTPServer(("0.0.0.0", 8000), Handler)
    print("Relatórios: http://127.0.0.1:8000")
    print("Para outros computadores: http://IP-DESTA-MAQUINA:8000")
    print(f"Modelo: {TEMPLATE_PATH}")
    server.serve_forever()
