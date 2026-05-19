"""
Calendário Fiscal Mensal — Shoulder
Envio automático via Gmail API todo dia 1º do mês.

CONFIGURAÇÃO INICIAL:
1. Instale as dependências:
   pip install google-auth google-auth-oauthlib google-auth-httplib2 google-api-python-client

2. Crie um projeto no Google Cloud Console (console.cloud.google.com):
   - Ative a API Gmail
   - Crie credenciais OAuth 2.0 (tipo: Desktop App)
   - Baixe o arquivo credentials.json e coloque na mesma pasta deste script

3. Na primeira execução, uma janela do navegador abrirá para autorização.
   O token é salvo em token.json — as próximas execuções são automáticas.

4. Agende via cron (Linux/Mac):
   crontab -e
   Adicione a linha abaixo para rodar todo dia 1º às 07h00:
   0 7 1 * * /usr/bin/python3 /caminho/para/enviar_calendario.py >> /caminho/log_fiscal.txt 2>&1

   No Windows, use o Agendador de Tarefas para rodar todo mês no dia 1.

5. Variáveis de configuração: ajuste a seção CONFIGURAÇÃO abaixo.
"""

import os
import base64
import logging
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# ─── CONFIGURAÇÃO ────────────────────────────────────────────────────────────

DESTINATARIO = "seu@email.com.br"          # e-mail de destino
REMETENTE    = "me"                         # "me" = conta autenticada no Gmail
EMPRESA      = "Shoulder"
SEGMENTO     = "Indústria e Varejo de Moda · Lucro Real"
CREDENCIAIS  = "credentials.json"          # arquivo baixado do Google Cloud
TOKEN        = "token.json"                # gerado automaticamente na 1ª execução
LOG_FILE     = "log_fiscal.txt"

SCOPES = ["https://www.googleapis.com/auth/gmail.send"]

# ─── DADOS DO CALENDÁRIO ──────────────────────────────────────────────────────

OBRIGACOES = [
    {
        "dia": 20, "nome": "FGTS — FGTS Digital", "tipo": "trab",
        "desc": "Recolhimento do FGTS — competência do mês anterior via FGTS Digital (GFD).",
        "prazo_federal": "Dia 20 do mês subsequente — antecipado para o 1º dia útil anterior se cair em não útil",
        "base_legal": "Lei 8.036/90 · FGTS Digital obrigatório desde 01/03/2024",
        "estados": None,
        "nota": "⚠️ O prazo antigo era dia 7 — foi alterado para dia 20 com o FGTS Digital (março/2024). Recolher via GFD exclusivamente por PIX até 21h59 de Brasília. Não há mais GRRF manual para novas competências.",
    },
    {
        "dia": 15, "nome": "EFD-REINF", "tipo": "fed",
        "desc": "Retenções e Outras Informações Fiscais — serviços e rendimentos.",
        "prazo_federal": "Dia 15 do mês seguinte (postergado para 1º dia útil se cair em não útil)",
        "base_legal": "IN RFB 2.043/2021 · IN RFB 2.133/2023",
        "estados": None,
        "nota": "R-2010, R-2020, R-4010/4020, R-2055. Integra automaticamente com a DCTFWeb.",
    },
    {
        "dia": 20, "nome": "INSS — Contribuição previdenciária", "tipo": "trab",
        "desc": "Contribuição previdenciária patronal e dos segurados — folha de pagamento.",
        "prazo_federal": "Dia 20 do mês seguinte — antecipado para o dia útil anterior se não útil",
        "base_legal": "Art. 30 Lei 8.212/91",
        "estados": None,
        "nota": "Via DARF gerado pela DCTFWeb para empresas no eSocial. Prazo de pagamento (dia 20) difere do prazo de entrega da DCTFWeb (último dia útil do mês).",
    },
    {
        "dia": 20, "nome": "IRRF — Salários e serviços", "tipo": "fed",
        "desc": "IR Retido na Fonte sobre salários, pró-labore e serviços.",
        "prazo_federal": "Dia 20 do mês subsequente — antecipado para o dia útil anterior se não útil",
        "base_legal": "RIR/2018 · IN RFB 1.500/2014",
        "estados": None,
        "nota": "DARF código 0561 (trabalho assalariado) e 1708 (serviços de PJ). O prazo de pagamento (dia 20) é diferente do prazo de entrega da DCTFWeb (último dia útil).",
    },
    {
        "dia": 25, "nome": "PIS / COFINS", "tipo": "fed",
        "desc": "Apuração e recolhimento — regime não cumulativo (Lucro Real).",
        "prazo_federal": "Dia 25 do mês seguinte ao fato gerador — antecipado se não útil",
        "base_legal": "Lei 10.637/02 · Lei 10.833/03 · Art. 18 MP 2.158-35/01 c/ Lei 11.933/09",
        "estados": None,
        "nota": "⚠️ O prazo é dia 25 — não dia 20. DARF: COFINS não cumulativa = código 5856 · PIS não cumulativo = código 6912. Verificar créditos: insumos, energia, aluguéis, depreciação.",
    },
    {
        "dia": 25, "nome": "IPI — Confecções", "tipo": "fed",
        "desc": "Imposto sobre Produtos Industrializados — apuração mensal.",
        "prazo_federal": "Dia 25 do mês subsequente (produtos em geral, inclusive confecções)",
        "base_legal": "RIPI · Decreto 7.212/2010",
        "estados": None,
        "nota": "Verificar NCM. Alíquota geralmente 0% para vestuário, mas obrigação de escrituração no SPED permanece.",
    },
    {
        "dia": 20, "nome": "ICMS — Recolhimento", "tipo": "est",
        "desc": "Apuração e pagamento do ICMS — prazo varia por estado e atividade.",
        "prazo_federal": None,
        "base_legal": None,
        "estados": [
            ("AC", "Dia 9"),  ("AL", "Dia 20"), ("AM", "Dia 9"),  ("AP", "Dia 20"),
            ("BA", "Dia 9/14"), ("CE", "Dia 9"), ("DF", "Dia 20"), ("ES", "Dia 9"),
            ("GO", "Dia 15"), ("MA", "Dia 20"), ("MG", "Dia 15"), ("MS", "Dia 20"),
            ("MT", "Dia 6/20"), ("PA", "Dia 20"), ("PB", "Dia 15"), ("PE", "Dia 15/20"),
            ("PI", "Dia 20"), ("PR", "Dia 10/20"), ("RJ", "Dia 10/20"), ("RN", "Dia 10"),
            ("RO", "Dia 20"), ("RR", "Dia 20"), ("RS", "Dia 12"), ("SC", "Dia 10/20"),
            ("SE", "Dia 20"), ("SP", "Dia 22"), ("TO", "Dia 20"),
        ],
        "nota": "SP: dia 22. RS: dia 12. BA: dia 9 comércio / 14 indústria. Verificar feriados estaduais.",
    },
    {
        "dia": 20, "nome": "GIA / DAPI / DIME", "tipo": "est",
        "desc": "Declaração mensal de ICMS — obrigação acessória estadual.",
        "prazo_federal": None,
        "base_legal": None,
        "estados": [
            ("SP", "Dia 20 — GIA"), ("RS", "Dia 20 — GIA/RS"), ("SC", "Dia 10 — DIME"),
            ("MG", "Dia 20 — DAPI"), ("PR", "Dia 20 — GIA-ICMS"), ("RJ", "Dia 20 — DIEF"),
            ("BA", "Dia 20 — GIA-BA"), ("Demais", "Até dia 20"),
        ],
        "nota": "Muitos estados já migraram para o SPED como obrigação principal.",
    },
    {
        "dia": 25, "nome": "EFD ICMS/IPI — SPED Fiscal", "tipo": "est",
        "desc": "Escrituração Fiscal Digital — livros fiscais ICMS e IPI.",
        "prazo_federal": "Dia 15 do mês seguinte (prazo federal padrão)",
        "base_legal": "Ajuste SINIEF 2/2009 · IN RFB 2.003/2021",
        "estados": [
            ("SP", "Dia 20"), ("MG", "Dia 15"), ("RS", "Dia 20"), ("RJ", "Dia 20"),
            ("PR", "Dia 15"), ("SC", "Dia 10"), ("BA", "Dia 15"), ("GO", "Dia 15"),
            ("Demais", "Dia 15"),
        ],
        "nota": "Validar no PVA antes da transmissão. Leiaute: Ato COTEPE/ICMS 44/2018.",
    },
    {
        "dia": 25, "nome": "EFD Contribuições — SPED PIS/COFINS", "tipo": "fed",
        "desc": "Escrituração Fiscal Digital das Contribuições Sociais.",
        "prazo_federal": "Até o 10º dia útil do 2º mês subsequente",
        "base_legal": "IN RFB 1.252/2012",
        "estados": None,
        "nota": "Atenção ao M100 (créditos PIS) e M500 (créditos COFINS).",
    },
    {
        "dia": "último dia útil", "nome": "IRPJ / CSLL — Estimativa mensal", "tipo": "fed",
        "desc": "Recolhimento por estimativa ou suspensão/redução com balanço — Lucro Real.",
        "prazo_federal": "Último dia útil do mês subsequente ao mês de apuração",
        "base_legal": "Art. 4º Lei 9.430/96 · IN RFB 1.700/2017",
        "estados": None,
        "nota": "⚠️ Prazo é o último dia útil — não dia 25. DARF 2362 (IRPJ estimativa) e 2430 (CSLL estimativa). Manter LALUR e LACS atualizados mensalmente.",
    },
    {
        "dia": "último dia útil", "nome": "DCTFWeb + MIT — Entrega mensal", "tipo": "fed",
        "desc": "Declaração unificada — extinguiu a DCTF PGD a partir de jan/2025. Consolida IRPJ, CSLL, IPI, PIS, COFINS, IRRF.",
        "prazo_federal": "Último dia útil do mês seguinte ao fato gerador",
        "base_legal": "IN RFB 2.237/2024 · IN RFB 2.248/2025",
        "estados": None,
        "nota": "⚠️ Mudança 2025: DCTF PGD extinta. Tributos declarados via MIT dentro da DCTFWeb. Pagamento dos tributos mantém datas próprias (dia 20 ou 25).",
    },
    {
        "dia": 20, "nome": "eSocial — Fechamento de folha", "tipo": "trab",
        "desc": "Eventos S-1200, S-1210 e S-1299 — remuneração e fechamento.",
        "prazo_federal": "S-1299: até dia 20 do mês seguinte (alinhado ao FGTS Digital)",
        "base_legal": "Manual eSocial v.S-1.2",
        "estados": None,
        "nota": "Sequência: S-1200 → S-1210 → S-1299. DCTFWeb e FGTS Digital gerados após S-1299. S-2299 (desligamento): 10 dias corridos após o desligamento.",
    },
    {
        "dia": 30, "nome": "ICMS-ST — Substituição Tributária", "tipo": "est",
        "desc": "Recolhimento do ICMS retido por ST nas operações interestaduais.",
        "prazo_federal": None,
        "base_legal": None,
        "estados": [
            ("SP", "Dia 22"), ("MG", "Dia 15"), ("RJ", "Dia 10"), ("RS", "Dia 10"),
            ("PR", "Dia 9"),  ("SC", "Dia 10"), ("BA", "Dia 9"),  ("GO", "Dia 15"),
            ("MT", "Dia 6"),  ("MS", "Dia 9"),  ("PE", "Dia 9"),  ("CE", "Dia 10"),
            ("Demais", "GNRE por operação"),
        ],
        "nota": "Verificar Convênio ICMS 92/15 e protocolos por NCM. EC 132/2023: revisar transferências entre filiais.",
    },
    {
        "dia": 30, "nome": "DIFAL — E-commerce e filiais", "tipo": "est",
        "desc": "Diferencial de alíquota ICMS nas vendas a consumidor final não contribuinte.",
        "prazo_federal": None,
        "base_legal": None,
        "estados": [
            ("SP", "Dia 15 mensal"), ("MG", "Dia 15 via GNRE"), ("RJ", "Por operação ou mensal"),
            ("RS", "Dia 20 mensal"), ("PR", "GNRE por operação"), ("SC", "GNRE por operação"),
            ("Demais", "GNRE por operação"),
        ],
        "nota": "LC 190/2022: inscrição estadual no destino ou GNRE. Verificar partilha FEF.",
    },
    {
        "dia": 30, "nome": "MIT — Importação/Exportação (Siscomex)", "tipo": "fed",
        "desc": "Licença de Importação, DI, DUIMP e demais declarações aduaneiras.",
        "prazo_federal": "Conforme tipo: LI (antes do embarque), DI (desembaraço)",
        "base_legal": "IN RFB 1.639/2016 · IN RFB 680/2006",
        "estados": None,
        "nota": "Verificar Drawback, Ex-Tarifário e NCM de insumos têxteis.",
    },
    {
        "dia": 30, "nome": "Fechamento contábil — CPC/IFRS", "tipo": "soc",
        "desc": "Balancete, conciliações, provisões e ajustes de acordo com CPCs e IFRS.",
        "prazo_federal": "Prazo interno — até último dia útil do mês",
        "base_legal": "CPC 00 · CPC 06 · CPC 16 · CPC 47 · IFRS 15 · IFRS 16",
        "estados": None,
        "nota": "Crítico para moda: CPC 06 (aluguéis de lojas), CPC 16 (estoques), CPC 47 (fidelidade e devoluções).",
    },
]

# ─── CORES POR TIPO ───────────────────────────────────────────────────────────

CORES = {
    "fed":  {"bg": "#E6F1FB", "txt": "#0C447C", "brd": "#85B7EB", "label": "Federal"},
    "est":  {"bg": "#FAECE7", "txt": "#712B13", "brd": "#F0997B", "label": "Estadual"},
    "trab": {"bg": "#EAF3DE", "txt": "#27500A", "brd": "#97C459", "label": "Trabalhista"},
    "soc":  {"bg": "#EEEDFE", "txt": "#3C3489", "brd": "#AFA9EC", "label": "Societária"},
}

# ─── GERAÇÃO DO HTML DO E-MAIL ────────────────────────────────────────────────

def gerar_html_email(mes: int, ano: int) -> str:
    mes_nome = [
        "Janeiro","Fevereiro","Março","Abril","Maio","Junho",
        "Julho","Agosto","Setembro","Outubro","Novembro","Dezembro"
    ][mes]
    hoje_fmt = datetime.now().strftime("%d/%m/%Y")

    linhas = ""
    for o in OBRIGACOES:
        c = CORES[o["tipo"]]
        badge = (
            f'<span style="background:{c["bg"]};color:{c["txt"]};'
            f'font-size:10px;padding:2px 8px;border-radius:99px;'
            f'border:1px solid {c["brd"]};white-space:nowrap;">'
            f'{c["label"]}</span>'
        )
        # Prazos por estado
        estados_html = ""
        if o["estados"]:
            rows_uf = "".join(
                f'<tr>'
                f'<td style="padding:4px 8px;font-weight:500;font-size:11px;color:#333;white-space:nowrap;">{uf}</td>'
                f'<td style="padding:4px 8px;font-size:11px;color:#564F19;font-weight:500;">{dia}</td>'
                f'</tr>'
                for uf, dia in o["estados"]
            )
            estados_html = f"""
            <table style="margin-top:8px;border-collapse:collapse;width:100%;max-width:500px;">
              <tr style="background:#f5f4f0;">
                <th style="padding:4px 8px;font-size:10px;color:#888;text-align:left;font-weight:500;">Estado</th>
                <th style="padding:4px 8px;font-size:10px;color:#888;text-align:left;font-weight:500;">Prazo</th>
              </tr>
              {rows_uf}
            </table>"""

        prazo_html = ""
        if o["prazo_federal"]:
            prazo_html = (
                f'<div style="margin-top:6px;background:#f5f4f0;border-radius:6px;'
                f'padding:6px 10px;font-size:11px;">'
                f'<strong>Prazo:</strong> {o["prazo_federal"]}'
                + (f' &nbsp;·&nbsp; <span style="color:#888;">{o["base_legal"]}</span>' if o["base_legal"] else "")
                + f'</div>'
            )

        nota_html = ""
        if o["nota"]:
            nota_html = (
                f'<div style="margin-top:8px;background:#FAEEDA;border:1px solid #EF9F27;'
                f'border-radius:6px;padding:6px 10px;font-size:11px;color:#633806;">'
                f'&#9432; {o["nota"]}</div>'
            )

        linhas += f"""
        <tr>
          <td style="padding:14px 16px;border-bottom:1px solid #eee;vertical-align:top;min-width:60px;">
            <span style="background:#000;color:#fff;font-size:11px;padding:3px 8px;
                         border-radius:99px;white-space:nowrap;font-weight:500;">
              Dia {o["dia"]}
            </span>
          </td>
          <td style="padding:14px 16px;border-bottom:1px solid #eee;vertical-align:top;">
            {badge}
          </td>
          <td style="padding:14px 16px;border-bottom:1px solid #eee;vertical-align:top;">
            <div style="font-weight:500;font-size:13px;color:#111;">{o["nome"]}</div>
            <div style="font-size:12px;color:#666;margin-top:3px;">{o["desc"]}</div>
            {prazo_html}
            {estados_html}
            {nota_html}
          </td>
        </tr>"""

    return f"""<!DOCTYPE html>
<html lang="pt-BR">
<head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1.0"></head>
<body style="margin:0;padding:0;background:#f7f5f1;font-family:'Helvetica Neue',Helvetica,Arial,sans-serif;">
<table width="100%" cellpadding="0" cellspacing="0" style="background:#f7f5f1;padding:24px 0;">
  <tr><td align="center">
    <table width="680" cellpadding="0" cellspacing="0" style="max-width:680px;width:100%;background:#fff;border-radius:12px;overflow:hidden;border:1px solid #e2ddd6;">

      <!-- Header -->
      <tr>
        <td style="background:#000;padding:24px 28px;">
          <div style="font-size:22px;letter-spacing:.12em;font-weight:300;color:#fff;">shoulder</div>
          <div style="font-size:12px;color:#C49863;margin-top:6px;">{SEGMENTO}</div>
          <div style="font-size:18px;font-weight:300;color:#fff;margin-top:12px;">
            Calendário Fiscal — {mes_nome} {ano}
          </div>
          <div style="font-size:11px;color:#888;margin-top:4px;">Gerado em {hoje_fmt}</div>
        </td>
      </tr>

      <!-- Intro -->
      <tr>
        <td style="padding:20px 28px 8px;">
          <p style="font-size:13px;color:#666;margin:0;line-height:1.6;">
            Obrigações fiscais, tributárias e societárias do mês de <strong>{mes_nome} {ano}</strong>.
            Verifique antecipações por feriados estaduais e federais antes dos vencimentos.
          </p>
        </td>
      </tr>

      <!-- Legenda -->
      <tr>
        <td style="padding:12px 28px;">
          <span style="background:#E6F1FB;color:#0C447C;font-size:10px;padding:3px 8px;border-radius:99px;border:1px solid #85B7EB;">Federal</span>
          &nbsp;
          <span style="background:#FAECE7;color:#712B13;font-size:10px;padding:3px 8px;border-radius:99px;border:1px solid #F0997B;">Estadual</span>
          &nbsp;
          <span style="background:#EAF3DE;color:#27500A;font-size:10px;padding:3px 8px;border-radius:99px;border:1px solid #97C459;">Trabalhista</span>
          &nbsp;
          <span style="background:#EEEDFE;color:#3C3489;font-size:10px;padding:3px 8px;border-radius:99px;border:1px solid #AFA9EC;">Societária</span>
        </td>
      </tr>

      <!-- Tabela de obrigações -->
      <tr>
        <td style="padding:0 28px 8px;">
          <table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;">
            <tr style="background:#f5f4f0;">
              <th style="padding:8px 16px;font-size:10px;color:#888;text-align:left;font-weight:500;width:70px;">Prazo</th>
              <th style="padding:8px 16px;font-size:10px;color:#888;text-align:left;font-weight:500;width:100px;">Tipo</th>
              <th style="padding:8px 16px;font-size:10px;color:#888;text-align:left;font-weight:500;">Obrigação e detalhes</th>
            </tr>
            {linhas}
          </table>
        </td>
      </tr>

      <!-- Aviso legal -->
      <tr>
        <td style="padding:16px 28px;border-top:1px solid #eee;">
          <p style="font-size:11px;color:#aaa;margin:0;line-height:1.6;">
            ⚠️ Prazos sujeitos a alteração por decreto ou feriados locais.
            Consulte sempre a SEFAZ de cada estado e o Diário Oficial antes dos vencimentos.
            Legislação de referência: 2025/2026.
          </p>
        </td>
      </tr>

      <!-- Footer -->
      <tr>
        <td style="background:#f5f4f0;padding:14px 28px;text-align:center;">
          <p style="font-size:11px;color:#aaa;margin:0;">
            <strong style="color:#C49863;">shoulder</strong> · Fiscal &amp; Contabilidade · Lucro Real ·
            Enviado automaticamente em {hoje_fmt}
          </p>
        </td>
      </tr>

    </table>
  </td></tr>
</table>
</body>
</html>"""


def gerar_texto_email(mes: int, ano: int) -> str:
    mes_nome = [
        "Janeiro","Fevereiro","Março","Abril","Maio","Junho",
        "Julho","Agosto","Setembro","Outubro","Novembro","Dezembro"
    ][mes]
    linhas = [
        f"CALENDÁRIO FISCAL — {mes_nome.upper()} {ano}",
        f"{EMPRESA} · {SEGMENTO}",
        f"Gerado em {datetime.now().strftime('%d/%m/%Y')}",
        "=" * 60,
        "",
    ]
    for o in OBRIGACOES:
        tipo = CORES[o["tipo"]]["label"]
        linhas.append(f"Dia {o['dia']:2d} | {tipo:<14} | {o['nome']}")
        linhas.append(f"          {o['desc']}")
        if o["prazo_federal"]:
            linhas.append(f"          Prazo: {o['prazo_federal']}")
        if o["estados"]:
            for uf, dia in o["estados"][:5]:
                linhas.append(f"          {uf}: {dia}")
            if len(o["estados"]) > 5:
                linhas.append(f"          ... e mais {len(o['estados'])-5} estados")
        if o["nota"]:
            linhas.append(f"          Atenção: {o['nota']}")
        linhas.append("")
    linhas.append("─" * 60)
    linhas.append("Prazos sujeitos a alteração. Consulte SEFAZ e DOU.")
    return "\n".join(linhas)


# ─── AUTENTICAÇÃO GMAIL ───────────────────────────────────────────────────────

def autenticar_gmail():
    creds = None
    if os.path.exists(TOKEN):
        creds = Credentials.from_authorized_user_file(TOKEN, SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CREDENCIAIS, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(TOKEN, "w") as token:
            token.write(creds.to_json())
    return build("gmail", "v1", credentials=creds)


# ─── ENVIO DO E-MAIL ──────────────────────────────────────────────────────────

def enviar_email(service, mes: int, ano: int):
    mes_nome = [
        "Janeiro","Fevereiro","Março","Abril","Maio","Junho",
        "Julho","Agosto","Setembro","Outubro","Novembro","Dezembro"
    ][mes]
    assunto = f"Calendário Fiscal — {mes_nome} {ano} · {EMPRESA}"

    msg = MIMEMultipart("alternative")
    msg["Subject"] = assunto
    msg["From"]    = REMETENTE
    msg["To"]      = DESTINATARIO

    msg.attach(MIMEText(gerar_texto_email(mes, ano), "plain", "utf-8"))
    msg.attach(MIMEText(gerar_html_email(mes, ano),  "html",  "utf-8"))

    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
    service.users().messages().send(userId="me", body={"raw": raw}).execute()
    return assunto


# ─── EXECUÇÃO PRINCIPAL ───────────────────────────────────────────────────────

def main():
    logging.basicConfig(
        filename=LOG_FILE,
        level=logging.INFO,
        format="%(asctime)s — %(levelname)s — %(message)s",
        datefmt="%d/%m/%Y %H:%M:%S",
    )

    hoje = datetime.now()
    mes  = hoje.month - 1   # mês de referência = mês atual (ou ajuste conforme preferir)
    ano  = hoje.year
    if mes < 0:
        mes = 11
        ano -= 1

    logging.info(f"Iniciando envio do calendário fiscal para {DESTINATARIO}…")

    try:
        service = autenticar_gmail()
        assunto = enviar_email(service, mes, ano)
        logging.info(f"E-mail enviado com sucesso: '{assunto}'")
        print(f"✓ Calendário fiscal enviado para {DESTINATARIO}")
    except HttpError as e:
        logging.error(f"Erro na API do Gmail: {e}")
        print(f"✗ Erro Gmail API: {e}")
    except FileNotFoundError:
        logging.error(f"Arquivo {CREDENCIAIS} não encontrado.")
        print(f"✗ Coloque o arquivo '{CREDENCIAIS}' na mesma pasta do script.")
    except Exception as e:
        logging.error(f"Erro inesperado: {e}")
        print(f"✗ Erro: {e}")


if __name__ == "__main__":
    main()
