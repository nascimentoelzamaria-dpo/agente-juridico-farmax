import os
import json
import uuid
import re
from datetime import datetime
from pathlib import Path
from flask import Flask, request, jsonify, send_file

app = Flask(__name__)
app.secret_key = os.urandom(24)

BASE_DIR = Path(__file__).parent

ADENDO_RT_LGPD = """
=== ADENDO TRANSVERSAL: REFORMA TRIBUTARIA E LGPD ===

REFORMA TRIBUTARIA (LC 214/2025):
- Contratos com preco fixo e vigencia > 12 meses DEVEM ter clausula de reajuste tributario contemplando IBS/CBS.
- Contratos com vigencia > 24 meses DEVEM ter clausula de equilibrio economico-financeiro.
- Tipos mais expostos: Fornecimento, Prestacao de Servicos, Empreitada, Transporte.
- DESVIO VERMELHO se: preco fixo + vigencia > 12 meses + sem clausula de reajuste tributario.
- DESVIO AMARELO se: vigencia > 12 meses + reajuste so por indice sem mencionar variacao tributaria.

LGPD (Lei 13.709/2018):
- Sempre verificar se o contrato envolve tratamento de dados pessoais.
- Se a contratada acessa dados pessoais da Farmax: clausula de OPERADOR obrigatoria.
- Se ambas as partes trocam dados: clausula de CONTROLADORES INDEPENDENTES.
- DESVIO VERMELHO se: contrato envolve dados pessoais + sem qualquer clausula LGPD.
- DESVIO AMARELO se: clausula LGPD presente mas incompleta.
- Prazo padrao de NDA: 5 anos. Inferior a 5 anos = DESVIO VERMELHO.
"""

PLAYBOOK_MAP = {
    "fornecimento": "pb_fornecimento.txt",
    "prestacao-de-servicos": "pb_prestacao-de-servicos.txt",
    "fornecimento-e-prestacao-de-servicos": "pb_fornecimento-e-prestacao-de-servicos.txt",
    "empreitada": "pb_empreitada.txt",
    "transporte": "pb_prestacao-de-servicos-de-transporte.txt",
    "parceria": "pb_parceria.txt",
    "comodato": "pb_comodato.txt",
    "autorizacao-de-uso-de-imagem": "pb_autorizacao-de-uso-de-imagem.txt",
    "procuracao": "pb_procuracao.txt",
    "termo-aditivo": "pb_termo-aditivo.txt",
    "nda": "pb_termo-de-confidencialidade.txt",
}

PLAYBOOK_LABELS = {
    "fornecimento": "Fornecimento",
    "prestacao-de-servicos": "Prestacao de Servicos",
    "fornecimento-e-prestacao-de-servicos": "Fornecimento e Prestacao de Servicos",
    "empreitada": "Empreitada",
    "transporte": "Prestacao de Servicos de Transporte",
    "parceria": "Parceria",
    "comodato": "Comodato",
    "autorizacao-de-uso-de-imagem": "Autorizacao de Uso de Imagem",
    "procuracao": "Procuracao",
    "termo-aditivo": "Termo Aditivo",
    "nda": "Termo de Confidencialidade / NDA",
}

def load_playbook(key):
    fname = PLAYBOOK_MAP.get(key)
    if not fname:
        return ""
    path = BASE_DIR / fname
    try:
        return path.read_text(encoding="utf-8") if path.exists() else ""
    except:
        return ""

def extract_text_from_docx(file_bytes):
    import io
    try:
        from docx import Document
        doc = Document(io.BytesIO(file_bytes))
        return "\n".join(p.text for p in doc.paragraphs if p.text.strip())
    except Exception as e:
        return f"[Erro ao ler DOCX: {e}]"

def extract_text_from_pdf(file_bytes):
    import io
    try:
        import pdfplumber
        parts = []
        with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
            for page in pdf.pages:
                t = page.extract_text()
                if t:
                    parts.append(t)
        return "\n".join(parts)
    except Exception as e:
        return f"[Erro ao ler PDF: {e}]"

def call_claude(api_key, system_prompt, user_prompt):
    import urllib.request
    payload = json.dumps({
        "model": "claude-sonnet-4-20250514",
        "max_tokens": 4000,
        "system": system_prompt,
        "messages": [{"role": "user", "content": user_prompt}]
    }).encode("utf-8")
    req = urllib.request.Request(
        "https://api.anthropic.com/v1/messages",
        data=payload,
        headers={
            "Content-Type": "application/json",
            "x-api-key": api_key,
            "anthropic-version": "2023-06-01"
        },
        method="POST"
    )
    with urllib.request.urlopen(req, timeout=120) as resp:
        data = json.loads(resp.read())
    return data["content"][0]["text"]

def build_system_prompt(playbook_text):
    return f"""Voce e um agente juridico especializado em revisao de contratos da empresa Farmax.
Sua funcao e analisar contratos recebidos de terceiros e identificar desvios em relacao ao padrao da Farmax.

PLAYBOOK DE REFERENCIA:
{playbook_text}

{ADENDO_RT_LGPD}

REGRAS:
1. Compare o contrato com o padrao do playbook.
2. Classifique cada desvio:
   - vermelho: risco alto
   - amarelo: atencao
   - verde: cosmetico/baixo impacto
3. Verifique SEMPRE clausula de Reforma Tributaria e LGPD.
4. NUNCA sugira aceite automatico de clausula divergente.

RESPONDA EXCLUSIVAMENTE em JSON valido, sem texto fora do JSON:
{{
  "tipo_identificado": "nome do tipo contratual",
  "confianca_classificacao": "alta|media|baixa",
  "resumo_executivo": "2-3 frases resumindo o contrato e principais riscos",
  "partes": {{"contratante": "nome", "contratada": "nome"}},
  "objeto": "descricao do objeto",
  "valor": "valor ou NAO IDENTIFICADO",
  "vigencia": "prazo ou NAO IDENTIFICADO",
  "desvios": [
    {{
      "clausula": "nome da clausula",
      "severidade": "vermelho|amarelo|verde",
      "descricao": "o que esta diferente do padrao",
      "risco": "impacto pratico se mantido",
      "recomendacao": "o que fazer"
    }}
  ],
  "verificacao_rt": {{
    "aplicavel": true,
    "clausula_presente": false,
    "observacao": "texto"
  }},
  "verificacao_lgpd": {{
    "dados_pessoais_identificados": true,
    "clausula_presente": false,
    "tipo_clausula_necessaria": "operador|controladores_independentes|nao_aplicavel",
    "observacao": "texto"
  }},
  "campos_faltantes": ["lista de campos nao preenchidos"],
  "recomendacao_final": "APROVADO_COM_RESSALVAS|NEGOCIAR|ESCALAR_JURIDICO",
  "justificativa_recomendacao": "explicacao da recomendacao final"
}}"""

@app.route("/")
def index():
    html_path = BASE_DIR / "index_static.html"
    if html_path.exists():
        return send_file(str(html_path))
    return "<h1>Sistema Agente Juridico Farmax</h1><p>Arquivo de interface nao encontrado.</p>", 500

@app.route("/analisar", methods=["POST"])
def analisar():
    api_key = request.form.get("api_key", "").strip()
    playbook_key = request.form.get("playbook_type", "").strip()
    auto_detect = request.form.get("auto_detect", "false") == "true"
    arquivo = request.files.get("contrato")

    if not api_key:
        return jsonify({"erro": "Chave de API nao informada."}), 400
    if not arquivo:
        return jsonify({"erro": "Nenhum arquivo enviado."}), 400

    filename = arquivo.filename.lower()
    file_bytes = arquivo.read()

    if filename.endswith(".docx"):
        texto = extract_text_from_docx(file_bytes)
    elif filename.endswith(".pdf"):
        texto = extract_text_from_pdf(file_bytes)
    else:
        return jsonify({"erro": "Formato nao suportado. Use PDF ou DOCX."}), 400

    if not texto.strip() or texto.startswith("[Erro"):
        return jsonify({"erro": f"Nao foi possivel extrair texto. {texto}"}), 400

    if auto_detect or not playbook_key:
        playbook_text = ""
        for key in PLAYBOOK_MAP:
            pb = load_playbook(key)
            if pb:
                playbook_text += f"\n\n=== PLAYBOOK: {PLAYBOOK_LABELS[key]} ===\n{pb[:2000]}"
    else:
        playbook_text = load_playbook(playbook_key)
        if not playbook_text:
            return jsonify({"erro": "Playbook nao encontrado."}), 400

    system_prompt = build_system_prompt(playbook_text)
    user_prompt = f"""Analise o contrato abaixo e gere o relatorio.

ARQUIVO: {arquivo.filename}
DATA: {datetime.now().strftime('%d/%m/%Y %H:%M')}

CONTRATO:
{texto[:12000]}

{"(Identifique automaticamente o tipo contratual)" if auto_detect else f"(Tipo indicado: {PLAYBOOK_LABELS.get(playbook_key, playbook_key)})"}
"""

    try:
        resposta_raw = call_claude(api_key, system_prompt, user_prompt)
        resposta_raw = re.sub(r'^```json\s*', '', resposta_raw.strip())
        resposta_raw = re.sub(r'\s*```$', '', resposta_raw.strip())
        resultado = json.loads(resposta_raw)
    except json.JSONDecodeError as e:
        return jsonify({"erro": f"Erro ao interpretar resposta da IA: {str(e)}"}), 500
    except Exception as e:
        err = str(e)
        if "401" in err or "authentication" in err.lower():
            return jsonify({"erro": "Chave de API invalida."}), 401
        if "429" in err:
            return jsonify({"erro": "Muitas requisicoes. Aguarde alguns segundos."}), 429
        return jsonify({"erro": f"Erro na analise: {err}"}), 500

    resultado["arquivo"] = arquivo.filename
    resultado["data_analise"] = datetime.now().strftime('%d/%m/%Y as %H:%M')
    resultado["id_analise"] = str(uuid.uuid4())[:8].upper()
    return jsonify(resultado)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
