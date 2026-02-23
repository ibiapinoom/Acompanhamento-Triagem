from __future__ import annotations  # Permite usar tipos como "list[str]" em versões modernas sem dor.

from pathlib import Path            # Manipular caminhos/pastas de forma segura (Windows/Linux).
from datetime import datetime       # Pegar data/hora atual.
from typing import List, Dict, Any, Optional  # Tipos para deixar o código mais claro.

from flask import Flask, render_template, request, jsonify  # Framework web + render HTML + receber dados + retornar JSON.
from openpyxl import Workbook, load_workbook               # Criar e editar arquivos Excel .xlsx.

APP_DIR = Path(__file__).resolve().parent  # Pasta onde este arquivo app.py está.
DATA_DIR = APP_DIR / "data"                # Pasta /data (vamos criar ela para guardar o Excel).
EXCEL_PATH = DATA_DIR / "atendimentos.xlsx"  # Caminho completo do Excel.
SHEET_NAME = "Registros"                   # Nome da aba dentro do Excel.

app = Flask(__name__)                      # Cria a aplicação web Flask.


def ensure_excel_file() -> None:
    """Garante que existe um Excel e que ele tem cabeçalho e aba correta."""
    DATA_DIR.mkdir(parents=True, exist_ok=True)  # Cria a pasta data/ se não existir.

    headers = ["Data/Hora", "PROTOCOLO", "CIRCUITO", "CLIENTE", "SERIAL", "TRATATIVA"]  # Cabeçalho das colunas.

    if EXCEL_PATH.exists():               # Se o arquivo Excel já existe...
        wb = load_workbook(EXCEL_PATH)    # Abre o Excel existente.
        if SHEET_NAME not in wb.sheetnames:  # Se a aba "Registros" não existir...
            ws = wb.create_sheet(SHEET_NAME)  # Cria a aba.
            ws.append(headers)                # Coloca o cabeçalho na primeira linha.
            wb.save(EXCEL_PATH)               # Salva o arquivo.
            return                            # Sai da função.

        ws = wb[SHEET_NAME]               # Pega a aba Registros.
        if ws.max_row == 0:               # Se por algum motivo a aba está vazia...
            ws.append(headers)            # Insere o cabeçalho.
            wb.save(EXCEL_PATH)           # Salva.
        return                            # Sai da função.

    wb = Workbook()                       # Se o arquivo não existe, cria um Excel novo.
    ws = wb.active                        # Pega a primeira aba criada automaticamente.
    ws.title = SHEET_NAME                 # Renomeia essa aba para "Registros".
    ws.append(headers)                    # Escreve o cabeçalho.
    wb.save(EXCEL_PATH)                   # Salva o Excel no disco.


def append_row(protocolo: str, circuito: str, cliente: str, serial: str, tratativa: str) -> None:
    """Adiciona uma nova linha no Excel com os campos preenchidos."""
    wb = load_workbook(EXCEL_PATH)        # Abre o Excel existente.
    ws = wb[SHEET_NAME]                   # Seleciona a aba "Registros".
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")  # Gera data/hora no padrão definido.
    ws.append([timestamp, protocolo, circuito, cliente, serial, tratativa])  # Adiciona uma linha nova.
    wb.save(EXCEL_PATH)                   # Salva o arquivo para persistir a linha.


def _parse_dt(s: str) -> Optional[datetime]:
    """Tenta converter uma string 'YYYY-MM-DD HH:MM:SS' em datetime. Se falhar, retorna None."""
    s = (s or "").strip()                 # Garante que não é None e remove espaços.
    if not s:                             # Se ficou vazio...
        return None                       # Retorna None.
    try:
        return datetime.strptime(s, "%Y-%m-%d %H:%M:%S")  # Converte string para datetime.
    except ValueError:
        return None                       # Se não deu certo, retorna None.


def read_all_records() -> List[Dict[str, Any]]:
    """Lê o Excel e devolve uma lista de registros (cada registro é um dicionário)."""
    ensure_excel_file()                   # Garante que o Excel existe antes de ler.

    wb = load_workbook(EXCEL_PATH, data_only=True)  # Abre o Excel (data_only evita fórmulas).
    ws = wb[SHEET_NAME]                   # Seleciona a aba.

    records: List[Dict[str, Any]] = []    # Lista onde vamos guardar os registros.

    for idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):  # Lê linha por linha, começando na 2 (pula cabeçalho).
        if not row or all(v is None or str(v).strip() == "" for v in row):  # Se linha vazia...
            continue                      # Pula.

        dt_raw, protocolo, circuito, cliente, serial, tratativa = row  # Desempacota colunas.

        if isinstance(dt_raw, datetime):  # Se a data veio como datetime...
            dt = dt_raw                   # Usa direto.
        else:
            dt = _parse_dt(str(dt_raw)) if dt_raw is not None else None  # Tenta converter se veio como string.

        records.append({                  # Adiciona um dicionário com os dados.
            "row_index": idx,             # Adiciona o índice da linha no Excel
            "data_hora": dt.strftime("%Y-%m-%d %H:%M:%S") if dt else (str(dt_raw) if dt_raw else ""),
            "dia": dt.strftime("%Y-%m-%d") if dt else "",
            "protocolo": str(protocolo or "").strip(),
            "circuito": str(circuito or "").strip(),
            "cliente": str(cliente or "").strip(),
            "serial": str(serial or "").strip(),
            "tratativa": str(tratativa or "").strip(),
        })

    records.sort(key=lambda r: r["data_hora"], reverse=True)  # Ordena do mais recente para o mais antigo.
    return records                        # Retorna a lista completa.


def filter_records(records: List[Dict[str, Any]],
                   date_from: str,
                   date_to: str,
                   protocolo: str,
                   cliente: str) -> List[Dict[str, Any]]:
    """Filtra a lista por intervalo de datas, protocolo e cliente (contains)."""
    df = (date_from or "").strip()        # Data inicial (YYYY-MM-DD).
    dt = (date_to or "").strip()          # Data final (YYYY-MM-DD).
    p = (protocolo or "").strip().lower() # Protocolo para filtro (minúsculo).
    c = (cliente or "").strip().lower()   # Cliente para filtro (minúsculo).

    dfrom = datetime.strptime(df, "%Y-%m-%d").date() if df else None  # Converte data inicial.
    dto = datetime.strptime(dt, "%Y-%m-%d").date() if dt else None    # Converte data final.

    out: List[Dict[str, Any]] = []        # Lista filtrada.

    for r in records:                      # Percorre todos os registros.
        if r["dia"]:                       # Se tem dia calculado...
            rd = datetime.strptime(r["dia"], "%Y-%m-%d").date()  # Converte em date.
        else:
            rd = None                      # Se não tem dia, marca como None.

        if dfrom and (rd is None or rd < dfrom):  # Se existe data inicial e registro é menor...
            continue                       # Pula.
        if dto and (rd is None or rd > dto):      # Se existe data final e registro é maior...
            continue                       # Pula.

        if p and p not in r["protocolo"].lower():  # Se filtro protocolo existe e não bate...
            continue                       # Pula.

        if c and c not in r["cliente"].lower():    # Se filtro cliente existe e não bate...
            continue                       # Pula.

        out.append(r)                      # Se passou tudo, inclui no resultado.

    return out                             # Retorna registros filtrados.


def count_by_day(records: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Conta quantas tratativas existem por dia (produtividade)."""
    counter: Dict[str, int] = {}          # Dicionário: dia -> quantidade.

    for r in records:                      # Percorre registros filtrados.
        day = r.get("dia") or ""          # Pega o dia.
        if not day:                        # Se não tem dia...
            continue                       # Pula.
        counter[day] = counter.get(day, 0) + 1  # Soma 1 na contagem do dia.

    days_sorted = sorted(counter.items(), key=lambda x: x[0], reverse=True)  # Ordena por dia desc.
    return [{"dia": d, "qtd": n} for d, n in days_sorted]  # Converte em lista para usar no HTML.


@app.get("/")                              # Rota principal.
def home():
    ensure_excel_file()                    # Garante Excel.
    
    # Buscar registros de hoje
    all_records = read_all_records()
    today = datetime.now().strftime("%d/%m/%Y")
    today_iso = datetime.now().strftime("%Y-%m-%d")  # Para comparação no banco
    today_records = [r for r in all_records if r.get("dia") == today_iso]
    
    return render_template("index.html", 
                         excel_path=str(EXCEL_PATH),
                         today_records=today_records,
                         today_date=today)


@app.post("/save")                         # Endpoint que recebe o salvar via JS.
def save():
    ensure_excel_file()                    # Garante Excel.

    payload = request.get_json(silent=True) or {}  # Lê o JSON do navegador.
    protocolo = (payload.get("protocolo") or "").strip()  # Extrai protocolo.
    circuito = (payload.get("circuito") or "").strip()    # Extrai circuito.
    cliente = (payload.get("cliente") or "").strip()      # Extrai cliente.
    serial = (payload.get("serial") or "").strip()        # Extrai serial.
    tratativa = (payload.get("tratativa") or "").strip()  # Extrai tratativa.

    if not tratativa:                       # Se tratativa estiver vazia...
        return jsonify({"ok": False, "error": "O campo TRATATIVA está vazio."}), 400  # Retorna erro.

    append_row(protocolo, circuito, cliente, serial, tratativa)  # Salva no Excel.
    return jsonify({"ok": True, "message": "Registro salvo no Excel com sucesso."})  # Responde OK.


@app.post("/update")                       # Endpoint para atualizar registro.
def update_record():
    ensure_excel_file()
    
    payload = request.get_json(silent=True) or {}
    row_index = payload.get("row_index")  # Índice da linha no Excel (começando em 2)
    
    if row_index is None:
        return jsonify({"ok": False, "error": "Índice da linha não informado."}), 400
    
    protocolo = (payload.get("protocolo") or "").strip()
    circuito = (payload.get("circuito") or "").strip()
    cliente = (payload.get("cliente") or "").strip()
    serial = (payload.get("serial") or "").strip()
    tratativa = (payload.get("tratativa") or "").strip()
    
    if not tratativa:
        return jsonify({"ok": False, "error": "O campo TRATATIVA está vazio."}), 400
    
    try:
        wb = load_workbook(EXCEL_PATH)
        ws = wb[SHEET_NAME]
        
        # Atualiza a linha (mantém a data/hora original)
        ws.cell(row=row_index, column=2).value = protocolo
        ws.cell(row=row_index, column=3).value = circuito
        ws.cell(row=row_index, column=4).value = cliente
        ws.cell(row=row_index, column=5).value = serial
        ws.cell(row=row_index, column=6).value = tratativa
        
        wb.save(EXCEL_PATH)
        return jsonify({"ok": True, "message": "Registro atualizado com sucesso."})
    except Exception as e:
        return jsonify({"ok": False, "error": f"Erro ao atualizar: {str(e)}"}), 500


@app.post("/delete")                       # Endpoint para deletar registro.
def delete_record():
    ensure_excel_file()
    
    payload = request.get_json(silent=True) or {}
    row_index = payload.get("row_index")
    
    if row_index is None:
        return jsonify({"ok": False, "error": "Índice da linha não informado."}), 400
    
    try:
        wb = load_workbook(EXCEL_PATH)
        ws = wb[SHEET_NAME]
        
        # Deleta a linha
        ws.delete_rows(row_index, 1)
        
        wb.save(EXCEL_PATH)
        return jsonify({"ok": True, "message": "Registro deletado com sucesso."})
    except Exception as e:
        return jsonify({"ok": False, "error": f"Erro ao deletar: {str(e)}"}), 500


@app.get("/today-records")                 # Endpoint para buscar registros de hoje (AJAX).
def get_today_records():
    ensure_excel_file()
    
    all_records = read_all_records()
    today = datetime.now().strftime("%Y-%m-%d")
    today_records = [r for r in all_records if r.get("dia") == today]
    
    return jsonify({"ok": True, "records": today_records})


@app.get("/historico")                     # Página de histórico.
def historico():
    ensure_excel_file()                    # Garante Excel.

    all_records = read_all_records()       # Lê todos os registros do Excel.

    # Pega filtros da URL (querystring): /historico?date_from=...&date_to=...&protocolo=...&cliente=...
    date_from = request.args.get("date_from", "")
    date_to = request.args.get("date_to", "")
    protocolo = request.args.get("protocolo", "")
    cliente = request.args.get("cliente", "")

    filtered = filter_records(all_records, date_from, date_to, protocolo, cliente)  # Aplica filtros.
    counts = count_by_day(filtered)        # Conta produtividade por dia.

    total = len(filtered)                  # Total de registros após o filtro.

    return render_template(                # Renderiza a página de histórico.
        "historico.html",
        excel_path=str(EXCEL_PATH),
        records=filtered,
        counts=counts,
        total=total,
        filters={                          # Envia filtros para manter preenchidos no formulário.
            "date_from": date_from,
            "date_to": date_to,
            "protocolo": protocolo,
            "cliente": cliente,
        }
    )


if __name__ == "__main__":                 # Executa só se você rodar "python app.py".
    ensure_excel_file()                    # Garante que o Excel existe antes de iniciar.
    app.run(host="127.0.0.1", port=5000, debug=True)  # Sobe servidor local em http://127.0.0.1:5000
