# GPON Notepad Web

Aplicação local (Flask) para registrar tratativas de atendimento GPON, com:
- Formulário de OS (protocolo, circuito, cliente, serial)
- Campo de tratativa
- Salvamento em Excel (.xlsx)
- Tela de histórico com filtros
- Timer de tratativa e fluxo de contato (em evolução)

## Como rodar
1. Instalar dependências:
   python -m pip install -r requirements.txt

2. Rodar:
   python app.py

3. Abrir:
   http://127.0.0.1:5000