# GridLamEdit

Aplicativo desktop modular para edição de laminados com interface feita em PySide6.

## Pré-requisitos
- Python 3.10 ou superior

## Instalação
1. (Opcional) Crie e ative um ambiente virtual:
   - Windows: `python -m venv .venv` e depois `.venv\Scripts\activate`
   - Linux/macOS: `python -m venv .venv` e depois `source .venv/bin/activate`
2. Instale as dependências: `pip install -r requirements.txt`

## Execução
Rode a interface inicial com:

```bash
python -m gridlamedit.app.main
```

Uma janela simples chamada **GridLamEdit** será exibida, pronta para ser expandida com novos módulos.

## Testes
1. Instale as dependências de teste (por exemplo, `pip install pytest`).
2. Execute a suíte: `python -m pytest -q`.
