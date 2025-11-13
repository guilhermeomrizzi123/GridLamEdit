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

## Geração do Executável (Windows)
1. (Opcional) Crie/ative um ambiente virtual (`python -m venv .venv` e `.\.venv\Scripts\activate`).
2. Instale as dependências: `pip install -r requirements.txt`.
3. Execute `build_exe.bat` para chamar o PyInstaller com os parâmetros abaixo:

   ```bat
   pyinstaller --noconfirm --clean --onedir --name GridLamEdit --noconsole ^
     --collect-submodules PySide6 --collect-data PySide6 ^
     --add-data "gridlamedit\resources;gridlamedit\resources" ^
     --add-data "Grid_Spreadsheet.xls;." ^
     --add-data "Grid_Spreadsheet_editado_RevA.xlsx;." ^
     --add-data "Grid_Spreadsheet_editado_RevB.xlsx;." ^
     gridlamedit\app\main.py
   ```

   > Ainda não há um arquivo `.ico` oficial. Assim que existir, inclua `--icon caminho\para\icone.ico` no comando.

4. O executável (e DLLs/recursos) ficará em `dist/GridLamEdit/`. Compacte essa pasta para distribuir; o usuário final só precisa extrair tudo e dar duplo clique em `GridLamEdit.exe`.
5. Após o primeiro build, o PyInstaller gera `GridLamEdit.spec`. Para rebuilds avançados, edite esse arquivo e execute `pyinstaller GridLamEdit.spec`.

## Testes
1. Instale as dependências de teste (por exemplo, `pip install pytest`).
2. Execute a suíte: `python -m pytest -q`.
