# Relatório OS 60+ — app de mesa

Interface gráfica para gerar os relatórios de OS/Aprovadas por gerenciadora,
sem precisar de Python instalado no computador de quem vai usar.

- `report_engine.py` — a regra de negócio original (`codigo_final_60_.py`), sem alterações.
- `app.py` — a janela (Tkinter), os seletores de arquivo e as otimizações de
  leitura/escrita de Excel (as mesmas usadas na versão web: leitor calamine,
  escrita "preguiçosa" de linhas grandes).
- `.github/workflows/build.yml` — compila um `RelatorioOS60.exe` (Windows,
  arquivo único, sem console) via GitHub Actions a cada push na `main`, e
  publica o `.exe` como artefato e como release.

## Rodar sem compilar (com Python instalado)

```
pip install -r requirements.txt
python app.py
```

## Gerar o .exe manualmente (Windows, com Python instalado)

```
pip install -r requirements.txt pyinstaller
pyinstaller --onefile --noconsole --name "RelatorioOS60" --collect-all python_calamine app.py
```

O executável fica em `dist/RelatorioOS60.exe` — copie esse único arquivo para
qualquer PC Windows; não precisa instalar Python nele.
