# GestorMan

Aplicação desktop em Python/Tkinter para gestão de manutenção.

## Recursos

- Cadastro e consulta de Ordens de Serviço, armazenadas em SQLite.
- Cadastro de almoxarifado, com carga de arquivos TXT.
- Carga de Ordens de Serviço por TXT.
- Filtros de pesquisa para Ordens de Serviço e almoxarifado.

## Como executar

Requer Python com Tkinter e não possui dependências externas.

```powershell
.\.venv\Scripts\python.exe .\app.py
```

Ou execute `app.bat` com um duplo clique. O ambiente virtual `.venv` usa
Python 3.13 e a aplicação não possui dependências externas para instalar.

Os bancos SQLite são criados e atualizados automaticamente ao abrir a aplicação.
