# Como utilizar o `porsche_sanitizer_agent.py`

O script `porsche_sanitizer_agent.py` permite aplicar automaticamente as regras de tratamento definidas no `schema.md` em novas bases de dados.

## Preparar os arquivos

Antes de executar o script, certifique-se de que os seguintes arquivos estejam na **mesma pasta** do seu notebook (`.ipynb`):

* `porsche_sanitizer_agent.py`
* `schema.md`
* `porsche_data_base.xlsx`

Essa organização permite que o script localize corretamente a planilha de entrada e o arquivo de regras de sanitização.

## Como executar

Instale as dependências (necessário apenas na primeira execução):

```bash
pip install pandas openpyxl
```

Em seguida, execute o script em uma célula do Jupyter Notebook:

```python
!python porsche_sanitizer_agent.py --input porsche_data_base.xlsx --schema schema.md --output porsche_data_base_sanitize.xlsx
```

## Resultado

Após a execução, o script irá:

* ler a planilha `porsche_data_base.xlsx`;
* aplicar todas as regras definidas no `schema.md`;
* gerar automaticamente a planilha tratada `porsche_data_base_sanitize.xlsx` na mesma pasta do projeto.
