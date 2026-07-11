# Projeto de AI Report

Este repositório contém o projeto **AI Report**, desenvolvido como parte do curso **"Aceleração: AI Reports com Excel, GPT Agents e Claude Code"** da DIO.

## Sobre o projeto

O objetivo deste projeto é realizar o tratamento de uma base de dados contendo informações de vendas de veículos da Porsche e, a partir desses dados tratados, gerar um dashboard interativo em HTML.

Todo o processo foi desenvolvido com o auxílio de Inteligência Artificial, seguindo a metodologia proposta pelo curso.

---

# Como o projeto foi desenvolvido

O desenvolvimento foi dividido em **três etapas principais**.

## 1ª Etapa — Tratamento da base de dados

Nesta etapa foi criado um agente de IA no ChatGPT responsável pelo tratamento da planilha `porsche_data_base.xlsx`.

Para isso, foi utilizado o arquivo `schema.md`, que contém todas as regras de validação e sanitização dos dados.

Após a execução do agente, foram gerados os seguintes arquivos:

* `porsche_data_base_sanitize.xlsx` — Base de dados tratada.
* `porsche_sanitizer_agent.py` — Script Python responsável por automatizar o processo de tratamento, permitindo sua reutilização em bases maiores.

### Como usar o Script Python
Para ver orientações de uso do arquivo `porsche_sanitizer_agent.py` no Jupiter Notebook, acesse o arquivo `orientação_script_python.md`. 

---

## 2ª Etapa — Validação dos dados

Após o tratamento da base de dados, foi realizada uma conferência manual para validar se todas as regras definidas no `schema.md` haviam sido aplicadas corretamente.

Com a validação concluída:

* Foi criada uma cópia da planilha tratada;
* As colunas contendo os dados originais foram removidas;
* Permaneceram apenas as colunas com os dados já sanitizados.

Como resultado, foi gerado o arquivo:

* `porsche_fordash_sanitize.xlsx`

Essa nova planilha foi utilizada como base para a construção do dashboard.

---

## 3ª Etapa — Desenvolvimento do dashboard

Com a base de dados pronta, foi solicitado ao ChatGPT que gerasse um dashboard em HTML utilizando o recurso de Canvas.

Após a primeira versão, foram realizadas diversas melhorias e ajustes visuais, funcionais e de usabilidade até chegar ao resultado final, representado pelo arquivo:

* `porsche_sales_dashboard.html`

---

# Prompts utilizados

Abaixo estão registrados os principais prompts utilizados durante o desenvolvimento do projeto.

## 1ª Etapa

```text
Crie um agente que leia o arquivo "porsche_data_base.xlsx" e aplique as regras do schema.md para gerar um novo arquivo tratado.
```

---

## 2ª Etapa

Nenhum.

---

## 3ª Etapa

```text
Utilizando o recurso de canvas, renderize uma dashboard em html ao lado. Esse dashboard deve ter filtros interativos e dinâmicos.

Filtros:
- Sale Date
- Model Year
- Pay Method
- State
- City
- Delivery Status

Perguntas de negócios e KPIs:
- Qaul a quantidade de vendas e receitas por estado?
- Qual a quantidade vendida por Model Year?
- Qual o modelo de carro que menos vendeu?
- Qual a quantidade por Delevery Status em cada Estado?
- Quero insights de vendas de carros populares com base nos dados em cada estado e cidade.

Sobre o UI/UX, se baseie no site oficial da porsche brasil:
https://www.porsche.com.brazil/pt/

Faça um design elegante e refinado.
```

# Orientações de Uso do Script Python (Agente)

Para ver orientações de uso do Script Python no Jupiter Notebook, acesse o arquivo `orientação_script_python.md`.
