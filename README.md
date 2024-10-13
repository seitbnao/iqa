# Função IQA para Excel

Este repositório contém uma função VBA (`IQA`) que calcula o Índice de Qualidade da Água (IQA) com base em vários parâmetros físicos e químicos da água. A função também inclui uma classificação do IQA usando a função `ClassificaIQA`.

## Tabela de Conteúdos

- [Descrição da Função](#descrição-da-função)
- [Parâmetros Considerados](#parâmetros-considerados)
- [Instalação](#instalação)
- [Como Usar](#como-usar)
- [Exemplo de Uso](#exemplo-de-uso)
- [Análise dos Resultados](#análise-dos-resultados)
- [Contribuições](#contribuições)
- [Licença](#licença)

## Descrição da Função

A função `IQA` calcula um índice baseado em várias características da água, fornecendo uma avaliação da qualidade da água em um formato numérico que pode ser facilmente interpretado.

### Função ClassificaIQA

A função `ClassificaIQA` classifica o valor de IQA em cinco categorias:

![image](https://github.com/user-attachments/assets/0953033b-b694-4faa-bcb8-0233f2a108d4)

## Parâmetros Considerados

A função IQA leva em consideração os seguintes parâmetros:

- **Oxigênio Dissolvido (mg/L)**
- **Coliformes Fecais (NMP/100mL)**
- **pH**
- **DBO (Demanda Biológica de Oxigênio)**
- **Nitrato (mg/L)**
- **Fosfato (mg/L)**
- **Temperatura (°C)**
- **Turbidez (NTU)**
- **Sólidos Totais (mg/L)**
- **Altitude (m)**

Cada parâmetro é ponderado de acordo com seu impacto na qualidade da água, e o resultado final é um valor que pode ser classificado em categorias de qualidade.

## Instalação

1. Abra o Excel.
2. Pressione `ALT + F11` para abrir o Editor do VBA.
3. No menu, clique em `Inserir` > `Módulo` para criar um novo módulo.
4. Cole o código da função `IQA`, `ClassificaIQA` e o subprocedimento `TestarIQA` no módulo.
5. Feche o Editor do VBA e retorne ao Excel.

## Como Usar

Para utilizar a função `IQA`, siga estas etapas no Excel:

### Passo 1: Inserir os Dados

1. **Abra o Excel** e crie uma nova planilha ou utilize uma já existente.
2. **Insira os seguintes dados** em sua planilha, começando na linha 2:

| B                     | C                         | D   | E   | F                | G                | H                  | I              | J                     | K           | L             | M             | N                   |
|-----------------------|---------------------------|-----|-----|------------------|------------------|--------------------|----------------|-----------------------|-------------|---------------|-----------------|---------------------|
| Oxigênio (mg/L)      | Coliformes (NMP/100mL)   | pH  | DBO | Nitrato (mg/L)   | Fosfato (mg/L)   | Temperatura (°C)   | Turbidez (NTU) | Sólidos Totais (mg/L) | Altitude (m) | Tipo de Fosfato | IQA             | CLASSIFICAÇÃO      |
| 6.5                   | 30                        | 7.2 | 3   | 10               | 1.5              | 20                 | 200            | 150                   | 500         | fosforo      | =IQA(B2:L2)    | =ClassificaIQA(M2) |

### Passo 2: Usar a Função IQA

1. **Selecione a célula** onde deseja calcular o IQA (neste exemplo, a célula M2).
2. **Insira a fórmula** que chama a função `IQA` para o intervalo de dados:

   ```excel
   =IQA(B2:L2)

### Exemplo Real
![image](https://github.com/user-attachments/assets/8864cdd6-0d7a-46c1-9b7e-7fdecf330432)

### Exemplo Exibindo as fórmulas
![image](https://github.com/user-attachments/assets/dfb9b771-5a18-4745-8820-9c7f9e18b433)

### Material de apoio
https://www.cetesb.sp.gov.br/aguas-interiores/wp-content/uploads/sites/12/2013/11/02.pdf

