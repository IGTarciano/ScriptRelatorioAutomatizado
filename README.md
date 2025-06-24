# ScriptRelatorioAutomatizado
Este script em Python é uma ferramenta para automatizar a geração de relatórios de emergência a partir de um arquivo de dados do Excel, focando especificamente em pedidos de um fornecedor principal. Ele processa os dados, filtra informações relevantes e organiza os resultados em diferentes abas de uma planilha Excel, facilitando a análise.
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

## Licença

Este projeto está licenciado sob a licença MIT - veja o arquivo [LICENSE](LICENSE) para detalhes.
# Funcionalidades Principais

## Leitura de Dados: O script inicia o processo lendo um arquivo Excel (.xlsx) que deve ser carregado pelo usuário.
Cálculo de Valor Total: Automaticamente calcula o Valor_Total de cada item do pedido, multiplicando a quantidade pela o valor total do item.

## Filtragem por Fornecedor: Filtra os dados para incluir apenas os registros da 'fornecedor específico'.

## Seleção e Renomeação de Colunas: Seleciona um conjunto específico de colunas para análise e as renomeia para termos mais compreensíveis e padronizados no relatório final (ex: 'PedidoCompra_EmpresaCod' para 'Filial').

## Mapeamento de Filiais: Converte os códigos numéricos das filiais para nomes descritivos, tornando o relatório mais legível.

## Classificação de Itens (ABC): Categoriza os itens com base em sua popularidade (Classificação ABC). Ele gera duas tabelas distintas:
  ### Itens A, B e C: Resumo dos itens de alta e média popularidade.
      Itens D, E e Inativos: Resumo dos itens de baixa popularidade ou inativos.
      Filtragem de Tipo de Pedido: Exclui pedidos do tipo 'R - REPOSICAO - TDB' do relatório analítico principal.
      Tratamento de Dados Ausentes: Preenche valores ausentes na coluna 'Classificação ABC' com 'S/class' e remove linhas onde o 'Cod. Montadora' está             vazio, garantindo a integridade dos dados.
## Geração de Relatório em Excel: Salva os resultados processados em um único arquivo Excel com três abas:
    Resumo_emergencia_Analitico: Contém a base de dados completa e processada.
    Resumo_emergencia_A_B_C: Resumo dos itens classificados como A, B e C.
    Resumo_emergencia_D_E_Inativos: Resumo dos itens classificados como D, E e Inativos.
    Tratamento de Erros: Inclui blocos try-except para lidar com potenciais erros durante o upload, processamento e salvamento dos dados, oferecendo         
    mensagens claras em caso de falha.

# Como Usar

## bnPreparação: Certifique-se de que o ambiente (preferencialmente Google Colab, devido ao uso de google.colab.files e google.colab.drive) tenha as bibliotecas openpyxl, pandas e numpy instaladas. O script já inclui a linha !pip install openpyxl --quiet para isso.

## Upload do Arquivo: Ao executar o script, ele solicitará o upload de um arquivo Excel contendo os dados de pedidos.

## Execução: O script processará automaticamente os dados.

## Acesso ao Relatório: O relatório final será salvo no seu Google Drive (em /content/drive/MyDrive/Acompanhamento/Resumo_emergencia.xlsx) e uma mensagem de sucesso será exibida com o caminho do arquivo.

# Autor

Igor Tarciano Gomes da Silva

www.linkedin.com/in/igor-tarciano-gomes

