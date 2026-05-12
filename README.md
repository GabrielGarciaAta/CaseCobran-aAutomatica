# Case Cobrança Automática
# 📧 Automação de Cobranças com Python e Outlook

![Python](https://img.shields.io/badge/python-3670A0?style=for-the-badge&logo=python&logoColor=ffdd54)
![Pandas](https://img.shields.io/badge/pandas-%23150458.svg?style=for-the-badge&logo=pandas&logoColor=white)
![Microsoft Outlook](https://img.shields.io/badge/Microsoft_Outlook-0078D4?style=for-the-badge&logo=microsoft-outlook&logoColor=white)

Este projeto automatiza o processo de verificação de inadimplência e envio de notificações de cobrança. O script processa uma base de dados em Excel, identifica clientes com pagamentos atrasados e dispara e-mails personalizados automaticamente através do Microsoft Outlook.

## 🚀 Funcionalidades

* **Processamento de Dados:** Realiza a leitura e análise de planilhas de contas a receber.
* **Filtro de Inadimplência:** Identifica automaticamente registros com status "Em aberto" onde a data de vencimento é inferior à data atual.
* **Comunicação Automatizada:** Gera e envia e-mails detalhando a Nota Fiscal (NF), data de vencimento e o valor total do débito.
* **Gestão de Contas:** Permite a seleção de uma conta específica configurada no Outlook para o envio das mensagens.

## 🛠️ Tecnologias Utilizadas

* **Python:** Linguagem principal do projeto.
* **Pandas:** Utilizada para manipulação e filtragem dos dados do Excel.
* **pywin32 (win32com.client):** Responsável pela integração e controle do aplicativo Microsoft Outlook.
* **Datetime:** Utilizada para a gestão de datas e identificação de atrasos.

## 📋 Pré-requisitos

Para o correto funcionamento do script, são necessários:

1.  **Microsoft Outlook:** Versão Desktop (Clássica) instalada e configurada.
2.  **Base de Dados:** Um arquivo Excel chamado `Contas a Receber.xlsx` contendo as seguintes colunas:
    * `CPF`
    * `Valor em aberto`
    * `Data Prevista para pagamento`
    * `Status`
    * `E-mail`
    * `NF`

### Instalação

```bash
pip install pywin32 pandas openpyxl