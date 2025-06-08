# Geração de PDFs em Lote - Mala Direta

## 📜 Descrição Geral
Solução de automação desenvolvida em VBA para otimizar o fluxo de trabalho no envio de Mala Direta. Transforma um processo manual e lento numa tarefa rápida e automatizada, gerando centenas de arquivos PDF personalizados e agrupados por cliente em questão de minutos. Além da geração dos arquivos, a solução pode ser estendida para automatizar completamente a comunicação, preparando e enviando os documentos por e-mail via Outlook.

---

## 🎯 O Problema Resolvido
O processo manual de gerar relatórios individualizados para múltiplos clientes é um grande consumidor de tempo e um foco de erros:
- Filtrar dados para cada cliente na base de dados.
- Gerar um documento Word para cada grupo de registos.
- Salvar cada documento como PDF, tendo o cuidado de nomear o arquivo corretamente.
- Organizar centenas de arquivos em pastas.

Além da geração dos arquivos, o fluxo de trabalho manual incluiria:
- Criar um e-mail para cada cliente no Outlook.
- Anexar o PDF correto ao destinatário correspondente, um grande ponto de falha.
- Enviar centenas de e-mails individualmente.

---

## ✨ A Solução
Esta ferramenta automatiza 100% do processo. Para se adaptar a diferentes necessidades, a ferramenta foi desenvolvida com três modos de operação distintos:

**1. Apenas Geração de PDFs:** O sistema lê a base de dados, agrupa as informações por cliente, gera um PDF consolidado para cada um e salva todos numa pasta predefinida. Ideal para arquivo ou para quando o envio não é imediato.

**2. Geração de PDFs com Preparo de E-mails (Modo de Revisão):** Após gerar cada PDF, a ferramenta cria um rascunho de e-mail no Outlook, já com o destinatário, assunto, corpo de texto e o PDF correto anexado. Permite uma revisão final antes do envio manual.

**3. Automação Completa (Envio Direto):** Executa todas as etapas anteriores e, no final, envia cada e-mail diretamente para a "Caixa de Saída" do Outlook, sem necessidade de intervenção manual. É a solução mais rápida para automação completa.

**O resultado é a conversão de um trabalho de horas numa tarefa de poucos minutos, com garantia de precisão e organização.**

---

## 🚀 Funcionalidades Principais
- **Geração de PDFs Agrupados por Cliente:** Consolida múltiplos registos de um mesmo cliente num único arquivo PDF.
- **Integração Completa com Outlook:** Prepara e/ou envia e-mails com destinatário, assunto e corpo da mensagem, anexando o PDF correspondente.
- **Modos de Operação Flexíveis:** Permite que o utilizador escolha entre apenas gerar os PDFs, criar rascunhos no Outlook para revisão, ou enviar todos os e-mails automaticamente.
- **Interface de Progresso Visual:** Um popup informa o utilizador sobre o progresso da tarefa em tempo real (ex: "Processando Cliente 20 de 124").
- **Log de Erros Inteligente:** Caso ocorra alguma falha, um arquivo de log `.txt` é criado automaticamente, listando apenas os itens que falharam e o motivo, facilitando a correção.
- **Criação Automática de Pastas:** Verifica se a pasta de destino existe e, caso não exista, cria-a automaticamente.
- **Performance Otimizada:** Utiliza conexão de banco de dados (ADO) para uma leitura ultra-rápida da fonte de dados e executa de forma otimizada para máxima velocidade.
- **Tratamento de Erros Robusto:** O código é construído para ser estável, com mecanismos que lidam com instabilidades do Outlook e previnem a interrupção do processo em grandes volumes.

---

## 📸 Demonstração Visual

| Descrição | Demonstração |
| :--- | :--- |
| **Popup de Progresso em Ação:** <br> Mostra o andamento do processo em tempo real, informando o cliente atual e a contagem total. | <img src="https://raw.githubusercontent.com/pricmendes/VBA/refs/heads/word/assets/rodando.jpg" width="150"> |
| **Resultado Final e Log de Erros:** <br> Apresenta um resumo da execução e informa se foi gerado um log de falhas. | <img src="https://raw.githubusercontent.com/pricmendes/VBA/refs/heads/word/assets/finalizado.jpg" width="150"> |

---

## ⚠️ Pontos de Atenção e Requisitos
Para garantir o funcionamento correto da ferramenta, o ambiente do utilizador precisa de cumprir os seguintes pré-requisitos:

| Requisito | Especificação | Observações Resumidas |
| :--- | :--- | :--- |
| **Pacote Office** | Office 2016 ou superior (incluindo M365) | Requer Word, Excel e Outlook(Versão Classic) instalados. |
| **Segurança de Macros**| Macros Habilitadas | Necessário clicar em "Habilitar Conteúdo" no aviso inicial. |
| **Driver de Acesso a Dados** | Microsoft Access Database Engine (2016+) | Essencial para a leitura otimizada do Excel. Pode requerer instalação manual. |
| **Arquitetura do Office** | Versão 32-bit ou 64-bit | A versão do Driver de Dados deve ser a mesma do Office instalado. |

---

## 🛠️ Tecnologias Utilizadas
- **VBA** (Visual Basic for Applications)
- Microsoft **Word**
- Microsoft **Excel**
- Microsoft **Outlook**
- **ADO** (ActiveX Data Objects)
- **FSO** (FileSystemObject)

---

## 💼 Contato
Interessado em automatizar a geração de documentos e outros processos na sua empresa? Entre em contato para discutir soluções personalizadas.

- **Email:** pricdados@gmail.com
- **LinkedIn:** [Priscila Cardoso](https://www.linkedin.com/in/priscila-mendes-sp/)
