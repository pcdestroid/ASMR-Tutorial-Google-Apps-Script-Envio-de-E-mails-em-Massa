function enviarEmailsEmMassa() {
  // Abra a planilha ativa e selecione a primeira aba
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const aba = planilha.getActiveSheet();

  // Obtenha os dados da planilha
  const dados = aba.getDataRange().getValues(); // Todas as linhas e colunas

  // Loop para percorrer os dados, começando na segunda linha (índice 1)
  for (let i = 1; i < dados.length; i++) {
    const nome = dados[i][0]; // Coluna Nome
    const email = dados[i][1]; // Coluna E-mail
    const mensagem = dados[i][2]; // Coluna Mensagem

    // Verifica se o e-mail está preenchido
    if (email) {
      // Envie o e-mail
      GmailApp.createDraft(email, `Olá, ${nome}!`, mensagem);
      console.log(`E-mail enviado para: ${nome} - ${email}`);
    }
  }
}
