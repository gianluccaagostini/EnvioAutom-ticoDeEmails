function enviarEmails() {
  var sheet = SpreadsheetApp.getActiveSheet();// Planilha ativa
  var startRow = 2; //Primeira linha de dados processado
  var numRows = sheet.getRange(2, 13).getValue(); // número de linhas a serem processadas da célula M2

  // buscar valores para a linha no intervalo
  var dataRange  = sheet.getRange(startRow, 1, numRows, 14);
  
  // buscar valores para a linha no intervalo
  var data = dataRange.getValues();
  
  
  for(i in data){
    var rec = data[i];
    var cliente = 
      {
        inicioAssinatura: rec[0],
        fimAssinatura: rec[1],
        prazoRenovacao: rec[2],
        servico: rec[3],
        codigoCliente: rec[4],
        codigoProduto: rec[5],
        valor: rec[6],
        primeiraCompra: rec[7],
        nome: rec[8],
        email: rec[9],
        telefone: rec[10],
        endereco: rec[11],
      };
        var template = HtmlService
        .createTemplateFromFile('modelo1');
    
        template.cliente = cliente;
    
        var message = template.evaluate().getContent();
    
        MailApp.sendEmail({
        to: cliente.email,
        subject: "Seu contrato de serviços está vencendo!! ",
        htmlBody: message
        });  
    

  }
}