function LIMPA() {

  var dataAtual = new Date()

  var mes = dataAtual.toLocaleString('pt-BR', { month: 'long' });
  var ano = dataAtual.getFullYear();
  var nome = mes + " / " + ano

  var ss = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1sOmk5msu5tT8FY0LLXjIpV0jaO3leBC0c7UQVL3OYjE/edit#gid=1543213291");
  var baseDeDados = ss.getSheetByName("ATUAL");
  var today = new Date();
  var lastDayOfMonth = new Date(today.getFullYear(), today.getMonth()+1, 0);

  if(today.getDate() == lastDayOfMonth.getDate() )
  {
  baseDeDados.copyTo(ss).setName(nome).hideSheet();
  console.log("copiado")

  var lastRow = baseDeDados.getMaxRows();
  baseDeDados.deleteRows(5, lastRow-4);
  }
  
}

function prevlogin(e) {
  const ps=PropertiesService.getUserProperties();
  const useremail=e.user.getEmail();
  const date=Utilities.formatDate(new Date(),Session.getScriptTimeZone(),"yyyyMMdd");
  let lastlogin=ps.getProperty(useremail);
  if(!lastlogin) {
    ps.setProperty(useremail,date)
    return false;
  }else if(lastlogin!=date) {
    ps.setProperty(useremail,date)
    return false;
  }else {
    return true;
  }
}

function f1(e) {

  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu("FLUXO");
  menu.addItem("Abrir", "openLink");
  menu.addToUi();

  /*if(!prevlogin(e)) {
   SpreadsheetApp.getUi().alert('Bem-vindo ao Agendamento de Carros da Embalagem! \n \n Por favor! \n Não adicione linhas extras na Página "ATUAL", o próprio Menu fará isso. \n \n A cada tarefa completa estamos mais perto do sucesso. \n \n Bom Trabalho!');
  }*/
  
}


function f2() {
  var link = "https://script.google.com/a/macros/whirlpool.com/s/AKfycbxBbB-yI2JutaNYBR-UtY7RQ_5zCvdeG2hIX6KSSdhA/dev";
  var html = '<script>window.open("' + link + '", "_blank");google.script.host.close();</script>';
  var ui = HtmlService.createHtmlOutput(html);
  SpreadsheetApp.getUi().showModelessDialog(ui, "EXECUTANDO PORTAL vs2.01..."); ///alterar conforma versão
}
