function logError(errorMessage) {
  var spreadsheetId = '1sOmk5msu5tT8FY0LLXjIpV0jaO3leBC0c7UQVL3OYjE'; // ID da planilha
  var sheetName = 'error';

  var sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(sheetName);

  if (!sheet) {
    throw new Error("A guia 'error' não foi encontrada na planilha.");
  }

  var timestamp = new Date();
  var rowData = [timestamp, errorMessage];
  sheet.appendRow(rowData);
}


function doGetPage() {
  try {
    var htmlOutput = HtmlService.createTemplateFromFile('index');
    return htmlOutput.evaluate().setTitle("FLUXO").setFaviconUrl("https://www.whirlpool.com/content/dam/business-unit/global-assets/images/favicons/favicon.ico").setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    var htmlOutput = HtmlService.createTemplateFromFile('error');
    return htmlOutput.evaluate().setTitle("FLUXO").setFaviconUrl("https://www.whirlpool.com/content/dam/business-unit/global-assets/images/favicons/favicon.ico").setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
}

function doGet() {
  try {
    var htmlOutput = HtmlService.createTemplateFromFile('index');
    return htmlOutput.evaluate().setTitle("FLUXO").setFaviconUrl("https://www.whirlpool.com/content/dam/business-unit/global-assets/images/favicons/favicon.ico").setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    var htmlOutput = HtmlService.createTemplateFromFile('error');
    return htmlOutput.evaluate().setTitle("FLUXO").setFaviconUrl("https://www.whirlpool.com/content/dam/business-unit/global-assets/images/favicons/favicon.ico").setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
}

function getNewHtml(e) {
      var html = HtmlService
        .createTemplateFromFile('index') // uses templated html
        .evaluate()
        .getContent();
return html;}

function gerarChaveUnica() {
  var planilhaURL = "https://docs.google.com/spreadsheets/d/11ZJ1GCIY5IL6-_GPAgFvxBF589uX8M26iXs_wsWG4ic/edit#gid=798358046"; // Substitua pelo URL da planilha desejada
  var planilha = SpreadsheetApp.openByUrl(planilhaURL); // Abre a planilha pelo URL
  
  var planilhaNome = "CHAVE"; // Substitua pelo nome da planilha desejada
  var planilhaAtiva = planilha.getSheetByName(planilhaNome); // Obtém a planilha pelo nome
  
  // Resto do código permanece o mesmo
  // Gerar uma chave única
  var chave = Utilities.getUuid();
  
  // Verificar se a chave já existe na planilha
  var chaveExiste = planilhaAtiva.createTextFinder(chave).findNext();
  
  // Se a chave já existe, gerar uma nova chave
  while (chaveExiste) {
    chave = Utilities.getUuid();
    chaveExiste = planilhaAtiva.createTextFinder(chave).findNext();
  }
  
  // Adicionar a chave à planilha
  //var ultimaLinha = planilhaAtiva.getLastRow() + 1;
  planilhaAtiva.getRange("A1").setValue(chave);
  console.log(chave)
}


function saveToSheet(dateString, nomeMoto, teleMoto, tpNum, horaAgenda, dataAgenda, timeFiscal, transp, placa, hourString, forn12, teleForn, status, horaEntrada, horaSaida, id, id2, conf, stats, n1, n2, n3, n4, n5, n6, n7, n8, n9, n10, m6) {
  // Verifica se a função está sendo executada, se sim, retorna false para evitar duplicação

  var ss = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1sOmk5msu5tT8FY0LLXjIpV0jaO3leBC0c7UQVL3OYjE/edit#gid=1543213291");
  var sheet = ss.getSheetByName("ATUAL");
  var registro = ss.getSheetByName("LOG DE REGISTROS");


  try {
    // Grava os dados na planilha
    sheet.appendRow([dateString, nomeMoto, teleMoto, tpNum, horaAgenda, dataAgenda, timeFiscal, transp, placa, hourString, forn12, teleForn, status, horaEntrada, horaSaida, id, id2, conf, stats, n1, n2, n3, n4, n5, n6, n7, n8, n9, n10, m6]);



    registro.appendRow([timeFiscal, teleForn, "", nomeMoto, dateString]);
    console.log("REGISTRADO NO LOG")

  } catch (e) { 

    Logger.log("ERRO DE REGISTRO");
    console.log("ERRO DE REGISTRO");
    registro.appendRow(["ERRO", "ERRO", "", "ERRO DE REGISTRO", "ERRO"]);
  
  
  }
    // Libera a função após o tempo de gravação na planilha

  return true;
}


function searchValueInColumnD(searchValue) {
var ss = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1sOmk5msu5tT8FY0LLXjIpV0jaO3leBC0c7UQVL3OYjE/edit#gid=1543213291")
var sheet = ss.getSheetByName("ATUAL");
var range = sheet.getDataRange().getValues();
for (var i = 0; i < range.length; i++) {
   if (range[i][3] == searchValue) {
      var foundValue = range[i][3]
      return foundValue 
    }
  }
  return false;
}

function createDropdown() {
    var ss = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1sOmk5msu5tT8FY0LLXjIpV0jaO3leBC0c7UQVL3OYjE/edit#gid=1543213291") // .getActiveSheet(); // Acessa a pagina aberta atual da planilha
    var sheet = ss.getSheetByName("Motoristas")
    var data = sheet.getDataRange().getValues();
    var select = document.getElementById("dropdown");

    for (var i = 0; i < data.length; i++) {
      var option = document.createElement("option");
      option.text = data[i][0];
      select.add(option);
    }
  }

function getData() {
    var ss = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1sOmk5msu5tT8FY0LLXjIpV0jaO3leBC0c7UQVL3OYjE/edit#gid=1543213291")
    var sheet = ss.getSheetByName("Motoristas");
    var data = sheet.getRange(3,1,sheet.getLastRow(),1).getValues();
    return data;
  }

////////

function createDropdownIndireto() {
    var ss = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1sOmk5msu5tT8FY0LLXjIpV0jaO3leBC0c7UQVL3OYjE/edit#gid=1543213291") // .getActiveSheet(); // Acessa a pagina aberta atual da planilha
    var sheet = ss.getSheetByName("Motoristas")
    var data = sheet.getDataRange().getValues();
    var select = document.getElementById("dropdown");

    for (var i = 0; i < data.length; i++) {
      var option = document.createElement("option");
      option.text = data[i][0];
      select.add(option);
    }
  }

function getDataIndireto() {
    var ss = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1sOmk5msu5tT8FY0LLXjIpV0jaO3leBC0c7UQVL3OYjE/edit#gid=1543213291")
    var sheet = ss.getSheetByName("Motoristas");
    var data = sheet.getRange(3,1,sheet.getLastRow(),1).getValues();
    return data;
  }

///////

function createDropdownSemTp() {
    var ss = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1sOmk5msu5tT8FY0LLXjIpV0jaO3leBC0c7UQVL3OYjE/edit#gid=1543213291") // .getActiveSheet(); // Acessa a pagina aberta atual da planilha
    var sheet = ss.getSheetByName("Motoristas")
    var data = sheet.getDataRange().getValues();
    var selectSemTp = document.getElementById("dropdownSemTp");

    for (var i = 0; i < data.length; i++) {
      var optionSemTp = document.createElement("option");
      optionSemTp.text = data[i][0];
      selectSemTp.add(option);
    }
  }



  function getDataSemTp() {
    var ss = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1sOmk5msu5tT8FY0LLXjIpV0jaO3leBC0c7UQVL3OYjE/edit#gid=1543213291")
    var sheet = ss.getSheetByName("Motoristas");
    var dataSemTp = sheet.getRange(3,1,sheet.getLastRow(),1).getValues();
    return dataSemTp;
  }



function createDropdownForn() {
  var ss = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1sOmk5msu5tT8FY0LLXjIpV0jaO3leBC0c7UQVL3OYjE/edit#gid=1543213291"); // .getActiveSheet(); // Acessa a pagina aberta atual da planilha
  var sheet = ss.getSheetByName("Fornecedores"); // altera o nome da página para "Fornecedores"
  var data = sheet.getDataRange().getValues();
  var selectForn = document.getElementById("dropdownForn");

  for (var i = 0; i < data.length; i++) {
    var optionForn = document.createElement("option");
    optionForn.text = data[i][1];
    selectForn.add(optionForn); 
  }
}


function getDataForn() {
  var ss = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1sOmk5msu5tT8FY0LLXjIpV0jaO3leBC0c7UQVL3OYjE/edit#gid=1626438448");
  var sheet = ss.getSheetByName("Fornecedores"); // altera o nome da página para "Fornecedores"
  var dataForn = sheet.getRange(2,2,sheet.getLastRow(),1).getValues(); // altera os parâmetros da getRange() para pegar a coluna K (11) e ajusta o número de linhas para pegar apenas os dados (subtraindo 2 da última linha)
  return dataForn;
}


////////////////

function createDropdownFornIndireto() {
  var ss = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1sOmk5msu5tT8FY0LLXjIpV0jaO3leBC0c7UQVL3OYjE/edit#gid=1543213291"); // .getActiveSheet(); // Acessa a pagina aberta atual da planilha
  var sheet = ss.getSheetByName("Fornecedores"); // altera o nome da página para "Fornecedores"
  var data = sheet.getDataRange().getValues();
  var selectForn = document.getElementById("dropdownFornIndireto");

  for (var i = 0; i < data.length; i++) {
    var optionForn = document.createElement("option");
    optionForn.text = data[i][1];
    selectForn.add(optionForn); 
  }
}


function getDataFornIndireto() {
  var ss = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1sOmk5msu5tT8FY0LLXjIpV0jaO3leBC0c7UQVL3OYjE/edit#gid=1626438448");
  var sheet = ss.getSheetByName("Fornecedores"); // altera o nome da página para "Fornecedores"
  var dataForn = sheet.getRange(2,2,sheet.getLastRow(),1).getValues(); // altera os parâmetros da getRange() para pegar a coluna K (11) e ajusta o número de linhas para pegar apenas os dados (subtraindo 2 da última linha)
  return dataForn;
}

////////////////

function searchForn(searchCodForn) {
  var ss = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1sOmk5msu5tT8FY0LLXjIpV0jaO3leBC0c7UQVL3OYjE/edit#gid=1626438448")
  var sheet3 = ss.getSheetByName("Fornecedores")
  var range3 = sheet3.getDataRange().getValues(); // get all the data in the sheet
  
  searchCodForn = searchCodForn.trim(); // remove any spaces at the beginning and end of the search value
  
  for (var i = 0; i < range3.length; i++) {
    var fornSearch = range3[i][1]; // remove any spaces at the beginning and end of the value in the spreadsheet
    if (fornSearch == searchCodForn) {
      var columnT1 = range3[i][2]
      
      SpreadsheetApp.flush();
      return { columnT1 };
    }
  }
  return "Não encontrado";
}

function searchFornIndireto(searchFornCodIndireto) {
  var ss = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1sOmk5msu5tT8FY0LLXjIpV0jaO3leBC0c7UQVL3OYjE/edit#gid=1626438448")
  var sheet3 = ss.getSheetByName("Fornecedores")
  var range3 = sheet3.getDataRange().getValues(); // get all the data in the sheet
  
  searchFornCodIndireto = searchFornCodIndireto.trim(); // remove any spaces at the beginning and end of the search value
  
  for (var i = 0; i < range3.length; i++) {
    var fornSearch = range3[i][1].trim(); // remove any spaces at the beginning and end of the value in the spreadsheet
    if (fornSearch == searchFornCodIndireto) {
      var columnT1 = range3[i][2]
      
      SpreadsheetApp.flush();
      return { columnT1 };
    }
  }
  return "Não encontrado";
}


function searchData(searchValue) {
  try {
    var ss = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1sOmk5msu5tT8FY0LLXjIpV0jaO3leBC0c7UQVL3OYjE/edit#gid=1036091630");
    var sheet = ss.getSheetByName("Espelho");
    var range = sheet.getDataRange().getValues(); // get all the data in the sheet

    for (var i = 0; i < range.length; i++) {
      if (range[i][0].split(" - ")[0] == searchValue) {
        var fornecedor = range[i][1];
        var transportadora = range[i][2];
        var dataAgendada = range[i][3];
        var horaAgendada = range[i][4];
        var tipoDeTp = "EMB."

        return { fornecedor, transportadora, dataAgendada, horaAgendada, tipoDeTp };
      }
    }

    var ss3 = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1sOmk5msu5tT8FY0LLXjIpV0jaO3leBC0c7UQVL3OYjE/edit#gid=1036091630");
    var sheet3 = ss3.getSheetByName("Espelho - Template Agendamento");
    var range3 = sheet3.getDataRange().getValues(); // get all the data in the sheet

    for (var i = 0; i < range3.length; i++) {
      if (range3[i][0] == searchValue) {
        var fornecedor = range3[i][3]; // D column
        var transportadora = range3[i][12]; // M column
        var dataAgendada = range3[i][10]; // K column
        var horaAgendada = "-";
        var tipoMouE = range3[i][2]; // tipo de tp
        var tipoDeTp = "EMB."

        if (tipoMouE == "M") {
          return "Não encontrado";
        }

        return { fornecedor, transportadora, dataAgendada, horaAgendada, tipoDeTp };
      }
    }

    // Terceira busca adicionada
    var ss4 = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1sOmk5msu5tT8FY0LLXjIpV0jaO3leBC0c7UQVL3OYjE/edit#gid=1673113086");
    var sheet4 = ss4.getSheetByName("Sub-Contratados");
    var range4 = sheet4.getDataRange().getValues();

    for (var i = 0; i < range4.length; i++) {
      for (var j = 0; j < range4[i].length; j++) {
        if (range4[i][j] == searchValue) {
          var fornecedor = range4[i][4]; // Coluna E
          var transportadora = range4[i][9]; // Coluna J
          var dataAgendada = range4[0][j]; // Primeira célula da coluna
          var horaAgendada = range4[i][6]; // Coluna G
          var tipoDeTp = "INDUST."

          return { fornecedor, transportadora, dataAgendada, horaAgendada, tipoDeTp };
        }
      }
    }



    var ss2 = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/11ZJ1GCIY5IL6-_GPAgFvxBF589uX8M26iXs_wsWG4ic/edit#gid=918319974");
    var sheet2 = ss2.getSheetByName("ORDENS");
    var range2 = sheet2.getDataRange().getValues();

    for (var i = 0; i < range2.length; i++) {
      if (range2[i][0].split(" - ")[0] == searchValue) {
        var fornecedor = range2[i][1].split(" - ")[0];
        var transportadora = range2[i][2].split(" - ")[1] || range2[i][2];
        var dataAgendada = range2[i][3];
        var horaAgendada = range2[i][4];
        var tipoDeTp = "INDUST."

        return { fornecedor, transportadora, dataAgendada, horaAgendada, tipoDeTp };
      }
    } 


    return "Não encontrado";
  } catch (e) {
    console.log("GS Search erro");
    return "Não encontrado";
  }
}


///////////*** */
function searchFone(searchMotoFone) {
  var url = 'https://docs.google.com/spreadsheets/d/1sOmk5msu5tT8FY0LLXjIpV0jaO3leBC0c7UQVL3OYjE/edit#gid=862230886';
  var ss = SpreadsheetApp.openByUrl(url);
  var sheet2 = ss.getSheetByName("Motoristas")
  var range2 = sheet2.getDataRange().getValues(); // get all the data in the sheet
  
  searchMotoFone = searchMotoFone.trim(); // remove any spaces at the beginning and end of the search value
  
  for (var i = 0; i < range2.length; i++) {
    var motoFone = range2[i][0].trim(); // remove any spaces at the beginning and end of the value in the spreadsheet
    if (motoFone == searchMotoFone) {
      var columnT1 = range2[i][1];
      
      SpreadsheetApp.flush();
      return { columnT1 };
    }
  }
  return "Não encontrado";
}


function searchFoneSemTp(searchMotoFoneSemTp) {
  var url = 'https://docs.google.com/spreadsheets/d/1sOmk5msu5tT8FY0LLXjIpV0jaO3leBC0c7UQVL3OYjE/edit#gid=862230886';
  var ss = SpreadsheetApp.openByUrl(url);
  var sheet2 = ss.getSheetByName("Motoristas")
  var range2 = sheet2.getDataRange().getValues(); // get all the data in the sheet
  
  searchMotoFoneSemTp = searchMotoFoneSemTp.trim(); // remove any spaces at the beginning and end of the search value
  
  for (var i = 0; i < range2.length; i++) {
    var motoFone = range2[i][0].trim(); // remove any spaces at the beginning and end of the value in the spreadsheet
    if (motoFone == searchMotoFoneSemTp) {
      var columnT1 = range2[i][1];
      
      SpreadsheetApp.flush();
      return { columnT1 };
    }
  }
  return "Não encontrado";
}

function searchFoneIndireto(searchMotoFoneIndireto) {
  var url = 'https://docs.google.com/spreadsheets/d/1sOmk5msu5tT8FY0LLXjIpV0jaO3leBC0c7UQVL3OYjE/edit#gid=862230886';
  var ss = SpreadsheetApp.openByUrl(url);
  var sheet2 = ss.getSheetByName("Motoristas")
  var range2 = sheet2.getDataRange().getValues(); // get all the data in the sheet
  
  searchMotoFoneIndireto = searchMotoFoneIndireto.trim(); // remove any spaces at the beginning and end of the search value
  
  for (var i = 0; i < range2.length; i++) {
    var motoFone = range2[i][0].trim(); // remove any spaces at the beginning and end of the value in the spreadsheet
    if (motoFone == searchMotoFoneIndireto) {
      var columnT1 = range2[i][1];
      
      SpreadsheetApp.flush();
      return { columnT1 };
    }
  }
  return "Não encontrado";
}

function checkLogin(username, password) {

  var url = 'https://docs.google.com/spreadsheets/d/11ZJ1GCIY5IL6-_GPAgFvxBF589uX8M26iXs_wsWG4ic/edit#gid=0';
  var ss = SpreadsheetApp.openByUrl(url);
  var webAppSheet = ss.getSheetByName("ACESS");
  var getLastRow =  webAppSheet.getLastRow();

  var found_record = '';

  for(var i = 1; i <= getLastRow; i++)
  {
   if(webAppSheet.getRange(i, 1).getValue() == username.toUpperCase() && 
     webAppSheet.getRange(i, 2).getValue() == password.toUpperCase())
   { 
     found_record = 'TRUE';
     webAppSheet.getRange(i, 6).setValue(new Date()).setNumberFormat ('dd/MM/yy - HH:MM am/pm');

   }    
  }
  if(found_record == '')
  {
    found_record = 'FALSE'; 
  }
  
  return found_record;
  
}

function getLevel(username) {
  var url = 'https://docs.google.com/spreadsheets/d/11ZJ1GCIY5IL6-_GPAgFvxBF589uX8M26iXs_wsWG4ic/edit#gid=0';
  var ss= SpreadsheetApp.openByUrl(url);
  var sheet = ss.getSheetByName("ACESS");
  var data = sheet.getDataRange().getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] == username.toUpperCase()) {
      return data[i][4];
    }
  }
  return "";
}

function getValorCelula() {
  var url = 'https://docs.google.com/spreadsheets/d/1GTh9iA3VLibtxmFJjjyP8BbS7B_LVopleQ0SFDAS37k/edit#gid=2016617566';
  var ss = SpreadsheetApp.openByUrl(url);
  var sheet = ss.getSheetByName("Weather");
  var valor = sheet.getRange("B1").getValue();
  return valor;
}

function getEmail(username) {
  var url = 'https://docs.google.com/spreadsheets/d/11ZJ1GCIY5IL6-_GPAgFvxBF589uX8M26iXs_wsWG4ic/edit#gid=0';
  var ss = SpreadsheetApp.openByUrl(url);
  var webAppSheet = ss.getSheetByName("ACESS");
  var getLastRow =  webAppSheet.getLastRow();
  var email = '';
  for(var i = 1; i <= getLastRow; i++) {
   if(webAppSheet.getRange(i, 1).getValue() == username.toUpperCase()) {
     email = webAppSheet.getRange(i, 3).getValue();
   }
  }
  return email;
}

function getLastLineOfColumnQ() {
  // Get the active spreadsheet
  var url = 'https://docs.google.com/spreadsheets/d/1sOmk5msu5tT8FY0LLXjIpV0jaO3leBC0c7UQVL3OYjE/edit#gid=1004243275';
  var ss= SpreadsheetApp.openByUrl(url);
  
  // Get the active sheet
  var sheet = ss.getSheetByName("ATUAL");
  
  // Get the last row number
  var lastRow = sheet.getLastRow();
  
  // Get the value of the last cell in column Q
  var lastValue = sheet.getRange("Q" + lastRow).getValue();
  
  return lastValue;
}


function AddRecord(usernamee, passwordd, email, phone, acess) {
  var url = 'https://docs.google.com/spreadsheets/d/11ZJ1GCIY5IL6-_GPAgFvxBF589uX8M26iXs_wsWG4ic/edit#gid=0';
  var ss= SpreadsheetApp.openByUrl(url);
  var webAppSheet = ss.getSheetByName("ACESS");
  var data = webAppSheet.getDataRange().getValues();
  var existingUsernames = data.map(function(row) { return row[0]; });
  
  if (existingUsernames.indexOf(usernamee) > -1) {
    return true;
  } else {
    var newRow = [usernamee, passwordd, email, phone, acess];
    webAppSheet.appendRow(newRow);
    return false;
  }
}

function AddRegistroMoto(nome, telefone, cnh, regis) {
  var url = 'https://docs.google.com/spreadsheets/d/1sOmk5msu5tT8FY0LLXjIpV0jaO3leBC0c7UQVL3OYjE/edit#gid=862230886';
  var ss = SpreadsheetApp.openByUrl(url);
  var webAppSheet = ss.getSheetByName("Motoristas");
  var newRow = [nome, telefone, cnh, regis]
  webAppSheet.appendRow(newRow);

  var url2 = 'https://docs.google.com/spreadsheets/d/11hDe7eEL7CT-UQxnt0Zx-oKObB6uJlSmxuN5aOFIpvE/edit#gid=0';
  var ss2 = SpreadsheetApp.openByUrl(url2);
  var webAppSheet2 = ss2.getSheetByName("Motoristas");
  var newRow2 = [nome, telefone, cnh, nome]
  webAppSheet2.appendRow(newRow2);
}

function AddRegistroForn(id, nome, regis) {
  var url = 'https://docs.google.com/spreadsheets/d/1sOmk5msu5tT8FY0LLXjIpV0jaO3leBC0c7UQVL3OYjE/edit#gid=862230886';
  var ss = SpreadsheetApp.openByUrl(url);
  var webAppSheet = ss.getSheetByName("Fornecedores");
  var newRow = [id, nome, "" ,regis]
  webAppSheet.appendRow(newRow);

  var url2 = 'https://docs.google.com/spreadsheets/d/11hDe7eEL7CT-UQxnt0Zx-oKObB6uJlSmxuN5aOFIpvE/edit#gid=0';
  var ss2 = SpreadsheetApp.openByUrl(url2);
  var webAppSheet2 = ss2.getSheetByName("Fornecedores");
  var newRow2 = [id, nome, "",regis]
  webAppSheet2.appendRow(newRow2);
}

function getSheetData() {
  var url = 'https://docs.google.com/spreadsheets/d/1GTh9iA3VLibtxmFJjjyP8BbS7B_LVopleQ0SFDAS37k/edit#gid=1926743166';
  var ss = SpreadsheetApp.openByUrl(url);
  var sheet = ss.getSheetByName("VarLet");
  var range = sheet.getRange("A4:N");
  var data = range.getValues();

  var uniqueData = [];
  for (var i = 0; i < data.length; i++) {
    var row = [];
    var rowHasValue = false;
    for (var j = 0; j < data[i].length; j++) {
      if (data[i][j] !== '' && data[i][j] !== null) {
        rowHasValue = true;
        row.push(data[i][j]);
      } else {
        row.push('');
      }
    }
    if (rowHasValue && uniqueData.indexOf(row) === -1) {
      uniqueData.push(row);
    }
  }

  return uniqueData;
}

function getSheetDataNotes() {
  var url = 'https://docs.google.com/spreadsheets/d/1GTh9iA3VLibtxmFJjjyP8BbS7B_LVopleQ0SFDAS37k/edit#gid=1926743166';
  var ss = SpreadsheetApp.openByUrl(url);
  var sheet = ss.getSheetByName("PPCP");
  var range = sheet.getRange("T14:V5000");
  var data = range.getValues();

  var uniqueDataNotes = [];
  for (var i = 0; i < data.length; i++) {
    var row = [];
    var rowHasValue = false;
    for (var j = 0; j < data[i].length; j++) {
      if (data[i][j] !== '') {
        rowHasValue = true;
        row.push(data[i][j]);
      } else {
        row.push('');
      }
    }
    if (rowHasValue && uniqueDataNotes.indexOf(row) === -1) {
      uniqueDataNotes.push(row);
    }
  }

  return uniqueDataNotes;
}

function getSheetDataLast5() {
  var url = 'https://docs.google.com/spreadsheets/d/1sOmk5msu5tT8FY0LLXjIpV0jaO3leBC0c7UQVL3OYjE/edit#gid=1543213291';
  var ss = SpreadsheetApp.openByUrl(url);
  var sheet = ss.getSheetByName("LAST");
  var range = sheet.getRange("A1:F9");
  var data = range.getValues();

  var uniqueDataLast5 = [];
  for (var i = 0; i < data.length; i++) {
    var row = [];
    var rowHasValue = false;
    for (var j = 0; j < data[i].length; j++) {
      if (data[i][j] != '') {
        rowHasValue = true;
        row.push(data[i][j]);
      } else {
        row.push('');
      }
    }
    if (rowHasValue && uniqueDataLast5.indexOf(row) === -1) {
      uniqueDataLast5.push(row);
    }
  }

  return uniqueDataLast5;
}

function getSheetDataLast5Nota() {
  var url = 'https://docs.google.com/spreadsheets/d/1sOmk5msu5tT8FY0LLXjIpV0jaO3leBC0c7UQVL3OYjE/edit#gid=1543213291';
  var ss = SpreadsheetApp.openByUrl(url);
  var sheet = ss.getSheetByName("LASTNota");
  var range = sheet.getRange("A1:F10");
  var data = range.getValues();
  data = data.reverse();

  var uniqueDataLast5 = [];
  for (var i = 0; i < data.length; i++) {
    var row = [];
    var rowHasValue = false;
    for (var j = 0; j < data[i].length; j++) {
      if (data[i][j] != '') {
        rowHasValue = true;
        row.push(data[i][j]);
      } else {
        row.push('');
      }
    }
    if (rowHasValue && uniqueDataLast5.indexOf(row) === -1) {
      uniqueDataLast5.push(row);
    }
  }

  return uniqueDataLast5;
}

function getSheetDataLast5Indireto() {
  var url = 'https://docs.google.com/spreadsheets/d/11hDe7eEL7CT-UQxnt0Zx-oKObB6uJlSmxuN5aOFIpvE/edit#gid=0';
  var ss = SpreadsheetApp.openByUrl(url);
  var sheet = ss.getSheetByName("LAST");
  var range = sheet.getRange("A1:F9");
  var data = range.getValues();

  var uniqueDataLast5 = [];
  for (var i = 0; i < data.length; i++) {
    var row = [];
    var rowHasValue = false;
    for (var j = 0; j < data[i].length; j++) {
      if (data[i][j] != '') {
        rowHasValue = true;
        row.push(data[i][j]);
      } else {
        row.push('');
      }
    }
    if (rowHasValue && uniqueDataLast5.indexOf(row) === -1) {
      uniqueDataLast5.push(row);
    }
  }

  return uniqueDataLast5;
}

function toggleConfirmation(rowIndex, id) {
  var url = 'https://docs.google.com/spreadsheets/d/1sOmk5msu5tT8FY0LLXjIpV0jaO3leBC0c7UQVL3OYjE/edit#gid=1568882513';
  var ss = SpreadsheetApp.openByUrl(url);
  var sheet = ss.getSheetByName("ATUAL");
  
  // Find the row with the matching ID in column AJ
  var lastRow = sheet.getLastRow();
  var idRange = sheet.getRange(2, 41, lastRow - 1, 1);
  var idValues = idRange.getValues();
  var matchingRow = null;
  for (var i = 0; i < idValues.length; i++) {
    if (idValues[i][0] == id) {
      matchingRow = i + 2;
      break;
    }
  }
  
  if (matchingRow) {
    var range = sheet.getRange(matchingRow, 30); // Column Y (25) for true/false toggle
    var currentValue = range.getValue();
    var newValue = (currentValue === true) ? true : true; //false : true;
    range.setValue(newValue);
    
    // Add the current time to column AK (37) if the value is set to true
    if (newValue) {
      var timeColumn = 42; // Column AK
      var timeRange = sheet.getRange(matchingRow, timeColumn);
      var currentTime = new Date();
      timeRange.setValue(currentTime);
    } else {
      // Clear the time from column AK if the value is set to false
      var timeColumn = 42; // Column AK
      var timeRange = sheet.getRange(matchingRow, timeColumn);
      timeRange.setValue("");
    }
    
    return newValue; // Return the new value of the cell
  }
  return null;
}


function getSheetDataLast20() {

  var url = 'https://docs.google.com/spreadsheets/d/1sOmk5msu5tT8FY0LLXjIpV0jaO3leBC0c7UQVL3OYjE/edit#gid=1568882513';
  var ss = SpreadsheetApp.openByUrl(url);
  var sheet = ss.getSheetByName("ATUAL");

var lastRow = sheet.getLastRow();
var numRows = Math.max(30, lastRow - 3); // Garante um mínimo de 30 linhas ou lastRow - 3 linhas
var startRow = Math.max(4, lastRow - numRows + 1); // Linha inicial considerando o número de linhas

var range = sheet.getRange(startRow, 29, numRows, 48); // Coluna X (24) até coluna AS (66)

  //var range = sheet.getRange("X4:AS");

  var data = range.getValues();
  data.reverse();
  var uniqueDataLast20 = [];

  for (var i = data.length - 1; i >= 0; i--) {

    if (data[i][16] != "" || data[i][17] != "") {
      console.log(data[i][24])
      var row = [];
      var rowHasValue = false;
      for (var j = 0; j < data[i].length; j++) {
        if (data[i][j] != '') {
          rowHasValue = true;
          row.push(data[i][j]);
        } else {
          row.push('');
        }
      }
      if (rowHasValue && data[i][18] != 'TRUE' && data[i][21] == 'TRUE' && uniqueDataLast20.indexOf(row) === -1) {
        uniqueDataLast20.push(row);
        console.log(data[i][21]);
        //console.log(uniqueDataLast20[i][0])
      }
    }
  }

  if (uniqueDataLast20 == "") {
  SpreadsheetApp.flush();
  var url = 'https://docs.google.com/spreadsheets/d/1sOmk5msu5tT8FY0LLXjIpV0jaO3leBC0c7UQVL3OYjE/edit#gid=1568882513';
  var ss = SpreadsheetApp.openByUrl(url);
  var sheet = ss.getSheetByName("ATUAL");
  SpreadsheetApp.flush();

  var lastRow = sheet.getLastRow();
  var numRows = Math.max(20, lastRow - 3); // Garante um mínimo de 20 linhas ou lastRow - 3 linhas
  var startRow = Math.max(4, lastRow - numRows + 1); // Linha inicial considerando o número de linhas

  var range = sheet.getRange(startRow, 29, numRows, 48); // Coluna X (24) até coluna AS (66)

  //var range = sheet.getRange("X4:AS");

  var data = range.getDisplayValues();
  data.reverse();
  var uniqueDataLast20 = [];

  for (var i = data.length - 1; i >= 0; i--) {

    if (data[i][16] !== null && data[i][16] !== "" || data[i][17] !== null && data[i][17] !== "" ) {
      var row = [];
      var rowHasValue = false;
      for (var j = 0; j < data[i].length; j++) {
        if (data[i][j] != '') {
          rowHasValue = true;
          row.push(data[i][j]);
        } else {
          row.push('');
        }
      }
      if (rowHasValue && data[i][18] != 'TRUE' && data[i][21] == 'TRUE' && uniqueDataLast20.indexOf(row) === -1) {
        uniqueDataLast20.push(row);
        //console.log(data[i][24])
         console.log(uniqueDataLast20[i][0])
      }
    }
  }
  }

  var tableHtml = "<table style='max-height: 600px; user-select: text !important;' id='sheetDataLast20' >";
  tableHtml += "<tbody>";
  tableHtml += "<tr>";
  tableHtml += "<th style='position: sticky; top: 0;'> LIBERADO </th>";
  tableHtml += "<th style='position: sticky; top: 0;'>EMBALAGEM</th>";
  tableHtml += "<th style='position: sticky; top: 0;'>INDUSTRIAL</th>";
  tableHtml += "<th style='position: sticky; top: 0;'>PLACA & TRANSP.</th>";
  tableHtml += "<th style='position: sticky; top: 0;'>CONFIRMAÇÃO</th>";
  tableHtml += "<th style='position: sticky; top: 0;'>ID</th>";
  tableHtml += "</tr>";

  for (var i = 0; i < uniqueDataLast20.length; i++) {
      tableHtml += "<tr>";
      tableHtml += "<td>" + uniqueDataLast20[i][19] + "</td>";
      if (uniqueDataLast20[i][16] != '') {
        tableHtml += "<td><button type='button' class='btn btn-outline-primary btn-sm' onclick='window.open(\"" + uniqueDataLast20[i][16] + "\")'>" + 'Embalagem' + "</button></td>";

      } else {
        tableHtml += "<td></td>";
      }
      if (uniqueDataLast20[i][17] != '') {
        tableHtml += "<td><button type='button' class='btn btn-outline-success btn-sm' onclick='window.open(\"" + uniqueDataLast20[i][17] + "\")'>"  + 'Industrial' + "</button></td>";
      } else {
        tableHtml += "<td></td>";
      }
      tableHtml += "<td>" + uniqueDataLast20[i][0] + "</td>";

      tableHtml += "<td><div class='checkbox-wrapper-31'>" +
        "<input type='checkbox' id='confirmation" + i + "' onchange='chamarToggleConfirmation(event, " + i + ",\"" + uniqueDataLast20[i][11] + "\")' " +
        (uniqueDataLast20[i][1] === 'TRUE' ? 'checked' : '') + ">" +
        "<svg viewBox='0 0 35.6 35.6'>" +
        "<circle class='background' cx='17.8' cy='17.8' r='17.8'></circle>" +
        "<circle class='stroke' cx='17.8' cy='17.8' r='14.37'></circle>" +
        "<polyline class='check' points='11.78 18.12 15.55 22.23 25.17 12.87'></polyline>" +
        "</svg></div></td>";


      /*tableHtml += "<td><input type='checkbox' id='confirmation" + i + "' onchange='chamarToggleConfirmation(event, " + i + ",\"" + uniqueDataLast20[i][11] + "\")' " + (uniqueDataLast20[i][1] === 'TRUE' ? 'checked' : '') + "></td>";*/

      tableHtml += "<td>" + uniqueDataLast20[i][11] + "</td>";
      
     /* tableHtml += "<td><button type='button' class='btn btn-outline-primary btn-sm' onclick='imprimirPdf(\"" + uniqueDataLast20[i][0] + "\", \"" + uniqueDataLast20[i][1] + "\")'>Imprimir</button></td>";*/
  
  tableHtml += "</tr>";
    }
    tableHtml += "</tbody>";
    tableHtml += "</table>";

  console.log(uniqueDataLast20)

  return tableHtml;
  
}

/////////////////////////////////////

function searchEmb(searchEmbala) {
  var ss = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1GTh9iA3VLibtxmFJjjyP8BbS7B_LVopleQ0SFDAS37k/edit#gid=1174848203")
  var sheet2 = ss.getSheetByName("BS EMB")
  var range2 = sheet2.getDataRange().getValues(); // get all the data in the sheet
  console.log(searchEmb);
  searchEmbala = searchEmbala.trim(); // remove any spaces at the beginning and end of the search value
  
  for (var i = 0; i < range2.length; i++) {
    var embSearch = range2[i][0].trim(); // remove any spaces at the beginning and end of the value in the spreadsheet
    if (embSearch == searchEmbala) {
      var columnT1 = range2[i][1];
      
      SpreadsheetApp.flush();
      return { columnT1 };
    }
  }
  return "Não encontrado";
}

function createDropdownEmb() {
    var ss = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1GTh9iA3VLibtxmFJjjyP8BbS7B_LVopleQ0SFDAS37k/edit#gid=931412639") // .getActiveSheet(); // Acessa a pagina aberta atual da planilha
    var sheet = ss.getSheetByName("BS EMB")
    var data = sheet.getDataRange().getValues();
    var selectEmb = document.getElementById("dropdownEmb");

    for (var i = 0; i < data.length; i++) {
      var optionEmb = document.createElement("option");
      optionEmb.text = data[i][0];
      selectEmb.add(option);
    }
  }

  function getDataEmb() {
    var ss = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1GTh9iA3VLibtxmFJjjyP8BbS7B_LVopleQ0SFDAS37k/edit#gid=1926743166") // .getActiveSheet(); // Acessa a pagina aberta atual da planilha
    var sheet = ss.getSheetByName("BS EMB")
    var dataEmb = sheet.getRange(3,1,sheet.getLastRow(),1).getValues();
    return dataEmb;
  }


  function AddRegistroEmb(nulo, embalageId, descricaoId,  qtx, nota, chave, sbaf, ajuste, dataHora, timeFiscal) {
  var url = 'https://docs.google.com/spreadsheets/d/1h4KXDgkBLBQ3SBm4d82uJDEzy5wOYdVPLwEhb15Fq1A/edit#gid=0';
  var ss = SpreadsheetApp.openByUrl(url);
  var webAppSheet = ss.getSheetByName("FX");
  var newRow = [nulo, embalageId, descricaoId,  qtx, nota, chave, sbaf, ajuste, dataHora, timeFiscal]
  webAppSheet.appendRow(newRow);
}

////////////////////////////////////888888888888888888

function searchEmbLoja(searchEmbalaLoja) {
  var ssLoja = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/15_iesDxT6NbdajRsGkn4peDgQQOvBf-kg42wt2A2Xk4/edit#gid=0")
  var sheetLoja = ssLoja.getSheetByName("bd")
  var rangeLoja = sheetLoja.getDataRange().getValues(); // get all the data in the sheet
  
  searchEmbalaLoja = searchEmbalaLoja.trim(); // remove any spaces at the beginning and end of the search value
  
  for (var i = 0; i < rangeLoja.length; i++) {
    var embSearchLoja = rangeLoja[i][0].trim(); // remove any spaces at the beginning and end of the value in the spreadsheet
    if (embSearchLoja == searchEmbalaLoja) {
      var columnT1 = rangeLoja[i][1];
      
      SpreadsheetApp.flush();
      return { columnT1 };
    }
  }
  return "Não encontrado";
}

function createDropdownEmbLoja() {
    var ssLoja = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/15_iesDxT6NbdajRsGkn4peDgQQOvBf-kg42wt2A2Xk4/edit#gid=0") // .getActiveSheet(); // Acessa a pagina aberta atual da planilha
    var sheetLoja = ssLoja.getSheetByName("bd")
    var dataLoja = sheetLoja.getDataRange().getValues();
    var selectEmbLoja = document.getElementById("dropdownEmbLoja");

    for (var i = 0; i < dataLoja.length; i++) {
      var optionEmbLoja = document.createElement("option");
      optionEmbLoja.text = dataLoja[i][0];
      selectEmbLoja.add(option);
    }
  }

  function getDataEmbLoja() {
    var ss = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/15_iesDxT6NbdajRsGkn4peDgQQOvBf-kg42wt2A2Xk4/edit#gid=0") // .getActiveSheet(); // Acessa a pagina aberta atual da planilha
    var sheet = ss.getSheetByName("bd")
    var dataEmbLoja = sheet.getRange(2,1,sheet.getLastRow(),1).getValues();
    return dataEmbLoja;
  }


  function AddRegistroEmbLoja(nulo, embalageId, descricaoId,  qtx, dataHora, timeFiscal, nulo2, area) {
  var url = 'https://docs.google.com/spreadsheets/d/15_iesDxT6NbdajRsGkn4peDgQQOvBf-kg42wt2A2Xk4/edit#gid=0';
  var ss = SpreadsheetApp.openByUrl(url);
  var webAppSheet = ss.getSheetByName("REGISTROS");
  var newRow = [nulo, embalageId, descricaoId,  qtx, dataHora, timeFiscal, nulo2, area]
  webAppSheet.appendRow(newRow);
}

function buscarDadosPlanilha(codigoForn) {
  var ss = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1sOmk5msu5tT8FY0LLXjIpV0jaO3leBC0c7UQVL3OYjE/edit#gid=1543213291") // .getActiveSheet(); // Acessa a pagina aberta atual da planilha
  var sheet = ss.getSheetByName("ATUAL")
  var dataRange = sheet.getDataRange();
  var values = dataRange.getValues();
  
  var transportadoraSenf = "";
  var placaSenf = "";
  var idComanda = "";
  var idTpSenf = "";
  var lacreSenf = "";
  
  for (var i = 1; i < values.length; i++) { // Começa a partir da segunda linha para ignorar o cabeçalho
    var codigo = values[i][11]; // Coluna L (índice 11)
    var status = values[i][13]; // Coluna N (índice 13)
    var transportadora = values[i][7]; // Coluna H (índice 7)
    var placa = values[i][8]; // Coluna I (índice 8)
    var comanda = values[i][16]; // Coluna Q (índice 16)
    var idTp = values[i][3];
    var lacre = values[i][48];
    var industOuemb = values[i][15];
    
    if (codigo == codigoForn && status !== "FINALIZADO" && industOuemb !== "INDUST." && idTp !== "DESCARGA") {
      transportadoraSenf = transportadora;
      placaSenf = placa;
      idComanda = comanda;
      idTpSenf = idTp;
      lacreSenf = lacre;
      break; // Interrompe o loop após encontrar a primeira correspondência
    }
  }
  
  // Retorna um objeto com os dados encontrados
  return {
    transportadora: transportadoraSenf,
    placa: placaSenf,
    idComanda: idComanda,
    idTpSenf: idTpSenf,
    lacre: lacreSenf
  };
}


function adicionarValores(nossasEmb, nossasQtx, delesEmb, delesQtx, lacre, id, first, placaSenf, codForn, transp, nomeForn, numTp, porcentagem, porcentagemIndust) {

  const lock = LockService.getScriptLock();

  lock.tryLock(50000); //tempo de 1 min para travar a senf não sobrescrever.
  if (lock.hasLock()) {
  var ss = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1Ed2e9aUMJZL4-hbvML42RqsYmfTqZTGCfNOvQ6MAurc/edit#gid=310078354");
  var sheet = ss.getSheetByName('PORTAL');

  var rangeEmbN = sheet.getRange('D19:D39');
  var rangeQuantidadesN = sheet.getRange('C19:C39');
  var rangeEmbD = sheet.getRange('D42:D61');
  var rangeQuantidadesD = sheet.getRange('C42:C61');
  var lugarId = sheet.getRange('G68');
  Utilities.sleep(5000);
  var lugarLacre = sheet.getRange('D89');
  var lugarName = sheet.getRange('E82');
  var lugarCodForn = sheet.getRange('D10');
  var lugarNomeForn = sheet.getRange('D11');
  var lugarPlaca = sheet.getRange('G69');
  var lugarTransp = sheet.getRange('G65');

  SpreadsheetApp.flush();

  var embNData = [];
  var qtxNData = [];
  var embDData = [];
  var qtxDData = [];

  // Cria a matriz com o número correto de linhas para embNData
  for (var i = 0; i < rangeEmbN.getNumRows(); i++) {
    var rowData = [nossasEmb[i] || ''];
    embNData.push(rowData);
  }

  // Cria a matriz com o número correto de linhas para qtxNData
  for (var i = 0; i < rangeQuantidadesN.getNumRows(); i++) {
    var rowData = [nossasQtx[i] || ''];
    qtxNData.push(rowData);
  }

  // Cria a matriz com o número correto de linhas para embDData
  for (var i = 0; i < rangeEmbD.getNumRows(); i++) {
    var rowData = [delesEmb[i] || ''];
    embDData.push(rowData);
  }

  // Cria a matriz com o número correto de linhas para qtxDData
  for (var i = 0; i < rangeQuantidadesD.getNumRows(); i++) {
    var rowData = [delesQtx[i] || ''];
    qtxDData.push(rowData);
  }

  // Define os valores nas faixas de destino
  rangeEmbN.setValues(embNData);

  rangeQuantidadesN.setValues(qtxNData);

  rangeEmbD.setValues(embDData);

  rangeQuantidadesD.setValues(qtxDData);

  lugarId.setValue(id);

  lugarLacre.setValue(lacre);

  lugarName.setValue(first);

  lugarCodForn.setValue(codForn);
 
  lugarNomeForn.setValue(nomeForn);

  lugarPlaca.setValue(placaSenf);

  lugarTransp.setValue(transp);

  SpreadsheetApp.flush();
  Utilities.sleep(1000);
  SpreadsheetApp.flush();

  activeSemImprimir(first, id, codForn, numTp, transp, nomeForn, placaSenf, lacre, porcentagem, porcentagemIndust);

  } lock.releaseLock();
}

function getSheetDataFila() {
  var url = 'https://docs.google.com/spreadsheets/d/1sOmk5msu5tT8FY0LLXjIpV0jaO3leBC0c7UQVL3OYjE/edit#gid=1543213291';
  var ss = SpreadsheetApp.openByUrl(url);

  var sheet = ss.getSheetByName("ATUAL");
  var range = sheet.getDataRange();
  var numRows = range.getNumRows();
  var startRow = Math.max(numRows - 30, 4);
  var data = range.getValues().slice(startRow);

  var filteredData = [];

  for (var i = 0; i < data.length; i++) {
    if (data[i][13] !== 'FINALIZADO' && data[i][13] !== 'DESCUMPRIDO' && data[i][12] !== "" && data[i][3] !== 'DESCARGA' && data[i][11] !== 1017305 && data[i][15] !== 'INDUST.') {
      filteredData.push(data[i][12] + " - " + data[i][11]);
    }
  }

  var formattedData = [];
  for (var j = 0; j < filteredData.length; j++) {
    var rowNumber = j + 1;
    var rowData = rowNumber + ". " + filteredData[j];
    formattedData.push(rowData);
  }

  return formattedData;
}

////////////////////////////////////888888888888888888

var url = 'https://docs.google.com/spreadsheets/d/1Ed2e9aUMJZL4-hbvML42RqsYmfTqZTGCfNOvQ6MAurc/edit#gid=310078354';
var ss = SpreadsheetApp.openByUrl(url);
var linkPlanilha = "1Ed2e9aUMJZL4-hbvML42RqsYmfTqZTGCfNOvQ6MAurc";
var pagEmbalagem = ss.getSheetByName('PORTAL');
var pagIndustrial = ss.getSheetByName('PORTAL INDUST.');
var pagColeta = ss.getSheetByName('PORTAL COLETA');
var data = pagEmbalagem.getRange("G11").getValue();
var idTemplate = "13gv-D49DEQNuqrIJxN50oPAVlH7dO9Ls";
var pastaDestino = DriveApp.getFolderById('13gv-D49DEQNuqrIJxN50oPAVlH7dO9Ls');
var idSheet = SpreadsheetApp.openById("1Ed2e9aUMJZL4-hbvML42RqsYmfTqZTGCfNOvQ6MAurc");
var idNossa = idSheet.getSheetByName("BACKEND")
var checkbox = SpreadsheetApp.newDataValidation().requireCheckbox().setAllowInvalid(false).build();

function generateUID () {
  var ALPHABET = '0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ?$%!@*&¬¢£';
  var rtn = '';
  for (var i = 0; i < 10; i++) {
    rtn += ALPHABET.charAt(Math.floor(Math.random() * ALPHABET.length));
  }
  return rtn;
}

function criarPDF(linkPlanilha, pagEmbalagem, nomeDoPdf, first, id, codfor, numTp, transp, placaSenf, nomeForn, lacre, porcentagem, porcentagemIndust) {
  SpreadsheetApp.flush();

  var nomeDoPdf = codfor + " - " + "NOSSA" + " - " + data + " # " + id

  pagEmbalagem.getRange("E82").setValue(first);
  SpreadsheetApp.flush();
  Utilities.sleep(5000);
  SpreadsheetApp.flush();

  const fr = 0, fc = 0, lc = 9, lr = 100;
  const url = "https://docs.google.com/spreadsheets/d/" + linkPlanilha + "/export" +
    "?format=pdf&" +
    "size=7&" +
    "fzr=true&" +
    "portrait=true&" +
    "fitw=true&" +
    "gridlines=false&" +
    "printtitle=false&" +
    "top_margin=0.25&" +
    "bottom_margin=0.25&" +
    "left_margin=0.3&" +
    "right_margin=0.3&" +
    "sheetnames=false&" +
    "pagenum=UNDEFINED&" +
    "attachment=true&" +
    "gid=" + pagEmbalagem.getSheetId() + '&' +
    "r1=" + fr + "&c1=" + fc + "&r2=" + lr + "&c2=" + lc;

  const params = { method: "GET", headers: { "authorization": "Bearer " + ScriptApp.getOAuthToken() } };
  const blob = UrlFetchApp.fetch(url, params).getBlob().setName(nomeDoPdf + '.pdf');

  // Gets the folder in Drive where the PDFs are stored.
  const pdfFile = pastaDestino.createFile(blob);
  var idPDF = pdfFile.getId();

  var pdf = DriveApp.getFileById(idPDF);

  var ultimaLinhaN = idNossa.getLastRow() // Pega a ultima linha vazia
  var proximaLinhaN = ultimaLinhaN + 1 // Acha a proxima linha vazia


  SpreadsheetApp.flush();

   //var nossaOuDeles = pagEmbalagem.getRange("D12").getValue();
   idNossa.getRange(proximaLinhaN, 2).setValue(first);
   idNossa.getRange(proximaLinhaN, 3).setValue(placaSenf);
   idNossa.getRange(proximaLinhaN, 5).setValue(codfor);
   idNossa.getRange(proximaLinhaN, 6).setValue(new Date()).setNumberFormat ('dd/MM/yy');
   idNossa.getRange(proximaLinhaN, 7).setValue(new Date()).setNumberFormat ('HH:MMam/pm');
   SpreadsheetApp.flush();
   idNossa.getRange(proximaLinhaN, 8).setValue("NOSSA");
   idNossa.getRange(proximaLinhaN, 9).setValue(pdfFile.getUrl());
   idNossa.getRange(proximaLinhaN, 10).setValue(nomeDoPdf);
   idNossa.getRange(proximaLinhaN, 13).setValue(id);
   
  SpreadsheetApp.flush();
  
  enviarEmailPersonalizado(idPDF, id, nomeForn, placaSenf, numTp, transp, first);
  //enviarEmail(idPDF, id, nomeForn, placaSenf, numTp, transp, first);
  atualizarPlanilhaLiberacao(id, first, lacre, pdfFile.getUrl(), "EMB.", porcentagem, "", codfor);

  return idPDF;

}

function getGreeting() {
  var dataAtual = new Date();
  var hora = dataAtual.getHours();

  if (hora < 12) {
    return "Bom dia";
  } else if (hora < 19) {
    return "Boa tarde";
  } else {
    return "Boa noite";
  }
}

function enviarEmail(pdf, idTp, nomeForn, placaSenf, numTp, transp, nomeConf) {

  var emailGrupo = "mdl-embalagem_rcl@whirlpool.com"; // Insira o email do seu grupo aqui
  var nomeGrupo = "Embalagem Whirlpool - RCL"; // Insira o nome do seu grupo aqui
  var dataAtual = new Date();
  var fusoHorario = "America/Sao_Paulo";
  var horaAtual = Utilities.formatDate(dataAtual, fusoHorario, "HH:mm - DD/MM");

  if (transp !== "JSL") {
    return;
  }

  var hora = horaAtual.substr(0, 5);
  var dia = dataAtual.getDate();
  var mes = dataAtual.toLocaleString('pt-BR', { month: 'long' });
  var ano = dataAtual.getFullYear();

  var resultado = hora + " - " + dia + " de " + mes + " de " + ano;
  var saudacao = getGreeting();

  var destinatarios = ["monitoramemto@jsl.com.br", "gcg-programacao_inbound@whirlpool.com", "mdl-embalagem_rcl@whirlpool.com"]; //*/["michel_bonfim@whirlpool.com"];
  var nomesDestinatarios = ["Time JSL", "Time Inbound", "Time de Embalagens"];
  var assunto = "[EMBALAGEM] - Externalização " + nomeForn + " - TP: " + numTp;

  for (var i = 0; i < destinatarios.length; i++) {
    var destinatario = destinatarios[i];
    var nomeDestinatario = nomesDestinatarios[i];
    var corpo = "<div style='display: flex; justify-content: center;'><table style='border: 1px solid #ccc; border-radius: 10px; padding: 10px; width: 50%; background: linear-gradient(to bottom, rgba(0, 123, 255,0.7) 0%, rgba(176, 196, 222, 1) 50%);'><tr><td>";
    corpo += '<div style="text-align: center;">';
    corpo += '<img src="https://www.whirlpooldigitalassets.com/content/dam/business-unit/global-assets/images/wdl-logos/whirlpool/whirlpool-logo-3.png" style="width: 15%;">';
    corpo += '</div>';
    corpo += "<div style='font-size: 25px; color:white; text-align: center;'>" + saudacao + ", " + nomeDestinatario + "!</div>\n\n";
    corpo += "<div style='font-size: 15px; text-align: center;'>" + "Processo de <B>embalagem</B> finalizado." + "</div>\n\n<br>";
    corpo += "<div style='border: 2px solid #f2f2f2; border-radius: 10px; padding: 10px; font-size: 15px; text-align: left;'>" +
              "Número de transporte: <b>" + numTp + "</b> " + "</div>\n\n<br>" + 
              "<div style='border: 2px solid #f2f2f2; border-radius: 10px; padding: 10px; font-size: 15px; text-align: left;'>" +
              "Fornecedor: <b>" + nomeForn + "</b>" + "</div>\n\n<br>" +
              "<div style='border: 2px solid #f2f2f2; border-radius: 10px; padding: 10px; font-size: 15px; text-align: left;'>" +
              "Placa: <i><b>" + placaSenf + "</b></i>" + "</div>\n\n<br>" +
              "<div style='border: 2px solid #f2f2f2; border-radius: 10px; padding: 10px; font-size: 15px; text-align: left;'>" +
              "Hórario: <b>"+resultado + "</b>" + "</div>\n\n<br>"+
              "<div style='border: 2px solid #f2f2f2; border-radius: 10px; padding: 10px; font-size: 15px; text-align: left;'>" +
              "ID de Rastreio: " + "<b>" + idTp + "</b>" + "</div>\n\n<br>" +
              "<div style='border: 2px solid #f2f2f2; border-radius: 10px; padding: 10px; font-size: 15px; text-align: left;'>" +
              "Conferente: " + "<b>" + nomeConf + "</b>" + "</div>\n\n";
              
    var anexos = [];

    if (destinatario !== "monitoramemto@jsl.com.br") {
     var arquivo1 = DriveApp.getFileById(pdf);
      anexos.push(arquivo1.getBlob());
    }

    corpo += "</td></tr><tr><td>";
    corpo += "<br><br>";
    corpo += '<div style="text-align: center; font-size: 16px; ">';
    corpo += '<img src="https://i.imgur.com/JTvXcJ2.png" style="width: 10%;">';
    corpo += "<br>";
    corpo += "Obrigado por fazer parte do nosso processo!"; + "</div>";
    corpo += '<div style="text-align: center; font-size: 12px; ">';
    corpo += "<hr style='border-top: 1px solid #ebebeb;'>"; // Separator line
    corpo += '<i>Esta é uma mensagem <b>automática</b>, qualquer dúvida entrar em contato: mdl-embalagem_rcl@whirlpool.com</i>';
    corpo += '</div></td></tr></table>';

  MailApp.sendEmail({
      to: destinatario,
      replyTo: emailGrupo, // Define o email de grupo como o remetente principal
      name: nomeGrupo, // Define o nome do grupo como o remetente principal
      subject: assunto,
      htmlBody: corpo,
      attachments: anexos,
      from: {
        name: nomeGrupo, // Define o nome do grupo como o remetente visível
        address: emailGrupo // Define o email de grupo como o remetente visível
      }
    });
  }
}

function atualizarteste () {
  atualizarPlanilhaLiberacao("%ptSWlqnH!", "RAFAEL", "666675", "https://drive.google.com/file/d/1vlUcbeb7vlLb_1VHifrxx-wwSAbItqOx/view", "EMB./IND.", '87,00%','',1003924)
}

function atualizarPlanilhaLiberacao(id, first, lacre, link, tipo, porcentagemEmb, porcentagemIndust, fornecedorCarro) {

  var planilhaPesquisa = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1sOmk5msu5tT8FY0LLXjIpV0jaO3leBC0c7UQVL3OYjE/edit#gid=1543213291");
  var abaPesquisa = planilhaPesquisa.getSheetByName("ATUAL");
  var dadosPesquisa = abaPesquisa.getDataRange().getValues();
  
  for (var i = 3; i < dadosPesquisa.length; i++) {
    var idPesquisa = dadosPesquisa[i][16];
    var idFinalizado = dadosPesquisa[i][48] || null;
    var idTeste = dadosPesquisa[i][15];
    var teste1 = parseFloat(dadosPesquisa[i][54]);
    var teste2 = parseFloat(dadosPesquisa[i][55]);


  if (fornecedorCarro == '1006932' && tipo == "EMB.") {
    porcentagemEmb = '100,00%';
  }

    if (id === idPesquisa) {
      console.log("ID1: " + typeof teste1 + teste1 + " - ID2: " + typeof teste2 + teste2)

      if (tipo == "EMB." ) {

        abaPesquisa.getRange(i + 1, 45).setValue(link);

      } else { abaPesquisa.getRange(i + 1, 46).setValue(link); }
      
      if (idTeste !== "EMB./IND.") {

        /*if (fornecedorCarro !== "COLETA") {
            var proximo = encontrarCarros(fornecedorCarro)
        }

        abaPesquisa.getRange(i + 1, 22).setValue(proximo);*/

        console.log(" Não compartilhado ! = " +idFinalizado)
        abaPesquisa.getRange(i + 1, 14).setValue("FINALIZADO");
        abaPesquisa.getRange(i + 1, 15).setDataValidation(checkbox).setValue(true);
        abaPesquisa.getRange(i + 1, 18).setValue(first);
        abaPesquisa.getRange(i + 1, 19).setValue(new Date()).setNumberFormat('HH:MM');
        abaPesquisa.getRange(i + 1, 44).setValue("ENVIADO - Não compartilhado");
        abaPesquisa.getRange(i + 1, 49).setValue(lacre);

        if (tipo == "EMB.") {
        abaPesquisa.getRange(i + 1, 53).setValue(porcentagemEmb);        
        } else {
          abaPesquisa.getRange(i + 1, 54).setValue(porcentagemIndust); 
        }

        //abaPesquisa.getRange(i + 1, 22).setValue("100%");
        abaPesquisa.getRange(i + 1, 52).setValue("CARRETA"); // MUDAR PARA AUTOMATICO DEPOIS

        return;
      }

      if (idTeste === "EMB./IND." && idFinalizado === "" || idFinalizado === null) {
        console.log(" Achou sem = " + idFinalizado)
        abaPesquisa.getRange(i + 1, 18).setValue(first);
        abaPesquisa.getRange(i + 1, 19).setValue(new Date()).setNumberFormat('HH:MM');
        abaPesquisa.getRange(i + 1, 44).setValue("ENVIADO - Primeira vez: " + lacre);
        abaPesquisa.getRange(i + 1, 49).setValue(lacre);

        if (teste1 == 0 && teste2 == 0 ) {
          console.log("Ambos zerados")
        abaPesquisa.getRange(i + 1, 53).setValue(porcentagemEmb);
        abaPesquisa.getRange(i + 1, 54).setValue(porcentagemIndust);
        } else if (teste1 !== "" || teste1 !== 0  && teste2 == "" || teste2 == 0 ) {
           console.log("Primeiro zerado")
        abaPesquisa.getRange(i + 1, 54).setValue(porcentagemIndust);
        } else if (teste2 !== "" || teste2 !== 0   && teste1 == "" || teste1 == 0 ) {
          console.log("Segundo zerado")
        abaPesquisa.getRange(i + 1, 53).setValue(porcentagemEmb);
        }

        /*var proximo = encontrarCarros(fornecedorCarro)
        abaPesquisa.getRange(i + 1, 22).setValue(proximo);*/

        break;
      } else {

        /*var proximo = encontrarCarros(fornecedorCarro)
        abaPesquisa.getRange(i + 1, 22).setValue(proximo);*/

        console.log(" Achou com = " + idFinalizado)
        abaPesquisa.getRange(i + 1, 14).setValue("FINALIZADO");
        abaPesquisa.getRange(i + 1, 15).setDataValidation(checkbox).setValue(true);
        abaPesquisa.getRange(i + 1, 18).setValue(first);
        abaPesquisa.getRange(i + 1, 19).setValue(new Date()).setNumberFormat('HH:MM');
        abaPesquisa.getRange(i + 1, 44).setValue("FINALIZADO - Segunda vez: " + lacre);
        abaPesquisa.getRange(i + 1, 49).setValue(lacre);

        if (teste1 == 0 && teste2 == 0 ) {
          console.log("Ambos zerados")
        abaPesquisa.getRange(i + 1, 53).setValue(porcentagemEmb);
        abaPesquisa.getRange(i + 1, 54).setValue(porcentagemIndust);
        } else if ( teste1 !== 0  && teste2 == 0 ) {
          console.log("Primeiro zerado")
        abaPesquisa.getRange(i + 1, 54).setValue(porcentagemIndust);

        /*if ((porcentagemIndust + teste1) > 1) { //COMPLETAR MAIS TARDE
          var real = porcentagemIndust + teste1;

          console.log((real-1) + " Sobra.")
        }*/

        } else if ( teste2 !== 0   &&  teste1 == 0 ) {
          console.log("Segundo zerado")
        abaPesquisa.getRange(i + 1, 53).setValue(porcentagemEmb);
        }

        


        //abaPesquisa.getRange(i + 1, 49).setValue("100%");
        abaPesquisa.getRange(i + 1, 52).setValue("CARRETA");

        break;
      }
    }
  }


}

function removerSimboloPorcentagem(porcentagem) {
  var valorSemPorcentagem = porcentagem.replace('%', '');
  var valorNumerico = parseFloat(valorSemPorcentagem);
  return valorNumerico;
}

/*
function encontrarCarros(fornecedor) {

  var planilhaPesquisa = SpreadsheetApp.openById('1sOmk5msu5tT8FY0LLXjIpV0jaO3leBC0c7UQVL3OYjE'); // Abre a planilha de pesquisa pelo ID
  var planilhaAtual = planilhaPesquisa.getSheetByName('ATUAL'); // Abre a planilha "ATUAL" da planilha de pesquisa
  var planilhaSemana = SpreadsheetApp.openById('1mQ9--hc0EVZhXtbrQ2CLzRGPTvX9aMXhNcV-abvv1Kk'); // Abre a planilha da semana pelo ID

  var numerosTransporteExistentes = planilhaAtual.getRange("D4:D").getValues().flat();
  var primeiroTransporteEncontrado = null;
  var carrosNaoRegistrados = 0;

  var diaAtual = new Date().getDay(); // Obtém o dia da semana atual (0 = Domingo, 1 = Segunda, ..., 6 = Sábado)

  // Reorganiza o array de dias da semana de acordo com o dia atual
  var diasReorganizados;
  if (diaAtual === 1) {
    diasReorganizados = ["Segunda-feira", "Terça-feira", "Quarta-feira"];
  } else if (diaAtual === 2) {
    diasReorganizados = ["Terça-feira", "Quarta-feira", "Quinta-feira"];
  } else if (diaAtual === 3) {
    diasReorganizados = ["Quarta-feira", "Quinta-feira", "Sexta-feira"];
  } else if (diaAtual === 4) {
    diasReorganizados = ["Quinta-feira", "Sexta-feira", "Sábado"];
  } else if (diaAtual === 5) {
    diasReorganizados = ["Sexta-feira", "Sábado", "Segunda-feira"];
  } else if (diaAtual === 6) {
    diasReorganizados = ["Sábado", "Segunda-feira", "Terça-feira"];
  } else if (diaAtual === 0) {
    diasReorganizados = ["Segunda-feira", "Terça-feira", "Quarta-feira"];
  }

  var today = new Date()
  var dataCompara = Utilities.formatDate(today, 'America/Sao_Paulo', 'dd/MM');

  console.log(diasReorganizados)

  for (var i = 0; i < diasReorganizados.length; i++) {
    var sheet = planilhaSemana.getSheetByName(diasReorganizados[i]); // Seleciona a página do dia da semana

    if (sheet) {
      var dados = sheet.getDataRange().getValues();

      for (var j = 0; j < dados.length; j++) {
        var codigoFornecedor = dados[j][6]; // Coluna G
        var numeroTransporte = dados[j][8]; // Coluna I
        var status = dados[j][15]; // Coluna P;
        var data = dados[j][34]; // Coluna AI formatar para data (12/10)
        var horario = dados[j][35]; // Coluna AJ formatar para hora (12:00)

        var dataFormatada = Utilities.formatDate(new Date(data), 'America/Sao_Paulo', 'dd/MM');
        var horaDate = new Date(horario);
        var horarioFormatado = Utilities.formatDate(horaDate, 'GMT-08', 'HH:mm');

        if (codigoFornecedor == fornecedor && numeroTransporte == 120 && status != "Sem Programação" && status != "-" && status !== "" && !numerosTransporteExistentes.includes(status)) {

            if (dataCompara > dataFormatada) {

              console.log(dataCompara +" ## " + dataFormatada)

              primeiroTransporteEncontrado = "Carro(s) atrasado(s) de " + dataFormatada  + " (" + status + ") -";

              break;

              } else {

            console.log(dataCompara +" ## " + dataFormatada)

          primeiroTransporteEncontrado = "Próximo: " + dataFormatada  + " às " + horarioFormatado + " (" + status + ") -";

          break;

            }
        }
      }

      if (primeiroTransporteEncontrado) {

            for (var i = 0; i < diasReorganizados.length; i++) {
              var sheet = planilhaSemana.getSheetByName(diasReorganizados[i]);

              if (sheet) {
                var dados = sheet.getDataRange().getValues();

                for (var j = 0; j < dados.length; j++) {
                  var codigoFornecedor = dados[j][6];
                  var numeroTransporte = dados[j][8];
                  var status = dados[j][15];

                  if (fornecedor == codigoFornecedor && numeroTransporte == 120 && status != "Sem Programação" && status != "-" && status !== "" && !numerosTransporteExistentes.includes(status)) {
                    carrosNaoRegistrados++;
                  }
                }
              }
            }

        if (carrosNaoRegistrados > 0) {
        primeiroTransporteEncontrado = primeiroTransporteEncontrado + " Rest: " + carrosNaoRegistrados;
        } else { primeiroTransporteEncontrado = primeiroTransporteEncontrado + " Sem rest." }       

      console.log(primeiroTransporteEncontrado)
        // Se o primeiro transporte foi encontrado, retorne-o
        return primeiroTransporteEncontrado;
      }
    }
  }

  // Se nenhum transporte foi encontrado, retorne uma mensagem indicando isso
  return "--";
}*/


function activeSemImprimir(first, id, codForn, numTp, transp, nomeForn, placaSenf, lacre, porcentagem, porcentagemIndust) {

  SpreadsheetApp.flush();

  var nomeDoPdf = codForn + " - " + "NOSSA" + " - " + data + " # " + id


  const lock = LockService.getScriptLock();
  SpreadsheetApp.flush();
  lock.tryLock(1000);
  if (lock.hasLock()) {

    criarPDF("1Ed2e9aUMJZL4-hbvML42RqsYmfTqZTGCfNOvQ6MAurc", pagEmbalagem, nomeDoPdf, first, id, codForn, numTp, transp, placaSenf, nomeForn, lacre, porcentagem, "")
    SpreadsheetApp.flush();
    Utilities.sleep(1000);
    SpreadsheetApp.flush();

    //pagEmbalagem.getRange("D11").clear({contentsOnly: true, skipFilteredRows: true});
    
    pagEmbalagem.getRange("C19:C39").clear({contentsOnly: true, skipFilteredRows: true});
    pagEmbalagem.getRange("C42:C61").clear({contentsOnly: true, skipFilteredRows: true});
    pagEmbalagem.getRange("D19:D39").clear({contentsOnly: true, skipFilteredRows: true});
    pagEmbalagem.getRange("D42:D61").clear({contentsOnly: true, skipFilteredRows: true});
    pagEmbalagem.getRange("G68").clear({contentsOnly: true, skipFilteredRows: true});
    pagEmbalagem.getRange("D89").clear({contentsOnly: true, skipFilteredRows: true});

    //SpreadsheetApp.getActiveSpreadsheet().toast("EM PROCESSO DE REGISTRO AGUARDO!","CARREGANDO...",1);
    //SpreadsheetApp.getActiveSpreadsheet().toast("SENF LANÇADA!","CONCLUÍDO",1);

    pagEmbalagem.getRange("E82").clear({contentsOnly: true, skipFilteredRows: true});

    pagEmbalagem.getRange("D10").clear({contentsOnly: true, skipFilteredRows: true});
    pagEmbalagem.getRange("D11:E11").clear({contentsOnly: true, skipFilteredRows: true});
    pagEmbalagem.getRange("G69").clear({contentsOnly: true, skipFilteredRows: true});
    pagEmbalagem.getRange("G65").clear({contentsOnly: true, skipFilteredRows: true});
    
  } lock.releaseLock();
    SpreadsheetApp.flush();
}


function getOAuthToken() {
  return ScriptApp.getOAuthToken();
}

/**
* creates a folder under a parent folder, and returns it's id. If the folder already exists
* then it is not created and it simply returns the id of the existing one
*/

var bdName = "BD Divisões";
function createOrGetFolder(folderName, parentFolderId) {
  try {
    var foldersIter = DriveApp.getFoldersByName(folderName),
      parentFolder = DriveApp.getFolderById(parentFolderId),
      folder;

    if (parentFolder) {
      if (foldersIter.hasNext()) {
        folder = foldersIter.next();
      } else {
        folder = parentFolder.createFolder(folderName);
      }
    } else {
      throw new Error("Parent Folder with id: " + parentFolderId + " not found");
    }

    return folder.getId();
  } catch (error) {
    return error;
  }
}

//////////////////////////////////UPDATE CRITICO

function aiExport () {

var url = 'https://docs.google.com/spreadsheets/d/1XSNM4TmUAaaS_7Pic_C3d1fHQIHI792V1XVly91ppgg/edit#gid=79616886';
var ss = SpreadsheetApp.openByUrl(url);
var bd = ss.getSheetByName("BD");
var hojeAgora = bd.getRange("B2").getValue();
var ui = SpreadsheetApp.getUi();
var zdiv = ss.getSheetByName("BD ZDIV");
var localZdiv = zdiv.getRange("A2");
var mblb = ss.getSheetByName("MBLB - Report Fornecedor");
var localMblb = mblb.getRange("C2");
var mb51 = ss.getSheetByName("MB51 - Report Fornecedor");
var localMb51 = mb51.getRange("C2");
var copyMb51 = ss.getSheetByName("mb51");
var copyZdiv = ss.getSheetByName("zdiv");
var copyMblb = ss.getSheetByName("mblb");

 zdiv.getRange("A2:M8637").clearContent();
 SpreadsheetApp.flush;
 mblb.getRange("C2:M1011").clearContent();
 SpreadsheetApp.flush;

 Utilities.sleep(1200);
  SpreadsheetApp.flush;
 copyMblb.getRange("A2:K1000").copyTo(localMblb)
 Utilities.sleep(1200);
  SpreadsheetApp.flush;
 copyZdiv.getRange("A6:M5524").copyTo(localZdiv)

  Utilities.sleep(3000);
  SpreadsheetApp.flush;
  ss.deleteSheet(copyZdiv);
  SpreadsheetApp.flush;
  ss.deleteSheet(copyMblb);

  //SpreadsheetApp.getActiveSpreadsheet().rename("Abastecimento de BD's e Report's Embalagem" + " - " + hojeAgora);

  //SpreadsheetApp.getActiveSpreadsheet().toast("\n - \n - \n Concluído!","Status",3);

}

function mainUpdadeCritico() {

  var url = 'https://docs.google.com/spreadsheets/d/1XSNM4TmUAaaS_7Pic_C3d1fHQIHI792V1XVly91ppgg/edit#gid=79616886';
  var ss = SpreadsheetApp.openByUrl(url);
  var copyZdiv = ss.getSheetByName("zdiv");
  var copyMblb = ss.getSheetByName("mblb");

  if (copyZdiv != null) { 
  SpreadsheetApp.flush;
  ss.deleteSheet(copyZdiv); }
  if (copyMblb != null) {
  SpreadsheetApp.flush;
  ss.deleteSheet(copyMblb); }

  //ZDIV
  var nameZ = "ZDIV"
  filename = "zdiv.xls"
  //toast(`Importando ${nameZ} do Google Drive ...`);
  let spreadsheetIdZ = convertExcelToGoogleSheets(filename);
  let importedSheetNameZ = importDataFromSpreadsheetZdiv(spreadsheetIdZ, filename);
  //toast(`${nameZ} IMPORTADO com SUCESSO para o Banco de Dados`);

  SpreadsheetApp.flush;
  Utilities.sleep(2000);

  //MBLB
  var nameM = "MBLB"
  filename = "mblb.xls"
  //toast(`Importando ${nameM} do Google Drive ...`);
  let spreadsheetIdM = convertExcelToGoogleSheets(filename);
  let importedSheetNameM = importDataFromSpreadsheetMblb(spreadsheetIdM, filename);
  //toast(`${nameM} IMPORTADO com SUCESSO para o Banco de Dados`);

  SpreadsheetApp.flush;
  Utilities.sleep(5000);

  delteFile("mblb")
  delteFile("mblb.xls") // change later idiot --'
  delteFile("zdiv") //
  delteFile("zdiv.xls")

  SpreadsheetApp.flush;
  ss.getSheetByName("zdiv").hideSheet();
  SpreadsheetApp.flush;
  ss.getSheetByName("mblb").hideSheet();

  limparECopiarPlanilha();

  return;
}

function convertExcelToGoogleSheets(fileName) {
  let files = DriveApp.getFilesByName(fileName);
  let excelFile = null;
  if(files.hasNext())
    excelFile = files.next();
  else
    return null;
  let blob = excelFile.getBlob();
  let config = {
    title: excelFile.getName(),
    parents: [{id: excelFile.getParents().next().getId()}],
    mimeType: MimeType.GOOGLE_SHEETS
  };
  let spreadsheet = Drive.Files.insert(config, blob);
  return spreadsheet.id;
}

function importDataFromSpreadsheetZdiv(spreadsheetId, sheetName) {
  let url = 'https://docs.google.com/spreadsheets/d/1XSNM4TmUAaaS_7Pic_C3d1fHQIHI792V1XVly91ppgg/edit#gid=79616886';
  let spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  let currentSpreadsheet = SpreadsheetApp.openByUrl(url);
  let newSheet = currentSpreadsheet.insertSheet().setName("zdiv");
  let dataToImport = spreadsheet.getSheetByName(sheetName).getDataRange();
  let range = newSheet.getRange(1,1,dataToImport.getNumRows(), dataToImport.getNumColumns());
  range.setValues(dataToImport.getValues());
  return newSheet.getName();
}

function importDataFromSpreadsheetMblb(spreadsheetId, sheetName) {
  let url = 'https://docs.google.com/spreadsheets/d/1XSNM4TmUAaaS_7Pic_C3d1fHQIHI792V1XVly91ppgg/edit#gid=79616886';
  let spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  let currentSpreadsheet = SpreadsheetApp.openByUrl(url);
  let newSheet = currentSpreadsheet.insertSheet().setName("mblb");
  let dataToImport = spreadsheet.getSheetByName(sheetName).getDataRange();
  let range = newSheet.getRange(1,1,dataToImport.getNumRows(), dataToImport.getNumColumns());
  range.setValues(dataToImport.getValues());
  return newSheet.getName();
}

function delteFile(myFileName) {
  var allFiles, idToDLET, myFolder, rtrnFromDLET, thisFile;

  myFolder = DriveApp.getFolderById('1Um58hjV4wJumsHYlOH7Qu85NYV8te3R2');

  allFiles = myFolder.getFilesByName(myFileName);

  while (allFiles.hasNext()) {//
    thisFile = allFiles.next();
    idToDLET = thisFile.getId();
    //Logger.log('idToDLET: ' + idToDLET);

    rtrnFromDLET = Drive.Files.remove(idToDLET);
  };
};

function limparECopiarPlanilha() {
  var spreadsheetId = "1XSNM4TmUAaaS_7Pic_C3d1fHQIHI792V1XVly91ppgg"; // ID da planilha
  var spreadsheetId2 = "1GTh9iA3VLibtxmFJjjyP8BbS7B_LVopleQ0SFDAS37k"; 
  var sheetName = "zdiv"; // Nome da página sourceSheet1
  var targetSheetName = "1"; // Nome da página de destino "1"
  console.log(limparECopiarPlanilha);
  
  var ss = SpreadsheetApp.openById(spreadsheetId);
  var sourceSheet = ss.getSheetByName(sheetName);
  var targetSheet = ss.getSheetByName(targetSheetName);
  
  // Limpar a planilha de destino "1"
  targetSheet.clear();
  
  // Obter o valor da célula F5 da sourceSheet1
  var checkCell = sourceSheet.getRange("F5").getValue();
  
  // Determinar as colunas para copiar da sourceSheet1 com base na condição
  var copyRange;
  if (checkCell !== "") {
    copyRange = sourceSheet.getRange("C6:E" + sourceSheet.getLastRow()).getValues();
    var additionalRange = sourceSheet.getRange("F6:O" + sourceSheet.getLastRow()).getValues();
    for (var i = 0; i < copyRange.length; i++) {
      copyRange[i] = copyRange[i].concat(additionalRange[i]);
    }
  } else {
    copyRange = sourceSheet.getRange("C6:E" + sourceSheet.getLastRow()).getValues();
    var additionalRange = sourceSheet.getRange("G6:P" + sourceSheet.getLastRow()).getValues();
    for (var i = 0; i < copyRange.length; i++) {
      copyRange[i] = copyRange[i].concat(additionalRange[i]);
    }
  }
  // Copiar os dados para a página de destino "1"
  targetSheet.getRange(1, 1, copyRange.length, copyRange[0].length).setValues(copyRange);
  
  // Copiar dados da sourceSheet2 para a página de destino "2"
  var sourceSheet2 = ss.getSheetByName("mblb"); // Nome da página sourceSheet2
  var copyRange2 = sourceSheet2.getRange("B3:M" + sourceSheet2.getLastRow()).getValues();
  var targetSheet2 = ss.getSheetByName("2"); // Nome da página de destino "2"
  targetSheet2.clear();
  targetSheet2.getRange(1, 1, copyRange2.length, copyRange2[0].length).setValues(copyRange2);
}


///////////////////////////////***************SENF INDUSTRIAL******************//////////////////////////////////
function buscarDadosIndustrial(codigoFornIndust) {
  var ss = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1sOmk5msu5tT8FY0LLXjIpV0jaO3leBC0c7UQVL3OYjE/edit#gid=1543213291") // .getActiveSheet(); // Acessa a pagina aberta atual da planilha
  
  var sheet = ss.getSheetByName("ATUAL");
  var dataRange = sheet.getDataRange();
  var values = dataRange.getValues();
  
  var transportadoraSenfIndust = "";
  var placaSenfIndust = "";
  var idComandaIndust = "";
  var idTpSenfIndust = "";
  var lacreIndustEs = "";

  
  for (var i = 1; i < values.length; i++) { // Começa a partir da segunda linha para ignorar o cabeçalho
    var codigo = values[i][11]; // Coluna L (índice 11)
    var status = values[i][13]; // Coluna N (índice 13)
    var transportadora = values[i][7]; // Coluna H (índice 7)
    var placa = values[i][8]; // Coluna I (índice 8)
    var comanda = values[i][16]; // Coluna Q (índice 16)
    var idTp = values[i][3];
    var lacreIndust = values[i][48];
    var industOuemb = values[i][15];
    
    if (codigo == codigoFornIndust && status !== "FINALIZADO" && industOuemb !== "EMB." && idTp !== "DESCARGA") {
      transportadoraSenfIndust = transportadora;
      placaSenfIndust = placa;
      idComandaIndust = comanda;
      idTpSenfIndust = idTp;
      lacreIndustEs = lacreIndust
      break; // Interrompe o loop após encontrar a primeira correspondência
    }
  }
  
  // Retorna um objeto com os dados encontrados
  return {
    transportadoraIndust: transportadoraSenfIndust,
    placaIndust: placaSenfIndust,
    idComandaIndust: idComandaIndust,
    idTpSenfIndust: idTpSenfIndust,
    lacreIndust: lacreIndustEs
  };
}

function getSheetDataFilaIndust() {
  var url = 'https://docs.google.com/spreadsheets/d/1sOmk5msu5tT8FY0LLXjIpV0jaO3leBC0c7UQVL3OYjE/edit#gid=1543213291';
  var ss = SpreadsheetApp.openByUrl(url);
  var sheet = ss.getSheetByName("ATUAL");

  var range = sheet.getDataRange();
  var numRows = range.getNumRows();
  var startRow = Math.max(numRows - 30, 4);
  var data = range.getValues().slice(startRow);

  var filteredDataIndust = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][13] !== 'FINALIZADO' && data[i][13] !== 'DESCUMPRIDO' && data[i][12] !== "" && data[i][3] !== 'DESCARGA' && data[i][15] !== 'EMB.') {
      filteredDataIndust.push(data[i][12]+" - "+data[i][11]);
      
    }
  }

  var formattedDataIndust = [];
  for (var j = 0; j < filteredDataIndust.length; j++) {
    var rowNumber = j + 1;
    var rowData = rowNumber + " .  " + filteredDataIndust[j];
    formattedDataIndust.push(rowData);
  }

  return formattedDataIndust;
}

function adicionarValoresIndust(embalagensCod, embalagensQtx, materialCod, materialQtx, lacreIndust, id, first, placaIndust, codForn, transp, nomeForn, numTp, pesoBruto, pesoLiquido, volumes, pallets, porcentagem, porcentagemIndust) {

  const lock = LockService.getScriptLock();

  lock.tryLock(40000); //tempo de 1 min para travar a senf não sobrescrever.
  if (lock.hasLock()) {
  var ss = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1Ed2e9aUMJZL4-hbvML42RqsYmfTqZTGCfNOvQ6MAurc/edit#gid=310078354");
  var sheet = ss.getSheetByName('PORTAL INDUST.');


  var rangeEmbN = sheet.getRange('D19:D32');
  var rangeQuantidadesN = sheet.getRange('C19:C32');
  var rangeEmbD = sheet.getRange('D35:D61');
  var rangeQuantidadesD = sheet.getRange('C35:C61');

  var lugarId = sheet.getRange('G67');
  var lugarLacre = sheet.getRange('D88');
  var lugarName = sheet.getRange('E81');
  var lugarCodForn = sheet.getRange('D10');
  var lugarNomeForn = sheet.getRange('D11');
  var lugarPlaca = sheet.getRange('G68');
  Utilities.sleep(5000);
  var lugarTransp = sheet.getRange('G65');

  SpreadsheetApp.flush();

  var lugarPesoBruto = sheet.getRange('G73');
  var lugarPesoLiquido = sheet.getRange('G74');
  var lugarVolume = sheet.getRange('G75');
  var lugarPallets = sheet.getRange('G72');

SpreadsheetApp.flush();

  var embNData = [];
  var qtxNData = [];
  var embDData = [];
  var qtxDData = [];

  // Cria a matriz com o número correto de linhas para embNData
  for (var i = 0; i < rangeEmbN.getNumRows(); i++) {
    var rowData = [materialCod[i] || ''];
    embNData.push(rowData);
  }

  // Cria a matriz com o número correto de linhas para qtxNData
  for (var i = 0; i < rangeQuantidadesN.getNumRows(); i++) {
    var rowData = [materialQtx[i] || ''];
    qtxNData.push(rowData);
  }

  // Cria a matriz com o número correto de linhas para embDData
  for (var i = 0; i < rangeEmbD.getNumRows(); i++) {
    var rowData = [embalagensCod[i] || ''];
    embDData.push(rowData);
  }

  // Cria a matriz com o número correto de linhas para qtxDData
  for (var i = 0; i < rangeQuantidadesD.getNumRows(); i++) {
    var rowData = [embalagensQtx[i] || ''];
    qtxDData.push(rowData);
  }

  // Define os valores nas faixas de destino
  rangeEmbN.setValues(embNData);

  rangeQuantidadesN.setValues(qtxNData);

  rangeEmbD.setValues(embDData);

  rangeQuantidadesD.setValues(qtxDData);

  lugarId.setValue(id);

  lugarLacre.setValue(lacreIndust);

  lugarName.setValue(first);

  lugarCodForn.setValue(codForn);
 
  lugarNomeForn.setValue(nomeForn);

  lugarPlaca.setValue(placaIndust);

  lugarTransp.setValue(transp);

  ///
  lugarPesoBruto.setValue(pesoBruto);

  lugarPesoLiquido.setValue(pesoLiquido);

  lugarVolume.setValue(volumes);

  lugarPallets.setValue(pallets);

  SpreadsheetApp.flush();
  Utilities.sleep(1000);
  SpreadsheetApp.flush();

 activeSemImprimirIndust(first, id, codForn, numTp, transp, nomeForn, placaIndust, lacreIndust, porcentagem, porcentagemIndust);

  } lock.releaseLock();
}

function activeSemImprimirIndust(first, id, codForn, numTp, transp, nomeForn, placaIndust, lacreIndust, porcentagem, porcentagemIndust) {

  SpreadsheetApp.flush();

  var nomeDoPdf = codForn + " - " + "INDUST." + " - " + data + " # " + id

  const lock = LockService.getScriptLock();
  SpreadsheetApp.flush();
  lock.tryLock(1000);
  if (lock.hasLock()) {

    criarPDFIndust("1Ed2e9aUMJZL4-hbvML42RqsYmfTqZTGCfNOvQ6MAurc", pagIndustrial, nomeDoPdf, first, id, codForn, numTp, transp, placaIndust, nomeForn, lacreIndust, "", porcentagemIndust);

    SpreadsheetApp.flush();
    Utilities.sleep(1000);
    SpreadsheetApp.flush();

    //pagEmbalagem.getRange("D11").clear({contentsOnly: true, skipFilteredRows: true});
    
    pagIndustrial.getRange("D19:D32").clear({contentsOnly: true, skipFilteredRows: true});
    pagIndustrial.getRange("C19:C32").clear({contentsOnly: true, skipFilteredRows: true});
    pagIndustrial.getRange("D35:D61").clear({contentsOnly: true, skipFilteredRows: true});
    pagIndustrial.getRange("C35:C61").clear({contentsOnly: true, skipFilteredRows: true});
    pagIndustrial.getRange("G67").clear({contentsOnly: true, skipFilteredRows: true});
    pagIndustrial.getRange("D88").clear({contentsOnly: true, skipFilteredRows: true});

    //SpreadsheetApp.getActiveSpreadsheet().toast("EM PROCESSO DE REGISTRO AGUARDO!","CARREGANDO...",1);
    //SpreadsheetApp.getActiveSpreadsheet().toast("SENF LANÇADA!","CONCLUÍDO",1);

    pagIndustrial.getRange("E81").clear({contentsOnly: true, skipFilteredRows: true});

    pagIndustrial.getRange("D10").clear({contentsOnly: true, skipFilteredRows: true});
    pagIndustrial.getRange("D11").clear({contentsOnly: true, skipFilteredRows: true});
    pagIndustrial.getRange("G68").clear({contentsOnly: true, skipFilteredRows: true});
    pagIndustrial.getRange("G65").clear({contentsOnly: true, skipFilteredRows: true});


    pagIndustrial.getRange("G73").clear({contentsOnly: true, skipFilteredRows: true});
    pagIndustrial.getRange("G74").clear({contentsOnly: true, skipFilteredRows: true});
    pagIndustrial.getRange("G75").clear({contentsOnly: true, skipFilteredRows: true});
    pagIndustrial.getRange("G72").clear({contentsOnly: true, skipFilteredRows: true});

    
  } lock.releaseLock();
    SpreadsheetApp.flush();
}

function getSheetDataFilaIndustConfirm() {
  var url = 'https://docs.google.com/spreadsheets/d/1sOmk5msu5tT8FY0LLXjIpV0jaO3leBC0c7UQVL3OYjE/edit#gid=1543213291';
  var ss = SpreadsheetApp.openByUrl(url);
  var sheet = ss.getSheetByName("ATUAL");

  var range = sheet.getDataRange();
  var numRows = range.getNumRows();
  var startRow = Math.max(numRows - 30, 4);
  var data = range.getValues().slice(startRow);

  var filteredDataIndustConfirm = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][13] !== 'FINALIZADO' && data[i][13] !== 'DESCUMPRIDO' && data[i][12] !== "" && data[i][3] !== 'DESCARGA' && data[i][11] !== 1017305 && data[i][15] !== 'INDUST.' && data[i][15] !== 'EMB./IND.') {
      filteredDataIndustConfirm.push(data[i][12]+" - "+data[i][16]);
      
    }
  }

  var formattedDataIndustConfirm = [];
  for (var j = 0; j < filteredDataIndustConfirm.length; j++) {
    var rowNumber = j + 1;
    var rowData = "  " + rowNumber + " .  " + filteredDataIndustConfirm[j];
    formattedDataIndustConfirm.push(rowData);
  }

  return formattedDataIndustConfirm;
}


function atualizarDadosPlanilhaIndustrialConfirm(planilhaId, sheetNome, colunaChave, chave, colunaAlterar, novoValor) {
  
  var ss = SpreadsheetApp.openById(planilhaId);
  var sheet = ss.getSheetByName(sheetNome);

  var data = sheet.getDataRange().getValues();
  var chaveEncontrada = false;

  for (var i = 0; i < data.length; i++) {
    var valorChave = data[i][colunaChave - 1];
    if (valorChave.toString() === chave.toString()) {
      var colunaAlterarIndex = sheet.getRange(1, colunaAlterar).getColumn();
      sheet.getRange(i + 1, colunaAlterarIndex).setValue(novoValor);
      chaveEncontrada = true;
      break;
    }
  }

  if (chaveEncontrada) {
    console.log("Dados atualizados na planilha com sucesso.");
  } else {
    console.log("Chave não encontrada na planilha.");
  }
}

function criarPDFIndust(linkPlanilha, pagIndustrial, nomeDoPdf, first, id, codfor, numTp, transp, placaSenf, nomeForn, lacre, porcentagem, porcentagemIndust) {
  SpreadsheetApp.flush();

  var nomeDoPdf = codfor + " - " + "INDUST" + " - " + data + " # " + id

  pagIndustrial.getRange("E82").setValue(first);
  SpreadsheetApp.flush();
  Utilities.sleep(1000);
  const fr = 0, fc = 0, lc = 9, lr = 100;
  const url = "https://docs.google.com/spreadsheets/d/" + linkPlanilha + "/export" +
    "?format=pdf&" +
    "size=7&" +
    "fzr=true&" +
    "portrait=true&" +
    "fitw=true&" +
    "gridlines=false&" +
    "printtitle=false&" +
    "top_margin=0.25&" +
    "bottom_margin=0.25&" +
    "left_margin=0.3&" +
    "right_margin=0.3&" +
    "sheetnames=false&" +
    "pagenum=UNDEFINED&" +
    "attachment=true&" +
    "gid=" + pagIndustrial.getSheetId() + '&' +
    "r1=" + fr + "&c1=" + fc + "&r2=" + lr + "&c2=" + lc;

  const params = { method: "GET", headers: { "authorization": "Bearer " + ScriptApp.getOAuthToken() } };
  const blob = UrlFetchApp.fetch(url, params).getBlob().setName(nomeDoPdf + '.pdf');

  // Gets the folder in Drive where the PDFs are stored.
  const pdfFile = pastaDestino.createFile(blob);
  var idPDF = pdfFile.getId();

  var pdf = DriveApp.getFileById(idPDF);

  var ultimaLinhaN = idNossa.getLastRow() // Pega a ultima linha vazia
  var proximaLinhaN = ultimaLinhaN + 1 // Acha a proxima linha vazia


  SpreadsheetApp.flush();

   //var nossaOuDeles = pagEmbalagem.getRange("D12").getValue();
   idNossa.getRange(proximaLinhaN, 2).setValue(first);
   idNossa.getRange(proximaLinhaN, 3).setValue(placaSenf);
   idNossa.getRange(proximaLinhaN, 5).setValue(codfor);
   idNossa.getRange(proximaLinhaN, 6).setValue(new Date()).setNumberFormat ('dd/MM/yy');
   idNossa.getRange(proximaLinhaN, 7).setValue(new Date()).setNumberFormat ('HH:MMam/pm');
   SpreadsheetApp.flush();
   idNossa.getRange(proximaLinhaN, 8).setValue("INDUST.");
   idNossa.getRange(proximaLinhaN, 9).setValue(pdfFile.getUrl());
   idNossa.getRange(proximaLinhaN, 10).setValue(nomeDoPdf);
   idNossa.getRange(proximaLinhaN, 13).setValue(id);
   
  SpreadsheetApp.flush();
  
  enviarEmailPersonalizadoIndustrial(idPDF, id, nomeForn, placaSenf, numTp, transp, first);
  //enviarIndust(idPDF, id, nomeForn, placaSenf, numTp, transp, first);

  atualizarPlanilhaLiberacao(id, first, lacre, pdfFile.getUrl(), "INDUST.","",porcentagemIndust, codfor);

  return idPDF;

}

function enviarIndust (idPDF, idTp, nomeForn, placaSenf, numTp, transp, nomeConf) {
  
  var anexosIndust = [];

  var emailGrupo = "mdl-embalagem_rcl@whirlpool.com"; // Insira o email do seu grupo aqui
  var nomeGrupo = "Industrial Whirlpool - RCL"; // Insira o nome do seu grupo aqui
  var dataAtual = new Date();
  var fusoHorario = "America/Sao_Paulo";
  var horaAtual = Utilities.formatDate(dataAtual, fusoHorario, "HH:mm - DD/MM");

  var hora = horaAtual.substr(0, 5);
  var dia = dataAtual.getDate();
  var mes = dataAtual.toLocaleString('pt-BR', { month: 'long' });
  var ano = dataAtual.getFullYear();

  var resultado = hora + " - " + dia + " de " + mes + " de " + ano;
  var saudacao = getGreeting();

  var destinatarios = ["gcg-programacao_inbound@whirlpool.com", "mdl-embalagem_rcl@whirlpool.com", "thiago_h_leite@whirlpool.com", "leticia_bernardes@whirlpool.com", "paulo_v_ometto@whirlpool.com"]; //*/["michel_bonfim@whirlpool.com"];
  var nomesDestinatarios = ["Time Inbound", "Time de Embalagens", "Thiago", "Letícia", "Paulo"];
  var assunto = "[INDUSTRIAL] - Externalização " + nomeForn + " - TP: " + numTp;

  for (var i = 0; i < destinatarios.length; i++) {
    var destinatario = destinatarios[i];
    var nomeDestinatario = nomesDestinatarios[i];
    var corpo = "<div style='display: flex; justify-content: center;'><table style='border: 1px solid #050505; border-radius: 10px; padding: 10px; width: 50%; background: linear-gradient(to bottom, rgba(0, 123, 255,0.5) 0%, rgba(0, 123, 255, 0.1) 50%);'><tr><td>";
    corpo += '<div style="text-align: center;">';
    corpo += '<img src="https://www.whirlpooldigitalassets.com/content/dam/business-unit/global-assets/images/wdl-logos/whirlpool/whirlpool-logo-3.png" style="width: 15%;" >';
    corpo += '</div>';
    corpo += "<div style='font-size: 25px; color:black; text-align: center;'>" + saudacao + ", " + nomeDestinatario + "!</div>\n\n";
    corpo += "<div style='font-size: 15px; text-align: center;'>" + "Processo de <B>industrialização</B> finalizado." + "</div>\n\n<br>";
    corpo += "<div style='border: 2px solid #ccc; border-radius: 10px; padding: 10px; font-size: 15px; text-align: left;'>" +
              "Número de transporte: <b>" + numTp + "</b> " + "</div>\n\n<br>" + 
              "<div style='border: 2px solid #ccc; border-radius: 10px; padding: 10px; font-size: 15px; text-align: left;'>" +
              "Fornecedor: <b>" + nomeForn + "</b>" + "</div>\n\n<br>" +
              "<div style='border: 2px solid #ccc; border-radius: 10px; padding: 10px; font-size: 15px; text-align: left;'>" +
              "Placa: <i><b>" + placaSenf + "</b></i>" + "</div>\n\n<br>" +
              "<div style='border: 2px solid #ccc; border-radius: 10px; padding: 10px; font-size: 15px; text-align: left;'>" +
              "Hórario: <b>" + resultado + "</b>" + "</div>\n\n<br>"+
              "<div style='border: 2px solid #ccc; border-radius: 10px; padding: 10px; font-size: 15px; text-align: left;'>" +
              "ID de Rastreio: " + "<b>" + idTp + "</b>" + "</div>\n\n<br>" +
              "<div style='border: 2px solid #ccc; border-radius: 10px; padding: 10px; font-size: 15px; text-align: left;'>" +
              "Conferente: " + "<b>" + nomeConf + "</b>" + "</div>\n\n";

    if (destinatario !== "monitoramemto@jsl.com.br") {
     var arquivo1 = DriveApp.getFileById(idPDF);
      anexosIndust.push(arquivo1.getBlob());
    }

    corpo += "</td></tr><tr><td>";
    corpo += "<br><br>";
    corpo += '<div style="text-align: center; font-size: 16px; ">';
    corpo += '<img src="https://i.imgur.com/JTvXcJ2.png" style="width: 10%;">';
    corpo += "<br>";
    corpo += "Obrigado por fazer parte do nosso processo!"; + "</div>";
    corpo += '<div style="text-align: center; font-size: 12px; ">';
    corpo += "<hr style='border-top: 1px solid #ebebeb;'>"; // Separator line
    corpo += '<i>Esta é uma mensagem <b>automática</b>, qualquer dúvida entrar em contato: mdl-embalagem_rcl@whirlpool.com</i>';
    corpo += '</div></td></tr></table>';

  MailApp.sendEmail({
      to: destinatario,
      replyTo: emailGrupo, // Define o email de grupo como o remetente principal
      name: nomeGrupo, // Define o nome do grupo como o remetente principal
      subject: assunto,
      htmlBody: corpo,
      attachments: anexosIndust,
      from: {
        name: nomeGrupo, // Define o nome do grupo como o remetente visível
        address: emailGrupo // Define o email de grupo como o remetente visível
      }
    });
  }
}


function saveToSheetIndireto(conferente, data, hora,codigoFornecedor, nomeFornecedor, tempo, nomeMotorista, telefone, placa, nf, pedido, email, time, pontos,veiculo, problema, tempoEmail, status, entrada,stsemail,emailenv,concl,emailend,conclpor,stsnota, tempoemail, id) {

    idFinal = id + generateUID ();

    status = "AGUARDANDO LANÇAMENTO";
    
    var ss = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/11hDe7eEL7CT-UQxnt0Zx-oKObB6uJlSmxuN5aOFIpvE/edit#gid=0")
    var sheet = ss.getSheetByName("ATUAL");
    sheet.appendRow([conferente, data, hora,codigoFornecedor, nomeFornecedor, tempo, nomeMotorista, telefone, placa.toUpperCase(), nf, pedido, email, time, pontos,veiculo, problema, tempoEmail, status, entrada,stsemail,emailenv,concl,emailend,conclpor,stsnota, tempoemail, idFinal]);

    return true;
}

function enviarEmailPersonalizado(pdf, id, forn, placa, tp, transp, conf) {

  var emailGrupo = "mdl-embalagem_rcl@whirlpool.com"; // Insira o email do seu grupo aqui
  var nomeGrupo = "Embalagem Whirlpool - RCL"; // Insira o nome do seu grupo aqui
  var dataAtual = new Date();
  var fusoHorario = "America/Sao_Paulo";
  var horaAtual = Utilities.formatDate(dataAtual, fusoHorario, "HH:mm - DD/MM");

  var hora = horaAtual.substr(0, 5);
  var dia = dataAtual.getDate();
  var mes = dataAtual.toLocaleString('pt-BR', { month: 'long' });
  var ano = dataAtual.getFullYear();

  var resultado = hora + " - " + dia + " de " + mes + " de " + ano;

  var destinatarios = ["gcg-programacao_inbound@whirlpool.com", "mdl-embalagem_rcl@whirlpool.com"];
  var nomesDestinatarios = ["Inbound", "Packing"];
  var assunto = "[EMBALAGEM] - Externalização " + forn + " - TP: " + tp;

  if (forn == "JAW PLASTICOS"){
            destinatarios.push("pcp@jawplasticos.com.br");
            nomesDestinatarios.push("JAW");
            destinatarios.push("programacao@jawplasticos.com.br");
            nomesDestinatarios.push("JAW");
            destinatarios.push("controle@jawplasticos.com.br");
            nomesDestinatarios.push("JAW");
  }

  if (forn == "ITP SYSTEMS") {
      destinatarios.push("jaqueline.matias@itpsystems.com.br");
      nomesDestinatarios.push("ITP");
      destinatarios.push("evelin@itpsystems.com.br");
      nomesDestinatarios.push("ITP");
      destinatarios.push("fabio@itpsystems.com.br");
      nomesDestinatarios.push("ITP");
  }

  if (forn == "BRASCABOS"){
            destinatarios.push("embalagens@brascabos.com.br");
            nomesDestinatarios.push("BRASCABOS");
            destinatarios.push("dconceicao@brascabos.com.br");
            nomesDestinatarios.push("BRASCABOS");
  }

  if (transp == "JSL") {
            destinatarios.push("monitoramemto@jsl.com.br");
            nomesDestinatarios.push("JSL");
  } 
          
  if (transp == "FTI") {
            destinatarios.push("filial.rc@ftilogistica.com.br");
            nomesDestinatarios.push("FTI");  
  }

  if (transp == "MOTOREX") {
            destinatarios.push("motorex.entregasrapidas@hotmail.com");
            nomesDestinatarios.push("MOTOREX");  
  }

  if (transp == "SULISTA") {
            destinatarios.push("comercial@sulista.com.br");
            nomesDestinatarios.push("SULISTA");  
  }

  if (transp == "NEPOMUCENO") {
            destinatarios.push("enzo.nepomuceno@jsl.com.br");
            nomesDestinatarios.push("NEPOMUCENO");  
  }

  for (var i = 0; i < destinatarios.length; i++) {
    var destinatario = destinatarios[i];
    var nome = nomesDestinatarios[i];

    var htmlTemplate = HtmlService.createTemplateFromFile('email');
    htmlTemplate.NOME = nome;
    htmlTemplate.FORNECEDOR = forn;
    htmlTemplate.NUMEROTRANSPORTE = tp;
    htmlTemplate.HORA = resultado;
    htmlTemplate.ID = id;
    htmlTemplate.PLACA = placa;
    htmlTemplate.CONFERENTE = conf;

    var corpoEmail = htmlTemplate.evaluate().getContent();

    corpoEmail = corpoEmail.replace(/%%NOME%%/g, nome);
    corpoEmail = corpoEmail.replace(/%%FORNECEDOR%%/g, forn);
    corpoEmail = corpoEmail.replace(/%%NUMEROTRANSPORTE%%/g, tp);
    corpoEmail = corpoEmail.replace(/%%HORA%%/g, resultado);
    corpoEmail = corpoEmail.replace(/%%ID%%/g, id);
    corpoEmail = corpoEmail.replace(/%%PLACA%%/g, placa);
    corpoEmail = corpoEmail.replace(/%%CONFERENTE%%/g, conf);

    var anexos = [];

   if (destinatario !== "monitoramemto@jsl.com.br" && destinatario !== "filial.rc@ftilogistica.com.br" && destinatario !== "enzo.nepomuceno@jsl.com.br" && destinatario !== "comercial@sulista.com.br" && destinatario !== "motorex.entregasrapidas@hotmail.com") {
     var arquivo1 = DriveApp.getFileById(pdf);
      anexos.push(arquivo1.getBlob());
    }
   


    MailApp.sendEmail({
      to: destinatario,
      replyTo: emailGrupo, // Define o email de grupo como o remetente principal
      name: nomeGrupo, // Define o nome do grupo como o remetente principal
      subject: assunto,
      htmlBody: corpoEmail,
      attachments: anexos,
      from: {
        name: nomeGrupo, // Define o nome do grupo como o remetente visível
        address: emailGrupo // Define o email de grupo como o remetente visível
      }
    });
  }
}

function enviarEmailPersonalizadoIndustrial(pdf, id, forn, placa, tp, transp, conf) {

  var emailGrupo = "mdl-embalagem_rcl@whirlpool.com"; // Insira o email do seu grupo aqui
  var nomeGrupo = "Industrial Whirlpool - RCL"; // Insira o nome do seu grupo aqui
  var dataAtual = new Date();
  var fusoHorario = "America/Sao_Paulo";
  var horaAtual = Utilities.formatDate(dataAtual, fusoHorario, "HH:mm - DD/MM");

  var hora = horaAtual.substr(0, 5);
  var dia = dataAtual.getDate();
  var mes = dataAtual.toLocaleString('pt-BR', { month: 'long' });
  var ano = dataAtual.getFullYear();

  var resultado = hora + " - " + dia + " de " + mes + " de " + ano;


  var destinatarios = ["gcg-programacao_inbound@whirlpool.com", "mdl-embalagem_rcl@whirlpool.com", "thiago_h_leite@whirlpool.com", "leticia_bernardes@whirlpool.com", "paulo_v_ometto@whirlpool.com", "anthony_silva@whirlpool.com", "diego_f_moura@whirlpool.com", "douglas_d_souza@whirlpool.com", "frank_w_lima@whirlpool.com", "joao_m_santos@whirlpool.com", "leandro_penteado@whirlpool.com", "lucas_aquino1@whirlpool.com", "lucca_rodrigues@whirlpool.com", "rafael_ventura1@whirlpool.com", "rodrigo_mignella@whirlpool.com","andre_p_donato@whirlpool.com"];

  var nomesDestinatarios = ["Inbound", "Packing", "Thiago", "Letícia", "Paulo", "Anthony", "Diego", "Douglas", "Frank", "João", "Leandro", "Lucas", "Lucca", "Rafael", "Rodrigo","Donato"];
  
  var assunto = "[INDUSTRIAL] - Externalização " + forn + " - TP: " + tp;

  if (forn == "JAW PLASTICOS"){
            destinatarios.push("pcp@jawplasticos.com.br");
            nomesDestinatarios.push("JAW");
            destinatarios.push("programacao@jawplasticos.com.br");
            nomesDestinatarios.push("JAW");
            destinatarios.push("controle@jawplasticos.com.br");
            nomesDestinatarios.push("JAW");
  }

  if (forn == "BRASCABOS"){
            destinatarios.push("embalagens@brascabos.com.br");
            nomesDestinatarios.push("BRASCABOS");
            destinatarios.push("dconceicao@brascabos.com.br");
            nomesDestinatarios.push("BRASCABOS");
  }

  if (forn == "ITP SYSTEMS") {
      destinatarios.push("lucas_aquino@whirlpool.com");
      nomesDestinatarios.push("AQUINO");
      destinatarios.push("jaqueline.matias@itpsystems.com.br");
      nomesDestinatarios.push("ITP");
      destinatarios.push("evelin@itpsystems.com.br");
      nomesDestinatarios.push("ITP");
      destinatarios.push("fabio@itpsystems.com.br");
      nomesDestinatarios.push("ITP");
  }

  if (transp == "JSL") {
            destinatarios.push("monitoramemto@jsl.com.br");
            nomesDestinatarios.push("JSL");
  } 
          
  if (transp == "FTI") {
            destinatarios.push("filial.rc@ftilogistica.com.br");
            nomesDestinatarios.push("FTI");  
  }

  if (transp == "MOTOREX") {
            destinatarios.push("motorex.entregasrapidas@hotmail.com");
            nomesDestinatarios.push("MOTOREX");  
  }

  if (transp == "SULISTA") {
            destinatarios.push("comercial@sulista.com.br");
            nomesDestinatarios.push("SULISTA");  
  }

  if (transp == "NEPOMUCENO") {
            destinatarios.push("enzo.nepomuceno@jsl.com.br");
            nomesDestinatarios.push("NEPOMUCENO");
  }



  for (var i = 0; i < destinatarios.length; i++) {
    var destinatario = destinatarios[i];
    var nome = nomesDestinatarios[i];

    var htmlTemplate2 = HtmlService.createTemplateFromFile('email2');
    htmlTemplate2.NOME = nome;
    htmlTemplate2.FORNECEDOR = forn;
    htmlTemplate2.NUMEROTRANSPORTE = tp;
    htmlTemplate2.HORA = resultado;
    htmlTemplate2.ID = id;
    htmlTemplate2.PLACA = placa;
    htmlTemplate2.CONFERENTE = conf;

    var corpoEmail = htmlTemplate2.evaluate().getContent();

    corpoEmail = corpoEmail.replace(/%%NOME%%/g, nome);
    corpoEmail = corpoEmail.replace(/%%FORNECEDOR%%/g, forn);
    corpoEmail = corpoEmail.replace(/%%NUMEROTRANSPORTE%%/g, tp);
    corpoEmail = corpoEmail.replace(/%%HORA%%/g, resultado);
    corpoEmail = corpoEmail.replace(/%%ID%%/g, id);
    corpoEmail = corpoEmail.replace(/%%PLACA%%/g, placa);
    corpoEmail = corpoEmail.replace(/%%CONFERENTE%%/g, conf);

    var anexos = [];

     var arquivo1 = DriveApp.getFileById(pdf);
      anexos.push(arquivo1.getBlob());
   


    MailApp.sendEmail({
      to: destinatario,
      replyTo: emailGrupo, // Define o email de grupo como o remetente principal
      name: nomeGrupo, // Define o nome do grupo como o remetente principal
      subject: assunto,
      htmlBody: corpoEmail,
      attachments: anexos,
      from: {
        name: nomeGrupo, // Define o nome do grupo como o remetente visível
        address: emailGrupo // Define o email de grupo como o remetente visível
      }
    });
  }
}

function getDadosTabela() {
  
  var ss = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1Ed2e9aUMJZL4-hbvML42RqsYmfTqZTGCfNOvQ6MAurc/edit#gid=624273769")
  var sheet = ss.getSheetByName("BD Indust.");
  var dataRange = sheet.getRange("A2:C" + sheet.getLastRow());
  var data = dataRange.getValues();

  var tabelaIndustrial = [];
  data.forEach(function(row) {
    tabelaIndustrial.push({
      "material": row[0].toString(),
      "descricao": row[1].toString(),
      "tipo": row[2].toString()
    });
  });

  return tabelaIndustrial;
}

function getSheetDataAvulso(forn) {
  var url = 'https://docs.google.com/spreadsheets/d/11lJunuQtMkuZnq8DlhFACzqAEqyY-nk4ivrvy1b_GkA/edit#gid=554135346';
  var ss = SpreadsheetApp.openByUrl(url);

  var sheet = ss.getSheetByName("Respostas ao formulário 1");
  var data = sheet.getDataRange().getValues();

  var filteredData = [];

  for (var i = 1; i < data.length; i++) { // Começa a partir da segunda linha (índice 1)
    if (data[i][10] == '' && data[i][2] == forn) {
      filteredData.push(forn + " - " + data[i][6]);
    } else { console.log(data[i][9]) + "FORN: " + data[i][6]}
  }

  var formattedData = [];
  for (var j = 0; j < filteredData.length; j++) {
    var rowNumber = j + 1;
    var rowData = rowNumber + ". " + filteredData[j];
    formattedData.push(rowData);
  }

  return formattedData;
}

function atualizarPlanilhaLiberacaoAvulso(id, chave) {

  var planilhaPesquisa = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/11lJunuQtMkuZnq8DlhFACzqAEqyY-nk4ivrvy1b_GkA/edit#gid=554135346');
  var abaPesquisa = planilhaPesquisa.getSheetByName("Respostas ao formulário 1");
  var dadosPesquisa = abaPesquisa.getDataRange().getValues();
  
  for (var i = 0; i < dadosPesquisa.length; i++) {
    var idPesquisa = dadosPesquisa[i][6];

    console.log(dadosPesquisa[i][6] + " - " + id)
    
    if (id == idPesquisa) {

      abaPesquisa.getRange(i + 1, 9).setValue(chave);
      abaPesquisa.getRange(i + 1, 11).setValue("CONFIRMADO");
      abaPesquisa.getRange(i + 1, 13).setValue(new Date()).setNumberFormat('dd/MM/yyyy');

    }
  }
}

function getDadosOcupacao() {
  var planilha = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1GTh9iA3VLibtxmFJjjyP8BbS7B_LVopleQ0SFDAS37k/edit#gid=1926743166"); // Substitua pelo ID da sua planilha
  var aba = planilha.getSheetByName("PROPRIEDADE"); // Substitua pelo nome da aba onde estão os dados

  var data = aba.getDataRange().getValues();
  var arrayDeDados = [];

  for (var i = 1; i < data.length; i++) { // Comece em 1 para evitar o cabeçalho
    var codigo = parseFloat(data[i][1]); // Converte para número
    var multiplo = parseFloat(data[i][8]); // Converte para número
    var metroQuadrado = parseFloat(data[i][9]); // Converte para número



    // Verifique se a conversão foi bem-sucedida antes de adicionar aos dados
    if (!isNaN(codigo) && !isNaN(multiplo) && !isNaN(metroQuadrado)) {
      arrayDeDados.push([codigo, multiplo, metroQuadrado]);
    }

    
  }

  console.log(arrayDeDados)
  return arrayDeDados;
}

function testeNotica( ) {
  verificarSaldoAtendimento(1008396)
}

function verificarSaldoAtendimento(fornecedorAlvo) {
  var today = new Date();
  var hoje = Utilities.formatDate(new Date(today), 'America/Sao_Paulo', 'dd/MM');
  var planilha = SpreadsheetApp.openById("1GTh9iA3VLibtxmFJjjyP8BbS7B_LVopleQ0SFDAS37k"); // Substitua pelo ID da sua planilha
  var aba = planilha.getSheetByName("Tabela"); // Substitua pelo nome da aba desejada
  var dados = aba.getRange("B5:L" + aba.getLastRow()).getValues();

  var propriedade = planilha.getSheetByName("Propriedade").getRange("B5:J" + aba.getLastRow()).getValues(); // Colunas B, I e J na aba "Propriedade"

  var saldoAtendimento = {}; // Um objeto para armazenar as informações dos fornecedores e embalagens
  var arrayRetorn = []; // Crie um array para armazenar as informações a serem retornadas

  for (var i = 0; i < dados.length; i++) {
    var fornecedor = parseFloat(dados[i][5]); // A coluna B contém o fornecedor
    var embalagem = dados[i][6].toString(); // A coluna C contém a embalagem

    if (fornecedor === fornecedorAlvo && (dados[i][1] === "Crítico" || (parseFloat(dados[i][1]) < 3 && parseFloat(dados[i][1]) > 0))) {

      // Verifique se o fornecedor está vazio e a embalagem não é papelão
      if (embalagem !== "Papelão" && embalagem !== "Plástico") {
        // Inicialize as informações do fornecedor e embalagem se não estiverem definidas
        if (!saldoAtendimento[embalagem]) {
          saldoAtendimento[embalagem] = { SaldoAtendimento: 0, ValorPropriedade: 0, PercentualAtendido: "Sem pedido" };
        }

        // Obtenha o saldo de atendimento original
        var saldoOriginal = dados[i][0];
        saldoAtendimento[embalagem].SaldoAtendimento += saldoOriginal;
        saldoAtendimento[embalagem].SaldoEmPosse = saldoOriginal || 0;

        var quantidade = dados[i][2];
        if (quantidade) {
          saldoAtendimento[embalagem].SaldoAtendimento -= quantidade;
          saldoAtendimento[embalagem].SaldoPedidos = quantidade;
        }
      }
    }
  }

  // Calcular o valor da propriedade
  for (var embalagem in saldoAtendimento) {
    var info = saldoAtendimento[embalagem];
    proximo = encontrarCarrosCarregar(fornecedorAlvo);

    // Encontre o código de embalagem na aba "Propriedade"
    for (var k = 0; k < propriedade.length; k++) {
      if (propriedade[k][0] == embalagem) {
        // Se encontrar, calcule o valor da propriedade
        var divisoes = propriedade[k][6]; // Quantidade de divisões na coluna I
        var multiplicacao = propriedade[k][8]; // Valor de multiplicação na coluna J

        if (divisoes > 0 && multiplicacao > 0) {
          valorCalculado = (saldoAtendimento[embalagem].SaldoEmPosse / divisoes) * multiplicacao;
          valorCalculado = Math.abs(valorCalculado);
        }

        // Atualize o valor da propriedade
        saldoAtendimento[embalagem].ValorPropriedade = valorCalculado.toFixed(2);
        saldoAtendimento[embalagem].ValorMultiplo = multiplicacao;

      }
    }

    if (info.ValorPropriedade < info.ValorMultiplo) {ocupacao = info.ValorMultiplo} else {ocupacao = info.ValorPropriedade}
    var retorno = "Fornecedor: " + fornecedorAlvo + "\n  Embalagem: " + embalagem + "\n  Saldo: " + info.SaldoEmPosse + " -  Pedido: " + info.SaldoPedidos + "\n  Saldo de Atendimento: " + info.SaldoAtendimento + "\n  Ocupação de carreta: " + ocupacao + "m².  " + "\n  " + proximo;
    arrayRetorn.push(retorno);
  }

  // Retorna o array de informações
  console.log(arrayRetorn)
  return arrayRetorn;
}

function encontrarCarrosCarregar(fornecedor) {
  var planilhaPesquisa = SpreadsheetApp.openById('1sOmk5msu5tT8FY0LLXjIpV0jaO3leBC0c7UQVL3OYjE'); // Abre a planilha de pesquisa pelo ID
  var planilhaAtual = planilhaPesquisa.getSheetByName("ATUAL"); // Abre a planilha "ATUAL" da planilha de pesquisa
  var planilhaSemana = SpreadsheetApp.openById('1mQ9--hc0EVZhXtbrQ2CLzRGPTvX9aMXhNcV-abvv1Kk'); // Abre a planilha da semana pelo ID

  var diasSemana = ["Segunda-feira", "Terça-feira", "Quarta-feira", "Quinta-feira", "Sexta-feira", "Sábado"]; // Lista de dias da semana

  var numerosTransporteExistentes = planilhaAtual.getRange("D4:D").getValues().flat();
  var primeiroTransporteEncontrado = null;
  var carrosNaoRegistrados = 0;

  var diaAtual = new Date().getDay(); // Obtém o dia da semana atual (0 = Domingo, 1 = Segunda, ..., 6 = Sábado)

  // Reorganiza o array de dias da semana de acordo com o dia atual
  var diasReorganizados = diasSemana.slice(diaAtual-1).concat(diasSemana.slice(0, diaAtual-5));

  console.log(diasReorganizados)

  for (var i = 0; i < diasReorganizados.length; i++) {
    var sheet = planilhaSemana.getSheetByName(diasReorganizados[i]); // Seleciona a página do dia da semana
    
    if (sheet) {
      var dados = sheet.getDataRange().getValues();

      for (var j = 0; j < dados.length; j++) {
        var codigoFornecedor = dados[j][6]; // Coluna G
        var numeroTransporte = dados[j][8]; // Coluna I
        var status = dados[j][15]; // Coluna P;
        var data = dados[j][34]; // Coluna AI formatar para data (12/10)
        var horario = dados[j][35]; // Coluna AJ formatar para hora (12:00)

        var dataFormatada = Utilities.formatDate(new Date(data), 'America/Sao_Paulo', 'dd/MM');
        var horaDate = new Date(horario);
        var horarioFormatado = Utilities.formatDate(horaDate, 'GMT-08', 'HH:mm');
        var today = new Date()
        var dataCompara = Utilities.formatDate(new Date(today), 'America/Sao_Paulo', 'dd/MM');

        if (codigoFornecedor == fornecedor && numeroTransporte == 120 && status != "Sem Programação" && status != "-" && status !== "" && !numerosTransporteExistentes.includes(status)) {

            if (dataCompara > dataFormatada) {

              primeiroTransporteEncontrado = "Carro(s) atrasado(s) de " + dataFormatada  + " (" + status + ") -";

              break;

              } else if (dataCompara = dataFormatada) {

              primeiroTransporteEncontrado = "Próximo: HOJE às " + horarioFormatado + " (" + status + ") -";

              break;

              } else {


          primeiroTransporteEncontrado = "Próximo: " + dataFormatada  + " às " + horarioFormatado + " (" + status + ") -";

          break;

            }
        }
      }

      if (primeiroTransporteEncontrado) {

            for (var i = 0; i < diasReorganizados.length; i++) {
              var sheet = planilhaSemana.getSheetByName(diasReorganizados[i]);

              if (sheet) {
                var dados = sheet.getDataRange().getValues();

                for (var j = 0; j < dados.length; j++) {
                  var codigoFornecedor = dados[j][6];
                  var numeroTransporte = dados[j][8];
                  var status = dados[j][15];

                  if (fornecedor == codigoFornecedor && numeroTransporte == 120 && status != "Sem Programação" && status != "-" && status !== "" && !numerosTransporteExistentes.includes(status)) {
                    carrosNaoRegistrados++;
                  }
                }
              }
            }

        if (carrosNaoRegistrados > 0) {
        primeiroTransporteEncontrado = primeiroTransporteEncontrado + " " + carrosNaoRegistrados + " Carro(s) restante(s)!";
        } else { primeiroTransporteEncontrado = primeiroTransporteEncontrado + " Nenhum programação!" }       

      console.log(primeiroTransporteEncontrado)
        // Se o primeiro transporte foi encontrado, retorne-o
        return primeiroTransporteEncontrado;
      }
    }
  }

  // Se nenhum transporte foi encontrado, retorne uma mensagem indicando isso
  console.log("nada")
  return "  Nenhuma programação de carro encontrada!";
}

function adicionarQuantidade(items) {
  var planilha = SpreadsheetApp.openById("1XSNM4TmUAaaS_7Pic_C3d1fHQIHI792V1XVly91ppgg");
  var guia = planilha.getSheetByName("2");

  for (var fornecedor in items) {
    if (items.hasOwnProperty(fornecedor)) {
      for (var codigo in items[fornecedor]) {
        if (items[fornecedor].hasOwnProperty(codigo)) {
          var quantidadeParaAdicionar = items[fornecedor][codigo].Quantidade;
          // Obtenha os valores da guia
          var dados = guia.getRange("A2:G" + guia.getLastRow()).getValues();

          for (var j = 0; j < dados.length; j++) {
            var linha = dados[j];
            if (linha[0] == fornecedor && linha[3] == codigo) { // Coluna A e D
              var celula = guia.getRange(j + 2, 7); // Coluna G (j + 2 porque começamos em A2)
              var valorAtual = celula.getValue();
              var novoValor = parseFloat(valorAtual) + quantidadeParaAdicionar;
              celula.setValue(novoValor);
              console.log(fornecedor + " - " + codigo + " - " + novoValor)
              break;
            }
          }
        }
      }
    }
  }
}

function checkLogout(chaveAtual) {
  var planilhaURL = "https://docs.google.com/spreadsheets/d/11ZJ1GCIY5IL6-_GPAgFvxBF589uX8M26iXs_wsWG4ic/edit#gid=798358046"; // Substitua pelo URL da planilha desejada
  var planilha = SpreadsheetApp.openByUrl(planilhaURL); // Abre a planilha pelo URL
  var planilhaAtiva = planilha.getSheetByName('CHAVE'); // Obtém a planilha pelo nome

  var chaveExistente = planilhaAtiva.getRange("A1").getValue();


  if (chaveAtual === chaveExistente) {
    console.log('chaves Iguais')
    return 'TRUE';
  } else { 
    
    console.log('chaves diferentes')
    return 'FALSE' }

}

function pegarChaveInicial () {
  var planilhaURL = "https://docs.google.com/spreadsheets/d/11ZJ1GCIY5IL6-_GPAgFvxBF589uX8M26iXs_wsWG4ic/edit#gid=798358046"; // Substitua pelo URL da planilha desejada
  var planilha = SpreadsheetApp.openByUrl(planilhaURL); // Abre a planilha pelo URL
  var planilhaAtiva = planilha.getSheetByName('CHAVE'); // Obtém a planilha pelo nome

  var chaveExistente = planilhaAtiva.getRange("A1").getValue();

  chaveExistente = chaveExistente.toString();
  
  console.log(chaveExistente)
  return chaveExistente;

}

function getSheetDataFilaColeta() {
  var url = 'https://docs.google.com/spreadsheets/d/1sOmk5msu5tT8FY0LLXjIpV0jaO3leBC0c7UQVL3OYjE/edit#gid=1543213291';
  var ss = SpreadsheetApp.openByUrl(url);

  var sheet = ss.getSheetByName("ATUAL");
  var range = sheet.getDataRange();
  var numRows = range.getNumRows();
  var startRow = Math.max(numRows - 30, 4);
  var data = range.getValues().slice(startRow);

  var filteredData = [];

  for (var i = 0; i < data.length; i++) {
    if (data[i][13] !== 'FINALIZADO' && data[i][13] !== 'DESCUMPRIDO' && data[i][12] !== "" && data[i][3] !== 'DESCARGA' && data[i][11] == 1017305) {
      filteredData.push(data[i][12] + " - " + data[i][11]);
    }
  }

  var formattedData = [];
  for (var j = 0; j < filteredData.length; j++) {
    var rowNumber = j + 1;
    var rowData = rowNumber + ". " + filteredData[j];
    formattedData.push(rowData);
  }

  return formattedData;
}

function buscarDadosColeta(codigoFornColeta) {
  var ss = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1sOmk5msu5tT8FY0LLXjIpV0jaO3leBC0c7UQVL3OYjE/edit#gid=1543213291") // .getActiveSheet(); // Acessa a pagina aberta atual da planilha
  
  var sheet = ss.getSheetByName("ATUAL");
  var dataRange = sheet.getDataRange();
  var values = dataRange.getValues();
  
  var transportadoraSenfColeta = "";
  var placaSenfColeta = "";
  var idComandaColeta = "";
  var idTpSenfColeta = "";
  var lacreColeta = "";

  
  for (var i = 1; i < values.length; i++) { // Começa a partir da segunda linha para ignorar o cabeçalho
    var codigo = values[i][11]; // Coluna L (índice 11)
    var status = values[i][13]; // Coluna N (índice 13)
    var transportadora = values[i][7]; // Coluna H (índice 7)
    var placa = values[i][8]; // Coluna I (índice 8)
    var comanda = values[i][16]; // Coluna Q (índice 16)
    var idTp = values[i][3];
    var lacreColeta = values[i][48];
    //var industOuemb = values[i][15];
    
    if (codigo == codigoFornColeta && status !== "FINALIZADO" && idTp !== "DESCARGA") {
      transportadoraSenfColeta = transportadora;
      placaSenfColeta = placa;
      idComandaColeta = comanda;
      idTpSenfColeta = idTp;
      lacreColeta = lacreColeta
      break; // Interrompe o loop após encontrar a primeira correspondência
    }
  }
  
  // Retorna um objeto com os dados encontrados
  return {
    transportadoraColeta: transportadoraSenfColeta,
    placaColeta: placaSenfColeta,
    idComandaColeta: idComandaColeta,
    idTpSenfColeta: idTpSenfColeta,
    lacreColeta: lacreColeta
  };
}

function adicionarColeta(nossasFornC, nossasEmbC, nossasQtxC, delesFornC, delesEmbC, delesQtxC, lacreC, idC, firstC, placaSenfC, transpC, porcentagemC, porcentagemIndustC) {
  const lock = LockService.getScriptLock();
  lock.tryLock(50000);

  if (lock.hasLock()) {
    var ss = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1Ed2e9aUMJZL4-hbvML42RqsYmfTqZTGCfNOvQ6MAurc/edit#gid=310078354");
    var sheet = ss.getSheetByName('PORTAL COLETA');

    var rangeFornNC = sheet.getRange('C19:C61');
    var rangeEmbNC = sheet.getRange('E19:E61');
    var rangeQuantidadesNC = sheet.getRange('D19:D61');

    var rangeFornDC = sheet.getRange('C116:C158');
    var rangeEmbDC = sheet.getRange('E116:E158');
    var rangeQuantidadesDC = sheet.getRange('D116:D158');
    var lugarIdC = sheet.getRange('H68');
    Utilities.sleep(2000);
    var lugarLacreC = sheet.getRange('D89');
    var lugarNameC = sheet.getRange('F82');

    var lugarPlacaC = sheet.getRange('H69');
    var lugarTranspC = sheet.getRange('H65');

    var fornNC = [];
    var embNDataC = [];
    var qtxNDataC = [];
    var fornDC = [];
    var embDDataC = [];
    var qtxDDataC = [];

    var fornAtualNC = ''; // Variável para controlar o fornecedor atual no array NC
    var fornAtualDC = ''; // Variável para controlar o fornecedor atual no array DC

    for (var i = 0; i < nossasFornC.length; i++) {
      if (nossasFornC[i] !== fornAtualNC) {
        adicionarLinhaEmBranco(fornNC, embNDataC, qtxNDataC);
        fornAtualNC = nossasFornC[i];
      }
      fornNC.push([nossasFornC[i] || '']);
      embNDataC.push([nossasEmbC[i] || '']);
      qtxNDataC.push([nossasQtxC[i] || '']);
    }

    for (var i = 0; i < delesFornC.length; i++) {
      if (delesFornC[i] !== fornAtualDC) {
        adicionarLinhaEmBranco(fornDC, embDDataC, qtxDDataC);
        fornAtualDC = delesFornC[i];
      }
      fornDC.push([delesFornC[i] || '']);
      embDDataC.push([delesEmbC[i] || '']);
      qtxDDataC.push([delesQtxC[i] || '']);
    }

    // Preencher com linhas em branco para completar 43 elementos
    while (fornNC.length < 43) {
      adicionarLinhaEmBranco(fornNC, embNDataC, qtxNDataC);
    }
    while (fornDC.length < 43) {
      adicionarLinhaEmBranco(fornDC, embDDataC, qtxDDataC);
    }

    rangeFornNC.setValues(fornNC);
    rangeQuantidadesNC.setValues(qtxNDataC);
    rangeEmbNC.setValues(embNDataC);

    rangeFornDC.setValues(fornDC);
    rangeQuantidadesDC.setValues(qtxDDataC);
    rangeEmbDC.setValues(embDDataC);

    lugarIdC.setValue(idC);
    lugarLacreC.setValue(lacreC);
    lugarNameC.setValue(firstC);
    lugarPlacaC.setValue(placaSenfC);
    lugarTranspC.setValue(transpC);

    Utilities.sleep(5000);
    criarPDFColeta("1Ed2e9aUMJZL4-hbvML42RqsYmfTqZTGCfNOvQ6MAurc", pagColeta, '', firstC, idC, '', '', '', placaSenfC, '', lacreC, porcentagemC, porcentagemIndustC);

    lock.releaseLock();
  }

}

function testedoteste () {
  criarPDFColeta("1Ed2e9aUMJZL4-hbvML42RqsYmfTqZTGCfNOvQ6MAurc", pagColeta, '', "teste", "teste", '', '', '', "teste", '', "teste", "teste", "teste");
}

function adicionarLinhaEmBranco(fornArray, embArray, qtxArray) {
  fornArray.push(['']);
  embArray.push(['']);
  qtxArray.push(['']);
}

function criarPDFColeta(linkPlanilha, pagColeta, nomeDoPdf, first, id, codfor, numTp, transp, placaSenf, nomeForn, lacre, porcentagem, porcentagemIndust) {
  SpreadsheetApp.flush();

  var nomeDoPdf = 1017305 + " - " + "COLETA" + " - " + data + " # " + id

  SpreadsheetApp.flush();
  Utilities.sleep(1000);
  const fr = 0, fc = 0, lc = 9, lr = 200;
  const url = "https://docs.google.com/spreadsheets/d/" + linkPlanilha + "/export" +
    "?format=pdf&" +
    "size=7&" +
    "fzr=true&" +
    "portrait=true&" +
    "fitw=true&" +
    "gridlines=false&" +
    "printtitle=false&" +
    "top_margin=0.25&" +
    "bottom_margin=0.25&" +
    "left_margin=0.3&" +
    "right_margin=0.3&" +
    "sheetnames=false&" +
    "pagenum=UNDEFINED&" +
    "attachment=true&" +
    "gid=" + pagColeta.getSheetId() + '&' +
    "r1=" + fr + "&c1=" + fc + "&r2=" + lr + "&c2=" + lc;

  const params = { method: "GET", headers: { "authorization": "Bearer " + ScriptApp.getOAuthToken() } };
  const blob = UrlFetchApp.fetch(url, params).getBlob().setName(nomeDoPdf + '.pdf');

  // Gets the folder in Drive where the PDFs are stored.
  const pdfFile = pastaDestino.createFile(blob);
  var idPDF = pdfFile.getId();

  var pdf = DriveApp.getFileById(idPDF);

  var ultimaLinhaN = idNossa.getLastRow() // Pega a ultima linha vazia
  var proximaLinhaN = ultimaLinhaN + 1 // Acha a proxima linha vazia


  SpreadsheetApp.flush();

   //var nossaOuDeles = pagEmbalagem.getRange("D12").getValue();
   idNossa.getRange(proximaLinhaN, 2).setValue(first);
   idNossa.getRange(proximaLinhaN, 3).setValue(placaSenf);
   idNossa.getRange(proximaLinhaN, 5).setValue(1017305);
   idNossa.getRange(proximaLinhaN, 6).setValue(new Date()).setNumberFormat ('dd/MM/yy');
   idNossa.getRange(proximaLinhaN, 7).setValue(new Date()).setNumberFormat ('HH:MMam/pm');
   SpreadsheetApp.flush();
   idNossa.getRange(proximaLinhaN, 8).setValue("COLETA");
   idNossa.getRange(proximaLinhaN, 9).setValue(pdfFile.getUrl());
   idNossa.getRange(proximaLinhaN, 10).setValue(nomeDoPdf);
   idNossa.getRange(proximaLinhaN, 13).setValue(id);
   
  SpreadsheetApp.flush();

  atualizarPlanilhaLiberacao(id, first, lacre, pdfFile.getUrl(), "EMB.",porcentagem ,porcentagemIndust , "COLETA");

  return idPDF;

}