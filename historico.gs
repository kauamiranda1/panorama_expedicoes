const WEBHOOK = PropertiesService.getScriptProperties().getProperty('SEATALK_WEBHOOK'); 

function salvarHistoricoTodasOpcoes() { 
 const ss = SpreadsheetApp.getActiveSpreadsheet(); 
 const abaBase = ss.getSheetByName('departed'); 
 let abaHistorico = ss.getSheetByName('Historico'); 

 if (!abaBase) { 
 throw new Error("A aba 'report' não foi encontrada."); 
 } 

 if (!abaHistorico) { 
 abaHistorico = ss.insertSheet('Historico'); 
 } 

 const pw = ss.getSpreadsheetTimeZone() || 'America/Sao_Paulo'; 
 const agora = new Date(); 
 const dataSnapshot = Utilities.formatDate(agora, pw, 'yyyy-MM-dd'); 
 const dataHoraExecucao = Utilities.formatDate(agora, pw, 'yyyy-MM-dd HH:mm:ss'); 

 const celulaFiltro = abaBase.getRange('A1'); 
 const valorOriginal = celulaFiltro.getValue(); 

 const opcoes = obterListaSuspensa(celulaFiltro); 

 if (!opcoes || opcoes.length === 0) { 
 throw new Error("Não foi possível encontrar opções na lista suspensa."); 
 } 

 if (abaHistorico.getLastRow() === 0) { 
 const ultimaColunaTemp = abaBase.getLastColumn(); 
 const cabecalho = abaBase.getRange(1, 1, 1, ultimaColunaTemp).getValues()[0]; 

 abaHistorico.appendRow([ 
 'DATA_REGISTRO', 
 'DATA_HORA_EXECUCAO', 
 'SELECAO', 
 ...cabecalho 
 ]); 
 } 

 let totalLinhas = 0; 

 for (const opcao of opcoes) { 
 celulaFiltro.setValue(opcao); 
 SpreadsheetApp.flush(); 
 Utilities.sleep(2000); 

 const ultimaLinha = abaBase.getLastRow(); 
 const ultimaColuna = abaBase.getLastColumn(); 

 if (ultimaLinha < 2) { 
 continue; 
 } 

 const dados = abaBase.getRange(1, 1, ultimaLinha, ultimaColuna).getValues(); 
 const linhas = dados.slice(1).filter(linha => linha.some(cel => cel !== "")); 

 if (linhas.length === 0) { 
 continue; 
 } 

 const bloco = linhas.map(linha => [ 
 dataSnapshot, 
 dataHoraExecucao, 
 opcao, 
 ...linha 
 ]); 

 abaHistorico.getRange( 
 abaHistorico.getLastRow() + 1, 
 1, 
 bloco.length, 
 bloco[0].length 
 ).setValues(bloco); 

 totalLinhas = totalLinhas + bloco.length; 
 } 

 let msg = ""; 

 if (totalLinhas > 0) { 
  msg = "**Histórico armazenado com sucesso**\nForam adicionadas: **" + totalLinhas + " linhas.**"; 
 } else { 
  msg = "Erro ao executar ou nenhuma linha foi adicionada."; 
 } 

 celulaFiltro.setValue(valorOriginal); 
 SpreadsheetApp.flush(); 
 enviar(msg); 
} 

function obterListaSuspensa(range) { 
 const regra = range.getDataValidation(); 

 if (!regra) { 
 return []; 
 } 

 const tipo = regra.getCriteriaType(); 
 const valores = regra.getCriteriaValues(); 

 if (tipo === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) { 
 return valores[0].filter(v => v !== ""); 
 } 

 if (tipo === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) { 
 const intervalo = valores[0]; 
 const dados = intervalo.getValues().flat().filter(v => v !== ""); 
 return [...new Set(dados)]; 
 } 

 return []; 
} 

function enviar(msg) { 
 UrlFetchApp.fetch(WEBHOOK, { 
 method: "post", 
 contentType: "application/json", 
 payload: JSON.stringify({ 
 tag: "text", 
 text: { content: msg } 
 }), 
 muteHttpExceptions: true 
 }); 
} 
