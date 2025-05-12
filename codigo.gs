/**
 * Radar Manual de Pedidos e Produtos
 * Versão: 1.5.0
 * Data da última atualização: 12/05/2025
 * 
 * Changelog:
 * 1.5.0 - Adição de campo Motivo na seção Devolução de Pedidos
 * 1.4.0 - Adição de campo Telefone nas seções Reembolso e Devolução, reorganização da data e dias de atraso
 * 1.3.0 - Adição dos campos Telefone e Transportadora na seção Pedidos em Atraso
 * 1.2.4 - Correção da função getDevolucoes() e carregarDevolucoes()
 * 1.2.3 - Melhorias nas funções de reembolsos arquivados e devoluções
 * 1.2.2 - Correção da função getReembolsosArquivados() e adição de funções de teste
 * 1.2.1 - Correção da função getDevolucoes() e getDevolucoesArquivadas()
 * 1.2.0 - Correção da função getReembolsosArquivados() e updateReembolso()
 * 1.1.0 - Adição de funcionalidades de reembolso e devolução
 * 1.0.0 - Versão inicial do sistema
 */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Sistema')
    .setTitle('Radar Manual de Pedidos e Produtos')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Radar Manual de Pedidos e Produtos')
    .addItem('Abrir Sistema', 'abrirSistema')
    .addToUi();
}

function abrirSistema() {
  var html = HtmlService.createHtmlOutputFromFile('Sistema')
    .setWidth(1200)
    .setHeight(800)
    .setTitle('Radar Manual de Pedidos e Produtos');
  SpreadsheetApp.getUi().showModalDialog(html, 'Radar Manual de Pedidos e Produtos');
}

// COMPRAS FUTURAS: FUNÇÕES BÁSICAS
function getProdutos() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ComprasFutura');
    if (!sheet) {
      criaPlanilhas();
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ComprasFutura');
    }
    
    if (sheet.getLastRow() <= 1) {
      return [];
    }
    
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
    return data;
  } catch (e) {
    return "ERRO: " + e.toString();
  }
}

function addProduto(nome) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ComprasFutura');
    if (!sheet) {
      criaPlanilhas();
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ComprasFutura');
    }
    
    var id = 1;
    if (sheet.getLastRow() > 1) {
      id = sheet.getRange(sheet.getLastRow(), 1).getValue() + 1;
    }
    
    sheet.appendRow([id, nome, new Date()]);
    return "OK";
  } catch (e) {
    return "ERRO: " + e.toString();
  }
}

function deleteProduto(id) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ComprasFutura');
    if (!sheet) return "Planilha não encontrada";
    
    for (var i = 2; i <= sheet.getLastRow(); i++) {
      if (sheet.getRange(i, 1).getValue() == id) {
        sheet.deleteRow(i);
        return "OK";
      }
    }
    return "Produto não encontrado";
  } catch (e) {
    return "ERRO: " + e.toString();
  }
}

// COMPRAS DO DIA: FUNÇÕES BÁSICAS
function getCompras() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ComprasDoDia');
    if (!sheet) {
      criaPlanilhas();
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ComprasDoDia');
    }
    
    if (sheet.getLastRow() <= 1) {
      return [];
    }
    
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
    return data;
  } catch (e) {
    return "ERRO: " + e.toString();
  }
}

function addCompra(produtoId, nomeProduto) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ComprasDoDia');
    if (!sheet) {
      criaPlanilhas();
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ComprasDoDia');
    }
    
    var id = 1;
    if (sheet.getLastRow() > 1) {
      id = sheet.getRange(sheet.getLastRow(), 1).getValue() + 1;
    }
    
    sheet.appendRow([id, produtoId, nomeProduto]);
    return "OK";
  } catch (e) {
    return "ERRO: " + e.toString();
  }
}

function deleteCompra(id) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ComprasDoDia');
    if (!sheet) return "Planilha não encontrada";
    
    for (var i = 2; i <= sheet.getLastRow(); i++) {
      if (sheet.getRange(i, 1).getValue() == id) {
        sheet.deleteRow(i);
        return "OK";
      }
    }
    return "Compra não encontrada";
  } catch (e) {
    return "ERRO: " + e.toString();
  }
}

// PRODUTOS DESATIVADOS: FUNÇÕES BÁSICAS
function getProdutosDesativados() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ProdutosDesativados');
    if (!sheet) {
      criaPlanilhas();
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ProdutosDesativados');
    }
    
    if (sheet.getLastRow() <= 1) {
      return [];
    }
    
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
    return data;
  } catch (e) {
    return "ERRO: " + e.toString();
  }
}

function addProdutoDesativado(produtoId, nomeProduto) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ProdutosDesativados');
    if (!sheet) {
      criaPlanilhas();
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ProdutosDesativados');
    }
    
    var id = 1;
    if (sheet.getLastRow() > 1) {
      id = sheet.getRange(sheet.getLastRow(), 1).getValue() + 1;
    }
    
    sheet.appendRow([id, produtoId, nomeProduto]);
    return "OK";
  } catch (e) {
    return "ERRO: " + e.toString();
  }
}

function deleteProdutoDesativado(id) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ProdutosDesativados');
    if (!sheet) return "Planilha não encontrada";
    
    for (var i = 2; i <= sheet.getLastRow(); i++) {
      if (sheet.getRange(i, 1).getValue() == id) {
        sheet.deleteRow(i);
        return "OK";
      }
    }
    return "Produto desativado não encontrado";
  } catch (e) {
    return "ERRO: " + e.toString();
  }
}

// PEDIDOS EM ATRASO: FUNÇÕES BÁSICAS
function getPedidosEmAtraso() {
  try {
    Logger.log("Inicio de getPedidosEmAtraso");
    
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('PedidosEmAtraso');
    if (!sheet) {
      Logger.log("Planilha PedidosEmAtraso não encontrada, criando...");
      criaPlanilhas();
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('PedidosEmAtraso');
    }
    
    var lastRow = sheet.getLastRow();
    Logger.log("Última linha com dados: " + lastRow);
    
    if (lastRow <= 1) {
      Logger.log("Nenhum dado encontrado na planilha PedidosEmAtraso");
      return [];
    }
    
    // Verifica se há dados
    var data = sheet.getRange(2, 1, lastRow - 1, 8).getValues(); // Ajustado para 8 colunas (incluindo Telefone e Transportadora)
    Logger.log("Dados brutos obtidos: " + JSON.stringify(data));
    
    // Filtra linhas vazias de forma mais efetiva
    var filteredData = [];
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      // Verifica se a linha tem dados (pelo menos o ID e o número do pedido)
      if (row[0] !== '' && row[0] !== null && row[1] !== '' && row[1] !== null) {
        // Formata a data corretamente
        if (row[6] instanceof Date) { // Ajustado para o novo índice da Data
          // Se for uma data, converte para formato de string
          var date = row[6];
          var day = date.getDate().toString().padStart(2, '0');
          var month = (date.getMonth() + 1).toString().padStart(2, '0');
          var year = date.getFullYear();
          row[6] = day + '/' + month + '/' + year;
        }
        filteredData.push(row);
      }
    }
    
    Logger.log("Dados filtrados: " + JSON.stringify(filteredData));
    return filteredData;
    
  } catch (e) {
    Logger.log("ERRO em getPedidosEmAtraso: " + e.toString());
    Logger.log("Stack trace: " + e.stack);
    throw new Error("ERRO: " + e.toString());
  }
}

function addPedidoEmAtraso(pedido, cliente, telefone, rastreamento, transportadora, dataPrevista, diasAtraso) {
  try {
    Logger.log("Iniciando addPedidoEmAtraso com dados:");
    Logger.log("Pedido: " + pedido);
    Logger.log("Cliente: " + cliente);
    Logger.log("Telefone: " + telefone);
    Logger.log("Rastreamento: " + rastreamento);
    Logger.log("Transportadora: " + transportadora);
    Logger.log("Data Prevista: " + dataPrevista + " (tipo: " + typeof dataPrevista + ")");
    Logger.log("Dias Atraso: " + diasAtraso);
    
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('PedidosEmAtraso');
    if (!sheet) {
      Logger.log("Criando planilha PedidosEmAtraso");
      criaPlanilhas();
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('PedidosEmAtraso');
    }
    
    var id = 1;
    if (sheet.getLastRow() > 1) {
      id = sheet.getRange(sheet.getLastRow(), 1).getValue() + 1;
    }
    
    // Converte a data string para objeto Date
    var dataObj = null;
    if (dataPrevista && typeof dataPrevista === 'string') {
      // Se estiver no formato YYYY-MM-DD, converte para Date
      dataObj = new Date(dataPrevista);
      Logger.log("Data convertida: " + dataObj);
    }
    
    // Garante que diasAtraso seja um número
    var diasAtrasoNum = typeof diasAtraso === 'number' ? diasAtraso : parseInt(diasAtraso) || 0;
    
    var dadosParaInserir = [id, pedido, cliente, telefone, rastreamento, transportadora, dataObj, diasAtrasoNum];
    Logger.log("Dados para inserir: " + JSON.stringify(dadosParaInserir));
    
    sheet.appendRow(dadosParaInserir);
    Logger.log("Linha adicionada com sucesso");
    
    // Força a recalculação da planilha
    SpreadsheetApp.flush();
    
    return "OK";
  } catch (e) {
    Logger.log("ERRO em addPedidoEmAtraso: " + e.toString());
    Logger.log("Stack trace: " + e.stack);
    throw new Error("ERRO: " + e.toString());
  }
}

function updatePedidoEmAtraso(id, pedido, cliente, telefone, rastreamento, transportadora, dataPrevista, diasAtraso) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('PedidosEmAtraso');
    if (!sheet) return "Planilha não encontrada";
    
    // Converte a data string para objeto Date
    var dataObj = null;
    if (dataPrevista && typeof dataPrevista === 'string') {
      dataObj = new Date(dataPrevista);
    }
    
    for (var i = 2; i <= sheet.getLastRow(); i++) {
      if (sheet.getRange(i, 1).getValue() == id) {
        sheet.getRange(i, 2).setValue(pedido);
        sheet.getRange(i, 3).setValue(cliente);
        sheet.getRange(i, 4).setValue(telefone);
        sheet.getRange(i, 5).setValue(rastreamento);
        sheet.getRange(i, 6).setValue(transportadora);
        sheet.getRange(i, 7).setValue(dataObj);
        sheet.getRange(i, 8).setValue(parseInt(diasAtraso));
        return "OK";
      }
    }
    return "Pedido não encontrado";
  } catch (e) {
    return "ERRO: " + e.toString();
  }
}

function deletePedidoEmAtraso(id) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('PedidosEmAtraso');
    if (!sheet) return "Planilha não encontrada";
    
    for (var i = 2; i <= sheet.getLastRow(); i++) {
      if (sheet.getRange(i, 1).getValue() == id) {
        sheet.deleteRow(i);
        return "OK";
      }
    }
    return "Pedido não encontrado";
  } catch (e) {
    return "ERRO: " + e.toString();
  }
}

// REEMBOLSO FINANCEIRO: FUNÇÕES BÁSICAS
function getReembolsos() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ReembolsoFinanceiro');
    if (!sheet) {
      criaPlanilhas();
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ReembolsoFinanceiro');
    }
    
    if (sheet.getLastRow() <= 1) {
      return [];
    }
    
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 10).getValues();
    
    // Formata as datas corretamente e filtra apenas reembolsos não arquivados
    var formattedData = [];
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      if (row[0] !== '' && row[0] !== null && row[9] !== true) { // Verifica se não está arquivado
        // Formata a data se for um objeto Date
        if (row[8] instanceof Date) {
          var date = row[8];
          var day = date.getDate().toString().padStart(2, '0');
          var month = (date.getMonth() + 1).toString().padStart(2, '0');
          var year = date.getFullYear();
          row[8] = day + '/' + month + '/' + year;
        }
        formattedData.push(row.slice(0, 9)); // Retorna apenas as primeiras 9 colunas
      }
    }
    
    return formattedData;
  } catch (e) {
    return "ERRO: " + e.toString();
  }
}

function addReembolso(pedido, cliente, telefone, tipo, motivo, valor, chavePix) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ReembolsoFinanceiro');
    if (!sheet) {
      criaPlanilhas();
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ReembolsoFinanceiro');
    }
    
    var id = 1;
    if (sheet.getLastRow() > 1) {
      id = sheet.getRange(sheet.getLastRow(), 1).getValue() + 1;
    }
    
    var dataAtual = new Date();
    sheet.appendRow([id, pedido, cliente, telefone, tipo, motivo, valor, chavePix, dataAtual, false]);
    return "OK";
  } catch (e) {
    return "ERRO: " + e.toString();
  }
}

function updateReembolso(id, pedido, cliente, telefone, tipo, motivo, valor, chavePix) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ReembolsoFinanceiro');
    if (!sheet) return "Planilha não encontrada";
    
    // Log para debug
    Logger.log("Atualizando reembolso ID: " + id);
    Logger.log("Pedido: " + pedido + ", Cliente: " + cliente + ", Telefone: " + telefone);
    Logger.log("Tipo: " + tipo + ", Motivo: " + motivo);
    Logger.log("Valor: " + valor + ", Chave Pix: " + chavePix);
    
    var encontrou = false;
    
    // Procurar o ID na primeira coluna
    for (var i = 2; i <= sheet.getLastRow(); i++) {
      var idCelula = sheet.getRange(i, 1).getValue();
      Logger.log("Verificando linha " + i + ", ID: " + idCelula);
      
      if (idCelula == id) { // Use == em vez de === para comparar números
        Logger.log("ID encontrado na linha " + i);
        encontrou = true;
        
        // Atualiza os valores
        sheet.getRange(i, 2).setValue(pedido);
        sheet.getRange(i, 3).setValue(cliente);
        sheet.getRange(i, 4).setValue(telefone);
        sheet.getRange(i, 5).setValue(tipo);
        sheet.getRange(i, 6).setValue(motivo);
        sheet.getRange(i, 7).setValue(parseFloat(valor) || 0);
        sheet.getRange(i, 8).setValue(chavePix);
        // Mantém a data original e o status de arquivamento
        
        Logger.log("Reembolso atualizado com sucesso");
        SpreadsheetApp.flush(); // Força a atualização imediata
        return "OK";
      }
    }
    
    if (!encontrou) {
      Logger.log("Reembolso não encontrado com ID: " + id);
      return "Reembolso não encontrado";
    }
  } catch (e) {
    Logger.log("ERRO em updateReembolso: " + e.toString());
    Logger.log("Stack: " + e.stack);
    return "ERRO: " + e.toString();
  }
}

function deleteReembolso(id) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ReembolsoFinanceiro');
    if (!sheet) return "Planilha não encontrada";
    
    for (var i = 2; i <= sheet.getLastRow(); i++) {
      if (sheet.getRange(i, 1).getValue() == id) {
        sheet.deleteRow(i);
        return "OK";
      }
    }
    return "Reembolso não encontrado";
  } catch (e) {
    return "ERRO: " + e.toString();
  }
}

function getReembolsosArquivados() {
  try {
    Logger.log("Iniciando getReembolsosArquivados()");
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ReembolsoFinanceiro');
    if (!sheet) {
      Logger.log("Planilha ReembolsoFinanceiro não encontrada, criando...");
      criaPlanilhas();
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ReembolsoFinanceiro');
    }
    
    var lastRow = sheet.getLastRow();
    Logger.log("Última linha com dados: " + lastRow);
    
    if (lastRow <= 1) {
      Logger.log("Nenhum dado encontrado na planilha ReembolsoFinanceiro");
      return [];
    }
    
    var data = sheet.getDataRange().getValues();
    Logger.log("Dados obtidos: total de " + (data.length-1) + " linhas");
    
    // Filtra apenas reembolsos arquivados e formata datas
    var filteredData = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      // Verifica se a linha tem dados e está arquivada
      if (row[0] !== '' && row[0] !== null && row[9] === true) {
        Logger.log("Reembolso arquivado encontrado na linha " + (i+1) + ": ID=" + row[0]);
        
        // Formata a data de cadastro se for um objeto Date
        var dataCadastro = '';
        if (row[8] instanceof Date) {
          var date = row[8];
          var day = date.getDate().toString().padStart(2, '0');
          var month = (date.getMonth() + 1).toString().padStart(2, '0');
          var year = date.getFullYear();
          dataCadastro = day + '/' + month + '/' + year;
        }
        
        // Formata a data de arquivamento se for um objeto Date
        var dataArquivamento = '';
        if (row[10] instanceof Date) {
          var date = row[10];
          var day = date.getDate().toString().padStart(2, '0');
          var month = (date.getMonth() + 1).toString().padStart(2, '0');
          var year = date.getFullYear();
          dataArquivamento = day + '/' + month + '/' + year;
        }
        
        // Adiciona todos os dados necessários
        filteredData.push([
          row[0],           // ID
          row[1],           // Pedido
          row[2],           // Cliente
          row[3],           // Telefone
          row[4],           // Tipo
          row[5],           // Motivo
          row[6],           // Valor
          row[7],           // Chave PIX
          dataCadastro,     // Data Cadastro (formatada)
          dataArquivamento  // Data Arquivamento (formatada)
        ]);
      }
    }
    
    Logger.log("Reembolsos arquivados encontrados: " + filteredData.length);
    // Exibe detalhes dos primeiros 5 reembolsos arquivados para diagnóstico
    for (var i = 0; i < Math.min(5, filteredData.length); i++) {
      Logger.log("Reembolso arquivado " + i + ": " + JSON.stringify(filteredData[i]));
    }
    
    return filteredData;
  } catch (e) {
    Logger.log("ERRO em getReembolsosArquivados: " + e.toString());
    Logger.log("Stack: " + e.stack);
    return "ERRO: " + e.toString();
  }
}

function arquivarReembolso(id) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ReembolsoFinanceiro');
    if (!sheet) return "Planilha não encontrada";
    
    for (var i = 2; i <= sheet.getLastRow(); i++) {
      if (sheet.getRange(i, 1).getValue() == id) {
        sheet.getRange(i, 10).setValue(true); // Arquivado = true
        sheet.getRange(i, 11).setValue(new Date()); // Data de arquivamento
        return "OK";
      }
    }
    return "Reembolso não encontrado";
  } catch (e) {
    return "ERRO: " + e.toString();
  }
}

function desarquivarReembolso(id) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ReembolsoFinanceiro');
    if (!sheet) return "Planilha não encontrada";
    
    for (var i = 2; i <= sheet.getLastRow(); i++) {
      if (sheet.getRange(i, 1).getValue() == id) {
        sheet.getRange(i, 10).setValue(false); // Arquivado = false
        sheet.getRange(i, 11).setValue(null); // Remove data de arquivamento
        return "OK";
      }
    }
    return "Reembolso não encontrado";
  } catch (e) {
    return "ERRO: " + e.toString();
  }
}

// DEVOLUÇÃO DE PEDIDOS: FUNÇÕES BÁSICAS - CORRIGIDAS
function getDevolucoes() {
  try {
    Logger.log("Iniciando getDevolucoes()");
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DevolucaoPedidos');
    if (!sheet) {
      Logger.log("Planilha DevolucaoPedidos não encontrada, criando...");
      criaPlanilhas();
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DevolucaoPedidos');
    }
    
    var lastRow = sheet.getLastRow();
    Logger.log("Última linha com dados: " + lastRow);
    
    if (lastRow <= 1) {
      Logger.log("Nenhum dado encontrado na planilha DevolucaoPedidos");
      return [];
    }
    
    var data = sheet.getDataRange().getValues();
    Logger.log("Dados obtidos: total de " + (data.length-1) + " linhas");
    
    // Mostra no log os valores da coluna Arquivado (índice 13) para diagnóstico
    for (var i = 1; i < Math.min(5, data.length); i++) {
      Logger.log("Linha " + i + " - Arquivado: " + data[i][13] + " (tipo: " + typeof data[i][13] + ")");
    }
    
    // Filtra apenas devoluções não arquivadas e formata datas
    var filteredData = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      // Modificação na verificação - comparando explicitamente com o booleano false ou a string "FALSE"
      if (row[0] !== '' && row[0] !== null && 
          (row[13] === false || row[13] === "FALSE" || row[13] === 0 || !row[13])) {
        Logger.log("Processando devolução não arquivada na linha " + (i+1) + ": ID=" + row[0]);
        
        // Formata a data de devolução se for um objeto Date
        if (row[5] instanceof Date) {
          var date = row[5];
          var day = date.getDate().toString().padStart(2, '0');
          var month = (date.getMonth() + 1).toString().padStart(2, '0');
          var year = date.getFullYear();
          row[5] = day + '/' + month + '/' + year;
        }
        
        // Formata a data de recebimento se for um objeto Date
        if (row[8] instanceof Date) {
          var date = row[8];
          var day = date.getDate().toString().padStart(2, '0');
          var month = (date.getMonth() + 1).toString().padStart(2, '0');
          var year = date.getFullYear();
          row[8] = day + '/' + month + '/' + year;
        }
        
        // Adiciona todas as colunas relevantes (incluindo Motivo)
        filteredData.push([
          row[0],  // ID
          row[1],  // Pedido
          row[2],  // Cliente
          row[3],  // Telefone
          row[4],  // Produtos
          row[5],  // Data_Devolucao (formatada)
          row[6],  // Quem_Paga
          row[7],  // Motivo
          row[8],  // Data_Recebimento (formatada)
          row[9],  // Parecer_Final
          row[10], // Devolvido_Estoque
          row[11]  // Conferencia_Final
        ]);
      }
    }
    
    Logger.log("Devoluções ativas filtradas: " + filteredData.length);
    return filteredData;
  } catch (e) {
    Logger.log("ERRO em getDevolucoes: " + e.toString());
    Logger.log("Stack: " + e.stack);
    return "ERRO: " + e.toString();
  }
}

function getDevolucoesArquivadas() {
  try {
    Logger.log("Iniciando getDevolucoesArquivadas()");
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DevolucaoPedidos');
    if (!sheet) {
      Logger.log("Planilha DevolucaoPedidos não encontrada, criando...");
      criaPlanilhas();
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DevolucaoPedidos');
    }
    
    var lastRow = sheet.getLastRow();
    Logger.log("Última linha com dados: " + lastRow);
    
    if (lastRow <= 1) {
      Logger.log("Nenhum dado encontrado na planilha DevolucaoPedidos");
      return [];
    }
    
    var data = sheet.getDataRange().getValues();
    Logger.log("Dados obtidos: total de " + (data.length-1) + " linhas");
    
    // Filtra apenas devoluções arquivadas e formata datas
    var filteredData = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (row[0] !== '' && row[0] !== null && row[13] === true) { // Verifica se está arquivado (coluna 14)
        Logger.log("Processando devolução arquivada na linha " + (i+1) + ": ID=" + row[0]);
        
        // Formata a data de devolução
        var dataDevolucao = '';
        if (row[5] instanceof Date) {
          var date = row[5];
          var day = date.getDate().toString().padStart(2, '0');
          var month = (date.getMonth() + 1).toString().padStart(2, '0');
          var year = date.getFullYear();
          dataDevolucao = day + '/' + month + '/' + year;
        }
        
        // Formata a data de arquivamento
        var dataArquivamento = '';
        if (row[14] instanceof Date) {
          var date = row[14];
          var day = date.getDate().toString().padStart(2, '0');
          var month = (date.getMonth() + 1).toString().padStart(2, '0');
          var year = date.getFullYear();
          dataArquivamento = day + '/' + month + '/' + year;
        }
        
        // Retorna as colunas necessárias para devoluções arquivadas (incluindo Motivo)
        filteredData.push([
          row[0],            // ID
          row[1],            // Pedido
          row[2],            // Cliente
          row[3],            // Telefone
          row[4],            // Produtos
          dataDevolucao,     // Data Devolução
          row[7],            // Motivo
          dataArquivamento   // Data Arquivamento
        ]);
      }
    }
    
    Logger.log("Devoluções arquivadas filtradas: " + filteredData.length);
    return filteredData;
  } catch (e) {
    Logger.log("ERRO em getDevolucoesArquivadas: " + e.toString());
    Logger.log("Stack: " + e.stack);
    return "ERRO: " + e.toString();
  }
}

function addDevolucao(pedido, cliente, telefone, produtos, dataDevolucao, quemPaga, motivo, dataRecebimento, parecer, devolvidoEstoque, conferenciaFinal) {
  try {
    Logger.log("Iniciando addDevolucao com dados:");
    Logger.log("Pedido: " + pedido);
    Logger.log("Cliente: " + cliente);
    Logger.log("Telefone: " + telefone);
    Logger.log("Produtos: " + produtos);
    Logger.log("Data Devolução: " + dataDevolucao);
    Logger.log("Quem Paga: " + quemPaga);
    Logger.log("Motivo: " + motivo);
    Logger.log("Data Recebimento: " + dataRecebimento);
    Logger.log("Parecer: " + parecer);
    Logger.log("Devolvido Estoque: " + devolvidoEstoque);
    Logger.log("Conferência Final: " + conferenciaFinal);
    
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DevolucaoPedidos');
    if (!sheet) {
      Logger.log("Criando planilha DevolucaoPedidos");
      criaPlanilhas();
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DevolucaoPedidos');
    }
    
    var id = 1;
    if (sheet.getLastRow() > 1) {
      id = sheet.getRange(sheet.getLastRow(), 1).getValue() + 1;
    }
    
    // Converte as datas string para objetos Date
    var dataDevObj = null;
    var dataRecObj = null;
    
    if (dataDevolucao && typeof dataDevolucao === 'string') {
      dataDevObj = new Date(dataDevolucao);
      Logger.log("Data devolução convertida: " + dataDevObj);
    }
    
    if (dataRecebimento && typeof dataRecebimento === 'string') {
      dataRecObj = new Date(dataRecebimento);
      Logger.log("Data recebimento convertida: " + dataRecObj);
    }
    
    // Preparando os dados para inserção
    var dadosParaInserir = [
      id, 
      pedido, 
      cliente,  // Cliente
      telefone, // Telefone
      produtos, 
      dataDevObj, 
      quemPaga, 
      motivo,   // Motivo
      dataRecObj, 
      parecer, 
      devolvidoEstoque, 
      conferenciaFinal,
      new Date(), // Data de criação
      false, // Arquivado (false)
      null  // Data de arquivamento
    ];
    
    Logger.log("Dados para inserir: " + JSON.stringify(dadosParaInserir));
    
    sheet.appendRow(dadosParaInserir);
    Logger.log("Linha adicionada com sucesso");
    
    // Força a recalculação da planilha
    SpreadsheetApp.flush();
    
    return "OK";
  } catch (e) {
    Logger.log("ERRO em addDevolucao: " + e.toString());
    Logger.log("Stack: " + e.stack);
    return "ERRO: " + e.toString();
  }
}

function updateDevolucao(id, pedido, cliente, telefone, produtos, dataDevolucao, quemPaga, motivo, dataRecebimento, parecer, devolvidoEstoque, conferenciaFinal) {
  try {
    Logger.log("Iniciando updateDevolucao para ID: " + id);
    
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DevolucaoPedidos');
    if (!sheet) {
      Logger.log("Planilha DevolucaoPedidos não encontrada");
      return "Planilha não encontrada";
    }
    
    // Converte as datas string para objetos Date
    var dataDevObj = null;
    var dataRecObj = null;
    
    if (dataDevolucao && typeof dataDevolucao === 'string') {
      dataDevObj = new Date(dataDevolucao);
      Logger.log("Data devolução convertida: " + dataDevObj);
    }
    
    if (dataRecebimento && typeof dataRecebimento === 'string') {
      dataRecObj = new Date(dataRecebimento);
      Logger.log("Data recebimento convertida: " + dataRecObj);
    }
    
    var encontrou = false;
    
    // Procurando o ID na planilha
    for (var i = 2; i <= sheet.getLastRow(); i++) {
      var idCelula = sheet.getRange(i, 1).getValue();
      Logger.log("Verificando linha " + i + ", ID: " + idCelula);
      
      if (idCelula == id) { // Use == em vez de === para comparar números
        Logger.log("ID encontrado na linha " + i);
        encontrou = true;
        
        // Atualiza os valores
        sheet.getRange(i, 2).setValue(pedido);
        sheet.getRange(i, 3).setValue(cliente);  // Cliente
        sheet.getRange(i, 4).setValue(telefone); // Telefone
        sheet.getRange(i, 5).setValue(produtos);
        sheet.getRange(i, 6).setValue(dataDevObj);
        sheet.getRange(i, 7).setValue(quemPaga);
        sheet.getRange(i, 8).setValue(motivo);   // Motivo
        sheet.getRange(i, 9).setValue(dataRecObj);
        sheet.getRange(i, 10).setValue(parecer);
        sheet.getRange(i, 11).setValue(devolvidoEstoque);
        sheet.getRange(i, 12).setValue(conferenciaFinal);
        
        Logger.log("Devolução atualizada com sucesso");
        SpreadsheetApp.flush(); // Força a atualização imediata
        return "OK";
      }
    }
    
    if (!encontrou) {
      Logger.log("Devolução não encontrada com ID: " + id);
      return "Devolução não encontrada";
    }
  } catch (e) {
    Logger.log("ERRO em updateDevolucao: " + e.toString());
    Logger.log("Stack: " + e.stack);
    return "ERRO: " + e.toString();
  }
}

function arquivarDevolucao(id) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DevolucaoPedidos');
    if (!sheet) return "Planilha não encontrada";
    
    for (var i = 2; i <= sheet.getLastRow(); i++) {
      if (sheet.getRange(i, 1).getValue() == id) {
        sheet.getRange(i, 14).setValue(true); // Arquivado = true (coluna 14)
        sheet.getRange(i, 15).setValue(new Date()); // Data de arquivamento (coluna 15)
        return "OK";
      }
    }
    return "Devolução não encontrada";
  } catch (e) {
    return "ERRO: " + e.toString();
  }
}

function desarquivarDevolucao(id) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DevolucaoPedidos');
    if (!sheet) return "Planilha não encontrada";
    
    for (var i = 2; i <= sheet.getLastRow(); i++) {
      if (sheet.getRange(i, 1).getValue() == id) {
        sheet.getRange(i, 14).setValue(false); // Arquivado = false (coluna 14)
        sheet.getRange(i, 15).setValue(null); // Remove data de arquivamento (coluna 15)
        return "OK";
      }
    }
    return "Devolução não encontrada";
  } catch (e) {
    return "ERRO: " + e.toString();
  }
}

function deleteDevolucao(id) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DevolucaoPedidos');
    if (!sheet) return "Planilha não encontrada";
    
    for (var i = 2; i <= sheet.getLastRow(); i++) {
      if (sheet.getRange(i, 1).getValue() == id) {
        sheet.deleteRow(i);
        return "OK";
      }
    }
    return "Devolução não encontrada";
  } catch (e) {
    return "ERRO: " + e.toString();
  }
}

// CRIAR PLANILHAS SE NÃO EXISTIREM
function criaPlanilhas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Cria/Verifica planilha de ComprasFutura
  var comprasFutura = ss.getSheetByName('ComprasFutura');
  if (!comprasFutura) {
    comprasFutura = ss.insertSheet('ComprasFutura');
    comprasFutura.getRange('A1:C1').setValues([['ID', 'Nome', 'Data']]);
  }
  
  // Cria/Verifica planilha de ComprasDoDia
  var comprasDoDia = ss.getSheetByName('ComprasDoDia');
  if (!comprasDoDia) {
    comprasDoDia = ss.insertSheet('ComprasDoDia');
    comprasDoDia.getRange('A1:C1').setValues([['ID', 'ID_Produto', 'Nome_Produto']]);
  }
  
  // Cria/Verifica planilha de ProdutosDesativados
  var produtosDesativados = ss.getSheetByName('ProdutosDesativados');
  if (!produtosDesativados) {
    produtosDesativados = ss.insertSheet('ProdutosDesativados');
    produtosDesativados.getRange('A1:C1').setValues([['ID', 'ID_Produto', 'Nome_Produto']]);
  }
  
  // Cria/Verifica planilha de PedidosEmAtraso
  var pedidosEmAtraso = ss.getSheetByName('PedidosEmAtraso');
  if (!pedidosEmAtraso) {
    pedidosEmAtraso = ss.insertSheet('PedidosEmAtraso');
    pedidosEmAtraso.getRange('A1:H1').setValues([['ID', 'Pedido', 'Cliente', 'Telefone', 'Rastreamento', 'Transportadora', 'Data_Envio', 'Dias_Atraso']]);
  }
  
  // Cria/Verifica planilha de ReembolsoFinanceiro
  var reembolsoFinanceiro = ss.getSheetByName('ReembolsoFinanceiro');
  if (!reembolsoFinanceiro) {
    reembolsoFinanceiro = ss.insertSheet('ReembolsoFinanceiro');
    reembolsoFinanceiro.getRange('A1:K1').setValues([
      ['ID', 'Pedido', 'Cliente', 'Telefone', 'Tipo', 'Motivo', 'Valor', 'Chave_Pix', 'Data_Cadastro', 'Arquivado', 'Data_Arquivamento']
    ]);
  }
  
  // Cria/Verifica planilha de DevolucaoPedidos
  var devolucaoPedidos = ss.getSheetByName('DevolucaoPedidos');
  if (!devolucaoPedidos) {
    devolucaoPedidos = ss.insertSheet('DevolucaoPedidos');
    devolucaoPedidos.getRange('A1:O1').setValues([
      ['ID', 'Pedido', 'Cliente', 'Telefone', 'Produtos', 'Data_Devolucao', 'Quem_Paga', 'Motivo', 'Data_Recebimento', 
       'Parecer_Final', 'Devolvido_Estoque', 'Conferencia_Final', 'Data_Criacao', 'Arquivado', 'Data_Arquivamento']
    ]);
  }
  
  return "Planilhas criadas";
}

// Função para migrar dados das abas antigas para as novas
function migrarDados() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Migrar de Produtos para ComprasFutura
  var produtosSheet = ss.getSheetByName('Produtos');
  if (produtosSheet) {
    var comprasFuturaSheet = ss.getSheetByName('ComprasFutura');
    if (!comprasFuturaSheet) {
      comprasFuturaSheet = ss.insertSheet('ComprasFutura');
      comprasFuturaSheet.getRange('A1:C1').setValues([['ID', 'Nome', 'Data']]);
    }
    
    if (produtosSheet.getLastRow() > 1) {
      var produtosData = produtosSheet.getRange(2, 1, produtosSheet.getLastRow() - 1, 3).getValues();
      for (var i = 0; i < produtosData.length; i++) {
        comprasFuturaSheet.appendRow(produtosData[i]);
      }
    }
  }
  
  // Migrar de Compras para ComprasDoDia
  var comprasSheet = ss.getSheetByName('Compras');
  if (comprasSheet) {
    var comprasDoDiaSheet = ss.getSheetByName('ComprasDoDia');
    if (!comprasDoDiaSheet) {
      comprasDoDiaSheet = ss.insertSheet('ComprasDoDia');
      comprasDoDiaSheet.getRange('A1:C1').setValues([['ID', 'ID_Produto', 'Nome_Produto']]);
    }
    
    if (comprasSheet.getLastRow() > 1) {
      var comprasData = comprasSheet.getRange(2, 1, comprasSheet.getLastRow() - 1, 3).getValues();
      for (var i = 0; i < comprasData.length; i++) {
        comprasDoDiaSheet.appendRow(comprasData[i]);
      }
    }
  }
  
  return "Migração de dados concluída!";
}

// Função para migrar a estrutura das planilhas para a versão 1.4.0
function migrarParaV1_4_0() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. Migrar ReembolsoFinanceiro - Adicionar campo Telefone
    var reembolsoSheet = ss.getSheetByName('ReembolsoFinanceiro');
    if (reembolsoSheet) {
      // Verificar se já tem a estrutura da v1.4.0
      var headers = reembolsoSheet.getRange(1, 1, 1, reembolsoSheet.getLastColumn()).getValues()[0];
      
      // Se não tiver o campo Telefone, adicionar
      if (headers[3] !== 'Telefone') {
        // Criar backup
        var backupSheet = ss.insertSheet('ReembolsoFinanceiro_Backup_v1_3_0');
        reembolsoSheet.getDataRange().copyTo(backupSheet.getRange(1, 1));
        
        // Inserir a nova coluna (Telefone) após a coluna Cliente
        reembolsoSheet.insertColumnAfter(3);
        reembolsoSheet.getRange(1, 4).setValue('Telefone');
        
        // Migrar dados existentes
        if (reembolsoSheet.getLastRow() > 1) {
          for (var i = 2; i <= reembolsoSheet.getLastRow(); i++) {
            reembolsoSheet.getRange(i, 4).setValue(''); // Telefone vazio para registros existentes
          }
        }
        
        Logger.log("Reembolso Financeiro atualizado para v1.4.0");
      }
    }
    
    // 2. Migrar DevolucaoPedidos - Adicionar campo Telefone
    var devolucaoSheet = ss.getSheetByName('DevolucaoPedidos');
    if (devolucaoSheet) {
      // Verificar se já tem a estrutura da v1.4.0
      var headers = devolucaoSheet.getRange(1, 1, 1, devolucaoSheet.getLastColumn()).getValues()[0];
      
      // Se não tiver o campo Telefone, adicionar
      if (headers[3] !== 'Telefone') {
        // Criar backup
        var backupSheet = ss.insertSheet('DevolucaoPedidos_Backup_v1_3_0');
        devolucaoSheet.getDataRange().copyTo(backupSheet.getRange(1, 1));
        
        // Inserir a nova coluna (Telefone) após a coluna Cliente
        devolucaoSheet.insertColumnAfter(3);
        devolucaoSheet.getRange(1, 4).setValue('Telefone');
        
        // Migrar dados existentes
        if (devolucaoSheet.getLastRow() > 1) {
          for (var i = 2; i <= devolucaoSheet.getLastRow(); i++) {
            devolucaoSheet.getRange(i, 4).setValue(''); // Telefone vazio para registros existentes
          }
        }
        
        Logger.log("Devolução de Pedidos atualizado para v1.4.0");
      }
    }
    
    return "Migração para v1.4.0 concluída com sucesso!";
  } catch (e) {
    Logger.log("ERRO em migrarParaV1_4_0: " + e.toString());
    Logger.log("Stack: " + e.stack);
    return "ERRO: " + e.toString();
  }
}

// Função para migrar a estrutura das planilhas para a versão 1.5.0
function migrarParaV1_5_0() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Migrar DevolucaoPedidos - Adicionar campo Motivo
    var devolucaoSheet = ss.getSheetByName('DevolucaoPedidos');
    if (devolucaoSheet) {
      // Verificar se já tem a estrutura da v1.5.0
      var headers = devolucaoSheet.getRange(1, 1, 1, devolucaoSheet.getLastColumn()).getValues()[0];
      
      // Se não tiver o campo Motivo, adicionar
      if (headers[7] !== 'Motivo') {
        // Criar backup
        var backupSheet = ss.insertSheet('DevolucaoPedidos_Backup_v1_4_0');
        devolucaoSheet.getDataRange().copyTo(backupSheet.getRange(1, 1));
        
        // Inserir a nova coluna (Motivo) após a coluna Quem_Paga
        devolucaoSheet.insertColumnAfter(7);
        devolucaoSheet.getRange(1, 8).setValue('Motivo');
        
        // Migrar dados existentes
        if (devolucaoSheet.getLastRow() > 1) {
          for (var i = 2; i <= devolucaoSheet.getLastRow(); i++) {
            devolucaoSheet.getRange(i, 8).setValue(''); // Motivo vazio para registros existentes
          }
        }
        
        Logger.log("Devolução de Pedidos atualizado para v1.5.0");
      }
    }
    
    return "Migração para v1.5.0 concluída com sucesso!";
  } catch (e) {
    Logger.log("ERRO em migrarParaV1_5_0: " + e.toString());
    Logger.log("Stack: " + e.stack);
    return "ERRO: " + e.toString();
  }
}

// FUNÇÕES DE TESTE E DEPURAÇÃO
function testarReembolsosArquivados() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ReembolsoFinanceiro');
    
    if (!sheet) {
      Logger.log("Planilha ReembolsoFinanceiro não encontrada");
      return "Planilha não encontrada";
    }
    
    Logger.log("Última linha com dados: " + sheet.getLastRow());
    Logger.log("Última coluna com dados: " + sheet.getLastColumn());
    
    var dados = sheet.getDataRange().getValues();
    
    Logger.log("Cabeçalhos:");
    Logger.log(JSON.stringify(dados[0]));
    
    if (dados.length <= 1) {
      return "Nenhum reembolso encontrado na planilha";
    }
    
    Logger.log("Todos os reembolsos:");
    var ativos = 0;
    var arquivados = 0;
    
    for (var i = 1; i < dados.length; i++) {
      // Índice 9 = Arquivado (true/false)
      if (dados[i][9] === true) {
        arquivados++;
        Logger.log("Reembolso arquivado " + i + ": " + JSON.stringify(dados[i]));
      } else {
        ativos++;
      }
    }
    
    Logger.log("Coluna 'Arquivado' (índice 9):");
    for (var i = 1; i < Math.min(10, dados.length); i++) {
      Logger.log("Linha " + (i+1) + ", valor arquivado: " + dados[i][9] + ", tipo: " + typeof dados[i][9]);
    }
    
    return "OK - Reembolsos encontrados: " + ativos + " ativos, " + arquivados + " arquivados";
  } catch (e) {
    Logger.log("Erro em testarReembolsosArquivados: " + e.toString());
    return "ERRO: " + e.toString();
  }
}

function testarDevolucoes() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DevolucaoPedidos');
    
    if (!sheet) {
      Logger.log("Planilha DevolucaoPedidos não encontrada");
      return "Planilha não encontrada";
    }
    
    Logger.log("Última linha com dados: " + sheet.getLastRow());
    Logger.log("Última coluna com dados: " + sheet.getLastColumn());
    
    var dados = sheet.getDataRange().getValues();
    
    Logger.log("Cabeçalhos:");
    Logger.log(JSON.stringify(dados[0]));
    
    if (dados.length <= 1) {
      return "Nenhuma devolução encontrada na planilha";
    }
    
    Logger.log("Todas as devoluções:");
    var ativas = 0;
    var arquivadas = 0;
    
    for (var i = 1; i < dados.length; i++) {
      // Índice 13 = Arquivado (true/false)
      if (dados[i][13] === true) {
        arquivadas++;
      } else {
        ativas++;
        Logger.log("Devolução ativa " + i + ": " + JSON.stringify(dados[i]));
      }
    }
    
    return "OK - Devoluções encontradas: " + ativas + " ativas, " + arquivadas + " arquivadas";
  } catch (e) {
    Logger.log("Erro em testarDevolucoes: " + e.toString());
    return "ERRO: " + e.toString();
  }
}

function testarPedidosEmAtraso() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('PedidosEmAtraso');
    
    if (!sheet) {
      Logger.log("Planilha PedidosEmAtraso não encontrada");
      return "Planilha não encontrada";
    }
    
    Logger.log("Última linha com dados: " + sheet.getLastRow());
    Logger.log("Última coluna com dados: " + sheet.getLastColumn());
    
    var dados = sheet.getDataRange().getValues();
    
    Logger.log("Todos os dados:");
    for (var i = 0; i < dados.length; i++) {
      Logger.log("Linha " + (i+1) + ": " + JSON.stringify(dados[i]));
    }
    
    return "OK - Dados encontrados: " + (dados.length - 1);
  } catch (e) {
    Logger.log("Erro em testarPedidosEmAtraso: " + e.toString());
    return "ERRO: " + e.toString();
  }
}
