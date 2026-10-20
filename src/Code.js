const SHEET_NAME = 'Tarefas';

function doGet() {
  // Template para incluir o módulo PLH.html no fim do Index.html.
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('FlowTask - Sistema de Gerenciamento de Tarefas')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function testConnection() {
  return { success: true, message: 'Conexão OK!', timestamp: new Date().toString() };
}

function getTasks() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      createTaskSheet();
      return [];
    }
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return [];
    }
    var data = sheet.getRange(2, 1, lastRow - 1, 15).getValues();
    var tasks = [];
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      if (!row[1] || row[1].toString().trim() === '') {
        continue;
      }
      var ultimaAnotacao = row[4] || '';
      if (ultimaAnotacao instanceof Date) {
        ultimaAnotacao = Utilities.formatDate(ultimaAnotacao, ssTimeZone(), 'dd/MM/yyyy');
      } else if (ultimaAnotacao) {
        ultimaAnotacao = ultimaAnotacao.toString();
        if (ultimaAnotacao.indexOf(' ') > -1) {
          ultimaAnotacao = ultimaAnotacao.split(' ')[0];
        }
      }
      var task = {
        id: row[0] ? row[0].toString() : generateId(),
        descricao: row[1] ? row[1].toString() : '',
        rotulo: row[2] ? row[2].toString() : '',
        inicio: row[3] ? formatDate(row[3]) : '',
        ultimaAnotacao: ultimaAnotacao,
        check: row[5] ? formatDate(row[5]) : '',
        dias: row[6] ? Number(row[6]) : 0,
        prazo: row[7] ? formatDate(row[7]) : '',
        bucket: row[8] ? row[8].toString() : 'Pendente',
        percentual: row[9] ? Number(row[9]) : 0,
        prioridade: row[10] ? Number(row[10]) : 0,
        anotacao: row[11] ? row[11].toString() : '',
        responsavel: row[12] ? row[12].toString() : '',
        etapas: row[13] ? row[13].toString() : '',
        aguardandoDesde: row[14] ? formatDate(row[14]) : ''
      };
      tasks.push(task);
    }
    return tasks;
  } catch (error) {
    Logger.log('ERRO em getTasks: ' + error.toString());
    throw error;
  }
}

function saveTask(task) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = createTaskSheet();
    }

    // Auto-corrige o cabeçalho da coluna Responsável em planilhas antigas.
    if (sheet.getRange(1, 13).getValue() !== 'Responsável') {
      sheet.getRange(1, 13).setValue('Responsável');
      sheet.getRange(1, 13).setFontWeight('bold').setBackground('#2c3e50').setFontColor('#ffffff').setHorizontalAlignment('center');
    }

    // Auto-corrige o cabeçalho da coluna Etapas em planilhas antigas.
    if (sheet.getRange(1, 14).getValue() !== 'Etapas') {
      sheet.getRange(1, 14).setValue('Etapas');
      sheet.getRange(1, 14).setFontWeight('bold').setBackground('#2c3e50').setFontColor('#ffffff').setHorizontalAlignment('center');
    }

    // Auto-corrige o cabeçalho da coluna AguardandoDesde em planilhas antigas.
    if (sheet.getRange(1, 15).getValue() !== 'AguardandoDesde') {
      sheet.getRange(1, 15).setValue('AguardandoDesde');
      sheet.getRange(1, 15).setFontWeight('bold').setBackground('#2c3e50').setFontColor('#ffffff').setHorizontalAlignment('center');
    }

    // Regra: tarefa Concluída sempre 100%.
    if (task.bucket === 'Concluído') {
      task.percentual = 1;
    }

    var checkDate = '';
    if (task.check) {
      try {
        checkDate = parseLocalDate(task.check);
      } catch (e) {
        checkDate = '';
      }
    }

    var dias = 0;
    if (task.inicio && checkDate) {
      var inicio = parseLocalDate(task.inicio);
      dias = Math.floor((checkDate - inicio) / (1000 * 60 * 60 * 24));
    }

    var taskId = task.id || generateId();

    var ultimaAnotacaoFormatada = task.ultimaAnotacao || '';
    if (ultimaAnotacaoFormatada) {
      if (ultimaAnotacaoFormatada.indexOf(' ') > -1) {
        ultimaAnotacaoFormatada = ultimaAnotacaoFormatada.split(' ')[0];
      }
      if (!ultimaAnotacaoFormatada.match(/^\d{2}\/\d{2}\/\d{4}$/)) {
        try {
          var tempDate = new Date(ultimaAnotacaoFormatada);
          if (!isNaN(tempDate.getTime())) {
            ultimaAnotacaoFormatada = Utilities.formatDate(tempDate, ssTimeZone(), 'dd/MM/yyyy');
          }
        } catch (e) {
          Logger.log('Erro ao formatar data da última anotação: ' + e);
        }
      }
    }

    // Localiza a linha existente (se houver) para preservar o "aguardando desde".
    var rowIndex = 0, prevBucket = '', prevAguardando = '';
    if (task.id) {
      var lastRowLookup = sheet.getLastRow();
      if (lastRowLookup > 1) {
        var idCol = sheet.getRange(2, 1, lastRowLookup - 1, 1).getValues();
        for (var k = 0; k < idCol.length; k++) {
          if (idCol[k][0] === task.id) {
            rowIndex = k + 2;
            prevBucket = sheet.getRange(rowIndex, 9).getValue();
            prevAguardando = sheet.getRange(rowIndex, 15).getValue();
            break;
          }
        }
      }
    }

    // "Aguardando desde": marca hoje ao entrar em Aguardando; preserva enquanto continuar; limpa ao sair.
    var aguardandoDesde = '';
    if (task.bucket === 'Aguardando') {
      if (String(prevBucket) === 'Aguardando' && prevAguardando) {
        aguardandoDesde = prevAguardando;
      } else {
        var hj = new Date();
        aguardandoDesde = new Date(hj.getFullYear(), hj.getMonth(), hj.getDate(), 12, 0, 0);
      }
    }

    var rowData = [
      taskId,
      task.descricao || '',
      task.rotulo || '',
      task.inicio ? parseLocalDate(task.inicio) : '',
      ultimaAnotacaoFormatada,
      checkDate || '',
      dias,
      task.prazo ? parseLocalDate(task.prazo) : '',
      task.bucket || 'Pendente',
      Number(task.percentual) || 0,
      Number(task.prioridade) || 0,
      task.anotacao || '',
      task.responsavel || '',
      task.etapas || '',
      aguardandoDesde
    ];

    if (rowIndex) {
      sheet.getRange(rowIndex, 1, 1, 15).setValues([rowData]);
      formatRow(sheet, rowIndex);
      return { success: true, message: 'Tarefa atualizada!' };
    }

    sheet.appendRow(rowData);
    formatRow(sheet, sheet.getLastRow());
    return { success: true, message: 'Tarefa criada!' };
  } catch (error) {
    Logger.log('ERRO em saveTask: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}

/**
 * Atualiza apenas status e percentual de uma tarefa (usado pelo Kanban).
 * Mais leve que reenviar a tarefa inteira.
 */
function updateTaskStage(taskId, bucket, percentual) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      return { success: false, message: 'Planilha não encontrada!' };
    }
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return { success: false, message: 'Nenhuma tarefa!' };
    }
    var data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < data.length; i++) {
      if (data[i][0] === taskId) {
        var rowIndex = i + 2;
        var prevBucket = sheet.getRange(rowIndex, 9).getValue();
        var prevAguardando = sheet.getRange(rowIndex, 15).getValue();
        sheet.getRange(rowIndex, 9).setValue(bucket);              // Bucket
        sheet.getRange(rowIndex, 10).setValue(Number(percentual)); // %
        sheet.getRange(rowIndex, 10).setNumberFormat('0%');
        // "Aguardando desde": marca hoje ao entrar em Aguardando; preserva; limpa ao sair.
        if (bucket === 'Aguardando') {
          if (!(String(prevBucket) === 'Aguardando' && prevAguardando)) {
            var hj = new Date();
            sheet.getRange(rowIndex, 15).setValue(new Date(hj.getFullYear(), hj.getMonth(), hj.getDate(), 12, 0, 0)).setNumberFormat('dd/MM/yyyy');
          }
        } else {
          sheet.getRange(rowIndex, 15).setValue('');
        }
        return { success: true, message: 'Tarefa movida!' };
      }
    }
    return { success: false, message: 'Tarefa não encontrada!' };
  } catch (error) {
    Logger.log('ERRO em updateTaskStage: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}

function deleteTask(taskId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      return { success: false, message: 'Planilha não encontrada!' };
    }
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return { success: false, message: 'Nenhuma tarefa!' };
    }
    var data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < data.length; i++) {
      if (data[i][0] === taskId) {
        sheet.deleteRow(i + 2);
        return { success: true, message: 'Tarefa excluída!' };
      }
    }
    return { success: false, message: 'Tarefa não encontrada!' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function exportToXLSX() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      return { success: false, message: 'Planilha não encontrada!' };
    }
    var tempSheet = sheet.copyTo(ss);
    tempSheet.setName('Export_Temp_' + new Date().getTime());
    var lastRow = tempSheet.getLastRow();
    var lastCol = tempSheet.getLastColumn();
    if (lastRow > 1) {
      var dateColumns = [4, 6, 8];
      for (var i = 0; i < dateColumns.length; i++) {
        tempSheet.getRange(2, dateColumns[i], lastRow - 1, 1).setNumberFormat('dd/MM/yyyy');
      }
      tempSheet.getRange(2, 10, lastRow - 1, 1).setNumberFormat('0%');
      tempSheet.getRange(1, 1, lastRow, lastCol).setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
      for (var row = 2; row <= lastRow; row++) {
        if (row % 2 === 0) {
          tempSheet.getRange(row, 1, 1, lastCol).setBackground('#f8f9fa');
        }
      }
    }
    var url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?format=xlsx&gid=' + tempSheet.getSheetId();
    Utilities.sleep(2000);
    ss.deleteSheet(tempSheet);
    return { success: true, url: url, message: 'Arquivo XLSX gerado com sucesso!' };
  } catch (error) {
    Logger.log('ERRO em exportToXLSX: ' + error.toString());
    return { success: false, message: 'Erro ao exportar: ' + error.toString() };
  }
}

function createTaskSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var existing = ss.getSheetByName(SHEET_NAME);
  if (existing) { ss.deleteSheet(existing); }
  var sheet = ss.insertSheet(SHEET_NAME);
  var headers = ['ID', 'Descrição', 'Rótulo', 'Início', 'Última Anotação', 'Check', 'D', 'Prazo', 'Bucket', '%', 'P', 'Anotação', 'Responsável', 'Etapas', 'AguardandoDesde'];
  sheet.getRange(1, 1, 1, 15).setValues([headers]);
  var headerRange = sheet.getRange(1, 1, 1, 15);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#2c3e50');
  headerRange.setFontColor('#ffffff');
  headerRange.setHorizontalAlignment('center');
  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 350);
  sheet.setColumnWidth(3, 150);
  sheet.setColumnWidth(4, 100);
  sheet.setColumnWidth(5, 100);
  sheet.setColumnWidth(6, 100);
  sheet.setColumnWidth(7, 50);
  sheet.setColumnWidth(8, 100);
  sheet.setColumnWidth(9, 100);
  sheet.setColumnWidth(10, 60);
  sheet.setColumnWidth(11, 50);
  sheet.setColumnWidth(12, 300);
  sheet.setColumnWidth(13, 150);
  sheet.setColumnWidth(14, 250);
  sheet.setColumnWidth(15, 120);
  sheet.setFrozenRows(1);
  return sheet;
}

function formatRow(sheet, row) {
  sheet.getRange(row, 4).setNumberFormat('dd/MM/yyyy');
  sheet.getRange(row, 6).setNumberFormat('dd/MM/yyyy');
  sheet.getRange(row, 8).setNumberFormat('dd/MM/yyyy');
  sheet.getRange(row, 10).setNumberFormat('0%');
  sheet.getRange(row, 15).setNumberFormat('dd/MM/yyyy');
}

function generateId() {
  return 'ID_' + new Date().getTime() + '_' + Math.floor(Math.random() * 1000);
}

// Fuso da planilha (não do script). Evita deslocamento de 1 dia quando os dois diferem.
var _SS_TZ = null;
function ssTimeZone() {
  if (!_SS_TZ) { _SS_TZ = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(); }
  return _SS_TZ;
}

function formatDate(date) {
  try {
    if (!date) return '';
    if (typeof date === 'string') return date;
    return Utilities.formatDate(date, ssTimeZone(), 'yyyy-MM-dd');
  } catch (error) {
    return '';
  }
}

// Faz o parse da data por partes, ancorando ao MEIO-DIA para que nenhuma diferença de
// fuso (planilha x script) empurre a data para o dia anterior/seguinte ao gravar.
function parseLocalDate(dateStr) {
  if (!dateStr) return '';
  if (dateStr instanceof Date) return dateStr;
  var s = dateStr.toString().trim();
  // Formato YYYY-MM-DD
  if (s.match(/^\d{4}-\d{2}-\d{2}$/)) {
    var p = s.split('-');
    return new Date(parseInt(p[0]), parseInt(p[1]) - 1, parseInt(p[2]), 12, 0, 0);
  }
  // Formato DD/MM/YYYY
  if (s.match(/^\d{2}\/\d{2}\/\d{4}$/)) {
    var p = s.split('/');
    return new Date(parseInt(p[2]), parseInt(p[1]) - 1, parseInt(p[0]), 12, 0, 0);
  }
  return new Date(s);
}

/* ============================================================================
 *  MÓDULO PESSOAL / FAMÍLIA  (aba Tarefas_Pessoais)
 *  Colunas: ID · Aluno · Matéria · Descrição · Data · Prazo · Status · % ·
 *           Anotação · Origem · RefID
 * ========================================================================== */
const SHEET_NAME_PERSONAL = 'Tarefas_Pessoais';

function createPersonalSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var existing = ss.getSheetByName(SHEET_NAME_PERSONAL);
  if (existing) { ss.deleteSheet(existing); }
  var sheet = ss.insertSheet(SHEET_NAME_PERSONAL);
  var headers = ['ID', 'Aluno', 'Matéria', 'Descrição', 'Data', 'Prazo', 'Status', '%', 'Anotação', 'Origem', 'RefID', 'EventoID'];
  sheet.getRange(1, 1, 1, 12).setValues([headers]);
  var headerRange = sheet.getRange(1, 1, 1, 12);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#8e24aa');
  headerRange.setFontColor('#ffffff');
  headerRange.setHorizontalAlignment('center');
  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 120);
  sheet.setColumnWidth(3, 140);
  sheet.setColumnWidth(4, 340);
  sheet.setColumnWidth(5, 100);
  sheet.setColumnWidth(6, 100);
  sheet.setColumnWidth(7, 100);
  sheet.setColumnWidth(8, 60);
  sheet.setColumnWidth(9, 260);
  sheet.setColumnWidth(10, 100);
  sheet.setColumnWidth(11, 200);
  sheet.setColumnWidth(12, 220);
  sheet.setFrozenRows(1);
  return sheet;
}

function getPersonalSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME_PERSONAL);
  if (!sheet) { sheet = createPersonalSheet(); }
  return sheet;
}

function formatPersonalRow_(sheet, row) {
  sheet.getRange(row, 5).setNumberFormat('dd/MM/yyyy'); // Data
  sheet.getRange(row, 6).setNumberFormat('dd/MM/yyyy'); // Prazo
  sheet.getRange(row, 8).setNumberFormat('0%');         // %
}

function getPersonalTasks() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME_PERSONAL);
    if (!sheet) { createPersonalSheet(); return []; }
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) { return []; }
    var data = sheet.getRange(2, 1, lastRow - 1, 11).getValues();
    var tasks = [];
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      if (!row[3] || row[3].toString().trim() === '') { continue; }
      tasks.push({
        id: row[0] ? row[0].toString() : generateId(),
        aluno: row[1] ? row[1].toString() : '',
        materia: row[2] ? row[2].toString() : '',
        descricao: row[3] ? row[3].toString() : '',
        data: row[4] ? formatDate(row[4]) : '',
        prazo: row[5] ? formatDate(row[5]) : '',
        status: row[6] ? row[6].toString() : 'Pendente',
        percentual: row[7] ? Number(row[7]) : 0,
        anotacao: row[8] ? row[8].toString() : '',
        origem: row[9] ? row[9].toString() : 'Manual',
        refId: row[10] ? row[10].toString() : ''
      });
    }
    return tasks;
  } catch (error) {
    Logger.log('ERRO em getPersonalTasks: ' + error.toString());
    throw error;
  }
}

function savePersonalTask(task) {
  try {
    var sheet = getPersonalSheet_();
    if (task.status === 'Concluído') { task.percentual = 1; }
    var taskId = task.id || generateId();
    var rowData = [
      taskId,
      task.aluno || '',
      task.materia || '',
      task.descricao || '',
      task.data ? parseLocalDate(task.data) : '',
      task.prazo ? parseLocalDate(task.prazo) : '',
      task.status || 'Pendente',
      Number(task.percentual) || 0,
      task.anotacao || '',
      task.origem || 'Manual',
      task.refId || ''
    ];
    if (task.id) {
      var lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
        for (var i = 0; i < ids.length; i++) {
          if (ids[i][0] === task.id) {
            var rowIndex = i + 2;
            sheet.getRange(rowIndex, 1, 1, 11).setValues([rowData]);
            formatPersonalRow_(sheet, rowIndex);
            sincronizarEventoDaTarefa_(taskId); // reflete no Calendar na hora (best-effort)
            return { success: true, message: 'Tarefa pessoal atualizada!' };
          }
        }
      }
    }
    sheet.appendRow(rowData);
    formatPersonalRow_(sheet, sheet.getLastRow());
    sincronizarEventoDaTarefa_(taskId); // reflete no Calendar na hora (best-effort)
    return { success: true, message: 'Tarefa pessoal criada!' };
  } catch (error) {
    Logger.log('ERRO em savePersonalTask: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}

function deletePersonalTask(taskId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME_PERSONAL);
    if (!sheet) { return { success: false, message: 'Planilha não encontrada!' }; }
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) { return { success: false, message: 'Nenhuma tarefa!' }; }
    var rows = sheet.getRange(2, 1, lastRow - 1, 12).getValues();
    for (var i = 0; i < rows.length; i++) {
      if (rows[i][0] === taskId) {
        var eventoId = rows[i][11] ? rows[i][11].toString() : '';
        if (eventoId) { // apaga o evento do Calendar (best-effort)
          try {
            var ev = getCalendarioDeveres_().getEventById(eventoId);
            if (ev) ev.deleteEvent();
          } catch (e) { Logger.log('Falha ao apagar evento: ' + e); }
        }
        sheet.deleteRow(i + 2);
        return { success: true, message: 'Tarefa pessoal excluída!' };
      }
    }
    return { success: false, message: 'Tarefa não encontrada!' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

/**
 * Endpoint chamado pela rotina do Classroom. Recebe uma lista de deveres:
 *   [{ aluno, materia, descricao, prazo?, data?, refId? }]
 * Insere apenas os que ainda não existem (dedup por RefID), marcando Origem=Classroom.
 * Retorna { success, added, skipped }.
 */
function addHomeworkTasks(items) {
  try {
    if (!items || !items.length) { return { success: true, added: 0, skipped: 0, message: 'Nada a inserir.' }; }
    var sheet = getPersonalSheet_();

    // Conjunto de RefIDs já presentes (dedup).
    var existingRefs = {};
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      var refs = sheet.getRange(2, 11, lastRow - 1, 1).getValues();
      for (var i = 0; i < refs.length; i++) {
        var r = refs[i][0] ? refs[i][0].toString() : '';
        if (r) { existingRefs[r] = true; }
      }
    }

    var added = 0, skipped = 0;
    var rows = [];
    for (var j = 0; j < items.length; j++) {
      var it = items[j];
      var refId = it.refId ? it.refId.toString() : '';
      if (refId && existingRefs[refId]) { skipped++; continue; }
      if (refId) { existingRefs[refId] = true; } // evita duplicar dentro do mesmo lote
      rows.push([
        generateId(),
        it.aluno || '',
        it.materia || '',
        it.descricao || '',
        it.data ? parseLocalDate(it.data) : new Date(),
        it.prazo ? parseLocalDate(it.prazo) : '',
        'Pendente',
        0,
        '',
        'Classroom',
        refId
      ]);
      added++;
    }

    if (rows.length) {
      var startRow = sheet.getLastRow() + 1;
      sheet.getRange(startRow, 1, rows.length, 11).setValues(rows);
      for (var k = 0; k < rows.length; k++) { formatPersonalRow_(sheet, startRow + k); }
    }
    return { success: true, added: added, skipped: skipped, message: added + ' adicionada(s), ' + skipped + ' já existia(m).' };
  } catch (error) {
    Logger.log('ERRO em addHomeworkTasks: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}

/* ============================================================================
 *  SINCRONIZAÇÃO COM O GOOGLE CALENDAR (2 vias)
 *  Cada tarefa Pessoal vira um evento de dia inteiro no calendário "Deveres da
 *  Escola". Cor por aluno; VERDE = concluído (nas duas direções).
 *  Coluna 12 (EventoID) guarda o id do evento (dedup).
 *
 *  Rode uma vez: criarGatilhosCalendario()  (cria o calendário + 2 gatilhos 15h e 0h)
 * ========================================================================== */
var CALENDAR_NAME = 'Deveres da Escola';

function criarGatilhosCalendario() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'sincronizarCalendario') { ScriptApp.deleteTrigger(t); }
  });
  ScriptApp.newTrigger('sincronizarCalendario').timeBased().everyDays(1).atHour(15).create();
  ScriptApp.newTrigger('sincronizarCalendario').timeBased().everyDays(1).atHour(0).create();
  var cal = getCalendarioDeveres_();
  Logger.log('Calendário "' + CALENDAR_NAME + '" pronto (id: ' + cal.getId() + ').');
  Logger.log('Gatilhos criados: sincronizarCalendario às 15h e às 0h.');
  Logger.log('Agora compartilhe esse calendário com sua esposa em calendar.google.com.');
}

function getCalendarioDeveres_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('CALENDAR_ID');
  if (id) {
    var cal = CalendarApp.getCalendarById(id);
    if (cal) return cal;
  }
  // procura por nome (evita criar duplicado) ou cria um novo
  var existentes = CalendarApp.getCalendarsByName(CALENDAR_NAME);
  var novo = (existentes && existentes.length) ? existentes[0] : CalendarApp.createCalendar(CALENDAR_NAME);
  props.setProperty('CALENDAR_ID', novo.getId());
  return novo;
}

function corDoAluno_(aluno) {
  var a = (aluno || '').toString().toLowerCase();
  if (a.indexOf('gabriel') > -1) return CalendarApp.EventColor.BLUE;
  if (a.indexOf('heitor') > -1) return CalendarApp.EventColor.ORANGE;
  return CalendarApp.EventColor.GRAY;
}

// getColor() devolve o id da cor como texto. Aceita os dois verdes do Google
// Calendar como "feito": '2' = Sálvia (claro) e '10' = Manjericão (escuro).
var CORES_FEITO = { '2': true, '10': true };

function tituloDesejado_(aluno, materia, descricao) {
  var t = '[' + (aluno || '?') + '] ' + (materia ? materia + ' — ' : '') + (descricao || '');
  return t.length > 120 ? t.substring(0, 117) + '…' : t;
}

/**
 * Sincroniza UMA tarefa (linha) com o Calendar. Cria/atualiza/conclui conforme o caso.
 * Retorna { criado, concluido, verde, atualizado } (booleans).
 */
function sincronizarEvento_(cal, sheet, rowIndex, row) {
  var res = { criado: false, concluido: false, verde: false, atualizado: false };
  var descricao = row[3];
  if (!descricao || descricao.toString().trim() === '') return res;

  var aluno = row[1] ? row[1].toString() : '';
  var materia = row[2] ? row[2].toString() : '';
  var concluido = (row[6] ? row[6].toString() : 'Pendente') === 'Concluído';
  var eventoId = row[11] ? row[11].toString() : '';
  var quandoEvento = (row[5] instanceof Date) ? row[5] : (row[4] instanceof Date ? row[4] : null); // Prazo, senão Data
  var titulo = tituloDesejado_(aluno, materia, descricao.toString());

  // Sem evento ainda
  if (!eventoId) {
    if (!concluido && quandoEvento) {
      var ev = cal.createAllDayEvent(titulo, quandoEvento, {
        description: descricao.toString() + '\n\n(Para marcar como FEITO, mude a cor deste evento para VERDE.)'
      });
      try { ev.setColor(corDoAluno_(aluno)); } catch (e) {}
      sheet.getRange(rowIndex, 12).setValue(ev.getId());
      res.criado = true;
    }
    return res;
  }

  // Já existe evento
  var evento = null;
  try { evento = cal.getEventById(eventoId); } catch (e) { evento = null; }
  if (!evento) return res; // apagado no calendário: respeita, não recria

  var corAtual = '';
  try { corAtual = evento.getColor(); } catch (e) { corAtual = ''; }
  var estaVerde = !!CORES_FEITO[corAtual];

  // Cal -> planilha: evento verde e tarefa pendente
  if (estaVerde && !concluido) {
    sheet.getRange(rowIndex, 7).setValue('Concluído');
    sheet.getRange(rowIndex, 8).setValue(1).setNumberFormat('0%');
    res.concluido = true;
    return res;
  }

  // planilha -> Cal: tarefa concluída e evento não-verde
  if (concluido && !estaVerde) {
    try { evento.setColor(CalendarApp.EventColor.GREEN); res.verde = true; } catch (e) {}
    return res;
  }

  // Atualiza título/data se a tarefa mudou (pendente, não-verde)
  if (!concluido) {
    try {
      if (evento.getTitle() !== titulo) { evento.setTitle(titulo); res.atualizado = true; }
      if (quandoEvento && evento.getAllDayStartDate) {
        var atual = evento.getAllDayStartDate();
        if (atual) {
          var dAtual = Utilities.formatDate(atual, ssTimeZone(), 'yyyy-MM-dd');
          var dNovo = Utilities.formatDate(quandoEvento, ssTimeZone(), 'yyyy-MM-dd');
          if (dAtual !== dNovo) { evento.setAllDayDate(quandoEvento); res.atualizado = true; }
        }
      }
    } catch (e) {}
  }
  return res;
}

function sincronizarCalendario() {
  try {
    var sheet = getPersonalSheet_();
    // auto-corrige o cabeçalho da coluna EventoID
    if (sheet.getRange(1, 12).getValue() !== 'EventoID') {
      sheet.getRange(1, 12).setValue('EventoID');
      sheet.getRange(1, 12).setFontWeight('bold').setBackground('#8e24aa').setFontColor('#ffffff').setHorizontalAlignment('center');
    }
    var criados = 0, concluidos = 0, verdes = 0, atualizados = 0;
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      var cal = getCalendarioDeveres_();
      var data = sheet.getRange(2, 1, lastRow - 1, 12).getValues();
      for (var i = 0; i < data.length; i++) {
        var r = sincronizarEvento_(cal, sheet, i + 2, data[i]);
        if (r.criado) criados++;
        if (r.concluido) concluidos++;
        if (r.verde) verdes++;
        if (r.atualizado) atualizados++;
      }
    }
    var lastSync = Utilities.formatDate(new Date(), ssTimeZone(), 'dd/MM/yyyy HH:mm');
    PropertiesService.getScriptProperties().setProperty('LAST_SYNC', lastSync);
    Logger.log('=== SYNC === criados:' + criados + ' concluidos:' + concluidos +
      ' verdes:' + verdes + ' atualizados:' + atualizados);
    return { success: true, criados: criados, concluidos: concluidos, verdes: verdes, atualizados: atualizados, lastSync: lastSync };
  } catch (error) {
    Logger.log('ERRO em sincronizarCalendario: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}

function getSyncStatus() {
  return { lastSync: PropertiesService.getScriptProperties().getProperty('LAST_SYNC') || '' };
}

/**
 * Status da automação para o app: última leitura do Classroom por filho (aba
 * "Execucoes", preenchida pelos scripts dos filhos) + última sincronização do Calendar.
 */
function getExecStatus() {
  var out = { execucoes: [], lastSync: '' };
  try {
    out.lastSync = PropertiesService.getScriptProperties().getProperty('LAST_SYNC') || '';
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var aba = ss.getSheetByName('Execucoes');
    if (aba && aba.getLastRow() > 1) {
      var data = aba.getRange(2, 1, aba.getLastRow() - 1, 3).getValues();
      for (var i = 0; i < data.length; i++) {
        if (data[i][0]) {
          out.execucoes.push({
            aluno: data[i][0].toString(),
            ultima: data[i][1] ? data[i][1].toString() : '',
            novas: data[i][2] === '' || data[i][2] == null ? '' : data[i][2]
          });
        }
      }
    }
  } catch (e) {
    Logger.log('getExecStatus: ' + e);
  }
  return out;
}

/** Sincroniza uma única tarefa pelo ID (best-effort) — usado ao salvar no app. */
function sincronizarEventoDaTarefa_(taskId) {
  try {
    var sheet = getPersonalSheet_();
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return;
    var data = sheet.getRange(2, 1, lastRow - 1, 12).getValues();
    for (var i = 0; i < data.length; i++) {
      if (data[i][0] === taskId) {
        var cal = getCalendarioDeveres_();
        sincronizarEvento_(cal, sheet, i + 2, data[i]);
        return;
      }
    }
  } catch (e) {
    Logger.log('sincronizarEventoDaTarefa_ (best-effort) falhou: ' + e);
  }
}

/* =====================================================================
 * WhatsApp (CallMeBot) — centralizado aqui no FlowTask (conta do dono).
 * Lê a aba Tarefas_Pessoais e avisa pai/mãe; NÃO depende das contas dos filhos.
 *
 * CONFIGURAÇÃO (Configurações do projeto → Propriedades do script):
 *   WHATS_NOTIFICAR = [{"nome":"Pai","phone":"5584XXXXXXXXX","apikey":"111111"},
 *                      {"nome":"Mãe","phone":"5584YYYYYYYYY","apikey":"222222"}]
 *   (a apikey é a que cada pessoa recebeu do CallMeBot no opt-in; phone com DDI 55)
 *
 * USO:
 *   testarWhats()        → envia mensagem de teste (loga o HTTP por destinatário).
 *   testarEntregas()     → roda o alerta "entregas de amanhã" na hora.
 *   testarNovas()        → roda a checagem de tarefas novas na hora.
 *   criarGatilhosWhats() → agenda entregas (17h/dia) + novas tarefas (de hora em hora).
 * ===================================================================== */

function getNotificar_() {
  try {
    var raw = PropertiesService.getScriptProperties().getProperty('WHATS_NOTIFICAR');
    if (!raw) return [];
    var arr = JSON.parse(raw);
    return Array.isArray(arr) ? arr : [];
  } catch (e) {
    Logger.log('WHATS_NOTIFICAR inválido (precisa ser JSON): ' + e);
    return [];
  }
}

/** Envia a mensagem para cada destino do WhatsApp via CallMeBot (best-effort, com log). */
function enviarWhatsApp_(msg) {
  try {
    var lista = getNotificar_();
    if (!lista.length) {
      Logger.log('WhatsApp: WHATS_NOTIFICAR vazio/ausente — configure a Propriedade do script.');
      return;
    }
    for (var i = 0; i < lista.length; i++) {
      var d = lista[i];
      var phone = d && d.phone ? String(d.phone) : '';
      var apikey = d && d.apikey ? String(d.apikey) : '';
      if (!phone || !apikey) {
        Logger.log('WhatsApp: pulando "' + ((d && d.nome) || '?') + '" — phone/apikey em branco.');
        continue;
      }
      var url = 'https://api.callmebot.com/whatsapp.php?phone=' + encodeURIComponent(phone) +
        '&apikey=' + encodeURIComponent(apikey) + '&text=' + encodeURIComponent(msg);
      var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      Logger.log('WhatsApp → ' + ((d && d.nome) || phone) + ': HTTP ' + resp.getResponseCode() +
        ' | ' + resp.getContentText().substring(0, 160));
      Utilities.sleep(1200);
    }
  } catch (e) {
    Logger.log('enviarWhatsApp_ falhou: ' + e);
  }
}

/**
 * Rode UMA VEZ para autorizar o acesso à internet (UrlFetch). Como NÃO tem try/catch,
 * o Google vai mostrar a tela de "Autorização necessária". Depois de permitir, o WhatsApp funciona.
 */
function autorizar() {
  var code = UrlFetchApp.fetch('https://www.google.com').getResponseCode();
  Logger.log('Permissão de internet OK (HTTP ' + code + '). Agora rode testarWhats.');
}

function testarWhats() {
  enviarWhatsApp_('✅ Teste FlowTask: se você recebeu isto, o WhatsApp está funcionando.');
}
function testarEntregas() { alertarEntregasAmanha(); }
function testarNovas() { alertarNovasTarefas(); }

/** Envia uma mensagem de EXEMPLO (conteúdo fictício) no mesmo formato das 17h, só para você ver como chega. */
function exemploEntregas() {
  var tz = ssTimeZone();
  var amanha = new Date(); amanha.setDate(amanha.getDate() + 1);
  var dataBR = Utilities.formatDate(amanha, tz, 'dd/MM');
  var msg = '🎒 Para entregar amanhã (' + dataBR + '):\n\n' +
    '👦 Gabriel:\n• Matemática: exercícios da página 42\n• Português: levar o livro de leitura\n\n' +
    '👦 Heitor:\n• Ciências: trazer o material da experiência\n\n' +
    '(mensagem de exemplo — a real é gerada automaticamente às 17h com os deveres da planilha)';
  enviarWhatsApp_(msg);
  Logger.log('Exemplo da mensagem das 17h enviado.');
}

function criarGatilhosWhats() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    var h = t.getHandlerFunction();
    if (h === 'alertarEntregasAmanha' || h === 'alertarNovasTarefas') { ScriptApp.deleteTrigger(t); }
  });
  ScriptApp.newTrigger('alertarEntregasAmanha').timeBased().everyDays(1).atHour(17).create();
  ScriptApp.newTrigger('alertarNovasTarefas').timeBased().everyHours(1).create();
  Logger.log('Gatilhos do WhatsApp criados: entregas 17h/dia + novas tarefas de hora em hora.');
}

/** Avisa no WhatsApp o que cada aluno precisa entregar AMANHÃ (uma mensagem só). */
function alertarEntregasAmanha() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_PERSONAL);
    if (!sheet) { Logger.log('Aba ' + SHEET_NAME_PERSONAL + ' não encontrada.'); return; }
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) { Logger.log('Sem tarefas pessoais.'); return; }

    var tz = ssTimeZone();
    var amanha = new Date(); amanha.setDate(amanha.getDate() + 1);
    var chave = Utilities.formatDate(amanha, tz, 'yyyy-MM-dd');

    var data = sheet.getRange(2, 1, lastRow - 1, 7).getValues(); // ID,Aluno,Matéria,Descrição,Data,Prazo,Status
    var porAluno = {};
    for (var i = 0; i < data.length; i++) {
      var r = data[i];
      var aluno = (r[1] || '').toString() || 'Aluno';
      var materia = (r[2] || '').toString();
      var descricao = (r[3] || '').toString();
      var prazo = r[5];
      var status = (r[6] || '').toString();
      if (status === 'Concluído') continue;
      if (!(prazo instanceof Date)) continue;
      if (Utilities.formatDate(prazo, tz, 'yyyy-MM-dd') !== chave) continue;
      (porAluno[aluno] = porAluno[aluno] || []).push('• ' + (materia ? materia + ': ' : '') + descricao);
    }

    var alunos = Object.keys(porAluno);
    if (!alunos.length) { Logger.log('Sem entregas para amanhã (' + chave + ').'); return; }
    var dataBR = Utilities.formatDate(amanha, tz, 'dd/MM');
    var partes = alunos.map(function (a) { return '👦 ' + a + ':\n' + porAluno[a].join('\n'); });
    enviarWhatsApp_('🎒 Para entregar amanhã (' + dataBR + '):\n\n' + partes.join('\n\n'));
    Logger.log('Alerta de entregas de amanhã enviado.');
  } catch (e) {
    Logger.log('alertarEntregasAmanha falhou: ' + e);
  }
}

/**
 * Detecta tarefas novas (origem Classroom) na planilha desde a última checagem e avisa.
 * Guarda os IDs já avisados em WHATS_NOTIFICADOS. Na 1ª execução só cria a linha de base
 * (não envia nada), para não disparar um monte de tarefas antigas.
 */
function alertarNovasTarefas() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_PERSONAL);
    if (!sheet) return;
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    var props = PropertiesService.getScriptProperties();
    var raw = props.getProperty('WHATS_NOTIFICADOS');
    var primeiraVez = !raw;
    var notificados = {};
    if (raw) { try { notificados = JSON.parse(raw) || {}; } catch (e) { notificados = {}; } }

    var data = sheet.getRange(2, 1, lastRow - 1, 10).getValues(); // até Origem (col 10)
    var novos = [];
    for (var i = 0; i < data.length; i++) {
      var r = data[i];
      var id = (r[0] || '').toString();
      var aluno = (r[1] || '').toString() || 'Aluno';
      var materia = (r[2] || '').toString();
      var descricao = (r[3] || '').toString();
      var origem = (r[9] || '').toString();
      if (!id || origem !== 'Classroom') continue;
      if (notificados[id]) continue;
      notificados[id] = true;
      if (!primeiraVez) novos.push({ aluno: aluno, linha: '• ' + (materia ? materia + ': ' : '') + descricao });
    }
    props.setProperty('WHATS_NOTIFICADOS', JSON.stringify(notificados));

    if (primeiraVez) {
      Logger.log('Baseline criado (' + Object.keys(notificados).length + ' tarefas) — nada enviado.');
      return;
    }
    if (!novos.length) { Logger.log('Nenhuma tarefa nova.'); return; }

    var porAluno = {};
    novos.forEach(function (n) { (porAluno[n.aluno] = porAluno[n.aluno] || []).push(n.linha); });
    var tz = ssTimeZone();
    var partes = Object.keys(porAluno).map(function (a) { return '📚 ' + a + ':\n' + porAluno[a].join('\n'); });
    enviarWhatsApp_('📚 Novas tarefas (' + Utilities.formatDate(new Date(), tz, 'dd/MM') + '):\n\n' + partes.join('\n\n'));
    Logger.log('Aviso de novas tarefas enviado (' + novos.length + ').');
  } catch (e) {
    Logger.log('alertarNovasTarefas falhou: ' + e);
  }
}
