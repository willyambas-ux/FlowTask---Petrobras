/* ============================================================================
 *  MÓDULO PLH SSUB — painel de integridade submarina
 *
 *  Fontes (o usuário cola os exports em A1, com cabeçalho):
 *    SRC_RT             GIDES  · TratamentoRT
 *    SRC_Planejamento   GIDES  · ConsultaInspecoes
 *    SRC_Historico      ISUB   · GridConsultaExecucoesExport
 *    SRC_Interligacoes  CIDR   · baixados
 *    Inspecoes_Jaqueta  controle próprio (colunas lidas pelo nome do cabeçalho)
 *
 *  Nada é gravado: getPlhData() lê as abas e calcula tudo a cada abertura,
 *  então os prazos ("vencida", "vence em 30 dias"...) estão sempre em dia.
 *  onEdit() só registra a data em que cada aba SRC_ foi colada.
 * ==========================================================================*/

var PLH_JANELA_PLANEJ = 120; // dias: "Planejar agora"

var PLH_FONTES = [
  { chave: 'RT',            aba: 'SRC_RT',            sistema: 'GIDES', exige: ['rt', 'delineamento'] },
  { chave: 'Planejamento',  aba: 'SRC_Planejamento',  sistema: 'GIDES', exige: ['vale para', 'periodicidade'] },
  { chave: 'Historico',     aba: 'SRC_Historico',     sistema: 'ISUB',  exige: ['patec aplic.', 'prazo emissao patec'] },
  { chave: 'Interligacoes', aba: 'SRC_Interligacoes', sistema: 'CIDR',  exige: ['inst. origem', 'inst. destino'] }
];
var PLH_ABA_JAQUETAS = 'SRC_Jaquetas'; // controle próprio, mantido direto na planilha
var PLH_ABA_JAQUETAS_ANTIGA = 'Jaquetas'; // nome anterior, aceito por compatibilidade

/* Colunas por posição (o GIDES repete nomes como "BDIEP" e "Situação"). */
var PLH_COLS = {
  RT: ['RT','Nota_SAP','Ativo','Instalacao_Resp','SPIE','Contribui_Indicador','Ambiente','ANP',
       'Objeto','Objeto_BDIEP','Objeto_Resp','Tipo_Componente','Componente','Componente_BDIEP','Tecnica',
       'Data_Inspecao','Data_Aprovacao','Data_Venc_Original','Data_Vencimento','Prioridade',
       'Dt_Alt_Prioridade','Resp_Alt_Prioridade','Qtd_Alt_Prioridade','Situacao_Cod','Local',
       'Descricao','Delineamento','Executante','Status_SAP','RT_Migrada','Situacao_Duto','Situacao_RT',
       'Data_Execucao','Data_Cancelamento','Observacao','Justificativa','Tipo_Servico','Tipo_Recurso',
       'Estimativa_Recurso','Info_Complementares','Necessidade_Material','Previsao_Execucao',
       'Prev_Contratacao','Real_Contratacao','Prev_Nivelamento','Real_Nivelamento','Local_SAP','Prazo',
       'Dt_Alt_Prazo','Resp_Alt_Prazo','Qtd_Alt_Prazo','Link_PATEC','Link_Relatorio'],
  Planejamento: ['Ativo','Instalacao_Resp','Objeto','Objeto_BDIEP','Configuracao','Tipo_Objeto',
       'Situacao_Duto','SPIE','Ambiente','ANP','Tipo_Componente','Nome_Componente','Componente_BDIEP',
       'Tramo_Suporte','Local_Instalacao','Extensao_m','Dt_Instalacao','Idade_anos','Planejamento',
       'Procedimento','Dt_Ultima_Inspecao','Periodicidade','Vale_Para','Dt_Proxima_Inspecao',
       'Dt_Validade_ACO','Dias_Conclusao_Planej','Dt_Aprovacao_PATEC','Class_Corrosividade',
       'PATEC_Aplicavel','Relatorio','Dt_Cadastro_Relatorio','Tipos_Recursos','TEADF_PIADF',
       'Analise_Moda','Link_Relatorio','Nome_PATEC','Link_PATEC','Nome_ACO','Link_ACO','Contribui_ICPI',
       'Tem_Descont_Critica','Qual_Descont_Critica','Local_Descont_Critica','Agente_Descont_Critica',
       'RTA_Descont_Critica'],
  Historico: ['Ativo','Instalacao_Resp','Interligacao','Componente_BDIEP','Componente','Procedimento',
       'Local_Instalacao','Ambiente','PATEC_Aplic','Dt_Execucao','Prazo_Emissao_PATEC',
       'Dt_Cadastro_Relatorio','Resp_Relatorio','Relatorios','Links_Relatorios','PATEC','Link_PATEC',
       'Status_PATEC_Sistema','Dt_Inicio_Elab','Elaborador','Dt_Envio_Verif','Verificador',
       'Dt_Envio_Aprov','Aprovador','Dt_Conclusao','Resultado_TEADF_PIADF','Analise_MODA',
       'Class_Corrosao','Cod_Planej_Posterior','Status_Planej_Posterior','Eh_Ultimo'],
  Interligacoes: ['Nome','Objeto_BDIEP','Funcao','Instalacao_Resp','Inst_Origem','Inst_Destino',
       'Classe','Fluido','Configuracao','Escopo_ANP','SPIE','Situacao']
};

/* ---------- registro da data de colagem ---------- */

function onEdit(e) {
  try {
    var nome = e && e.range ? e.range.getSheet().getName() : '';
    if (nome.indexOf('SRC_') === 0 || nome === PLH_ABA_JAQUETAS) {
      PropertiesService.getDocumentProperties()
        .setProperty('PLH_ATUALIZADO_' + nome, new Date().toISOString());
    }
  } catch (err) { /* simples gatilho: nunca interromper a edição */ }
}

/* ---------- utilitários ---------- */

function plhTxt_(v) {
  if (v === null || v === undefined) return '';
  if (v instanceof Date) return Utilities.formatDate(v, ssTimeZone(), 'dd/MM/yyyy');
  return String(v).replace(/ /g, ' ').replace(/­/g, '').trim();
}

function plhNorm_(v) {
  return plhTxt_(v).toLowerCase()
    .normalize('NFD').replace(/[̀-ͯ]/g, '')
    .replace(/\s+/g, ' ');
}

function plhHoje_() {
  var h = new Date();
  return new Date(h.getFullYear(), h.getMonth(), h.getDate());
}

/** Date | "dd/mm/aaaa" | "d/m/aa" (primeira linha da célula) → Date à meia-noite, ou null. */
function plhData_(v) {
  if (v === null || v === undefined || v === '') return null;
  if (v instanceof Date) return isNaN(v) ? null : new Date(v.getFullYear(), v.getMonth(), v.getDate());
  var s = String(v).split(/[\r\n]/)[0].trim();
  var m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
  if (!m) return null;
  var dia = Number(m[1]), mes = Number(m[2]), a = Number(m[3]);
  if (!dia || !mes || mes > 12 || dia > 31) return null; // ex.: 0/0/0, 00/01/00
  if (a < 100) a += 2000;
  var d = new Date(a, mes - 1, dia);
  return isNaN(d) ? null : d;
}

function plhIso_(d) {
  return d ? Utilities.formatDate(d, ssTimeZone(), 'yyyy-MM-dd') : '';
}

function plhIsoBR_(iso) {
  if (!iso) return '-';
  var p = iso.split('-');
  return p[2] + '/' + p[1] + '/' + p[0];
}

function plhDias_(d, hoje) {
  return d ? Math.round((d - hoje) / 86400000) : null;
}

function plhNum_(v) {
  if (typeof v === 'number') return v;
  var s = plhTxt_(v).replace(/\./g, '').replace(',', '.');
  var n = parseFloat(s);
  return isNaN(n) ? null : n;
}

/** PXA-3 → PXA03 · PBIQ1 → PBIQ01 · PPE1A mantido. */
function plhPlataforma_(x) {
  var p = plhTxt_(x).toUpperCase().replace(/[-\s]/g, '');
  if (!p) return '';
  var m = p.match(/^([A-Z]*)(\d+)$/);
  if (m) return m[1] + (m[2].length < 2 ? '0' + m[2] : m[2]);
  return p;
}

function plhCorta_(s, n) {
  s = plhTxt_(s);
  return s.length > n ? s.substring(0, n - 1) + '…' : s;
}

/* ---------- leitura das abas ---------- */

function plhLerFonte_(ss, f) {
  var props = PropertiesService.getDocumentProperties();
  var info = { base: f.chave, sistema: f.sistema, aba: f.aba, linhas: 0, ok: false, erro: '',
               atualizado: props.getProperty('PLH_ATUALIZADO_' + f.aba) || '' };
  var sh = ss.getSheetByName(f.aba);
  if (!sh) { info.erro = 'Aba não encontrada'; return { info: info, rows: [] }; }
  info.gid = sh.getSheetId();
  var vals = sh.getDataRange().getValues();
  if (vals.length < 2) { info.erro = 'Aba vazia: cole o export em A1'; return { info: info, rows: [] }; }
  var cab = vals[0].map(plhNorm_);
  var faltam = f.exige.filter(function (c) { return cab.indexOf(c) < 0; });
  if (faltam.length) {
    info.erro = 'Cabeçalho não confere (faltam: ' + faltam.join(', ') + '). Export colado na aba errada?';
    return { info: info, rows: [] };
  }
  var nomes = PLH_COLS[f.chave];
  var rows = [];
  for (var i = 1; i < vals.length; i++) {
    var r = vals[i], o = {};
    for (var c = 0; c < nomes.length; c++) o[nomes[c]] = r[c];
    rows.push(o);
  }
  info.ok = true;
  return { info: info, rows: rows };
}

/** Aba Jaquetas: lida pelo nome do cabeçalho (é controle próprio e pode ganhar colunas).
 *  Uma linha por plataforma × zona (Emersa/Submersa) × nível (N1/N2). */
function plhLerJaquetas_(ss) {
  var props = PropertiesService.getDocumentProperties();
  var info = { base: 'Jaquetas', sistema: 'Controle próprio', aba: PLH_ABA_JAQUETAS, linhas: 0,
               ok: false, erro: '', atualizado: props.getProperty('PLH_ATUALIZADO_' + PLH_ABA_JAQUETAS) || '' };
  var sh = ss.getSheetByName(PLH_ABA_JAQUETAS) || ss.getSheetByName(PLH_ABA_JAQUETAS_ANTIGA);
  if (!sh) { info.erro = 'Aba não encontrada'; return { info: info, rows: [] }; }
  info.aba = sh.getName(); info.gid = sh.getSheetId();
  var vals = sh.getDataRange().getValues();
  if (vals.length < 2) { info.erro = 'Aba vazia'; return { info: info, rows: [] }; }
  var cab = vals[0].map(plhNorm_);
  var col = function () {
    for (var i = 0; i < arguments.length; i++) { var k = cab.indexOf(arguments[i]); if (k >= 0) return k; }
    return -1;
  };
  var ix = {
    uf: col('uf'), plataforma: col('plataforma'), acesso: col('acesso'),
    zona: col('zona', 'inspecao'), nivel: col('nivel'), plano: col('plano'),
    statusPlano: col('status do plano', 'status'),
    perDias: col('periodicidade (dias)', 'periodicidade [dias]', 'periodicidade dias'),
    ordem: col('ordem'), ultima: col('ultima inspecao'), lote: col('lote de controle', 'lote inspecao'),
    du: col('du/rl aplicado', 'aplicar du ou rl'), proxima: col('proxima inspecao'),
    proxOS: col('proxima os', 'proxima os.'), juntas: col('juntas criticas'),
    obs: col('observacao'), acoes: col('acoes')
  };
  if (ix.plataforma < 0 || ix.ultima < 0 || ix.nivel < 0) {
    info.erro = 'Cabeçalho não confere (Plataforma, Nível e Última inspeção são obrigatórios)';
    return { info: info, rows: [] };
  }
  // "-", "—" e "0/0/0" na planilha de controle significam "vazio"
  var g = function (r, k) {
    if (ix[k] < 0) return '';
    var v = r[ix[k]];
    if (typeof v === 'string' && /^\s*([-—–]+|0\/0\/0+)\s*$/.test(v)) return '';
    return v;
  };
  var rows = [], ignoradas = 0;
  for (var i = 1; i < vals.length; i++) {
    var r = vals[i];
    if (!plhTxt_(g(r, 'plataforma'))) continue;
    // Sem plano e sem nenhuma data: plataforma sem inspeção de jaqueta (ex.: PAG-2, de concreto)
    if (!plhTxt_(g(r, 'plano')) && !plhData_(g(r, 'ultima')) && !plhData_(g(r, 'proxima'))) { ignoradas++; continue; }
    var zona = plhTxt_(g(r, 'zona')), niv = plhTxt_(g(r, 'nivel'));
    rows.push({
      UF: plhTxt_(g(r, 'uf')), Plataforma: plhTxt_(g(r, 'plataforma')), Acesso: plhTxt_(g(r, 'acesso')),
      Zona: zona, Niv: niv, Nivel: (niv + ' ' + zona).trim(),
      Plano: plhTxt_(g(r, 'plano')), Status_Plano: plhTxt_(g(r, 'statusPlano')),
      Periodicidade_Dias: plhNum_(g(r, 'perDias')), Ordem: plhTxt_(g(r, 'ordem')),
      Dt_Ultima: plhData_(g(r, 'ultima')), Lote: plhTxt_(g(r, 'lote')), DU: plhTxt_(g(r, 'du')),
      Dt_Proxima: plhData_(g(r, 'proxima')), Prox_OS: plhTxt_(g(r, 'proxOS')),
      Juntas_Criticas: plhTxt_(g(r, 'juntas')), Observacao: plhTxt_(g(r, 'obs')), Acoes: plhTxt_(g(r, 'acoes'))
    });
  }
  info.ok = true;
  info.ignoradas = ignoradas;
  return { info: info, rows: rows };
}

/** SRC_ICPI (painel BI, opcional): só a OS e o Status OS de cada demanda. Lido pelo nome do cabeçalho. */
var PLH_ABA_ICPI = 'SRC_ICPI';
function plhLerICPI_(ss) {
  var props = PropertiesService.getDocumentProperties();
  var info = { base: 'ICPI (OS)', sistema: 'Painel BI', aba: PLH_ABA_ICPI, linhas: 0, ok: false, erro: '',
               opcional: true, atualizado: props.getProperty('PLH_ATUALIZADO_' + PLH_ABA_ICPI) || '' };
  var sh = ss.getSheetByName(PLH_ABA_ICPI);
  if (!sh) { info.erro = 'Aba não encontrada (opcional)'; return { info: info, rows: [] }; }
  info.gid = sh.getSheetId();
  var vals = sh.getDataRange().getValues();
  if (vals.length < 2) { info.erro = 'Aba vazia (opcional)'; return { info: info, rows: [] }; }
  var cab = vals[0].map(plhNorm_);
  var col = function (n) { return cab.indexOf(n); };
  var ix = { objeto: col('objeto submarino'), proc: col('procedimento'), uep: col('uep'), venc: col('venc. min.'),
             os: col('os'), statusOS: col('status os'), grupo: col('grupo') };
  if (ix.objeto < 0 || ix.proc < 0 || ix.venc < 0) { info.erro = 'Cabeçalho não confere (Objeto Submarino, Procedimento, Venc. Mín.)'; return { info: info, rows: [] }; }
  var g = function (r, k) { return ix[k] < 0 ? '' : r[ix[k]]; };
  var rows = [], cerquilhas = 0;
  for (var i = 1; i < vals.length; i++) {
    var r = vals[i];
    if (!plhTxt_(g(r, 'objeto'))) continue;
    if (/^#+$/.test(plhTxt_(g(r, 'venc'))) || /^#+$/.test(plhTxt_(g(r, 'os')))) cerquilhas++;
    var os = plhTxt_(g(r, 'os'));
    if (/^#+$/.test(os)) os = '';
    rows.push({ objeto: plhTxt_(g(r, 'objeto')), proc: plhTxt_(g(r, 'proc')), uep: plhTxt_(g(r, 'uep')),
                venc: plhData_(g(r, 'venc')), os: (os === '-' ? '' : os), statusOS: plhTxt_(g(r, 'statusOS')),
                grupo: plhTxt_(g(r, 'grupo')) });
  }
  info.ok = true;
  info.linhas = rows.length;
  if (cerquilhas) info.aviso = cerquilhas + ' linha(s) com ##### em Venc. Mín./OS: no Excel, alargue as colunas (Ctrl+A e duplo clique na divisa entre colunas) antes de copiar';
  return { info: info, rows: rows };
}

/* ---------- regras de negócio ---------- */

function plhStatusPrazo_(dias, janela) {
  if (dias === null) return 'Sem data';
  if (dias < 0) return 'Vencida';
  if (dias <= janela) return 'Planejar agora';
  if (dias <= 365) return 'Próximos 12 meses';
  return 'No prazo';
}

function plhCategoriaRT_(descr) {
  var t = plhTxt_(descr).toLowerCase();
  var tem = function (ks) { return ks.some(function (k) { return t.indexOf(k) >= 0; }); };
  if (tem(['vão', 'calçamento'])) return 'Vão livre';
  if (tem(['contato entre'])) return 'Contato entre dutos';
  if (tem(['anodo', 'potencia', 'catódica', 'eletroquímico', 'contato elétrico'])) return 'Proteção catódica';
  if (tem(['revestimento', 'pintura', '(tmp)'])) return 'Revestimento / pintura';
  if (tem(['braçadeira', 'estojo'])) return 'Fixação / braçadeiras';
  if (tem(['cercamento', 'capina', 'sinaliza', 'marco'])) return 'Faixa / sinalização';
  if (tem(['inspe'])) return 'Inspeção complementar';
  return 'Outros';
}

function plhUrgencia_(situacao, dias) {
  var s = situacao || '';
  if (s === 'Verificar data (futura)' || s === 'Sem plano cadastrado') return '5 - Verificar dado';
  if (s.indexOf('Vencida') === 0 || s === 'PATEC atrasado' || (dias !== null && dias < 0)) return '1 - Vencido';
  if (dias !== null && dias <= 30) return '2 - Até 30 dias';
  if (dias !== null && dias <= 90) return '3 - Até 90 dias';
  return '4 - Acima de 90 dias';
}

/* ---------- função chamada pelo painel ---------- */

function getPlhData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hoje = plhHoje_();
  var fontes = [], src = {};
  PLH_FONTES.forEach(function (f) {
    var r = plhLerFonte_(ss, f);
    src[f.chave] = r.rows; fontes.push(r.info);
  });
  var jq = plhLerJaquetas_(ss);
  fontes.push(jq.info);
  var icpi = plhLerICPI_(ss);

  /* --- Planejamento --- */
  var plan = [];
  src.Planejamento.forEach(function (r) {
    if (!plhTxt_(r.Objeto)) return;
    var prox = plhData_(r.Dt_Proxima_Inspecao), ult = plhData_(r.Dt_Ultima_Inspecao);
    var dias = plhDias_(prox, hoje);
    plan.push({
      plat: plhPlataforma_(r.Instalacao_Resp), objId: plhTxt_(r.Objeto_BDIEP), objeto: plhTxt_(r.Objeto),
      comp: plhTxt_(r.Nome_Componente), compId: plhTxt_(r.Componente_BDIEP), proc: plhTxt_(r.Procedimento),
      per: plhTxt_(r.Periodicidade), planej: plhTxt_(r.Planejamento), situacao: plhTxt_(r.Situacao_Duto),
      ult: plhIso_(ult), prox: plhIso_(prox), dias: dias, status: plhStatusPrazo_(dias, PLH_JANELA_PLANEJ),
      semAnterior: !ult, descCritica: plhNorm_(r.Tem_Descont_Critica) === 'sim'
    });
  });

  /* --- RT (só as em aberto vão para o painel) --- */
  var rts = [], rtTotal = 0;
  src.RT.forEach(function (r) {
    var num = plhTxt_(r.RT);
    if (!num) return;
    rtTotal++;
    var cod = plhTxt_(r.Situacao_Cod).toUpperCase();
    if (cod === 'E' || cod === 'C') return;
    var venc = plhData_(r.Data_Vencimento), orig = plhData_(r.Data_Venc_Original);
    var dias = plhDias_(venc, hoje);
    var st = !venc ? 'Sem data' : dias < 0 ? 'Vencida' : dias <= 30 ? 'Vence em 30 dias'
           : dias <= 90 ? 'Vence em 90 dias' : 'No prazo';
    rts.push({
      rt: num, nota: plhTxt_(r.Nota_SAP), plat: plhPlataforma_(r.Instalacao_Resp),
      objId: plhTxt_(r.Objeto_BDIEP), objeto: plhTxt_(r.Objeto),
      comp: plhTxt_(r.Tipo_Componente) + ' [' + plhTxt_(r.Componente) + ']',
      prio: plhTxt_(r.Prioridade), cat: plhCategoriaRT_(r.Descricao), descr: plhCorta_(r.Descricao, 400),
      venc: plhIso_(venc), dias: dias, status: st,
      reaprazada: !!(venc && orig && venc.getTime() !== orig.getTime())
    });
  });

  /* --- Histórico / PATEC --- */
  var patec = [];
  var limiteAntigo = new Date(hoje.getTime() - 365 * 86400000);
  src.Historico.forEach(function (r) {
    var compId = plhTxt_(r.Componente_BDIEP);
    if (!compId) return;
    var comp = plhTxt_(r.Componente);
    var tipo = comp.split('[')[0].trim();
    if (tipo.indexOf('Ponto') === 0) return; // monitoramento topside fora do painel
    var exec = plhData_(r.Dt_Execucao), prazo = plhData_(r.Prazo_Emissao_PATEC);
    var dias = plhDias_(prazo, hoje);
    var aplic = plhTxt_(r.PATEC_Aplic).toLowerCase(), st = plhTxt_(r.Status_PATEC_Sistema);
    var temRel = !!plhTxt_(r.Relatorios);
    var status =
      (aplic.indexOf('não aplicável') >= 0 || aplic.indexOf('somente execução') >= 0) ? 'Não requer PATEC'
      : st === 'Concluído' ? 'Concluído'
      : (!st && exec && exec < limiteAntigo) ? 'Histórico sem PATEC'
      : (dias !== null && dias < 0) ? 'PATEC atrasado'
      : st ? st
      : temRel ? 'PATEC não iniciado' : 'Sem relatório';
    var ultimo = plhNorm_(r.Eh_Ultimo) === 'sim';
    var pendente = ['Concluído', 'Não requer PATEC', 'Histórico sem PATEC'].indexOf(status) < 0;
    if (!ultimo && !pendente) return;
    patec.push({
      plat: plhPlataforma_(r.Instalacao_Resp), objeto: plhTxt_(r.Interligacao), comp: comp, compId: compId,
      proc: plhTxt_(r.Procedimento), exec: plhIso_(exec), patec: plhTxt_(r.PATEC), status: status,
      prazo: plhIso_(prazo), dias: dias, corrosao: plhTxt_(r.Class_Corrosao),
      relatorio: plhCorta_(r.Relatorios, 120), ultimo: ultimo, pendente: pendente
    });
  });

  /* --- Objetos (CIDR + ANM/PLEM que só aparecem nas outras bases) --- */
  var objetos = [], porId = {};
  var extVida = {};
  plan.forEach(function (p) { if (plhNorm_(p.situacao).indexOf('extensao') >= 0) extVida[p.objId] = true; });
  var tipoExtremidade = function (x) {
    var t = plhTxt_(x).toUpperCase();
    if (!t) return '';
    if (t.indexOf('SISTEMA ANM') === 0) return 'ANM';
    if (t.indexOf('EXTREMIDADE LIVRE') >= 0) return 'Extremidade livre';
    if (t === 'PLEM') return 'PLEM';
    return /\d/.test(t) ? 'Plataforma' : 'Terra';
  };
  src.Interligacoes.forEach(function (r) {
    var id = plhTxt_(r.Objeto_BDIEP);
    if (!id || porId[id]) return;
    var tO = tipoExtremidade(r.Inst_Origem), tD = tipoExtremidade(r.Inst_Destino);
    var o = {
      id: id, nome: plhTxt_(r.Nome), funcao: plhTxt_(r.Funcao), plat: plhPlataforma_(r.Instalacao_Resp),
      origem: tO === 'Plataforma' ? plhPlataforma_(r.Inst_Origem) : plhTxt_(r.Inst_Origem),
      destino: tD === 'Plataforma' ? plhPlataforma_(r.Inst_Destino) : plhTxt_(r.Inst_Destino),
      classe: plhTxt_(r.Classe), fluido: plhTxt_(r.Fluido), config: plhTxt_(r.Configuracao),
      anp: plhTxt_(r.Escopo_ANP), spie: plhTxt_(r.SPIE), situacao: plhTxt_(r.Situacao),
      extVida: !!extVida[id], cadastro: 'CIDR'
    };
    porId[id] = o; objetos.push(o);
  });
  var ehEquip = function (n) { var u = plhTxt_(n).toUpperCase(); return u.indexOf('SISTEMA ANM') === 0 || u.indexOf('PLEM') === 0; };
  var addEquip = function (id, nome, plat, situacao) {
    if (!id || porId[id] || !ehEquip(nome)) return;
    var plem = plhTxt_(nome).toUpperCase().indexOf('PLEM') === 0;
    var o = { id: id, nome: nome, funcao: plem ? 'PLEM' : 'Sistema ANM', plat: plat, origem: '', destino: '',
              classe: plem ? 'PLEM' : 'ANM', fluido: '', config: '', anp: '', spie: '', situacao: situacao || '',
              extVida: false, cadastro: 'Fora do CIDR' };
    porId[id] = o; objetos.push(o);
  };
  plan.forEach(function (p) { addEquip(p.objId, p.objeto, p.plat, p.situacao); });
  rts.forEach(function (r) { addEquip(r.objId, r.objeto, r.plat, ''); });
  patec.forEach(function (h) { addEquip(h.compId, h.comp, h.plat, ''); });
  // nome da interligação → id, para ligar o histórico (que vem por nome) ao objeto
  var idPorNome = {};
  objetos.forEach(function (o) { idPorNome[plhNorm_(o.nome)] = o.id; });
  patec.forEach(function (h) { h.objId = idPorNome[plhNorm_(h.objeto)] || h.compId; });

  /* --- Jaquetas: fica a inspeção mais recente de cada plataforma + nível --- */
  // A linha mais nova vale; campos que ela deixou em branco (plano, periodicidade...) vêm da anterior.
  var platsCadastro = {};
  objetos.forEach(function (o) { if (o.plat) platsCadastro[o.plat] = true; });
  var platJaq = function (x) {
    var p = plhPlataforma_(x);
    return (!/\d/.test(p) && platsCadastro[p + '01']) ? p + '01' : p; // PBIQ → PBIQ01
  };
  var jaqMap = {};
  jq.rows.slice().sort(function (a, b) {
    return (a.Dt_Ultima ? a.Dt_Ultima.getTime() : 0) - (b.Dt_Ultima ? b.Dt_Ultima.getTime() : 0);
  }).forEach(function (r) {
    var k = platJaq(r.Plataforma) + '|' + r.Nivel;
    var base = jaqMap[k] || {};
    var novo = {};
    Object.keys(r).forEach(function (c) {
      var v = r[c];
      novo[c] = (v === '' || v === null || v === undefined) ? base[c] : v;
    });
    // observação, ações e lote descrevem a inspeção atual: não herdar
    novo.Observacao = r.Observacao; novo.Acoes = r.Acoes; novo.Lote = r.Lote;
    jaqMap[k] = novo;
  });
  var jaquetas = Object.keys(jaqMap).sort().map(function (k) {
    var r = jaqMap[k];
    var calc = (r.Dt_Ultima && r.Periodicidade_Dias) ? new Date(r.Dt_Ultima.getTime() + r.Periodicidade_Dias * 86400000) : null;
    var ref = r.Dt_Proxima || calc;
    var dias = plhDias_(ref, hoje);
    // Sem última inspeção: vale a próxima programada (SAP); sem nenhuma das duas, está vencida.
    var status = !r.Plano ? 'Sem plano cadastrado'
      : (!r.Dt_Ultima && !ref) ? 'Vencida (sem registro)'
      : (r.Dt_Ultima && r.Dt_Ultima > hoje) ? 'Verificar data (futura)'
      : plhStatusPrazo_(dias, PLH_JANELA_PLANEJ);
    var vencida = status.indexOf('Vencida') === 0;
    var juntas = !!r.Juntas_Criticas;
    var prio = vencida && juntas ? '1 - Crítica' : vencida ? '2 - Alta'
      : ['Planejar agora', 'Verificar data (futura)', 'Sem plano cadastrado'].indexOf(status) >= 0 ? '3 - Média' : '4 - Baixa';
    return {
      plat: platJaq(r.Plataforma), plataforma: r.Plataforma, nivel: r.Nivel, plano: r.Plano,
      per: r.Periodicidade_Dias, ult: plhIso_(r.Dt_Ultima), prox: plhIso_(ref), dias: dias, status: status,
      prio: prio, juntas: r.Juntas_Criticas, obs: r.Observacao, acoes: r.Acoes,
      uf: r.UF, acesso: r.Acesso, zona: r.Zona, niv: r.Niv, statusPlano: r.Status_Plano,
      ordem: r.Ordem, lote: r.Lote, proxOS: r.Prox_OS, semRegistro: !r.Dt_Ultima,
      faltaDU: !!r.Dt_Ultima && plhNorm_(r.DU) !== 'sim',
      reprogramada: !!(r.Dt_Proxima && calc && Math.abs(r.Dt_Proxima - calc) > 3 * 86400000)
    };
  });

  /* --- Fluxo das inspeções: demanda → PLANCOR → OS → execução → relatório no ISUB → PATEC --- */
  var fluxo = plhMontarFluxo_(plan, src.Historico, icpi.rows, idPorNome, hoje);

  /* --- Pendências (lista de trabalho única) --- */
  var pend = [];
  var push = function (p) { p.urg = plhUrgencia_(p.situacao, p.dias); pend.push(p); };
  plan.forEach(function (p) {
    if (p.status !== 'Vencida' && p.status !== 'Planejar agora') return;
    push({ origem: 'Inspeção de duto', plat: p.plat, objId: p.objId, objeto: p.objeto, comp: p.comp,
           atividade: p.proc, descr: p.semAnterior ? 'Primeira inspeção, sem registro anterior' : 'Última inspeção: ' + plhIsoBR_(p.ult),
           prazo: p.prox, dias: p.dias, situacao: p.status, ref: 'Planejamento ' + (p.planej || '-'),
           chave: 'PLAN ' + p.planej });
  });
  rts.forEach(function (r) {
    if (['Vencida', 'Vence em 30 dias', 'Vence em 90 dias'].indexOf(r.status) < 0) return;
    push({ origem: 'RT', plat: r.plat, objId: r.objId, objeto: r.objeto, comp: r.comp,
           atividade: 'RT ' + r.rt + ' · ' + r.cat, descr: r.descr, prazo: r.venc, dias: r.dias,
           situacao: r.status, ref: 'Nota SAP ' + (r.nota || '-') + ' · Prioridade ' + (r.prio || '-'),
           chave: 'RT ' + r.rt });
  });
  patec.forEach(function (h) {
    if (['PATEC atrasado', 'PATEC não iniciado', 'Em elaboração', 'Em aprovação', 'Em verificação', 'Sem relatório'].indexOf(h.status) < 0) return;
    push({ origem: 'PATEC', plat: h.plat, objId: h.objId, objeto: h.objeto, comp: h.comp, atividade: h.proc,
           descr: 'Executada em ' + plhIsoBR_(h.exec) + (h.patec ? ' · ' + h.patec : ''), prazo: h.prazo, dias: h.dias,
           situacao: h.status, ref: h.relatorio || '-', chave: 'PATEC ' + h.compId + ' ' + h.exec });
  });
  jaquetas.forEach(function (j) {
    if (['1 - Crítica', '2 - Alta', '3 - Média'].indexOf(j.prio) < 0) return;
    push({ origem: 'Jaqueta', plat: j.plat, objId: '', objeto: 'Jaqueta ' + j.plataforma, comp: j.nivel,
           atividade: 'Inspeção ' + j.nivel,
           descr: [j.juntas ? 'Juntas críticas: ' + j.juntas : '', j.obs, j.acoes ? 'Ações: ' + j.acoes : ''].filter(String).join(' · '),
           prazo: j.prox, dias: j.dias, situacao: j.status,
           ref: ['Plano ' + (j.plano || '-'), j.proxOS ? 'Próx. OS ' + j.proxOS : '', j.acesso].filter(String).join(' · '),
           chave: 'JAQ ' + j.plat + ' ' + j.nivel });
  });
  pend.sort(function (a, b) {
    if (a.urg !== b.urg) return a.urg < b.urg ? -1 : 1;
    var da = a.dias === null ? 1e9 : a.dias, db = b.dias === null ? 1e9 : b.dias;
    return da - db;
  });

  /* --- Validação --- */
  var valid = [];
  var vistos = {};
  plan.concat(rts).forEach(function (x) {
    if (!x.objId || vistos[x.objId]) return;
    vistos[x.objId] = true;
    if (!porId[x.objId]) valid.push({ objeto: x.objeto, problema: 'Objeto sem cadastro no CIDR' });
  });
  var normSit = function (s) { var t = plhNorm_(s); return t.indexOf('extensao') >= 0 ? 'em operacao' : t; };
  var sitVista = {};
  plan.forEach(function (p) {
    if (sitVista[p.objId]) return;
    sitVista[p.objId] = true;
    var o = porId[p.objId];
    if (o && o.cadastro === 'CIDR' && o.situacao && normSit(o.situacao) !== normSit(p.situacao)) {
      valid.push({ objeto: p.objeto, problema: 'Situação diferente: planejamento "' + p.situacao + '" × CIDR "' + o.situacao + '"' });
    }
  });
  jaquetas.forEach(function (j) {
    var nome = 'Jaqueta ' + j.plataforma + ' ' + j.nivel;
    if (j.status === 'Verificar data (futura)') valid.push({ objeto: nome, problema: 'Última inspeção com data futura' });
    if (j.status === 'Sem plano cadastrado') valid.push({ objeto: nome, problema: 'Sem plano de inspeção na aba Jaquetas' });
  });

  /* --- Plataformas --- */
  var plats = {};
  objetos.forEach(function (o) { if (o.plat) plats[o.plat] = true; });
  jaquetas.forEach(function (j) { if (j.plat) plats[j.plat] = true; });

  fontes.push(icpi.info);
  fontes.forEach(function (f) {
    if (f.base === 'ICPI (OS)') return;
    f.linhas = f.base === 'RT' ? rtTotal : f.base === 'Planejamento' ? plan.length
      : f.base === 'Historico' ? src.Historico.length : f.base === 'Interligacoes' ? src.Interligacoes.length
      : jq.rows.length;
  });

  return {
    geradoEm: Utilities.formatDate(new Date(), ssTimeZone(), 'dd/MM/yyyy HH:mm'),
    planilhaUrl: ss.getUrl(),
    janela: PLH_JANELA_PLANEJ,
    fontes: fontes, plataformas: Object.keys(plats).sort(),
    objetos: objetos, plan: plan, rts: rts, patec: patec, jaquetas: jaquetas,
    pend: pend, valid: valid, fluxo: fluxo,
    etapasFluxo: PLH_ETAPAS
  };
}


/* ============================================================================
 *  FLUXO DAS INSPEÇÕES (uma linha do tempo por inspeção)
 *  Etapas: Demanda · PLANCOR · OS · Execução · Relatório no ISUB · Elaboração · Verificação · Aprovação
 *  Estados de cada etapa: ok (feita), atual (onde está parada), pend (ainda não chegou),
 *  na (não se aplica), sem (sem fonte de dados ainda).
 * ==========================================================================*/
var PLH_ETAPAS = ['Demanda', 'PLANCOR', 'OS', 'Execução', 'Relatório ISUB', 'Elaboração', 'Verificação', 'Aprovação'];
var PLH_FLUXO_HORIZONTE = 730;   // demandas com vencimento em até 24 meses (a tela filtra)
var PLH_FLUXO_EXEC_DIAS = 365;   // execuções do último ano ainda em tratamento
var PLH_FLUXO_CONCL_DIAS = 60;   // concluídas recentemente continuam visíveis por 60 dias

/** Onde a inspeção acontece: as UEPs (PXA-3 → PXA03, PAT1 → PAT01); sem UEP (trecho no leito), a plataforma responsável. */
function plhLocais_(ueps, platResp) {
  var l = [];
  (ueps || []).forEach(function (u) { var p = plhPlataforma_(u); if (p && l.indexOf(p) < 0) l.push(p); });
  return l.length ? l : (platResp ? [platResp] : []);
}

function plhMontarFluxo_(plan, historico, icpiRows, idPorNome, hoje) {
  var fluxo = [];
  var dBR = function (d) { return d ? Utilities.formatDate(d, ssTimeZone(), 'dd/MM/yyyy') : ''; };
  var etapa = function (estado, data, info) { return { estado: estado, data: data || '', info: info || '' }; };

  // ICPI agrupado por objeto + procedimento
  var icpiPor = {};
  icpiRows.forEach(function (r) {
    var k = plhNorm_(r.objeto) + '|' + plhNorm_(r.proc);
    (icpiPor[k] = icpiPor[k] || []).push(r);
  });
  var acharICPI = function (objeto, proc, proxIso) {
    var lista = icpiPor[plhNorm_(objeto) + '|' + plhNorm_(proc)] || [];
    if (!proxIso) return [];
    var alvo = new Date(proxIso + 'T00:00:00').getTime();
    return lista.filter(function (r) { return r.venc && Math.abs(r.venc.getTime() - alvo) <= 3 * 86400000; });
  };

  // 1) Demandas: planejamentos vencidos ou que vencem no horizonte,
  //    agrupados como no ICPI (objeto + procedimento + vencimento; vários componentes viram uma demanda)
  var grupos = {}, ordem = [];
  plan.forEach(function (p) {
    if (p.dias === null || p.dias > PLH_FLUXO_HORIZONTE) return;
    if (plhNorm_(p.situacao).indexOf('permanentemente') >= 0) return;
    var k = plhNorm_(p.objeto) + '|' + plhNorm_(p.proc) + '|' + p.prox;
    if (!grupos[k]) { grupos[k] = { p: p, comps: [], planejs: [] }; ordem.push(k); }
    grupos[k].comps.push(p.comp); grupos[k].planejs.push(p.planej);
  });
  ordem.forEach(function (k) {
    var g = grupos[k], p = g.p;
    var cand = acharICPI(p.objeto, p.proc, p.prox);
    var comOS = cand.filter(function (r) { return r.os; });
    var ueps = cand.map(function (r) { return r.uep; }).filter(String);
    var replanejar = comOS.some(function (r) { return /replanej/i.test(r.statusOS); });
    var infoOS;
    if (!cand.length) infoOS = 'Sem OS (demanda não encontrada no ICPI)';
    else if (cand.length === 1) infoOS = comOS.length ? 'OS ' + comOS[0].os + (comOS[0].statusOS ? ' · ' + comOS[0].statusOS : '') : 'Sem OS';
    else infoOS = cand.map(function (r) {
      return (r.uep || 'trecho') + ': ' + (r.os ? 'OS ' + r.os + (r.statusOS ? ' · ' + r.statusOS : '') : 'sem OS');
    }).join(' | ');
    var temOS = comOS.length > 0, todasComOS = cand.length > 0 && comOS.length === cand.length;
    var vencida = p.dias < 0;
    // Com OS em todas as pontas: OS feita, parada em Execução (programação/detalhamento/carteira do barco).
    // Sem OS (ou só em parte das pontas): parada em OS.
    var eOS, eExec, atualIdx;
    if (todasComOS) {
      eOS = etapa('ok', '', comOS.map(function (r) { return 'OS ' + r.os + (cand.length > 1 && r.uep ? ' (' + r.uep + ')' : ''); }).join(', '));
      eExec = etapa(replanejar || vencida ? 'alerta' : 'atual', '', 'Aguardando execução · ' + infoOS);
      atualIdx = 3;
    } else {
      eOS = etapa(vencida || replanejar ? 'alerta' : 'atual', '', infoOS);
      eExec = etapa('pend');
      atualIdx = 2;
    }
    var e = [
      etapa('ok', p.prox, vencida ? 'Vencida há ' + (-p.dias) + ' dias' : 'Vence em ' + p.dias + ' dias'),
      etapa('sem', '', 'Ainda sem fonte do PLANCOR'),
      eOS, eExec, etapa('pend'), etapa('pend'), etapa('pend'), etapa('pend')
    ];
    var stOS = comOS.map(function (r) { return r.statusOS; }).filter(String);
    fluxo.push({
      tipo: 'demanda', chave: 'FLX ' + g.planejs.join('+'), plat: p.plat, objId: p.objId, objeto: p.objeto,
      comp: g.comps.join(' · '), proc: p.proc, uep: ueps.join(', '), planej: g.planejs.join(', '),
      locais: plhLocais_(ueps, p.plat),
      prazo: p.prox, dataRef: p.prox, dias: p.dias, os: comOS.map(function (r) { return r.os; }).join(', '),
      statusOS: stOS.join(', '), etapaAtual: atualIdx, etapas: e,
      situacao: !temOS ? 'Aguardando OS' : todasComOS ? 'OS ' + (stOS[0] || 'aberta') : 'OS parcial',
      alerta: p.dias < 0 || replanejar
    });
  });

  // 2) Execuções do último ano ainda em tratamento (relatório / PATEC)
  var limiteExec = new Date(hoje.getTime() - PLH_FLUXO_EXEC_DIAS * 86400000);
  var limiteConcl = new Date(hoje.getTime() - PLH_FLUXO_CONCL_DIAS * 86400000);
  var planPorCodigo = {};
  plan.forEach(function (p) { if (p.planej) planPorCodigo[p.planej] = p; });
  historico.forEach(function (r) {
    var comp = plhTxt_(r.Componente);
    if (!plhTxt_(r.Componente_BDIEP) || comp.indexOf('Ponto') === 0) return;
    var exec = plhData_(r.Dt_Execucao);
    if (!exec || exec < limiteExec) return;
    var aplic = plhTxt_(r.PATEC_Aplic).toLowerCase();
    var semPatec = aplic.indexOf('não aplicável') >= 0 || aplic.indexOf('somente execução') >= 0;
    var st = plhTxt_(r.Status_PATEC_Sistema);
    var cad = plhData_(r.Dt_Cadastro_Relatorio), ini = plhData_(r.Dt_Inicio_Elab), envV = plhData_(r.Dt_Envio_Verif),
        envA = plhData_(r.Dt_Envio_Aprov), concl = plhData_(r.Dt_Conclusao), prazo = plhData_(r.Prazo_Emissao_PATEC);
    var concluido = st === 'Concluído' || !!concl;
    // fora da tela: concluído há mais de 60 dias, ou sem PATEC com relatório cadastrado há mais de 60 dias
    if (concluido && (concl || cad || exec) < limiteConcl) return;
    if (semPatec && cad && cad < limiteConcl) return;

    var dias = plhDias_(prazo, hoje);
    var e = [etapa('ok', '', 'Atendida'), etapa('ok', '', ''), etapa('ok', '', ''),
             etapa('ok', plhIso_(exec), 'Executada em ' + dBR(exec))];
    var atual;
    if (!cad) { e.push(etapa('atual', '', 'Relatório ainda não cadastrado')); atual = 4; }
    else e.push(etapa('ok', plhIso_(cad), 'Cadastrado em ' + dBR(cad) + (plhTxt_(r.Resp_Relatorio) ? ' por ' + plhTxt_(r.Resp_Relatorio).split(/[\r\n]/)[0] : '')));
    if (semPatec) {
      e.push(etapa('na', '', 'PATEC não aplicável'), etapa('na'), etapa('na'));
      if (atual === undefined) atual = 8;
    } else {
      var pessoa = function (x) { return plhTxt_(x).split(/[\r\n]/)[0]; };
      var passo = function (feito, emAndamento, dataFeito, quem, rotulo) {
        if (feito) return etapa('ok', plhIso_(dataFeito), rotulo + (quem ? ' · ' + quem : ''));
        if (emAndamento) return etapa(dias !== null && dias < 0 ? 'alerta' : 'atual', '', (quem ? quem + ' · ' : '') + 'em andamento');
        return etapa('pend');
      };
      var elabFeita = !!envV || concluido, verifFeita = !!envA || concluido, aprovFeita = concluido;
      var elab = passo(elabFeita, cad && !elabFeita, envV, pessoa(r.Elaborador), 'Enviado p/ verificação ' + dBR(envV));
      if (cad && !elabFeita && !ini && !st) elab = etapa(dias !== null && dias < 0 ? 'alerta' : 'atual', '', 'PATEC não iniciado');
      e.push(elab);
      e.push(passo(verifFeita, elabFeita && !verifFeita, envA, pessoa(r.Verificador), 'Enviado p/ aprovação ' + dBR(envA)));
      e.push(passo(aprovFeita, verifFeita && !aprovFeita, concl, pessoa(r.Aprovador), 'Aprovado ' + dBR(concl)));
      if (atual === undefined) atual = aprovFeita ? 8 : verifFeita ? 7 : elabFeita ? 6 : 5;
    }
    var p = planPorCodigo[plhTxt_(r.Cod_Planej_Posterior)];
    var interlig = plhTxt_(r.Interligacao);
    var uepsExec = p ? acharICPI(p.objeto, p.proc, p.prox).map(function (c) { return c.uep; }).filter(String) : [];
    var platResp = plhPlataforma_(r.Instalacao_Resp);
    var situacao = atual === 8 ? (semPatec ? 'Relatório cadastrado (sem PATEC)' : 'PATEC concluído')
      : atual === 4 ? 'Aguardando cadastro do relatório'
      : ['', '', '', '', '', 'PATEC em elaboração', 'PATEC em verificação', 'PATEC em aprovação'][atual];
    if (atual === 5 && !ini && !st) situacao = 'PATEC não iniciado';
    fluxo.push({
      tipo: 'execucao', chave: 'FLX EXEC ' + plhTxt_(r.Componente_BDIEP) + ' ' + plhIso_(exec),
      plat: plhPlataforma_(r.Instalacao_Resp), objId: idPorNome[plhNorm_(interlig)] || (p ? p.objId : ''),
      objeto: interlig, comp: comp, proc: plhTxt_(r.Procedimento), uep: uepsExec.join(', '),
      locais: plhLocais_(uepsExec, platResp), planej: plhTxt_(r.Cod_Planej_Posterior),
      prazo: semPatec ? '' : plhIso_(prazo), dataRef: plhIso_(exec), dias: semPatec ? null : dias, os: '', statusOS: '',
      patec: plhTxt_(r.PATEC), etapaAtual: atual, etapas: e, situacao: situacao,
      alerta: !semPatec && !concluido && dias !== null && dias < 0
    });
  });

  fluxo.sort(function (a, b) {
    if (a.etapaAtual !== b.etapaAtual) return a.etapaAtual - b.etapaAtual;
    var da = a.dias === null ? 1e9 : a.dias, db = b.dias === null ? 1e9 : b.dias;
    return da - db;
  });
  return fluxo;
}
