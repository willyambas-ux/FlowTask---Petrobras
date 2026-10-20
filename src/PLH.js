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
var PLH_ABA_JAQUETAS = 'Inspecoes_Jaqueta';

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
  var a = Number(m[3]); if (a < 100) a += 2000;
  var d = new Date(a, Number(m[2]) - 1, Number(m[1]));
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

/** Inspecoes_Jaqueta: lida pelo nome do cabeçalho (é controle próprio e pode ganhar colunas). */
function plhLerJaquetas_(ss) {
  var props = PropertiesService.getDocumentProperties();
  var info = { base: 'Jaquetas', sistema: 'Controle próprio', aba: PLH_ABA_JAQUETAS, linhas: 0,
               ok: false, erro: '', atualizado: props.getProperty('PLH_ATUALIZADO_' + PLH_ABA_JAQUETAS) || '' };
  var sh = ss.getSheetByName(PLH_ABA_JAQUETAS);
  if (!sh) { info.erro = 'Aba não encontrada'; return { info: info, rows: [] }; }
  var vals = sh.getDataRange().getValues();
  if (vals.length < 2) { info.erro = 'Aba vazia'; return { info: info, rows: [] }; }
  var cab = vals[0].map(plhNorm_);
  var col = function (nome) { return cab.indexOf(nome); };
  var ix = {
    plataforma: col('plataforma'), nivel: col('tipo inspecao'), plano: col('plano'),
    perDias: col('periodicidade dias'), ultima: col('ultima inspecao'), proxima: col('proxima inspecao'),
    lote: col('lote inspecao'), atual: col('atual/final'), status: col('status'),
    obs: col('observacao'), juntas: col('juntas criticas')
  };
  if (ix.plataforma < 0 || ix.ultima < 0) { info.erro = 'Cabeçalho não confere'; return { info: info, rows: [] }; }
  var g = function (r, k) { return ix[k] < 0 ? '' : r[ix[k]]; };
  var rows = [];
  for (var i = 1; i < vals.length; i++) {
    var r = vals[i];
    if (!plhTxt_(g(r, 'plataforma'))) continue;
    rows.push({
      Plataforma: plhTxt_(g(r, 'plataforma')), Nivel: plhTxt_(g(r, 'nivel')), Plano: plhTxt_(g(r, 'plano')),
      Periodicidade_Dias: plhNum_(g(r, 'perDias')), Dt_Ultima: plhData_(g(r, 'ultima')),
      Dt_Proxima: plhData_(g(r, 'proxima')), Lote: plhTxt_(g(r, 'lote')), Atual: plhTxt_(g(r, 'atual')),
      Status_Manual: plhTxt_(g(r, 'status')), Observacao: plhTxt_(g(r, 'obs')),
      Juntas_Criticas: plhTxt_(g(r, 'juntas'))
    });
  }
  info.ok = true;
  info.temJuntas = ix.juntas >= 0;
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
    // observação e status manual descrevem a inspeção atual: não herdar
    novo.Observacao = r.Observacao; novo.Status_Manual = r.Status_Manual; novo.Lote = r.Lote;
    jaqMap[k] = novo;
  });
  var jaquetas = Object.keys(jaqMap).sort().map(function (k) {
    var r = jaqMap[k];
    var calc = (r.Dt_Ultima && r.Periodicidade_Dias) ? new Date(r.Dt_Ultima.getTime() + r.Periodicidade_Dias * 86400000) : null;
    var ref = r.Dt_Proxima || calc;
    var dias = plhDias_(ref, hoje);
    var status = !r.Plano ? 'Sem plano cadastrado'
      : !r.Dt_Ultima ? 'Vencida (sem registro)'
      : r.Dt_Ultima > hoje ? 'Verificar data (futura)'
      : plhStatusPrazo_(dias, PLH_JANELA_PLANEJ);
    var vencida = status.indexOf('Vencida') === 0;
    var juntas = !!r.Juntas_Criticas;
    var prio = vencida && juntas ? '1 - Crítica' : vencida ? '2 - Alta'
      : ['Planejar agora', 'Verificar data (futura)', 'Sem plano cadastrado'].indexOf(status) >= 0 ? '3 - Média' : '4 - Baixa';
    return {
      plat: platJaq(r.Plataforma), plataforma: r.Plataforma, nivel: r.Nivel, plano: r.Plano,
      per: r.Periodicidade_Dias, ult: plhIso_(r.Dt_Ultima), prox: plhIso_(ref), dias: dias, status: status,
      prio: prio, juntas: r.Juntas_Criticas, obs: r.Observacao,
      faltaDU: r.Status_Manual.toUpperCase().indexOf('FALTA DU') >= 0,
      divergencia: !!(r.Dt_Proxima && calc && Math.abs(r.Dt_Proxima - calc) > 3 * 86400000)
    };
  });

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
           atividade: 'Inspeção ' + j.nivel, descr: [j.obs, j.juntas ? 'Juntas críticas: ' + j.juntas : ''].filter(String).join(' · '),
           prazo: j.prox, dias: j.dias, situacao: j.status, ref: 'Plano ' + (j.plano || '-'),
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
    if (j.status === 'Verificar data (futura)') valid.push({ objeto: 'Jaqueta ' + j.plataforma + ' ' + j.nivel, problema: 'Última inspeção com data futura' });
    else if (j.divergencia) valid.push({ objeto: 'Jaqueta ' + j.plataforma + ' ' + j.nivel, problema: 'Próxima inspeção ≠ última + periodicidade' });
  });

  /* --- Plataformas --- */
  var plats = {};
  objetos.forEach(function (o) { if (o.plat) plats[o.plat] = true; });
  jaquetas.forEach(function (j) { if (j.plat) plats[j.plat] = true; });

  fontes.forEach(function (f) {
    f.linhas = f.base === 'RT' ? rtTotal : f.base === 'Planejamento' ? plan.length
      : f.base === 'Historico' ? src.Historico.length : f.base === 'Interligacoes' ? src.Interligacoes.length
      : jq.rows.length;
  });

  return {
    geradoEm: Utilities.formatDate(new Date(), ssTimeZone(), 'dd/MM/yyyy HH:mm'),
    janela: PLH_JANELA_PLANEJ,
    fontes: fontes, plataformas: Object.keys(plats).sort(),
    objetos: objetos, plan: plan, rts: rts, patec: patec, jaquetas: jaquetas,
    pend: pend, valid: valid
  };
}
