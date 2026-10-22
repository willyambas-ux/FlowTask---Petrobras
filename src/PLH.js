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
       'Classe','Fluido','Configuracao','Escopo_ANP','SPIE','Situacao','Agente_Descont','Ano_Construcao',
       'Ano_Operacao','Atestado_Descomiss','Capacidade','Caract_Fluido','Classe_Pressao','Classif_Duto',
       'Classif_Sistema','Codigo_Duto','Desinvestimento','Just_PMO','Just_PMOA','Local_Descont','Local_Instalacao',
       'Motivo_NaoPig_Limpeza','Motivo_NaoPig_Inspecao','Necessario_PMCI','Norma','PDID','Pig_Inspecao',
       'Pig_Limpeza','PMO','PMOA','Qual_Descont','Rotina_Pig','RTA_Descont','TAG_PSH','Tem_Descont']
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

function plhNulo_(v) { var t = plhTxt_(v); return /^null$/i.test(t) ? '' : t.replace(/\.0$/, ''); }
/** número vindo do CIDR: célula numérica direto; texto em formato brasileiro ("3031,398") pelo plhNum_ */
function plhNumC_(v) { if (typeof v === 'number') return v; var t = plhNulo_(v); return t ? plhNum_(t) : null; }
function plhAno_(v) { var n = plhNum_(plhNulo_(v)); return n && n > 1900 && n < 2100 ? Math.round(n) : null; }
/** 'OL 12" PAT-2/PXA-3' → 12 */
function plhDiametro_(nome) { var m = plhTxt_(nome).match(/(\d+(?:[.,]\d+)?)\s*"/); return m ? plhNum_(m[1]) : null; }

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

/** SRC_Componentes (CIDR, aba de componentes; opcional): todos os componentes de cada duto,
 *  inclusive os sem planejamento. Lido pelo nome do cabeçalho (o export pode mudar de ordem).
 *  O código BDIEP é reconhecido pelo conteúdo (OBJBDIEP:nnn): o que for de duto vira o duto, o resto é o componente. */
var PLH_ABA_COMPCIDR = 'SRC_Componentes';
function plhLerCompCIDR_(ss, objetos, plan) {
  var props = PropertiesService.getDocumentProperties();
  var info = { base: 'Componentes', sistema: 'CIDR', aba: PLH_ABA_COMPCIDR, linhas: 0, ok: false, erro: '', opcional: true,
               atualizado: props.getProperty('PLH_ATUALIZADO_' + PLH_ABA_COMPCIDR) || '' };
  var res = { info: info, porDuto: {}, total: 0, semDuto: 0 };
  var sh = ss.getSheetByName(PLH_ABA_COMPCIDR);
  if (!sh) { info.erro = 'Aba não encontrada (opcional)'; return res; }
  info.gid = sh.getSheetId();
  var vals = sh.getDataRange().getValues();
  var h = -1;
  for (var i = 0; i < Math.min(vals.length, 20); i++) {
    if (vals[i].map(plhNorm_).some(function (c) { return /^extensao/.test(c); })) { h = i; break; }
  }
  if (h < 0) { info.erro = vals.length > 1 ? 'Cabeçalho não encontrado (coluna Extensão)' : 'Aba vazia (opcional): cole o export de componentes do CIDR'; return res; }
  var cab = vals[h].map(plhNorm_);
  var achar = function (re, nao) {
    for (var k = 0; k < cab.length; k++) if (re.test(cab[k]) && !(nao && nao.test(cab[k]))) return k;
    return -1;
  };
  var ix = {
    duto: achar(/^(nome|interligacao|nome da interligacao|nome do duto|duto|objeto)$/),
    comp: achar(/componente/, /tipo|bdiep|codigo/),
    tipo: achar(/tipo/), ordem: achar(/^ordem$/),
    ext: achar(/^extensao/), esp: achar(/espessura/), vida: achar(/vida util/), dn: achar(/diametro nominal/),
    anp: achar(/escopo anp/), sit: achar(/^situacao/), local: achar(/local de instalacao/),
    resp: achar(/inst\. resp|instalacao resp/), orig: achar(/origem/), dest: achar(/destino/)
  };
  var idObj = {}, idPorNome = {};
  objetos.forEach(function (o) { idObj[o.id] = o; idPorNome[plhNorm_(o.nome)] = o.id; });
  var g = function (r, k) { return ix[k] < 0 ? '' : r[ix[k]]; };
  var normPonta = function (x) { var t = plhTxt_(x); return plhEhPlataforma_(t) ? plhPlataforma_(t) : plhNorm_(t); };
  var linhasC = [];
  for (var j = h + 1; j < vals.length; j++) {
    var r = vals[j];
    if (!r.some(function (v) { return plhTxt_(v) !== ''; })) continue;
    var ids = [];
    r.forEach(function (v) { var m = plhTxt_(v).match(/OBJBDIEP:\s*\d+/i); if (m) ids.push(m[0].toUpperCase().replace(/\s/g, '')); });
    var dutoId = ids.filter(function (x) { return idObj[x]; })[0] || idPorNome[plhNorm_(g(r, 'duto'))] || '';
    var compNome = plhTxt_(g(r, 'comp'));
    var compId = ids.filter(function (x) { return x !== dutoId; })[0] || '';
    if (!compId) { var b = compNome.match(/\[(\d+)\]/); if (b) compId = 'OBJBDIEP:' + b[1]; }
    var tipo = plhTxt_(g(r, 'tipo')) || (compNome ? compNome.split('[')[0].trim() : '');
    linhasC.push({ dutoId: dutoId, id: compId, nome: compNome, tipo: tipo,
              ext: plhNumC_(g(r, 'ext')), esp: plhNumC_(g(r, 'esp')), vida: plhNumC_(g(r, 'vida')),
              dn: plhNumC_(g(r, 'dn')), anp: plhNulo_(g(r, 'anp')), situacao: plhNulo_(g(r, 'sit')),
              local: plhEhPlataforma_(g(r, 'local')) ? plhPlataforma_(g(r, 'local')) : '',
              ordem: plhNumC_(g(r, 'ordem')),
              trio: [normPonta(g(r, 'resp')), normPonta(g(r, 'orig')), normPonta(g(r, 'dest'))].join('|') });
  }
  res.total = linhasC.length;

  // Sem nome/código do duto no export: liga pelos blocos de linhas seguidas com a mesma
  // Inst. Resp. + Origem + Destino, desempatando pelo diâmetro (do nome do duto) — cabo/umbilical não tem diâmetro.
  if (linhasC.length && !linhasC.some(function (c) { return c.dutoId; })) {
    var blocos = [];
    linhasC.forEach(function (c) {
      var u = blocos[blocos.length - 1];
      if (u && u.trio === c.trio) u.linhas.push(c); else blocos.push({ trio: c.trio, linhas: [c], usado: false });
    });
    blocos.forEach(function (bl) {
      var cont = {};
      bl.linhas.forEach(function (c) { if (c.dn) cont[c.dn] = (cont[c.dn] || 0) + 1; });
      var ks = Object.keys(cont).sort(function (a, b) { return cont[b] - cont[a]; });
      bl.dn = ks.length ? Number(ks[0]) : null;
    });
    objetos.filter(function (o) { return o.cadastro === 'CIDR'; }).forEach(function (o) {
      var trio = [o.plat, o.origem, o.destino].map(normPonta).join('|');
      var flex = /flex/i.test(String(o.config || '').replace(/\u00ad/g, ''));
      var bl = blocos.filter(function (b) {
        if (b.usado || b.trio !== trio) return false;
        if (o.diam) return b.dn === o.diam;
        return flex ? !b.dn : true;
      })[0];
      if (!bl) return;
      bl.usado = true;
      bl.linhas.forEach(function (c) { c.dutoId = o.id; c.porBloco = true; });
    });
  }

  // componente sem código: casa com o planejamento pela extensão (mesmo duto)
  var planPor = {};
  (plan || []).forEach(function (p) {
    if (!p.objId || !p.compId || p.ext === null || p.ext === undefined) return;
    var lst = planPor[p.objId] = planPor[p.objId] || [];
    if (!lst.some(function (x) { return x.compId === p.compId; })) lst.push({ compId: p.compId, comp: p.comp, tipo: p.tipoComp, ext: p.ext, usado: false });
  });
  linhasC.forEach(function (c) {
    if (c.id || !c.dutoId || c.ext === null) return;
    var m = (planPor[c.dutoId] || []).filter(function (x) { return !x.usado && Math.abs(x.ext - c.ext) < 0.01; })[0];
    if (m) { m.usado = true; c.id = m.compId; c.nome = m.comp; c.tipo = m.tipo; c.porExtensao = true; }
  });
  linhasC.forEach(function (c) {
    if (!c.dutoId) { res.semDuto++; return; }
    if (!c.nome) c.nome = (c.tipo || 'Componente CIDR') + (c.id ? ' [' + c.id.split(':')[1] + ']' : '');
    (res.porDuto[c.dutoId] = res.porDuto[c.dutoId] || []).push(c);
  });
  info.ok = true;
  info.linhas = res.total;
  if (res.total && res.semDuto === res.total) info.aviso = 'Nenhuma linha ligada a um duto: inclua a coluna com o nome ou o código BDIEP do duto';
  else if (res.semDuto) info.aviso = res.semDuto + ' linha(s) sem duto identificado';
  if (res.total && ix.comp < 0 && !Object.keys(res.porDuto).some(function (k) { return res.porDuto[k].some(function (c) { return c.id; }); }))
    info.aviso = (info.aviso ? info.aviso + '; ' : '') + 'sem código BDIEP do componente: não dá para ligar ao planejamento';
  return res;
}

/** SRC_PLANCOR (opcional): planilha de insumos da última implantação, colada como está.
 *  Lida pelo nome do cabeçalho. Só interessa saber se cada demanda está ou não no PLANCOR. */
var PLH_ABA_PLANCOR = 'SRC_PLANCOR';
function plhLerPlancor_(ss) {
  var props = PropertiesService.getDocumentProperties();
  var info = { base: 'PLANCOR', sistema: 'Insumos PLANCOR', aba: PLH_ABA_PLANCOR, linhas: 0, ok: false, erro: '',
               opcional: true, semIdade: true, // muda só a cada implantação: fica fora do aviso de "base antiga"
               atualizado: props.getProperty('PLH_ATUALIZADO_' + PLH_ABA_PLANCOR) || '' };
  var sh = ss.getSheetByName(PLH_ABA_PLANCOR);
  if (!sh) { info.erro = 'Aba não encontrada (opcional)'; return { info: info, rows: [] }; }
  info.gid = sh.getSheetId();
  var vals = sh.getDataRange().getValues();
  // o cabeçalho pode não estar na linha 1 (texto colado antes): procura a linha com "Tarefa" e "UEP"
  var h = -1;
  for (var i = 0; i < Math.min(vals.length, 20); i++) {
    var n = vals[i].map(plhNorm_);
    if (n.indexOf('tarefa') >= 0 && n.indexOf('uep') >= 0) { h = i; break; }
  }
  if (h < 0) { info.erro = 'Cabeçalho não encontrado (Tarefa, UEP, Duto, Observações)'; return { info: info, rows: [] }; }
  var cab = vals[h].map(plhNorm_);
  var col = function () {
    for (var k = 0; k < arguments.length; k++) { var c = cab.indexOf(arguments[k]); if (c >= 0) return c; }
    return -1;
  };
  var ix = { tarefa: col('tarefa'), uep: col('uep'), duto: col('duto'), equip: col('equipamento'),
             obs: col('observacoes', 'observacao'), ini: col('data inicio tarefa'), fim: col('data fim tarefa'),
             ref: col('numero referencia insumo'), projeto: col('projeto') };
  var g = function (r, k) { return ix[k] < 0 ? '' : r[ix[k]]; };
  var rows = [];
  for (var j = h + 1; j < vals.length; j++) {
    var r = vals[j];
    var tarefa = plhTxt_(g(r, 'tarefa'));
    if (!tarefa || plhNorm_(tarefa) === 'tarefa') continue;
    rows.push({ tarefa: tarefa, uep: plhTxt_(g(r, 'uep')), duto: plhTxt_(g(r, 'duto')), equip: plhTxt_(g(r, 'equip')),
                obs: plhTxt_(g(r, 'obs')), ini: plhData_(g(r, 'ini')), fim: plhData_(g(r, 'fim')),
                ref: plhTxt_(g(r, 'ref')), projeto: plhTxt_(g(r, 'projeto')), linha: j + 1 });
  }
  info.ok = true;
  info.linhas = rows.length;
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
  var plancor = plhLerPlancor_(ss);

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
      semAnterior: !ult, descCritica: plhNorm_(r.Tem_Descont_Critica) === 'sim',
      tipoComp: plhTxt_(r.Tipo_Componente), vale: plhTxt_(r.Vale_Para), ext: plhNum_(r.Extensao_m),
      local: plhPlataforma_(r.Local_Instalacao), tramo: plhTxt_(r.Tramo_Suporte)
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
      extVida: !!extVida[id], cadastro: 'CIDR',
      anoOper: plhAno_(r.Ano_Operacao) || plhAno_(r.Ano_Construcao), anoConstr: plhAno_(r.Ano_Construcao),
      diam: plhDiametro_(r.Nome), sistema: plhTxt_(r.Classif_Sistema), norma: plhNulo_(r.Norma),
      pressao: plhNulo_(r.Classe_Pressao), pigInsp: plhNulo_(r.Pig_Inspecao), pigLimp: plhNulo_(r.Pig_Limpeza),
      pmo: plhNulo_(r.PMO), pmoa: plhNulo_(r.PMOA)
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
  var extPor = {}, compVisto = {}, procsPor = {};
  plan.forEach(function (p) {
    if (!p.objId) return;
    (procsPor[p.objId] = procsPor[p.objId] || {})[p.proc] = true;
    if (!p.compId || compVisto[p.compId] || !p.ext) return;
    compVisto[p.compId] = true;
    extPor[p.objId] = (extPor[p.objId] || 0) + p.ext;
  });
  objetos.forEach(function (o) {
    o.extM = extPor[o.id] ? Math.round(extPor[o.id]) : null;
    o.procs = Object.keys(procsPor[o.id] || {}).sort();
    if (o.cadastro === 'CIDR' && o.situacao && plan.some(function (p) { return p.objId === o.id; })) {
      var sitG = plan.filter(function (p) { return p.objId === o.id; })[0].situacao;
      o.sitGides = sitG;
    }
  });
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
  var nomesObj = {};
  plan.forEach(function (p) { nomesObj[plhNorm_(p.objeto)] = p.objeto; });
  objetos.forEach(function (o) { nomesObj[plhNorm_(o.nome)] = o.nome; });
  var plancorRes = plhLigarPlancor_(plancor, fluxo, rts, jaquetas, nomesObj, plan, hoje);
  if (plancor.info.ok) plancor.info.linhas = plancorRes.suas + ' suas (de ' + plancorRes.linhas + ')';
  var icpiPainel = plhMontarIcpi_(ss, icpi.rows, fluxo, hoje);
  var compCIDR = plhLerCompCIDR_(ss, objetos, plan);
  var matriz = plhMontarMatriz_(ss, plan, objetos, icpi.rows, compCIDR);

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
           situacao: r.status, ref: 'Nota SAP ' + (r.nota || '-') + ' · Prioridade ' + (r.prio || '-') + (r.plancor ? ' · No PLANCOR (' + r.plancor + ')' : ''),
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
           ref: ['Plano ' + (j.plano || '-'), j.proxOS ? 'Próx. OS ' + j.proxOS : '', j.acesso, j.plancor ? 'No PLANCOR (' + j.plancor + ')' : ''].filter(String).join(' · '),
           chave: 'JAQ ' + j.plat + ' ' + j.nivel });
  });
  fluxo.forEach(function (x) {
    if (x.tipo !== 'demanda' || x.plancor !== 'nao' || x.dias === null || x.dias > 365) return;
    push({ origem: 'PLANCOR', plat: x.plat, objId: x.objId, objeto: x.objeto, comp: x.comp,
           atividade: x.proc + ' fora do PLANCOR',
           descr: 'Demanda que vence em ' + plhIsoBR_(x.prazo) + ' não foi encontrada na última implantação do PLANCOR',
           prazo: x.prazo, dias: x.dias, situacao: 'Fora do PLANCOR', ref: 'Planejamento ' + (x.planej || '-'),
           chave: 'PLANCOR ' + x.planej });
  });
  jaquetas.forEach(function (j) {
    if (!j.foraPlancor) return;
    push({ origem: 'PLANCOR', plat: j.plat, objId: '', objeto: 'Jaqueta ' + j.plataforma, comp: j.nivel,
           atividade: 'Inspeção ' + j.nivel + ' fora do PLANCOR',
           descr: 'Próxima inspeção em ' + plhIsoBR_(j.prox) + ' e nenhuma linha da jaqueta com esse nível no PLANCOR',
           prazo: j.prox, dias: j.dias, situacao: 'Fora do PLANCOR', ref: 'Plano ' + (j.plano || '-'),
           chave: 'PLANCOR JAQ ' + j.plat + ' ' + j.nivel });
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

  plancorRes.semDemanda.forEach(function (l) {
    valid.push({ objeto: l.objeto, problema: 'No PLANCOR (insumo ' + (l.ref || '-') + ', ' + l.proc + (l.lado ? ' lado ' + l.lado : '') +
      ', início ' + plhIsoBR_(l.ini) + ') sem demanda correspondente no GIDES' });
  });

  /* --- Plataformas --- */
  var plats = {};
  objetos.forEach(function (o) { if (o.plat) plats[o.plat] = true; });
  jaquetas.forEach(function (j) { if (j.plat) plats[j.plat] = true; });

  fontes.push(icpi.info);
  fontes.push(plancor.info);
  fontes.push(compCIDR.info);
  fontes.forEach(function (f) {
    if (f.base === 'ICPI (OS)' || f.base === 'PLANCOR' || f.base === 'Componentes') return;
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
    pend: pend, valid: valid, fluxo: fluxo, icpi: icpiPainel, plancor: plancorRes, matriz: matriz,
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
    var pessoa1 = function (x) { return plhTxt_(x).split(/[\r\n]/)[0]; };
    var comQuem = atual === 4 ? 'Cadastro do relatório' : atual === 5 ? pessoa1(r.Elaborador)
      : atual === 6 ? pessoa1(r.Verificador) : atual === 7 ? pessoa1(r.Aprovador) : '';
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
      patec: plhTxt_(r.PATEC), etapaAtual: atual, etapas: e, situacao: situacao, comQuem: comQuem,
      naoIniciado: atual === 5 && !ini && !st,
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


/* ============================================================================
 *  PLANCOR (etapa 2 do fluxo): a demanda está na última implantação?
 *  - ok    : linha que identifica o duto (coluna Duto/Observações, com o lado) + procedimento
 *  - prov  : linha genérica ("Duto rígido de produção de óleo") do mesmo campo + procedimento, perto da data
 *  - nao   : nada encontrado
 *  - na    : inspeção emersa (PIDR-4 / PIDF-4), feita pelo inspetor, não vai para o PLANCOR
 *  O PLANCOR registra a plataforma dona do projeto (PXA01), não o lado: comparação por campo (PXA, PAT...).
 * ==========================================================================*/
var PLH_PLANCOR_NA = /^(PID[RF]-4\b|FAIXA|EXT_)/i; // emersa, faixa em terra, extensão de vida: sem barco
var PLH_PLANCOR_JANELA_ANTES = 365;  // dias: linha do PLANCOR até 1 ano antes do vencimento
var PLH_PLANCOR_JANELA_DEPOIS = 180; // ... ou até 6 meses depois (replanejada)

function plhCampo_(p) { var m = plhPlataforma_(p).match(/^[A-Z]+/); return m ? m[0] : ''; }

function plhProcPlancor_(tarefa, obs) {
  var s = (tarefa + ' ' + obs).toUpperCase();
  var m = s.match(/PI-?(PLEM|ANM)(-PE)?\b/);
  if (m) return 'PI-' + m[1] + (m[2] || '');
  m = s.match(/PID([RF])-?(\d+)/);
  if (m) return 'PID' + m[1] + '-' + m[2];
  if (plhTxt_(obs)) return ''; // descrição própria sem procedimento (ex.: monitoramento de válvulas da ANM)
  if (/PLEM|PLAEM/.test(s)) return 'PI-PLEM';
  if (/ANM/.test(s)) return 'PI-ANM';
  return '';
}

function plhLigarPlancor_(plancor, fluxo, rts, jaquetas, nomesObj, plan, hoje) {
  var res = { carregado: plancor.info.ok && plancor.rows.length > 0, atualizado: plancor.info.atualizado || '',
              gid: plancor.info.gid, linhas: plancor.rows.length, suas: 0, ok: 0, prov: 0, nao: 0, na: 0,
              rtsNoPlancor: 0, jaqNoPlancor: 0, semDemanda: [] };
  var demandas = fluxo.filter(function (x) { return x.tipo === 'demanda'; });
  var dBR = function (d) { return d ? Utilities.formatDate(d, ssTimeZone(), 'dd/MM/yyyy') : ''; };
  var etapa = function (estado, info) { return { estado: estado, data: '', info: info }; };

  // campos que são nossos (para ignorar as linhas de outros ativos)
  var campos = {};
  demandas.forEach(function (x) { (x.locais || []).concat([x.plat]).forEach(function (p) { var c = plhCampo_(p); if (c) campos[c] = true; }); });
  jaquetas.forEach(function (j) { var c = plhCampo_(j.plat); if (c) campos[c] = true; });
  rts.forEach(function (r) { var c = plhCampo_(r.plat); if (c) campos[c] = true; });

  var nomes = Object.keys(nomesObj).sort(function (a, b) { return b.length - a.length; });
  var linhas = [];
  plancor.rows.forEach(function (r) {
    var campo = plhCampo_(r.uep);
    if (!campos[campo]) return;
    res.suas++;
    var dutoN = plhNorm_(r.duto.replace(/\s*-\s*lado\s+.*$/i, ''));
    var objeto = nomesObj[dutoN] || '';
    if (!objeto) {
      var txt = ' ' + plhNorm_(r.duto + ' ' + r.equip + ' ' + r.obs) + ' ';
      for (var i = 0; i < nomes.length; i++) if (txt.indexOf(nomes[i]) >= 0) { objeto = nomesObj[nomes[i]]; break; }
    }
    var lado = (r.duto + ' ' + r.obs).match(/lado\s+([A-Z]+\s*-?\s*\d+)/i);
    var anm = (r.duto + ' ' + r.equip).match(/\bANM\s+([A-Z]{2,3})\s*-?\s*0*(\d+)/i);
    var rt = r.obs.match(/\bRT\s*[-:nº°.]*\s*(\d{4,6})\b/i);
    linhas.push({ r: r, campo: campo, objeto: objeto, lado: lado ? plhPlataforma_(lado[1]) : '',
                  proc: plhProcPlancor_(r.tarefa, r.obs), rt: rt ? rt[1] : '',
                  anm: anm ? new RegExp('\\b' + anm[1].toUpperCase() + '-?0*' + anm[2] + '\\b') : null,
                  jaqueta: /jaqueta/i.test(r.tarefa + ' ' + r.duto), usada: false });
  });

  // RTs: número da RT nas observações
  var rtPor = {};
  linhas.forEach(function (l) { if (l.rt) rtPor[l.rt] = l; });
  rts.forEach(function (r) {
    var l = rtPor[plhTxt_(r.rt).replace(/\.0$/, '')];
    if (l) { l.usada = true; r.plancor = 'insumo ' + (l.r.ref || 'linha ' + l.r.linha) + (l.r.ini ? ', ' + dBR(l.r.ini) : ''); res.rtsNoPlancor++; }
  });

  // Jaquetas: "Jaqueta PAT-3" + "N2 SUBMERSA" nas observações (linhas sem nível/zona, como ACFM ou
  // medição de potencial, não identificam o nível). Cada nível fica com a linha mais próxima da próxima inspeção.
  var jaqLinhas = [];
  linhas.forEach(function (l) {
    if (!l.jaqueta) return;
    var m = l.r.duto.match(/jaqueta\s+([A-Z]+\s*-?\s*\d+)/i) || l.r.obs.match(/jaqueta\s+(?:da\s+)?([A-Z]+\s*-?\s*\d+)/i);
    var o = l.r.obs.toUpperCase();
    var zn = o.match(/\bN([123])\s+(SUBMERSA|EMERSA)/);
    l.usada = true; // linha de jaqueta nunca vale para duto
    if (!m || !zn) return;
    jaqLinhas.push({ l: l, plat: plhPlataforma_(m[1]), nivel: 'n' + zn[1] + ' ' + zn[2].toLowerCase() });
  });
  jaquetas.forEach(function (j) {
    var alvo = j.prox ? new Date(j.prox + 'T00:00:00').getTime() : hoje.getTime();
    var melhor = null;
    jaqLinhas.forEach(function (c) {
      if (c.usada || c.plat !== j.plat || plhNorm_(j.nivel) !== c.nivel) return;
      var d = c.l.r.ini ? Math.abs(c.l.r.ini.getTime() - alvo) : 1e15;
      if (!melhor || d < melhor.d) melhor = { c: c, d: d };
    });
    if (melhor) {
      melhor.c.usada = true; res.jaqNoPlancor++;
      j.plancor = 'insumo ' + (melhor.c.l.r.ref || 'linha ' + melhor.c.l.r.linha) + (melhor.c.l.r.ini ? ', ' + dBR(melhor.c.l.r.ini) : '');
    } else if (res.carregado && /submersa/i.test(j.nivel) && j.plano && j.dias !== null && j.dias <= 365) {
      j.foraPlancor = true; // submersa (precisa de embarcação) vencendo em até 12 meses e fora do PLANCOR
    }
  });

  var procOk = function (dProc, lProc, exato) {
    if (!lProc) return false;
    var a = plhNorm_(dProc), b = plhNorm_(lProc);
    return exato ? a === b : (a === b || a.indexOf(b + '-') === 0);
  };
  var camposDem = function (x) {
    var c = {};
    (x.locais || []).concat([x.plat]).forEach(function (p) { var k = plhCampo_(p); if (k) c[k] = true; });
    return c;
  };
  var distancia = function (x, l) {
    if (!l.r.ini || !x.prazo) return null;
    var d = (l.r.ini.getTime() - new Date(x.prazo + 'T00:00:00').getTime()) / 86400000;
    if (d < -PLH_PLANCOR_JANELA_ANTES || d > PLH_PLANCOR_JANELA_DEPOIS) return null;
    return Math.abs(d);
  };
  // pareamento guloso: menor distância de data primeiro, cada linha serve uma demanda só
  var parear = function (lista, pode) {
    var pares = [];
    demandas.forEach(function (x, ix) {
      if (x.plancor) return;
      lista.forEach(function (l, il) {
        if (l.usada || !pode(x, l)) return;
        var d = distancia(x, l);
        if (d !== null) pares.push({ x: x, l: l, d: d });
      });
    });
    pares.sort(function (a, b) { return a.d - b.d; });
    pares.forEach(function (p) {
      if (p.x.plancor || p.l.usada) return;
      p.l.usada = true; p.x._plancorLinha = p.l; p.x.plancor = lista === identificadas ? 'ok' : 'prov';
    });
  };

  demandas.forEach(function (x) { if (PLH_PLANCOR_NA.test(x.proc)) x.plancor = 'na'; });
  if (res.carregado) {
    var identificadas = linhas.filter(function (l) { return l.objeto && !l.jaqueta && !l.rt && l.proc; });
    var genericas = linhas.filter(function (l) { return !l.objeto && !l.anm && !l.jaqueta && !l.rt && l.proc; });
    var ladoOk = function (x, l) { return !l.lado || l.lado === x.plat || (x.locais || []).indexOf(l.lado) >= 0; };
    [true, false].forEach(function (exato) {
      parear(identificadas, function (x, l) {
        return plhNorm_(l.objeto) === plhNorm_(x.objeto) && procOk(x.proc, l.proc, exato) && ladoOk(x, l);
      });
    });
    linhas.forEach(function (l) {
      if (!l.anm || !l.proc) return;
      demandas.forEach(function (x) {
        if (x.plancor || !procOk(x.proc, l.proc, true) || !l.anm.test(String(x.objeto).toUpperCase())) return;
        if (distancia(x, l) === null) return;
        x.plancor = 'ok'; x._plancorLinha = l; l.usada = true;
      });
    });
    [true, false].forEach(function (exato) {
      parear(genericas, function (x, l) { return camposDem(x)[l.campo] && procOk(x.proc, l.proc, exato); });
    });
    // no PLANCOR, mas sem demanda no GIDES perto da data (procura em todo o planejamento, não só no horizonte do fluxo)
    var limite = new Date(hoje.getTime() + PLH_FLUXO_HORIZONTE * 86400000);
    identificadas.forEach(function (l) {
      if (l.usada || !l.r.ini || l.r.ini > limite) return;
      var temPlan = plan.some(function (p) {
        return plhNorm_(p.objeto) === plhNorm_(l.objeto) && procOk(p.proc, l.proc, false) && distancia({ prazo: p.prox }, l) !== null;
      });
      if (!temPlan) res.semDemanda.push({ objeto: l.objeto, proc: l.proc, lado: l.lado, ini: plhIso_(l.r.ini), ref: l.r.ref, obs: plhCorta_(l.r.obs, 120) });
    });
  }

  demandas.forEach(function (x) {
    var l = x._plancorLinha; delete x._plancorLinha;
    if (!res.carregado && x.plancor !== 'na') { x.plancor = ''; x.etapas[1] = etapa('sem', 'Aba SRC_PLANCOR ainda sem dados'); return; }
    if (!x.plancor) x.plancor = 'nao';
    res[x.plancor]++;
    var onde = l ? ' · insumo ' + (l.r.ref || 'linha ' + l.r.linha) + (l.r.ini ? ' · início ' + dBR(l.r.ini) : '') + ' · ' + l.r.tarefa : '';
    if (x.plancor === 'ok') x.etapas[1] = etapa('ok', 'No PLANCOR' + onde);
    else if (x.plancor === 'prov') x.etapas[1] = etapa('prov', 'Provável: linha genérica do campo ' + l.campo + ' com ' + l.proc + onde);
    else if (x.plancor === 'na') x.etapas[1] = etapa('na', 'Inspeção emersa: não vai para o PLANCOR');
    else {
      x.etapas[1] = etapa('alerta', 'Não encontrada no PLANCOR');
      x.alertaPlancor = true;
    }
  });
  return res;
}


/* ============================================================================
 *  ICPI SSUB: lista de ações do ano, previsões (otimista/realista/pessimista)
 *  e projeção do indicador. Regra: objeto em dia = todas as inspeções com
 *  relatório registrado até o vencimento. ICPI = objetos em dia / objetos com plano.
 *  As previsões e notas ficam na aba PLH_ICPI_Acoes (gravada pelo FlowTask).
 * ==========================================================================*/
var PLH_ICPI_META = 62.3;
var PLH_ABA_ACOES = 'PLH_ICPI_Acoes';
var PLH_ACOES_COLS = ['Chave', 'Ano', 'Ação', 'Objeto', 'Procedimento', 'Lado', 'Vencimento', 'Gerência', 'Recurso',
  'Início plan', 'Fim proj (realista)', '% Concl.', 'Status', 'Observações', 'Justificativa',
  'Prev. Otimista', 'Prev. Realista', 'Prev. Pessimista', 'Atualizado em', 'Atualizado por'];
var PLH_MESES = ['janeiro', 'fevereiro', 'março', 'abril', 'maio', 'junho', 'julho', 'agosto', 'setembro', 'outubro', 'novembro', 'dezembro'];

function plhChaveAcao_(objeto, proc, lado, ano) {
  return [plhNorm_(objeto), plhNorm_(proc), lado || '', String(ano)].join('|');
}

function plhAbaAcoes_(ss) {
  var sh = ss.getSheetByName(PLH_ABA_ACOES);
  if (!sh) {
    sh = ss.insertSheet(PLH_ABA_ACOES);
    sh.getRange(1, 1, 1, PLH_ACOES_COLS.length).setValues([PLH_ACOES_COLS])
      .setFontWeight('bold').setBackground('#2c3e50').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    sh.setTabColor('#8e24aa');
  }
  return sh;
}

function plhLerAcoes_(sh) {
  var vals = sh.getDataRange().getValues();
  var cab = vals[0] || [];
  var ix = {};
  PLH_ACOES_COLS.forEach(function (c) { ix[c] = cab.indexOf(c); });
  var porChave = {}, lista = [];
  for (var i = 1; i < vals.length; i++) {
    var r = vals[i], g = function (c) { return ix[c] < 0 ? '' : r[ix[c]]; };
    var chave = plhTxt_(g('Chave'));
    if (!chave) continue;
    var pct = g('% Concl.');
    var a = {
      chave: chave, linha: i + 1, ano: Number(g('Ano')) || null, acao: plhTxt_(g('Ação')), objeto: plhTxt_(g('Objeto')),
      proc: plhTxt_(g('Procedimento')), lado: plhTxt_(g('Lado')), venc: plhIso_(plhData_(g('Vencimento'))),
      gerencia: plhTxt_(g('Gerência')), recurso: plhTxt_(g('Recurso')), inicio: plhIso_(plhData_(g('Início plan'))),
      fim: plhIso_(plhData_(g('Fim proj (realista)'))),
      pct: (pct === '' || pct === null) ? null
        : typeof pct === 'number' ? (pct > 0 && pct <= 1 ? Math.round(pct * 100) : pct)  // 0,9 (formato %) ou 90
        : plhNum_(String(pct).replace('%', '')),
      status: plhTxt_(g('Status')), obs: plhTxt_(g('Observações')), just: plhTxt_(g('Justificativa')),
      otim: plhTxt_(g('Prev. Otimista')), real: plhTxt_(g('Prev. Realista')), pess: plhTxt_(g('Prev. Pessimista')),
      atualizado: plhIso_(plhData_(g('Atualizado em')))
    };
    porChave[chave] = a; lista.push(a);
  }
  return { porChave: porChave, lista: lista };
}

/** Sugestão de gerência/recurso pelo procedimento (o usuário pode trocar). */
function plhSugestaoRecurso_(proc) {
  return /^PID[RF]-4/i.test(proc) ? { gerencia: 'ENGI', recurso: 'Inspetor' } : { gerencia: 'SUB', recurso: 'Embarcação' };
}

function plhMontarIcpi_(ss, icpiRows, fluxo, hoje) {
  var ano = hoje.getFullYear();
  var fimAno = new Date(ano, 11, 31);
  var sh = plhAbaAcoes_(ss);
  var guardadas = plhLerAcoes_(sh);
  var acoes = [], vistas = {}, novas = [];

  var nomeAcao = function (objeto, proc, uep) { return objeto + ' - ' + proc + (uep ? ' - Lado ' + uep : ''); };
  var mesmoLado = function (x, lado) { return !lado || (x.locais || []).indexOf(lado) >= 0; };

  // inspeções do ano ainda pendentes no ICPI (inclui vencidas)
  icpiRows.forEach(function (r) {
    if (!r.venc || r.venc > fimAno) return;
    var lado = plhPlataforma_(r.uep);
    var chave = plhChaveAcao_(r.objeto, r.proc, lado, r.venc.getFullYear());
    if (vistas[chave]) return;
    vistas[chave] = true;
    var g = guardadas.porChave[chave];
    if (!g) novas.push([chave, r.venc.getFullYear(), nomeAcao(r.objeto, r.proc, r.uep), r.objeto, r.proc, r.uep, r.venc]);
    acoes.push({ chave: chave, ano: r.venc.getFullYear(), acao: nomeAcao(r.objeto, r.proc, r.uep), objeto: r.objeto, proc: r.proc,
                 lado: lado, ladoTxt: r.uep, venc: plhIso_(r.venc), os: r.os, statusOS: r.statusOS, pendenteICPI: true, g: g || null });
  });
  // ações guardadas do ano que já saíram do ICPI (concluídas ou replanejadas)
  guardadas.lista.forEach(function (g) {
    if (vistas[g.chave] || g.ano !== ano) return;
    vistas[g.chave] = true;
    acoes.push({ chave: g.chave, ano: g.ano, acao: g.acao, objeto: g.objeto, proc: g.proc, lado: plhPlataforma_(g.lado),
                 ladoTxt: g.lado, venc: g.venc, os: '', statusOS: '', pendenteICPI: false, g: g });
  });
  if (novas.length) {
    var ini = Math.max(sh.getLastRow(), 1) + 1;
    sh.getRange(ini, 1, novas.length, 7).setValues(novas);
  }

  // situação automática pelo fluxo (execução e relatório no ISUB / OS)
  acoes.forEach(function (a) {
    var execs = fluxo.filter(function (x) {
      return x.tipo === 'execucao' && plhNorm_(x.objeto) === plhNorm_(a.objeto) && plhNorm_(x.proc) === plhNorm_(a.proc) &&
        mesmoLado(x, a.lado) && (x.dataRef || '').slice(0, 4) === String(a.ano);
    }).sort(function (p, q) { return p.dataRef < q.dataRef ? 1 : -1; });
    var dem = fluxo.filter(function (x) {
      return x.tipo === 'demanda' && plhNorm_(x.objeto) === plhNorm_(a.objeto) && plhNorm_(x.proc) === plhNorm_(a.proc) &&
        x.prazo === a.venc && mesmoLado(x, a.lado);
    })[0];
    var ex = execs[0];
    var cad = ex && ex.etapas[4] && ex.etapas[4].estado === 'ok' ? ex.etapas[4].data : '';
    if (ex && cad && (!a.pendenteICPI || cad >= (a.ano + '-01-01'))) {
      a.auto = { situacao: 'concluida', etapa: 'Relatório no ISUB em ' + plhIsoBR_(cad), pct: 100, status: 'Concluido', data: cad };
    } else if (ex) {
      a.auto = { situacao: 'executada', etapa: 'Executada em ' + plhIsoBR_(ex.dataRef) + ', aguardando relatório no ISUB', pct: 90, status: 'Em andamento', data: ex.dataRef };
    } else if (!a.pendenteICPI) {
      a.auto = { situacao: 'saiu', etapa: 'Fora do ICPI do ano (já atendida ou replanejada)', pct: null, status: '', data: '' };
    } else {
      var etapa = dem ? dem.situacao : (a.os ? 'OS ' + a.os + (a.statusOS ? ' · ' + a.statusOS : '') : 'Sem OS');
      a.auto = { situacao: dem && dem.etapaAtual === 3 ? 'barco' : 'pendente', etapa: etapa, pct: 0, status: 'Em andamento', data: '' };
    }
    var sug = plhSugestaoRecurso_(a.proc);
    var g = a.g || {};
    a.gerencia = g.gerencia || sug.gerencia;
    a.recurso = g.recurso || sug.recurso;
    a.inicio = g.inicio || ''; a.fim = g.fim || '';
    a.pct = (g.pct !== null && g.pct !== undefined && g.pct !== '') ? g.pct : a.auto.pct;
    a.status = a.auto.situacao === 'concluida' ? 'Concluido' : (g.status || a.auto.status);
    a.obs = g.obs || ''; a.just = g.just || '';
    a.otim = g.otim || ''; a.real = g.real || ''; a.pess = g.pess || '';
    a.atualizado = g.atualizado || '';
    a.concluida = a.auto.situacao === 'concluida' || plhNorm_(a.status).indexOf('conclu') === 0;
    var d = a.venc ? plhDias_(new Date(a.venc + 'T00:00:00'), hoje) : null;
    a.dias = d;
    a.prazo = d === null ? '' : d < 0 ? 'Venc.' : d <= 90 ? '<90d' : '>90d';
    delete a.g;
  });
  acoes.sort(function (p, q) { return (p.venc || '9') < (q.venc || '9') ? -1 : (p.venc || '9') > (q.venc || '9') ? 1 : (p.acao < q.acao ? -1 : 1); });

  // objetos do indicador e as inspeções de cada um (para a projeção)
  var objs = {};
  icpiRows.forEach(function (r) {
    var o = objs[r.objeto] = objs[r.objeto] || { objeto: r.objeto, grupo: r.grupo, itens: [] };
    var lado = plhPlataforma_(r.uep);
    o.itens.push({ venc: plhIso_(r.venc), chave: r.venc ? plhChaveAcao_(r.objeto, r.proc, lado, r.venc.getFullYear()) : '' });
  });
  return {
    meta: PLH_ICPI_META, ano: ano, hoje: plhIso_(hoje), meses: PLH_MESES,
    objetos: Object.keys(objs).map(function (k) { return objs[k]; }), acoes: acoes,
    abaGid: sh.getSheetId()
  };
}

/** Grava as previsões e notas de uma ação (chamado pelo formulário da tela ICPI). */
function salvarIcpiAcao(a) {
  var lock = LockService.getDocumentLock();
  lock.waitLock(20000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = plhAbaAcoes_(ss);
    var vals = sh.getDataRange().getValues();
    var cab = vals[0];
    var col = function (c) { return cab.indexOf(c) + 1; };
    var linha = 0;
    for (var i = 1; i < vals.length; i++) if (plhTxt_(vals[i][col('Chave') - 1]) === a.chave) { linha = i + 1; break; }
    if (!linha) {
      linha = Math.max(sh.getLastRow(), 1) + 1;
      sh.getRange(linha, 1, 1, 7).setValues([[a.chave, a.ano, a.acao, a.objeto, a.proc, a.ladoTxt || '',
        a.venc ? new Date(a.venc + 'T12:00:00') : '']]);
    }
    var dataOuVazio = function (iso) { return iso ? new Date(iso + 'T12:00:00') : ''; };
    var campos = {
      'Gerência': a.gerencia || '', 'Recurso': a.recurso || '', 'Início plan': dataOuVazio(a.inicio),
      'Fim proj (realista)': dataOuVazio(a.fim), '% Concl.': (a.pct === null || a.pct === '' || a.pct === undefined) ? '' : Number(a.pct),
      'Status': a.status || '', 'Observações': a.obs || '', 'Justificativa': a.just || '',
      'Prev. Otimista': a.otim || '', 'Prev. Realista': a.real || '', 'Prev. Pessimista': a.pess || '',
      'Atualizado em': new Date(), 'Atualizado por': (Session.getActiveUser().getEmail() || '').split('@')[0]
    };
    Object.keys(campos).forEach(function (c) { if (col(c) > 0) sh.getRange(linha, col(c)).setValue(campos[c]); });
    return { success: true };
  } finally {
    lock.releaseLock();
  }
}


/* ============================================================================
 *  MATRIZ DE PLANEJAMENTO: objeto × procedimento, com a periodicidade cadastrada no GIDES
 *  comparada com a regra de referência (aba PLH_Regras, editável pelo usuário).
 *  Limite: o export de planejamento só traz componentes que têm algum planejamento;
 *  componente sem nenhum plano não aparece (só o duto inteiro sem plano é detectado, via CIDR).
 * ==========================================================================*/
var PLH_ABA_REGRAS = 'PLH_Regras';
var PLH_REGRAS_COLS = ['Tipo componente', 'Procedimento', 'Periodicidade padrão (meses)', 'Obrigatório', 'Aceita no lugar', 'Observação'];
// Ponto de partida: o mais comum no cadastro atual. O usuário corrige na aba.
var PLH_REGRAS_PADRAO = [
  ['Riser Rígido', 'PIDR-4', 36, 'Sim', '', 'Emersa (inspetor)'],
  ['Riser Rígido', 'PIDR-3', 72, 'Sim', 'PIDR-8', 'Riser submerso; PIDR-8 pode substituir'],
  ['Riser Rígido', 'PIDR-8', 72, 'Não', '', 'Diferenciada'],
  ['Trecho Rígido', 'PIDR-1', '', 'Sim', 'PIDR-8', 'Leito marinho; hoje todos os trechos estão com PIDR-8'],
  ['Trecho Rígido', 'PIDR-8', 72, 'Não', '', 'Diferenciada'],
  ['Riser Flexível', 'PIDF-4', 48, 'Sim', '', 'Emersa (inspetor)'],
  ['Riser Flexível', 'PIDF-3', 72, 'Sim', 'PIDF-1', ''],
  ['Riser Flexível', 'PIDF-1', 120, 'Não', '', ''],
  ['Flow Flexível', 'PIDF-1', 120, 'Sim', '', ''],
  ['SISTEMA ANM', 'PI-ANM', 36, 'Sim', '', ''],
  ['SISTEMA ANM', 'PI-ANM-PE', 36, 'Sim', '', ''],
  ['PLEM', 'PI-PLEM', 72, 'Sim', '', ''],
  ['PLEM', 'PI-PLEM-PE', 72, 'Sim', '', ''],
  ['J-TUBE', 'PIDR-8', 72, 'Não', '', '']
];
var PLH_MATRIZ_GRUPOS = [
  { grupo: 'Rígido', procs: ['PIDR-0', 'PIDR-1', 'PIDR-3', 'PIDR-4', 'PIDR-8'] },
  { grupo: 'Flexível', procs: ['PIDF-1', 'PIDF-3', 'PIDF-4'] },
  { grupo: 'Equipamentos', procs: ['PI-ANM', 'PI-ANM-PE', 'PI-PLEM', 'PI-PLEM-PE'] }
];

function plhLerRegras_(ss) {
  var sh = ss.getSheetByName(PLH_ABA_REGRAS), criada = false;
  if (!sh) {
    try {
      sh = ss.insertSheet(PLH_ABA_REGRAS);
      sh.getRange(1, 1, 1, PLH_REGRAS_COLS.length).setValues([PLH_REGRAS_COLS])
        .setFontWeight('bold').setBackground('#2c3e50').setFontColor('#ffffff');
      sh.getRange(2, 1, PLH_REGRAS_PADRAO.length, PLH_REGRAS_COLS.length).setValues(PLH_REGRAS_PADRAO);
      sh.setFrozenRows(1);
      sh.setTabColor('#00a085');
      criada = true;
    } catch (e) {
      return { gid: null, criada: false, lista: PLH_REGRAS_PADRAO.map(plhRegraDe_), erro: 'Não foi possível criar a aba PLH_Regras: ' + e.message };
    }
  }
  var vals = sh.getDataRange().getValues();
  var cab = (vals[0] || []).map(plhNorm_);
  var ix = function (n) { return cab.indexOf(plhNorm_(n)); };
  var c = PLH_REGRAS_COLS.map(ix);
  var lista = [];
  for (var i = 1; i < vals.length; i++) {
    var r = vals[i];
    var linha = c.map(function (k) { return k < 0 ? '' : r[k]; });
    if (!plhTxt_(linha[0]) || !plhTxt_(linha[1])) continue;
    lista.push(plhRegraDe_(linha));
  }
  return { gid: sh.getSheetId(), criada: criada, lista: lista };
}
function plhRegraDe_(l) {
  return { tipo: plhTxt_(l[0]), proc: plhTxt_(l[1]).toUpperCase(), padrao: plhNum_(l[2]),
           obrig: /^s/i.test(plhTxt_(l[3])), alt: plhTxt_(l[4]).toUpperCase().split(/[\s,;\/]+/).filter(String),
           obs: plhTxt_(l[5]) };
}

/** "72 Meses" → {meses:72}; "3160 Dias" → {meses:103.8, dias:3160} */
function plhPeriodo_(txt) {
  var t = plhNorm_(txt), n = plhNum_((t.match(/[\d.,]+/) || [''])[0]);
  if (n === null) return { meses: null, dias: null, txt: plhTxt_(txt) };
  if (/dia/.test(t)) return { meses: Math.round(n / 30.4375 * 10) / 10, dias: n, txt: plhTxt_(txt) };
  if (/ano/.test(t)) return { meses: n * 12, dias: null, txt: plhTxt_(txt) };
  return { meses: n, dias: null, txt: plhTxt_(txt) };
}

function plhMontarMatriz_(ss, plan, objetos, icpiRows, compCIDR) {
  var regras = plhLerRegras_(ss);
  var refPer = plhLerPeriodicidade_(ss);
  var objPorId = {};
  objetos.forEach(function (o) { objPorId[o.id] = o; });
  var icpiPor = {};
  (icpiRows || []).forEach(function (r) { var k = plhNorm_(r.objeto) + '|' + plhNorm_(r.proc); (icpiPor[k] = icpiPor[k] || []).push(r); });
  var ladoICPI = function (p) {
    if (!p.prox) return '';
    var alvo = new Date(p.prox + 'T00:00:00').getTime();
    var u = {};
    (icpiPor[plhNorm_(p.objeto) + '|' + plhNorm_(p.proc)] || []).forEach(function (r) {
      if (r.venc && Math.abs(r.venc.getTime() - alvo) <= 3 * 86400000 && r.uep) u[plhPlataforma_(r.uep)] = true;
    });
    var k = Object.keys(u);
    return k.length === 1 ? k[0] : '';
  };
  var regraPor = {}, exigidas = {};
  regras.lista.forEach(function (g) {
    regraPor[plhNorm_(g.tipo) + '|' + g.proc] = g;
    if (g.obrig) (exigidas[plhNorm_(g.tipo)] = exigidas[plhNorm_(g.tipo)] || []).push(g);
  });

  var linhas = {}, ordem = [];
  var linhaDe = function (nome, id, plat, situacao, classe) {
    var k = plhNorm_(nome);
    if (!linhas[k]) {
      linhas[k] = { objId: id || '', objeto: nome, plat: plat || '', situacao: situacao || '', classe: classe || '',
                    cel: {}, comps: {}, nPlanos: 0 };
      ordem.push(k);
    }
    var l = linhas[k];
    if (!l.objId && id) l.objId = id;
    if (!l.plat && plat) l.plat = plat;
    if (!l.situacao && situacao) l.situacao = situacao;
    return l;
  };
  objetos.forEach(function (o) {
    if (o.cadastro !== 'CIDR') return; // ANM/PLEM fora do CIDR entram pelo próprio planejamento do duto
    linhaDe(o.nome, o.id, o.plat, o.situacao, o.classe);
  });

  var procsVistos = {};
  plan.forEach(function (p) {
    if (!p.objeto || !p.proc) return;
    var l = linhaDe(p.objeto, p.objId, p.plat, p.situacao, '');
    l.situacao = p.situacao || l.situacao; // situação do GIDES (mais atual que o CIDR)
    l.nPlanos++;
    var proc = p.proc.toUpperCase();
    procsVistos[proc] = true;
    var per = plhPeriodo_(p.per);
    var regra = regraPor[plhNorm_(p.tipoComp) + '|' + proc];
    var ref = plhRefPeriodicidade_(refPer, proc, objPorId[p.objId]);
    var item = {
      proc: proc, comp: p.comp, compId: p.compId, tipo: p.tipoComp, planej: p.planej, per: per.txt, meses: per.meses,
      emDias: per.dias !== null, vale: p.vale, ult: p.ult, prox: p.prox, dias: p.dias, status: p.status,
      padrao: ref ? ref.meses : null, padraoFonte: ref ? ref.fonte : '',
      semRegra: !regra
    };
    // Referência = aba PLH_Periodicidade (procedimento × fluido). Fora = periodicidade MAIOR que a referência
    // (inspeção mais espaçada que a norma); menor é mais conservador e só fica registrado. Sem referência, não compara.
    item.foraPadrao = !!(ref && per.meses !== null && per.meses > ref.meses + 0.5);
    item.abaixoRef = !!(ref && per.meses !== null && per.meses < ref.meses - 0.5);
    (l.cel[proc] = l.cel[proc] || { itens: [] }).itens.push(item);
    // flexível: o mesmo componente aparece uma vez por ponta (Local Instalação / Tramo diferentes)
    var chaveC = (p.compId || p.comp) + (p.local ? '|' + p.local : '');
    var cc = l.comps[chaveC] = l.comps[chaveC] ||
      { chave: chaveC, comp: p.tramo ? p.tipoComp + ' [' + p.tramo + ']' : p.comp, compId: p.compId, tipo: p.tipoComp, ext: p.ext,
        local: p.local, procs: {}, itens: [], ladosICPI: {} };
    cc.procs[proc] = true;
    cc.itens.push(item);
    var lado = ladoICPI(p);
    if (lado) cc.ladosICPI[lado] = true;
  });

  // Procedimento exigido que falta no duto: a regra vale para o duto (não para cada componente),
  // porque o GIDES divide o riser em componente emerso (PIDR-4) e submerso (PIDR-3/PIDR-8).
  ordem.forEach(function (k) {
    var l = linhas[k];
    var tipos = {}, procsObj = {};
    Object.keys(l.comps).forEach(function (ck) {
      var c = l.comps[ck];
      tipos[plhNorm_(c.tipo)] = c.tipo;
      Object.keys(c.procs).forEach(function (p) { procsObj[p] = true; });
    });
    Object.keys(tipos).forEach(function (t) {
      (exigidas[t] || []).forEach(function (g) {
        if (procsObj[g.proc] || g.alt.some(function (a) { return procsObj[a]; })) return;
        procsVistos[g.proc] = true;
        var cel = l.cel[g.proc] = l.cel[g.proc] || { itens: [] };
        (cel.faltam = cel.faltam || []).push({ tipo: tipos[t], alt: g.alt });
      });
    });
    // Lados do riser: componentes emersos (PIDR-4/PIDF-4) × submersos (PIDR-3/PIDR-8/PIDF-3/PIDF-1)
    ['Rígido', 'Flexível'].forEach(function (mat) {
      var em = mat === 'Rígido' ? ['PIDR-4'] : ['PIDF-4'];
      var sub = mat === 'Rígido' ? ['PIDR-3', 'PIDR-8'] : ['PIDF-3', 'PIDF-1'];
      var nEm = 0, nSub = 0;
      Object.keys(l.comps).forEach(function (ck) {
        var c = l.comps[ck];
        if (plhNorm_(c.tipo) !== plhNorm_('Riser ' + mat)) return;
        if (em.some(function (p) { return c.procs[p]; })) nEm++;
        if (sub.some(function (p) { return c.procs[p]; })) nSub++;
      });
      if (nEm && nSub && nEm !== nSub) {
        var alvo = nEm < nSub ? em[0] : sub[0];
        if (!l.cel[alvo] && nEm < nSub) l.cel[alvo] = { itens: [] };
        (l.lados = l.lados || []).push({ mat: mat, em: nEm, sub: nSub, proc: nEm < nSub ? em[0] : null });
      }
    });
  });

  var comps = plhLerComponentes_(ss);
  var porDuto = (compCIDR && compCIDR.porDuto) || {};
  ordem.forEach(function (k) { linhas[k].esq = plhEsqueleto_(linhas[k], objPorId[linhas[k].objId], comps, porDuto[linhas[k].objId] || []); });
  plhGravarComponentes_(comps, ordem.map(function (k) { return linhas[k]; }));

  var resumo = { objetos: 0, semPlano: 0, posicao: 0, lados: 0, faltando: 0, foraPadrao: 0, divergente: 0, emDias: 0, valeProximo: 0, vencidas: 0, semRegra: 0 };
  var lista = ordem.map(function (k) {
    var l = linhas[k];
    var cel = {};
    Object.keys(l.cel).forEach(function (proc) {
      var c = l.cel[proc], it = c.itens;
      var pers = {};
      it.forEach(function (i) { pers[i.meses === null ? i.per : i.meses] = true; });
      var prox = it.map(function (i) { return i.prox; }).filter(String).sort()[0] || '';
      var diasMin = it.reduce(function (m, i) { return i.dias === null ? m : (m === null || i.dias < m ? i.dias : m); }, null);
      cel[proc] = {
        n: it.length, itens: it, faltam: c.faltam || [],
        per: it.length ? Object.keys(pers).map(function (x) { return isNaN(x) ? x : Number(x); }) : [],
        prox: prox, dias: diasMin,
        divergente: Object.keys(pers).length > 1,
        foraPadrao: it.some(function (i) { return i.foraPadrao; }),
        emDias: it.some(function (i) { return i.emDias; }),
        valeProximo: it.some(function (i) { return /pr[oó]xim/i.test(i.vale); }),
        semRegra: it.length > 0 && it.every(function (i) { return i.semRegra; }),
        lados: (l.lados || []).filter(function (x) { return x.proc === proc; })[0] || null
      };
    });
    var flags = { semPlano: l.nPlanos === 0, posicao: l.esq ? l.esq.alertas.length : 0, lados: (l.lados || []).length, faltando: 0, foraPadrao: 0, divergente: 0, emDias: 0, valeProximo: 0, vencidas: 0, semRegra: 0 };
    Object.keys(cel).forEach(function (p) {
      var c = cel[p];
      if (c.faltam.length) flags.faltando++;
      ['foraPadrao', 'divergente', 'emDias', 'valeProximo', 'semRegra'].forEach(function (f) { if (c[f]) flags[f]++; });
      if (c.dias !== null && c.dias < 0) flags.vencidas++;
    });
    resumo.objetos++;
    if (flags.semPlano) resumo.semPlano++;
    if (flags.posicao) resumo.posicao++;
    ['lados', 'faltando', 'foraPadrao', 'divergente', 'emDias', 'valeProximo', 'vencidas', 'semRegra'].forEach(function (f) { resumo[f] += flags[f]; });
    return { objId: l.objId, objeto: l.objeto, plat: l.plat, situacao: l.situacao, classe: l.classe,
             nPlanos: l.nPlanos, nComps: Object.keys(l.comps).length, cel: cel, flags: flags, lados: l.lados || [], esq: l.esq };
  });
  lista.sort(function (a, b) { return a.plat === b.plat ? (a.objeto < b.objeto ? -1 : 1) : (a.plat < b.plat ? -1 : 1); });

  var colunas = [];
  var usados = {};
  PLH_MATRIZ_GRUPOS.forEach(function (g) {
    g.procs.forEach(function (p) { if (procsVistos[p]) { colunas.push({ proc: p, grupo: g.grupo }); usados[p] = true; } });
  });
  Object.keys(procsVistos).sort().forEach(function (p) { if (!usados[p]) colunas.push({ proc: p, grupo: 'Outros' }); });

  return { colunas: colunas, linhas: lista, resumo: resumo, refGid: refPer.gid, refErro: refPer.erro || '', refLinhas: refPer.lista.length, compGid: comps.gid, compErro: comps.erro || '', limiteEmerso: comps.limite,
           regras: regras.lista, regrasGid: regras.gid, regrasCriada: regras.criada, regrasErro: regras.erro || '' };
}


/* ============================================================================
 *  COMPONENTES: posição de cada componente no duto (esqueleto)
 *  Rígido:   lado A emerso · lado A submerso · leito · lado B submerso · lado B emerso
 *            emerso = riser curto (≤ limite, padrão 15 m) → PIDR-4
 *            submerso = riser longo → PIDR-3 (PIDR-8 substitui)
 *            leito = trecho → PIDR-1 (PIDR-8 substitui)
 *  Flexível: lado A riser · leito · lado B riser (riser: PIDF-4 + PIDF-3/PIDF-1)
 *  Lado: UEP do ICPI (pela demanda do componente); senão, a correção da aba PLH_Componentes.
 * ==========================================================================*/
var PLH_ABA_COMP = 'PLH_Componentes';
var PLH_COMP_COLS = ['Componente BDIEP', 'Duto', 'Componente', 'Tipo', 'Extensão (m)', 'Procedimentos',
                     'Posição (automática)', 'Lado (automático)', 'Posição (corrigir)', 'Lado (corrigir)', 'Motivo / observação'];
var PLH_COMP_LIMITE_PADRAO = 15;
var PLH_POS_ROTULO = { emerso: 'Riser emerso', submerso: 'Riser submerso', leito: 'Flowline', riser: 'Riser', equip: 'Equipamento', terra: 'Trecho terrestre', outro: 'Outro componente' };
var PLH_POS_ESPERADO = {
  rigido: { emerso: { base: ['PIDR-4'], subst: [] }, submerso: { base: ['PIDR-3'], subst: ['PIDR-8'] }, leito: { base: ['PIDR-1'], subst: ['PIDR-8'] } },
  flexivel: { riser: { base: ['PIDF-4', 'PIDF-3'], subst: ['PIDF-1'] }, leito: { base: ['PIDF-1'], subst: [] } }
};

function plhLerComponentes_(ss) {
  var res = { sh: null, gid: null, porId: {}, limite: PLH_COMP_LIMITE_PADRAO, erro: '', linhas: 0 };
  var sh = ss.getSheetByName(PLH_ABA_COMP);
  if (!sh) return res; // aba opcional: o usuário preferiu não manter correções manuais
  if (false) {
    try {
      sh = ss.insertSheet(PLH_ABA_COMP);
      sh.getRange(1, 1, 1, PLH_COMP_COLS.length).setValues([PLH_COMP_COLS])
        .setFontWeight('bold').setBackground('#2c3e50').setFontColor('#ffffff');
      sh.getRange(1, 9, 1, 3).setBackground('#00a085');
      sh.setFrozenRows(1);
      sh.setTabColor('#00a085');
      sh.getRange(1, 13).setValue('Limite emerso (m)');
      sh.getRange(2, 13).setValue(PLH_COMP_LIMITE_PADRAO);
      sh.getRange(1, 13).setFontWeight('bold');
    } catch (e) { res.erro = 'Não foi possível criar a aba PLH_Componentes: ' + e.message; return res; }
  }
  res.sh = sh; res.gid = sh.getSheetId();
  var vals = sh.getDataRange().getValues();
  var cab = (vals[0] || []).map(plhNorm_);
  var ix = function (n) { return cab.indexOf(plhNorm_(n)); };
  var cLim = ix('Limite emerso (m)');
  if (cLim >= 0 && vals.length > 1) { var lim = plhNum_(vals[1][cLim]); if (lim) res.limite = lim; }
  var c = { id: ix('Componente BDIEP'), pos: ix('Posição (corrigir)'), lado: ix('Lado (corrigir)'), motivo: ix('Motivo / observação') };
  for (var i = 1; i < vals.length; i++) {
    var id = c.id < 0 ? '' : plhTxt_(vals[i][c.id]);
    if (!id) continue;
    res.linhas++;
    var g = function (k) { return c[k] < 0 ? '' : plhTxt_(vals[i][c[k]]); };
    res.porId[id] = { pos: plhPosicaoDe_(g('pos')), lado: plhPlataforma_(g('lado')), motivo: g('motivo'), existe: true };
  }
  return res;
}

/** "Riser emerso" / "submerso" / "leito" / "trecho" → chave da posição */
function plhPosicaoDe_(t) {
  var n = plhNorm_(t);
  if (!n) return '';
  if (/^emer|emerso/.test(n)) return 'emerso';
  if (/submer/.test(n)) return 'submerso';
  if (/leito|trecho|flow/.test(n)) return 'leito';
  if (/equip|anm|plem/.test(n)) return 'equip';
  if (/terra/.test(n)) return 'terra';
  if (/riser/.test(n)) return 'riser';
  return '';
}

function plhEhPlataforma_(x) { return /^P[A-Z]{1,4}\d{2}[A-Z]?$/.test(plhPlataforma_(x)); }

function plhEsqueleto_(l, obj, comps, cidr) {
  var lista = Object.keys(l.comps).map(function (k) { return l.comps[k]; });
  // componentes do cadastro do CIDR: completam os dados dos que têm plano e entram os que não têm
  var cidrPor = {};
  (cidr || []).forEach(function (c) { if (c.id) cidrPor[c.id] = c; });
  var comPlano = {};
  lista.forEach(function (c) { if (c.compId) comPlano[c.compId] = true; });
  (cidr || []).forEach(function (c, i) {
    if (c.id && comPlano[c.id]) return;
    lista.push({ chave: c.id || ('CIDR' + i + '|' + (l.objId || l.objeto)), compId: c.id, comp: c.nome, tipo: c.tipo, ext: c.ext,
                 local: c.local, ordem: c.ordem, procs: {}, itens: [], ladosICPI: {}, semPlano: true,
                 foraANP: !/^s/i.test(c.anp || ''), anp: c.anp, situacao: c.situacao });
  });
  if (!lista.length) return null;
  lista.forEach(function (c) { if (c.ordem === undefined && c.compId && cidrPor[c.compId]) c.ordem = cidrPor[c.compId].ordem; });
  var tipoTxt = lista.map(function (c) { return plhNorm_(c.tipo); }).join(' ');
  var mat = /rigido/.test(tipoTxt) ? 'rigido' : /flexivel/.test(tipoTxt) ? 'flexivel' : null;
  if (!mat && obj && obj.config) { var cf = plhNorm_(String(obj.config).replace(/\u00ad/g, '')); mat = /rigido/.test(cf) ? 'rigido' : /flexivel/.test(cf) ? 'flexivel' : null; }
  var esperado = mat ? PLH_POS_ESPERADO[mat] : {};
  // pontas do duto (CIDR): só plataforma tem riser
  var pontas = [];
  if (obj) [obj.origem, obj.destino].forEach(function (x) { if (plhEhPlataforma_(x)) { var p = plhPlataforma_(x); if (pontas.indexOf(p) < 0) pontas.push(p); } });

  // Ordem do cadastro de componentes do CIDR: segue o sentido do duto (origem → destino).
  // Riser antes do trecho = ponta de origem; depois = ponta de destino.
  var ordTrechos = lista.filter(function (c) { return /trecho|flow/.test(plhNorm_(c.tipo)) && c.ordem !== null && c.ordem !== undefined; })
    .map(function (c) { return c.ordem; });
  var ladoPelaOrdem = function (c) {
    if (!ordTrechos.length || c.ordem === null || c.ordem === undefined || !obj || !/riser/.test(plhNorm_(c.tipo))) return '';
    var ponta = c.ordem < Math.min.apply(null, ordTrechos) ? obj.origem : c.ordem > Math.max.apply(null, ordTrechos) ? obj.destino : '';
    return plhEhPlataforma_(ponta) ? plhPlataforma_(ponta) : '';
  };
  var out = [];
  lista.forEach(function (c) {
    var t = plhNorm_(c.tipo), auto = '';
    var nomesP = Object.keys(c.procs);
    if (nomesP.length && nomesP.every(function (x) { return /^(FAIXA|EXT_)/.test(x); })) auto = 'terra';
    else if (c.semPlano && !t) auto = 'indef';
    else if (c.semPlano && !/riser|trecho|flow/.test(t)) auto = /anm|plem|j-tube/.test(t) ? 'equip' : 'outro';
    else if (/trecho|flow/.test(t)) auto = 'leito';
    else if (/riser rigido/.test(t)) {
      // o procedimento decide; a extensão só desempata
      var temEm = !!c.procs['PIDR-4'], temSub = !!(c.procs['PIDR-3'] || c.procs['PIDR-8'] || c.procs['PIDR-1']);
      if (temEm && !temSub) auto = 'emerso';
      else if (temSub && !temEm) auto = 'submerso';
      else auto = (c.ext !== null && c.ext !== undefined && c.ext <= comps.limite) ? 'emerso' : 'submerso';
    } else if (/riser flexivel/.test(t)) auto = 'riser';
    else auto = 'equip';
    var ladosAuto = Object.keys(c.ladosICPI);
    var ladoICPI = ladosAuto.length === 1 ? ladosAuto[0] : '';
    var ladoOrdem = ladoPelaOrdem(c);
    var ladoAuto = ladoOrdem || c.local || ladoICPI;
    var fonte = ladoOrdem ? 'CIDR (ordem)' : c.local ? 'GIDES (local)' : ladoICPI ? 'ICPI' : '';
    var corr = comps.porId[c.chave] || {};
    var pos = corr.pos || auto, lado = corr.lado || ladoAuto;
    if (corr.lado) fonte = 'corrigido';
    if (pos === 'leito' || pos === 'equip' || pos === 'terra' || pos === 'outro' || pos === 'indef') { lado = corr.lado || ''; fonte = lado ? fonte : ''; }
    // lado sobrando: duto com só uma plataforma nas pontas
    if (!lado && (pos === 'emerso' || pos === 'submerso' || pos === 'riser') && pontas.length === 1) { lado = pontas[0]; fonte = 'ponta única'; }
    var conflito = (ladoOrdem && ladoICPI && ladoOrdem !== ladoICPI) ? ladoICPI : '';
    var esp = esperado[pos];
    var nomes = Object.keys(c.procs);
    var estado = 'ok', subst = '';
    if (esp) {
      var temBase = esp.base.filter(function (b) { return c.procs[b]; });
      var temSubst = esp.subst.filter(function (b) { return c.procs[b]; });
      if (temBase.length === esp.base.length) estado = 'ok';
      else if (temSubst.length && pos !== 'riser') { estado = 'substitui'; subst = temSubst[0] + ' no lugar da ' + esp.base[0]; }
      else if (pos === 'riser' && c.procs['PIDF-4'] && temSubst.length) { estado = 'substitui'; subst = temSubst[0] + ' no lugar da PIDF-3'; }
      else estado = 'falta';
    }
    if (c.semPlano) estado = c.foraANP ? 'foraanp' : 'semplano';
    if (c.semPlano && c.foraANP && (pos === 'emerso' || pos === 'submerso' || pos === 'riser' || pos === 'leito' || pos === 'indef')) pos = 'outro';
    if (c.semPlano && pos === 'outro' && estado === 'semplano') estado = 'acessorio';
    var dc = cidrPor[c.compId] || {};
    out.push({
      chave: c.chave, compId: c.compId, comp: c.comp, tipo: c.tipo, ext: c.ext === undefined ? null : c.ext,
      pos: pos, posAuto: auto, posManual: !!corr.pos, lado: lado, ladoAuto: ladoAuto, ladoManual: !!corr.lado,
      ladoDuvida: !c.local && !ladoOrdem && ladosAuto.length > 1 ? ladosAuto.join('/') : '',
      ladoFonte: fonte, ladoConflitoICPI: conflito, ordem: c.ordem === undefined ? null : c.ordem,
      motivo: corr.motivo || '', procs: nomes.sort(), estado: estado, subst: subst, semPlano: !!c.semPlano,
      esp: dc.esp === undefined ? null : dc.esp, dn: dc.dn === undefined ? null : dc.dn, vida: dc.vida === undefined ? null : dc.vida,
      anp: dc.anp || c.anp || '', sitComp: dc.situacao || c.situacao || '',
      esperado: esp ? esp.base.join(' + ') : '',
      itens: c.itens.map(function (i) { return { proc: i.proc, per: i.per, meses: i.meses, ult: i.ult, prox: i.prox, dias: i.dias, status: i.status, vale: i.vale, planej: i.planej,
        padrao: i.padrao, padraoFonte: i.padraoFonte, foraPadrao: i.foraPadrao, abaixoRef: !!i.abaixoRef }; })
    });
  });

  // posições esperadas
  var slots = [];
  var slot = function (pos, lado) { slots.push({ pos: pos, lado: lado || '' }); };
  if (mat === 'rigido') {
    if (pontas[0]) { slot('emerso', pontas[0]); slot('submerso', pontas[0]); }
    slot('leito', '');
    if (pontas[1]) { slot('submerso', pontas[1]); slot('emerso', pontas[1]); }
  } else if (mat === 'flexivel') {
    if (pontas[0]) slot('riser', pontas[0]);
    slot('leito', '');
    if (pontas[1]) slot('riser', pontas[1]);
  }
  var noSlot = function (c, sl) { return c.pos === sl.pos && (sl.pos === 'leito' || c.lado === sl.lado); };
  // dedução: um componente sem lado e uma única posição vazia daquele tipo → vai para ela
  ['emerso', 'submerso', 'riser'].forEach(function (pos) {
    var semLado = out.filter(function (c) { return c.pos === pos && !c.lado; });
    var vazios = slots.filter(function (sl) { return sl.pos === pos && !out.some(function (c) { return noSlot(c, sl); }); });
    if (semLado.length === 1 && vazios.length === 1) { semLado[0].lado = vazios[0].lado; semLado[0].ladoDeduzido = true; semLado[0].ladoFonte = 'deduzido'; }
    // componentes sem plano (do CIDR) completam as posições vazias quando a quantidade bate
    var semLadoSP = out.filter(function (c) { return c.pos === pos && !c.lado && c.semPlano; });
    var vaz2 = slots.filter(function (sl) { return sl.pos === pos && !out.some(function (c) { return noSlot(c, sl); }); });
    if (semLadoSP.length === 1 && vaz2.length === 1) { semLadoSP[0].lado = vaz2[0].lado; semLadoSP[0].ladoDeduzido = true; semLadoSP[0].ladoFonte = 'deduzido'; }
  });
  // componentes do CIDR sem tipo e sem plano: ocupam as posições vazias pela extensão
  // (o maior até o limite → riser emerso; até 100 m → riser submerso; acima → leito). O resto fica como "outro".
  var indef = out.filter(function (c) { return c.pos === 'indef'; }).sort(function (a, b) { return (b.ext || 0) - (a.ext || 0); });
  slots.forEach(function (sl) {
    if (out.some(function (c) { return c.pos !== 'indef' && noSlot(c, sl); })) return;
    var cabe = function (e) {
      if (sl.pos === 'riser') return true; // cabo/umbilical: o CIDR não traz extensão por ponta
      if (e === null || e === undefined) return false;
      if (sl.pos === 'emerso') return e >= 1 && e <= comps.limite;
      if (sl.pos === 'submerso') return e > comps.limite && e <= 100;
      if (sl.pos === 'leito') return e > 100;
      if (sl.pos === 'riser') return true;
      return false;
    };
    var c = indef.filter(function (x) { return x.pos === 'indef' && cabe(x.ext); })[0];
    if (!c) return;
    c.pos = sl.pos; c.lado = sl.lado; c.provavel = true; c.ladoFonte = 'deduzido';
    c.esperado = esperado[sl.pos] ? esperado[sl.pos].base.join(' + ') : '';
  });
  out.forEach(function (c) { if (c.pos === 'indef') { c.pos = 'outro'; if (c.estado === 'semplano') c.estado = 'naoclass'; } });

  var posicoes = [], alertas = [];
  slots.forEach(function (sl) {
    var cs = out.filter(function (c) { return noSlot(c, sl); }).sort(function (a, b) { return (a.ordem || 0) - (b.ordem || 0); });
    var p = { pos: sl.pos, lado: sl.lado, rotulo: PLH_POS_ROTULO[sl.pos] + (sl.lado ? ' ' + sl.lado : ''), comps: cs,
              esperado: esperado[sl.pos] ? esperado[sl.pos].base.join(' + ') : '' };
    if (!cs.length) {
      p.vazio = true;
      var semLadoMesmaPos = out.some(function (c) { return c.pos === sl.pos && !c.lado; });
      if (!(mat === 'flexivel' && sl.pos === 'leito')) alertas.push(p.rotulo + ': ' + (semLadoMesmaPos ? 'há componente desta posição sem lado identificado'
        : 'nenhum componente ' + ((cidr || []).length ? 'identificado' : 'com planejamento')) + ' (esperado ' + p.esperado + ')');
    }
    cs.forEach(function (c) {
      if (c.estado === 'falta') alertas.push(p.rotulo + ' (' + c.comp + '): sem ' + c.esperado);
      if (c.ladoConflitoICPI) alertas.push(p.rotulo + ' (' + c.comp + '): pela ordem do CIDR o lado é ' + c.lado + ', mas o ICPI aponta ' + c.ladoConflitoICPI);
      if (c.estado === 'semplano') alertas.push(p.rotulo + ' (' + c.comp + (c.ext && c.comp.indexOf(' m') < 0 ? ', ' + c.ext + ' m' : '') + '): ' +
        (c.provavel ? 'componente do CIDR sem planejamento, provavelmente esta posição' : 'sem planejamento') + ', esperado ' + c.esperado);
    });
    if (cs.filter(function (c) { return !c.semPlano; }).length > 1 && sl.pos !== 'leito') alertas.push(p.rotulo + ': ' + cs.length + ' componentes na mesma posição');
    posicoes.push(p);
  });
  var usados = {};
  posicoes.forEach(function (p) { p.comps.forEach(function (c) { usados[c.chave] = true; }); });
  var fora = out.filter(function (c) { return !usados[c.chave]; }).sort(function (a, b) { return (a.ordem || 0) - (b.ordem || 0); });
  fora.forEach(function (c) {
    if (c.pos === 'equip' || c.pos === 'terra' || c.pos === 'outro') return;
    alertas.push(c.comp + ': ' + PLH_POS_ROTULO[c.pos] + (c.lado ? ' lado ' + c.lado + ' (não é ponta do duto no CIDR)' : ' sem lado identificado'));
  });
  return { mat: mat, pontas: pontas, posicoes: posicoes, fora: fora, alertas: alertas, comps: out };
}

/** Acrescenta na aba os componentes que ainda não estão lá (não altera linhas existentes). */
function plhGravarComponentes_(comps, linhas) {
  return; // desativado: as correções manuais não são usadas
  if (!comps.sh) return;
  var novas = [];
  linhas.forEach(function (l) {
    if (!l.esq) return;
    l.esq.comps.forEach(function (c) {
      if (!c.chave || comps.porId[c.chave] || /^CIDR\d/.test(c.chave)) return;
      comps.porId[c.chave] = { existe: true };
      novas.push([c.chave, l.objeto, c.comp, c.tipo, c.ext === null ? '' : c.ext, c.procs.join(', '),
                  PLH_POS_ROTULO[c.posAuto] || '', c.ladoAuto || '', '', '', '']);
    });
  });
  if (!novas.length) return;
  novas.sort(function (a, b) { return a[1] === b[1] ? (a[2] < b[2] ? -1 : 1) : (a[1] < b[1] ? -1 : 1); });
  try {
    var ini = Math.max(comps.sh.getLastRow(), 1) + 1;
    comps.sh.getRange(ini, 1, novas.length, PLH_COMP_COLS.length).setValues(novas);
  } catch (e) { comps.erro = 'Não foi possível gravar na aba PLH_Componentes: ' + e.message; }
}


/* ============================================================================
 *  PERIODICIDADE DE REFERÊNCIA (aba PLH_Periodicidade, editável pelo usuário)
 *  Procedimento × fluido → periodicidade em meses e a referência (ex.: N-1487).
 *  Fluido do duto vem do CIDR: ÁGUA → "Água"; ÓLEO/GÁS → "Óleo e gás"; linha "Todos" vale para qualquer fluido.
 * ==========================================================================*/
var PLH_ABA_PERIOD = 'PLH_Periodicidade';
var PLH_PERIOD_COLS = ['Procedimento', 'Fluido', 'Periodicidade (meses)', 'Referência', 'Observação'];
var PLH_PERIOD_PADRAO = [
  ['PIDR-4', 'Óleo e gás', 36, 'N-1487', 'Região 4 (emersa): 3 anos'],
  ['PIDR-4', 'Água', 60, 'N-1487', 'Região 4 (emersa): 5 anos'],
  ['PIDF-4', 'Óleo e gás', 36, 'N-1487', 'Região 4 (emersa): 3 anos'],
  ['PIDF-4', 'Água', 60, 'N-1487', 'Região 4 (emersa): 5 anos'],
  ['PIDR-3', 'Óleo e gás', 72, 'N-1487', 'Região 3: 6 anos'],
  ['PIDR-3', 'Água', 96, 'N-1487', 'Região 3: 8 anos'],
  ['PIDF-3', 'Óleo e gás', 72, 'N-1487', 'Região 3: 6 anos'],
  ['PIDF-3', 'Água', 96, 'N-1487', 'Região 3: 8 anos'],
  ['PIDR-1', 'Óleo e gás', 72, 'N-1487', 'Região 1: 6 anos'],
  ['PIDR-1', 'Água', 96, 'N-1487', 'Região 1: 8 anos'],
  ['PIDF-1', 'Óleo e gás', 72, 'N-1487', 'Região 1: 6 anos'],
  ['PIDF-1', 'Água', 96, 'N-1487', 'Região 1: 8 anos'],
  ['PIDR-0', 'Óleo e gás', 72, 'N-1487', '6 anos'],
  ['PIDR-0', 'Água', 96, 'N-1487', '8 anos']
];

function plhFluidoChave_(t) {
  var n = plhNorm_(t);
  if (!n || /todos|qualquer/.test(n)) return 'todos';
  if (/agua|aquaduto/.test(n)) return 'agua';
  if (/oleo|gas|petroleo/.test(n)) return 'oleogas';
  return n;
}

function plhLerPeriodicidade_(ss) {
  var res = { gid: null, lista: [], por: {}, erro: '' };
  var sh = ss.getSheetByName(PLH_ABA_PERIOD);
  if (!sh) {
    try {
      sh = ss.insertSheet(PLH_ABA_PERIOD);
      sh.getRange(1, 1, 1, PLH_PERIOD_COLS.length).setValues([PLH_PERIOD_COLS])
        .setFontWeight('bold').setBackground('#2c3e50').setFontColor('#ffffff');
      sh.getRange(2, 1, PLH_PERIOD_PADRAO.length, PLH_PERIOD_COLS.length).setValues(PLH_PERIOD_PADRAO);
      sh.setFrozenRows(1);
      sh.setTabColor('#00a085');
    } catch (e) {
      res.erro = 'Não foi possível criar a aba PLH_Periodicidade: ' + e.message;
      PLH_PERIOD_PADRAO.forEach(function (l) { plhAddRef_(res, l); });
      return res;
    }
  }
  res.gid = sh.getSheetId();
  var vals = sh.getDataRange().getValues();
  var cab = (vals[0] || []).map(plhNorm_);
  var c = PLH_PERIOD_COLS.map(function (n) { return cab.indexOf(plhNorm_(n)); });
  for (var i = 1; i < vals.length; i++) {
    var l = c.map(function (k) { return k < 0 ? '' : vals[i][k]; });
    if (!plhTxt_(l[0]) || plhNumC_(l[2]) === null) continue;
    plhAddRef_(res, l);
  }
  return res;
}
function plhAddRef_(res, l) {
  var r = { proc: plhTxt_(l[0]).toUpperCase(), fluido: plhTxt_(l[1]) || 'Todos', meses: plhNumC_(l[2]),
            fonte: plhTxt_(l[3]) || 'PLH_Periodicidade', obs: plhTxt_(l[4]) };
  res.lista.push(r);
  res.por[r.proc + '|' + plhFluidoChave_(r.fluido)] = r;
}
/** referência para o procedimento no fluido do duto; sem fluido conhecido, só a linha "Todos" */
function plhRefPeriodicidade_(refPer, proc, obj) {
  if (!refPer || !refPer.lista.length) return null;
  var f = obj ? plhFluidoChave_(obj.fluido || obj.funcao) : 'todos';
  if (f === 'todos' || f === '') return refPer.por[proc + '|todos'] || null;
  return refPer.por[proc + '|' + f] || refPer.por[proc + '|todos'] || null;
}
