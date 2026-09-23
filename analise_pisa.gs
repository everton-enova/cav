/**
 * ============================================================================
 *  gerarAbaAnalise()  -  Gera a aba "análise" (PISA 2025) com destaque para a
 *  BAHIA, cria gráficos e ALIMENTA a aba "índice" com navegação.
 * ============================================================================
 *
 *  COMO USAR
 *  1. Abra a planilha no Google Sheets:
 *     https://docs.google.com/spreadsheets/d/17_rRhQOboOpdtFtyj3CmPVfR1WBHFKRccKKoi1LH3dA/edit
 *  2. Menu  Extensões  >  Apps Script
 *  3. Apague o conteúdo e cole TODO este arquivo. Salve (Ctrl+S).
 *  4. Selecione a função  gerarAbaAnalise  e clique em  Executar.
 *  5. Autorize o script quando pedido (é a sua própria conta).
 *
 *  O que o script faz:
 *   - Cria/atualiza a aba "análise" com:
 *       0. Bahia em destaque (pontos fortes primeiro, depois o panorama)
 *       1. Desempenho de todas as UFs + ranking
 *       2. Proficiência (% abaixo do Nível 2 e % Níveis 5-6)
 *       3. Diferenças de gênero
 *       4. Status socioeconômico (ESCS)
 *       5. Tipo de escola
 *       6. Resiliência socioeconômica (ranking)
 *       7. Notas metodológicas
 *   - Cria 3 gráficos: ranking por UF (Bahia destacada), radar de equidade e
 *     linha comparando a Bahia com os estados vizinhos.
 *   - Reescreve a aba "índice" como um sumário navegável (links nativos).
 *   - Padroniza a tipografia de TODA a planilha: fonte Inter, escala de
 *     tamanhos e paleta de cores consistentes (inclusive as 17 tabelas de origem).
 *
 *  Para reaplicar SÓ o visual, sem regerar a análise, rode  padronizarPlanilha().
 *  Obs.: a fonte Inter precisa estar disponível no seletor de fontes do Sheets.
 *
 *  Fonte: PISA 2025. A linha "Brasil" está vazia nas tabelas originais, então
 *  usamos a média simples entre as 27 UFs como referência nacional.
 * ============================================================================
 */

var UF_DESTAQUE = 'Bahia';
var ABA_INDICE = 'índice';

var UFS = [
  'Acre', 'Alagoas', 'Amapá', 'Amazonas', 'Bahia', 'Ceará', 'Distrito Federal',
  'Espírito Santo', 'Goiás', 'Maranhão', 'Mato Grosso', 'Mato Grosso do Sul',
  'Minas Gerais', 'Pará', 'Paraíba', 'Paraná', 'Pernambuco', 'Piauí',
  'Rio de Janeiro', 'Rio Grande do Norte', 'Rio Grande do Sul', 'Rondônia',
  'Roraima', 'Santa Catarina', 'São Paulo', 'Sergipe', 'Tocantins'
];

var DOMINIOS = ['Ciências', 'Leitura', 'Matemática', 'Resolução de problemas'];

var ABAS_DESEMPENHO = {
  'Ciências': 'Tabela I.B2.1',
  'Leitura': 'Tabela I.B2.2',
  'Matemática': 'Tabela I.B2.3',
  'Resolução de problemas': 'Tabela I.B2.4'
};

var ABAS_PROFICIENCIA = {
  'Ciências':               { aba: 'Tabela I.B2.9',  abaixo: [2, 4],      topo: [14, 16] },
  'Leitura':                { aba: 'Tabela I.B2.10', abaixo: [2, 4, 6, 8], topo: [16, 18] },
  'Matemática':             { aba: 'Tabela I.B2.11', abaixo: [2, 4, 6, 8], topo: [16, 18] },
  'Resolução de problemas': { aba: 'Tabela I.B2.12', abaixo: [2, 4],      topo: [12, 14] }
};

var ABAS_GENERO = {
  'Ciências': 'Tabela I.B2.17',
  'Leitura': 'Tabela I.B2.18',
  'Matemática': 'Tabela I.B2.19',
  'Resolução de problemas': 'Tabela I.B2.20'
};

var ABAS_SES = {
  'Ciências': 'Tabela I.B2.29',
  'Leitura': 'Tabela I.B2.30',
  'Matemática': 'Tabela I.B2.31',
  'Resolução de problemas': 'Tabela I.B2.32'
};

var ABA_ESCOLA = 'Tabela I.B2.77';

// descrições originais da aba índice (fallback — nunca se perdem)
var DESCRICOES = {
  'análise': 'Análise consolidada do PISA 2025 com destaque para a Bahia',
  'Tabela I.B2.1': 'Pontuação média e variação no desempenho em ciências',
  'Tabela I.B2.2': 'Pontuação média e variação no desempenho em leitura',
  'Tabela I.B2.3': 'Pontuação média e variação no desempenho em matemática',
  'Tabela I.B2.4': 'Pontuação média e variação no desempenho em resolução computacional de problemas',
  'Tabela I.B2.9': 'Percentual de estudantes em cada nível de proficiência em ciências',
  'Tabela I.B2.10': 'Percentual de estudantes em cada nível de proficiência em leitura',
  'Tabela I.B2.11': 'Percentual de estudantes em cada nível de proficiência em matemática',
  'Tabela I.B2.12': 'Percentual de estudantes em cada nível de proficiência em resolução computacional de problemas',
  'Tabela I.B2.17': 'Desempenho em ciências, por gênero',
  'Tabela I.B2.18': 'Desempenho em leitura, por gênero',
  'Tabela I.B2.19': 'Desempenho em matemática, por gênero',
  'Tabela I.B2.20': 'Desempenho em resolução computacional de problemas, por gênero',
  'Tabela I.B2.29': 'Status socioeconômico e desempenho em ciências',
  'Tabela I.B2.30': 'Status socioeconômico e desempenho em leitura',
  'Tabela I.B2.31': 'Status socioeconômico e desempenho em matemática',
  'Tabela I.B2.32': 'Status socioeconômico e desempenho em resolução computacional de problemas',
  'Tabela I.B2.77': 'Tipo de escola'
};

// estados vizinhos / do Nordeste para comparação direta com a Bahia
var VIZINHOS = ['Bahia', 'Alagoas', 'Sergipe', 'Pernambuco', 'Paraíba',
  'Rio Grande do Norte', 'Ceará', 'Piauí', 'Maranhão'];

// ===========================================================================
//  DESIGN SYSTEM — tipografia e cores padronizadas
// ===========================================================================
var FONTE = 'Inter';
var TAM = { titulo: 16, secao: 13, subsecao: 11, header: 10, corpo: 10, nota: 9 };
var COR = {
  vermelho: '#c8102e',   // títulos
  azul: '#1a4f9c',       // seções
  azulClaro: '#d9e2f3',  // subseções
  cinza: '#eef3fb',      // cabeçalhos de tabela
  cinza2: '#e8f0fe',     // cabeçalhos de bloco
  amarelo: '#fff2cc',    // destaque Bahia
  verde: '#e6f4ea',      // resultado positivo
  zebra1: '#ffffff',     // faixa par
  zebra2: '#f5f8fd',     // faixa ímpar
  texto: '#202124',
  branco: '#ffffff'
};


// ---------------------------------------------------------------------------
//  Ponto de entrada
// ---------------------------------------------------------------------------
function gerarAbaAnalise() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var nomeAba = 'análise';

  var antiga = ss.getSheetByName(nomeAba);
  if (antiga) ss.deleteSheet(antiga);

  var sh = ss.insertSheet(nomeAba);
  sh.setTabColor('#c8102e');

  // ---------- coleta ----------
  var mediasPorDom = {}, rankPorDom = {}, dispersaoPorDom = {};
  DOMINIOS.forEach(function (dom) {
    mediasPorDom[dom] = lerColunaPorUF(ABAS_DESEMPENHO[dom], 2);
    dispersaoPorDom[dom] = lerColunaPorUF(ABAS_DESEMPENHO[dom], 16);
    rankPorDom[dom] = ranking(mediasPorDom[dom]);
  });

  var abaixoPorDom = {}, topoPorDom = {};
  DOMINIOS.forEach(function (dom) {
    var cfg = ABAS_PROFICIENCIA[dom];
    abaixoPorDom[dom] = somaColunasPorUF(cfg.aba, cfg.abaixo);
    topoPorDom[dom] = somaColunasPorUF(cfg.aba, cfg.topo);
  });

  var generoPorDom = {};
  DOMINIOS.forEach(function (dom) { generoPorDom[dom] = lerColunasPorUF(ABAS_GENERO[dom], [2, 17, 32]); });

  var sesR2 = {}, sesGap = {}, sesResil = {};
  DOMINIOS.forEach(function (dom) {
    sesR2[dom] = lerColunaPorUF(ABAS_SES[dom], 2);
    sesGap[dom] = lerColunaPorUF(ABAS_SES[dom], 20);
    sesResil[dom] = lerColunaPorUF(ABAS_SES[dom], 23);
  });

  var privPorUF = lerColunaPorUF(ABA_ESCOLA, 2);

  var generoDifMedia = {};
  DOMINIOS.forEach(function (d) {
    var s = 0, n = 0;
    UFS.forEach(function (u) { var v = val(generoPorDom[d], u, 2); if (typeof v === 'number') { s += v; n++; } });
    generoDifMedia[d] = n ? s / n : null;
  });

  // ---------- buffer ----------
  var linhas = [], estilo = [];
  function add(arr, st) { linhas.push(arr); estilo.push(st || ''); return linhas.length; }
  function branco() { return add(['']); }

  // ====================== CABEÇALHO ======================
  add(['PISA 2025 — ANÁLISE DOS RESULTADOS POR UNIDADE FEDERATIVA'], 'titulo');
  add(['DESTAQUE: BAHIA.  Fonte: abas "Tabela I.B2.*". Gerado em ' +
       Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm') + '.']);
  add(['A linha da Bahia aparece destacada em todos os blocos. ' +
       'Referência "média das UFs" = média simples das 27 unidades federativas.']);
  branco();

  // ====================== BLOCO 0 ======================
  add(['0. BAHIA EM DESTAQUE'], 'secao');
  branco();

  add(['0.1 Pontos fortes da Bahia (comparação com a média das UFs)'], 'subsecao');
  add(['Indicador', 'Domínio', 'Bahia', 'Média das UFs', 'Situação'], 'header');

  var forcas = [];

  var cntResil = 0;
  DOMINIOS.forEach(function (d) {
    var v = sesResil[d][UF_DESTAQUE], m = mediaDeMapa(sesResil[d]);
    var melhor = v >= m; if (melhor) cntResil++;
    add(['Alunos resilientes socioeconômicos (%)', d, num(v), num(m),
         melhor ? 'ACIMA da média (+' + (v - m).toFixed(1) + ' p.p.)' : 'abaixo da média'], melhor ? 'bom' : 'bahia');
  });
  if (cntResil > 0) forcas.push('Resiliência socioeconômica acima da média das UFs em ' + cntResil +
    ' de 4 domínios: alunos desfavorecidos da Bahia alcançam o quartil superior de desempenho em proporção maior que a média nacional.');

  var cntGap = 0;
  DOMINIOS.forEach(function (d) {
    var v = sesGap[d][UF_DESTAQUE], m = mediaDeMapa(sesGap[d]);
    var melhor = v <= m; if (melhor) cntGap++;
    add(['Desigualdade socioeconômica (gap Q4-Q1, pts)', d, num(v), num(m),
         melhor ? 'MENOR que a média (-' + (m - v).toFixed(0) + ' pts)' : 'acima da média'], melhor ? 'bom' : 'bahia');
  });
  if (cntGap > 0) forcas.push('Gap socioeconômico (quartil superior - inferior) abaixo da média das UFs em ' + cntGap +
    ' de 4 domínios: a distância entre os alunos mais ricos e mais pobres é comparativamente menor.');

  var cntR2 = 0;
  DOMINIOS.forEach(function (d) {
    var v = sesR2[d][UF_DESTAQUE], m = mediaDeMapa(sesR2[d]);
    var melhor = v <= m; if (melhor) cntR2++;
    add(['R² do ESCS (% variância explicada)', d, num(v), num(m),
         melhor ? 'MENOR que a média' : 'acima da média'], melhor ? 'bom' : 'bahia');
  });
  if (cntR2 > 0) forcas.push('O status socioeconômico explica MENOS da variância do desempenho na Bahia do que na média das UFs em ' +
    cntR2 + ' de 4 domínios — desempenho menos determinado pela origem social.');

  var cntDisp = 0;
  DOMINIOS.forEach(function (d) {
    var v = dispersaoPorDom[d][UF_DESTAQUE], m = mediaDeMapa(dispersaoPorDom[d]);
    var melhor = v <= m; if (melhor) cntDisp++;
    add(['Amplitude 90-10 (menor = mais homogêneo)', d, num(v), num(m),
         melhor ? 'MENOR que a média' : 'acima da média'], melhor ? 'bom' : 'bahia');
  });
  if (cntDisp > 0) forcas.push('Menor dispersão entre os 10% piores e 10% melhores alunos (90-10) em ' + cntDisp +
    ' de 4 domínios — rede mais homogênea entre os extremos.');

  var gCien = generoPorDom['Ciências'];
  add(['Diferença de gênero em Ciências (meninos - meninas)', 'Ciências',
       num(val(gCien, UF_DESTAQUE, 2)), num(generoDifMedia['Ciências']),
       val(gCien, UF_DESTAQUE, 2) === 0 ? 'IGUALDADE plena' : 'pequena diferença'],
       val(gCien, UF_DESTAQUE, 2) === 0 ? 'bom' : 'bahia');
  if (val(gCien, UF_DESTAQUE, 2) === 0)
    forcas.push('Igualdade de gênero em Ciências: meninas e meninos da Bahia têm exatamente a mesma média (365 pontos).');

  branco();
  add(['0.2 Resumo dos pontos fortes'], 'subsecao');
  if (forcas.length === 0) forcas.push('Consulte o bloco 0.1 para as comparações detalhadas.');
  forcas.forEach(function (t) { add(['• ' + t], 'bom'); });
  branco();

  add(['0.3 Desempenho médio e posição da Bahia (panorama completo)'], 'subsecao');
  add(['UF', 'Ciências', 'Leitura', 'Matemática', 'Resolução de problemas', 'Ranking médio'], 'header');
  var lh = [UF_DESTAQUE];
  DOMINIOS.forEach(function (d) { lh.push(num(mediasPorDom[d][UF_DESTAQUE])); });
  lh.push(num(mediaRank(rankPorDom, UF_DESTAQUE)));
  add(lh, 'bahia');
  var lm = ['Média das UFs'];
  DOMINIOS.forEach(function (d) { lm.push(num(mediaDeMapa(mediasPorDom[d]))); });
  lm.push(''); add(lm, 'media');
  var lb = ['Melhor UF'];
  DOMINIOS.forEach(function (d) { lb.push(melhorUF(mediasPorDom[d])[0]); });
  lb.push(''); add(lb, 'media');
  add(['Posição da Bahia (1 = melhor)'].concat(DOMINIOS.map(function (d) { return rankPorDom[d][UF_DESTAQUE]; })).concat(['']), 'bahia');
  branco();

  // ====================== BLOCO 1 ======================
  add(['1. DESEMPENHO MÉDIO DE TODAS AS UFs (pontuação PISA, 0-1000)'], 'secao');
  add(['UF', 'Ciências', 'Leitura', 'Matemática', 'Resolução de problemas', 'Ranking médio'], 'header');
  UFS.slice().sort(porRankingMedio(rankPorDom)).forEach(function (uf) {
    add([uf, num(mediasPorDom['Ciências'][uf]), num(mediasPorDom['Leitura'][uf]),
         num(mediasPorDom['Matemática'][uf]), num(mediasPorDom['Resolução de problemas'][uf]),
         num(mediaRank(rankPorDom, uf))], uf === UF_DESTAQUE ? 'bahia' : '');
  });
  branco();

  // ====================== BLOCO 2 ======================
  add(['2. PROFICIÊNCIA: % ABAIXO DO NÍVEL 2 E % NÍVEIS 5-6'], 'secao');
  add(['UF', 'Abaixo N2 Ciências', 'N5-6 Ciências', 'Abaixo N2 Leitura', 'N5-6 Leitura',
       'Abaixo N2 Matemática', 'N5-6 Matemática', 'Abaixo N2 Res.probl.', 'N5-6 Res.probl.'], 'header');
  UFS.slice().sort(porRankingMedio(rankPorDom)).forEach(function (uf) {
    add([uf,
      num(abaixoPorDom['Ciências'][uf]), num(topoPorDom['Ciências'][uf]),
      num(abaixoPorDom['Leitura'][uf]), num(topoPorDom['Leitura'][uf]),
      num(abaixoPorDom['Matemática'][uf]), num(topoPorDom['Matemática'][uf]),
      num(abaixoPorDom['Resolução de problemas'][uf]), num(topoPorDom['Resolução de problemas'][uf])
    ], uf === UF_DESTAQUE ? 'bahia' : '');
  });
  branco();

  // ====================== BLOCO 3 ======================
  add(['3. DIFERENÇA DE GÊNERO (meninos - meninas; + = meninos melhor)'], 'secao');
  add(['UF', 'Meninas Ciências', 'Meninos Ciências', 'Dif. Ciências',
       'Meninas Leitura', 'Meninos Leitura', 'Dif. Leitura',
       'Meninas Matemática', 'Meninos Matemática', 'Dif. Matemática',
       'Meninas Res.probl.', 'Meninos Res.probl.', 'Dif. Res.probl.'], 'header');
  UFS.forEach(function (uf) {
    add([uf,
      val(generoPorDom['Ciências'], uf, 0), val(generoPorDom['Ciências'], uf, 1), val(generoPorDom['Ciências'], uf, 2),
      val(generoPorDom['Leitura'], uf, 0), val(generoPorDom['Leitura'], uf, 1), val(generoPorDom['Leitura'], uf, 2),
      val(generoPorDom['Matemática'], uf, 0), val(generoPorDom['Matemática'], uf, 1), val(generoPorDom['Matemática'], uf, 2),
      val(generoPorDom['Resolução de problemas'], uf, 0), val(generoPorDom['Resolução de problemas'], uf, 1),
      val(generoPorDom['Resolução de problemas'], uf, 2)
    ], uf === UF_DESTAQUE ? 'bahia' : '');
  });
  branco();

  // ====================== BLOCO 4 ======================
  add(['4. STATUS SOCIOECONÔMICO (ESCS)'], 'secao');
  add(['UF', 'R² Ciências (%)', 'Gap Q4-Q1 Ciências', 'Resilientes Ciências (%)',
       'R² Leitura (%)', 'Gap Q4-Q1 Leitura', 'Resilientes Leitura (%)',
       'R² Matemática (%)', 'Gap Q4-Q1 Matemática', 'Resilientes Matemática (%)',
       'R² Res.probl. (%)', 'Gap Q4-Q1 Res.probl.', 'Resilientes Res.probl. (%)'], 'header');
  UFS.forEach(function (uf) {
    add([uf,
      val2(sesR2, 'Ciências', uf), val2(sesGap, 'Ciências', uf), val2(sesResil, 'Ciências', uf),
      val2(sesR2, 'Leitura', uf), val2(sesGap, 'Leitura', uf), val2(sesResil, 'Leitura', uf),
      val2(sesR2, 'Matemática', uf), val2(sesGap, 'Matemática', uf), val2(sesResil, 'Matemática', uf),
      val2(sesR2, 'Resolução de problemas', uf), val2(sesGap, 'Resolução de problemas', uf), val2(sesResil, 'Resolução de problemas', uf)
    ], uf === UF_DESTAQUE ? 'bahia' : '');
  });
  branco();

  // ====================== BLOCO 5 ======================
  add(['5. TIPO DE ESCOLA (% de estudantes)'], 'secao');
  add(['UF', '% Privada independente', '% Pública (pública + privada dependente do governo)'], 'header');
  UFS.forEach(function (uf) {
    var p = num(privPorUF[uf]);
    add([uf, p, p === null ? null : Math.round((100 - p) * 10) / 10], uf === UF_DESTAQUE ? 'bahia' : '');
  });
  branco();

  // ====================== BLOCO 6: RESILIÊNCIA ======================
  add(['6. RESILIÊNCIA SOCIOECONÔMICA — RANKING'], 'secao');
  add(['UF', 'Resilientes Ciências (%)', 'Resilientes Leitura (%)', 'Resilientes Matemática (%)',
       'Resilientes Res.probl. (%)', 'Média (%)'], 'header');
  var resilMedia = mediaPorUF(sesResil);
  UFS.slice().sort(function (a, b) { return (resilMedia[b] || 0) - (resilMedia[a] || 0); }).forEach(function (uf) {
    add([uf,
      num(sesResil['Ciências'][uf]), num(sesResil['Leitura'][uf]),
      num(sesResil['Matemática'][uf]), num(sesResil['Resolução de problemas'][uf]),
      Math.round((resilMedia[uf] || 0) * 10) / 10
    ], uf === UF_DESTAQUE ? 'bahia' : '');
  });
  branco();

  // ====================== BLOCO 7 ======================
  add(['7. NOTAS METODOLÓGICAS E ALERTAS'], 'secao');
  [
    '• Fonte: PISA 2025. Os resultados são por unidade federativa.',
    '• A linha "Brasil" está vazia nas tabelas originais: a referência usada é a média simples das 27 UFs.',
    '• Erros padrão (E.P.) aparecem NEGATIVOS na fonte. As magnitudes são coerentes com',
    '  erros padrão; use o valor absoluto.',
    '• A ordem das UFs está embaralhada em algumas abas (I.B2.11, I.B2.12, I.B2.77): o script',
    '  localiza cada UF pelo nome, não pela posição.',
    '• Marcadores "c" nas colunas de E.P. indicam anotações/supressões de estimativa.',
    '• "Abaixo do Nível 2" = soma das categorias abaixo do nível mínimo do PISA.',
    '• Indicadores do bloco 0.1: Resilientes (maior melhor); Gap Q4-Q1, R² do ESCS e',
    '  amplitude 90-10 (menor melhor = mais equidade/homogeneidade).',
    '• Ranking médio = média das posições nos 4 domínios (1 = melhor).',
    '• Associação não implica causalidade.'
  ].forEach(function (t) { add([t]); });

  // ---------- escrita ----------
  var nCol = maxColunas(linhas);
  sh.clear();
  sh.getRange(1, 1, linhas.length, nCol).setValues(padronizar(linhas));

  for (var i = 0; i < linhas.length; i++) {
    var r = i + 1, st = estilo[i];
    if (st === 'titulo') sh.getRange(r, 1, 1, nCol).setBackground('#c8102e').setFontColor('#ffffff').setFontWeight('bold').setFontSize(13);
    else if (st === 'secao') sh.getRange(r, 1, 1, nCol).setBackground('#1a4f9c').setFontColor('#ffffff').setFontWeight('bold').setFontSize(12);
    else if (st === 'subsecao') sh.getRange(r, 1, 1, nCol).setBackground('#d9e2f3').setFontWeight('bold');
    else if (st === 'header') sh.getRange(r, 1, 1, nCol).setBackground('#e8f0fe').setFontWeight('bold');
    else if (st === 'bahia') sh.getRange(r, 1, 1, nCol).setBackground('#fff2cc').setFontWeight('bold');
    else if (st === 'bom') sh.getRange(r, 1, 1, nCol).setBackground('#e6f4ea');
    else if (st === 'media') sh.getRange(r, 1, 1, nCol).setFontStyle('italic');
  }
  sh.setColumnWidth(1, 210);
  for (var c2 = 2; c2 <= nCol; c2++) sh.setColumnWidth(c2, 95);
  sh.setFrozenRows(1);
  sh.getRange(1, 1, linhas.length, nCol).setNumberFormat('0.0');
  sh.getRange(2, 1, 2, nCol).setNumberFormat('@');

  // ---------- gráficos ----------
  criarGraficos_(sh, linhas.length, nCol, {
    mediasPorDom: mediasPorDom, rankPorDom: rankPorDom, dispersaoPorDom: dispersaoPorDom,
    sesGap: sesGap, sesResil: sesResil, generoPorDom: generoPorDom
  });

  // ---------- alimenta a aba índice ----------
  var secoes = [];
  linhas.forEach(function (l, k) { if (estilo[k] === 'secao') secoes.push({ titulo: String(l[0]), linha: k + 1 }); });
  alimentarIndice_(sh, secoes);

  // ---------- padronização visual de TODA a planilha (fonte Inter) ----------
  try { padronizarPlanilha(); } catch (e) { Logger.log('Falha ao padronizar visual: ' + e); }

  ss.toast('Aba "análise" gerada e aba "índice" atualizada!', 'Concluído', 6);
  return sh.getName();
}


// ===========================================================================
//  Gráficos
// ===========================================================================
function criarGraficos_(sh, ultimaLinha, nCol, dados) {
  var helperCol = nCol + 2;
  var rowBar = ultimaLinha + 2;
  var rowRadar = rowBar + 31;

  // ----- Bar: ranking médio por UF, com a Bahia em série separada -----
  var helper = [];
  helper.push(['UF', 'Demais UFs', 'Bahia']);
  UFS.slice().sort(porRankingMedio(dados.rankPorDom)).forEach(function (uf) {
    var v = mediaRank(dados.rankPorDom, uf);
    if (uf === UF_DESTAQUE) helper.push([uf, '', v]);
    else helper.push([uf, v, '']);
  });
  sh.getRange(rowBar, helperCol, helper.length, 3).setValues(helper);

  var barChart = sh.newChart()
    .setChartType(Charts.ChartType.BAR)
    .addRange(sh.getRange(rowBar, helperCol, helper.length, 3))
    .setPosition(ultimaLinha + 3, 1, 0, 0)
    .setOption('title', 'Ranking médio por UF (menor = melhor) — Bahia em destaque')
    .setOption('legend', { position: 'top' })
    .setOption('colors', ['#9aa0a6', '#c8102e'])
    .setOption('width', 620).setOption('height', 900)
    .build();
  sh.insertChart(barChart);

  // ----- Radar: índice de equidade (0-100) Bahia vs média das UFs -----
  var desempenho = mediaPorUF(dados.mediasPorDom);
  var resiliencia = mediaPorUF(dados.sesResil);
  var equidadeSes = mediaPorUF(dados.sesGap);
  var homogeneidade = mediaPorUF(dados.dispersaoPorDom);
  var generoAbs = {};
  UFS.forEach(function (u) {
    var s = 0, n = 0;
    DOMINIOS.forEach(function (d) { var v = val(dados.generoPorDom[d], u, 2); if (typeof v === 'number') { s += Math.abs(v); n++; } });
    generoAbs[u] = n ? s / n : null;
  });

  var indicadores = [
    { label: 'Desempenho médio', mapa: desempenho, maiorMelhor: true },
    { label: 'Resiliência', mapa: resiliencia, maiorMelhor: true },
    { label: 'Equidade socioeconômica', mapa: equidadeSes, maiorMelhor: false },
    { label: 'Homogeneidade (90-10)', mapa: homogeneidade, maiorMelhor: false },
    { label: 'Igualdade de gênero', mapa: generoAbs, maiorMelhor: false }
  ];

  var radar = [['Dimensão', 'Bahia', 'Média das UFs']];
  indicadores.forEach(function (ind) {
    var sc = normalizaScore(ind.mapa, ind.maiorMelhor);
    radar.push([ind.label, sc(ind.mapa[UF_DESTAQUE]), sc(mediaDeMapa(ind.mapa))]);
  });
  sh.getRange(rowRadar, helperCol, radar.length, 3).setValues(radar);

  var radarChart = sh.newChart()
    .setChartType(Charts.ChartType.RADAR)
    .addRange(sh.getRange(rowRadar, helperCol, radar.length, 3))
    .setPosition(ultimaLinha + 3, 9, 0, 0)
    .setOption('title', 'Índice de equidade (0-100) — Bahia vs média das UFs')
    .setOption('legend', { position: 'top' })
    .setOption('colors', ['#c8102e', '#9aa0a6'])
    .setOption('vAxis', { viewWindow: { min: 0, max: 100 } })
    .setOption('width', 620).setOption('height', 480)
    .build();
  sh.insertChart(radarChart);

  // ----- Linha: Bahia vs estados vizinhos (Nordeste) por domínio -----
  var rowViz = rowRadar + 8;
  var viz = [['Domínio'].concat(VIZINHOS)];
  DOMINIOS.forEach(function (d) {
    var linha = [d];
    VIZINHOS.forEach(function (uf) { linha.push(num(dados.mediasPorDom[d][uf])); });
    viz.push(linha);
  });
  sh.getRange(rowViz, helperCol, viz.length, 1 + VIZINHOS.length).setValues(viz);

  var coresViz = ['#c8102e', '#9aa0a6', '#b7b7b7', '#c8c8c8', '#d9d9d9', '#e0e0e0', '#e8e8e8', '#efefef', '#f5f5f5'];
  var lineChart = sh.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(sh.getRange(rowViz, helperCol, viz.length, 1 + VIZINHOS.length))
    .setPosition(ultimaLinha + 30, 9, 0, 0)
    .setOption('title', 'Bahia e estados vizinhos — desempenho por domínio')
    .setOption('legend', { position: 'right' })
    .setOption('colors', coresViz)
    .setOption('pointSize', 6)
    .setOption('width', 620).setOption('height', 480)
    .build();
  sh.insertChart(lineChart);

  // esconde as colunas de apoio
  sh.hideColumns(helperCol, 10);
}


// ===========================================================================
//  Alimenta a aba "índice" com sumário navegável (com links)
// ===========================================================================
function alimentarIndice_(shAnalise, secoes) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var idx = ss.getSheetByName(ABA_INDICE) || ss.getSheetByName('indice');
  if (!idx) idx = ss.insertSheet(ABA_INDICE, 0);

  // guarda descrições já existentes (ignora lixo de fórmula/#ERROR!)
  var desc = {};
  if (idx.getLastRow() > 0) {
    idx.getDataRange().getValues().forEach(function (row) {
      var a = String(row[0] || '').trim();
      var b = String(row[1] || '').trim();
      if (!a || !b) return;
      if (a === 'Documento' || a === 'ÍNDICE') return;
      if (a.charAt(0) === '=' || a.indexOf('#ERROR') === 0) return;
      if (b.charAt(0) === '=' || b.indexOf('#ERROR') === 0) return;
      desc[a] = b;
    });
  }
  function descricao(nome) { return desc[nome] || DESCRICOES[nome] || ''; }

  idx.clear();

  var lines = [], types = [];
  function push(r, t) { lines.push(r); types.push(t || ''); }
  push(['ÍNDICE'], 'titulo');
  push(['Documento', 'Descrição / Acesso'], 'header');
  push(['análise', descricao('análise')], 'destaque');
  ss.getSheets().forEach(function (s) {
    var n = s.getName();
    if (n === ABA_INDICE || n === 'análise' || n === 'indice') return;
    push([n, descricao(n)], '');
  });
  push(['']);
  push(['ANÁLISE PISA 2025 — IR PARA O BLOCO'], 'secao');
  secoes.forEach(function (s) { push([s.titulo, ''], 'link'); });

  var nCol = 2;
  idx.getRange(1, 1, lines.length, nCol).setValues(padronizar(lines));

  // formatação
  for (var i = 0; i < lines.length; i++) {
    var r = i + 1, t = types[i];
    if (t === 'titulo') idx.getRange(r, 1, 1, nCol).setBackground('#c8102e').setFontColor('#ffffff').setFontWeight('bold').setFontSize(13);
    else if (t === 'header') idx.getRange(r, 1, 1, nCol).setBackground('#e8f0fe').setFontWeight('bold');
    else if (t === 'destaque') idx.getRange(r, 1, 1, nCol).setBackground('#fff2cc').setFontWeight('bold');
    else if (t === 'secao') idx.getRange(r, 1, 1, nCol).setBackground('#1a4f9c').setFontColor('#ffffff').setFontWeight('bold');
  }

  // links nativos (RichText) — funcionam em qualquer idioma, sem fórmula
  var ssId = ss.getId();
  var base = 'https://docs.google.com/spreadsheets/d/' + ssId + '/edit';
  var gidAnalise = shAnalise.getSheetId();
  var linhasSecao = {};
  secoes.forEach(function (s) { linhasSecao[s.titulo] = s.linha; });
  for (var j = 0; j < lines.length; j++) {
    var nome = String(lines[j][0] || '');
    if (types[j] === 'destaque' || (j > 1 && types[j] === '' && ss.getSheetByName(nome))) {
      var alvo = ss.getSheetByName(nome);
      if (alvo) aplicarLink_(idx, j + 1, nome, base + '#gid=' + alvo.getSheetId() + '&range=A1');
    } else if (types[j] === 'link' && linhasSecao[nome]) {
      aplicarLink_(idx, j + 1, nome, base + '#gid=' + gidAnalise + '&range=A' + linhasSecao[nome]);
    }
  }

  idx.setColumnWidth(1, 320);
  idx.setColumnWidth(2, 640);
  idx.setFrozenRows(2);
}


// ===========================================================================
//  Auxiliares de leitura
// ===========================================================================
function lerColunaPorUF(nomeAba, col) {
  var mapa = lerColunasPorUF(nomeAba, [col]), out = {};
  Object.keys(mapa).forEach(function (uf) { out[uf] = mapa[uf][0]; });
  return out;
}

function lerColunasPorUF(nomeAba, cols) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nomeAba);
  if (!sh) throw new Error('Aba não encontrada: ' + nomeAba);
  var dados = sh.getDataRange().getValues(), out = {};
  for (var i = 0; i < dados.length; i++) {
    var nome = String(dados[i][0]).trim();
    if (UFS.indexOf(nome) === -1) continue;
    out[nome] = cols.map(function (c) { return num(dados[i][c - 1]); });
  }
  return out;
}

function somaColunasPorUF(nomeAba, cols) {
  var mapa = lerColunasPorUF(nomeAba, cols), out = {};
  Object.keys(mapa).forEach(function (uf) {
    var s = 0, achou = false;
    mapa[uf].forEach(function (v) { if (v !== null) { s += v; achou = true; } });
    out[uf] = achou ? s : null;
  });
  return out;
}

// ===========================================================================
//  Auxiliares de cálculo
// ===========================================================================
function ranking(mapa) {
  var pares = UFS.filter(function (u) { return typeof mapa[u] === 'number'; }).map(function (u) { return [u, mapa[u]]; });
  pares.sort(function (a, b) { return b[1] - a[1]; });
  var out = {};
  pares.forEach(function (p, i) { out[p[0]] = i + 1; });
  return out;
}

function mediaRank(rankPorDom, uf) {
  var soma = 0, n = 0;
  DOMINIOS.forEach(function (d) { if (rankPorDom[d][uf]) { soma += rankPorDom[d][uf]; n++; } });
  return n ? soma / n : null;
}

function porRankingMedio(rankPorDom) {
  return function (a, b) { return (mediaRank(rankPorDom, a) || 99) - (mediaRank(rankPorDom, b) || 99); };
}

function mediaDeMapa(m) {
  var s = 0, n = 0;
  Object.keys(m).forEach(function (k) { if (typeof m[k] === 'number') { s += m[k]; n++; } });
  return n ? s / n : null;
}

function mediaPorUF(mapaPorDom) {
  var out = {};
  UFS.forEach(function (u) {
    var s = 0, n = 0;
    DOMINIOS.forEach(function (d) { var v = mapaPorDom[d][u]; if (typeof v === 'number') { s += v; n++; } });
    out[u] = n ? s / n : null;
  });
  return out;
}

function melhorUF(m) {
  var best = null, bv = -Infinity;
  UFS.forEach(function (u) { if (typeof m[u] === 'number' && m[u] > bv) { bv = m[u]; best = u; } });
  return [best, bv];
}

function normalizaScore(mapa, maiorMelhor) {
  var vals = UFS.map(function (u) { return mapa[u]; }).filter(isNum);
  var mn = Math.min.apply(null, vals), mx = Math.max.apply(null, vals);
  return function (v) {
    if (!isNum(v) || mx === mn) return 50;
    var x = (v - mn) / (mx - mn);
    if (!maiorMelhor) x = 1 - x;
    return Math.round(x * 1000) / 10;
  };
}

function val(mapa, uf, idx) { return (mapa[uf] ? mapa[uf][idx] : null); }
function val2(mapaPorDom, dom, uf) { return mapaPorDom[dom][uf]; }

function isNum(v) { return typeof v === 'number' && isFinite(v); }

function num(v) {
  if (v === '' || v === null || v === undefined) return null;
  if (typeof v === 'number') return isFinite(v) ? v : null;
  var s = String(v).replace(',', '.').replace(/[^0-9.\-]/g, '');
  var n = parseFloat(s);
  return isFinite(n) ? n : null;
}

function escaparFormula(s) { return String(s).replace(/"/g, '""'); }

function aplicarLink_(sheet, row, texto, url) {
  try {
    var v = SpreadsheetApp.newRichTextValue().setText(texto).setLinkUrl(url).build();
    sheet.getRange(row, 1).setRichTextValue(v);
  } catch (e) {
    // se o link falhar por qualquer motivo, garante ao menos o texto
    sheet.getRange(row, 1).setValue(texto);
  }
}

function maxColunas(linhas) {
  var m = 0;
  linhas.forEach(function (l) { if (l.length > m) m = l.length; });
  return m;
}

function padronizar(linhas) {
  var m = maxColunas(linhas);
  return linhas.map(function (l) {
    var r = l.slice();
    while (r.length < m) r.push('');
    return r;
  });
}


// ===========================================================================
//  PADRONIZAÇÃO VISUAL — fonte Inter + tamanhos + cores consistentes
//  Rode  padronizarPlanilha()  a qualquer momento para reaplicar o padrão.
// ===========================================================================
function padronizarPlanilha() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.getSheets().forEach(function (sh) {
    var n = sh.getName();
    try {
      if (n === 'análise') estilizarAnalise_(sh);
      else if (n === ABA_INDICE || n === 'indice') estilizarIndice_(sh);
      else estilizarTabelaFonte_(sh);
    } catch (e) {
      Logger.log('Não foi possível estilizar a aba "' + n + '": ' + e);
    }
  });
}

function estilizarTabelaFonte_(sh) {
  var lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  if (lastRow < 1 || lastCol < 1) return;
  var dados = sh.getRange(1, 1, lastRow, lastCol).getValues();

  // primeira linha de dados (Brasil ou uma UF)
  var primeiraDados = -1;
  for (var i = 0; i < dados.length; i++) {
    var a = String(dados[i][0] || '').trim();
    if (a === 'Brasil' || UFS.indexOf(a) !== -1) { primeiraDados = i + 1; break; }
  }

  // base: Inter em toda a área usada
  sh.getRange(1, 1, lastRow, lastCol)
    .setFontFamily(FONTE).setFontColor(COR.texto).setVerticalAlignment('middle');

  // corpo da tabela
  var inicio = primeiraDados > 0 ? primeiraDados : 1;
  sh.getRange(inicio, 1, lastRow - inicio + 1, lastCol).setFontSize(TAM.corpo);

  // título (1ª linha não vazia antes da tabela)
  var titleRow = -1;
  var limite = primeiraDados > 0 ? primeiraDados - 1 : Math.min(lastRow, 3);
  for (var i = 0; i < limite; i++) {
    if (temConteudo(dados[i])) { titleRow = i + 1; break; }
  }
  if (titleRow > 0) {
    sh.getRange(titleRow, 1, 1, lastCol)
      .setFontSize(TAM.titulo).setFontWeight('bold').setFontColor(COR.vermelho);
  }

  // cabeçalhos (linhas não vazias entre título e dados)
  for (var i = 1; i < primeiraDados - 1; i++) {
    if (i + 1 === titleRow) continue;
    if (temConteudo(dados[i])) {
      sh.getRange(i + 1, 1, 1, lastCol)
        .setFontSize(TAM.header).setFontWeight('bold')
        .setBackground(COR.cinza).setFontColor(COR.azul)
        .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
    }
  }

  // nome da UF em negrito + faixas zebra no bloco de dados
  if (primeiraDados > 0) {
    var ultimaDados = primeiraDados;
    for (var i = 0; i < dados.length; i++) {
      var a = String(dados[i][0] || '').trim();
      if (a === 'Brasil' || UFS.indexOf(a) !== -1) ultimaDados = i + 1;
    }
    sh.getRange(primeiraDados, 1, ultimaDados - primeiraDados + 1, 1).setFontWeight('bold');
    aplicarZebra_(sh, primeiraDados - 1, ultimaDados, lastCol);
  }

  // notas de rodapé
  for (var i = Math.max(inicio - 1, 0); i < lastRow; i++) {
    var t = String(dados[i][0] || '');
    if (t.charAt(0) === '*' || /^(Nota|Consulte|As regi|\d\.)/i.test(t)) {
      sh.getRange(i + 1, 1, 1, lastCol).setFontSize(TAM.nota).setFontStyle('italic').setFontColor('#5f6368');
    }
  }
}

function estilizarAnalise_(sh) {
  var lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  if (lastRow < 1) return;
  sh.getRange(1, 1, lastRow, lastCol)
    .setFontFamily(FONTE).setFontSize(TAM.corpo).setFontColor(COR.texto).setVerticalAlignment('top');

  sh.getRange(1, 1, 1, lastCol)
    .setFontSize(TAM.titulo).setFontWeight('bold').setBackground(COR.vermelho).setFontColor(COR.branco);

  var vals = sh.getRange(1, 1, lastRow, 1).getValues();
  for (var i = 0; i < lastRow; i++) {
    var t = String(vals[i][0] || '');
    var row = i + 1;
    if (/^\d+\.\d+\s/.test(t)) {
      sh.getRange(row, 1, 1, lastCol).setFontSize(TAM.subsecao).setFontWeight('bold').setBackground(COR.azulClaro);
    } else if (/^\d+\.\s/.test(t)) {
      sh.getRange(row, 1, 1, lastCol).setFontSize(TAM.secao).setFontWeight('bold').setBackground(COR.azul).setFontColor(COR.branco);
    } else if (t === 'UF' || t === 'Indicador') {
      sh.getRange(row, 1, 1, lastCol).setFontSize(TAM.header).setFontWeight('bold').setBackground(COR.cinza2);
    } else if (t === UF_DESTAQUE) {
      sh.getRange(row, 1, 1, lastCol).setFontWeight('bold').setBackground(COR.amarelo);
    } else if (t.charAt(0) === '•') {
      sh.getRange(row, 1).setFontSize(TAM.nota);
    }
  }

  // faixas zebra sob cada cabeçalho 'UF' (não mexe no bloco 0.1, que é colorido)
  var c1 = sh.getRange(1, 1, lastRow, 1).getValues();
  for (var i = 0; i < lastRow; i++) {
    if (String(c1[i][0] || '') === 'UF') {
      var fim = i + 1;
      for (var k = i + 1; k < lastRow; k++) {
        var a = String(c1[k][0] || '');
        if (a === '' || a === 'UF' || a === 'Indicador' || /^\d+\./.test(a)) break;
        fim = k + 1;
      }
      if (fim > i + 1) aplicarZebra_(sh, i + 1, fim, lastCol);
    }
  }

  // reaplica o destaque da Bahia por cima das faixas
  for (var i = 0; i < lastRow; i++) {
    var t = String(c1[i][0] || '');
    if (t === UF_DESTAQUE || /^Posição da Bahia/.test(t)) {
      sh.getRange(i + 1, 1, 1, lastCol).setBackground(COR.amarelo).setFontWeight('bold');
    }
  }
}

function estilizarIndice_(sh) {
  var lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  if (lastRow < 1) return;
  sh.getRange(1, 1, lastRow, lastCol)
    .setFontFamily(FONTE).setFontSize(TAM.corpo).setFontColor(COR.texto).setVerticalAlignment('middle');
  sh.getRange(1, 1, 1, lastCol)
    .setFontSize(TAM.titulo).setFontWeight('bold').setBackground(COR.vermelho).setFontColor(COR.branco);
  sh.getRange(2, 1, 1, lastCol).setFontWeight('bold').setBackground(COR.cinza2);

  var vals = sh.getRange(1, 1, lastRow, 1).getValues();
  for (var i = 0; i < lastRow; i++) {
    var t = String(vals[i][0] || '');
    var row = i + 1;
    if (t === 'análise') {
      sh.getRange(row, 1, 1, lastCol).setBackground(COR.amarelo).setFontWeight('bold');
    } else if (/^ANÁLISE/.test(t)) {
      sh.getRange(row, 1, 1, lastCol).setBackground(COR.azul).setFontColor(COR.branco).setFontWeight('bold');
    }
  }
}

function temConteudo(linha) {
  for (var i = 0; i < linha.length; i++) if (String(linha[i]).trim() !== '') return true;
  return false;
}

/** Aplica faixas zebra de headerRow+1 até endRow (1-based), em uma única chamada. */
function aplicarZebra_(sh, headerRow, endRow, lastCol) {
  var n = endRow - headerRow;
  if (n < 1 || lastCol < 1) return;
  var colors = [];
  for (var i = 0; i < n; i++) {
    var cor = (i % 2 === 0) ? COR.zebra1 : COR.zebra2;
    var linha = [];
    for (var j = 0; j < lastCol; j++) linha.push(cor);
    colors.push(linha);
  }
  sh.getRange(headerRow + 1, 1, n, lastCol).setBackgrounds(colors);
}
