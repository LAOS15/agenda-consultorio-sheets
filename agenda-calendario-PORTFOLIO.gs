/**
 * AGENDA DO CONSULTÓRIO — calendário de segunda a sábado (Google Planilhas)
 * VERSÃO PORTFÓLIO: todos os nomes e telefones de exemplo são fictícios.
 *
 * INSTALAR (uma vez): planilha nova > Extensões > Apps Script > apague tudo e cole este código >
 * Salvar > escolha "configurarPlanilha" > Executar > autorize > volte à planilha e recarregue (F5).
 *
 * USAR:
 *  - Na aba do mês, digite o paciente na célula do dia/horário com a sigla da psicóloga:
 *      Maria (AN)        → individual
 *      Maria (AN+BE)     → em equipe
 *    (se o paciente tiver psicóloga cadastrada na aba Pacientes, a sigla entra sozinha)
 *  - Na célula ao lado, escolha o status. Cores, feriados, bloqueios e resumo são automáticos.
 *  - Faltou/Cancelou: escolha o motivo na aba "Faltas e cancelamentos".
 */

const MESES = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto',
  'Setembro', 'Outubro', 'Novembro', 'Dezembro'];
const DIAS = ['Seg', 'Ter', 'Qua', 'Qui', 'Sex', 'Sáb'];
const STATUS = ['Agendada', 'Confirmada', 'Compareceu', 'Faltou', 'Cancelou'];
const MOTIVOS = ['Doença', 'Imprevisto pessoal', 'Trabalho', 'Esqueceu', 'Questão financeira',
  'Transporte / clima', 'Remarcada', 'Desmarcado pela psicóloga', 'Sem justificativa', 'Outro'];
const COR = {
  Agendada: '#FFF2CC', Confirmada: '#D9EAD3', Compareceu: '#CFE2F3', Faltou: '#F4CCCC', Cancelou: '#E0E0E0',
  FERIADO: '#EADCF8', BLOQ: '#CCCCCC', FORA: '#F3F3F3', APAGADO: '#FAFAFA', AZUL: '#1F4E79', SEMANA: '#DDEBF7'
};
const ABA = { RESUMO: 'Resumo', PAC: 'Pacientes', FALTAS: 'Faltas e cancelamentos', PSI: 'Psicólogas',
  FER: 'Feriados', CFG: 'Config' };
const COLS = 13;                       // A = horário; cada dia usa 2 colunas (paciente, status)
const pc = d => 2 + 2 * d;             // coluna do paciente no dia d (0 = Seg ... 5 = Sáb)


// ============================== MENU ==============================

function onOpen() {
  SpreadsheetApp.getUi().createMenu('🗓️ Consultório')
    .addItem('Criar mês', 'menuCriarMes')
    .addItem('Repetir paciente fixo', 'repetirFixo')
    .addItem('Tirar paciente fixo das próximas semanas', 'tirarFixo')
    .addItem('Atualizar cores e resumo', 'atualizarTudo')
    .addSeparator()
    .addItem('Preencher com dados de exemplo (portfólio)', 'preencherExemplo')
    .addItem('Configurar planilha (primeira vez)', 'configurarPlanilha')
    .addToUi();
}


// ========================= CONFIGURAÇÃO INICIAL =========================

function configurarPlanilha() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetLocale('pt_BR');
  ss.setSpreadsheetTimeZone('America/Sao_Paulo');

  let a = tabela(ss, ABA.CFG, ['Item', 'Valor', 'Dica'], [220, 90, 460]);
  if (a) {
    a.getRange('B2:B4').setNumberFormat('@');
    a.getRange(2, 1, 3, 3).setValues([
      ['Primeiro horário', '08:00', 'Vale para os meses criados depois.'],
      ['Fim do expediente', '20:00', 'Nos meses já criados, é só digitar outro horário na coluna A.'],
      ['Duração da sessão (min)', '60', 'Ex.: 50 ou 60']]);
    a.getRange('B2:B4').setBackground('#FFF8DC');
  }

  a = tabela(ss, ABA.PSI, ['Nome', 'Sigla', 'Cor (pinte a célula)'].concat(DIAS), [170, 60, 130, 150, 150, 150, 150, 150, 150]);
  if (a) {
    a.getRange('A2:B5').setValues([['Psicóloga 1', 'P1'], ['Psicóloga 2', 'P2'], ['Psicóloga 3', 'P3'], ['Psicóloga 4', 'P4']]);
    a.getRange('C2:C5').setBackgrounds([['#F9CB9C'], ['#B6D7A8'], ['#9FC5E8'], ['#D5A6BD']]);
    a.getRange('D2:I').setNumberFormat('@');
    a.getRange('D1').setNote('Horários em que atende. Ex.: 08:00-12:00, 14:00-18:00\n' +
      'Vazio = não atende nesse dia.\nLinha inteira vazia = sem bloqueio.');
    a.getRange('B1').setNote('Sigla curta (2 ou 3 letras). É o que vai entre parênteses no calendário.');
  }

  a = tabela(ss, ABA.PAC, ['Nome', 'Telefone', 'Responsável (se menor)', 'Psicóloga (sigla)', 'Observações (não clínicas)'],
    [200, 130, 200, 130, 300]);
  if (a) a.getRange('D2:D').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInRange(ss.getSheetByName(ABA.PSI).getRange('B2:B'), true).build());

  a = tabela(ss, ABA.FALTAS, ['Data', 'Horário', 'Paciente', 'Status', 'Motivo', 'Registrado em'], [95, 70, 200, 95, 190, 130]);
  if (a) {
    a.getRange('A2:B').setNumberFormat('@');
    a.getRange('E2:E').setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(MOTIVOS, true).setAllowInvalid(true).build()).setBackground('#FFF8DC');
  }

  a = tabela(ss, ABA.FER, ['Data', 'Feriado'], [100, 260]);
  if (a) {
    a.getRange('A2:A').setNumberFormat('@');
    a.getRange('A1').setNote('Os nacionais entram sozinhos. Pode acrescentar feriados da cidade (ex.: 20/01/2027).');
  }

  a = tabela(ss, ABA.RESUMO, ['Mês:', ''], [300, 110, 110, 110, 110, 110]);
  if (a) a.getRange('B1').setNumberFormat('@').setBackground('#FFF8DC').setFontColor('black');

  const hoje = new Date(), ano = Number(fmt(hoje, 'yyyy')), mes = Number(fmt(hoje, 'M'));
  criarMes(ss, ano, mes);
  criarMes(ss, mes === 12 ? ano + 1 : ano, mes === 12 ? 1 : mes + 1);

  ['Página1', 'Sheet1', 'Planilha1'].forEach(n => { const s = ss.getSheetByName(n); if (s) ss.deleteSheet(s); });
  listaMeses(ss);
  resumo(ss);
  onOpen();
  ss.setActiveSheet(ss.getSheetByName(ABA.PSI));
  SpreadsheetApp.getUi().alert('Pronto! ✅\n\n' +
    '1. Aba "Psicólogas": nomes, siglas, cor e horários de cada uma.\n' +
    '2. Aba "Config": primeiro horário, fim do expediente e duração da sessão.\n' +
    '   (Se mudar, apague as abas dos meses e use 🗓️ Consultório > Criar mês.)\n' +
    '3. Aba "Pacientes": cadastre os pacientes.\n' +
    '4. Depois use 🗓️ Consultório > Atualizar cores e resumo.\n\n' +
    'Compartilhe a planilha com as psicólogas como Editor (não use "qualquer pessoa com o link").');
}

function tabela(ss, nome, cab, larguras) {
  if (ss.getSheetByName(nome)) return null;
  const a = ss.insertSheet(nome);
  a.getRange(1, 1, 1, cab.length).setValues([cab]).setFontWeight('bold').setBackground(COR.AZUL).setFontColor('white');
  a.setFrozenRows(1);
  larguras.forEach((l, i) => a.setColumnWidth(i + 1, l));
  return a;
}


// ============================== MESES ==============================

function menuCriarMes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), ui = SpreadsheetApp.getUi();
  const r = ui.prompt('Criar mês', 'Qual mês? (ex.: 12/2026)', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const m = r.getResponseText().match(/(\d{1,2})\s*\/\s*(\d{4})/);
  if (!m || m[1] < 1 || m[1] > 12) return ui.alert('Use o formato 12/2026.');
  const qtd = numero(ui, 'Quantos meses criar a partir dele? (1 a 12)');
  if (!qtd) return;
  let ano = Number(m[2]), mes = Number(m[1]), aba;
  for (let i = 0; i < Math.min(qtd, 12); i++) {
    aba = criarMes(ss, ano, mes);
    if (++mes > 12) { mes = 1; ano++; }
  }
  listaMeses(ss);
  ss.setActiveSheet(aba);
}

function criarMes(ss, ano, mes) {
  const nome = MESES[mes - 1] + ' ' + ano;
  if (ss.getSheetByName(nome)) return ss.getSheetByName(nome);
  garantirFeriados(ss, ano);

  const c = lerCfg(ss), horas = [];
  for (let t = c.ini; t + c.dur <= c.fim; t += c.dur) horas.push(hhmm(t));
  const ultimo = new Date(Date.UTC(ano, mes, 0)).getUTCDate();
  let seg = 1 - (new Date(Date.UTC(ano, mes - 1, 1)).getUTCDay() + 6) % 7;   // segunda-feira da 1ª semana

  const L = [linha(['🗓️ ' + nome]), linha(['Ver:', 'Todas']), linha([])];
  const blocos = [];
  for (let w = 1; seg <= ultimo; seg += 7) {
    const dias = DIAS.map((_, d) => (seg + d >= 1 && seg + d <= ultimo) ? seg + d : null);
    if (!dias.some(x => x)) continue;
    blocos.push(L.length + 1);
    L.push(linha(['SEMANA ' + (w++)]));
    const cab = linha(['Horário']);
    dias.forEach((dd, d) => { cab[pc(d) - 1] = dd ? DIAS[d] + ' ' + dm(dd, mes) : '—'; });
    L.push(cab);
    horas.forEach(h => L.push(linha([h])));
    L.push(linha([]));
  }

  const a = ss.insertSheet(nome);
  if (a.getMaxRows() < L.length) a.insertRowsAfter(a.getMaxRows(), L.length - a.getMaxRows());
  if (a.getMaxColumns() > COLS) a.deleteColumns(COLS + 1, a.getMaxColumns() - COLS);
  a.getRange(1, 1, L.length, COLS).setNumberFormat('@').setValues(L).setVerticalAlignment('middle');
  a.setColumnWidth(1, 60);
  DIAS.forEach((_, d) => { a.setColumnWidth(pc(d), 165); a.setColumnWidth(pc(d) + 1, 95); });
  a.setFrozenRows(2);
  a.setFrozenColumns(1);
  a.getRange('A1').setFontSize(15).setFontWeight('bold').setFontColor(COR.AZUL);
  a.getRange('A2').setFontWeight('bold');
  a.getRange('B2').setBackground('#FFF8DC').setNote('Escolha uma sigla para ver só os horários dela.');
  const leg = ['Agendada', 'Confirmada', 'Compareceu', 'Faltou', 'Cancelou', 'Feriado', 'Indisponível'];
  a.getRange(3, 2, 1, 7).setValues([leg]).setFontSize(9).setHorizontalAlignment('center')
    .setBackgrounds([leg.map(x => COR[x] || (x === 'Feriado' ? COR.FERIADO : COR.BLOQ))]);

  const regra = SpreadsheetApp.newDataValidation().requireValueInList(STATUS, true).build();
  blocos.forEach(r => {
    a.getRange(r, 1, 1, COLS).setBackground(COR.SEMANA).setFontWeight('bold');
    a.getRange(r + 1, 1, 1, COLS).setFontColor('white').setFontWeight('bold').setHorizontalAlignment('center');
    DIAS.forEach((_, d) => {
      a.getRange(r + 1, pc(d), 1, 2).merge();
      a.getRange(r + 2, pc(d) + 1, horas.length, 1).setDataValidation(regra).setFontSize(9);
    });
    a.getRange(r + 2, 1, horas.length, 1).setFontWeight('bold').setBackground('#F2F2F2').setHorizontalAlignment('center');
    a.getRange(r + 1, 1, horas.length + 1, COLS)
      .setBorder(true, true, true, true, true, true, '#BFBFBF', SpreadsheetApp.BorderStyle.SOLID);
  });

  atualizarFiltro(ss, a);
  pintar(ss, a);
  ordenarAbas(ss);
  return a;
}

/** Lê a aba do mês: blocos de semana, com os dias e as linhas de horário. */
function lerMes(a, dur) {
  const im = infoAba(a.getName()), blocos = [];
  const v = a.getRange(1, 1, a.getLastRow(), COLS).getDisplayValues();
  let b = null;
  v.forEach((l, i) => {
    if (i < 3) return;
    if (l[0] === 'Horário') {
      b = { linha: i + 1, slots: [], dias: DIAS.map((_, d) => {
        const m = String(l[pc(d) - 1]).match(/(\d{2})\/(\d{2})/);
        return m ? { a: im.ano, m: Number(m[2]), d: Number(m[1]) } : null;
      }) };
      blocos.push(b);
      return;
    }
    const h = lerHora(l[0]);
    if (b && h !== null) b.slots.push({ linha: i + 1, min: h, v: l });
  });
  blocos.forEach(bl => bl.slots.forEach((s, k) => {
    const prox = bl.slots[k + 1];
    s.fim = prox && prox.min > s.min ? prox.min : s.min + dur;
  }));
  return blocos;
}

/** Pinta o mês: cor da psicóloga no nome, cor do status, feriados, bloqueios e filtro. */
function pintar(ss, a) {
  const c = lerCfg(ss), P = lerPsis(ss), fer = lerFeriados(ss);
  const filtro = a.getRange('B2').getDisplayValue().trim().toUpperCase();
  const pf = filtro && filtro !== 'TODAS' ? P.por[filtro] : null;
  const comHorario = P.lista.filter(p => p.tem);

  lerMes(a, c.dur).forEach(b => {
    const cab = ['Horário'], fundoCab = [COR.AZUL];
    b.dias.forEach((x, d) => {
      const f = x && fer[chave(x)];
      cab.push(x ? DIAS[d] + ' ' + dm(x.d, x.m) + (f ? ' · ' + f : '') : '—', '');
      fundoCab.push(...Array(2).fill(!x ? '#999999' : f ? '#7E57C2' : COR.AZUL));
    });
    a.getRange(b.linha, 1, 1, COLS).setValues([cab]).setBackgrounds([fundoCab]);
    if (!b.slots.length) return;

    const bg = [], fc = [], fl = [], notas = [];
    b.slots.forEach(s => {
      const rb = [], rf = [], rl = [], rn = [];
      b.dias.forEach((x, d) => {
        const txt = s.v[pc(d) - 1].trim(), st = s.v[pc(d)].trim();
        if (!x) { rb.push(COR.FORA, COR.FORA); rf.push('#BBBBBB', '#BBBBBB'); rl.push('none', 'none'); rn.push('', ''); return; }
        const livre = pf ? atende(pf, d, s.min, s.fim)
          : !comHorario.length || comHorario.some(p => atende(p, d, s.min, s.fim));
        const base = fer[chave(x)] ? COR.FERIADO : livre ? '#FFFFFF' : COR.BLOQ;
        if (!txt) { rb.push(base, base); rf.push('black', 'black'); rl.push('none', 'none'); rn.push('', ''); return; }

        const sg = siglas(txt), corSt = COR[st] || '#FFFFFF';
        const corPsi = sg.length && P.por[sg[0]] ? P.por[sg[0]].cor : corSt;
        const avisos = sg.filter(k => !P.por[k]).map(k => 'Sigla ' + k + ' não cadastrada')
          .concat(sg.filter(k => P.por[k] && !atende(P.por[k], d, s.min, s.fim)).map(k => k + ' não atende neste horário'));
        if (!sg.length) avisos.push('Falta a sigla da psicóloga. Ex.: Maria (AN)');
        if (pf && sg.indexOf(filtro) < 0) {
          rb.push(COR.APAGADO, COR.APAGADO); rf.push('#BBBBBB', '#BBBBBB'); rl.push('none', 'none');
        } else {
          const cancel = st === 'Cancelou';
          rb.push(corPsi, corSt); rf.push(cancel ? '#777777' : 'black', 'black'); rl.push(cancel ? 'line-through' : 'none', 'none');
        }
        rn.push(avisos.length ? '⚠ ' + avisos.join('\n⚠ ') : '', '');
      });
      bg.push(rb); fc.push(rf); fl.push(rl); notas.push(rn);
    });
    a.getRange(b.slots[0].linha, 2, b.slots.length, COLS - 1)
      .setBackgrounds(bg).setFontColors(fc).setFontLines(fl).setNotes(notas);
  });
}


// ============================ AUTOMÁTICO ============================

function onEdit(e) {
  if (!e || !e.range) return;
  const ss = e.source, a = e.range.getSheet(), nome = a.getName();
  if (nome === ABA.RESUMO) return resumo(ss);
  if (!infoAba(nome)) return;

  const r1 = e.range.getRow(), r2 = e.range.getLastRow(), c1 = e.range.getColumn(), c2 = e.range.getLastColumn();
  if (r2 <= 3) { if (r1 <= 2 && r2 >= 2 && c1 <= 2) pintar(ss, a); return; }   // filtro "Ver"

  const c = lerCfg(ss), ref = lerPacientes(ss);
  const umaCelula = r1 === r2 && c1 === c2;
  let pedirMotivo = false;

  lerMes(a, c.dur).forEach(b => b.slots.forEach(s => {
    if (s.linha < r1 || s.linha > r2) return;
    b.dias.forEach((x, d) => {
      const col = pc(d);
      if (col + 1 < c1 || col > c2) return;
      const cel = a.getRange(s.linha, col, 1, 2);
      let txt = s.v[col - 1].trim(), st = s.v[col].trim();
      if (!x) { if (txt || st) { cel.clearContent(); ss.toast('Esse dia é de outro mês.'); } return; }
      if (!txt) { if (st) cel.clearContent(); return; }

      // Trocou o paciente de um horário que estava cancelado/faltou → novo agendamento
      const trocou = umaCelula && c1 === col && e.oldValue && nomePac(String(e.oldValue)) !== nomePac(txt);
      if (!st || (trocou && (st === 'Faltou' || st === 'Cancelou'))) st = 'Agendada';
      const sigla = ref[nomePac(txt).toLowerCase()];
      if (!siglas(txt).length && sigla) txt = nomePac(txt) + ' (' + sigla + ')';
      if (txt !== s.v[col - 1].trim() || st !== s.v[col].trim()) cel.setValues([[txt, st]]);

      if (st === 'Faltou' || st === 'Cancelou') pedirMotivo = registrar(ss, x, s.min, txt, st) || pedirMotivo;
      else desregistrar(ss, x, s.min, txt);
    });
  }));

  pintar(ss, a);
  if (ss.getSheetByName(ABA.RESUMO).getRange('B1').getDisplayValue() === nome) resumo(ss);
  if (pedirMotivo) ss.toast('Escolha o motivo na aba "Faltas e cancelamentos".', 'Falta / cancelamento', 8);
}

/** Guarda falta/cancelamento. Devolve true se ainda falta o motivo. */
function registrar(ss, x, min, txt, st) {
  const f = ss.getSheetByName(ABA.FALTAS), data = dm(x.d, x.m) + '/' + x.a, hora = hhmm(min);
  const i = acharRegistro(f, data, hora, txt);
  if (i) {
    f.getRange(i, 3, 1, 2).setValues([[txt, st]]);
    return !f.getRange(i, 5).getDisplayValue();
  }
  f.appendRow([data, hora, txt, st, '', fmt(new Date(), 'dd/MM/yyyy HH:mm')]);
  return true;
}

function desregistrar(ss, x, min, txt) {
  const f = ss.getSheetByName(ABA.FALTAS);
  const i = acharRegistro(f, dm(x.d, x.m) + '/' + x.a, hhmm(min), txt);
  if (i) f.deleteRow(i);
}

function acharRegistro(f, data, hora, txt) {
  if (f.getLastRow() < 2) return 0;
  const v = f.getRange(2, 1, f.getLastRow() - 1, 3).getDisplayValues();
  const p = nomePac(txt).toLowerCase();
  const k = v.findIndex(l => l[0] === data && l[1] === hora && nomePac(l[2]).toLowerCase() === p);
  return k < 0 ? 0 : k + 2;
}


// ============================== RESUMO ==============================

function resumo(ss) {
  const R = ss.getSheetByName(ABA.RESUMO);
  if (!R) return;
  const nome = R.getRange('B1').getDisplayValue(), a = ss.getSheetByName(nome), im = infoAba(nome);
  R.getRange(3, 1, R.getMaxRows() - 2, 6).clear();
  if (!a || !im) { R.getRange('A3').setValue('Escolha o mês em B1.'); return; }

  const P = lerPsis(ss), n = { Agendada: 0, Confirmada: 0, Compareceu: 0, Faltou: 0, Cancelou: 0 };
  const por = {}, motivos = {}, faltas = {};
  let equipe = 0, semMotivo = 0;
  const pp = k => por[k] = por[k] || [0, 0, 0, 0, 0];   // compareceu, faltou, cancelou, em equipe, a vir

  lerMes(a, lerCfg(ss).dur).forEach(b => b.slots.forEach(s => b.dias.forEach((x, d) => {
    const txt = s.v[pc(d) - 1].trim(), st = s.v[pc(d)].trim();
    if (!x || !txt || st === 'Faltou' || st === 'Cancelou') return;     // essas vêm da aba de faltas
    const sg = siglas(txt);
    n[st] = (n[st] || 0) + 1;
    sg.forEach(k => pp(k)[st === 'Compareceu' ? 0 : 4]++);
    if (sg.length > 1) { equipe++; sg.forEach(k => pp(k)[3]++); }
  })));

  const f = ss.getSheetByName(ABA.FALTAS), fim = '/' + ('0' + im.mes).slice(-2) + '/' + im.ano;
  (f.getLastRow() > 1 ? f.getRange(2, 1, f.getLastRow() - 1, 5).getDisplayValues() : []).forEach(l => {
    if (!l[0].endsWith(fim) || !n.hasOwnProperty(l[3])) return;
    n[l[3]]++;
    siglas(l[2]).forEach(k => pp(k)[l[3] === 'Faltou' ? 1 : 2]++);
    const m = l[4] || '(sem motivo)';
    motivos[m] = (motivos[m] || 0) + 1;
    if (!l[4]) semMotivo++;
    if (l[3] === 'Faltou') faltas[nomePac(l[2])] = (faltas[nomePac(l[2])] || 0) + 1;
  });

  const total = STATUS.reduce((t, s) => t + n[s], 0);
  const pct = (x, y) => y ? Math.round(100 * x / y) + '%' : '0%';
  const ordem = o => Object.keys(o).sort((x, y) => o[y] - o[x]);
  let r = 3;
  r = secao(R, r, ['TOTAIS DO MÊS', ''], [
    ['Consultas marcadas', total], ['Compareceram', n.Compareceu], ['Confirmadas (ainda vão vir)', n.Confirmada],
    ['Agendadas sem confirmação', n.Agendada], ['Faltas', n.Faltou], ['Cancelamentos', n.Cancelou],
    ['Sessões em equipe', equipe], ['Taxa de faltas', pct(n.Faltou, n.Compareceu + n.Faltou)],
    ['Taxa de cancelamento', pct(n.Cancelou, total)], ['Faltas/cancelamentos sem motivo', semMotivo]]);
  const siglasPsi = P.lista.map(p => p.sigla).concat(Object.keys(por).filter(k => !P.por[k]));
  r = secao(R, r, ['POR PSICÓLOGA', 'Compareceu', 'Faltou', 'Cancelou', 'Em equipe', 'A vir'],
    siglasPsi.map(k => [(P.por[k] ? P.por[k].nome + ' ' : '') + '(' + k + ')'].concat(por[k] || [0, 0, 0, 0, 0])));
  r = secao(R, r, ['MOTIVOS (faltas e cancelamentos)', 'Vezes'], ordem(motivos).map(m => [m, motivos[m]]));
  secao(R, r, ['QUEM MAIS FALTOU', 'Faltas'], ordem(faltas).slice(0, 5).map(p => [p, faltas[p]]));
}

function secao(R, r, cab, linhas) {
  R.getRange(r, 1, 1, cab.length).setValues([cab]).setFontWeight('bold').setBackground(COR.AZUL).setFontColor('white');
  if (!linhas.length) linhas = [['Nada no mês'].concat(Array(cab.length - 1).fill(''))];
  R.getRange(r + 1, 1, linhas.length, cab.length).setValues(linhas).setHorizontalAlignment('left');
  return r + linhas.length + 2;
}


// ========================== PACIENTE FIXO ==========================

function repetirFixo() { fixo(true); }
function tirarFixo() { fixo(false); }

function fixo(criar) {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), ui = SpreadsheetApp.getUi();
  const a = ss.getActiveSheet(), cel = a.getActiveCell(), dur = lerCfg(ss).dur;
  const d = Math.floor((cel.getColumn() - 2) / 2);
  let b0, s0;
  if (infoAba(a.getName()) && d >= 0 && d < 6) {
    lerMes(a, dur).forEach(b => b.slots.forEach(s => { if (s.linha === cel.getRow()) { b0 = b; s0 = s; } }));
  }
  if (!s0 || !b0.dias[d] || !s0.v[pc(d) - 1].trim()) {
    return ui.alert('Clique no nome do paciente no calendário e escolha a opção de novo.');
  }
  const txt = s0.v[pc(d) - 1].trim(), x = b0.dias[d], base = Date.UTC(x.a, x.m - 1, x.d);
  let vezes = 52, passo = 1;
  if (criar) {
    vezes = numero(ui, nomePac(txt) + ' — ' + DIAS[d] + ' ' + hhmm(s0.min) + '\n\nQuantas vezes repetir? (ex.: 8)');
    if (!vezes) return;
    passo = numero(ui, 'De quantas em quantas semanas?\n1 = toda semana · 2 = a cada 15 dias');
    if (!passo) return;
    vezes = Math.min(vezes, 52);
  } else if (ui.alert('Tirar ' + nomePac(txt) + ' de ' + DIAS[d] + ' ' + hhmm(s0.min) +
    ' das próximas semanas?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;

  const feitos = [], pulados = [], abas = {};
  for (let k = 1; k <= vezes; k++) {
    const dt = new Date(base + 7 * k * passo * 864e5), ano = dt.getUTCFullYear(), mes = dt.getUTCMonth() + 1, dia = dt.getUTCDate();
    let alvo = ss.getSheetByName(MESES[mes - 1] + ' ' + ano);
    if (!alvo) { if (!criar) break; alvo = criarMes(ss, ano, mes); }
    let s = null;
    lerMes(alvo, dur).forEach(b => { if (b.dias[d] && b.dias[d].d === dia) s = b.slots.find(z => z.min === s0.min) || s; });
    const quando = dm(dia, mes);
    if (!s) { if (criar) pulados.push(quando + ' (horário não existe)'); continue; }
    const atual = s.v[pc(d) - 1].trim(), rg = alvo.getRange(s.linha, pc(d), 1, 2);
    if (criar) {
      if (atual) { pulados.push(quando + ' (ocupado: ' + atual + ')'); continue; }
      rg.setValues([[txt, 'Agendada']]);
    } else {
      if (nomePac(atual).toLowerCase() !== nomePac(txt).toLowerCase()) continue;
      rg.clearContent();
    }
    feitos.push(quando);
    abas[alvo.getName()] = alvo;
  }
  Object.keys(abas).forEach(n => pintar(ss, abas[n]));
  listaMeses(ss);
  resumo(ss);
  ui.alert((criar ? 'Marcado em: ' : 'Retirado de: ') + (feitos.join(', ') || 'nenhuma data') +
    (pulados.length ? '\n\n⚠ Não marcado:\n' + pulados.join('\n') : ''));
}


// ============================ MANUTENÇÃO ============================

function atualizarTudo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.getSheets().filter(a => infoAba(a.getName())).forEach(a => { atualizarFiltro(ss, a); pintar(ss, a); });
  listaMeses(ss);
  resumo(ss);
  ss.toast('Cores, feriados e resumo atualizados. ✅');
}

function atualizarFiltro(ss, a) {
  a.getRange('B2').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(['Todas'].concat(lerPsis(ss).lista.map(p => p.sigla)), true).build());
}

function listaMeses(ss) {
  const R = ss.getSheetByName(ABA.RESUMO);
  const nomes = ss.getSheets().map(a => a.getName()).filter(infoAba);
  if (!R || !nomes.length) return;
  R.getRange('B1').setNumberFormat('@')   // texto: senão o Planilhas transforma "Outubro 2026" em data
    .setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(nomes, true).build());
  if (nomes.indexOf(R.getRange('B1').getDisplayValue()) < 0) {
    const atual = MESES[Number(fmt(new Date(), 'M')) - 1] + ' ' + fmt(new Date(), 'yyyy');
    R.getRange('B1').setValue(nomes.indexOf(atual) >= 0 ? atual : nomes[0]);
  }
}

function ordenarAbas(ss) {
  const meses = ss.getSheets().map(a => a.getName()).filter(infoAba)
    .sort((x, y) => { const a = infoAba(x), b = infoAba(y); return (a.ano - b.ano) || (a.mes - b.mes); });
  const ativa = ss.getActiveSheet();
  let pos = 1;
  [ABA.RESUMO].concat(meses, [ABA.PAC, ABA.FALTAS, ABA.PSI, ABA.FER, ABA.CFG]).forEach(n => {
    const a = ss.getSheetByName(n);
    if (!a) return;
    ss.setActiveSheet(a);
    ss.moveActiveSheet(pos++);
  });
  ss.setActiveSheet(ativa);
}

function garantirFeriados(ss, ano) {
  const f = ss.getSheetByName(ABA.FER);
  const v = f.getLastRow() > 1 ? f.getRange(2, 1, f.getLastRow() - 1, 1).getDisplayValues() : [];
  if (v.some(l => l[0].endsWith('/' + ano))) return;
  const p = pascoa(ano), mov = (n, nome) => { const t = new Date(p + n * 864e5); return [t.getUTCDate(), t.getUTCMonth() + 1, nome]; };
  const lista = [[1, 1, 'Confraternização Universal'], [21, 4, 'Tiradentes'], [1, 5, 'Dia do Trabalho'],
    [7, 9, 'Independência'], [12, 10, 'N. Sra. Aparecida'], [2, 11, 'Finados'], [15, 11, 'Proclamação da República'],
    [20, 11, 'Consciência Negra'], [25, 12, 'Natal'],
    mov(-48, 'Carnaval'), mov(-47, 'Carnaval'), mov(-2, 'Sexta-feira Santa'), mov(60, 'Corpus Christi')]
    .sort((x, y) => (x[1] - y[1]) || (x[0] - y[0]));
  f.getRange(f.getLastRow() + 1, 1, lista.length, 2).setNumberFormat('@')
    .setValues(lista.map(x => [dm(x[0], x[1]) + '/' + ano, x[2]]));
}


// ============================== LEITURA ==============================

function lerCfg(ss) {
  const v = ss.getSheetByName(ABA.CFG).getRange('B2:B4').getDisplayValues();
  const ini = lerHora(v[0][0]), fim = lerHora(v[1][0]);
  return { ini: ini === null ? 480 : ini, fim: fim === null ? 1200 : fim, dur: Math.max(Number(v[2][0]) || 60, 10) };
}

function lerPsis(ss) {
  const a = ss.getSheetByName(ABA.PSI), n = a.getLastRow() - 1, lista = [], por = {};
  if (n < 1) return { lista, por };
  const v = a.getRange(2, 1, n, 9).getDisplayValues(), cores = a.getRange(2, 3, n, 1).getBackgrounds();
  v.forEach((l, i) => {
    const sg = l[1].trim().toUpperCase();
    if (!sg) return;
    const horas = l.slice(3, 9).map(t => String(t).split(/[,;]/).map(faixa => {
      const m = faixa.match(/(\d{1,2})[:h](\d{2})\s*(?:-|a|às|até)\s*(\d{1,2})[:h](\d{2})/);
      return m ? [Number(m[1]) * 60 + Number(m[2]), Number(m[3]) * 60 + Number(m[4])] : null;
    }).filter(Boolean));
    const p = { nome: l[0].trim() || sg, sigla: sg, cor: cores[i][0], horas: horas, tem: horas.some(h => h.length) };
    lista.push(p);
    por[sg] = p;
  });
  return { lista, por };
}

function atende(p, d, ini, fim) {
  return !p.tem || p.horas[d].some(h => ini >= h[0] && fim <= h[1]);
}

function lerPacientes(ss) {
  const a = ss.getSheetByName(ABA.PAC), mapa = {};
  if (a.getLastRow() < 2) return mapa;
  a.getRange(2, 1, a.getLastRow() - 1, 4).getDisplayValues().forEach(l => {
    if (l[0].trim() && l[3].trim()) mapa[l[0].trim().toLowerCase()] = l[3].trim().toUpperCase();
  });
  return mapa;
}

function lerFeriados(ss) {
  const f = ss.getSheetByName(ABA.FER), mapa = {};
  if (f.getLastRow() < 2) return mapa;
  f.getRange(2, 1, f.getLastRow() - 1, 2).getDisplayValues().forEach(l => {
    const m = l[0].match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
    if (m) mapa[dm(Number(m[1]), Number(m[2])) + '/' + m[3]] = l[1];
  });
  return mapa;
}


// ============================== AJUDA ==============================

function infoAba(nome) {
  const m = String(nome).match(new RegExp('^(' + MESES.join('|') + ') (\\d{4})$'));
  return m ? { mes: MESES.indexOf(m[1]) + 1, ano: Number(m[2]) } : null;
}

function siglas(txt) {
  const m = String(txt).match(/\(([^)]*)\)\s*$/);
  return m ? m[1].toUpperCase().split(/[^A-Z0-9À-Ü]+/).filter(String) : [];
}

function nomePac(txt) { return String(txt).replace(/\s*\([^)]*\)\s*$/, '').trim(); }

function lerHora(t) {
  const m = String(t).match(/^\s*(\d{1,2})[:h](\d{2})/);
  return m && Number(m[1]) < 24 && Number(m[2]) < 60 ? Number(m[1]) * 60 + Number(m[2]) : null;
}

function hhmm(min) { return ('0' + Math.floor(min / 60)).slice(-2) + ':' + ('0' + (min % 60)).slice(-2); }
function dm(d, m) { return ('0' + d).slice(-2) + '/' + ('0' + m).slice(-2); }
function chave(x) { return dm(x.d, x.m) + '/' + x.a; }
function linha(arr) { return arr.concat(Array(COLS - arr.length).fill('')); }
function fmt(data, padrao) { return Utilities.formatDate(data, 'America/Sao_Paulo', padrao); }

function numero(ui, msg) {
  const r = ui.prompt('🗓️ Consultório', msg, ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return 0;
  const n = parseInt(r.getResponseText(), 10);
  if (!(n > 0)) { ui.alert('Digite um número maior que zero.'); return 0; }
  return n;
}

/** Data da Páscoa (para Carnaval, Sexta-feira Santa e Corpus Christi). */
function pascoa(ano) {
  const a = ano % 19, b = Math.floor(ano / 100), c = ano % 100, d = Math.floor(b / 4), e = b % 4;
  const f = Math.floor((b + 8) / 25), g = Math.floor((b - f + 1) / 3), h = (19 * a + b - d - g + 15) % 30;
  const i = Math.floor(c / 4), k = c % 4, l = (32 + 2 * e + 2 * i - h - k) % 7, m = Math.floor((a + 11 * h + 22 * l) / 451);
  const mes = Math.floor((h + l - 7 * m + 114) / 31), dia = ((h + l - 7 * m + 114) % 31) + 1;
  return Date.UTC(ano, mes - 1, dia);
}


// ======================= DADOS DE EXEMPLO (PORTFÓLIO) =======================

/** Preenche psicólogas, pacientes, consultas e faltas FICTÍCIOS para demonstração. */
function preencherExemplo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), ui = SpreadsheetApp.getUi();
  if (ui.alert('Dados de exemplo', 'Isso coloca psicólogas, pacientes e consultas FICTÍCIOS.\n' +
    'Use só na cópia do portfólio. Continuar?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;

  ss.getSheetByName(ABA.PSI).getRange('A2:I5').setValues([
    ['Dra. Helena Prado', 'HP', '', '08:00-12:00, 14:00-18:00', '08:00-12:00', '14:00-20:00', '08:00-12:00, 14:00-18:00', '08:00-12:00', ''],
    ['Dra. Bianca Torres', 'BT', '', '', '14:00-20:00', '08:00-12:00', '14:00-20:00', '08:00-12:00, 14:00-18:00', '08:00-12:00'],
    ['Dr. Rafael Moura', 'RM', '', '08:00-12:00', '08:00-12:00', '', '08:00-12:00', '14:00-18:00', '08:00-12:00'],
    ['Dra. Camila Neves', 'CN', '', '14:00-20:00', '', '08:00-12:00, 14:00-18:00', '', '14:00-20:00', '']]);

  const pacs = [['Ana Souza', 'HP'], ['Bruno Lima', 'HP'], ['Carla Mendes', 'HP'], ['Diego Alves', 'BT'],
    ['Eduarda Rocha', 'BT'], ['Felipe Costa', 'BT'], ['Gabriela Dias', 'RM'], ['Henrique Melo', 'RM'],
    ['Isabela Freitas', 'RM'], ['João Pereira', 'CN'], ['Larissa Gomes', 'CN'], ['Marcos Ribeiro', 'CN']];
  ss.getSheetByName(ABA.PAC).getRange(2, 1, pacs.length, 5).setValues(pacs.map((p, i) => [p[0],
    '(11) 90000-00' + ('0' + (i + 1)).slice(-2), i === 2 || i === 8 ? 'Responsável Fictício' : '', p[1],
    'Dado fictício para demonstração']));

  let semente = 7;
  const rnd = () => (semente = (semente * 16807) % 2147483647) / 2147483647;
  const P = lerPsis(ss), fer = lerFeriados(ss), dur = lerCfg(ss).dur;
  const hoje = Number(fmt(new Date(), 'yyyyMMdd')), agora = fmt(new Date(), 'dd/MM/yyyy HH:mm');
  const log = [], usados = {};

  ss.getSheets().filter(a => infoAba(a.getName())).forEach(a => lerMes(a, dur).forEach(b => {
    if (!b.slots.length) return;
    const vals = b.slots.map(s => s.v.slice(1, COLS));
    b.slots.forEach((s, i) => b.dias.forEach((x, d) => {
      if (!x || fer[chave(x)] || vals[i][2 * d] || rnd() > 0.45) return;
      const livres = P.lista.filter(p => p.tem && atende(p, d, s.min, s.fim));
      if (!livres.length) return;
      const p = livres[Math.floor(rnd() * livres.length)];
      const opc = pacs.filter(q => q[1] === p.sigla);
      const nome = opc[Math.floor(rnd() * opc.length)][0];
      if (usados[chave(x) + nome]) return;                 // mesmo paciente 1 vez por dia
      usados[chave(x) + nome] = true;
      let sg = p.sigla;
      if (livres.length > 1 && rnd() < 0.12) sg += '+' + livres.find(q => q !== p).sigla;   // em equipe
      const txt = nome + ' (' + sg + ')', r = rnd();
      const st = x.a * 10000 + x.m * 100 + x.d < hoje
        ? (r < 0.75 ? 'Compareceu' : r < 0.87 ? 'Faltou' : 'Cancelou')
        : (r < 0.45 ? 'Confirmada' : 'Agendada');
      vals[i][2 * d] = txt;
      vals[i][2 * d + 1] = st;
      if (st === 'Faltou' || st === 'Cancelou') {
        log.push([chave(x), hhmm(s.min), txt, st, log.length % 6 === 5 ? '' : MOTIVOS[Math.floor(rnd() * MOTIVOS.length)], agora]);
      }
    }));
    a.getRange(b.slots[0].linha, 2, vals.length, COLS - 1).setValues(vals);
  }));

  const f = ss.getSheetByName(ABA.FALTAS);
  if (log.length) f.getRange(f.getLastRow() + 1, 1, log.length, 6).setValues(log);
  atualizarTudo();
  ui.alert('Pronto! ✅ Dados fictícios preenchidos.\n\nVeja as abas dos meses e o Resumo.\n' +
    'Para o portfólio, compartilhe como "Leitor".');
}
