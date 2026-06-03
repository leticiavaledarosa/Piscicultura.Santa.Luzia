// ============================================================
// AquaLuzia — Google Apps Script v4
// Piscicultura Santa Luzia — COMPLETO
// ============================================================

const CHAVE_SECRETA = "AQUALUZIA2026";
const PASTA_FOTOS_ID = "1dCbTHQVZEL-ncLN5461kbWPJ8TP5AErq";

function doPost(e) {
  try {
    let payload;
    if (e.postData && e.postData.type === 'application/json') {
      payload = JSON.parse(e.postData.contents);
    } else if (e.parameter && e.parameter.payload) {
      payload = JSON.parse(e.parameter.payload);
    } else if (e.postData && e.postData.contents) {
      payload = JSON.parse(e.postData.contents);
    } else {
      return resposta("erro", "Sem dados");
    }
    return processar(payload);
  } catch (err) {
    return resposta("erro", err.message);
  }
}

function doGet(e) {
  // Rota para relatório diário
  if (e.parameter && e.parameter.tipo === 'relatorio') {
    try {
      const data = e.parameter.data || Utilities.formatDate(
        new Date(Date.now() - 86400000), "America/Sao_Paulo", "yyyy-MM-dd"
      );
      return gerarRelatorio(data);
    } catch(err) {
      return resposta("erro", err.message);
    }
  }
  // Rota normal com payload
  if (e.parameter && e.parameter.payload) {
    try {
      const payload = JSON.parse(e.parameter.payload);
      return processar(payload);
    } catch(err) {
      return resposta("erro", err.message);
    }
  }
  return resposta("ok", "AquaLuzia online!");
}

// ===================== RELATÓRIO DIÁRIO =====================
function gerarRelatorio(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const resultado = { data, gerado: new Date().toISOString() };

  // Alimentação
  try {
    const abaAli = ss.getSheetByName("Alimentação Diária");
    if (abaAli) {
      const rows = abaAli.getDataRange().getValues();
      const header = rows[0];
      const dadosDia = rows.slice(1).filter(r => r[0] && r[0].toString().includes(data));
      
      resultado.racao_total = dadosDia.reduce((s,r) => s + (parseFloat(r[header.indexOf('Kg Fornecido')])||0), 0).toFixed(1);
      resultado.mortes_total = dadosDia.reduce((s,r) => s + (parseInt(r[header.indexOf('Mortes')])||0), 0);
      resultado.arr_count = dadosDia.length;
      resultado.gaiolas_ok = [...new Set(dadosDia.map(r => r[header.indexOf('Gaiola')]))].length;
      
      // Por setor
      resultado.setores = {};
      ['A','B','C','D','E'].forEach(s => {
        const reg = dadosDia.filter(r => r[header.indexOf('Setor')] === s);
        resultado.setores[s] = {
          racao: reg.reduce((sum,r) => sum+(parseFloat(r[header.indexOf('Kg Fornecido')])||0), 0).toFixed(1),
          tratos: reg.length,
          gaiolas: [...new Set(reg.map(r => r[header.indexOf('Gaiola')]))].length,
          mortes: reg.reduce((sum,r) => sum+(parseInt(r[header.indexOf('Mortes')])||0), 0)
        };
      });

      // CA estimado
      if(parseFloat(resultado.racao_total) > 0){
        resultado.ca_estimado = (parseFloat(resultado.racao_total) / 
          Math.max(1, parseFloat(resultado.racao_total) * 0.7)).toFixed(2);
      }
    }
  } catch(e) { resultado.ali_erro = e.message; }

  // Hídrico
  try {
    const abaH = ss.getSheetByName("Hídrico");
    if (abaH) {
      const rows = abaH.getDataRange().getValues();
      const header = rows[0];
      const dadosDia = rows.slice(1).filter(r => r[0] && r[0].toString().includes(data));
      
      if (dadosDia.length > 0) {
        const odIdx = header.indexOf('OD (mg/L)');
        const phIdx = header.indexOf('pH');
        const tempIdx = header.indexOf('Temp Água (°C)');
        const tempArIdx = header.indexOf('Temp Ar (°C)');
        const corIdx = header.indexOf('Cor da Água');
        const localIdx = header.indexOf('Local');
        
        const ods = dadosDia.map(r => parseFloat(r[odIdx])).filter(v => !isNaN(v));
        const phs = dadosDia.map(r => parseFloat(r[phIdx])).filter(v => !isNaN(v));
        const temps = dadosDia.map(r => parseFloat(r[tempIdx])).filter(v => !isNaN(v));
        
        resultado.hidrico = {
          od_medio: ods.length ? (ods.reduce((a,b)=>a+b,0)/ods.length).toFixed(1) : '--',
          ph_medio: phs.length ? (phs.reduce((a,b)=>a+b,0)/phs.length).toFixed(1) : '--',
          temp_agua: temps.length ? (temps.reduce((a,b)=>a+b,0)/temps.length).toFixed(1) : '--',
          temp_ar: dadosDia[0] ? (dadosDia[0][tempArIdx]||'--').toString() : '--',
          pontos: dadosDia.map(r => ({
            local: r[localIdx]||'',
            od: r[odIdx]||'',
            ph: r[phIdx]||'',
            temp: r[tempIdx]||'',
            cor: r[corIdx]||''
          }))
        };
        
        const tAgua = parseFloat(resultado.hidrico.temp_agua);
        const tAr = parseFloat(resultado.hidrico.temp_ar);
        if (!isNaN(tAgua) && !isNaN(tAr)) {
          resultado.hidrico.diff_temp = (tAgua - tAr).toFixed(1);
        }
      }
    }
  } catch(e) { resultado.h_erro = e.message; }

  // Biometrias
  try {
    const abaBio = ss.getSheetByName("Biometria");
    if (abaBio) {
      const rows = abaBio.getDataRange().getValues();
      const header = rows[0];
      const dadosDia = rows.slice(1).filter(r => r[0] && r[0].toString().includes(data));
      resultado.biometrias = dadosDia.map(r => ({
        gaiola: r[header.indexOf('Gaiola')]||'',
        peso: r[header.indexOf('Peso Médio (g)')]||'',
        amostras: r[header.indexOf('Nº Amostras')]||'',
        gpd: r[header.indexOf('Ganho Diário (g/dia)')]||'',
        ca: r[header.indexOf('CA Acumulado')]||'',
        fator_k: r[header.indexOf('Fator K')]||''
      }));
    }
  } catch(e) { resultado.bio_erro = e.message; }

  // Despescas
  try {
    const abaDesp = ss.getSheetByName("Despescas");
    if (abaDesp) {
      const rows = abaDesp.getDataRange().getValues();
      const header = rows[0];
      const dadosDia = rows.slice(1).filter(r => r[0] && r[0].toString().includes(data));
      resultado.despescas = dadosDia.map(r => ({
        gaiola: r[header.indexOf('Gaiola')]||'',
        biomassa: r[header.indexOf('Biomassa (kg)')]||'',
        ca_final: r[header.indexOf('CA Final')]||'',
        sobrevivencia: r[header.indexOf('Sobrevivência (%)')]||'',
        destino: r[header.indexOf('Destino')]||''
      }));
    }
  } catch(e) { resultado.desp_erro = e.message; }

  // Diagnósticos
  try {
    const abaDiag = ss.getSheetByName("Diagnósticos");
    if (abaDiag) {
      const rows = abaDiag.getDataRange().getValues();
      const header = rows[0];
      const dadosDia = rows.slice(1).filter(r => r[0] && r[0].toString().includes(data));
      resultado.diagnosticos = dadosDia.map(r => ({
        gaiola: r[header.indexOf('Gaiola')]||'',
        sintomas: r[header.indexOf('Sintomas')]||'',
        diagnostico: (r[header.indexOf('Diagnóstico IA')]||'').substring(0, 100)
      }));
    }
  } catch(e) { resultado.diag_erro = e.message; }

  return ContentService
    .createTextOutput(JSON.stringify(resultado))
    .setMimeType(ContentService.MimeType.JSON);
}

function processar(payload) {
  if (payload.chave !== CHAVE_SECRETA) return resposta("erro", "Chave inválida");
  const tipo = payload.tipo;
  const data = payload.data;
  if (tipo === "gaiola")    salvarGaiola(data);
  if (tipo === "hidrico")   salvarHidrico(data);
  if (tipo === "biometria") salvarBiometria(data);
  if (tipo === "foto")      salvarFoto(data);
  if (tipo === "lote")      salvarLote(data);
  if (tipo === "despesca")  salvarDespesca(data);
  if (tipo === "diagnostico") salvarDiagnostico(data);
  return resposta("ok", "Salvo!");
}

// ===================== ALIMENTAÇÃO DIÁRIA =====================
function salvarGaiola(d) {
  const aba = getOuCriarAba("Alimentação Diária", [
    "Data","Gaiola","Setor","Arraç. Nº","Horário",
    "Clima","Temp Ar (°C)",
    "Fase do Lote","Qualidade Visual Água",
    "Tratador 1","Tratador 2",
    "Tipo Ração","Kg Fornecido","Consumo","Sobra (kg)",
    "Comportamento","Lesões","Parasitas","Predador","Coloração",
    "Mortes","Causa Morte",
    "GPS","Observações","Foto"
  ]);
  aba.appendRow([
    d.data||"", d.cage_id||"", d.setor||"", d.arr_num||1, d.hora||"",
    d.clima||"", d.temp_ar||"",
    d.fase||"", d.agua_visual||"",
    d.trat1||"", d.trat2||"",
    d.racao_tipo||"", d.kg||0, d.consumo||"", d.sobra||0,
    d.comportamento||"", d.mortes||0, d.causa||"",
    d.gps||"", d.obs||"", d.foto||""
  ]);
}

// ===================== HÍDRICO — 1 linha por ponto =====================
function salvarHidrico(d) {
  const aba = getOuCriarAba("Hídrico", [
    "Data","Horário","Responsável","Clima","Temp Ar (°C)",
    "Ponto","Local",
    "Temp Água (°C)","OD (mg/L)","pH",
    "Secchi (cm)","Amônia (mg/L)","Nitrito (mg/L)","Nível (m)",
    "Cor da Água","Floração Algas","Peixes Superfície",
    "GPS","Observações"
  ]);
  
  const locais = ["Entrada","Meio","Saída"];
  const pontos = d.pontos || [{},{},{}];
  const pv = (pt, k) => pt && pt[k] !== undefined && pt[k] !== null ? pt[k] : "";
  
  // Uma linha por ponto
  pontos.forEach((p, idx) => {
    aba.appendRow([
      d.data||"", d.hora||"", d.resp||"", d.clima||"", d.temp_ar||"",
      idx + 1, locais[idx] || ("Ponto " + (idx+1)),
      pv(p,'temp'), pv(p,'od'), pv(p,'ph'),
      pv(p,'secchi'), pv(p,'amonia'), pv(p,'nitrito'), pv(p,'nivel'),
      pv(p,'cor'), pv(p,'algas'), pv(p,'sup'),
      pv(p,'gps'), pv(p,'obs')
    ]);
  });
}

// ===================== BIOMETRIA =====================
function salvarBiometria(d) {
  const aba = getOuCriarAba("Biometria", [
    "Data","Gaiola","Horário","Responsável","Fase","Clima",
    "Nº Amostras","Peso Médio (g)","Menor (g)","Maior (g)",
    "Desvio Padrão (g)","Comprimento (cm)",
    "Condição Corporal","Coloração","Comportamento",
    "Lesões","Parasitas","Manejo",
    "Fator K","Biomassa Est. (kg)","CA Acumulado","Ganho Diário (g/dia)",
    "Data Povoamento","Qtd Inicial","Peso Inicial (g)","Ração Acum. (kg)",
    "Observações"
  ]);
  // Calcula indicadores automaticamente
  const peso = parseFloat(d.peso) || 0;
  const tam = parseFloat(d.tam) || 0;
  const qtd = parseInt(d.lote_qtd) || 0;
  const pi = parseFloat(d.lote_pesoinicial) || 0;
  const ra = parseFloat(d.lote_racao) || 0;
  
  // Fator K = 100 * peso / comprimento³
  const fatorK = tam > 0 ? (100 * peso / Math.pow(tam, 3)).toFixed(3) : "";
  
  // Biomassa estimada
  const biomassa = qtd > 0 ? ((peso * qtd) / 1000).toFixed(1) : "";
  
  // CA acumulado
  let ca = "";
  if(ra > 0 && peso > 0 && qtd > 0 && pi > 0){
    const bioAtual = (peso * qtd) / 1000;
    const bioInicial = (pi * qtd) / 1000;
    const ganho = bioAtual - bioInicial;
    if(ganho > 0) ca = (ra / ganho).toFixed(2);
  }
  
  // Ganho diário
  let gpd = "";
  if(d.lote_data && pi > 0 && peso > 0){
    const dataP = new Date(d.lote_data);
    const dataAtual = new Date(d.data);
    const dias = Math.floor((dataAtual - dataP) / 86400000);
    if(dias > 0) gpd = ((peso - pi) / dias).toFixed(2);
  }

  aba.appendRow([
    d.data||"", d.gaiola||"", d.hora||"", d.resp||"", d.fase||"", d.clima||"",
    d.amostras||0, d.peso||0, d.min||0, d.max||0,
    d.dp||0, d.tam||0,
    d.cond||"", d.coloracao||"", d.comportamento||"",
    d.lesoes||"", d.parasitas||"", d.manejo||"",
    fatorK, biomassa, ca, gpd,
    d.lote_data||"", d.lote_qtd||0, d.lote_pesoinicial||0, d.lote_racao||0,
    d.obs||""
  ]);
}

// ===================== LOTES =====================
function salvarLote(d) {
  const aba = getOuCriarAba("Lotes", [
    "Data Cadastro","Gaiola","Setor","Data Entrada",
    "Qtd Alevinos","Peso Inicial (g)","Fase","Fornecedor",
    "Origem","Responsável","Observações","Status"
  ]);
  aba.appendRow([
    d.data_cadastro||"", d.cage_id||"", d.setor||"",
    d.data_entrada||"", d.qtd||0, d.peso_inicial||0,
    d.fase||"", d.fornecedor||"", d.origem||"",
    d.resp||"", d.obs||"", d.status||"ativo"
  ]);
}

// ===================== FOTOS =====================
function salvarFoto(d) {
  try {
    const aba = getOuCriarAba("Fotos", [
      "Data","Gaiola","Nome Arquivo","Link Drive"
    ]);

    // Se já tem link direto (Drive API do app)
    if(d.link) {
      aba.appendRow([d.data||"", d.cage_id||"", d.nome||"", d.link]);
      return d.link;
    }

    // Se tem base64, faz upload via Apps Script
    if(d.base64) {
      const pasta = DriveApp.getFolderById(PASTA_FOTOS_ID);
      const nomeSubpasta = d.data || Utilities.formatDate(new Date(), "America/Sao_Paulo", "yyyy-MM-dd");
      let subpasta;
      const subs = pasta.getFoldersByName(nomeSubpasta);
      subpasta = subs.hasNext() ? subs.next() : pasta.createFolder(nomeSubpasta);
      const base64 = d.base64.includes(',') ? d.base64.split(',')[1] : d.base64;
      const blob = Utilities.newBlob(Utilities.base64Decode(base64), 'image/jpeg', d.nome || 'foto.jpg');
      const arquivo = subpasta.createFile(blob);
      arquivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      const linkFoto = arquivo.getUrl();
      aba.appendRow([d.data||"", d.cage_id||"", d.nome||"", linkFoto]);
      return linkFoto;
    }

    return null;
  } catch(err) {
    Logger.log("Erro foto: " + err.message);
    return null;
  }
}

// ===================== UTILITÁRIOS =====================
function getOuCriarAba(nome, cabecalho) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let aba = ss.getSheetByName(nome);
  if (!aba) {
    aba = ss.insertSheet(nome);
    aba.appendRow(cabecalho);
    const h = aba.getRange(1, 1, 1, cabecalho.length);
    h.setBackground("#1E3D2F");
    h.setFontColor("#E8DFC8");
    h.setFontWeight("bold");
    h.setFontSize(11);
    aba.setFrozenRows(1);
    aba.setColumnWidth(1, 100);
  }
  return aba;
}

function resposta(status, msg) {
  return ContentService
    .createTextOutput(JSON.stringify({ status, msg, ts: new Date().toISOString() }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ===================== CRIAR ABAS =====================
// Execute esta função UMA VEZ para criar todas as abas na planilha
function criarAbas() {
  getOuCriarAba("Alimentação Diária", [
    "Data","Gaiola","Setor","Arraç. Nº","Horário",
    "Clima","Temp Ar (°C)","Fase do Lote","Qualidade Visual Água",
    "Tratador 1","Tratador 2","Tipo Ração","Kg Fornecido","Consumo","Sobra (kg)",
    "Comportamento","Mortes","Causa Morte","GPS","Observações","Foto"
  ]);
  Logger.log("✅ Alimentação Diária OK");

  getOuCriarAba("Hídrico", [
    "Data","Horário","Responsável","Clima","Temp Ar (°C)",
    "Ponto","Local","Temp Água (°C)","OD (mg/L)","pH",
    "Secchi (cm)","Amônia (mg/L)","Nitrito (mg/L)","Nível (m)",
    "Cor da Água","Floração Algas","Peixes Superfície","GPS","Observações"
  ]);
  Logger.log("✅ Hídrico OK");

  getOuCriarAba("Biometria", [
    "Data","Gaiola","Horário","Responsável","Fase","Clima",
    "Nº Amostras","Peso Médio (g)","Menor (g)","Maior (g)",
    "Desvio Padrão (g)","Comprimento Médio (cm)",
    "Condição Corporal","Coloração","Comportamento",
    "Lesões (%)","Parasitas (%)","Manejo",
    "Fator K","Biomassa Est. (kg)","CA Acumulado","Ganho Diário (g/dia)",
    "Data Povoamento","Qtd Inicial","Peso Inicial (g)","Ração Acum. (kg)",
    "Observações"
  ]);
  Logger.log("✅ Biometria OK");

  getOuCriarAba("Lotes", [
    "Data Cadastro","Gaiola","Setor","Data Entrada",
    "Qtd Alevinos","Peso Inicial (g)","Fase","Fornecedor",
    "Origem","Responsável","Observações","Status"
  ]);
  Logger.log("✅ Lotes OK");

  getOuCriarAba("Despescas", [
    "Data","Gaiola","Peso Final (g)","Qtd Despescada",
    "Biomassa (kg)","Sobrevivência (%)","CA Final",
    "Destino","Responsável","Data Entrada","Dias Ciclo",
    "Qtd Inicial","Peso Inicial (g)","Ração Total (kg)","Observações"
  ]);
  Logger.log("✅ Despescas OK");

  getOuCriarAba("Diagnósticos", [
    "Data","Hora","Gaiola","Sintomas","Diagnóstico IA",
    "Temp Água (°C)","OD (mg/L)","Mortalidade"
  ]);
  Logger.log("✅ Diagnósticos OK");

  getOuCriarAba("Fotos", [
    "Data","Gaiola","Nome Arquivo","Link Drive"
  ]);
  Logger.log("✅ Fotos OK");

  Logger.log("✅ TODAS AS 7 ABAS CRIADAS COM SUCESSO!");
}

// ===================== DESPESCA =====================
function salvarDespesca(d) {
  const aba = getOuCriarAba("Despescas", [
    "Data","Gaiola","Peso Final (g)","Qtd Despescada",
    "Biomassa (kg)","Sobrevivência (%)","CA Final",
    "Destino","Responsável","Data Entrada","Dias Ciclo",
    "Qtd Inicial","Peso Inicial (g)","Ração Total (kg)","Observações"
  ]);
  const diasCiclo = d.lote_data_entrada ? 
    Math.floor((new Date(d.data) - new Date(d.lote_data_entrada)) / 86400000) : '';
  aba.appendRow([
    d.data||"", d.cage_id||"", d.peso_final||0, d.qtd_final||0,
    d.biomassa||0, d.sobrevivencia||0, "",
    d.destino||"", d.resp||"", d.lote_data_entrada||"", diasCiclo,
    d.lote_qtd_inicial||0, d.lote_peso_inicial||0, d.racao_total||0, d.obs||""
  ]);
}