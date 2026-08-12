// ═══════════════════════════════════════════════════════════
// SCRIPT DA API - SISTEMA DE LAUDOS (GOOGLE SHEETS)
// ═══════════════════════════════════════════════════════════
// Instruções:
// 1. Cole este código no Google Apps Script (Extensões > Apps Script)
// 2. Salve (Ctrl+S).
// 3. Clique em "Implantar" > "Nova implantação".
// 4. Selecione o tipo "App da Web".
// 5. Em "Quem pode acessar", coloque "Qualquer pessoa".
// 6. Autorize os acessos pedidos pelo Google.
// 7. Copie a "URL do app da Web" gerada.

const SPREADSHEET_NAME = "Laudos";
const SECRET_KEY = "perito:icrim123";

function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SPREADSHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SPREADSHEET_NAME);
    sheet.appendRow(["mobile_id", "data_sincronizacao", "dados_json", "baixado"]);
    sheet.getRange("A1:D1").setFontWeight("bold");
  }
  return sheet;
}

// Lida com as requisições GET (Do Streamlit puxando dados)
function doGet(e) {
  // Verifica senha simples enviada por parâmetro ?key=...
  if (!e.parameter.key || e.parameter.key !== SECRET_KEY) {
    return buildCorsResponse({error: "Não autorizado"});
  }

  const sheet = setupSheet();
  const data = sheet.getDataRange().getValues();
  
  if (data.length <= 1) {
    return buildCorsResponse([]);
  }

  const results = [];
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[3] !== "SIM") {
      results.push({
        row_index: i + 1,
        mobile_id: row[0],
        data_sincronizacao: row[1],
        dados_json: row[2]
      });
    }
  }

  return buildCorsResponse(results);
}

// Helper: retorna JSON com cabeçalhos CORS para o navegador aceitar
function buildCorsResponse(obj) {
  const output = ContentService.createTextOutput(JSON.stringify(obj));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}

// Lida com as requisições POST (Do Celular enviando ou Streamlit confirmando baixa)
function doPost(e) {
  try {
    const jsonString = e.postData.contents;
    const requestData = JSON.parse(jsonString);
    
    // Verifica a senha no POST
    if (requestData.key !== SECRET_KEY) {
      return buildCorsResponse({error: "Não autorizado"});
    }
    
    const sheet = setupSheet();
    
    // Ação: Confirmar que os laudos foram baixados (enviado pelo Streamlit)
        // Ação: Excluir ocorrência da planilha (enviado pelo Celular/HTML)
    if (requestData.action === "delete" && requestData.mobile_id) {
      const targetId = String(requestData.mobile_id).trim();
      const allData = sheet.getDataRange().getValues();
      let deletedCount = 0;
      
      // Percorre de baixo para cima para deletar sem desalinhar linhas
      for (let i = allData.length - 1; i >= 1; i--) {
        const rowMobileId = String(allData[i][0] || "").trim();
        let jsonOcorrencia = "";
        try {
          const parsed = JSON.parse(allData[i][2] || "{}");
          jsonOcorrencia = String(parsed.ocorrencia || "").trim();
        } catch(e) {}

        if (rowMobileId === targetId || (jsonOcorrencia && jsonOcorrencia === targetId)) {
          sheet.deleteRow(i + 1);
          deletedCount++;
        }
      }
      return buildCorsResponse({status: "ok", message: "Excluído " + deletedCount + " linha(s)"});
    }

    if (requestData.action === "mark_downloaded" && requestData.rows) {
      requestData.rows.forEach(rowIndex => {
        sheet.getRange(rowIndex, 4).setValue("SIM");
      });
      return buildCorsResponse({status: "ok", message: "Atualizado"});
    }
    
    // Ação Padrão: Salvar novo laudo (enviado pelo Celular)
    if (requestData.mobile_id && requestData.dados) {
      const now = new Date().toISOString();
      const stringDados = JSON.stringify(requestData.dados);
      
      const allData = sheet.getDataRange().getValues();
      let updated = false;
      for (let i = 1; i < allData.length; i++) {
        if (allData[i][0] === requestData.mobile_id) {
          sheet.getRange(i + 1, 2).setValue(now);
          sheet.getRange(i + 1, 3).setValue(stringDados);
          sheet.getRange(i + 1, 4).setValue("");
          updated = true;
          break;
        }
      }
      
      if (!updated) {
        sheet.appendRow([requestData.mobile_id, now, stringDados, ""]);
      }
      
      return buildCorsResponse({status: "ok", message: "Sincronizado"});
    }
    
    return buildCorsResponse({error: "Payload inválido"});

  } catch(error) {
    return buildCorsResponse({error: error.toString()});
  }
}

// Preflight CORS para navegadores
function doOptions(e) {
  return buildCorsResponse({status: "ok"});
}
