const HF_TOKEN = "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"; // Reemplazalo con tu token
let processedItems = [];

document.getElementById("fileInput").addEventListener("change", handleFile);
document.getElementById("exportBtn").addEventListener("click", exportToExcel);

async function handleFile(event) {
  const file = event.target.files[0];
  const data = await file.arrayBuffer();
  const workbook = XLSX.read(data);
  const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });

  const output = document.getElementById("output");
  output.innerHTML = "";
  processedItems = [];

  for (let row of rows) {
    for (let cell of row) {
      if (!cell || typeof cell !== "string") continue;

      const match = cell.match(/([a-zA-Z]{2,5}-?\d{3,6})/);
      const isLikelyCode = match ? match[1] : null;

      if (isLikelyCode) {
        try {
          const description = await classifyDescription(cell);
          const item = {
            code: isLikelyCode,
            description,
            status: "no-confirmado"
          };
          processedItems.push(item);
          renderItem(item);
          await delay(600); // <-- importante para evitar bloqueos
        } catch (error) {
          console.error("Error al procesar:", cell, error);
        }
      }
    }
  }

  document.getElementById("exportBtn").style.display = "block";
}
async function classifyDescription(text) {
  const response = await fetch("https://api-inference.huggingface.co/models/dslim/bert-base-NER", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${HF_TOKEN}`,
      "Content-Type": "application/json"
    },
    body: JSON.stringify({ inputs: text })
  });

  const result = await response.json();

  if (Array.isArray(result) && result[0]) {
    const tokens = result[0].filter(token =>
      ["ORG", "MISC", "PER", "PRODUCT"].includes(token.entity_group)
    );
    return tokens.map(t => t.word).join(" ") || text;
  }

  return text;
}

function renderItem(item) {
  const output = document.getElementById("output");

  const div = document.createElement("div");
  div.className = `item not-confirmed`;

  div.innerHTML = `
    <span class="code">${item.code}</span> - <span class="desc">${item.description}</span><br/>
    <label>Estado:
      <select>
        <option value="confirmado">✅ Confirmado</option>
        <option value="no-confirmado" selected>⚠️ No confirmado</option>
        <option value="no-encontrado">❌ No encontrado</option>
      </select>
    </label>
  `;

  const select = div.querySelector("select");
  select.addEventListener("change", () => {
    item.status = select.value;
    div.className = `item ${mapStatusToClass(item.status)}`;
  });

  output.appendChild(div);
}

function mapStatusToClass(status) {
  if (status === "confirmado") return "confirmed";
  if (status === "no-encontrado") return "not-found";
  return "not-confirmed";
}

function exportToExcel() {
  const wsData = [["Código", "Descripción", "Estado"]];

  processedItems.forEach(item => {
    wsData.push([item.code, item.description, item.status]);
  });

  const ws = XLSX.utils.aoa_to_sheet(wsData);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Ítems Verificados");

  XLSX.writeFile(wb, "verificados.xlsx");
}

function delay(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}
