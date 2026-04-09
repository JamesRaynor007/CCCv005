// Constantes y variables
const DAYS = 365; // Número de días en un año
const OBJECTIVES = {
  DSO: 15,
  DIO: 30,
  DPO: 45,
  CCC: 60
};

// Función para crear la hoja de Excel (template)
function createTemplate() {
  const wb = XLSX.utils.book_new();

  const wsData = [
    ['Account', 'Amount'],
    ['Account Receivable (current)', 10000],
    ['Total Credit Sales (current)', 100000],
    ['Average Inventory (current)', 20000],
    ['Cost of Goods Sold (current)', 80000],
    ['Average Account Payable (current)', 15000],
    ['Account Receivable (last)', 12000],
    ['Total Credit Sales (last)', 95000],
    ['Average Inventory (last)', 22000],
    ['Cost of Goods Sold (last)', 78000],
    ['Average Account Payable (last)', 14000],
  ];

  const ws = XLSX.utils.aoa_to_sheet(wsData);
  XLSX.utils.book_append_sheet(wb, ws, 'Template');

  // Descargar el archivo
  XLSX.writeFile(wb, 'CashCycleTemplate.xlsx');
}

// Función para leer archivo XLSX cargado
function readWorkbook(file, callback) {
  const reader = new FileReader();
  reader.onload = function(e) {
    const data = e.target.result;
    const wb = XLSX.read(data, { type: 'binary' });
    callback(wb);
  };
  reader.readAsBinaryString(file);
}

// Función para extraer datos del template cargado
function extractDataFromWorkbook(wb) {
  const sheetName = wb.SheetNames[0];
  const sheet = wb.Sheets[sheetName];
  const json = XLSX.utils.sheet_to_json(sheet, { header: 1 });

  const data = {};
  json.forEach(row => {
    if (row.length >= 2) {
      const account = row[0].toString().toLowerCase();
      const amount = parseFloat(row[1]);
      if (account.includes('account receivable')) {
        if (!data.current) data.current = {};
        data.current.ar = amount;
        data.lastAr = amount; // Asumiendo iguales para simplificar, ajustar si se necesita distinguir
      } else if (account.includes('total credit sales')) {
        if (!data.current) data.current = {};
        data.current.sales = amount;
        data.lastSales = amount;
      } else if (account.includes('average inventory')) {
        if (!data.current) data.current = {};
        data.current.inventory = amount;
        data.lastInventory = amount;
      } else if (account.includes('cost of goods sold')) {
        if (!data.current) data.current = {};
        data.current.cogs = amount;
        data.lastCogs = amount;
      } else if (account.includes('average account payable')) {
        if (!data.current) data.current = {};
        data.current.payable = amount;
        data.lastPayable = amount;
      }
    }
  });
  return data;
}

// Función para calcular ratios
function computeRatios(data) {
  const current = data.current || {};
  const last = data.last || {};

  const safeCurrentSales = current.sales || 1;
  const safeCurrentCogs = current.cogs || 1;
  const safeLastSales = last.sales || 1;
  const safeLastCogs = last.cogs || 1;

  const arCurrent = current.ar || 0;
  const inventoryCurrent = current.inventory || 0;
  const payableCurrent = current.payable || 0;

  const arLast = last.ar || 0;
  const inventoryLast = last.inventory || 0;
  const payableLast = last.payable || 0;

  // Calculando DSO
  const currentDSO = (arCurrent / safeCurrentSales) * DAYS;
  const lastDSO = (arLast / safeLastSales) * DAYS;

  // Calculando DIO
  const currentDIO = (inventoryCurrent / safeCurrentCogs) * DAYS;
  const lastDIO = (inventoryLast / safeLastCogs) * DAYS;

  // Calculando DPO
  const currentDPO = (payableCurrent / safeCurrentCogs) * DAYS;
  const lastDPO = (payableLast / safeLastCogs) * DAYS;

  // Calculando CCC
  const currentCCC = currentDSO + currentDIO - currentDPO;
  const lastCCC = lastDSO + lastDIO - lastDPO;

  const results = [
    { name: 'DSO', current: currentDSO, last: lastDSO, obj: OBJECTIVES.DSO },
    { name: 'DIO', current: currentDIO, last: lastDIO, obj: OBJECTIVES.DIO },
    { name: 'DPO', current: currentDPO, last: lastDPO, obj: OBJECTIVES.DPO },
    { name: 'CCC', current: currentCCC, last: lastCCC, obj: OBJECTIVES.CCC }
  ];

  // Ajustar cálculo de % to OBJ
  results.forEach(r => {
    // Fórmula corregida: ((Valor Actual - Objetivo) / Objetivo) * 100
    r.percentObj = ((r.current - r.obj) / r.obj) * 100;

    // Evolución respecto al pasado
    r.percentEvo = r.last === 0 ? 0 : 100 * (r.current - r.last) / r.last;

    // Clases para estilos (puedes ajustar según estilos CSS)
    if (r.name !== 'DPO') {
      r.deltaClass = (r.current <= r.obj) ? 'good' : 'bad';
    } else {
      r.deltaClass = (r.current >= r.obj) ? 'good' : 'bad';
    }
    r.evoClass = (r.current >= r.last) ? 'evo-bad' : 'evo-good';
  });

  return results;
}

// Función para mostrar resultados en la tabla
function displayResults(results) {
  const tbody = document.querySelector('#results-table tbody');
  tbody.innerHTML = '';

  results.forEach(r => {
    const tr = document.createElement('tr');

    // Indicador
    const tdIndicator = document.createElement('td');
    tdIndicator.textContent = r.name;
    tr.appendChild(tdIndicator);

    // Valor Actual
    const tdCurrent = document.createElement('td');
    tdCurrent.textContent = r.current.toFixed(2);
    tr.appendChild(tdCurrent);

    // % to OBJ (columna intercambiada)
    const tdPercentObj = document.createElement('td');
    tdPercentObj.textContent = r.percentObj.toFixed(2) + '%';
    tr.appendChild(tdPercentObj);

    // Valor Objetivo
    const tdObj = document.createElement('td');
    tdObj.textContent = r.obj;
    tr.appendChild(tdObj);

    // Valor Pasado
    const tdLast = document.createElement('td');
    tdLast.textContent = r.last.toFixed(2);
    tr.appendChild(tdLast);

    // Estilos opcionales
    tr.children[0].style.fontWeight = 'bold';

    tbody.appendChild(tr);
  });
}

// Evento para calcular
document.getElementById('calcular').addEventListener('click', () => {
  const data = {
    current: {
      ar: parseFloat(document.getElementById('ar-current').value) || 0,
      sales: parseFloat(document.getElementById('sales-current').value) || 0,
      inventory: parseFloat(document.getElementById('inventory-current').value) || 0,
      cogs: parseFloat(document.getElementById('cogs-current').value) || 0,
      payable: parseFloat(document.getElementById('payable-current').value) || 0
    },
    last: {
      ar: parseFloat(document.getElementById('ar-last').value) || 0,
      sales: parseFloat(document.getElementById('sales-last').value) || 0,
      inventory: parseFloat(document.getElementById('inventory-last').value) || 0,
      cogs: parseFloat(document.getElementById('cogs-last').value) || 0,
      payable: parseFloat(document.getElementById('payable-last').value) || 0
    }
  };
  const results = computeRatios(data);
  displayResults(results);
});

// Evento para descargar plantilla XLSX
document.getElementById('download-template').addEventListener('click', () => {
  createTemplate();
});

// Evento para cargar archivo XLSX
document.getElementById('upload-file').addEventListener('change', (e) => {
  const file = e.target.files[0];
  if (file) {
    readWorkbook(file, (wb) => {
      const data = extractDataFromWorkbook(wb);
      // Asignar valores a los inputs
      document.getElementById('ar-current').value = data.current?.ar || '';
      document.getElementById('sales-current').value = data.current?.sales || '';
      document.getElementById('inventory-current').value = data.current?.inventory || '';
      document.getElementById('cogs-last').value = data.lastCogs || '';
      document.getElementById('payable-last').value = data.lastPayable || '';
    });
  }
});
