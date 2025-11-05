(function(){
const fileInput = document.getElementById('fileInput');
const processBtn = document.getElementById('processBtn');
const downloadBtn = document.getElementById('downloadBtn');
const imprimirBtn = document.getElementById('imprimirBtn');

const outputLiquidacion = document.getElementById('outputLiquidacion');
const outputEtiquetas = document.getElementById('outputEtiquetas');

const btnLiquidacion = document.getElementById('btnLiquidacion');
const btnEtiquetas = document.getElementById('btnEtiquetas');
const encabezadoLiquidacion = document.getElementById('encabezadoLiquidacion');

let lastPivotWB = null;
let pivotArrayLiquidacion = [];
let pivotArrayEtiquetas = [];
let columnasLiquidacion = [];
let columnasEtiquetas = [];

// CONFIGURACIÓN LIQUIDACIÓN DE TÚNEL
const colMapLiquidacion = {"Item":"Item","Proforma":"# Prof","Importador":"Importador","Unid x Master":"Empaque","Talla":"Talla","Lote Cliente":"Lote cliente","Lote M":"Lote mascara","Coche/Pallets":"# Coche","No. Cajas":"Cajas Liquidadas","Master":"Masters","Cajas Sobrantes":"# Cajas Sobrante","Libras":"Total de Libras"};
const groupFieldsLiquidacion = ["Item","# Prof","Importador","Empaque","Talla","Lote cliente","Lote mascara","# Coche"];
const sumFieldsLiquidacion = ["Cajas Liquidadas","Masters","# Cajas Sobrante","Total de Libras"];

// CONFIGURACIÓN REQUERIMIENTO DE ETIQUETAS
const colMapEtiquetas = {
  "Item":"Item",
  "Importador":"Importador",
  "Talla":"Talla",
  "Lote M":"Lote M",
  "Lote Cliente":"Lote Cliente",
  "Proforma":"Proforma",
  "Unid x Master":"Unid x Master",
  "No. Cajas":"No. Cajas",
  "conversion":"conversion",
  "SUBTOTAL MASTER":"SUBTOTAL MASTER",
  "CAJA SOBRANTE":"CAJA SOBRANTE",
  "master sobrante":"master sobrante",
  "TOTAL A IMPRIMIR":"TOTAL A IMPRIMIR"
};
const groupFieldsEtiquetas = ["Item","Importador","Talla","Lote M","Lote Cliente","Proforma","Unid x Master"];
const sumFieldsEtiquetas = ["No. Cajas","SUBTOTAL MASTER","CAJA SOBRANTE","master sobrante","TOTAL A IMPRIMIR"];

fileInput.addEventListener('change', ()=>{ processBtn.disabled = !fileInput.files || fileInput.files.length === 0; });

processBtn.addEventListener('click', async () => {
  const f = fileInput.files[0];
  if(!f) return;
  processBtn.disabled = true;

  try {
    const data = await f.arrayBuffer();
    const wb = XLSX.read(data, {type:'array'});
    const ws = wb.Sheets[wb.SheetNames[0]];
    const range = XLSX.utils.decode_range(ws['!ref']);
    const lastRow = range.e.r + 1;

    // --- LIQUIDACIÓN DE TÚNEL ---
    const headerArrLiq = XLSX.utils.sheet_to_json(ws, {range:"A2:Q2", header:1})[0] || [];
    const rawDataLiq = XLSX.utils.sheet_to_json(ws, {range:`A3:Q${lastRow}`, header: headerArrLiq, defval:""});
    pivotArrayLiquidacion = rawDataLiq.map(r=>{
      const obj = {};
      for(const key in colMapLiquidacion){ obj[colMapLiquidacion[key]] = r[key] ?? ""; }
      return obj;
    }).filter(r => { if(!r["# Coche"]) return false; return groupFieldsLiquidacion.some(f => f !== "# Coche" && r[f] !== "" && r[f] != null); });
    columnasLiquidacion = [...groupFieldsLiquidacion,...sumFieldsLiquidacion];
    renderizarTabla(pivotArrayLiquidacion, columnasLiquidacion, outputLiquidacion, sumFieldsLiquidacion);

    // --- REQUERIMIENTO DE ETIQUETAS ---
    const headerArrEtiq = XLSX.utils.sheet_to_json(ws, {range:"A2:Z2", header:1})[0] || [];
    const rawDataEtiq = XLSX.utils.sheet_to_json(ws, {range:`A3:Z${lastRow}`, header: headerArrEtiq, defval:""});

    pivotArrayEtiquetas = rawDataEtiq.map(r=>{
      const obj = {};
      for(const key in colMapEtiquetas){ obj[colMapEtiquetas[key]] = r[key] ?? ""; }

      const unidadesMaster = parseFloat(obj["Unid x Master"]) || 0;
      const noCajas = parseFloat(obj["No. Cajas"]) || 0;

      obj["conversion"] = unidadesMaster > 0 ? (noCajas / unidadesMaster).toFixed(2) : 0;
      obj["SUBTOTAL MASTER"] = unidadesMaster > 0 ? Math.floor(parseFloat(obj["conversion"])) : 0;
      obj["CAJA SOBRANTE"] = unidadesMaster > 0 ? unidadesMaster * (parseFloat(obj["conversion"]) - Math.floor(parseFloat(obj["conversion"]))) : 0;
      obj["master sobrante"] = 0;
      obj["TOTAL A IMPRIMIR"] = parseFloat(obj["SUBTOTAL MASTER"]) + parseFloat(obj["master sobrante"]);

      return obj;
    }).filter(r => groupFieldsEtiquetas.some(f => r[f] !== "" && r[f] != null));

    // 🔹 AGRUPAR SOLO FILAS QUE NO TIENEN PROFORMA CON "VA"
    const groupedEtiquetas = {};
    const filasSinAgruparVA = [];

    pivotArrayEtiquetas.forEach(row => {
      const proforma = row["Proforma"]?.toString().trim() || "";

      if(proforma.startsWith("VA")) {
        filasSinAgruparVA.push(row);
        return;
      }

      const key = `${row["Lote M"]}||${row["Lote Cliente"]}`;
      if(!groupedEtiquetas[key]){
        groupedEtiquetas[key] = {...row};
      } else {
        sumFieldsEtiquetas.forEach(f=>{
          const val = parseFloat(row[f]) || 0;
          groupedEtiquetas[key][f] = (parseFloat(groupedEtiquetas[key][f]) || 0) + val;
        });
        const unidadesMaster = parseFloat(groupedEtiquetas[key]["Unid x Master"]) || 0;
        const noCajas = parseFloat(groupedEtiquetas[key]["No. Cajas"]) || 0;
        const conversion = unidadesMaster > 0 ? (noCajas / unidadesMaster).toFixed(2) : 0;
        groupedEtiquetas[key]["conversion"] = conversion;
        groupedEtiquetas[key]["SUBTOTAL MASTER"] = unidadesMaster > 0 ? Math.floor(conversion) : 0;
        groupedEtiquetas[key]["CAJA SOBRANTE"] = unidadesMaster > 0 ? unidadesMaster * (conversion - Math.floor(conversion)) : 0;
        groupedEtiquetas[key]["TOTAL A IMPRIMIR"] = groupedEtiquetas[key]["SUBTOTAL MASTER"] + (groupedEtiquetas[key]["master sobrante"]||0);
      }
    });

    pivotArrayEtiquetas = [...filasSinAgruparVA, ...Object.values(groupedEtiquetas)];

    // 🔹 CALCULAR MASTER SOBRANTE (solo filas sin VA)
    const lotesMap = {};
    pivotArrayEtiquetas.forEach(row => {
      const lote = row["Lote M"];
      const proforma = row["Proforma"]?.toString().trim() || "";
      if(!lote) return;

      const vaKey = proforma.startsWith("VA") ? "VA" : "";
      const key = `${lote}||${vaKey}`;

      if(!lotesMap[key]) lotesMap[key] = [];
      if(!proforma.startsWith("VA")) lotesMap[key].push(row);
    });

    Object.values(lotesMap).forEach(filas => {
      if(filas.length === 0) return;
      const sumaSobrante = filas.reduce((acc, r) => acc + (parseFloat(r["CAJA SOBRANTE"]) || 0), 0);
      const filaMayorCaja = filas.reduce((maxR, r) => (parseFloat(r["CAJA SOBRANTE"]) || 0) > (parseFloat(maxR["CAJA SOBRANTE"])||0) ? r : maxR, filas[0]);
      const unidadesMaster = parseFloat(filaMayorCaja["Unid x Master"]) || 0;

      if(unidadesMaster > 0 && Math.floor(sumaSobrante / unidadesMaster) >= 1){
        filaMayorCaja["master sobrante"] = Math.floor(sumaSobrante / unidadesMaster);
      } else {
        filaMayorCaja["master sobrante"] = 0;
      }
      filaMayorCaja["TOTAL A IMPRIMIR"] = parseFloat(filaMayorCaja["SUBTOTAL MASTER"]) + parseFloat(filaMayorCaja["master sobrante"]);
    });

    // 🔹 Ordenar tabla: filas VA al inicio, resto por Lote M
    pivotArrayEtiquetas.sort((a, b) => {
      const proformaA = a["Proforma"]?.toString().trim() || "";
      const proformaB = b["Proforma"]?.toString().trim() || "";

      const isVA_A = proformaA.startsWith("VA");
      const isVA_B = proformaB.startsWith("VA");
      if (isVA_A && !isVA_B) return -1;
      if (!isVA_A && isVA_B) return 1;

      const valA = a["Lote M"]?.toString().trim() || "";
      const valB = b["Lote M"]?.toString().trim() || "";
      const numA = parseFloat(valA), numB = parseFloat(valB);
      if (!isNaN(numA) && !isNaN(numB)) return numA - numB;
      return valA.localeCompare(valB, 'es', { numeric: true });
    });

    columnasEtiquetas = [
      ...groupFieldsEtiquetas,
      "No. Cajas",
      "conversion",
      "SUBTOTAL MASTER",
      "CAJA SOBRANTE",
      "master sobrante",
      "TOTAL A IMPRIMIR"
    ];

    renderizarTabla(pivotArrayEtiquetas, columnasEtiquetas, outputEtiquetas, sumFieldsEtiquetas);

    lastPivotWB = wb;
    downloadBtn.disabled = false;
    imprimirBtn.disabled = false;
    processBtn.disabled = false;

    outputLiquidacion.style.display='block';
    outputEtiquetas.style.display='none';
    encabezadoLiquidacion.style.display='flex';

  } catch(err){ console.error(err); alert(err.message||err); processBtn.disabled=false; }
});

// ---------------- FUNCIONES DE TABLA ----------------
function renderizarTabla(datos, columnas, contenedor, sumFields){
  contenedor.innerHTML='';
  const table = document.createElement('table');
  const thead = document.createElement('thead');
  const trh = document.createElement('tr');
  columnas.forEach(c=>{
    const th = document.createElement('th'); th.textContent=c;
    const btn = document.createElement('button'); btn.textContent='🔽'; btn.className='filter-btn'; btn.dataset.col=c;
    btn.addEventListener('click', e=>{
      const drop = th.querySelector('.filter-dropdown');
      drop.style.display = drop.style.display==='block'?'none':'block';
    });
    th.appendChild(btn);
    const dropdown = document.createElement('div'); dropdown.className='filter-dropdown'; dropdown.id='filter-'+c;
    th.appendChild(dropdown);
    trh.appendChild(th);
  });
  thead.appendChild(trh);
  table.appendChild(thead);

  const tbody = document.createElement('tbody');
  table.appendChild(tbody);

  const tfoot = document.createElement('tfoot');
  const trTotal = document.createElement('tr');
  columnas.forEach(c=>{
    const td=document.createElement('td');
    if(sumFields.includes(c)){
      const total = datos.reduce((acc,r)=>acc+(parseFloat(r[c])||0),0);
      td.textContent = total.toFixed(2);
    } else if(c==="Item"){ td.textContent="TOTAL GENERAL"; }
    trTotal.appendChild(td);
  });
  tfoot.appendChild(trTotal);
  table.appendChild(tfoot);

  contenedor.appendChild(table);

  crearFiltros(datos, columnas, sumFields);
  aplicarFiltros(datos, columnas, sumFields);
}

function crearFiltros(datos, columnas, sumFields){
  columnas.forEach(c=>{
    const container = document.getElementById('filter-'+c);
    if(!container) return;
    container.innerHTML='';

    const valoresUnicos = [...new Set(datos.map(r => String(r[c] ?? "").trim()))].sort((a,b)=>{
      const numA = parseFloat(a), numB = parseFloat(b);
      if(!isNaN(numA) && !isNaN(numB)) return numA - numB;
      return a.localeCompare(b, 'es', {numeric:true});
    });

    const labelTodos = document.createElement('label');
    const cbTodos = document.createElement('input');
    cbTodos.type='checkbox';
    cbTodos.checked=true;
    cbTodos.dataset.todos='true';
    cbTodos.addEventListener('change',()=>{
      const checkboxes = Array.from(container.querySelectorAll('input[type=checkbox]')).filter(i=>!i.dataset.todos);
      checkboxes.forEach(i=>i.checked=cbTodos.checked);
      aplicarFiltros(datos, columnas, sumFields);
    });
    labelTodos.appendChild(cbTodos);
    labelTodos.appendChild(document.createTextNode('Todos'));
    container.appendChild(labelTodos);

    valoresUnicos.forEach(v=>{
      const label = document.createElement('label');
      const cb = document.createElement('input');
      cb.type='checkbox';
      cb.value=v;
      cb.checked=true;
      cb.addEventListener('change',()=>{
        const all = container.querySelector('input[data-todos]');
        const totalChecks = container.querySelectorAll('input[type=checkbox]:not([data-todos])').length;
        const checksMarcados = container.querySelectorAll('input[type=checkbox]:not([data-todos]):checked').length;
        all.checked = (checksMarcados === totalChecks);
        aplicarFiltros(datos, columnas, sumFields);
      });
      label.appendChild(cb);
      label.appendChild(document.createTextNode(v === "" ? "(vacío)" : v));
      container.appendChild(label);
    });
  });
}

function aplicarFiltros(datos, columnas, sumFields){
  const contenedor = (columnas === columnasEtiquetas) ? outputEtiquetas : outputLiquidacion;
  const table = contenedor.querySelector('table');
  const tbody = table.querySelector('tbody');
  tbody.innerHTML='';

  const filtrado = datos.filter(row => columnas.every(c => {
    const container = document.getElementById('filter-'+c);
    if(!container) return true;
    const checkboxes = Array.from(container.querySelectorAll('input[type=checkbox]')).filter(i => !i.dataset.todos);
    const seleccionados = checkboxes.filter(cb => cb.checked).map(cb => cb.value);
    const valorFila = String(row[c] ?? "").trim();
    return seleccionados.includes(valorFila);
  }));

  filtrado.forEach((row)=>{
    const tr=document.createElement('tr');
    columnas.forEach(c=>{
      const td=document.createElement('td');

      if(c === "master sobrante"){
        td.textContent = row[c] || 0;
      } else if(c === "TOTAL A IMPRIMIR"){
        td.textContent = (parseFloat(row["SUBTOTAL MASTER"]) + (parseFloat(row["master sobrante"])||0)).toFixed(2);
      } else if(sumFields.includes(c)){
        td.textContent = (parseFloat(row[c]||0)).toFixed(2);
      } else {
        td.textContent = row[c] || "";
      }

      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });

  actualizarFooter(filtrado, columnas, sumFields, table);
}

function actualizarFooter(filtrado, columnas, sumFields, table){
  const tfoot = table.querySelector('tfoot');
  const trTotal = tfoot.querySelector('tr');
  columnas.forEach((c,i)=>{
    if(sumFields.includes(c)){
      const total = filtrado.reduce((acc,r)=>acc + (parseFloat(r[c])||0), 0);
      trTotal.children[i].textContent = total.toFixed(2);
    }
  });
}

// Alternar vistas
btnLiquidacion.addEventListener('click', ()=>{
  outputLiquidacion.style.display='block';
  outputEtiquetas.style.display='none';
  encabezadoLiquidacion.style.display='flex';
  document.getElementById('encabezadoEtiquetas').style.display='none';
});

btnEtiquetas.addEventListener('click', ()=>{
  outputLiquidacion.style.display='none';
  outputEtiquetas.style.display='block';
  encabezadoLiquidacion.style.display='none';
  document.getElementById('encabezadoEtiquetas').style.display='flex';
});

// Descargar e imprimir
downloadBtn.addEventListener('click', ()=>{
  if(!lastPivotWB) return;
  const wbout = XLSX.write(lastPivotWB,{bookType:'xlsx',type:'array'});
  const blob = new Blob([wbout],{type:"application/octet-stream"});
  const a = document.createElement('a'); 
  a.href = URL.createObjectURL(blob); 
  a.download="tabla_dinamica_resultado.xlsx"; 
  document.body.appendChild(a); 
  a.click(); 
  a.remove();
});

imprimirBtn.addEventListener('click', ()=>{ 
  encabezadoLiquidacion.style.display = (outputLiquidacion.style.display === 'block') ? 'flex' : 'none';
  window.print(); 
});

document.addEventListener('click', e=>{
  document.querySelectorAll('.filter-dropdown').forEach(dd=>{
    if(!dd.contains(e.target) && !dd.previousSibling.contains(e.target)){ dd.style.display='none'; }
  });
});
})();
