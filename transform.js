// transform.js
// Módulo para transformar archivos Excel

const Excel = require('exceljs');
const path = require('path');
const fs = require('fs').promises;

// Encabezados exactos en el orden indicado
const HEADERS = [
  'SOCIEDAD','VISTA','SERIE','NUM_FRA','NUMERO_MANUAL','CIF','CLIENTE','D347','FEC_FRA',
  'ARREN_LOCALNEG','ARREN_REFCATASTRAL','ARREN_SITUACION','TIPO_FRA','NUM.FRA_ORIGEN','DELEGACION',
  'IMPORTE_BRUTO','TIPO_FRA_SII','CLAVE_RE','CLAVE_RE_AD1','CLAVE_RE_AD2','TIPO_OPERACION','FEC_EXPED',
  'FRA_SIMPLIFICADA','FRA_SIMPLI_IDEN','FRA_NOIDEN_DEST','IGIC_BIEN25','IGIC_DOCU25','IGIC_PROT25',
  'IGIC_NOTA25','DIARIO1','BASE1','IVA1','CUOTA1','TOTAL','CTA_DEUDORA','SCTA_DEUDORA','BASE_RETENCION',
  'PORCENTAJE_RETENCION','IMPORTE_RETENCION','CTA_RETENCION','SCTA_RETENCION','FRA_IVA_DIFERIDO','SERIE_DEVENGO',
  'CTA_DIFERIDO','SCTA_DIFERIDO','PORCENTAJE_IRCM','BASE_IRCM','IMPORTE_IRCM','CTA_IRCM','SCTA_IRCM',
  'CLAVE_340','NUMERO_FACTURAS','PRIMER_NUMERO','ULTIMO_NUMERO','EMIT_TERCEROS','CENTRO_COSTE','CON_CONCEPTO',
  'CON_CANTIDAD','CON_PRECIO','CON_IMPORTE','CON_DIARIO','CON_BASE','CON_PORCENTAJE','CON_IVA','ING_CTA',
  'ING_SCTA','ING_IMPORTE','ING_PROMOCION','ING_TIPO','ING_APARTADO','ING_CAPITULO','ING_PARTIDA','TIPO_INMUEBLE',
  'CLAVE1','CLAVE2','CLAVE3','CLAVE4','ING_UNIDADES'
];

/**
 * Transforma un archivo Excel
 * @param {string} inputFilePath - Ruta del archivo Excel de entrada
 * @param {string} outputFilePath - Ruta del archivo Excel de salida
 * @param {Function} progressCallback - Función callback para reportar progreso (current, total)
 * @param {number} serie - Valor de serie (4=CIELO, 5=NOVA, 6=GREEN)
 * @returns {Promise<void>}
 */
async function transformExcel(inputFilePath, outputFilePath, progressCallback = null, serie = 6) {
  const wbIn = new Excel.Workbook();
  await wbIn.xlsx.readFile(inputFilePath);
  
  // Calcular fórmulas antes de leer valores (importante para obtener resultados de fórmulas)
  wbIn.calcProperties.fullCalcOnLoad = true;

  const sheetIn = wbIn.worksheets[0];

  const wbOut = new Excel.Workbook();
  // Configurar cálculo automático para que las fórmulas se ejecuten
  wbOut.calcProperties.fullCalcOnLoad = true;
  const sheetOut = wbOut.addWorksheet('Output');

  // Escribir encabezados
  sheetOut.addRow(HEADERS);

  const promoNames = { '85': 'Grenn', '86': 'Cielo', '87': 'Nova' };

  // IMPORTANTE: ignorar encabezado → comenzar desde fila 2
  const firstDataRow = 2;
  const totalRows = sheetIn.rowCount - firstDataRow + 1;

  for (let r = firstDataRow; r <= sheetIn.rowCount; r++) {
    const inRow = sheetIn.getRow(r);

    const val = (cellRef) => {
      const c = inRow.getCell(cellRef);
      if (!c || c.value === null || typeof c.value === 'undefined') return '';
      if (typeof c.value === 'object') {
        if (c.value.richText) return c.value.richText.map(x => x.text).join('');
        if (c.value.text) return c.value.text;
        if (c.value.result) return c.value.result;
        return c.value;
      }
      return c.value;
    };

    // Valores importados por nombre de columna
    const A = val('A');
    const B = val('B');
    const C = val('C');
    const E = val('E');
    const I = val('I');
    const K = parseFloat(val('K') || 0);
    const O = parseFloat(val('O') || 0);
    
    // Función auxiliar para parsear números de manera más robusta
    const parseNum = (cellValue) => {
      // Si ya es un número, retornarlo directamente
      if (typeof cellValue === 'number') {
        return isNaN(cellValue) ? 0 : cellValue;
      }
      // Si es null, undefined o string vacío, retornar 0
      if (!cellValue && cellValue !== 0) return 0;
      // Convertir a string y limpiar
      const str = String(cellValue).trim();
      if (str === '') return 0;
      // Reemplazar coma por punto para decimales
      const cleaned = str.replace(',', '.');
      const num = parseFloat(cleaned);
      return isNaN(num) ? 0 : num;
    };
    
    // Leer AF y AG del archivo de entrada (AE se calculará con fórmula)
    const AF = parseNum(val('AF') || 0);
    const AG = parseNum(val('AG') || 0);

    // Constantes y reglas
    const SOCIEDAD = 3;
    const VISTA = 1;
    // SERIE viene del parámetro (4=CIELO, 5=NOVA, 6=GREEN)
    const SERIE = serie;

    // fila de salida (contando encabezado)
    const currentOutputRowIndex = sheetOut.rowCount + 1;

    // Función para convertir índice de columna a letra de Excel (0=A, 1=B, ..., 30=AE, 32=AG)
    const colToLetter = (colIndex) => {
      if (colIndex < 0) return '';
      let result = '';
      let num = colIndex;
      while (num >= 0) {
        result = String.fromCharCode(65 + (num % 26)) + result;
        num = Math.floor(num / 26) - 1;
      }
      return result;
    };

    // Encontrar índices de columnas que contienen los valores de AE y AG
    // BASE1 contiene el valor de AE y está en la columna AE (índice 30)
    // CUOTA1 contiene el valor calculado y está en la columna AG (índice 32)
    const colAE_index = HEADERS.indexOf('BASE1'); // BASE1 está en la columna AE (índice 30)
    const colAF_index = HEADERS.indexOf('IVA1'); // BASE1 está en la columna AE (índice 31)
    const colAG_index = HEADERS.indexOf('CUOTA1'); // CUOTA1 está en la columna AG (índice 32)
    
    // Convertir a letras de Excel
    const colAE = colToLetter(colAE_index);
    const colAF = colToLetter(colAF_index);
    const colAG = colToLetter(colAG_index);

    // Fórmula en sintaxis de ExcelJS (inglés, con coma como separador)
    const NUM_FRA_formula = `RIGHT(E${currentOutputRowIndex},3)`;
    
    // Fórmulas para referenciar AE y AG (formato Excel: =+AE2)
    // ExcelJS requiere el formato sin el signo = cuando usas { formula: '...' }
    const formulaAE = `${colAE}${currentOutputRowIndex}`;
    const formulaAG = `${colAG}${currentOutputRowIndex}`;
    const formulaAF = `${colAF}${currentOutputRowIndex}`;
    
    const D347 = 'S';
    const ARREN_LOCALNEG = 'N';

    const DIARIO1 = 1;
    // BASE1 debe contener el valor calculado de K + O del archivo de entrada (no una fórmula)
    const BASE1 = K + O;  // Suma directa de los valores leídos del archivo de entrada
    const IVA1 = 10;
    // CUOTA1 = AE * AF / 100 (fórmula que referencia a BASE1 y usa el valor de AF del archivo de entrada)
    // Si AF debe ser una referencia, necesitaríamos saber en qué columna está en el archivo de salida
    const CUOTA1_formula = `${colAE}${currentOutputRowIndex}*${colAF}${currentOutputRowIndex}/100`;
    // TOTAL = AE + AG (fórmula que referencia a BASE1 y CUOTA1)
    const TOTAL_formula = `${colAE}${currentOutputRowIndex}+${colAG}${currentOutputRowIndex}`;
    const CTA_DEUDORA = 4300;
    const SCTA_DEUDORA = 0;

    // BASE_RETENCION debe referenciar a AE (BASE1)
    const BASE_RETENCION_formula = formulaAE;

    const FRA_IVA_DIFERIDO = 'N';

    // CON_IMPORTE debe referenciar a AE (BASE1)
    const CON_IMPORTE_formula = formulaAE;
    const CON_DIARIO = 1;
    // CON_BASE debe referenciar a AE (BASE1)
    const CON_BASE_formula = formulaAE;
    const CON_PORCENTAJE = 10;
    // CON_IVA debe referenciar a AG (CUOTA1)
    const CON_IVA_formula = formulaAG;

    const ING_CTA = 7060;
    // ING_SCTA basado en SERIE: 4=2, 5=3, 6=4
    const ING_SCTA = serie === 4 ? 2 : serie === 5 ? 3 : 4;
    // ING_IMPORTE debe referenciar a AE (BASE1)
    const ING_IMPORTE_formula = formulaAE;
    // ING_PROMOCION basado en ING_SCTA: 2=86, 3=87, 4=85
    const ING_PROMOCION = ING_SCTA === 2 ? 86 : ING_SCTA === 3 ? 87 : 85;
    const ING_APARTADO = '02';
    const ING_CAPITULO = '01';
    const ING_PARTIDA = '001';

    const promoName = promoNames[String(ING_PROMOCION)] || String(ING_PROMOCION);
    const CON_CONCEPTO = `${promoName}House Reserva ${A}`;

    const outRowValues = [
      SOCIEDAD, VISTA, SERIE,
      { formula: NUM_FRA_formula },
      B,              // NUMERO_MANUAL
      I,              // CIF
      '', D347, C,
      ARREN_LOCALNEG, '', '', '', '', '',  // J, K, L, M, N, O (no se copian, se referencian en la fórmula)
      { formula: formulaAE },  // IMPORTE_BRUTO =+AE2 (referencia a BASE1 que contiene AE)
      '',              // TIPO_FRA_SII (debe estar vacío)
      '', '', '', '', '', '',
      '', '', '', '', '', '',
      // Nota: Las columnas K y O están en las posiciones 10 y 14 del array
      // K (índice 10) y O (índice 14) se copian del archivo de entrada para que BASE1 pueda referenciarlas
      DIARIO1, BASE1, IVA1, { formula: CUOTA1_formula }, { formula: TOTAL_formula }, CTA_DEUDORA, SCTA_DEUDORA, { formula: BASE_RETENCION_formula }, // BASE1 = K+O (valor calculado), CUOTA1, TOTAL, BASE_RETENCION =+AE2
      '', '', '', '', FRA_IVA_DIFERIDO, '',
      '', '', '', '', '', '', '',
      '', '', '', '', '', '', CON_CONCEPTO,
      '', '', { formula: CON_IMPORTE_formula }, CON_DIARIO, { formula: CON_BASE_formula }, CON_PORCENTAJE, { formula: CON_IVA_formula }, // CON_IMPORTE, CON_BASE =+AE2, CON_IVA =+AG2
      ING_CTA, ING_SCTA, { formula: ING_IMPORTE_formula }, ING_PROMOCION, '', ING_APARTADO, // ING_IMPORTE =+AE2
      ING_CAPITULO, ING_PARTIDA, '', '', '', '', '', ''
    ];

    sheetOut.addRow(outRowValues);

    // Reportar progreso
    if (progressCallback) {
      const current = r - firstDataRow + 1;
      progressCallback(current, totalRows);
    }
  }

  await wbOut.xlsx.writeFile(outputFilePath);
}

// Si se ejecuta directamente desde línea de comandos, mantener compatibilidad
if (require.main === module) {
  if (process.argv.length < 4) {
    console.error('Uso: node transform.js input.xlsx output.xlsx [serie]');
    process.exit(1);
  }

  const [,, INPUT_FILE, OUTPUT_FILE, SERIE_ARG] = process.argv;
  const serieValue = SERIE_ARG ? parseInt(SERIE_ARG) : 6;

  transformExcel(INPUT_FILE, OUTPUT_FILE, null, serieValue)
    .then(() => {
      console.log(`Archivo generado correctamente: ${OUTPUT_FILE}`);
    })
    .catch((error) => {
      console.error('Error al transformar el archivo:', error);
      process.exit(1);
    });
}

module.exports = { transformExcel };
