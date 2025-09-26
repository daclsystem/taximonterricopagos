import { AbonoRecord } from '../types/excel';

export interface ExcelData {
  fileName: string;
  sheets: Array<{
    name: string;
    data: any[];
  }>;
}

export function combineExcelData(data1: ExcelData, data2: ExcelData): { records: AbonoRecord[]; totalRecords: number; sources: string[]; processedAt: Date } {
  const records: AbonoRecord[] = [];
  
  console.log(`🔄 COMBINANDO DATOS: ${data1.fileName} + ${data2.fileName}`);
  
  // Procesar archivo 1 (BBVA) - mapeo simplificado: doc. identidad - titular - cuenta - importe - situacion
  data1.sheets.forEach((sheet, sheetIndex) => {
    console.log(`Procesando hoja BBVA ${sheetIndex}: ${sheet.name}`);
    
    sheet.data.forEach((row, index) => {
      // Manejar tanto objetos como arrays para BBVA
      let rowData: any;
      
      if (typeof row === 'object' && row !== null && !Array.isArray(row)) {
        // Si viene como objeto (con cabeceras como propiedades)
        rowData = row;
        // Función para obtener el valor de una columna manejando variaciones de acentos y caracteres extraños
        const getColumnValue = (rowObj: any, columnName: string, fallbackName?: string): string => {
          // Intentar con el nombre exacto primero
          if (rowObj[columnName] !== undefined) {
            return String(rowObj[columnName] || '');
          }
          
          // Si hay un nombre alternativo, intentarlo
          if (fallbackName && rowObj[fallbackName] !== undefined) {
            return String(rowObj[fallbackName] || '');
          }
          
          // Buscar en todas las claves que contengan la palabra (sin acentos y sin caracteres extraños)
          const normalizedColumnName = columnName.normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase();
          
          for (const key of Object.keys(rowObj)) {
            // Normalizar la clave y limpiar caracteres extraños
            const normalizedKey = key
              .normalize('NFD')
              .replace(/[\u0300-\u036f]/g, '') // Quitar acentos
              .replace(/[^\w\s]/g, '') // Quitar caracteres especiales excepto letras, números y espacios
              .toLowerCase();
            
            // Buscar coincidencias parciales
            if (normalizedKey.includes(normalizedColumnName) || normalizedColumnName.includes(normalizedKey)) {
              return String(rowObj[key] || '');
            }
          }
          
          // Buscar por patrones específicos para casos conocidos
          if (columnName.toLowerCase().includes('situacion')) {
            for (const key of Object.keys(rowObj)) {
              if (key.toLowerCase().includes('situacion') || key.toLowerCase().includes('situaci')) {
                return String(rowObj[key] || '');
              }
            }
          }
          
          return '';
        };

        const record: AbonoRecord = {
          id: `${data1.fileName}_${index}`,
          // Mapeo correcto para BBVA: Titular(Archivo) / Doc.Identidad / Cuenta / Importe / Situación
          beneficiario: getColumnValue(rowData, 'Titular(Archivo)'), 
          documento_tipo: '', 
          documento: getColumnValue(rowData, 'Doc.Identidad').replace(/^L\s*-\s*/i, ''), // Doc.Identidad para BBVA, quitando "L - "
          documento_2: '',
          documento_3: '',
          monto_mn: 0,
          monto: parseFloat(getColumnValue(rowData, 'Importe')) || 0, // Importe para BBVA
          tc: '',
          monto_abonado: 0,
          monto_abonado_2: 0,
          cuenta_tipo: '',
          cuenta_numero: getColumnValue(rowData, 'Cuenta').replace(/-/g, ''), // Cuenta para BBVA
          cuenta_nombre: '',
          estado: getColumnValue(rowData, 'Situación', 'Situacion'), // Situación con manejo de caracteres extraños
          observaciones: '',
          banco: 'BBVA',
          origen: data1.fileName
        };
        
        console.log(`🔍 DEBUGGING BBVA COMBINED:`);
        console.log(`  - 'Titular(Archivo)': "${rowData['Titular(Archivo)']}"`);
        console.log(`  - 'Doc.Identidad': "${rowData['Doc.Identidad']}"`);
        console.log(`  - 'Cuenta': "${rowData['Cuenta']}"`);
        console.log(`  - 'Importe': "${rowData['Importe']}"`);
        console.log(`  - 'Situación' (con tilde): "${rowData['Situación']}"`);
        console.log(`  - 'Situacion' (sin tilde): "${rowData['Situacion']}"`);
        
        // Buscar claves que contengan "situaci" para debugging
        const situacionKeys = Object.keys(rowData).filter(key => 
          key.toLowerCase().includes('situaci')
        );
        console.log(`  - Claves que contienen 'situaci':`, situacionKeys);
        situacionKeys.forEach(key => {
          console.log(`    - "${key}": "${rowData[key]}" → estado`);
        });
        
        console.log(`📊 RESULTADO MAPEADO:`);
        console.log(`  - beneficiario: "${record.beneficiario}"`);
        console.log(`  - documento: "${record.documento}"`);
        console.log(`  - cuenta_numero: "${record.cuenta_numero}"`);
        console.log(`  - monto: ${record.monto}`);
        console.log(`  - estado: "${record.estado}"`);
        console.log(`🔑 TODAS LAS CLAVES:`, Object.keys(rowData));
        console.log(`========================`);
        
        console.log(`BBVA COMBINED Record ${index}: beneficiario="${record.beneficiario}", documento="${record.documento}", monto=${record.monto}, estado="${record.estado}"`);
        
        if (record.beneficiario || record.monto > 0 || record.estado || record.cuenta_numero) {
          records.push(record);
        }
      } else if (Array.isArray(row)) {
        // Si viene como array (lógica anterior)
        const rowArray = row as any[];
        const record: AbonoRecord = {
          id: `${data1.fileName}_${index}`,
          // Mapeo correcto para BBVA: Titular(Banco) / Doc.Identidad / Cuenta / Importe / Situación
          beneficiario: String(rowArray[5] || ''), // Columna 6: Titular(Banco)
          documento_tipo: '', 
          documento: String(rowArray[6] || ''), // Columna 7: Doc.Identidad
          documento_2: '',
          documento_3: '',
          monto_mn: 0,
          monto: parseFloat(String(rowArray[7] || '0')) || 0, // Columna 8: Importe
          tc: '',
          monto_abonado: 0,
          monto_abonado_2: 0,
          cuenta_tipo: '',
          cuenta_numero: String(rowArray[8] || '').replace(/-/g, ''), // Columna 9: Cuenta
          cuenta_nombre: '',
          estado: String(rowArray[9] || ''), // Columna 10: Situación
          observaciones: '',
          banco: 'BBVA',
          origen: data1.fileName
        };
        
        if (record.beneficiario || record.monto > 0 || record.estado || record.cuenta_numero) {
          records.push(record);
        }
      }
    });
  });
  
  // Procesar archivo 2 (BCP) - mapeo simplificado: documento - beneficiario - cuenta - monto - estado
  data2.sheets.forEach((sheet, sheetIndex) => {
    console.log(`Procesando hoja BCP ${sheetIndex}: ${sheet.name}`);
    
    sheet.data.forEach((row, index) => {
      // Manejar tanto objetos como arrays para BCP
      let rowData: any;
      
      if (typeof row === 'object' && row !== null && !Array.isArray(row)) {
        // Si viene como objeto (con cabeceras como propiedades)
        rowData = row;
        const record: AbonoRecord = {
          id: `${data2.fileName}_${index}`,
          // Mapeo correcto para BCP: Beneficiario - Nombre / Documento / Cuenta - Número / Monto / Estado
          beneficiario: String(rowData['Beneficiario - Nombre'] || ''), 
          documento_tipo: '', 
          documento: String(rowData['Documento'] || ''), // Documento para BCP
          documento_2: '',
          documento_3: '',
          monto_mn: 0,
          monto: parseFloat(String(rowData['Monto'] || '0')) || 0, 
          tc: '',
          monto_abonado: 0,
          monto_abonado_2: 0,
          cuenta_tipo: '',
          cuenta_numero: String(rowData['Cuenta - Número'] || '').replace(/-/g, ''), 
          cuenta_nombre: '',
          estado: String(rowData['Estado'] || '') || '', // Estado para BCP
          observaciones: '',
          banco: 'BCP',
          origen: data2.fileName
        };
        
        console.log(`=== BCP RECORD ${index} ===`);
        console.log(`Datos originales del Excel:`, rowData);
        console.log(`🔍 DEBUGGING BCP:`);
        console.log(`  - 'Beneficiario - Nombre': "${rowData['Beneficiario - Nombre']}"`);
        console.log(`  - 'Documento': "${rowData['Documento']}"`);
        console.log(`  - 'Cuenta - Número': "${rowData['Cuenta - Número']}"`);
        console.log(`  - 'Monto': "${rowData['Monto']}"`);
        console.log(`  - 'Estado': "${rowData['Estado']}"`);
        console.log(`🔍 BUSCANDO VARIACIONES DE 'Documento':`);
        console.log(`  - 'Documento - Tip': "${rowData['Documento - Tip']}"`);
        console.log(`  - 'Documento - 1': "${rowData['Documento - 1']}"`);
        console.log(`  - 'Documen': "${rowData['Documen']}"`);
        console.log(`🔍 TODAS LAS CLAVES DISPONIBLES:`, Object.keys(rowData));
        console.log(`📊 RESULTADO MAPEADO:`);
        console.log(`  - beneficiario: "${record.beneficiario}"`);
        console.log(`  - documento: "${record.documento}"`);
        console.log(`  - cuenta_numero: "${record.cuenta_numero}"`);
        console.log(`  - monto: ${record.monto}`);
        console.log(`  - estado: "${record.estado}"`);
        console.log(`🔑 TODAS LAS CLAVES:`, Object.keys(rowData));
        console.log(`========================`);
        
        console.log(`BCP COMBINED Record ${index}: beneficiario="${record.beneficiario}", documento="${record.documento}", monto=${record.monto}, estado="${record.estado}"`);
        
        if (record.beneficiario || record.monto > 0 || record.estado || record.cuenta_numero) {
          records.push(record);
        }
      } else if (Array.isArray(row)) {
        // Si viene como array, usar la lógica anterior
        const rowArray = row as any[];
        
        const record: AbonoRecord = {
          id: `${data2.fileName}_${index}`,
          // Mapeo correcto para BCP: Beneficiario - Nombre / Documento / Cuenta - Número / Monto / Estado
          beneficiario: String(rowArray[0] || ''), 
          documento_tipo: '', 
          documento: String(rowArray[2] || '') || '-', // Columna C: Documento
          documento_2: '',
          documento_3: '',
          monto_mn: 0,
          monto: parseFloat(String(rowArray[6] || '0')) || 0, 
          tc: '',
          monto_abonado: 0,
          monto_abonado_2: 0,
          cuenta_tipo: '',
          cuenta_numero: String(rowArray[11] || '').replace(/-/g, ''), 
          cuenta_nombre: '',
          estado: String(rowArray[12] || '') || '', // Columna M: Estado
          observaciones: '',
          banco: 'BCP',
          origen: data2.fileName
        };
        
        if (record.beneficiario || record.monto > 0 || record.estado || record.cuenta_numero) {
          records.push(record);
        }
      }
    });
  });
  
  console.log(`✅ COMBINACIÓN COMPLETADA: ${records.length} registros totales`);
  return {
    records,
    totalRecords: records.length,
    sources: [data1.fileName, data2.fileName],
    processedAt: new Date()
  };
}

export async function exportCombinedToXLSX(data: { records: AbonoRecord[]; totalRecords: number; sources: string[]; processedAt: Date }, filename: string = 'Carga_de_Abonos.xlsx'): Promise<void> {
  if (!data || !data.records || data.records.length === 0) {
    console.log('No hay datos para exportar');
    return;
  }

  try {
    // Importar ExcelJS dinámicamente
    const ExcelJS = await import('exceljs');
    
    // Crear un nuevo workbook
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Carga de Abonos');

    // Definir las columnas
    worksheet.columns = [
      { header: 'ITEM', key: 'item', width: 8 },
      { header: 'BENEFICIARIO', key: 'beneficiario', width: 30 },
      { header: 'DOCUMENTO', key: 'documento', width: 15 },
      { header: 'CUENTA', key: 'cuenta', width: 20 },
      { header: 'MONTO', key: 'monto', width: 15 },
      { header: 'ESTADO', key: 'estado', width: 20 },
      { header: 'BANCO', key: 'banco', width: 10 }
    ];

    // Estilizar la fila de encabezados
    worksheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFF' } };
    worksheet.getRow(1).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: '366092' }
    };

    // Agregar los datos
    data.records.forEach((record, index) => {
      const row = worksheet.addRow({
        item: index + 1,
        beneficiario: record.beneficiario,
        documento: record.documento || record.documento_2 || record.documento_3 || record.documento_tipo || '-',
        cuenta: record.cuenta_numero,
        monto: record.monto,
        estado: record.estado,
        banco: record.banco
      });

      // Colorear la columna banco según el tipo
      const bancoCell = row.getCell('banco');
      if (record.banco === 'BBVA') {
        bancoCell.font = { color: { argb: '2563EB' } }; // Azul
      } else {
        bancoCell.font = { color: { argb: '7C3AED' } }; // Morado
      }

      // Formatear la columna de monto
      const montoCell = row.getCell('monto');
      montoCell.numFmt = 'S/ #,##0.00';
    });

    // Aplicar bordes a toda la tabla
    worksheet.eachRow((row) => {
      row.eachCell((cell) => {
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' }
        };
      });
    });

    // Generar el archivo
    const buffer = await workbook.xlsx.writeBuffer();
    
    // Crear y descargar el archivo
    const blob = new Blob([buffer], { 
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
    });
    
    const url = window.URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    window.URL.revokeObjectURL(url);

    console.log('✅ Archivo XLSX exportado exitosamente');
  } catch (error) {
    console.error('❌ Error exportando XLSX:', error);
    throw new Error('No se pudo exportar el archivo XLSX');
  }
}
