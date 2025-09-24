/**
 * Tipos de datos para el sistema de carga de abonos - Taxi Monterrico
 */

export interface ExcelRow {
  [key: string]: any;
}

export interface AbonoRecord {
  id?: string;
  beneficiario: string;
  documento_tipo: string;
  documento: string;
  documento_2: string;
  documento_3: string;
  monto_mn: number;
  monto: number;
  tc: string;
  monto_abonado: number;
  monto_abonado_2: number;
  cuenta_tipo: string;
  cuenta_numero: string;
  cuenta_nombre: string;
  estado: string;
  observaciones: string;
  banco: string;
  origen: string; // Indica de qué archivo proviene (BCP o BBVA)
}

export interface ExcelData {
  fileName: string;
  sheets: ExcelSheet[];
  totalRows: number;
  uploadedAt: Date;
}

export interface ExcelSheet {
  name: string;
  data: ExcelRow[];
  headers: string[];
  rowCount: number;
}

export interface FileUploadState {
  file: File | null;
  data: ExcelData | null;
  isLoading: boolean;
  error: string | null;
}
export interface CombinedData {
  records: AbonoRecord[];
  totalRecords: number;
  sources: string[];
  processedAt: Date;
}