/**
 * Componente para visualizar datos de Excel procesados
 */
import React, { useState } from 'react';
import { Download, Eye, Table, BarChart3 } from 'lucide-react';
import { ExcelData } from '../types/excel';
import { exportToCSV } from '../utils/excelProcessorBCP';

interface DataViewerProps {
  data: ExcelData;
  title: string;
  bankName?: string;
}

export const DataViewer: React.FC<DataViewerProps> = ({ data, title, bankName }) => {
  const [selectedSheet, setSelectedSheet] = useState(0);
  const [viewMode, setViewMode] = useState<'preview' | 'full'>('preview');

  const currentSheet = data.sheets[selectedSheet];
  const previewData = viewMode === 'preview' 
    ? currentSheet.data.slice(0, 10) 
    : currentSheet.data;

  const handleExportCSV = () => {
    if (currentSheet && currentSheet.data.length > 0) {
      exportToCSV(currentSheet.data, `${data.fileName}_${currentSheet.name}`);
    }
  };

  if (!currentSheet) {
    return (
      <div className="bg-white rounded-lg border border-gray-200 p-6">
        <p className="text-gray-500 text-center">No hay datos para mostrar</p>
      </div>
    );
  }

  return (
    <div className="bg-white rounded-lg border border-gray-200 shadow-sm">
      <div className="p-4 border-b border-gray-200">
        <div className="flex items-center justify-between mb-4">
          <h3 className="text-lg font-semibold text-gray-800 flex items-center">
            <Table className="h-5 w-5 mr-2 text-blue-500" />
            {title}
          </h3>
          <div className="flex items-center space-x-2">
            <button
              onClick={() => setViewMode(viewMode === 'preview' ? 'full' : 'preview')}
              className="inline-flex items-center px-3 py-1.5 border border-gray-300 rounded-md text-sm font-medium text-gray-700 bg-white hover:bg-gray-50 focus:outline-none focus:ring-2 focus:ring-blue-500"
            >
              <Eye className="h-4 w-4 mr-1" />
              {viewMode === 'preview' ? 'Ver Todo' : 'Vista Previa'}
            </button>
            <button
              onClick={handleExportCSV}
              className="inline-flex items-center px-3 py-1.5 border border-gray-300 rounded-md text-sm font-medium text-gray-700 bg-white hover:bg-gray-50 focus:outline-none focus:ring-2 focus:ring-blue-500"
            >
              <Download className="h-4 w-4 mr-1" />
              Exportar CSV
            </button>
          </div>
        </div>

        {/* Información del archivo */}
        <div className="grid grid-cols-1 md:grid-cols-4 gap-4 mb-4">
          <div className="bg-blue-50 p-3 rounded-lg">
            <p className="text-xs text-blue-600 font-medium">Archivo</p>
            <p className="text-sm text-blue-900 truncate">{data.fileName}</p>
          </div>
          <div className="bg-green-50 p-3 rounded-lg">
            <p className="text-xs text-green-600 font-medium">Hojas</p>
            <p className="text-sm text-green-900">{data.sheets.length}</p>
          </div>
          <div className="bg-purple-50 p-3 rounded-lg">
            <p className="text-xs text-purple-600 font-medium">Total Filas</p>
            <p className="text-sm text-purple-900">{data.totalRows}</p>
          </div>
          <div className="bg-orange-50 p-3 rounded-lg">
            <p className="text-xs text-orange-600 font-medium">Subido</p>
            <p className="text-sm text-orange-900">
              {data.uploadedAt.toLocaleDateString('es-ES')}
            </p>
          </div>
        </div>

        {/* Selector de hojas */}
        {data.sheets.length > 1 && (
          <div className="flex flex-wrap gap-2">
            {data.sheets.map((sheet, index) => (
              <button
                key={index}
                onClick={() => setSelectedSheet(index)}
                className={`px-3 py-1.5 rounded-md text-sm font-medium transition-colors ${
                  selectedSheet === index
                    ? 'bg-blue-500 text-white'
                    : 'bg-gray-100 text-gray-700 hover:bg-gray-200'
                }`}
              >
                {sheet.name} ({sheet.rowCount})
              </button>
            ))}
          </div>
        )}
      </div>

      {/* Tabla de datos */}
      <div className="p-4">
        {previewData.length > 0 ? (
          <div className="overflow-x-auto">
            <table className="min-w-full table-auto">
              <thead>
                <tr className="bg-gray-50">
                  {bankName && (
                    <th className="px-4 py-2 text-left text-xs font-medium text-gray-500 uppercase tracking-wider border-b border-gray-200">
                      Banco
                    </th>
                  )}
                  {currentSheet.headers.map((header, index) => (
                    <th
                      key={index}
                      className="px-4 py-2 text-left text-xs font-medium text-gray-500 uppercase tracking-wider border-b border-gray-200"
                    >
                      {header}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody className="bg-white divide-y divide-gray-200">
                {previewData.map((row, rowIndex) => (
                  <tr key={rowIndex} className="hover:bg-gray-50">
                    {bankName && (
                      <td className="px-4 py-2 text-sm font-medium text-blue-600 border-b border-gray-100">
                        {bankName}
                      </td>
                    )}
                    {currentSheet.headers.map((header, colIndex) => (
                      <td
                        key={colIndex}
                        className="px-4 py-2 text-sm text-gray-900 border-b border-gray-100 max-w-xs truncate"
                        title={String(row[header] || '')}
                      >
                        {String(row[header] || '')}
                      </td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
            {viewMode === 'preview' && currentSheet.data.length > 10 && (
              <div className="mt-4 text-center">
                <p className="text-sm text-gray-500">
                  Mostrando 10 de {currentSheet.data.length} filas
                </p>
              </div>
            )}
          </div>
        ) : (
          <div className="text-center py-8">
            <BarChart3 className="h-12 w-12 text-gray-400 mx-auto mb-4" />
            <p className="text-gray-500">La hoja seleccionada no tiene datos</p>
          </div>
        )}
      </div>
    </div>
  );
};