/**
 * Componente para visualizar datos combinados de abonos - Taxi Monterrico
 */
import React, { useState } from 'react';
import { Download, Filter, Search, DollarSign, Users, Calendar, FileText } from 'lucide-react';
import { CombinedData } from '../types/excel';
import { exportCombinedToCSV } from '../utils/excelProcessor';

interface CombinedDataViewerProps {
  data: CombinedData;
}

export const CombinedDataViewer: React.FC<CombinedDataViewerProps> = ({ data }) => {
  const [searchTerm, setSearchTerm] = useState('');
  const [statusFilter, setStatusFilter] = useState<'all' | 'pendiente' | 'procesado' | 'cancelado'>('all');
  const [sourceFilter, setSourceFilter] = useState<string>('all');

  // Si no hay datos, mostrar tabla vacía
  if (!data) {
    return (
      <div className="bg-white rounded-lg border border-gray-200 shadow-sm">
        <div className="p-6 border-b border-gray-200">
          <div className="flex items-center justify-between mb-6">
            <div className="flex items-center">
              <img 
                src="https://taximonterrico.com/assets/logo_variante-CoJ5dU2i.png" 
                alt="Taxi Monterrico" 
                className="h-10 w-auto mr-4"
              />
              <div>
                <h3 className="text-xl font-bold text-gray-900">Carga de Abonos</h3>
                <p className="text-sm text-gray-600">Tabla de procesamiento de datos</p>
              </div>
            </div>
          </div>
        </div>

        {/* Tabla vacía */}
        <div className="overflow-x-auto w-full">
          <table className="w-[97%] mx-auto divide-y divide-gray-200" style={{ width: '97%' }}>
            <thead className="bg-gray-50">
              <tr>
                <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">
                  ITEM
                </th>
                <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">
                  BENEFICIARIO
                </th>
                <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">
                  DOCUMENTO
                </th>
                <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">
                  # DOCUMENTO
                </th>
                <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">
                  MONTO
                </th>
                <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">
                  CUENTA
                </th>
                <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">
                  ESTADO
                </th>
                <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">
                  OBSERVACIÓN
                </th>
                <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">
                  BANCO
                </th>
              </tr>
            </thead>
            <tbody className="bg-white divide-y divide-gray-200">
              <tr>
                <td colSpan={9} className="px-6 py-12 text-center">
                  <div className="text-center">
                    <FileText className="h-12 w-12 text-gray-400 mx-auto mb-4" />
                    <p className="text-gray-500 text-lg font-medium">Tabla de Abonos</p>
                    <p className="text-gray-400 text-sm mt-2">Sube archivos de BBVA o BCP para ver los datos aquí</p>
                  </div>
                </td>
              </tr>
            </tbody>
          </table>
        </div>
      </div>
    );
  }

  const filteredRecords = data.records.filter(record => {
    const matchesSearch = searchTerm === '' || 
      record.beneficiario?.toLowerCase().includes(searchTerm.toLowerCase()) ||
      record.documento?.toLowerCase().includes(searchTerm.toLowerCase()) ||
      record.cuenta_numero?.toLowerCase().includes(searchTerm.toLowerCase());
    
    const matchesStatus = statusFilter === 'all' || record.estado === statusFilter;
    const matchesSource = sourceFilter === 'all' || record.origen === sourceFilter;
    
    return matchesSearch && matchesStatus && matchesSource;
  });

  const totalMonto = filteredRecords.reduce((sum, record) => sum + (record.monto || record.monto_mn), 0);
  const uniqueClients = new Set(filteredRecords.map(r => r.beneficiario)).size;

  const getStatusColor = (status: string) => {
    switch (status) {
      case 'ABONO CORRECTO': 
      case 'TERMINADA OK': return 'bg-green-100 text-green-800';
      case 'ERROR': 
      case 'RECHAZADO': return 'bg-red-100 text-red-800';
      default: return 'bg-yellow-100 text-yellow-800';
    }
  };

  const handleExport = () => {
    const exportData = {
      ...data,
      records: filteredRecords
    };
    exportCombinedToCSV(exportData);
  };

  return (
    <div className="bg-white rounded-lg border border-gray-200 shadow-sm">
      <div className="p-6 border-b border-gray-200">
        <div className="flex items-center justify-between mb-6">
          <div className="flex items-center">
            <img 
              src="https://taximonterrico.com/assets/logo_variante-CoJ5dU2i.png" 
              alt="Taxi Monterrico" 
              className="h-10 w-auto mr-4"
            />
            <div>
              <h3 className="text-xl font-bold text-gray-900">Carga de Abonos</h3>
              <p className="text-sm text-gray-600">Datos combinados y procesados</p>
            </div>
          </div>
          <button
            onClick={handleExport}
            className="inline-flex items-center px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 focus:outline-none focus:ring-2 focus:ring-blue-500 transition-colors"
          >
            <Download className="h-4 w-4 mr-2" />
            Exportar CSV
          </button>
        </div>

        {/* Estadísticas */}
        <div className="grid grid-cols-1 md:grid-cols-4 gap-4 mb-6">
          <div className="bg-blue-50 p-4 rounded-lg">
            <div className="flex items-center">
              <FileText className="h-8 w-8 text-blue-600" />
              <div className="ml-3">
                <p className="text-sm font-medium text-blue-600">Total Registros</p>
                <p className="text-2xl font-bold text-blue-900">{filteredRecords.length}</p>
              </div>
            </div>
          </div>
          <div className="bg-green-50 p-4 rounded-lg">
            <div className="flex items-center">
              <DollarSign className="h-8 w-8 text-green-600" />
              <div className="ml-3">
                <p className="text-sm font-medium text-green-600">Monto Total</p>
                <p className="text-2xl font-bold text-green-900">S/ {totalMonto.toLocaleString('es-PE', { minimumFractionDigits: 2 })}</p>
              </div>
            </div>
          </div>
          <div className="bg-purple-50 p-4 rounded-lg">
            <div className="flex items-center">
              <Users className="h-8 w-8 text-purple-600" />
              <div className="ml-3">
                <p className="text-sm font-medium text-purple-600">Clientes</p>
                <p className="text-2xl font-bold text-purple-900">{uniqueClients}</p>
              </div>
            </div>
          </div>
          <div className="bg-orange-50 p-4 rounded-lg">
            <div className="flex items-center">
              <Calendar className="h-8 w-8 text-orange-600" />
              <div className="ml-3">
                <p className="text-sm font-medium text-orange-600">Procesado</p>
                <p className="text-sm font-bold text-orange-900">
                  {data.processedAt.toLocaleDateString('es-PE')}
                </p>
              </div>
            </div>
          </div>
        </div>

        {/* Filtros */}
        <div className="flex flex-col sm:flex-row gap-4 mb-6">
          <div className="flex-1">
            <div className="relative">
              <Search className="absolute left-3 top-1/2 transform -translate-y-1/2 text-gray-400 h-4 w-4" />
              <input
                type="text"
                placeholder="Buscar por beneficiario, documento o cuenta..."
                value={searchTerm}
                onChange={(e) => setSearchTerm(e.target.value)}
                className="w-full pl-10 pr-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent"
              />
            </div>
          </div>
          <select
            value={statusFilter}
            onChange={(e) => setStatusFilter(e.target.value as any)}
            className="px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent"
          >
            <option value="all">Todos los estados</option>
            <option value="ABONO CORRECTO">Abono Correcto</option>
            <option value="TERMINADA OK">Terminada OK</option>
            <option value="ERROR">Error</option>
          </select>
          <select
            value={sourceFilter}
            onChange={(e) => setSourceFilter(e.target.value)}
            className="px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent"
          >
            <option value="all">Todos los archivos</option>
            {data.sources.map((source, index) => (
              <option key={index} value={source}>{source}</option>
            ))}
          </select>
        </div>
      </div>

      {/* Tabla de datos - ACTUALIZADA 2024 */}
      <div className="overflow-x-auto w-full">
        <table className="w-[97%] mx-auto divide-y divide-gray-200" style={{ width: '97%' }}>
          <thead className="bg-gray-50">
            <tr>
              <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">
                ITEM
              </th>
              <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">
                BENEFICIARIO
              </th>
              <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">
                DOCUMENTO
              </th>
              <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">
                # DOCUMENTO
              </th>
              <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">
                MONTO
              </th>
              <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">
                CUENTA
              </th>
              <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">
                ESTADO
              </th>
              <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">
                OBSERVACIÓN
              </th>
              <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">
                BANCO
              </th>
            </tr>
          </thead>
          <tbody className="bg-white divide-y divide-gray-200">
            {filteredRecords.map((record, index) => (
              <tr key={record.id || index} className="hover:bg-gray-50">
                <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-900">
                  {index + 1}
                </td>
                <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-900">
                  {record.beneficiario}
                </td>
                <td className="px-6 py-4 whitespace-nowrap text-sm font-medium text-gray-900">
                  {record.documento_tipo || '-'}
                </td>
                <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-900">
                  {record.documento || '-'}
                </td>
                <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-900">
                  {record.monto_abonado ? `S/ ${record.monto_abonado.toLocaleString('es-PE', { minimumFractionDigits: 2 })}` : 
                   record.monto ? `S/ ${record.monto.toLocaleString('es-PE', { minimumFractionDigits: 2 })}` : 
                   record.monto_mn ? `S/ ${record.monto_mn.toLocaleString('es-PE', { minimumFractionDigits: 2 })}` : '-'}
                </td>
                <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-900">
                  {record.cuenta_numero || record.cuenta_tipo || '-'}
                </td>
                <td className="px-6 py-4 whitespace-nowrap">
                  <span className={`inline-flex px-2 py-1 text-xs font-semibold rounded-full ${getStatusColor(record.estado)}`}>
                    {record.estado}
                  </span>
                </td>
                <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-900">
                  {record.observaciones}
                </td>
                <td className="px-6 py-4 whitespace-nowrap text-sm font-medium text-blue-600">
                  {record.banco || 'BCP'}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
        
        {filteredRecords.length === 0 && (
          <div className="text-center py-12">
            <Filter className="h-12 w-12 text-gray-400 mx-auto mb-4" />
            <p className="text-gray-500">No se encontraron registros con los filtros aplicados</p>
          </div>
        )}
      </div>
    </div>
  );
};