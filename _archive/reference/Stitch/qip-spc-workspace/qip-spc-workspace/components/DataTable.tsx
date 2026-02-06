import React from 'react';
import { Search, Filter } from 'lucide-react';
import { Measurement, Status } from '../types';

interface DataTableProps {
  measurements: Measurement[];
}

export const DataTable: React.FC<DataTableProps> = ({ measurements }) => {
  return (
    <div className="bg-white rounded-lg shadow-sm border border-gray-200 flex flex-col h-full">
      <div className="p-4 border-b border-gray-200 flex justify-between items-center">
        <h3 className="font-bold text-gray-800">Data Table</h3>
        <div className="flex gap-2">
           <div className="relative">
             <Search className="w-4 h-4 absolute left-2 top-2.5 text-gray-400" />
             <input 
               type="text" 
               placeholder="Search data..." 
               className="pl-8 pr-3 py-2 border border-gray-300 rounded text-sm focus:outline-none focus:ring-1 focus:ring-indigo-500 w-48"
             />
           </div>
           <button className="flex items-center gap-1 px-3 py-2 border border-gray-300 rounded hover:bg-gray-50 text-sm font-medium text-gray-700">
             <Filter className="w-4 h-4" />
             Filter
           </button>
        </div>
      </div>
      <div className="overflow-auto custom-scrollbar flex-1">
        <table className="w-full text-left border-collapse">
          <thead className="bg-gray-50 sticky top-0 z-10">
            <tr>
              <th className="px-4 py-3 text-xs font-semibold text-gray-600 uppercase tracking-wider">Timestamp</th>
              <th className="px-4 py-3 text-xs font-semibold text-gray-600 uppercase tracking-wider">Sample ID</th>
              <th className="px-4 py-3 text-xs font-semibold text-gray-600 uppercase tracking-wider">Measurement Value</th>
              <th className="px-4 py-3 text-xs font-semibold text-gray-600 uppercase tracking-wider">Operator</th>
              <th className="px-4 py-3 text-xs font-semibold text-gray-600 uppercase tracking-wider">Machine</th>
              <th className="px-4 py-3 text-xs font-semibold text-gray-600 uppercase tracking-wider">Status</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-gray-200">
            {measurements.map((row) => (
              <tr key={row.id} className={`text-sm hover:bg-gray-50 ${row.status === Status.OutOfControl ? 'bg-red-50 hover:bg-red-100' : ''}`}>
                <td className="px-4 py-3 whitespace-nowrap text-gray-700">{row.timestamp}</td>
                <td className="px-4 py-3 whitespace-nowrap text-gray-700">{row.sampleId}</td>
                <td className={`px-4 py-3 whitespace-nowrap font-medium ${row.status === Status.OutOfControl ? 'text-red-700' : 'text-gray-900'}`}>{row.value}</td>
                <td className="px-4 py-3 whitespace-nowrap text-gray-600">{row.operator}</td>
                <td className="px-4 py-3 whitespace-nowrap text-gray-600">{row.machine}</td>
                <td className="px-4 py-3 whitespace-nowrap">
                  <span className={`px-2 py-1 rounded-full text-xs font-medium ${
                    row.status === Status.Normal 
                      ? 'bg-green-100 text-green-700' 
                      : row.status === Status.OutOfControl 
                        ? 'bg-red-600 text-white' 
                        : 'bg-yellow-100 text-yellow-800'
                  }`}>
                    {row.status}
                  </span>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
};