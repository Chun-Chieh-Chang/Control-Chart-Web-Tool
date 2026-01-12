import React from 'react';
import { Upload, ChevronLeft, ChevronDown } from 'lucide-react';

export const Sidebar: React.FC = () => {
  return (
    <div className="w-72 bg-white h-screen border-r border-gray-200 flex flex-col shrink-0">
      <div className="p-4 bg-indigo-900 text-white flex justify-between items-center">
        <h2 className="font-semibold text-lg">Configuration</h2>
        <ChevronLeft className="w-5 h-5 cursor-pointer hover:text-gray-300" />
      </div>

      <div className="p-4 flex-1 overflow-y-auto">
        <div className="mb-6">
          <label className="block text-sm font-semibold text-gray-700 mb-2">File Upload</label>
          <div className="border-2 border-dashed border-gray-300 rounded-lg p-6 flex flex-col items-center justify-center bg-gray-50 hover:bg-gray-100 transition-colors cursor-pointer relative group">
             <input type="file" className="absolute inset-0 opacity-0 cursor-pointer" />
            <Upload className="w-8 h-8 text-gray-400 mb-2 group-hover:text-indigo-600" />
            <p className="text-xs text-center text-gray-500">
              Drag files here or click to upload
              <br />
              (CSV, XLSX)
            </p>
          </div>
        </div>

        <div className="mb-4">
            <div className="flex justify-between items-center mb-2 cursor-pointer">
                <h3 className="font-bold text-sm text-gray-800">Parameters</h3>
                <ChevronDown className="w-4 h-4 text-gray-500" />
            </div>
            
            <div className="space-y-4">
                <div>
                    <label className="block text-xs font-medium text-gray-600 mb-1">Data Source</label>
                    <select className="w-full border border-gray-300 rounded px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500">
                        <option>Data Source</option>
                        <option>Machine 1</option>
                        <option>Machine 2</option>
                    </select>
                </div>

                <div>
                    <label className="block text-xs font-medium text-gray-600 mb-1">Time Range (Last 30 Days)</label>
                    <select className="w-full border border-gray-300 rounded px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500">
                        <option>Last 30 Days</option>
                        <option>Last 7 Days</option>
                        <option>Today</option>
                    </select>
                </div>

                <div>
                    <label className="block text-xs font-medium text-gray-600 mb-1">Control Limits (Default)</label>
                    <select className="w-full border border-gray-300 rounded px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500">
                        <option>Control Limits (Default)</option>
                        <option>Custom</option>
                    </select>
                </div>

                <div className="grid grid-cols-2 gap-2">
                    <div>
                        <label className="block text-xs font-medium text-gray-600 mb-1">Target Mean</label>
                        <input type="text" placeholder="Mean" className="w-full border border-gray-300 rounded px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500" />
                    </div>
                    <div>
                        <label className="block text-xs font-medium text-gray-600 mb-1">Sigma</label>
                        <input type="text" defaultValue="5" className="w-full border border-gray-300 rounded px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500" />
                    </div>
                </div>
            </div>
        </div>

        <button className="w-full bg-indigo-900 text-white py-2 rounded-md hover:bg-indigo-800 transition-colors mt-4 text-sm font-medium shadow-sm">
            Apply Settings
        </button>
      </div>

      <div className="p-4 border-t border-gray-200">
          <div className="flex justify-end">
             <ChevronLeft className="w-5 h-5 text-gray-400 rotate-180" />
          </div>
      </div>
    </div>
  );
};