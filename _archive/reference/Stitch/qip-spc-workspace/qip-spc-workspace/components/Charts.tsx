import React from 'react';
import { LineChart, Line, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, ReferenceLine } from 'recharts';
import { ChartDataPoint } from '../types';

interface ChartsProps {
  data: ChartDataPoint[];
}

export const Charts: React.FC<ChartsProps> = ({ data }) => {
  return (
    <div className="grid grid-cols-1 lg:grid-cols-2 gap-6 mb-6">
      {/* X-bar Chart */}
      <div className="bg-white p-4 rounded-lg shadow-sm border border-gray-200">
        <div className="flex justify-between items-center mb-4">
          <h3 className="font-bold text-gray-800">X-bar Chart</h3>
          <div className="flex gap-4 text-xs">
             <div className="flex items-center gap-1"><span className="w-2 h-2 bg-indigo-900 rounded-sm"></span> X bar</div>
             <div className="flex items-center gap-1"><span className="w-2 h-2 bg-teal-500 rounded-sm"></span> Measurement</div>
             <div className="flex items-center gap-1"><span className="w-2 h-2 bg-gray-800 rounded-sm"></span> Control Limit</div>
          </div>
        </div>
        <div className="h-64 w-full">
          <ResponsiveContainer width="100%" height="100%">
            <LineChart data={data} margin={{ top: 5, right: 30, left: -20, bottom: 5 }}>
              <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e5e7eb" />
              <XAxis dataKey="name" tick={{fontSize: 10}} interval={2} axisLine={false} tickLine={false} />
              <YAxis domain={['dataMin - 2', 'dataMax + 2']} tick={{fontSize: 10}} axisLine={false} tickLine={false} />
              <Tooltip 
                contentStyle={{backgroundColor: '#1f2937', color: 'white', border: 'none', borderRadius: '4px', fontSize: '12px'}}
                itemStyle={{color: '#9ca3af'}}
                labelStyle={{color: '#fff', fontWeight: 'bold'}}
              />
              <ReferenceLine y={18.5} label={{ value: 'Up. Lmt', position: 'right', fontSize: 10, fill: '#374151' }} stroke="#374151" strokeDasharray="3 3" />
              <ReferenceLine y={12.5} label={{ value: 'Low Lmt', position: 'right', fontSize: 10, fill: '#374151' }} stroke="#374151" strokeDasharray="3 3" />
              <ReferenceLine y={15.2} stroke="#3b82f6" strokeWidth={2} />
              <Line type="monotone" dataKey="value" stroke="#0d9488" strokeWidth={2} dot={{r: 3, fill: '#0d9488', strokeWidth: 1, stroke: '#fff'}} activeDot={{ r: 6 }} />
            </LineChart>
          </ResponsiveContainer>
        </div>
      </div>

      {/* R Chart */}
      <div className="bg-white p-4 rounded-lg shadow-sm border border-gray-200">
        <div className="flex justify-between items-center mb-4">
          <h3 className="font-bold text-gray-800">R Chart</h3>
          <div className="flex gap-4 text-xs">
             <div className="flex items-center gap-1"><span className="w-2 h-2 bg-indigo-900 rounded-sm"></span> R-bar</div>
             <div className="flex items-center gap-1"><span className="w-2 h-2 bg-teal-500 rounded-sm"></span> Range</div>
             <div className="flex items-center gap-1"><span className="w-2 h-2 bg-gray-800 rounded-sm"></span> Control Limit</div>
          </div>
        </div>
        <div className="h-64 w-full">
          <ResponsiveContainer width="100%" height="100%">
            <LineChart data={data} margin={{ top: 5, right: 30, left: -20, bottom: 5 }}>
              <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e5e7eb" />
              <XAxis dataKey="name" tick={{fontSize: 10}} interval={2} axisLine={false} tickLine={false} />
              <YAxis domain={[0, 15]} tick={{fontSize: 10}} axisLine={false} tickLine={false} />
              <Tooltip 
                 contentStyle={{backgroundColor: '#1f2937', color: 'white', border: 'none', borderRadius: '4px', fontSize: '12px'}}
              />
              <ReferenceLine y={12} label={{ value: 'Limit', position: 'right', fontSize: 10, fill: '#374151' }} stroke="#374151" strokeDasharray="3 3" />
              <ReferenceLine y={0} label={{ value: 'Low', position: 'right', fontSize: 10, fill: '#374151' }} stroke="#374151" strokeDasharray="3 3" />
              <ReferenceLine y={5} stroke="#3b82f6" strokeWidth={2} />
              <Line type="monotone" dataKey="range" stroke="#0d9488" strokeWidth={2} dot={{r: 3, fill: '#0d9488', strokeWidth: 1, stroke: '#fff'}} activeDot={{ r: 6 }} />
            </LineChart>
          </ResponsiveContainer>
        </div>
      </div>
    </div>
  );
};