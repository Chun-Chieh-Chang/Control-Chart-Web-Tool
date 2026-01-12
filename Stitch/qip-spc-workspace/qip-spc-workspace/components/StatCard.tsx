import React from 'react';
import { TrendingUp, Activity, AlertTriangle, CheckCircle, Minus } from 'lucide-react';
import { KPI } from '../types';

interface StatCardProps {
  data: KPI;
}

export const StatCard: React.FC<StatCardProps> = ({ data }) => {
  const getIcon = () => {
    switch (data.icon) {
      case 'trend-up': return <TrendingUp className="w-5 h-5 text-green-500" />;
      case 'check-circle': return <CheckCircle className="w-6 h-6 text-green-600" />;
      case 'minus': return <Minus className="w-6 h-6 text-teal-500" />;
      case 'activity': return <Activity className="w-6 h-6 text-teal-600" />;
      case 'alert-triangle': return <AlertTriangle className="w-6 h-6 text-red-600" />;
      default: return null;
    }
  };

  const isWarning = data.status === 'bad';

  return (
    <div className="bg-white p-4 rounded-lg shadow-sm border border-gray-100 flex flex-col justify-between h-28 relative overflow-hidden">
      <div className="flex justify-between items-start">
        <div>
          <h3 className="text-sm font-semibold text-gray-600">{data.label}</h3>
          <div className={`text-2xl font-bold mt-1 ${isWarning ? 'text-red-600' : 'text-gray-900'}`}>{data.value}</div>
        </div>
        <div className={`p-1 rounded-full ${isWarning ? 'bg-red-50' : ''}`}>
           {getIcon()}
        </div>
      </div>
      <div className="text-xs text-gray-500 font-medium">
        {data.subValue}
      </div>
      {/* Decorative line/wave for specific cards based on screenshot */}
      {data.label === 'Standard Deviation' && (
          <div className="absolute bottom-4 right-4 text-teal-500 opacity-50">
             <Activity className="w-8 h-8" />
          </div>
      )}
    </div>
  );
};