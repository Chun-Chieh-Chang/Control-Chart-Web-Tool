import React, { useState } from 'react';
import { AlertTriangle, Bot } from 'lucide-react';
import { Anomaly } from '../types';
import { analyzeAnomalyWithAI } from '../services/geminiService';

interface AnomalyListProps {
  anomalies: Anomaly[];
  contextData: any;
}

export const AnomalyList: React.FC<AnomalyListProps> = ({ anomalies, contextData }) => {
  const [analyzingId, setAnalyzingId] = useState<string | null>(null);
  const [analysisResult, setAnalysisResult] = useState<{id: string, text: string} | null>(null);

  const handleInvestigate = async (anomaly: Anomaly) => {
    setAnalyzingId(anomaly.id);
    const result = await analyzeAnomalyWithAI(anomaly, contextData);
    setAnalysisResult({ id: anomaly.id, text: result });
    setAnalyzingId(null);
  };

  const closeAnalysis = () => {
    setAnalysisResult(null);
  }

  return (
    <div className="bg-white rounded-lg shadow-sm border border-gray-200 flex flex-col h-full relative">
      <div className="p-4 border-b border-gray-200">
        <h3 className="font-bold text-gray-800">Quality Summary/Anomaly List</h3>
      </div>
      
      {/* Summary Stat Bar */}
      <div className="grid grid-cols-3 divide-x divide-gray-200 border-b border-gray-200 bg-gray-50">
          <div className="p-3">
              <div className="text-xs text-gray-500 font-medium uppercase">Total Samples</div>
              <div className="text-xl font-bold text-gray-800">1,200</div>
          </div>
          <div className="p-3">
              <div className="text-xs text-gray-500 font-medium uppercase">Anomalies</div>
              <div className="text-xl font-bold text-red-600 flex items-center gap-2">
                  15 <span className="text-sm font-normal text-red-500">(1.25%)</span>
              </div>
          </div>
          <div className="p-3">
              <div className="text-xs text-gray-500 font-medium uppercase">Process Capability (Cp)</div>
              <div className="text-xl font-bold text-gray-800">1.33</div>
          </div>
      </div>

      <div className="overflow-auto custom-scrollbar flex-1 p-2 space-y-2">
        {anomalies.map((anomaly) => (
            <div key={anomaly.id} className="flex items-center justify-between p-3 rounded-md bg-white border border-gray-100 hover:shadow-md transition-shadow">
                <div className="flex items-start gap-3">
                    <AlertTriangle className="w-5 h-5 text-red-600 mt-0.5 shrink-0" />
                    <div>
                        <div className="text-xs text-gray-500 font-mono mb-0.5">{anomaly.timestamp}</div>
                        <div className="text-sm font-medium text-gray-800">{anomaly.message}</div>
                    </div>
                </div>
                <div className="flex gap-2 shrink-0">
                    <button 
                      onClick={() => handleInvestigate(anomaly)}
                      disabled={analyzingId === anomaly.id}
                      className="px-3 py-1.5 bg-white border border-gray-300 text-gray-700 text-xs font-medium rounded hover:bg-gray-50 flex items-center gap-1 transition-all"
                    >
                       {analyzingId === anomaly.id ? (
                         <span className="animate-pulse">AI Analyzing...</span>
                       ) : (
                         <>
                           <Bot className="w-3 h-3 text-indigo-600" />
                           Investigate
                         </>
                       )}
                    </button>
                    <button className="px-3 py-1.5 bg-white border border-gray-300 text-gray-700 text-xs font-medium rounded hover:bg-gray-50">
                        Acknowledge
                    </button>
                </div>
            </div>
        ))}
      </div>

      {/* AI Analysis Modal Overlay */}
      {analysisResult && (
        <div className="absolute inset-0 bg-black/10 backdrop-blur-[1px] flex items-center justify-center p-4 z-20">
            <div className="bg-white rounded-lg shadow-xl border border-indigo-100 w-full max-w-md overflow-hidden animate-in fade-in zoom-in duration-200">
                <div className="bg-indigo-900 text-white px-4 py-3 flex justify-between items-center">
                    <div className="flex items-center gap-2">
                        <Bot className="w-5 h-5" />
                        <h4 className="font-semibold text-sm">AI Root Cause Analysis</h4>
                    </div>
                    <button onClick={closeAnalysis} className="text-indigo-200 hover:text-white">&times;</button>
                </div>
                <div className="p-4">
                    <div className="text-sm text-gray-700 leading-relaxed space-y-2">
                        {analysisResult.text.split('\n').map((line, i) => (
                            <p key={i}>{line}</p>
                        ))}
                    </div>
                    <div className="mt-4 flex justify-end">
                        <button onClick={closeAnalysis} className="px-4 py-2 bg-indigo-600 text-white text-sm rounded hover:bg-indigo-700">
                            Close & Log
                        </button>
                    </div>
                </div>
            </div>
        </div>
      )}
    </div>
  );
};