import React from 'react';
import { Header } from './components/Header';
import { Sidebar } from './components/Sidebar';
import { StatCard } from './components/StatCard';
import { Charts } from './components/Charts';
import { DataTable } from './components/DataTable';
import { AnomalyList } from './components/AnomalyList';
import { MOCK_KPIS, MOCK_CHART_DATA, MOCK_MEASUREMENTS, MOCK_ANOMALIES } from './constants';

const App: React.FC = () => {
  return (
    <div className="flex h-screen w-full bg-slate-50 text-gray-900 font-sans overflow-hidden">
      {/* Sidebar Configuration */}
      <Sidebar />

      {/* Main Content Area */}
      <div className="flex-1 flex flex-col min-w-0">
        <Header />

        <main className="flex-1 overflow-y-auto p-6 scroll-smooth">
          {/* Section Header */}
          <div className="mb-6">
             <h1 className="text-2xl font-bold text-gray-800">Analysis Workspace</h1>
          </div>

          {/* KPI Cards Row */}
          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-5 gap-4 mb-6">
            {MOCK_KPIS.map((kpi, index) => (
              <StatCard key={index} data={kpi} />
            ))}
          </div>

          {/* Charts Row */}
          <Charts data={MOCK_CHART_DATA} />

          {/* Bottom Row: Data Table and Anomaly List */}
          <div className="grid grid-cols-1 xl:grid-cols-2 gap-6 h-[500px]">
             <DataTable measurements={MOCK_MEASUREMENTS} />
             <AnomalyList anomalies={MOCK_ANOMALIES} contextData={MOCK_CHART_DATA.slice(-5)} />
          </div>
        </main>
      </div>
    </div>
  );
};

export default App;