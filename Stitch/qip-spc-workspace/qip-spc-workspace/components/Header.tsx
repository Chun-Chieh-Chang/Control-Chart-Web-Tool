import React from 'react';
import { Search, Bell, User, Menu } from 'lucide-react';

export const Header: React.FC = () => {
  return (
    <header className="h-16 bg-indigo-950 text-white flex items-center justify-between px-6 shrink-0">
      <div className="flex items-center gap-8">
        <div className="flex items-center gap-2">
           <div className="w-8 h-8 rounded bg-gradient-to-br from-blue-400 to-indigo-600 flex items-center justify-center font-bold text-lg">Q</div>
           <span className="font-semibold text-lg tracking-tight">QIP SPC Workspace</span>
        </div>
        
        <nav className="hidden md:flex items-center space-x-1">
            <a href="#" className="px-4 py-2 text-sm text-gray-300 hover:text-white transition-colors">Dashboard</a>
            <a href="#" className="px-4 py-2 text-sm bg-indigo-800 rounded text-white font-medium">Analysis</a>
            <a href="#" className="px-4 py-2 text-sm text-gray-300 hover:text-white transition-colors">Reports</a>
            <a href="#" className="px-4 py-2 text-sm text-gray-300 hover:text-white transition-colors">Settings</a>
        </nav>
      </div>

      <div className="flex items-center gap-4">
        <div className="relative">
            <Bell className="w-5 h-5 text-gray-300 hover:text-white cursor-pointer" />
            <span className="absolute -top-1 -right-1 w-2 h-2 bg-pink-500 rounded-full"></span>
        </div>
        <div className="flex items-center gap-2 cursor-pointer hover:bg-indigo-800 px-2 py-1 rounded transition-colors">
            <div className="w-8 h-8 bg-gray-300 rounded-full flex items-center justify-center text-gray-700">
                <User className="w-5 h-5" />
            </div>
            <span className="text-sm font-medium hidden sm:block">John Doe</span>
            <ChevronDownIcon className="w-3 h-3 text-gray-400" />
        </div>
      </div>
    </header>
  );
};

const ChevronDownIcon = ({ className }: { className?: string }) => (
    <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className={className}><path d="m6 9 6 6 6-6"/></svg>
)
