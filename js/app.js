// ============================================================================
// Main Application Controller
// ============================================================================

import { DataInput } from './data-input.js';
import { BatchAnalysis } from './batch-analysis.js';
import { CavityAnalysis } from './cavity-analysis.js';
import { GroupAnalysis } from './group-analysis.js';
import { ExcelExport } from './excel-export.js';

class SPCApp {
    constructor() {
        this.currentLanguage = 'zh';
        this.workbook = null;
        this.selectedItem = null;
        this.analysisResults = null;
        
        this.init();
    }
    
    init() {
        this.setupLanguageToggle();
        this.setupFileUpload();
        this.setupEventListeners();
        this.updateLanguage();
    }
    
    // ========================================================================
    // Language Support
    // ========================================================================
    
    setupLanguageToggle() {
        const langBtn = document.getElementById('langBtn');
        const langText = document.getElementById('langText');
        
        langBtn.addEventListener('click', () => {
            this.currentLanguage = this.currentLanguage === 'zh' ? 'en' : 'zh';
            langText.textContent = this.currentLanguage === 'zh' ? 'EN' : '中文';
            this.updateLanguage();
        });
    }
    
    updateLanguage() {
        const elements = document.querySelectorAll('[data-en][data-zh]');
        elements.forEach(el => {
            const text = this.currentLanguage === 'zh' ? el.dataset.zh : el.dataset.en;
            if (el.tagName === 'BUTTON' || el.tagName === 'P' || el.tagName === 'H1' || 
                el.tagName === 'H2' || el.tagName === 'H3' || el.tagName === 'SPAN') {
                el.textContent = text;
            } else {
                el.innerHTML = text;
            }
        });
    }
    
    t(zh, en) {
        return this.currentLanguage === 'zh' ? zh : en;
    }
    
    // ========================================================================
    // File Upload
    // ========================================================================
    
    setupFileUpload() {
        const uploadZone = document.getElementById('uploadZone');
        const fileInput = document.getElementById('fileInput');
        const fileInfo = document.getElementById('fileInfo');
        const fileName = document.getElementById('fileName');
        const resetBtn = document.getElementById('resetBtn');
        
        // Click to upload
        uploadZone.addEventListener('click', () => fileInput.click());
        
        // Drag and drop
        uploadZone.addEventListener('dragover', (e) => {
            e.preventDefault();
            uploadZone.classList.add('dragover');
        });
        
        uploadZone.addEventListener('dragleave', () => {
            uploadZone.classList.remove('dragover');
        });
        
        uploadZone.addEventListener('drop', (e) => {
            e.preventDefault();
            uploadZone.classList.remove('dragover');
            const files = e.dataTransfer.files;
            if (files.length > 0) {
                this.handleFile(files[0]);
            }
        });
        
        // File input change
        fileInput.addEventListener('change', (e) => {
            if (e.target.files.length > 0) {
                this.handleFile(e.target.files[0]);
            }
        });
        
        // Reset button
        resetBtn.addEventListener('click', () => {
            this.resetApp();
        });
    }
    
    async handleFile(file) {
        if (!file.name.match(/\.(xlsx|xls)$/i)) {
            alert(this.t('請選擇 Excel 檔案 (.xlsx 或 .xls)', 'Please select an Excel file (.xlsx or .xls)'));
            return;
        }
        
        this.showLoading(this.t('讀取檔案中...', 'Reading file...'));
        
        try {
            // Read file using SheetJS
            const arrayBuffer = await file.arrayBuffer();
            this.workbook = XLSX.read(new Uint8Array(arrayBuffer), { type: 'array' });
            
            // Update UI
            document.getElementById('uploadZone').style.display = 'none';
            document.getElementById('fileInfo').style.display = 'flex';
            document.getElementById('fileName').textContent = file.name;
            
            // Show inspection item selection
            this.showInspectionItems();
            
            this.hideLoading();
        } catch (error) {
            this.hideLoading();
            alert(this.t('檔案讀取失敗：', 'File reading failed: ') + error.message);
            console.error(error);
        }
    }
    
    // ========================================================================
    // Inspection Item Selection
    // ========================================================================
    
    showInspectionItems() {
        const itemList = document.getElementById('itemList');
        itemList.innerHTML = '';
        
        const sheets = this.workbook.SheetNames.filter(name => {
            // Exclude summary sheets
            return !['摘要', 'Summary', '統計', '說明'].includes(name) &&
                   !name.match(/分析|配置/);
        });
        
        sheets.forEach(sheetName => {
            const ws = this.workbook.Sheets[sheetName];
            const data = XLSX.utils.sheet_to_json(ws, { header: 1 });
            
            // Get target value from row 2, column B
            const target = data[1] && data[1][1] ? data[1][1] : 'N/A';
            
            const card = document.createElement('div');
            card.className = 'item-card';
            card.innerHTML = `
                <div class="item-info">
                    <h3>${sheetName}</h3>
                    <p>Target: ${target}</p>
                </div>
                <div class="item-badge">${this.t('選擇', 'Select')}</div>
            `;
            
            card.addEventListener('click', () => {
                this.selectedItem = sheetName;
                this.showAnalysisOptions();
            });
            
            itemList.appendChild(card);
        });
        
        document.getElementById('step2').style.display = 'block';
        document.getElementById('step2').scrollIntoView({ behavior: 'smooth' });
    }
    
    // ========================================================================
    // Analysis Type Selection
    // ========================================================================
    
    showAnalysisOptions() {
        document.getElementById('step3').style.display = 'block';
        document.getElementById('step3').scrollIntoView({ behavior: 'smooth' });
    }
    
    setupEventListeners() {
        // Analysis buttons
        document.querySelectorAll('[data-analysis]').forEach(btn => {
            btn.addEventListener('click', (e) => {
                const analysisType = e.target.dataset.analysis;
                this.executeAnalysis(analysisType);
            });
        });
        
        // Download buttons
        document.getElementById('downloadExcel')?.addEventListener('click', () => {
            this.downloadExcel();
        });
        
        document.getElementById('downloadCharts')?.addEventListener('click', () => {
            this.downloadCharts();
        });
    }
    
    // ========================================================================
    // Execute Analysis
    // ========================================================================
    
    async executeAnalysis(type) {
        this.showLoading(this.t('分析中...', 'Analyzing...'));
        
        try {
            const ws = this.workbook.Sheets[this.selectedItem];
            const dataInput = new DataInput(ws);
            
            let analysis;
            switch (type) {
                case 'batch':
                    analysis = new BatchAnalysis(dataInput, this.currentLanguage);
                    break;
                case 'cavity':
                    analysis = new CavityAnalysis(dataInput, this.currentLanguage);
                    break;
                case 'group':
                    analysis = new GroupAnalysis(dataInput, this.currentLanguage);
                    break;
            }
            
            this.analysisResults = await analysis.execute();
            this.displayResults(type);
            
            this.hideLoading();
        } catch (error) {
            this.hideLoading();
            alert(this.t('分析失敗：', 'Analysis failed: ') + error.message);
            console.error(error);
        }
    }
    
    displayResults(type) {
        const resultsSection = document.getElementById('results');
        const resultsContent = document.getElementById('resultsContent');
        
        resultsContent.innerHTML = this.analysisResults.html;
        
        // Render charts
        if (this.analysisResults.charts) {
            this.analysisResults.charts.forEach((chartConfig, index) => {
                const canvas = document.getElementById(chartConfig.canvasId);
                if (canvas) {
                    new Chart(canvas, chartConfig.config);
                }
            });
        }
        
        resultsSection.style.display = 'block';
        resultsSection.scrollIntoView({ behavior: 'smooth' });
    }
    
    // ========================================================================
    // Export Functions
    // ========================================================================
    
    async downloadExcel() {
        this.showLoading(this.t('生成 Excel 中...', 'Generating Excel...'));
        
        try {
            const exporter = new ExcelExport(this.analysisResults, this.selectedItem, this.currentLanguage);
            await exporter.generate();
            
            this.hideLoading();
        } catch (error) {
            this.hideLoading();
            alert(this.t('Excel 生成失敗：', 'Excel generation failed: ') + error.message);
            console.error(error);
        }
    }
    
    async downloadCharts() {
        this.showLoading(this.t('匯出圖表中...', 'Exporting charts...'));
        
        try {
            // Export all charts as PNG images
            const charts = document.querySelectorAll('.chart-container canvas');
            const zip = new JSZip();
            
            for (let i = 0; i < charts.length; i++) {
                const canvas = charts[i];
                const dataUrl = canvas.toDataURL('image/png');
                const base64 = dataUrl.split(',')[1];
                zip.file(`chart_${i + 1}.png`, base64, { base64: true });
            }
            
            const content = await zip.generateAsync({ type: 'blob' });
            const link = document.createElement('a');
            link.href = URL.createObjectURL(content);
            link.download = `SPC_Charts_${new Date().toISOString().split('T')[0]}.zip`;
            link.click();
            
            this.hideLoading();
        } catch (error) {
            this.hideLoading();
            alert(this.t('圖表匯出失敗：', 'Chart export failed: ') + error.message);
            console.error(error);
        }
    }
    
    // ========================================================================
    // UI Helpers
    // ========================================================================
    
    showLoading(text) {
        const overlay = document.getElementById('loadingOverlay');
        const loadingText = document.getElementById('loadingText');
        loadingText.textContent = text;
        overlay.classList.add('active');
    }
    
    hideLoading() {
        document.getElementById('loadingOverlay').classList.remove('active');
    }
    
    resetApp() {
        this.workbook = null;
        this.selectedItem = null;
        this.analysisResults = null;
        
        document.getElementById('fileInput').value = '';
        document.getElementById('uploadZone').style.display = 'block';
        document.getElementById('fileInfo').style.display = 'none';
        document.getElementById('step2').style.display = 'none';
        document.getElementById('step3').style.display = 'none';
        document.getElementById('results').style.display = 'none';
        
        window.scrollTo({ top: 0, behavior: 'smooth' });
    }
}

// Initialize app when DOM is ready
document.addEventListener('DOMContentLoaded', () => {
    window.spcApp = new SPCApp();
});
