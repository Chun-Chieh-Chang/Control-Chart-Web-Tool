/**
 * QIP Data Extract - Application Controller with Mouse Range Selection
 * VBA-like range selection using mouse drag on preview table
 */

var QIPExtractApp = {
    workbook: null,
    fileName: '',
    processingResults: null,
    selectionMode: null,  // 'cavity' or 'data'
    selectionTarget: null,  // target input element ID
    selectionStart: null,
    selectionEnd: null,

    init: function () {
        console.log('QIPExtractApp initializing...');
        this.cacheElements();
        this.bindEvents();
        this.loadSavedConfigs();
    },

    cacheElements: function () {
        this.els = {
            uploadZone: document.getElementById('qip-upload-zone'),
            fileInput: document.getElementById('qip-file-input'),
            fileInfo: document.getElementById('qip-file-info'),
            fileName: document.getElementById('qip-file-name'),
            removeFile: document.getElementById('qip-remove-file'),

            productCode: document.getElementById('qip-product-code'),
            cavityCount: document.getElementById('qip-cavity-count'),
            worksheetSelectGroup: document.getElementById('qip-worksheet-select-group'),
            worksheetSelect: document.getElementById('qip-worksheet-select'),
            previewBtn: document.getElementById('qip-preview-btn'),
            cavityGroups: document.getElementById('qip-cavity-groups'),

            configName: document.getElementById('qip-config-name'),
            saveConfig: document.getElementById('qip-save-config'),
            loadConfig: document.getElementById('qip-load-config'),

            startProcess: document.getElementById('qip-start-process'),
            progress: document.getElementById('qip-progress'),
            progressBar: document.getElementById('qip-progress-bar'),
            progressText: document.getElementById('qip-progress-text'),

            resultSection: document.getElementById('qip-result-section'),
            resultSummary: document.getElementById('qip-result-summary'),
            downloadExcel: document.getElementById('qip-download-excel'),
            sendToSpc: document.getElementById('qip-send-to-spc'),
            errorLog: document.getElementById('qip-error-log'),
            errorList: document.getElementById('qip-error-list'),

            previewPanel: document.getElementById('qip-preview-panel'),
            previewContent: document.getElementById('qip-preview-content'),
            selectionModeIndicator: document.getElementById('qip-selection-mode'),
            selectionTypeText: document.getElementById('qip-selection-type'),
            prevSheet: document.getElementById('qip-prev-sheet'),
            nextSheet: document.getElementById('qip-next-sheet')
        };
    },
