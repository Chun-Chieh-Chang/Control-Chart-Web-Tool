/**
 * ErrorLogger - 錯誤日誌模組
 * 對應 VBA ErrorLogger.bas
 */
class ErrorLogger {
    constructor() {
        this.errors = [];
    }

    /**
     * 記錄錯誤信息
     * @param {string} sheetName - 工作表名稱
     * @param {string} errorMessage - 錯誤訊息
     */
    logError(sheetName, errorMessage) {
        const errorType = this.determineErrorType(errorMessage);
        this.errors.push({
            sheetName,
            errorType,
            errorMessage,
            timestamp: new Date().toISOString()
        });
        console.error(`[${errorType}] ${sheetName}: ${errorMessage}`);
    }

    /**
     * 確定錯誤類型
     * @param {string} errorMessage 
     * @returns {string}
     */
    determineErrorType(errorMessage) {
        const msg = errorMessage.toLowerCase();

        if (msg.includes('格式') || msg.includes('format')) {
            return '格式錯誤';
        } else if (msg.includes('數據') || msg.includes('data')) {
            return '數據錯誤';
        } else if (msg.includes('圖表') || msg.includes('chart')) {
            return '圖表錯誤';
        } else if (msg.includes('工作表') || msg.includes('worksheet') || msg.includes('sheet')) {
            return '工作表錯誤';
        } else if (msg.includes('範圍') || msg.includes('range')) {
            return '範圍錯誤';
        } else {
            return '一般錯誤';
        }
    }

    /**
     * 檢查是否有錯誤記錄
     * @returns {boolean}
     */
    hasErrors() {
        return this.errors.length > 0;
    }

    /**
     * 獲取所有錯誤
     * @returns {Array}
     */
    getErrors() {
        return [...this.errors];
    }

    /**
     * 獲取錯誤數量
     * @returns {number}
     */
    getErrorCount() {
        return this.errors.length;
    }

    /**
     * 清除錯誤日誌
     */
    clearErrors() {
        this.errors = [];
    }

    /**
     * 獲取錯誤摘要
     * @returns {Object}
     */
    getSummary() {
        const summary = {
            total: this.errors.length,
            byType: {},
            bySheet: {}
        };

        for (const error of this.errors) {
            // 按類型統計
            if (!summary.byType[error.errorType]) {
                summary.byType[error.errorType] = 0;
            }
            summary.byType[error.errorType]++;

            // 按工作表統計
            if (!summary.bySheet[error.sheetName]) {
                summary.bySheet[error.sheetName] = 0;
            }
            summary.bySheet[error.sheetName]++;
        }

        return summary;
    }

    /**
     * 格式化錯誤列表為 HTML
     * @returns {string}
     */
    toHTML() {
        if (this.errors.length === 0) {
            return '<p>無錯誤記錄</p>';
        }

        let html = '<table class="error-table">';
        html += '<thead><tr><th>工作表</th><th>類型</th><th>訊息</th><th>時間</th></tr></thead>';
        html += '<tbody>';

        for (const error of this.errors) {
            const time = new Date(error.timestamp).toLocaleTimeString();
            html += `<tr>
                <td>${this.escapeHtml(error.sheetName)}</td>
                <td>${this.escapeHtml(error.errorType)}</td>
                <td>${this.escapeHtml(error.errorMessage)}</td>
                <td>${time}</td>
            </tr>`;
        }

        html += '</tbody></table>';
        return html;
    }

    /**
     * HTML 轉義
     * @param {string} text 
     * @returns {string}
     */
    escapeHtml(text) {
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    }
}

// 導出供其他模組使用
if (typeof module !== 'undefined' && module.exports) {
    module.exports = ErrorLogger;
}
