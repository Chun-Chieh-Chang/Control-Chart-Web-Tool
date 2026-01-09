// ============================================================================
// Batch Analysis Module - X-Bar R Control Charts
// ============================================================================

import { SPCEngine } from './spc-engine.js';

export class BatchAnalysis {
    constructor(dataInput, language = 'zh') {
        this.dataInput = dataInput;
        this.lang = language;
        this.results = null;
    }

    async execute() {
        // Get data matrix [batch][cavity]
        const dataMatrix = this.dataInput.getDataMatrix();
        const batchNames = this.dataInput.batchNames;
        const specs = this.dataInput.getSpecs();
        const cavityCount = this.dataInput.getCavityCount();

        // Calculate X-Bar R control limits
        const xbarR = SPCEngine.calculateXBarRLimits(dataMatrix);

        // Detect out-of-control points
        const xBarOOC = SPCEngine.detectOutOfControlPoints(
            xbarR.xBar.data,
            xbarR.xBar.UCL,
            xbarR.xBar.LCL
        );

        const rOOC = SPCEngine.detectOutOfControlPoints(
            xbarR.R.data,
            xbarR.R.UCL,
            xbarR.R.LCL
        );

        // Generate HTML results
        const html = this.generateHTML(xbarR, batchNames, specs, xBarOOC, rOOC);

        // Generate chart configurations
        const charts = this.generateCharts(xbarR, batchNames, specs, xBarOOC, rOOC);

        this.results = {
            type: 'batch',
            html: html,
            charts: charts,
            data: {
                xbarR: xbarR,
                batchNames: batchNames,
                specs: specs,
                cavityCount: cavityCount
            }
        };

        return this.results;
    }

    generateHTML(xbarR, batchNames, specs, xBarOOC, rOOC) {
        const t = (zh, en) => this.lang === 'zh' ? zh : en;

        // Summary statistics
        const summaryHTML = `
            <div class="results-summary">
                <div class="stat-card">
                    <div class="stat-label">${t('樣本數 (n)', 'Sample Size (n)')}</div>
                    <div class="stat-value">${xbarR.summary.n}</div>
                </div>
                <div class="stat-card">
                    <div class="stat-label">${t('批號數 (k)', 'Subgroups (k)')}</div>
                    <div class="stat-value">${xbarR.summary.k}</div>
                </div>
                <div class="stat-card">
                    <div class="stat-label">${t('總平均', 'Grand Average')} (X̿)</div>
                    <div class="stat-value">${SPCEngine.round(xbarR.summary.xDoubleBar, 4)}</div>
                </div>
                <div class="stat-card">
                    <div class="stat-label">${t('平均全距', 'Average Range')} (R̄)</div>
                    <div class="stat-value">${SPCEngine.round(xbarR.summary.rBar, 4)}</div>
                </div>
            </div>
        `;

        // Control limits table
        const limitsHTML = `
            <h3 class="chart-title">${t('管制界限', 'Control Limits')}</h3>
            <table class="data-table">
                <thead>
                    <tr>
                        <th>${t('圖表', 'Chart')}</th>
                        <th>${t('上管制限 (UCL)', 'Upper Control Limit')}</th>
                        <th>${t('中心線 (CL)', 'Center Line')}</th>
                        <th>${t('下管制限 (LCL)', 'Lower Control Limit')}</th>
                    </tr>
                </thead>
                <tbody>
                    <tr>
                        <td><strong>X̄ ${t('圖', 'Chart')}</strong></td>
                        <td>${SPCEngine.round(xbarR.xBar.UCL, 4)}</td>
                        <td>${SPCEngine.round(xbarR.xBar.CL, 4)}</td>
                        <td>${SPCEngine.round(xbarR.xBar.LCL, 4)}</td>
                    </tr>
                    <tr>
                        <td><strong>R ${t('圖', 'Chart')}</strong></td>
                        <td>${SPCEngine.round(xbarR.R.UCL, 4)}</td>
                        <td>${SPCEngine.round(xbarR.R.CL, 4)}</td>
                        <td>${SPCEngine.round(xbarR.R.LCL, 4)}</td>
                    </tr>
                </tbody>
            </table>
        `;

        // Specification limits
        const specsHTML = `
            <h3 class="chart-title">${t('規格界限', 'Specification Limits')}</h3>
            <table class="data-table">
                <thead>
                    <tr>
                        <th>${t('目標值 (Target)', 'Target')}</th>
                        <th>${t('上規格限 (USL)', 'Upper Spec Limit')}</th>
                        <th>${t('下規格限 (LSL)', 'Lower Spec Limit')}</th>
                    </tr>
                </thead>
                <tbody>
                    <tr>
                        <td>${specs.target}</td>
                        <td>${specs.usl}</td>
                        <td>${specs.lsl}</td>
                    </tr>
                </tbody>
            </table>
        `;

        // Charts placeholder
        const chartsHTML = `
            <div class="chart-container">
                <h3 class="chart-title">${t('X̄ 管制圖 (平均值)', 'X-Bar Control Chart (Average)')}</h3>
                <canvas id="xbarChart"></canvas>
            </div>
            <div class="chart-container">
                <h3 class="chart-title">${t('R 管制圖 (全距)', 'R Control Chart (Range)')}</h3>
                <canvas id="rChart"></canvas>
            </div>
        `;

        // Out-of-control alerts
        const oocCount = xBarOOC.filter(p => p.outOfControl).length + rOOC.filter(p => p.outOfControl).length;
        let alertHTML = '';
        if (oocCount > 0) {
            alertHTML = `
                <div style="padding: 1rem; background: rgba(239, 68, 68, 0.1); border: 2px solid #ef4444; border-radius: 0.75rem; margin: 2rem 0;">
                    <h3 style="color: #ef4444; margin-bottom: 0.5rem;">⚠️ ${t('偵測到管制界限外的點', 'Out-of-Control Points Detected')}</h3>
                    <p>${t('共', 'Total')} ${oocCount} ${t('個點超出管制界限，請檢查製程。', 'points are out of control. Please review the process.')}</p>
                </div>
            `;
        }

        return summaryHTML + specsHTML + limitsHTML + alertHTML + chartsHTML;
    }

    generateCharts(xbarR, batchNames, specs, xBarOOC, rOOC) {
        const t = (zh, en) => this.lang === 'zh' ? zh : en;

        // Colors for points (red if out of control)
        const xBarColors = xBarOOC.map(p => p.outOfControl ? '#ef4444' : '#3b82f6');
        const rColors = rOOC.map(p => p.outOfControl ? '#ef4444' : '#3b82f6');

        return [
            {
                canvasId: 'xbarChart',
                config: {
                    type: 'line',
                    data: {
                        labels: batchNames.slice(0, xbarR.xBar.data.length),
                        datasets: [
                            {
                                label: t('平均值', 'X-Bar'),
                                data: xbarR.xBar.data,
                                borderColor: '#3b82f6',
                                backgroundColor: xBarColors,
                                pointBackgroundColor: xBarColors,
                                pointBorderColor: xBarColors,
                                pointRadius: 5,
                                pointHoverRadius: 7,
                                borderWidth: 2,
                                tension: 0.1
                            },
                            {
                                label: t('上管制限 (UCL)', 'UCL'),
                                data: Array(xbarR.xBar.data.length).fill(xbarR.xBar.UCL),
                                borderColor: '#ef4444',
                                borderDash: [5, 5],
                                borderWidth: 2,
                                pointRadius: 0,
                                fill: false
                            },
                            {
                                label: t('中心線 (CL)', 'CL'),
                                data: Array(xbarR.xBar.data.length).fill(xbarR.xBar.CL),
                                borderColor: '#10b981',
                                borderWidth: 2,
                                pointRadius: 0,
                                fill: false
                            },
                            {
                                label: t('下管制限 (LCL)', 'LCL'),
                                data: Array(xbarR.xBar.data.length).fill(xbarR.xBar.LCL),
                                borderColor: '#ef4444',
                                borderDash: [5, 5],
                                borderWidth: 2,
                                pointRadius: 0,
                                fill: false
                            }
                        ]
                    },
                    options: {
                        responsive: true,
                        maintainAspectRatio: true,
                        aspectRatio: 2,
                        plugins: {
                            legend: {
                                position: 'top',
                                labels: {
                                    font: { size: 12, family: 'Microsoft JhengHei' }
                                }
                            },
                            title: {
                                display: false
                            }
                        },
                        scales: {
                            y: {
                                beginAtZero: false,
                                title: {
                                    display: true,
                                    text: t('平均數', 'Average'),
                                    font: { size: 14, weight: 'bold' }
                                }
                            },
                            x: {
                                title: {
                                    display: true,
                                    text: t('樣本號數', 'Sample Number'),
                                    font: { size: 14, weight: 'bold' }
                                }
                            }
                        }
                    }
                }
            },
            {
                canvasId: 'rChart',
                config: {
                    type: 'line',
                    data: {
                        labels: batchNames.slice(0, xbarR.R.data.length),
                        datasets: [
                            {
                                label: t('全距', 'Range (R)'),
                                data: xbarR.R.data,
                                borderColor: '#3b82f6',
                                backgroundColor: rColors,
                                pointBackgroundColor: rColors,
                                pointBorderColor: rColors,
                                pointRadius: 5,
                                pointHoverRadius: 7,
                                borderWidth: 2,
                                tension: 0.1
                            },
                            {
                                label: t('上管制限 (UCL)', 'UCL'),
                                data: Array(xbarR.R.data.length).fill(xbarR.R.UCL),
                                borderColor: '#ef4444',
                                borderDash: [5, 5],
                                borderWidth: 2,
                                pointRadius: 0,
                                fill: false
                            },
                            {
                                label: t('中心線 (CL)', 'CL'),
                                data: Array(xbarR.R.data.length).fill(xbarR.R.CL),
                                borderColor: '#10b981',
                                borderWidth: 2,
                                pointRadius: 0,
                                fill: false
                            }
                        ]
                    },
                    options: {
                        responsive: true,
                        maintainAspectRatio: true,
                        aspectRatio: 2,
                        plugins: {
                            legend: {
                                position: 'top',
                                labels: {
                                    font: { size: 12, family: 'Microsoft JhengHei' }
                                }
                            },
                            title: {
                                display: false
                            }
                        },
                        scales: {
                            y: {
                                beginAtZero: true,
                                title: {
                                    display: true,
                                    text: t('全距', 'Range'),
                                    font: { size: 14, weight: 'bold' }
                                }
                            },
                            x: {
                                title: {
                                    display: true,
                                    text: t('樣本號數', 'Sample Number'),
                                    font: { size: 14, weight: 'bold' }
                                }
                            }
                        }
                    }
                }
            }
        ];
    }
}
