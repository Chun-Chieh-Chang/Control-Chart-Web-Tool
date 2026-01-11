// ============================================================================
// Group Analysis Module - Min-Max-Avg Control Charts
// ============================================================================

import { SPCEngine } from './spc-engine.js';

export class GroupAnalysis {
    constructor(dataInput, language = 'zh') {
        this.dataInput = dataInput;
        this.lang = language;
        this.results = null;
    }

    async execute() {
        const specs = this.dataInput.getSpecs();
        const batchNames = this.dataInput.batchNames;
        const dataMatrix = this.dataInput.getDataMatrix();

        // Calculate Min, Max, Avg for each batch
        const groupStats = dataMatrix.map((batchData, index) => {
            const filtered = batchData.filter(v => v !== null && !isNaN(v));

            if (filtered.length === 0) {
                return {
                    batch: batchNames[index] || `Batch ${index + 1}`,
                    min: 0,
                    max: 0,
                    avg: 0,
                    range: 0,
                    count: 0
                };
            }

            const min = SPCEngine.min(filtered);
            const max = SPCEngine.max(filtered);
            const avg = SPCEngine.mean(filtered);

            return {
                batch: batchNames[index] || `Batch ${index + 1}`,
                min: min,
                max: max,
                avg: avg,
                range: max - min,
                count: filtered.length
            };
        });

        // Generate HTML results
        const html = this.generateHTML(groupStats, specs);

        // Generate chart configurations
        const charts = this.generateCharts(groupStats, specs);

        this.results = {
            type: 'group',
            html: html,
            charts: charts,
            data: {
                groupStats: groupStats,
                specs: specs
            }
        };

        return this.results;
    }

    generateHTML(groupStats, specs) {
        const t = (zh, en) => this.lang === 'zh' ? zh : en;

        // Summary
        const avgOfAvgs = SPCEngine.mean(groupStats.map(s => s.avg));
        const avgRange = SPCEngine.mean(groupStats.map(s => s.range));
        const maxRange = SPCEngine.max(groupStats.map(s => s.range));

        const summaryHTML = `
            <div class="results-summary">
                <div class="stat-card">
                    <div class="stat-label">${t('批號數量', 'Batch Count')}</div>
                    <div class="stat-value">${groupStats.length}</div>
                </div>
                <div class="stat-card">
                    <div class="stat-label">${t('總平均', 'Overall Average')}</div>
                    <div class="stat-value">${SPCEngine.round(avgOfAvgs, 4)}</div>
                </div>
                <div class="stat-card">
                    <div class="stat-label">${t('平均全距', 'Average Range')}</div>
                    <div class="stat-value">${SPCEngine.round(avgRange, 4)}</div>
                </div>
                <div class="stat-card">
                    <div class="stat-label">${t('最大全距', 'Max Range')}</div>
                    <div class="stat-value">${SPCEngine.round(maxRange, 4)}</div>
                </div>
            </div>
        `;

        // Statistics table
        let tableRows = '';
        groupStats.forEach(stat => {
            const inSpec = stat.min >= specs.lsl && stat.max <= specs.usl;
            const rowStyle = inSpec ? '' : 'background-color: rgba(239, 68, 68, 0.1);';

            tableRows += `
                <tr style="${rowStyle}">
                    <td><strong>${stat.batch}</strong></td>
                    <td>${SPCEngine.round(stat.avg, 4)}</td>
                    <td>${SPCEngine.round(stat.max, 4)}</td>
                    <td>${SPCEngine.round(stat.min, 4)}</td>
                    <td>${SPCEngine.round(stat.range, 4)}</td>
                    <td>${stat.count}</td>
                    <td>${specs.usl}</td>
                    <td>${specs.target}</td>
                    <td>${specs.lsl}</td>
                </tr>
            `;
        });

        const tableHTML = `
            <h3 class="chart-title">${t('群組統計數據', 'Group Statistics')}</h3>
            <div style="overflow-x: auto;">
                <table class="data-table">
                    <thead>
                        <tr>
                            <th>${t('批號', 'Batch')}</th>
                            <th>${t('平均值 (Avg)', 'Average')}</th>
                            <th>${t('最大值 (Max)', 'Maximum')}</th>
                            <th>${t('最小值 (Min)', 'Minimum')}</th>
                            <th>${t('全距 (Range)', 'Range')}</th>
                            <th>${t('樣本數 (n)', 'Sample Size')}</th>
                            <th>USL</th>
                            <th>Target</th>
                            <th>LSL</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${tableRows}
                    </tbody>
                </table>
            </div>
        `;

        // Chart placeholders
        const chartsHTML = `
            <div class="chart-container">
                <h3 class="chart-title">${t('Min-Max-Avg 管制圖', 'Min-Max-Avg Control Chart')}</h3>
                <canvas id="groupMinMaxAvgChart"></canvas>
            </div>
            <div class="chart-container">
                <h3 class="chart-title">${t('全距 (Range) 圖 - 模穴間變異', 'Range Chart - Inter-Cavity Variation')}</h3>
                <canvas id="groupRangeChart"></canvas>
            </div>
        `;

        return summaryHTML + tableHTML + chartsHTML;
    }

    generateCharts(groupStats, specs) {
        const t = (zh, en) => this.lang === 'zh' ? zh : en;
        const labels = groupStats.map(s => s.batch);

        return [
            // Min-Max-Avg chart
            {
                canvasId: 'groupMinMaxAvgChart',
                config: {
                    type: 'line',
                    data: {
                        labels: labels,
                        datasets: [
                            {
                                label: t('最大值 (Max)', 'Max'),
                                data: groupStats.map(s => s.max),
                                borderColor: '#ef4444',
                                backgroundColor: 'transparent',
                                pointRadius: 0,
                                borderWidth: 1.5,
                                fill: false
                            },
                            {
                                label: t('平均值 (Avg)', 'Avg'),
                                data: groupStats.map(s => s.avg),
                                borderColor: '#3b82f6',
                                backgroundColor: '#3b82f6',
                                pointRadius: 5,
                                pointHoverRadius: 7,
                                borderWidth: 2.5,
                                fill: false
                            },
                            {
                                label: t('最小值 (Min)', 'Min'),
                                data: groupStats.map(s => s.min),
                                borderColor: '#ef4444',
                                backgroundColor: 'transparent',
                                pointRadius: 0,
                                borderWidth: 1.5,
                                fill: false
                            },
                            {
                                label: 'USL',
                                data: Array(labels.length).fill(specs.usl),
                                borderColor: '#ff9800',
                                borderDash: [5, 5],
                                borderWidth: 2,
                                pointRadius: 0,
                                fill: false
                            },
                            {
                                label: 'Target',
                                data: Array(labels.length).fill(specs.target),
                                borderColor: '#10b981',
                                borderWidth: 1.5,
                                pointRadius: 0,
                                fill: false
                            },
                            {
                                label: 'LSL',
                                data: Array(labels.length).fill(specs.lsl),
                                borderColor: '#ff9800',
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
                        aspectRatio: 2.5,
                        plugins: {
                            legend: {
                                position: 'top',
                                labels: {
                                    font: { size: 12 }
                                }
                            }
                        },
                        scales: {
                            y: {
                                title: {
                                    display: true,
                                    text: t('測量值', 'Measurement Value'),
                                    font: { size: 14, weight: 'bold' }
                                }
                            },
                            x: {
                                title: {
                                    display: true,
                                    text: t('生產批號', 'Batch Number'),
                                    font: { size: 14, weight: 'bold' }
                                }
                            }
                        }
                    }
                }
            },
            // Range chart
            {
                canvasId: 'groupRangeChart',
                config: {
                    type: 'line',
                    data: {
                        labels: labels,
                        datasets: [
                            {
                                label: t('全距 (模穴間變異)', 'Range (Inter-Cavity Variation)'),
                                data: groupStats.map(s => s.range),
                                borderColor: '#8b5cf6',
                                backgroundColor: '#8b5cf6',
                                pointRadius: 5,
                                pointHoverRadius: 7,
                                borderWidth: 2,
                                fill: false
                            }
                        ]
                    },
                    options: {
                        responsive: true,
                        maintainAspectRatio: true,
                        aspectRatio: 2.5,
                        plugins: {
                            legend: {
                                position: 'top'
                            }
                        },
                        scales: {
                            y: {
                                beginAtZero: true,
                                title: {
                                    display: true,
                                    text: t('全距 (Range)', 'Range'),
                                    font: { size: 14, weight: 'bold' }
                                }
                            },
                            x: {
                                title: {
                                    display: true,
                                    text: t('生產批號', 'Batch Number'),
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
