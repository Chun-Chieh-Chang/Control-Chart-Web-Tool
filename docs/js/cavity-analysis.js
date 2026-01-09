// ============================================================================
// Cavity Analysis Module - Cavity Comparison & Capability Assessment
// ============================================================================

import { SPCEngine } from './spc-engine.js';

export class CavityAnalysis {
    constructor(dataInput, language = 'zh') {
        this.dataInput = dataInput;
        this.lang = language;
        this.results = null;
    }

    async execute() {
        const specs = this.dataInput.getSpecs();
        const cavityNames = this.dataInput.getCavityNames();
        const cavityCount = this.dataInput.getCavityCount();

        // Analyze each cavity
        const cavityStats = [];
        for (let i = 0; i < cavityCount; i++) {
            const data = this.dataInput.getCavityBatchData(i);
            const capability = SPCEngine.calculateProcessCapability(data, specs.usl, specs.lsl);

            cavityStats.push({
                name: cavityNames[i],
                ...capability
            });
        }

        // Generate HTML results
        const html = this.generateHTML(cavityStats, specs);

        // Generate chart configurations
        const charts = this.generateCharts(cavityStats, specs);

        this.results = {
            type: 'cavity',
            html: html,
            charts: charts,
            data: {
                cavityStats: cavityStats,
                specs: specs
            }
        };

        return this.results;
    }

    generateHTML(cavityStats, specs) {
        const t = (zh, en) => this.lang === 'zh' ? zh : en;

        // Summary
        const avgCpk = SPCEngine.mean(cavityStats.map(s => s.Cpk));
        const minCpk = SPCEngine.min(cavityStats.map(s => s.Cpk));
        const maxCpk = SPCEngine.max(cavityStats.map(s => s.Cpk));

        const summaryHTML = `
            <div class="results-summary">
                <div class="stat-card">
                    <div class="stat-label">${t('模穴數量', 'Cavity Count')}</div>
                    <div class="stat-value">${cavityStats.length}</div>
                </div>
                <div class="stat-card">
                    <div class="stat-label">${t('平均 Cpk', 'Average Cpk')}</div>
                    <div class="stat-value">${SPCEngine.round(avgCpk, 3)}</div>
                </div>
                <div class="stat-card">
                    <div class="stat-label">${t('最小 Cpk', 'Min Cpk')}</div>
                    <div class="stat-value" style="color: ${minCpk < 1.33 ? '#ef4444' : '#10b981'}">${SPCEngine.round(minCpk, 3)}</div>
                </div>
                <div class="stat-card">
                    <div class="stat-label">${t('最大 Cpk', 'Max Cpk')}</div>
                    <div class="stat-value">${SPCEngine.round(maxCpk, 3)}</div>
                </div>
            </div>
        `;

        // Statistics table
        let tableRows = '';
        cavityStats.forEach(stat => {
            const cpkColor = SPCEngine.getCapabilityColor(stat.Cpk);
            const ppkColor = SPCEngine.getCapabilityColor(stat.Ppk);

            tableRows += `
                <tr>
                    <td><strong>${stat.name}</strong></td>
                    <td>${SPCEngine.round(stat.mean, 4)}</td>
                    <td>${SPCEngine.round(stat.withinStdDev, 4)}</td>
                    <td>${SPCEngine.round(stat.overallStdDev, 4)}</td>
                    <td>${SPCEngine.round(stat.range, 4)}</td>
                    <td>${SPCEngine.round(stat.max, 4)}</td>
                    <td>${SPCEngine.round(stat.min, 4)}</td>
                    <td>${SPCEngine.round(stat.Cp, 3)}</td>
                    <td style="background-color: ${cpkColor.bg}; color: ${cpkColor.text}; font-weight: bold;">${SPCEngine.round(stat.Cpk, 3)}</td>
                    <td>${SPCEngine.round(stat.Pp, 3)}</td>
                    <td style="background-color: ${ppkColor.bg}; color: ${ppkColor.text}; font-weight: bold;">${SPCEngine.round(stat.Ppk, 3)}</td>
                    <td>${stat.count}</td>
                </tr>
            `;
        });

        const tableHTML = `
            <h3 class="chart-title">${t('模穴統計摘要', 'Cavity Statistics Summary')}</h3>
            <div style="overflow-x: auto;">
                <table class="data-table">
                    <thead>
                        <tr>
                            <th>${t('模穴', 'Cavity')}</th>
                            <th>${t('平均值', 'Mean')}</th>
                            <th>${t('組內標準差', 'Within StdDev')}</th>
                            <th>${t('整體標準差', 'Overall StdDev')}</th>
                            <th>${t('範圍', 'Range')}</th>
                            <th>${t('最大值', 'Max')}</th>
                            <th>${t('最小值', 'Min')}</th>
                            <th>Cp</th>
                            <th>Cpk</th>
                            <th>Pp</th>
                            <th>Ppk</th>
                            <th>${t('樣本數', 'Count')}</th>
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
                <h3 class="chart-title">${t('模穴平均值比較', 'Cavity Mean Comparison')}</h3>
                <canvas id="cavityMeanChart"></canvas>
            </div>
            <div class="chart-container">
                <h3 class="chart-title">${t('模穴 Cpk 比較', 'Cavity Cpk Comparison')}</h3>
                <canvas id="cavityCpkChart"></canvas>
            </div>
            <div class="chart-container">
                <h3 class="chart-title">${t('模穴標準差比較', 'Cavity Standard Deviation Comparison')}</h3>
                <canvas id="cavityStdDevChart"></canvas>
            </div>
        `;

        // Capability legend
        const legendHTML = `
            <div style="padding: 1.5rem; background: #f9fafb; border-radius: 0.75rem; margin: 2rem 0;">
                <h3 style="margin-bottom: 1rem;">${t('能力指標評級', 'Capability Index Rating')}</h3>
                <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 1rem;">
                    <div style="display: flex; align-items: center; gap: 0.5rem;">
                        <div style="width: 20px; height: 20px; background: #c6efce; border-radius: 4px;"></div>
                        <span>Cpk ≥ 1.67: ${t('優異', 'Excellent')}</span>
                    </div>
                    <div style="display: flex; align-items: center; gap: 0.5rem;">
                        <div style="width: 20px; height: 20px; background: #c6efce; border-radius: 4px;"></div>
                        <span>Cpk ≥ 1.33: ${t('良好', 'Good')}</span>
                    </div>
                    <div style="display: flex; align-items: center; gap: 0.5rem;">
                        <div style="width: 20px; height: 20px; background: #ffeb9c; border-radius: 4px;"></div>
                        <span>Cpk ≥ 1.0: ${t('可接受', 'Acceptable')}</span>
                    </div>
                    <div style="display: flex; align-items: center; gap: 0.5rem;">
                        <div style="width: 20px; height: 20px; background: #ffc7ce; border-radius: 4px;"></div>
                        <span>Cpk < 1.0: ${t('不足', 'Insufficient')}</span>
                    </div>
                </div>
            </div>
        `;

        return summaryHTML + tableHTML + legendHTML + chartsHTML;
    }

    generateCharts(cavityStats, specs) {
        const t = (zh, en) => this.lang === 'zh' ? zh : en;
        const labels = cavityStats.map(s => s.name);

        // Bar colors based on Cpk values
        const cpkColors = cavityStats.map(s => {
            const color = SPCEngine.getCapabilityColor(s.Cpk);
            return color.bg;
        });

        return [
            // Mean comparison chart
            {
                canvasId: 'cavityMeanChart',
                config: {
                    type: 'line',
                    data: {
                        labels: labels,
                        datasets: [
                            {
                                label: t('平均值', 'Mean'),
                                data: cavityStats.map(s => s.mean),
                                borderColor: '#3b82f6',
                                backgroundColor: '#3b82f6',
                                pointRadius: 6,
                                pointHoverRadius: 8,
                                borderWidth: 2
                            },
                            {
                                label: t('目標值', 'Target'),
                                data: Array(labels.length).fill(specs.target),
                                borderColor: '#10b981',
                                borderWidth: 2,
                                pointRadius: 0,
                                fill: false
                            },
                            {
                                label: 'USL',
                                data: Array(labels.length).fill(specs.usl),
                                borderColor: '#ef4444',
                                borderDash: [5, 5],
                                borderWidth: 2,
                                pointRadius: 0,
                                fill: false
                            },
                            {
                                label: 'LSL',
                                data: Array(labels.length).fill(specs.lsl),
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
                        aspectRatio: 2.5,
                        plugins: {
                            legend: {
                                position: 'top'
                            }
                        },
                        scales: {
                            y: {
                                title: {
                                    display: true,
                                    text: t('平均值', 'Mean Value')
                                }
                            },
                            x: {
                                title: {
                                    display: true,
                                    text: t('模穴', 'Cavity')
                                }
                            }
                        }
                    }
                }
            },
            // Cpk comparison chart
            {
                canvasId: 'cavityCpkChart',
                config: {
                    type: 'bar',
                    data: {
                        labels: labels,
                        datasets: [
                            {
                                label: 'Cpk',
                                data: cavityStats.map(s => s.Cpk),
                                backgroundColor: cpkColors,
                                borderColor: cpkColors.map(c => c.replace('0.1', '1')),
                                borderWidth: 2
                            },
                            {
                                label: t('優異 (1.67)', 'Excellent (1.67)'),
                                data: Array(labels.length).fill(1.67),
                                type: 'line',
                                borderColor: '#10b981',
                                borderDash: [5, 5],
                                borderWidth: 2,
                                pointRadius: 0,
                                fill: false
                            },
                            {
                                label: t('良好 (1.33)', 'Good (1.33)'),
                                data: Array(labels.length).fill(1.33),
                                type: 'line',
                                borderColor: '#f59e0b',
                                borderDash: [5, 5],
                                borderWidth: 2,
                                pointRadius: 0,
                                fill: false
                            },
                            {
                                label: t('可接受 (1.0)', 'Acceptable (1.0)'),
                                data: Array(labels.length).fill(1.0),
                                type: 'line',
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
                                    text: 'Cpk'
                                }
                            },
                            x: {
                                title: {
                                    display: true,
                                    text: t('模穴', 'Cavity')
                                }
                            }
                        }
                    }
                }
            },
            // Standard deviation comparison chart
            {
                canvasId: 'cavityStdDevChart',
                config: {
                    type: 'line',
                    data: {
                        labels: labels,
                        datasets: [
                            {
                                label: t('組內標準差', 'Within StdDev'),
                                data: cavityStats.map(s => s.withinStdDev),
                                borderColor: '#3b82f6',
                                backgroundColor: '#3b82f6',
                                pointRadius: 5,
                                borderWidth: 2
                            },
                            {
                                label: t('整體標準差', 'Overall StdDev'),
                                data: cavityStats.map(s => s.overallStdDev),
                                borderColor: '#ef4444',
                                backgroundColor: '#ef4444',
                                pointRadius: 5,
                                borderWidth: 2
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
                                    text: t('標準差', 'Standard Deviation')
                                }
                            },
                            x: {
                                title: {
                                    display: true,
                                    text: t('模穴', 'Cavity')
                                }
                            }
                        }
                    }
                }
            }
        ];
    }
}
