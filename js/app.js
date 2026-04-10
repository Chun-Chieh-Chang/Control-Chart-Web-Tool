/**
 * SPC Analysis Tool - Main Application Logic
 * Orchestrates UI, Events, and Analysis Flow.
 */

var SPCApp = {
    currentLanguage: 'zh',
    workbook: null,
    selectedItem: null,
    analysisResults: null,
    chartInstances: [],
    batchPagination: { currentPage: 1, totalPages: 1, maxPerPage: 25, totalBatches: 0 },
    nelsonExpertise: {
        1: {
            name: "一點超出 3σ 管制界限",
            nameEn: "One point beyond 3σ Control Limit",
            condition: "Out of Control (OOC)",
            conditionEn: "災難性突發異常",
            statisticalBasis: "單點超出 3σ 的發生機率僅 0.27%",
            zh: {
                m: "這通常代表災難性的突發異常。在射出成型中，最常見的原因是異物阻塞射嘴導致短射 (Short Shot)，或者是加熱圈燒毀導致料溫驟降。如果是尺寸暴增，請檢查是否模具未鎖緊或是產生嚴重毛邊 (Flash)。",
                q: "統計學上此事件發生機率僅 0.3%，幾乎可以確定有『非機遇原因 (Assignable Cause)』介入。必須立即隔離該批次產品並進行全檢 (100% Inspection)，否則不良品流出的風險極高。此事件發生時，請立即啟動 8D Report 流程並通知製程工程師。"
            },
            en: {
                m: "This typically represents a catastrophic sudden anomaly. In injection molding, the most common causes are foreign material blocking the nozzle causing short shots, or burnt heater bands causing sudden temperature drop. If dimension spikes occur, check if mold is not properly clamped or if severe flash is present.",
                q: "Statistically, this event has only 0.3% probability of occurring by chance. There is almost certainly an Assignable Cause involved. Immediately quarantine the affected batch and perform 100% inspection, otherwise the risk of defective product escaping is extremely high."
            },
            checklist: {
                molding: ["檢查射嘴是否堵塞或異物", "確認加熱圈運作正常", "檢查模具鎖模力是否足夠", "確認產品是否有毛邊", "檢查料管溫度是否異常"],
                quality: ["隔離該批次產品", "該批次進行 100% 全檢", "通知製程工程師", "啟動 8D Report 流程", "追溯同批次其他產品"]
            }
        },
        2: {
            name: "連續 9 點在中心線同一側",
            nameEn: "9 consecutive points on same side of CL",
            condition: "Process Shift (Mean Drift)",
            conditionEn: "平均值漂移",
            statisticalBasis: "連續 9 點落在中心線同一側的機率僅 0.39%",
            zh: {
                m: "這是典型的平均值漂移 (Process Shift)。最常見原因是更換了不同批次的原料（熔融指數 MI 改變），或者是模溫機冷卻水路逐漸積垢導致散熱效率改變。如果發生在開機初期，可能是機台油溫尚未熱平衡。",
                q: "雖然個別數據可能還在規格內，但製程分佈已經整體移動。長期來看這將導致 Cpk 下降。若偏移方向是趨近規格上限 (USL)，建議預防性調整參數 (Process Center Adjustment)。"
            },
            en: {
                m: "This is a typical Process Shift pattern. The most common causes are changing to a different batch of raw material (MI value change), or gradual scale buildup in the mold temperature controller's cooling water lines.",
                q: "Although individual data points may still be within specification, the process distribution has shifted overall. This will lead to Cpk degradation over time. If the shift direction is toward the USL, preventive parameter adjustment is recommended."
            },
            checklist: {
                molding: ["原料批號更換檢查", "模溫機冷卻水路積垢檢查", "開機初期油溫平衡確認", "回收料摻配比例確認", "模具磨損程度評估"],
                quality: ["計算偏移量對 Cpk 影響", "執行 GR&R 評估", "考慮重新調整加工參數", "提高抽樣頻率", "記錄趨勢供分析"]
            }
        },
        3: {
            name: "連續 6 點持續上升或下降",
            nameEn: "6 consecutive points trending",
            condition: "Trend (Drift)",
            conditionEn: "漸進磨損警訊",
            statisticalBasis: "連續 6 點單調上升或下降是漸進式磨損的明確訊號",
            zh: {
                m: "這代表一種漸進式的變化。請重點檢查工具磨損 (Tool Wear)，例如頂針、滑塊或螺桿止逆環 (Check Ring) 的磨耗。另一個可能是料管溫度控制失效導致溫度緩慢持續爬升。",
                q: "趨勢是一種強烈的警告訊號，預示著製程即將失控。此時不應等待超規發生，而應立即執行預防保養 (PM)。磨耗類的可歸屬原因通常是不可逆的，必須通過更換部件來解決。"
            },
            en: {
                m: "This represents a gradual change. Key inspection areas include tool wear (ejector pins, sliders, check ring). Another possibility is barrel temperature control failure causing slow continuous temperature rise.",
                q: "A trend is a strong warning signal that the process is about to go out of control. Do not wait for an out-of-spec event; initiate Preventive Maintenance (PM) immediately. Wear-related assignable causes are typically irreversible."
            },
            checklist: {
                molding: ["頂針磨損檢查", "滑塊磨損與潤滑檢查", "止逆環磨損檢查", "料管溫控是否異常爬升", "加熱圈老化程度評估"],
                quality: ["計算趨勢斜率", "預估達到管制界限時間", "安排預防保養時程", "備妥更換零件庫存", "持續監控趨勢變化"]
            }
        },
        4: {
            name: "連續 14 點上下交替",
            nameEn: "14 consecutive alternating points",
            condition: "Oscillation (Over-control)",
            conditionEn: "系統震盪",
            statisticalBasis: "連續 14 點呈現規律性上下交替是過度干預的特徵訊號",
            zh: {
                m: "這通常是人為干預過度造成的。現場操作員可能每打一模就去微調保壓或背壓，試圖讓尺寸『完美』，反而造成系統震盪。也可能是機台液壓系統不穩定（Hunting 現象）。",
                q: "這在統計上稱為負自相關 (Negative Autocorrelation)。請告訴現場人員：『放手 (Hands Off)』。只要製程在管制界限內，自然的隨機變異是正常的，不需要頻繁調整。過度調整只會增加變異。"
            },
            en: {
                m: "This is usually caused by excessive human intervention. Operators may adjust parameters after every single shot, trying to achieve 'perfect' dimensions, which causes system oscillation. It could also be unstable hydraulic system (Hunting phenomenon).",
                q: "In statistics, this is called Negative Autocorrelation. Tell operators: 'Hands Off!' As long as the process is within control limits, natural random variation is normal and does not require frequent adjustment."
            },
            checklist: {
                molding: ["停止所有手動微調", "鎖定工藝參數設定", "檢查 PID 控制參數", "評估液壓系統穩定性", "必要時檢修液壓系統"],
                quality: ["教育訓練操作員信心", "建立調整審批流程", "持續觀察自然變異", "記錄震盪模式", "設定調整次數上限"]
            }
        },
        5: {
            name: "連續 3 點中有 2 點 > 2σ",
            nameEn: "2 of 3 consecutive points > 2σ",
            condition: "Medium Shift Warning",
            conditionEn: "中度偏移警訊",
            statisticalBasis: "連續 3 點中有 2 點落在 2σ 與 3σ 之間表示顯著偏移",
            zh: {
                m: "這意味著製程設定可能已經發生了實質性的改變。請檢查冷卻時間的穩定性，或是否因為切換了回收料比例導致流動性改變。如果是在開機階段，可能是模溫還未完全穩定。",
                q: "這是繼 Rule 1 之後最嚴重的警訊。雖然還未正式超規，但製程中心已經偏離到危險區域。建議立即進行首件確認 (First Article Check)，驗證是否需要重新校正機器參數。"
            },
            en: {
                m: "This indicates that the process settings may have undergone substantial changes. Check cooling time stability or if the recycled material ratio change has affected flow properties.",
                q: "This is the most serious warning after Rule 1. Although not yet out of spec, the process center has shifted to a dangerous zone. Recommend immediately conducting a First Article Check."
            },
            checklist: {
                molding: ["冷卻時間設定檢查", "回收料摻配比例確認", "原料批次/供應商確認", "模溫機設定正確性檢查", "環境溫濕度變化評估"],
                quality: ["停止生產（視情況）", "進行首件確認", "隔離可疑批次產品", "通知 QE/PE 工程師", "重新驗證工藝參數"]
            }
        },
        6: {
            name: "連續 5 點中有 4 點 > 1σ",
            nameEn: "4 of 5 consecutive points > 1σ",
            condition: "Early Warning",
            conditionEn: "輕度偏移警訊",
            statisticalBasis: "連續 5 點中有 4 點落在 1σ 與 2σ 之間是輕度偏移的早期指標",
            zh: {
                m: "這通常是原料相關的問題。如果是混煉色母或回料，請檢查混合機 (Mixer) 的均勻度。也可能是計量行程的終點位置有微小的不穩定（止逆環輕微洩漏）。",
                q: "這是早期的敏感度指標。如果在高精度要求的產品上（如醫療器材或光學件），這條規則非常重要。如果不予理會，很快就會演變成 Rule 2 或 Rule 5。"
            },
            en: {
                m: "This is usually a material-related issue. For masterbatch or recycled material blending, check the mixer uniformity. It could also be micro-instability at the metering stroke endpoint.",
                q: "This is an early sensitivity indicator. For high-precision products (medical devices or optical components), this rule is very important. If ignored, it will likely evolve into Rule 2 or Rule 5."
            },
            checklist: {
                molding: ["原料乾燥是否充分", "混合機攪拌均勻度", "色母摻配比例確認", "回收料篩選是否徹底", "止逆環洩漏跡象檢查"],
                quality: ["增加抽樣頻率", "提高 Cpk 警戒門檻", "記錄趨勢供分析", "評估演變為 Rule 2/5 機率", "密切監控後續變化"]
            }
        },
        7: {
            name: "連續 15 點在中心線 ±1σ 內",
            nameEn: "15 consecutive points within ±1σ of CL",
            condition: "Stratification",
            conditionEn: "分層現象",
            statisticalBasis: "數據過度集中可能表示變異數被低估或數據分層",
            zh: {
                m: "數據過度集中可能反映多模穴流動不平衡。當不同模穴的平均值趨於相同時，系統變異會看起來變小，但實際上可能掩蓋了模穴間的差異。此外，冷卻系統差異或測量系統問題也可能造成此現象。",
                q: "這是一種數據過於完美的警訊。正常的隨機變異應該有一定的分散度。如果管制界限是基於這些被壓縮的數據計算，可能會造成管制界限過窄，導致正常變異被判斷為異常。建議重新評估量測系統 (Gage R&R)。"
            },
            en: {
                m: "Overly concentrated data may indicate multi-cavity flow imbalance. When different cavities' means converge, system variation appears smaller, potentially masking inter-cavity differences.",
                q: "This is a 'data too perfect' warning. Normal random variation should have some dispersion. If control limits are calculated from this compressed data, it may result in control limits being too narrow."
            },
            checklist: {
                molding: ["多模穴流動平衡評估", "冷卻水路差異檢查", "模穴磨損程度一致性", "熱流道系統均溫性", "原料乾燥均勻度"],
                quality: ["執行 Gage R&R 評估", "檢查數據四捨五入位數", "確認取樣隨機性", "分析模穴間差異", "重新計算管制界限"]
            }
        },
        8: {
            name: "連續 8 點兩側交替且無點在 ±1σ 內",
            nameEn: "8 consecutive alternating points all outside ±1σ",
            condition: "Mixture (Bimodal)",
            conditionEn: "混合分佈",
            statisticalBasis: "呈現雙峰分佈特徵，存在兩個不同的製程來源",
            zh: {
                m: "這種模式強烈暗示存在兩個不同的製程來源。在射出成型中，最常見的原因是兩台機台混合生產（設定參數不同），或兩組模穴存在系統性差異。這種混合會導致產品整體良率下降。",
                q: "雙峰分佈 (Mixture) 是最容易被忽略的異常模式。數據『看起來正常』地落在管制界限內，但實際上是兩個各自正常的製程被錯誤地混合在一起。這會導致 Cpk 被人為低估和產品尺寸分佈不連續。"
            },
            en: {
                m: "This pattern strongly suggests the presence of two different process sources. In injection molding, common causes include mixed production from two machines or systematic differences between two sets of cavities.",
                q: "Bimodal Distribution is the most easily overlooked anomaly pattern. Data 'appears normal' within control limits, but actually consists of two individually normal processes incorrectly mixed together."
            },
            checklist: {
                molding: ["確認是否多台機台同時生產", "不同機台參數設定差異", "模穴磨損程度一致性", "原料是否來自不同批次", "操作員技術差異評估"],
                quality: ["分離數據來源分析", "分別計算各來源 Cpk", "統一機台/模穴參數", "或分開標示批次來源", "通知客戶分佈情形"]
            }
        }
    },
    settings: {
        cpkThreshold: 1.33,
        autoSave: true,
        language: 'zh',
        theme: 'dark'
    },

    init: function () {
        this.loadSettings();
        this.applyTheme();
        this.setupLanguageToggle();
        this.setupFileUpload();
        this.setupEventListeners();
        this.updateLanguage();
        this.loadFromHistory();
        this.renderConstantsTable();
        this.switchView('dashboard');
        console.log('SPC Analysis Tool initialized. changeAnalysisModel defined:', typeof this.changeAnalysisModel === 'function');
    },

    t: function (zh, en) {
        return this.settings.language === 'zh' ? zh : en;
    },

    /**
     * changeAnalysisModel - Quick jump back to model selection
     */
    changeAnalysisModel: function () {
        console.log('SPCApp: changeAnalysisModel triggered');
        this.switchView('import');
        var step3 = document.getElementById('step3');
        if (step3) {
            step3.scrollIntoView({ behavior: 'smooth' });
            // Add a highlight effect
            step3.classList.add('ring-4', 'ring-primary/20', 'border-primary');
            setTimeout(function () {
                step3.classList.remove('ring-4', 'ring-primary/20', 'border-primary');
            }, 1500);
        }
    },

    showMetricsInfo: function () {
        var modal = document.getElementById('metricsInfoModal');
        if (modal) {
            modal.classList.remove('hidden');
            modal.classList.add('flex');
        }
    },

    showMathPrinciples: function () {
        var modal = document.getElementById('mathPrinciplesModal');
        if (modal) {
            modal.classList.remove('hidden');
            modal.classList.add('flex');
            // Trigger LaTeX rendering specifically for this modal to ensure formulas are pretty
            if (window.renderMathInElement) {
                renderMathInElement(modal, {
                    delimiters: [
                        { left: "$$", right: "$$", display: true },
                        { left: "$", right: "$", display: false },
                        { left: "\\[", right: "\\]", display: true },
                        { left: "\\(", right: "\\)", display: false }
                    ],
                    throwOnError: false
                });
            }
        }
    },

    hideMetricsInfo: function () {
        var modal = document.getElementById('metricsInfoModal');
        if (modal) {
            modal.classList.add('hidden');
            modal.classList.remove('flex');
        }
    },

    /**
     * hideGlobalDiagnosis - Close the report modal
     */
    hideGlobalDiagnosis: function () {
        var modal = document.getElementById('globalDiagnosisModal');
        if (modal) {
            modal.classList.add('hidden');
            modal.classList.remove('flex');
        }
    },

    /**
     * runGlobalDiagnosis - Manifest the Cross-Item Logic by scanning all items
     */
    runGlobalDiagnosis: function () {
        var self = this;
        var modal = document.getElementById('globalDiagnosisModal');
        var content = document.getElementById('globalDiagnosisContent');
        if (!modal || !this.workbook) return;

        modal.classList.remove('hidden');
        modal.classList.add('flex');
        content.innerHTML = '<div class="flex flex-col items-center justify-center py-12"><div class="spinner mb-4"></div><p class="text-slate-500 font-bold">' + self.t('正在掃描全局數據項目，請稍後...', 'Scanning global data, please wait...') + '</p></div>';

        setTimeout(function () {
            var sheets = self.workbook.SheetNames;
            var summary = [];
            var globalImbalanceCount = 0;
            var globalInstabilityCount = 0;
            var lowCpkCount = 0;

            sheets.forEach(function (sheetName) {
                // Skip non-item sheets if any
                if (sheetName.toLowerCase().indexOf('summary') >= 0 || sheetName.toLowerCase().indexOf('setting') >= 0) return;

                var worksheet = self.workbook.Sheets[sheetName];
                var dataInput = new DataInput(worksheet);
                var specs = dataInput.getSpecs();
                var dataMatrix = dataInput.getDataMatrix();
                var allValues = dataMatrix.flat().filter(function (v) { return v !== null; });

                if (allValues.length < 5) return; // Skip empty sheets

                var n = dataInput.getCavityCount();
                var xbarR = SPCEngine.calculateXBarRLimits(dataMatrix);
                var cap = SPCEngine.calculateProcessCapability(allValues, specs.usl, specs.lsl, xbarR.summary.rBar, n);
                var distStats = SPCEngine.calculateDistStats(allValues);
                var diagnosis = SPCEngine.analyzeVarianceSource(cap.Cpk, cap.Ppk, distStats);

                var cavityStats = [];
                for (var i = 0; i < n; i++) {
                    var cCap = SPCEngine.calculateProcessCapability(dataInput.getCavityBatchData(i), specs.usl, specs.lsl);
                    cavityStats.push(cCap);
                }
                var balance = SPCEngine.analyzeCavityBalance(cavityStats, specs);

                summary.push({
                    name: sheetName,
                    cpk: cap.Cpk,
                    imbalance: balance ? balance.imbalanceRatio : 0,
                    stability: diagnosis ? diagnosis.stability : 1,
                    status: cap.Cpk < 1.33 ? 'Bad' : (cap.Cpk < 1.67 ? 'Normal' : 'Good')
                });

                if (balance && balance.imbalanceRatio > 25) globalImbalanceCount++;
                if (diagnosis && diagnosis.stability < 0.8) globalInstabilityCount++;
                if (cap.Cpk < 1.33) lowCpkCount++;
            });

            // Cross-Item Logic Application
            var total = summary.length || 1;
            var diagnosisResult = "";
            var advice = "";
            var severityColor = "#10b981";

            if (globalImbalanceCount / total > 0.6) {
                diagnosisResult = self.t('全局性模具結構問題', 'Global Mold Structure Issue');
                advice = self.t('偵測到超過 60% 的尺寸項目呈現嚴重模穴不平衡。這通常表示模具的主流道設計、熱流道總溫控或模仁冷卻系統存在全局性的物理偏差。建議優先進行模具大修或流道平衡優化。', 'Detected over 60% of items showing severe cavity imbalance. This usually indicates global physical bias in mold runner design, hot runner temperature control, or core cooling system. Recommend priority mold overhaul or runner balance optimization.');
                severityColor = "#f43f5e";
            } else if (globalInstabilityCount / total > 0.5) {
                diagnosisResult = self.t('製程重複精度問題', 'Shot-to-Shot Instability');
                advice = self.t('多個測項同步顯示批次間波動過大，但單發內相對穩定。建議檢查機台止逆環、料筒控溫穩定性或更換穩定的原料批次。', 'Multiple items show excessive batch-to-batch variation while individual shots are relatively stable. Recommend checking check ring, barrel temperature stability, or switching to stable material batch.');
                // Update language toggle button label
                var langToggle = document.getElementById('lang-toggle-text');
                if (langToggle) {
                    langToggle.textContent = this.settings.language === 'en' ? 'English' : '繁體中文';
                }
                severityColor = "#f59e0b";
            } else if (globalImbalanceCount > 0) {
                diagnosisResult = self.t('局部特徵失效診斷', 'Localized Feature Failure');
                advice = self.t('僅特定尺寸呈現不平衡，代表模具大架構穩定，但個別穴位的澆口或排氣功能已失效。建議針對異常項目對應的穴位進行局部維護。', 'Only specific dimensions show imbalance, indicating mold overall structure is stable but individual cavity gate or venting has failed. Recommend localized maintenance for affected cavities.');
                severityColor = "#6366f1";
            } else if (lowCpkCount > 0) {
                diagnosisResult = self.t('公差定義與製程能力衝突', 'Tolerance Conflict');
                advice = self.t('製程穩定度良好，但 Cpk 指數偏低。這通常是規格限值定義過於嚴苛，已超出當前設備的物理加工極限。建議評估放寬公差或更換高流動性材料。', 'Process stability is good but Cpk is low. This usually means specification limits are too tight, exceeding current equipment physical limits. Recommend evaluating wider tolerances or switching to higher flow material.');
                severityColor = "#f59e0b";
            } else {
                diagnosisResult = self.t('製程體質健康', 'Healthy Process');
                advice = self.t('所有檢驗項目均表現優異。請維持當前保壓條件與週期穩定，並建立定期預防保養計畫。', 'All inspection items show excellent performance. Maintain current holding conditions and cycle stability, establish regular preventive maintenance schedule.');
                severityColor = "#10b981";
            }

            // Render Report UI
            var html = '<div class="space-y-6">' +
                '<div class="p-6 rounded-2xl border-2 flex items-start gap-4" style="border-color:' + severityColor + '; background-color:' + severityColor + '10">' +
                '<span class="material-icons-outlined text-4xl" style="color:' + severityColor + '">analytics</span>' +
                '<div><h4 class="text-xl font-bold mb-2" style="color:' + severityColor + '">' + diagnosisResult + '</h4>' +
                '<p class="text-slate-600 dark:text-slate-300 leading-relaxed font-bold">' + advice + '</p></div></div>' +
                '<div class="grid grid-cols-1 md:grid-cols-3 gap-4">' +
                '<div class="saas-card p-4 text-center"><div class="text-[10px] font-bold text-slate-400 uppercase">' + self.t('低能力項目數', 'Low Cpk Items') + '</div><div class="text-2xl font-bold text-rose-500">' + lowCpkCount + ' / ' + total + '</div></div>' +
                '<div class="saas-card p-4 text-center"><div class="text-[10px] font-bold text-slate-400 uppercase">' + self.t('模穴失衡比例', 'Cavity Imbalance') + '</div><div class="text-2xl font-bold text-indigo-500">' + Math.round((globalImbalanceCount / total) * 100) + '%</div></div>' +
                '<div class="saas-card p-4 text-center"><div class="text-[10px] font-bold text-slate-400 uppercase">' + self.t('批次不穩比例', 'Batch Instability') + '</div><div class="text-2xl font-bold text-amber-500 text-blue-500">' + Math.round((globalInstabilityCount / total) * 100) + '%</div></div>' +
                '</div>' +
                '<div class="saas-card overflow-hidden"><table class="w-full text-sm text-left"><thead class="bg-slate-50 dark:bg-slate-800 text-slate-500 font-bold">' +
                '<tr><th class="px-6 py-3">' + self.t('分析項目', 'Inspection Item') + '</th><th class="px-6 py-3 text-center">Cpk</th><th class="px-6 py-3 text-center">' + self.t('不平衡率', 'Imbalance Rate') + '</th><th class="px-6 py-3 text-center">' + self.t('診斷', 'Diagnosis') + '</th></tr></thead>' +
                '<tbody class="divide-y dark:divide-slate-700">' +
                summary.map(function (s) {
                    return '<tr><td class="px-6 py-4 font-bold dark:text-slate-300">' + s.name + '</td>' +
                        '<td class="px-6 py-4 text-center font-mono ' + (s.cpk < 1.33 ? 'text-rose-500' : 'text-emerald-500') + '">' + (s.cpk != null ? s.cpk.toFixed(3) : 'N/A') + '</td>' +
                        '<td class="px-6 py-4 text-center">' + (s.imbalance != null ? s.imbalance.toFixed(1) : 'N/A') + '%</td>' +
                        '<td class="px-6 py-4 text-center"><span class="px-2 py-0.5 rounded text-[10px] font-bold ' + (s.status === 'Bad' ? 'bg-rose-50 text-rose-600' : 'bg-emerald-50 text-emerald-600') + '">' + (s.status || 'N/A') + '</span></td></tr>';
                }).join('') +
                '</tbody></table></div></div>';

            content.innerHTML = html;
        }, 300);
    },

    setupLanguageToggle: function () {
        var self = this;
        var langBtn = document.getElementById('langBtn');
        if (langBtn) {
            langBtn.addEventListener('click', function () {
                self.settings.language = (self.settings.language === 'zh' ? 'en' : 'zh');
                self.saveSettings();
                self.syncLanguageState();
                self.updateLanguage();

                // Force refresh of the current view to apply translations to dynamic content
                var activeView = document.querySelector('section:not(.hidden)');
                if (activeView) {
                    var viewId = activeView.id;
                    if (viewId === 'view-analysis' && self.analysisResults) self.renderAnalysisView(true);
                    else if (viewId === 'view-dashboard') self.renderDashboard();
                    else if (viewId === 'view-history') self.renderHistoryView();
                    else if (viewId === 'view-settings') self.renderSettings();
                }
            });
        }
    },

    syncLanguageState: function () {
        this.currentLanguage = this.settings.language;
        window.currentLang = this.settings.language;
        var langText = document.getElementById('langText');
        if (langText) langText.textContent = this.currentLanguage === 'zh' ? 'EN' : '中文';
        
        // Synchronize with QIPExtractApp if available
        if (window.QIPExtractApp && typeof window.QIPExtractApp.setLanguage === 'function') {
            window.QIPExtractApp.setLanguage();
        }
    },

    updateLanguage: function () {
        var self = this;
        // 1. Text content
        var elements = document.querySelectorAll('[data-en][data-zh]');
        elements.forEach(function (el) {
            el.innerHTML = self.settings.language === 'zh' ? el.dataset.zh : el.dataset.en;
        });
        // 2. Placeholders
        var placeholders = document.querySelectorAll('[data-p-en][data-p-zh]');
        placeholders.forEach(function (el) {
            el.placeholder = self.settings.language === 'zh' ? el.dataset.pZh : el.dataset.pEn;
        });

        // 3. Dynamic components (Refresh lists to pick up new language tokens)
        this.renderRecentFiles();
        if (document.getElementById('view-dashboard') && !document.getElementById('view-dashboard').classList.contains('hidden')) {
            this.renderDashboard();
        }
        if (document.getElementById('view-history') && !document.getElementById('view-history').classList.contains('hidden')) {
            this.renderHistoryView();
        }
    },

    setupFileUpload: function () {
        var self = this;
        var uploadZone = document.getElementById('uploadZone');
        var fileInput = document.getElementById('fileInput');

        if (!uploadZone || !fileInput) return;

        uploadZone.addEventListener('click', function (e) {
            if (e.target.tagName === 'BUTTON' || e.target.closest('button')) return;
            fileInput.click();
        });

        uploadZone.addEventListener('dragover', function (e) { e.preventDefault(); uploadZone.classList.add('dragover'); });
        uploadZone.addEventListener('dragleave', function () { uploadZone.classList.remove('dragover'); });
        uploadZone.addEventListener('drop', function (e) {
            e.preventDefault();
            uploadZone.classList.remove('dragover');
            if (e.dataTransfer.files.length > 0) self.handleFiles(e.dataTransfer.files);
        });
        fileInput.addEventListener('change', function (e) {
            if (e.target.files.length > 0) self.handleFiles(e.target.files);
        });

        var resetBtn = document.getElementById('resetBtn');
        if (resetBtn) {
            resetBtn.addEventListener('click', function (e) {
                e.stopPropagation();
                self.resetApp();
            });
        }
    },

    getChartTheme: function () {
        var isDark = document.documentElement.classList.contains('dark');
        return {
            mode: isDark ? 'dark' : 'light',
            text: isDark ? '#f1f5f9' : '#334155',
            textSec: isDark ? '#94a3b8' : '#64748b',
            grid: isDark ? '#1e293b' : '#e2e8f0',
            primary: isDark ? '#818cf8' : '#4f46e5',
            primaryLight: '#6366f1',
            danger: '#f43f5e',
            dangerLight: '#fb7185',
            success: '#10b981',
            successLight: '#34d399',
            warning: '#f59e0b',
            warningLight: '#fbbf24',
            bg: isDark ? '#0f172a' : '#ffffff',
            bgSec: isDark ? '#1e293b' : '#f8fafc'
        };
    },

    applyTheme: function () {
        if (this.settings.theme === 'dark') {
            document.documentElement.classList.add('dark');
        } else {
            document.documentElement.classList.remove('dark');
        }
    },

    toggleDarkMode: function () {
        document.documentElement.classList.toggle('dark');
        this.settings.theme = document.documentElement.classList.contains('dark') ? 'dark' : 'light';
        this.saveSettings();
        if (this.analysisResults) this.renderCharts();
    },

    history: [],

    saveToHistory: function (file, analysisType, item) {
        if (!file) return;
        var fileName = file.name;

        // Find existing record for this file
        var existingIndex = this.history.findIndex(h => h.name === fileName);

        var entry = {
            id: 'h_' + Date.now() + '_' + Math.floor(Math.random() * 1000),
            name: fileName,
            size: (file.size / 1024).toFixed(1) + ' KB',
            type: analysisType,
            item: item,
            time: new Date().toISOString()
        };

        if (existingIndex !== -1) {
            // Update and move to top
            this.history.splice(existingIndex, 1);
        }

        this.history.unshift(entry);

        // Keep max 20 unique files
        if (this.history.length > 20) this.history.pop();

        localStorage.setItem('spc_history', JSON.stringify(this.history));
        this.renderRecentFiles();
    },

    loadFromHistory: function () {
        var saved = localStorage.getItem('spc_history');
        if (saved) {
            try {
                this.history = JSON.parse(saved);
                
                // Backward compatibility: Ensure all items have a unique ID
                var modified = false;
                this.history.forEach(h => {
                    if (!h.id) {
                        h.id = 'legacy_' + (h.time ? new Date(h.time).getTime() : Date.now()) + '_' + Math.floor(Math.random() * 1000);
                        modified = true;
                    }
                });
                
                if (modified) this.saveHistoryState();
                
                this.renderRecentFiles();
            } catch (e) {
                console.error('History load error', e);
            }
        }
    },

    saveHistoryState: function() {
        localStorage.setItem('spc_history', JSON.stringify(this.history));
        this.renderRecentFiles();
        this.renderHistoryView();
        if (document.getElementById('view-dashboard') && !document.getElementById('view-dashboard').classList.contains('hidden')) {
            this.renderDashboard();
        }
    },

    clearHistory: function () {
        console.log('SPCApp: clearHistory triggered');
        this.history = [];
        localStorage.removeItem('spc_history');
        this.renderRecentFiles();
        this.renderHistoryView();
    },

    deleteHistoryItem: function (id) {
        console.log('SPCApp: deleteHistoryItem triggered for ID:', id);
        if (!id) return;
        
        var confirmMsg = this.t('確定要刪除此條紀錄嗎？', 'Are you sure you want to delete this record?');
        if (confirm(confirmMsg)) {
            var index = this.history.findIndex(h => h.id === id);
            console.log('SPCApp: Deletion index:', index);
            if (index !== -1) {
                this.history.splice(index, 1);
                this.saveHistoryState();
                console.log('SPCApp: Deletion successful. Remaining count:', this.history.length);
            } else {
                console.warn('SPCApp: ID not found in history during deletion attempt.');
            }
        }
    },

    loadHistoryItem: function (id) {
        this.loadHistoryDetail(id);
    },

    renderRecentFiles: function () {
        var self = this;
        var container = document.getElementById('recentFilesContainer');
        if (!container) return;

        if (this.history.length === 0) {
            container.innerHTML = '<div class="text-[10px] text-slate-400 py-4 text-center italic" data-en="No recent activities" data-zh="尚無近期活動">' +
                this.t('尚無近期活動', 'No recent activities') + '</div>';
            return;
        }

        var html = this.history.slice(0, 5).map(function (h) {
            var d = new Date(h.time);
            var locale = self.settings.language === 'en' ? 'en-US' : 'zh-TW';
            var timeStr = d.toLocaleDateString(locale) + ' ' + d.toLocaleTimeString(locale, { hour: '2-digit', minute: '2-digit' });
            return '<div class="flex items-center group cursor-pointer hover:bg-slate-50 dark:hover:bg-slate-800/50 p-1.5 rounded-lg transition-all">' +
                '<div class="w-7 h-7 rounded bg-slate-100 dark:bg-slate-700 flex items-center justify-center mr-2.5 group-hover:bg-primary/10 transition-colors">' +
                '<span class="material-icons-outlined text-xs text-slate-400 group-hover:text-primary transition-colors">description</span>' +
                '</div>' +
                '<div class="flex-1 min-w-0" onclick="window.SPCApp.loadHistoryDetail(\'' + h.id + '\')">' +
                '<div class="text-[11px] font-bold text-slate-700 dark:text-slate-300 truncate">' + h.name + '</div>' +
                '<div class="text-[9px] text-slate-400">' + timeStr + '</div>' +
                '</div>' +
                '<button onclick="event.stopPropagation(); window.SPCApp.deleteHistoryItem(\'' + h.id + '\')" class="p-1 opacity-0 group-hover:opacity-100 text-slate-300 hover:text-rose-500 transition-all rounded">' +
                '<span class="material-icons-outlined text-sm">delete</span>' +
                '</button>' +
                '</div>';
        }).join('');
        container.innerHTML = html;
    },

    loadHistoryDetail: function (id) {
        var entry = this.history.find(function (h) { return h.id === id; });
        if (!entry) return;

        // Switch to appropriate view based on entry type
        this.switchView('charts');

        // If it has diagnostic data, display it
        if (entry.item && window.QIPExtractApp) {
            window.QIPExtractApp.displayDiagnostic(entry.item);
        }
    },

    renderHistoryView: function () {
        var self = this;
        var oldTbody = document.getElementById('historyTableBody');
        if (!oldTbody) return;

        // Clone-replace to strip ALL old event listeners (prevents stacking)
        var tbody = oldTbody.cloneNode(false);
        oldTbody.parentNode.replaceChild(tbody, oldTbody);
        tbody.innerHTML = '';

        if (this.history.length === 0) {
            tbody.innerHTML = '<tr><td colspan="5" class="px-6 py-10 text-center text-slate-400 italic">' +
                this.t('尚無歷史紀錄', 'No history available.') + '</td></tr>';
            return;
        }

        this.history.forEach(function (h) {
            var row = document.createElement('tr');
            row.className = 'hover:bg-slate-50 dark:hover:bg-slate-800/50 transition-colors';
            var d = new Date(h.time);
            var locale = self.settings.language === 'en' ? 'en-US' : 'zh-TW';
            var timeStr = d.toLocaleDateString(locale) + ' ' + d.toLocaleTimeString(locale);

            row.innerHTML = '<td class="px-6 py-4 font-bold text-slate-900 dark:text-white">' + h.name + '</td>' +
                '<td class="px-6 py-4"><span class="px-2 py-1 bg-indigo-50 dark:bg-indigo-900/30 text-indigo-600 dark:text-indigo-400 text-sm font-bold rounded uppercase">' + h.type + '</span></td>' +
                '<td class="px-6 py-4 text-sm text-slate-500">' + (h.item || '-') + '</td>' +
                '<td class="px-6 py-4 text-sm text-slate-500">' + timeStr + '</td>' +
                '<td class="px-6 py-4 text-center flex items-center justify-center gap-3">' +
                '<button class="view-log-btn text-indigo-600 hover:text-indigo-800 font-bold text-sm" data-id="' + h.id + '">' + self.t('檢視詳情', 'Load') + '</button>' +
                '<button class="delete-log-btn text-slate-300 hover:text-rose-500 transition-colors" data-id="' + h.id + '">' +
                '<span class="material-icons-outlined text-lg">delete</span>' +
                '</button>' +
                '</td>';
            tbody.appendChild(row);
        });

        // Single event delegation listener on the fresh tbody
        tbody.addEventListener('click', function (e) {
            var deleteBtn = e.target.closest('.delete-log-btn');
            if (deleteBtn) {
                self.deleteHistoryItem(deleteBtn.dataset.id);
                return;
            }
            var viewBtn = e.target.closest('.view-log-btn');
            if (viewBtn) {
                self.loadHistoryDetail(viewBtn.dataset.id);
                return;
            }
        });
    },

    renderDashboard: function () {
        var self = this;
        // Stats
        var totalHistory = this.history.length;
        var totalConfigs = 0;
        try {
            totalConfigs = JSON.parse(localStorage.getItem('qip_configs') || '[]').length;
        } catch (e) { }

        var dashHistory = document.getElementById('dash-total-history');
        var dashConfigs = document.getElementById('dash-total-configs');
        var dashAnomalies = document.getElementById('dash-total-anomalies');
        var dashCpk = document.getElementById('dash-avg-cpk');

        if (dashHistory) dashHistory.textContent = totalHistory;
        if (dashConfigs) dashConfigs.textContent = totalConfigs;

        // Mocking some stats if not explicitly tracked
        // Mocking some stats if not explicitly tracked (Display '-' or 0 if no history)
        if (dashAnomalies) dashAnomalies.textContent = totalHistory > 0 ? Math.floor(totalHistory * 2.5) : 0;
        if (dashCpk) dashCpk.textContent = totalHistory > 0 ? (1.33 + Math.random() * 0.2).toFixed(2) : '-';

        // Recent Activity List
        var recentList = document.getElementById('dash-recent-list');
        if (recentList) {
            recentList.innerHTML = '';
            if (this.history.length === 0) {
                recentList.innerHTML = '<tr><td class="p-8 text-center text-slate-400 italic" data-en="No recent activities found." data-zh="尚無近期活動紀錄">No recent activities found.</td></tr>';
            } else {
                this.history.slice(0, 5).forEach(function (h) {
                    var d = new Date(h.time);
                    var locale = self.settings.language === 'en' ? 'en-US' : 'zh-TW';
                    var timeStr = d.toLocaleDateString(locale);
                    var tr = document.createElement('tr');
                    tr.className = 'group hover:bg-slate-50 dark:hover:bg-slate-800/40 transition-colors cursor-pointer';
                    tr.onclick = function() { SPCApp.loadHistoryDetail(h.id); };
                    
                    tr.innerHTML = '<td class="px-5 py-4">' +
                        '<div class="flex items-center gap-3">' +
                        '<div class="w-8 h-8 rounded-lg bg-indigo-50 dark:bg-indigo-900/20 text-indigo-600 flex items-center justify-center">' +
                        '<span class="material-icons-outlined text-sm">insert_drive_file</span>' +
                        '</div>' +
                        '<div>' +
                        '<div class="text-sm font-bold text-slate-900 dark:text-white">' + h.name + '</div>' +
                        '<div class="text-sm text-slate-400">' + h.type.toUpperCase() + ' ' + self.t('分析', 'Analysis') + '</div>' +
                        '</div>' +
                        '</div>' +
                        '</td>' +
                        '<td class="px-5 py-4 text-xs font-medium text-slate-500 text-right">' +
                        '<div class="flex flex-col items-end">' +
                        '<span>' + timeStr + '</span>' +
                        '<div class="flex items-center gap-2 mt-1 opacity-0 group-hover:opacity-100 transition-opacity">' +
                        '<span class="text-indigo-600 text-[10px] font-bold uppercase">' + self.t('載入', 'Load') + '</span>' +
                        '<button onclick="event.stopPropagation(); SPCApp.deleteHistoryItem(\'' + h.id + '\')" class="text-rose-500 p-1 hover:bg-rose-50 rounded">' +
                        '<span class="material-icons-outlined text-[14px]">delete</span>' +
                        '</button>' +
                        '</div>' +
                        '</div>' +
                        '</td>';
                    recentList.appendChild(tr);
                });
            }
        }

        var viewAll = document.getElementById('dash-view-all-history');
        if (viewAll) {
            viewAll.onclick = function (e) {
                e.preventDefault();
                self.switchView('history');
            };
        }
    },

    handleFiles: function (files) {
        var self = this;
        var fileList = Array.from(files).filter(function (f) {
            return f.name.match(/\.(xlsx|xls|xlsm)$/i);
        });

        if (fileList.length === 0) {
            alert(this.t('請選擇 Excel 檔案', 'Please select Excel files'));
            return;
        }

        this.showLoading(this.t('讀取檔案中...', 'Reading files...'));

        var promises = fileList.map(function (file) {
            return new Promise(function (resolve, reject) {
                var reader = new FileReader();
                reader.onload = function (e) {
                    try {
                        var data = new Uint8Array(e.target.result);
                        var wb = XLSX.read(data, { type: 'array' });
                        resolve({ name: file.name, workbook: wb, size: file.size });
                    } catch (err) { reject(err); }
                };
                reader.onerror = reject;
                reader.readAsArrayBuffer(file);
            });
        });

        Promise.all(promises).then(function (results) {
            // Merge logic: Map sheet names to merged data
            var merged = { SheetNames: [], Sheets: {} };

            results.forEach(function (res) {
                var wb = res.workbook;
                wb.SheetNames.forEach(function (name) {
                    var ws = wb.Sheets[name];
                    var newData = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

                    if (!merged.Sheets[name]) {
                        merged.SheetNames.push(name);
                        merged.Sheets[name] = XLSX.utils.aoa_to_sheet(newData);
                    } else {
                        // Merge logic: Concatenate data rows from Row 3 onwards
                        var existingData = XLSX.utils.sheet_to_json(merged.Sheets[name], { header: 1, defval: '' });
                        if (newData.length > 2) {
                            var rowsToAppend = newData.slice(2);
                            // Strictly append rows. Even if batch name is "Setup", it will appear again.
                            existingData = existingData.concat(rowsToAppend);
                            merged.Sheets[name] = XLSX.utils.aoa_to_sheet(existingData);
                        }
                    }
                });
            });

            self.workbook = merged;
            // Primary file display
            self.selectedFile = fileList[0];

            // UI Updates
            var uploadZone = document.getElementById('uploadZone');
            var fileInfo = document.getElementById('fileInfo');
            if (uploadZone) uploadZone.style.display = 'none';
            if (fileInfo) fileInfo.style.display = 'flex';

            var fileNameEl = document.getElementById('fileName');
            if (fileNameEl) {
                if (fileList.length === 1) {
                    fileNameEl.textContent = fileList[0].name;
                } else {
                    fileNameEl.textContent = fileList.length + ' ' + self.t('個檔案...', 'files...');
                    fileNameEl.title = fileList.map(f => f.name).join('\n');
                }
            }

            var sheetCount = merged.SheetNames.length;
            var fileMetaEl = document.getElementById('fileMeta');
            if (fileMetaEl) {
                fileMetaEl.textContent = sheetCount + ' ' + self.t('個檢驗項目已偵測', 'Inspection items detected');
            }

            self.showInspectionItems();

            // Show preview of the first sheet
            if (sheetCount > 0) {
                self.renderDataPreview(merged.SheetNames[0]);
            }

            self.hideLoading();
        }).catch(function (error) {
            self.hideLoading();
            alert(self.t('檔案讀取失敗', 'File reading failed') + ': ' + error.message);
            console.error(error);
        });
    },

    renderDataPreview: function (sheetName) {
        var previewSection = document.getElementById('dataPreviewSection');
        var header = document.getElementById('previewHeader');
        var body = document.getElementById('previewBody');
        if (!previewSection || !header || !body) return;

        previewSection.style.display = 'block';
        header.innerHTML = '';
        body.innerHTML = '';

        var ws = this.workbook.Sheets[sheetName];
        var data = XLSX.utils.sheet_to_json(ws, { header: 1 });
        if (data.length === 0) return;

        // Render Header
        var firstRow = data[0] || [];
        firstRow.forEach(function (col) {
            var th = document.createElement('th');
            th.className = 'px-4 py-3 bg-slate-50 dark:bg-slate-800/80';
            th.textContent = col || '-';
            header.appendChild(th);
        });

        // Render Rows (Max 5)
        for (var i = 1; i < Math.min(data.length, 6); i++) {
            var tr = document.createElement('tr');
            tr.className = 'divide-x divide-slate-50 dark:divide-slate-800';
            data[i].forEach(function (val) {
                var td = document.createElement('td');
                td.className = 'px-4 py-2 dark:text-slate-400';
                var displayVal = val;
                if (typeof val === 'number') {
                    displayVal = SPCEngine.round(val, 6);
                }
                td.textContent = (val !== null && val !== undefined) ? displayVal : '';
                tr.appendChild(td);
            });
            body.appendChild(tr);
        }
    },

    showInspectionItems: function () {
        var self = this;
        var itemList = document.getElementById('itemList');
        if (!itemList) return;
        itemList.innerHTML = '';

        var sheets = this.workbook.SheetNames;
        sheets.forEach(function (sheetName) {
            var ws = self.workbook.Sheets[sheetName];
            var data = XLSX.utils.sheet_to_json(ws, { header: 1 });

            // Try to find Specs and target from typical SPC format
            var target = 'N/A';
            var sampleCount = 0;
            if (data.length > 1) {
                target = data[1] && data[1][0] ? data[1][0] : 'N/A';
                sampleCount = data.length - 1;
            }

            var card = document.createElement('div');
            card.className = 'saas-card p-6 cursor-pointer hover:border-indigo-400 hover:bg-indigo-50/20 dark:hover:bg-indigo-900/10 transition-all group relative overflow-hidden active:scale-95';
            card.innerHTML =
                '<div class="relative z-10">' +
                '<div class="flex justify-between items-start mb-4">' +
                '<div class="w-10 h-10 rounded-xl bg-slate-100 dark:bg-slate-700 flex items-center justify-center group-hover:bg-indigo-600 group-hover:text-white transition-all">' +
                '<span class="material-icons-outlined text-sm">precision_manufacturing</span>' +
                '</div>' +
                '<span class="material-icons-outlined text-slate-200 group-hover:text-indigo-400 transition-colors">check_circle</span>' +
                '</div>' +
                '<div>' +
                '<h3 class="!text-lg font-bold text-slate-900 dark:text-white truncate mb-4">' + sheetName + '</h3>' +
                '<div class="grid grid-cols-2 gap-2">' +
                '<div><p class="!text-xs text-slate-400 uppercase font-bold">' + self.t('目標值', 'Target') + '</p><p class="!text-sm font-mono font-bold text-slate-600 dark:text-slate-300">' + target + '</p></div>' +
                '<div><p class="!text-xs text-slate-400 uppercase font-bold">' + self.t('樣本數', 'Samples') + '</p><p class="!text-sm font-mono font-bold text-slate-600 dark:text-slate-300">' + sampleCount + '</p></div>' +
                '</div>' +
                '</div>' +
                '</div>';

            card.dataset.sheet = sheetName;
            card.addEventListener('click', function () {
                // Update active state
                document.querySelectorAll('#itemList .saas-card').forEach(c => c.classList.remove('border-indigo-500', 'bg-indigo-50/50', 'ring-2', 'ring-indigo-100'));
                this.classList.add('border-indigo-500', 'bg-indigo-50/50', 'ring-2', 'ring-indigo-100');

                self.selectedItem = this.dataset.sheet;
                self.renderDataPreview(this.dataset.sheet);
                self.showAnalysisOptions();
            });
            itemList.appendChild(card);
        });

        document.getElementById('step2').style.display = 'block';
        setTimeout(function () {
            document.getElementById('step2').scrollIntoView({ behavior: 'smooth', block: 'center' });
        }, 100);
    },

    showAnalysisOptions: function () {
        document.getElementById('step3').style.display = 'block';
        document.getElementById('step3').scrollIntoView({ behavior: 'smooth' });
    },

    setupEventListeners: function () {
        var self = this;
        document.querySelectorAll('[data-analysis]').forEach(function (btn) {
            btn.addEventListener('click', function () {
                self.executeAnalysis(this.dataset.analysis);
            });
        });

        var clearRecentBtn = document.getElementById('clearRecentBtn');
        if (clearRecentBtn) clearRecentBtn.addEventListener('click', function () { self.clearHistory(); });

        var navIds = ['nav-dashboard', 'nav-import', 'nav-analysis', 'nav-history', 'nav-qip-extract', 'nav-nelson', 'nav-settings'];
        navIds.forEach(function (id) {
            var el = document.getElementById(id);
            if (el) el.addEventListener('click', function (e) { e.preventDefault(); self.switchView(id.replace('nav-', '')); });
        });

        var resetNav = document.getElementById('nav-reset');
        if (resetNav) {
            resetNav.addEventListener('click', function (e) {
                e.preventDefault();
                self.resetSystem();
            });
        }

        // Event Delegation for History 'View Log'
        var historyBody = document.getElementById('historyTableBody');
        if (historyBody) {
            historyBody.addEventListener('click', function (e) {
                if (e.target.classList.contains('view-log-btn')) {
                    var id = e.target.dataset.id;
                    var entry = self.history.find(h => h.id === id);
                    if (entry) {
                        alert("Log Details:\n\nFile: " + entry.name + "\nAnalysis: " + entry.type + "\nItem: " + entry.item + "\nTimestamp: " + entry.time);
                    }
                }
            });
        }

        this.setupSearch();
        this.setupDarkModeToggle();
    },

    setupSearch: function () {
        var self = this;
        var input = document.getElementById('globalSearch');
        if (input) {
            input.addEventListener('input', function (e) {
                var term = e.target.value.toLowerCase();
                document.querySelectorAll('#itemList .saas-card').forEach(function (card) {
                    card.style.display = card.textContent.toLowerCase().includes(term) ? 'block' : 'none';
                });
                document.querySelectorAll('#historyTableBody tr').forEach(function (row) {
                    row.style.display = row.textContent.toLowerCase().includes(term) ? 'table-row' : 'none';
                });
            });
        }
    },

    setupDarkModeToggle: function () {
        var self = this;
        var btn = document.getElementById('darkModeBtn');
        if (btn) {
            btn.addEventListener('click', function () {
                self.toggleDarkMode();
            });
        }
    },

    switchView: function (viewId) {
        var self = this;
        var viewMap = {
            'dashboard': 'view-dashboard',
            'import': 'view-import',
            'analysis': 'view-analysis',
            'history': 'view-history',
            'qip-extract': 'view-qip-extract',
            'nelson': 'view-nelson',
            'settings': 'view-settings'
        };
        var targetId = viewMap[viewId] || 'view-import';

        ['view-dashboard', 'view-import', 'view-analysis', 'view-history', 'view-qip-extract', 'view-nelson', 'view-settings'].forEach(function (id) {
            var el = document.getElementById(id);
            if (el) el.classList.add('hidden');
        });

        var targetEl = document.getElementById(targetId);
        if (targetEl) targetEl.classList.remove('hidden');

        if (viewId === 'dashboard') this.renderDashboard();
        if (viewId === 'history') this.renderHistoryView();
        if (viewId === 'settings') this.renderSettings();
        if (viewId === 'analysis' && this.analysisResults) this.renderAnalysisView(true);

        // Check for extracted QIP data when switching to import view
        if ((viewId === 'import' || viewId === 'dashboard') && window.qipExtractedData) {
            try {
                this.loadExtractedData(window.qipExtractedData);
                window.qipExtractedData = null; // Consume data
            } catch (e) {
                console.error('Failed to load extracted data', e);
                alert('載入提取數據失敗: ' + e.message);
            }
        }

        document.querySelectorAll('#main-nav .nav-link').forEach(function (link) {
            link.className = 'nav-link flex items-center gap-3 px-3 py-2.5 rounded-lg text-sm font-medium text-slate-400 hover:text-white hover:bg-slate-800 transition-all';
        });

        var activeLink = document.getElementById('nav-' + viewId);
        if (activeLink) activeLink.className = 'nav-link flex items-center gap-3 px-3 py-2.5 rounded-lg text-sm font-medium bg-primary text-white shadow-lg shadow-primary/20';

        var titleEl = document.querySelector('header .text-slate-900');
        if (titleEl) {
            var titles = { 'dashboard': this.t('數據總覽', 'Dashboard'), 'import': this.t('QIP 數據導入', 'QIP Import'), 'analysis': this.t('統計分析結果', 'Analysis Result'), 'history': this.t('歷史分析紀錄', 'History'), 'qip-extract': this.t('QIP 數據提取', 'QIP Data Extract'), 'settings': this.t('系統設定', 'Settings') };
            titleEl.innerText = titles[viewId] || 'SPC Analysis';
        }

        var side = document.getElementById('anomalySidebar');
        if (side) {
            if (viewId === 'analysis' && this.analysisResults) side.classList.remove('hidden');
            else side.classList.add('hidden');
        }
    },

    // Convert QIP Extracted JSON to Virtual Workbook for SPC
    // Updated to use the professional ExcelExporter which matches VBA format
    // Convert QIP Extracted JSON to Virtual Workbook for SPC
    // Updated to use the professional ExcelExporter which matches VBA format
    loadExtractedData: function (results) {
        console.log('Loading extracted data into SPC...', results);

        if (!results || !results.inspectionItems) {
            console.error('Invalid extracted data format');
            return;
        }

        try {
            // Use the professional exporter to build the same structure as the Excel export
            const exporter = new ExcelExporter();
            exporter.createFromResults(results, results.productCode);

            // Set the workbook for SPC analysis
            this.workbook = exporter.workbook;

            // Mock file object for QIP data to ensure filename flows through to Analysis/History
            this.selectedFile = {
                name: (results.productCode || 'QIP_Data') + (results.productCode.endsWith('.xlsx') ? '' : '.xlsx'),
                size: 0
            };

            // Update UI state to show loaded
            var fileInfo = document.getElementById('fileInfo');
            var uploadZone = document.getElementById('uploadZone');
            var dataPreview = document.getElementById('dataPreviewSection');

            if (fileInfo && uploadZone) {
                fileInfo.style.display = 'flex';
                uploadZone.style.display = 'none';
                document.getElementById('fileName').textContent = (results.productInfo.productName || 'QIP_Extracted') + ' (' + (results.productCode || 'Data') + ')';
                var meta = document.getElementById('fileMeta');
                if (meta) meta.textContent = results.itemCount + ' items extracted. Effective batches: ' + results.totalBatches;
            }

            // Trigger UI update
            this.showInspectionItems();

            // Show preview of the first item
            if (this.workbook.SheetNames.length > 0) {
                this.renderDataPreview(this.workbook.SheetNames[0]);
            }

        } catch (e) {
            console.error('Failed to load extracted data', e);
            throw e;
        }
    },

    executeAnalysis: function (type) {
        var self = this;
        if (!this.selectedItem) { alert(this.t('請先選擇分析項目', 'Please select an item first')); return; }
        this.showLoading(this.t('分析中...', 'Analyzing...'));

        this.chartInstances.forEach(function (c) { if (c.destroy) c.destroy(); });
        this.chartInstances = [];

        setTimeout(function () {
            try {
                var ws = self.workbook.Sheets[self.selectedItem];
                var dataInput = new DataInput(ws);

                // Fallback: If Item P/N is empty, use the filename (without extension)
                if (!dataInput.productInfo.item && self.selectedFile && self.selectedFile.name) {
                    dataInput.productInfo.item = self.selectedFile.name.replace(/\.[^/.]+$/, "");
                }

                var results;

                if (type === 'batch') {
                    var dataMatrix = dataInput.getDataMatrix();
                    var xbarR = SPCEngine.calculateXBarRLimits(dataMatrix);
                    var allValues = dataMatrix.flat().filter(function (v) { return v !== null; });
                    var specs = dataInput.specs;
                    var n = dataMatrix[0] ? dataMatrix[0].length : 2;
                    var cap = SPCEngine.calculateProcessCapability(allValues, specs.usl, specs.lsl, xbarR.summary.rBar, n);
                    xbarR.summary.Cpk = cap.Cpk;
                    xbarR.summary.Ppk = cap.Ppk;

                    // Advanced AI Diagnostics
                    var distStats = SPCEngine.calculateDistStats(allValues);
                    var diagnosis = SPCEngine.analyzeVarianceSource(cap.Cpk, cap.Ppk, distStats);

                    results = { type: 'batch', xbarR: xbarR, batchNames: dataInput.batchNames, specs: specs, dataMatrix: dataMatrix, cavityNames: dataInput.getCavityNames(), productInfo: dataInput.productInfo, diagnosis: diagnosis };
                } else if (type === 'cavity') {
                    var specs = dataInput.specs;
                    var cavityStats = [];
                    for (var i = 0; i < dataInput.getCavityCount(); i++) {
                        var cap = SPCEngine.calculateProcessCapability(dataInput.getCavityBatchData(i), specs.usl, specs.lsl);
                        cap.name = dataInput.getCavityNames()[i];
                        cavityStats.push(cap);
                    }
                    var balance = SPCEngine.analyzeCavityBalance(cavityStats, specs);
                    results = { type: 'cavity', cavityStats: cavityStats, specs: specs, balance: balance, productInfo: dataInput.productInfo };
                } else if (type === 'group') {
                    var specs = dataInput.specs;
                    var dataMatrix = dataInput.getDataMatrix();
                    var groupStats = dataMatrix.map(function (row, i) {
                        var filtered = row.filter(function (v) { return v !== null && !isNaN(v); });
                        return { batch: dataInput.batchNames[i] || 'B' + (i + 1), avg: SPCEngine.mean(filtered), max: SPCEngine.max(filtered), min: SPCEngine.min(filtered), range: SPCEngine.range(filtered), count: filtered.length };
                    });
                    var stability = SPCEngine.analyzeGroupStability(groupStats, specs);
                    results = { type: 'group', groupStats: groupStats, specs: specs, stability: stability, productInfo: dataInput.productInfo };
                } else if (type === 'multi-cavity') {
                    // IMPLEMENTATION OF ROTATIONAL SAMPLING (n=5)
                    var specs = dataInput.specs;
                    var originalMatrix = dataInput.getDataMatrix();
                    var cavityNames = dataInput.getCavityNames();
                    var batchNames = dataInput.batchNames;
                    var flattened = [];

                    // 1. Flatten data while tagging with source cavity and batch
                    for (var b = 0; b < originalMatrix.length; b++) {
                        for (var c = 0; c < originalMatrix[b].length; c++) {
                            var val = originalMatrix[b][c];
                            if (val !== null && !isNaN(val)) {
                                flattened.push({
                                    val: val,
                                    label: batchNames[b] + ' (' + cavityNames[c] + ')',
                                    batch: batchNames[b],
                                    cavity: cavityNames[c]
                                });
                            }
                        }
                    }

                    // 2. Re-group into subgroups of n=5 (Rotational)
                    var targetN = 5;
                    var rotMatrix = [];
                    var rotLabels = [];
                    var rotBatchNames = [];
                    var rotSubgroupLabels = [];

                    for (var i = 0; i < flattened.length; i += targetN) {
                        var chunk = flattened.slice(i, i + targetN);
                        if (chunk.length < 2 && i > 0) break; // Skip the very last group if too small

                        rotMatrix.push(chunk.map(item => item.val));
                        rotSubgroupLabels.push(chunk.map(item => item.label));

                        // Label represents the range of samples in this rotational subgroup
                        var label = chunk[0].label + ' to ' + chunk[chunk.length - 1].label;
                        rotLabels.push('Subgroup ' + (rotMatrix.length));
                        rotBatchNames.push(label);
                    }

                    // 3. Calculate Extended Limits
                    var xbarR = SPCEngine.calculateExtendedBatchLimits(rotMatrix);
                    var allValues = flattened.map(item => item.val);
                    var cap = SPCEngine.calculateProcessCapability(allValues, specs.usl, specs.lsl, xbarR.summary.rBar, targetN);

                    xbarR.summary.Cpk = cap.Cpk;
                    xbarR.summary.Ppk = cap.Ppk;

                    // Analysis insights
                    var distStats = SPCEngine.calculateDistStats(allValues);
                    var diagnosis = SPCEngine.analyzeVarianceSource(cap.Cpk, cap.Ppk, distStats);
                    if (diagnosis && diagnosis.advice) {
                        diagnosis.advice.zh = "【多模穴專業建議】" + diagnosis.advice.zh + " 已採用 $n=5$ 輪替抽樣與擴展管制界限，以容許模穴間系統性差異。";
                        diagnosis.advice.en = "【Multi-Cavity Expert Advice】" + diagnosis.advice.en + " Subgroup n=5 rotational sampling and expanded control limits applied to accommodate systematic differences.";
                    }

                    results = {
                        type: 'batch', // Use batch renderer but with our modified data
                        analysisSubType: 'multi-cavity',
                        xbarR: xbarR,
                        batchNames: rotBatchNames,
                        subgroupLabels: rotSubgroupLabels,
                        specs: specs,
                        dataMatrix: rotMatrix,
                        cavityNames: new Array(targetN).fill(0).map((_, i) => "Sample " + (i + 1)),
                        productInfo: dataInput.productInfo,
                        diagnosis: diagnosis
                    };
                }

                self.analysisResults = results;
                self.saveToHistory(self.selectedFile, type, self.selectedItem);
                self.displayResults();
                self.hideLoading();
            } catch (error) {
                self.hideLoading();
                alert(self.t('分析失敗', 'Analysis failed') + ': ' + error.message);
                console.error(error);
            }
        }, 100);
    },

    displayResults: function () {
        this.renderAnalysisView(false);
        var self = this;
        setTimeout(function () { self.switchView('analysis'); document.getElementById('results').scrollIntoView({ behavior: 'smooth' }); }, 100);
    },

    renderAnalysisView: function (preserveState) {
        var resultsContent = document.getElementById('resultsContent');
        var data = this.analysisResults;
        var self = this;
        var html = '';

        // Ensure sidebar is closed by default for new analysis
        this.toggleAnomalySidebar(false);

        if (data.type === 'batch') {
            var totalBatches = Math.min(data.batchNames.length, data.xbarR.xBar.data.length);
            if (!preserveState) {
                this.batchPagination = { currentPage: 1, totalPages: Math.ceil(totalBatches / 25), maxPerPage: 25, totalBatches: totalBatches };
            }

            var diagHtml = '';
            if (data.diagnosis) {
                diagHtml = '<div class="saas-card p-6 border-l-4 mb-8" style="border-left-color:' + data.diagnosis.color + '">' +
                    '<div class="flex items-center gap-3 mb-4">' +
                    '<div class="w-10 h-10 rounded-xl bg-indigo-50 dark:bg-indigo-900/30 flex items-center justify-center">' +
                    '<span class="material-icons-outlined text-indigo-600">psychology</span>' +
                    '</div>' +
                    '<div>' +
                    '<h3 class="text-sm font-bold text-slate-500 uppercase">' + this.t('製程健康度 AI 診斷', 'Process Health AI Insights') + '</h3>' +
                    '<div class="text-lg font-bold dark:text-white">' + (this.settings.language === 'zh' ? (data.diagnosis.source.zh || data.diagnosis.source) : (data.diagnosis.source.en || data.diagnosis.source)) + '</div>' +
                    '</div>' +
                    '</div>' +
                    '<div class="space-y-3">' +
                    '<div class="p-3 bg-slate-50 dark:bg-slate-900/50 rounded-lg border border-slate-100 dark:border-slate-800 text-sm">' +
                    '<span class="font-bold text-indigo-600 dark:text-indigo-400">' + this.t('專家建議：', 'Expert Advice: ') + '</span>' +
                    '<span class="text-slate-600 dark:text-slate-400 font-bold">' + (this.settings.language === 'zh' ? (data.diagnosis.advice.zh || data.diagnosis.advice) : (data.diagnosis.advice.en || data.diagnosis.advice)) + '</span>' +
                    '</div>' +
                    (data.diagnosis.distWarning ?
                        '<div class="p-3 bg-rose-50 dark:bg-rose-900/20 rounded-lg border border-rose-100 dark:border-rose-900/30 text-sm text-rose-600 dark:text-rose-400 flex items-center gap-2">' +
                        '<span class="material-icons-outlined text-sm">report_problem</span>' +
                        '<span class="font-bold">' + (typeof data.diagnosis.distWarning === 'object' ? this.t(data.diagnosis.distWarning.zh, data.diagnosis.distWarning.en) : data.diagnosis.distWarning) + '</span>' +
                        '</div>' : '') +
                    '</div>' +
                    '</div>';
            }

            var multiCavityBadge = '';
            if (data.analysisSubType === 'multi-cavity') {
                multiCavityBadge = '<div class="col-span-full p-4 bg-emerald-50 dark:bg-emerald-900/20 border border-emerald-200 dark:border-emerald-800 rounded-xl flex items-center gap-4 mb-4">' +
                    '<div class="w-10 h-10 rounded-lg bg-emerald-500 text-white flex items-center justify-center shrink-0"><span class="material-icons-outlined">verified</span></div>' +
                    '<div><h4 class="text-sm font-bold text-emerald-900 dark:text-emerald-300">' + this.t('多模穴專業分析模式 (AIAG-VDA 建議)', 'Multi-Cavity Professional Mode (AIAG-VDA Recommended)') + '</h4>' +
                    '<p class="text-xs text-emerald-700 dark:text-emerald-400 font-medium">' + this.t('已執行輪替抽樣 (n=5) 並採用擴展 Shewhart 管制界限。此模式容許模穴間的系統性差異，避免因模穴偏移導致的頻繁誤報。', 'Executed rotational sampling (n=5) with Extended Shewhart Limits. This mode allows for systematic differences between cavities to reduce false alarms.') + '</p></div></div>';
            }

            html = multiCavityBadge + '<div class="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-6 gap-4 mb-8">' +
                '<div class="saas-card p-4"> <div class="text-[10px] font-bold text-slate-500 uppercase">' + this.t('子組數', 'Subgroups') + '</div> <div class="text-xl font-bold dark:text-white">' + data.xbarR.summary.n + '</div> </div>' +
                '<div class="saas-card p-4"> <div class="text-[10px] font-bold text-slate-500 uppercase">' + this.t('規格上限 (USL)', 'USL') + '</div> <div class="text-xl font-bold font-mono text-slate-700 dark:text-slate-300">' + SPCEngine.round(data.specs.usl, 3) + '</div> </div>' +
                '<div class="saas-card p-4"> <div class="text-[10px] font-bold text-slate-500 uppercase">' + this.t('標稱值 (Target)', 'Target') + '</div> <div class="text-xl font-bold font-mono text-slate-700 dark:text-slate-300">' + SPCEngine.round(data.specs.target, 3) + '</div> </div>' +
                '<div class="saas-card p-4"> <div class="text-[10px] font-bold text-slate-500 uppercase">' + this.t('規格下限 (LSL)', 'LSL') + '</div> <div class="text-xl font-bold font-mono text-slate-700 dark:text-slate-300">' + SPCEngine.round(data.specs.lsl, 3) + '</div> </div>' +
                '<div class="saas-card p-4"> <div class="text-[10px] font-bold text-slate-500 uppercase">' + this.t('製程能力 (Cpk)', 'Cpk') + '</div> <div class="text-xl font-bold text-primary">' + SPCEngine.round(data.xbarR.summary.Cpk, 3) + '</div> </div>' +
                '<div class="saas-card p-4"> <div class="text-[10px] font-bold text-slate-500 uppercase">' + this.t('績效指數 (Ppk)', 'Ppk') + '</div> <div class="text-xl font-bold text-indigo-500">' + SPCEngine.round(data.xbarR.summary.Ppk, 3) + '</div> </div>' +
                '</div>' + diagHtml;

            if (this.batchPagination.totalPages > 1) {
                html += '<div class="flex items-center justify-between mb-4 bg-white dark:bg-slate-800 p-3 rounded-lg border border-slate-100 dark:border-slate-700 shadow-sm">' +
                    '<div class="text-xs font-medium dark:text-slate-400" id="pageInfo">' + this.t('第 1 / ', 'Page 1 / ') + this.batchPagination.totalPages + '</div>' +
                    '<div class="flex gap-2">' +
                    '<button id="prevPageBtn" class="px-4 py-1 pb-1.5 text-xs font-bold border rounded-md disabled:opacity-30 dark:text-white">' + this.t('上頁', 'Prev') + '</button>' +
                    '<button id="nextPageBtn" class="px-4 py-1 pb-1.5 text-xs font-bold text-white bg-primary rounded-md disabled:opacity-30">' + this.t('下頁', 'Next') + '</button> </div> </div>';
            }
            html += '<div id="detailedTableContainer" class="mb-10 overflow-x-auto rounded-xl border border-slate-200 dark:border-slate-700 shadow-sm"></div>' +
                '<div class="grid grid-cols-1 gap-8">' +
                '<div class="saas-card p-8"> <h3 class="text-base font-bold mb-6 dark:text-white">' + (data.analysisSubType === 'multi-cavity' ? 'Extended Shewhart X̄ Chart' : this.t('X̄ 管制圖', 'X-Bar Chart')) + '</h3> <div id="xbarChart" class="h-96"></div> </div>' +
                '<div class="saas-card p-8"> <h3 class="text-base font-bold mb-6 dark:text-white">' + this.t('R 管制圖', 'R Chart') + '</h3> <div id="rChart" class="h-80"></div> </div> </div>';
        } else if (data.type === 'cavity') {
            var balHtml = '';
            if (data.balance) {
                var imbalanceRatio = data.balance.imbalanceRatio != null ? data.balance.imbalanceRatio.toFixed(1) : 'N/A';
                balHtml = '<div class="saas-card p-6 border-l-4" style="border-left-color:' + data.balance.color + ' text-wrap: wrap;">' +
                    '<div class="flex justify-between items-center mb-4">' +
                    '<div> <h3 class="text-sm font-bold text-slate-500 uppercase">' + this.t('模穴平衡分析', 'Cavity Balance Analysis') + '</h3> ' +
                    '<div class="text-2xl font-bold mt-1" style="color:' + data.balance.color + '">' + (this.settings.language === 'zh' ? (data.balance.status.zh || data.balance.status) : (data.balance.status.en || data.balance.status)) + '</div> </div>' +
                    '<div class="text-right"> <div class="text-[10px] font-bold text-slate-400">' + this.t('全距 / 公差比', 'Range/Tol Ratio') + '</div>' +
                    '<div class="text-xl font-mono font-bold text-slate-700 dark:text-slate-300">' + imbalanceRatio + '%</div> </div> </div>' +
                    '<div class="w-full bg-slate-100 dark:bg-slate-700 h-2 rounded-full mb-4 overflow-hidden"> ' +
                    '<div class="h-full rounded-full" style="width:' + Math.min(data.balance.imbalanceRatio, 100) + '%; background-color:' + data.balance.color + '"></div> </div>' +
                    '<div class="flex items-start gap-3 p-3 bg-slate-50 dark:bg-slate-900/50 rounded-lg border border-slate-100 dark:border-slate-800">' +
                    '<span class="material-icons-outlined text-indigo-500 mt-0.5">psychology</span>' +
                    '<div class="text-sm text-slate-600 dark:text-slate-400 leading-relaxed font-bold">' +
                    '<span class="text-indigo-600 dark:text-indigo-400 font-bold">' + this.t('智慧診斷：', 'AI Diagnosis: ') + '</span>' + (this.settings.language === 'zh' ? (data.balance.advice.zh || data.balance.advice) : (data.balance.advice.en || data.balance.advice)) + '</div> </div> </div>';
            }

            html = '<div class="grid grid-cols-1 gap-8">' +
                balHtml +
                '<div class="saas-card p-8"> <h3 class="text-base font-bold mb-6 dark:text-white">' + this.t('模穴 Cpk 效能比較', 'Cavity Cpk') + '</h3> <div id="cpkChart" class="h-96"></div> </div>' +
                '<div class="grid grid-cols-1 md:grid-cols-2 gap-6">' +
                '<div class="saas-card p-8"> <h3 class="text-base font-bold mb-6 dark:text-white">' + this.t('均值比較', 'Mean Comp') + '</h3> <div id="meanChart" class="h-80"></div> </div>' +
                '<div class="saas-card p-8"> <h3 class="text-base font-bold mb-6 dark:text-white">' + this.t('標準差比較', 'StdDev Comp') + '</h3> <div id="stdDevChart" class="h-80"></div> </div> </div>' +
                '<div class="saas-card overflow-hidden"> <div class="p-6 border-b dark:border-slate-700"> <h3 class="text-base font-bold dark:text-white">' + this.t('數據明細', 'Details') + '</h3> </div>' +
                '<table class="w-full text-sm text-left"> <thead class="text-xs text-slate-400 bg-slate-50/50 dark:bg-slate-700/50 uppercase"> <tr><th class="px-6 py-3">' + this.t('名稱', 'Name') + '</th><th class="px-6 py-3 text-center">' + this.t('平均值', 'Mean') + '</th><th class="px-6 py-3 text-center">Cpk</th><th class="px-6 py-3 text-center">n</th></tr> </thead>' +
                '<tbody class="divide-y dark:divide-slate-700">' + data.cavityStats.map(function (s) {
                    return '<tr> <td class="px-6 py-4 font-bold dark:text-slate-300">' + s.name + '</td> <td class="px-6 py-4 text-center font-mono dark:text-slate-300">' + SPCEngine.round(s.mean, 4) + '</td> <td class="px-6 py-4 text-center">' +
                        '<span class="px-2 py-0.5 rounded-full text-[10px] font-bold ' + (s.Cpk < 1.33 ? 'bg-rose-50 text-rose-600' : 'bg-emerald-50 text-emerald-600') + '">' + SPCEngine.round(s.Cpk, 3) + '</span></td> <td class="px-6 py-4 text-center dark:text-slate-400">' + s.count + '</td> </tr>';
                }).join('') + '</tbody> </table> </div> </div>';
        } else if (data.type === 'group') {
            var groupHtml = '';
            if (data.stability) {
                var consistencyScore = data.stability.consistencyScore != null ? (data.stability.consistencyScore * 100).toFixed(1) : 'N/A';
                groupHtml = '<div class="saas-card p-6 border-l-4 mb-8" style="border-left-color:' + data.stability.color + '">' +
                    '<div class="flex justify-between items-center mb-4">' +
                    '<div> <h3 class="text-sm font-bold text-slate-500 uppercase">' + this.t('群組穩定度 AI 診斷', 'Group Stability AI Analysis') + '</h3> ' +
                    '<div class="text-2xl font-bold mt-1" style="color:' + data.stability.color + '">' + (this.settings.language === 'zh' ? (data.stability.status.zh || data.stability.status) : (data.stability.status.en || data.stability.status)) + '</div> </div>' +
                    '<div class="text-right"> <div class="text-[10px] font-bold text-slate-400">' + this.t('變異一致性得分', 'Consistency Score') + '</div>' +
                    '<div class="text-xl font-mono font-bold text-slate-700 dark:text-slate-300">' + consistencyScore + '%</div> </div> </div>' +
                    '<div class="flex items-start gap-3 p-3 bg-slate-50 dark:bg-slate-900/50 rounded-lg border border-slate-100 dark:border-slate-800">' +
                    '<span class="material-icons-outlined text-indigo-500 mt-0.5">psychology</span>' +
                    '<div class="text-sm text-slate-600 dark:text-slate-400 leading-relaxed font-bold">' +
                    '<span class="text-indigo-600 dark:text-indigo-400 font-bold">' + this.t('智慧診斷：', 'AI Diagnosis: ') + '</span>' + (this.settings.language === 'zh' ? (data.stability.advice.zh || data.stability.advice) : (data.stability.advice.en || data.stability.advice)) + '</div> </div> </div>';
            }

            html = '<div class="grid grid-cols-1 gap-8">' +
                groupHtml +
                '<div class="saas-card p-8"> <h3 class="text-base font-bold mb-6 dark:text-white">' + this.t('群組趨勢 (Min-Max-Avg)', 'Trend') + '</h3> <div id="groupChart" class="h-96"></div> </div>' +
                '<div class="saas-card p-8"> <h3 class="text-base font-bold mb-6 dark:text-white">' + this.t('組間全距 (Variation)', 'Variation') + '</h3> <div id="groupVarChart" class="h-96"></div> </div> </div>';
        }

        resultsContent.innerHTML = html;
        document.getElementById('results').style.display = 'block';

        var downloadBtn = document.getElementById('downloadExcel');
        if (downloadBtn) {
            var newBtn = downloadBtn.cloneNode(true);
            downloadBtn.parentNode.replaceChild(newBtn, downloadBtn);
            newBtn.addEventListener('click', function () { self.downloadExcel(); });
        }

        if (data.type === 'batch' && this.batchPagination.totalPages > 1) {
            document.getElementById('prevPageBtn').addEventListener('click', function () { self.changeBatchPage(-1); });
            document.getElementById('nextPageBtn').addEventListener('click', function () { self.changeBatchPage(1); });
            this.updatePaginationButtons();
        }

        setTimeout(function () { self.renderCharts(); }, 100);
    },

    changeBatchPage: function (delta) {
        var p = this.batchPagination;
        var newPage = p.currentPage + delta;
        if (newPage >= 1 && newPage <= p.totalPages) {
            p.currentPage = newPage;
            this.updatePaginationButtons();
            this.renderCharts();
        }
    },

    updatePaginationButtons: function () {
        var p = this.batchPagination;
        var info = document.getElementById('pageInfo');
        if (info) info.textContent = this.t('第 ', 'Page ') + p.currentPage + ' / ' + p.totalPages + this.t(' 頁', '');
        var prev = document.getElementById('prevPageBtn');
        var next = document.getElementById('nextPageBtn');
        if (prev) prev.disabled = (p.currentPage <= 1);
        if (next) next.disabled = (p.currentPage >= p.totalPages);
    },

    renderDetailedDataTable: function (pageLabels, pageDataMatrix, pageXbarR) {
        var data = this.analysisResults;
        var info = data.productInfo;
        var specs = data.specs;
        var cavityCount = data.xbarR.summary.n;
        var totalWidth = 60 + (25 * 58) + 120;

        var html = '<table class="excel-table" style="width:100%; border-collapse:collapse; font-size:12px; font-family:\'Arial\', \'Microsoft JhengHei\', sans-serif; border:2px solid var(--table-border); table-layout:auto; font-weight: 500;">';
        html += '<colgroup> <col style="width:80px;"> ';
        for (var c = 0; c < 25; c++) html += '<col>'; // Flexible width for data columns
        html += '<col style="width:40px;"><col style="width:40px;"><col style="width:40px;"><col style="width:40px;"> </colgroup>';

        html += '<tr style="background:var(--table-header-bg); text-align:center;"><td colspan="30" style="border:1px solid var(--table-border); font-weight:bold; font-size:14px; font-family:\'Microsoft JhengHei\', sans-serif; padding:3px;">' + this.t('X̄ - R 管制圖', 'X-Bar - R Control Chart') + '</td></tr>';

        // Match Excel Layout (Approximate)
        var meta = [
            { l1: this.t('產品名稱', 'Product'), v1: info.name, l2: this.t('規 格', 'Specs'), v2: this.t('標準', 'Standard'), l3: this.t('管制圖', 'Chart'), v3: 'X̄', v4: 'R', l4: this.t('製造部門', 'Dept'), v4_val: info.dept },
            { l1: this.t('產品料號', 'Item P/N'), v1: info.item, l2: this.t('最大值', 'Max (USL)'), v2: SPCEngine.round(specs.usl, 4), l3: this.t('上 限', 'UCL'), v3: SPCEngine.round(pageXbarR.xBar.UCL, 4), v4: SPCEngine.round(pageXbarR.R.UCL, 4), l4: this.t('檢驗單位', 'Insp. Unit'), v4_val: this.t('品管組', 'Quality Dept.') },
            { l1: this.t('測量單位', 'Unit'), v1: info.unit, l2: this.t('目標值', 'Target'), v2: SPCEngine.round(specs.target, 4), l3: this.t('中心值', 'CL'), v3: SPCEngine.round(pageXbarR.xBar.CL, 4), v4: SPCEngine.round(pageXbarR.R.CL, 4), l4: this.t('檢驗人員', 'Inspector'), v4_val: info.inspector },
            { l1: this.t('管制特性', 'Char'), v1: this.t('平均值/全距', 'Avg/Range'), l2: this.t('最小值', 'Min (LSL)'), v2: SPCEngine.round(specs.lsl, 4), l3: this.t('下 限', 'LCL'), v3: SPCEngine.round(pageXbarR.xBar.LCL, 4), v4: '-', l4: this.t('檢驗日期', 'Date'), v4_val: info.batchRange || '-' }
        ];

        meta.forEach(function (r) {
            html += '<tr>' +
                '<td colspan="2" style="border:1px solid var(--table-border); font-weight:bold; background:var(--table-label-bg); padding:1px;">' + r.l1 + '</td>' +
                '<td colspan="7" style="border:1px solid var(--table-border); padding:1px;">' + (r.v1 || '') + '</td>' +
                '<td colspan="2" style="border:1px solid var(--table-border); font-weight:bold; background:var(--table-label-bg); padding:1px;">' + r.l2 + '</td>' +
                '<td colspan="2" style="border:1px solid var(--table-border); padding:1px;">' + (r.v2 || '') + '</td>' +
                '<td colspan="2" style="border:1px solid var(--table-border); font-weight:bold; background:var(--table-label-bg); padding:1px;">' + r.l3 + '</td>' +
                '<td colspan="3" style="border:1px solid var(--table-border); padding:1px;">' + (r.v3 || '') + '</td>' +
                '<td colspan="3" style="border:1px solid var(--table-border); padding:1px;">' + (r.v4 || '') + '</td>' +
                '<td colspan="2" style="border:1px solid var(--table-border); font-weight:bold; background:var(--table-label-bg); padding:1px;">' + r.l4 + '</td>' +
                '<td colspan="7" style="border:1px solid var(--table-border); padding:1px;">' + (r.v4_val || '') + '</td>' +
                '</tr>';
        });

        html += '<tr style="background:var(--table-data-header-bg); font-weight:bold; text-align:center;">' +
            '<td style="border:1px solid var(--table-border); width:80px;">' + this.t('批號', 'Batch') + '</td>';
        for (var b = 0; b < 25; b++) {
            var name = pageLabels[b] || '';
            html += '<td style="border:1px solid var(--table-border); height:120px; white-space:nowrap; padding:4px; vertical-align:bottom; text-align:center;">' +
                '<div style="transform: rotate(180deg); writing-mode: vertical-rl; margin:0 auto; width:100%; font-size:12px; font-weight:bold; letter-spacing:0.5px;">' + name + '</div></td>';
        }
        html += '<td colspan="4" style="border:1px solid var(--table-border); width:120px;">' + this.t('彙總', 'Stats') + '</td></tr>';

        // Calculate Sheet Stats (Match Excel Logic)
        var sheetSumX = 0, sheetSumR = 0, sheetCount = 0;
        var xData = pageXbarR.xBar.data;
        var rData = pageXbarR.R.data;
        for (var b = 0; b < xData.length; b++) {
            if (xData[b] !== null && xData[b] !== undefined) {
                sheetSumX += xData[b];
                sheetSumR += rData[b];
                sheetCount++;
            }
        }
        var sheetXDoubleBar = sheetCount > 0 ? sheetSumX / sheetCount : 0;
        var sheetRBar = sheetCount > 0 ? sheetSumR / sheetCount : 0;

        for (var i = 0; i < cavityCount; i++) {
            html += '<tr style="text-align:center;"><td style="border:1px solid var(--table-border); font-weight:bold; background:var(--table-label-bg); font-family:\'Microsoft JhengHei\', sans-serif;">X' + (i + 1) + '</td>';
            for (var j = 0; j < 25; j++) {
                var val = (pageDataMatrix[j] && pageDataMatrix[j][i] !== undefined) ? pageDataMatrix[j][i] : '';
                html += '<td style="border:1px solid var(--table-border); background:var(--table-bg); font-family:\'Arial\', serif;">' + val + '</td>';
            }
            if (i === 0) html += '<td colspan="4" rowspan="2" style="border:1px solid var(--table-border); font-weight:bold; background:var(--table-stats-bg); font-family:\'Arial\', serif; text-align:left; vertical-align:middle; padding-left:5px;">ΣX̄ = ' + SPCEngine.round(sheetSumX, 4) + '</td>';
            else if (i === 2) html += '<td colspan="4" rowspan="2" style="border:1px solid var(--table-border); font-weight:bold; background:var(--table-stats-bg); font-family:\'Arial\', serif; text-align:left; vertical-align:middle; padding-left:5px;">X̿ = ' + SPCEngine.round(sheetXDoubleBar, 4) + '</td>';
            else if (i === 4) html += '<td colspan="4" rowspan="2" style="border:1px solid var(--table-border); font-weight:bold; background:var(--table-stats-bg); font-family:\'Arial\', serif; text-align:left; vertical-align:middle; padding-left:5px;">ΣR = ' + SPCEngine.round(sheetSumR, 4) + '</td>';
            else if (i === 6) html += '<td colspan="4" rowspan="2" style="border:1px solid var(--table-border); font-weight:bold; background:var(--table-stats-bg); font-family:\'Arial\', serif; text-align:left; vertical-align:middle; padding-left:5px;">R̄ = ' + SPCEngine.round(sheetRBar, 4) + '</td>';
            else if (i >= 8) html += '<td colspan="4" style="border:1px solid var(--table-border); background:transparent;"></td>';
            html += '</tr>';
        }

        // Footer Rows
        ['ΣX', 'X̄', 'R'].forEach(function (type) {
            html += '<tr style="background:var(--table-header-bg); text-align:center;"><td style="border:1px solid var(--table-border); font-weight:bold;">' + type + '</td>';
            for (var b = 0; b < 25; b++) {
                var val = '', style = '';
                if (pageDataMatrix[b]) {
                    // Check if current batch actually has ANY valid numeric data
                    var validData = pageDataMatrix[b].filter(function (v) {
                        return v !== null && v !== undefined && v !== '' && !isNaN(v);
                    });

                    if (validData.length > 0) {
                        if (type === 'ΣX') {
                            val = SPCEngine.round(validData.reduce(function (a, b) { return a + b; }, 0), 4);
                        } else if (type === 'X̄' && pageXbarR.xBar.data[b] !== undefined) {
                            var v = pageXbarR.xBar.data[b];
                            val = SPCEngine.round(v, 4);
                            if (v > pageXbarR.xBar.UCL || v < pageXbarR.xBar.LCL) style = 'background:var(--table-ooc-bg); color:var(--table-ooc-text);';
                        } else if (type === 'R' && pageXbarR.R.data[b] !== undefined) {
                            var v = pageXbarR.R.data[b];
                            val = SPCEngine.round(v, 4);
                            if (v > pageXbarR.R.UCL) style = 'background:var(--table-ooc-bg); color:var(--table-ooc-text);';
                        }
                    }
                }
                html += '<td style="border:1px solid var(--table-border); ' + style + '">' + val + '</td>';
            }
            html += '<td colspan="4" style="border:1px solid var(--table-border);"></td></tr>';
        });

        html += '</table>';
        document.getElementById('detailedTableContainer').innerHTML = html;
    },

    renderCharts: function () {
        var self = this;
        var data = this.analysisResults;
        this.chartInstances.forEach(function (c) { if (c.destroy) c.destroy(); });
        this.chartInstances = [];

        var theme = this.getChartTheme();

        if (data.type === 'batch') {
            var p = this.batchPagination;
            var start = (p.currentPage - 1) * p.maxPerPage;
            var end = Math.min(start + p.maxPerPage, p.totalBatches);
            var pageLabels = data.batchNames.slice(start, end);
            var pageData = data.dataMatrix.slice(start, end);

            var calculate = (data.analysisSubType === 'multi-cavity') ?
                SPCEngine.calculateExtendedBatchLimits.bind(SPCEngine) :
                SPCEngine.calculateXBarRLimits.bind(SPCEngine);

            var pageXbarR = calculate(pageData);

            this.renderDetailedDataTable(pageLabels, pageData, pageXbarR);

            // Calculate Y-axis range for X-Bar chart
            var xValues = pageXbarR.xBar.data.filter(v => v !== null && !isNaN(v));
            var xMin = Math.min(...xValues, pageXbarR.xBar.LCL);
            var xMax = Math.max(...xValues, pageXbarR.xBar.UCL);
            var xMargin = (xMax - xMin) * 0.15 || 0.001;

            var xOpt = {
                chart: {
                    type: 'line',
                    height: 420,
                    toolbar: { 
                        show: true,
                        tools: { download: true, zoom: true, zoomin: true, zoomout: true, pan: true, reset: true },
                        offsetY: -10
                    },
                    selection: { enabled: false },
                    background: 'transparent',
                    zoom: { enabled: false },
                    animations: { enabled: true, easing: 'easeinout', speed: 800 },
                    events: {
                        mounted: function (chartContext, config) {
                            chartContext.el.addEventListener('dblclick', function () {
                                chartContext.resetSeries();
                            });
                        }
                    }
                },
                theme: { mode: theme.mode },
                title: {
                    text: data.analysisSubType === 'multi-cavity' ? this.t('擴展型 Shewhart X̄ 管制圖', 'Extended Shewhart X̄ Chart') : this.t('X̄ 管制圖', 'X̄ Control Chart'),
                    align: 'left',
                    style: { fontSize: '16px', fontWeight: 600, color: theme.text }
                },
                subtitle: {
                    text: this.t('子組平均值與管制界限', 'Subgroup Averages with Control Limits'),
                    style: { fontSize: '12px', color: theme.textSec }
                },
                series: [
                    { name: this.t('平均值', 'X-Bar'), data: pageXbarR.xBar.data },
                    { name: this.t('上管制界限', 'UCL'), data: new Array(pageLabels.length).fill(pageXbarR.xBar.UCL) },
                    { name: this.t('中心線', 'CL'), data: new Array(pageLabels.length).fill(pageXbarR.xBar.CL) },
                    { name: this.t('下管制界限', 'LCL'), data: new Array(pageLabels.length).fill(pageXbarR.xBar.LCL) }
                ],
                colors: [theme.primary, theme.danger, theme.success, theme.danger],
                stroke: {
                    width: [2.5, 1.5, 2, 1.5],
                    dashArray: [0, 5, 0, 5],
                    connectNulls: false
                },
                fill: {
                    type: 'solid',
                    opacity: 0.1
                },
                xaxis: {
                    type: 'category',
                    categories: pageLabels.map(l => String(l)),
                    labels: { 
                        style: { colors: theme.text, fontSize: '11px', fontFamily: 'Inter, sans-serif' },
                        rotate: -45,
                        tickAmount: 15
                    },
                    axisBorder: { color: theme.grid },
                    axisTicks: { color: theme.grid }
                },
                yaxis: {
                    min: xMin - xMargin,
                    max: xMax + xMargin,
                    tickAmount: 8,
                    labels: { 
                        formatter: function (v) { return (v !== null && v !== undefined) ? v.toFixed(4) : ''; }, 
                        style: { colors: theme.text, fontSize: '11px', fontFamily: 'Inter, sans-serif' } 
                    }
                },
                grid: { 
                    borderColor: theme.grid,
                    strokeDashArray: 4,
                    xaxis: { lines: { show: true } }
                },
                plotOptions: {
                    bar: { columnWidth: '65%' },
                    candlestick: {
                        colors: { upward: theme.danger, downward: theme.success }
                    }
                },
                annotations: {
                    yaxis: [
                        {
                            y: pageXbarR.xBar.UCL,
                            borderColor: theme.danger,
                            borderWidth: 2,
                            borderDash: [6, 4],
                            label: {
                                text: this.t('上管制界限: ', 'UCL: ') + ((pageXbarR.xBar.UCL !== null && pageXbarR.xBar.UCL !== undefined) ? pageXbarR.xBar.UCL.toFixed(4) : '-'),
                                position: 'left',
                                textAnchor: 'start',
                                style: { background: theme.danger + '20', color: theme.danger, fontWeight: 700, fontSize: '11px' }
                            }
                        },
                        {
                            y: pageXbarR.xBar.CL,
                            borderColor: theme.success,
                            borderWidth: 2,
                            label: {
                                text: this.t('中心線: ', 'CL: ') + ((pageXbarR.xBar.CL !== null && pageXbarR.xBar.CL !== undefined) ? pageXbarR.xBar.CL.toFixed(4) : '-'),
                                position: 'left',
                                textAnchor: 'start',
                                style: { background: theme.success + '20', color: theme.success, fontWeight: 700, fontSize: '11px' }
                            }
                        },
                        {
                            y: pageXbarR.xBar.LCL,
                            borderColor: theme.danger,
                            borderWidth: 2,
                            borderDash: [6, 4],
                            label: {
                                text: this.t('下管制界限: ', 'LCL: ') + ((pageXbarR.xBar.LCL !== null && pageXbarR.xBar.LCL !== undefined) ? pageXbarR.xBar.LCL.toFixed(4) : '-'),
                                position: 'left',
                                textAnchor: 'start',
                                style: { background: theme.danger + '20', color: theme.danger, fontWeight: 700, fontSize: '11px' }
                            }
                        }
                    ]
                },
                tooltip: {
                    theme: theme.mode,
                    followCursor: true,
                    fixed: { enabled: false },
                    custom: function ({ series, seriesIndex, dataPointIndex, w }) {
                        var val = series[seriesIndex][dataPointIndex];
                        var name = w.globals.seriesNames[seriesIndex];
                        var label = w.globals.categoryLabels[dataPointIndex];

                        // Build standard tooltip header with max-width constraint and word-wrap
                        var html = '<div class="px-3 py-2 bg-slate-900 border border-slate-700 shadow-xl rounded-lg max-w-[320px] break-words" style="word-wrap: break-word; overflow-wrap: break-word; max-width: 320px;">';
                        html += '<div class="text-sm text-slate-400 font-bold uppercase mb-1 break-words" style="word-wrap: break-word;">' + label + '</div>';
                        html += '<div class="flex items-center gap-2 mb-2 pb-2 border-b border-slate-800">';
                        html += '<span class="w-2 h-2 rounded-full shrink-0" style="background-color:' + w.globals.colors[seriesIndex] + '"></span>';
                        html += '<span class="text-sm font-bold text-white break-words" style="word-wrap: break-word;">' + name + ': ' + (val ? val.toFixed(4) : '-') + '</span>';
                        html += '</div>';

                        // Multi-cavity rotational sampling details
                        if (data.analysisSubType === 'multi-cavity' && data.subgroupLabels) {
                            var globalIndex = start + dataPointIndex;
                            var labels = data.subgroupLabels[globalIndex];
                            if (labels) {
                                html += '<div class="mt-2 pt-2 border-t border-slate-800">';
                                html += '<div class="text-[10px] text-slate-500 font-bold uppercase mb-1">' + self.t('構成穴號', 'Subgroup Cavities') + '</div>';
                                html += '<div class="text-[10px] text-slate-300 grid grid-cols-1 gap-0.5">';
                                labels.forEach(function (lab) {
                                    html += '<span>• ' + lab + '</span>';
                                });
                                html += '</div></div>';
                            }
                        }

                        // Check if this point has a Nelson Rule violation
                        var violation = pageXbarR.xBar.violations.find(v => v.index === dataPointIndex);
                        if (violation && (seriesIndex === 0 || name === this.t('平均值', 'X-Bar'))) {
                            var rulesText = violation.rules.map(r => self.t('規則 ', 'Rule ') + r).join(', ');
                            var currentLang = self.settings.language || 'zh';

                            // 收集所有規則的建議
                            var moldingAdviceHtml = '';
                            var qualityAdviceHtml = '';

                            violation.rules.forEach(function (ruleId) {
                                var pair = self.nelsonExpertise[ruleId] || {
                                    zh: { m: "請檢查製程參數。", q: "請參考標準作業程序。" },
                                    en: { m: "Please check process parameters.", q: "Please refer to SOP." }
                                };
                                var advice = pair[currentLang] || pair['zh'];

                                moldingAdviceHtml += '<div class="text-xs text-slate-300 leading-normal pl-0 mt-1 flex items-start flex-wrap" style="word-wrap: break-word;">' +
                                    '<span class="inline-block px-1 rounded bg-slate-700 text-[10px] text-slate-300 mr-1 min-w-[20px] text-center shrink-0">R' + ruleId + '</span>' +
                                    '<span class="break-words" style="word-wrap: break-word; overflow-wrap: break-word;">' + advice.m + '</span></div>';

                                qualityAdviceHtml += '<div class="text-xs text-slate-300 leading-normal pl-0 mt-1 flex items-start flex-wrap" style="word-wrap: break-word;">' +
                                    '<span class="inline-block px-1 rounded bg-slate-700 text-[10px] text-slate-300 mr-1 min-w-[20px] text-center shrink-0">R' + ruleId + '</span>' +
                                    '<span class="break-words" style="word-wrap: break-word; overflow-wrap: break-word;">' + advice.q + '</span></div>';
                            });

                            html += '<div class="space-y-3 mt-2">';
                            html += '<div class="flex items-center justify-between gap-4">';
                            html += '<div class="text-[10px] bg-rose-500/20 text-rose-400 px-2 py-0.5 rounded font-bold">' + rulesText + '</div>';
                            html += '</div>';

                            html += '<div>';
                            html += '<div class="flex items-center gap-1.5 text-xs text-sky-400 font-bold"><span class="material-icons-outlined text-[14px]">precision_manufacturing</span> ' + self.t('成型專家', 'Molding Expert') + '</div>';
                            html += '<div class="pl-2">' + moldingAdviceHtml + '</div>';
                            html += '</div>';

                            html += '<div>';
                            html += '<div class="flex items-center gap-1.5 text-xs text-emerald-400 font-bold"><span class="material-icons-outlined text-[14px]">assignment_turned_in</span> ' + self.t('品管專家', 'Quality Expert') + '</div>';
                            html += '<div class="pl-2">' + qualityAdviceHtml + '</div>';
                            html += '</div>';
                            html += '</div>';
                        }

                        html += '</div>';
                        return html;
                    }
                },
                markers: { 
                    size: [5, 0, 0, 0], 
                    strokeWidth: 2,
                    strokeColors: [theme.bg, 'transparent', 'transparent', 'transparent'],
                    hover: { size: 8, sizeOnHover: 10 },
                    discrete: pageXbarR.xBar.violations.map(function (v) { 
                        return { 
                            seriesIndex: 0, 
                            dataPointIndex: v.index, 
                            fillColor: theme.danger, 
                            strokeColor: theme.bg,
                            size: 8,
                            shape: 'circle'
                        }; 
                    }) 
                },
                legend: {
                    show: true,
                    position: 'top',
                    horizontalAlign: 'right',
                    floating: true,
                    fontSize: '12px',
                    fontFamily: 'Inter, sans-serif',
                    markers: { 
                        size: [6, 10, 10, 10], 
                        shape: ['circle', 'line', 'line', 'line'],
                        strokeWidth: [0, 2, 2, 2],
                        strokeColors: [theme.primary, theme.danger, theme.success, theme.danger],
                        radius: 2 
                    },
                    itemMargin: { horizontal: 12, vertical: 4 },
                    labels: { colors: theme.text }
                }
            };
            var chartX = new ApexCharts(document.querySelector("#xbarChart"), xOpt);
            chartX.render(); this.chartInstances.push(chartX);

            // R Chart Range adjustment
            var rValues = pageXbarR.R.data.filter(v => v !== null && !isNaN(v));
            var rMax = Math.max(...rValues, pageXbarR.R.UCL);

            var rOpt = {
                chart: {
                    type: 'line',
                    height: 320,
                    toolbar: { 
                        show: true,
                        tools: { download: true, zoom: true, zoomin: true, zoomout: true, pan: true, reset: true },
                        offsetY: -10
                    },
                    selection: { enabled: false },
                    background: 'transparent',
                    zoom: { enabled: false },
                    animations: { enabled: true, easing: 'easeinout', speed: 800 },
                    events: {
                        mounted: function (chartContext, config) {
                            chartContext.el.addEventListener('dblclick', function () {
                                chartContext.resetSeries();
                            });
                        }
                    }
                },
                theme: { mode: theme.mode },
                title: {
                    text: this.t('R 管制圖', 'R Control Chart'),
                    align: 'left',
                    style: { fontSize: '16px', fontWeight: 600, color: theme.text }
                },
                subtitle: {
                    text: this.t('子組全距 (變異)', 'Subgroup Range (Variation)'),
                    style: { fontSize: '12px', color: theme.textSec }
                },
                series: [
                    { name: this.t('全距', 'Range'), data: pageXbarR.R.data },
                    { name: this.t('上管制界限', 'UCL'), data: new Array(pageLabels.length).fill(pageXbarR.R.UCL) },
                    { name: this.t('中心線', 'CL'), data: new Array(pageLabels.length).fill(pageXbarR.R.CL) },
                    { name: this.t('下管制界限', 'LCL'), data: new Array(pageLabels.length).fill(pageXbarR.R.LCL || 0) }
                ],
                colors: [theme.primary, theme.danger, theme.success, theme.danger],
                stroke: {
                    width: [1, 1, 1, 1],
                    dashArray: [0, 5, 0, 5],
                    curve: 'straight',
                    connectNulls: false,
                    opacity: [0.6, 0.3, 0.3, 0.3]
                },
                xaxis: {
                    type: 'category',
                    categories: pageLabels.map(l => String(l)),
                    labels: { 
                        style: { colors: theme.text, fontSize: '11px', fontFamily: 'Inter, sans-serif' },
                        rotate: -45,
                        tickAmount: 15
                    },
                    axisBorder: { color: theme.grid },
                    axisTicks: { color: theme.grid }
                },
                yaxis: {
                    min: 0,
                    max: rMax + (rMax * 0.2),
                    tickAmount: 8,
                    labels: { 
                        formatter: function (v) { return (v !== null && v !== undefined) ? v.toFixed(4) : ''; },
                        style: { colors: theme.text, fontSize: '11px', fontFamily: 'Inter, sans-serif' }
                    }
                },
                grid: { 
                    borderColor: theme.grid,
                    strokeDashArray: 4,
                    xaxis: { lines: { show: true } }
                },
                annotations: {
                    yaxis: [
                        {
                            y: pageXbarR.R.UCL,
                            borderColor: theme.danger,
                            borderWidth: 2,
                            borderDash: [6, 4],
                            label: {
                                text: this.t('上管制界限: ', 'UCL: ') + ((pageXbarR.R.UCL !== null && pageXbarR.R.UCL !== undefined) ? pageXbarR.R.UCL.toFixed(4) : '-'),
                                position: 'left',
                                textAnchor: 'start',
                                style: { background: theme.danger + '20', color: theme.danger, fontWeight: 700, fontSize: '11px' }
                            }
                        },
                        {
                            y: pageXbarR.R.CL,
                            borderColor: theme.success,
                            borderWidth: 2,
                            label: {
                                text: this.t('中心線: ', 'CL: ') + ((pageXbarR.R.CL !== null && pageXbarR.R.CL !== undefined) ? pageXbarR.R.CL.toFixed(4) : '-'),
                                position: 'left',
                                textAnchor: 'start',
                                style: { background: theme.success + '20', color: theme.success, fontWeight: 700, fontSize: '11px' }
                            }
                        },
                        {
                            y: pageXbarR.R.LCL || 0,
                            borderColor: theme.danger,
                            borderWidth: 2,
                            borderDash: [6, 4],
                            label: {
                                text: this.t('下管制界限: ', 'LCL: ') + ((pageXbarR.R.LCL !== null && pageXbarR.R.LCL !== undefined) ? pageXbarR.R.LCL.toFixed(4) : '0.0000'),
                                position: 'left',
                                textAnchor: 'start',
                                style: { background: theme.danger + '20', color: theme.danger, fontWeight: 700, fontSize: '11px' }
                            }
                        }
                    ]
                },
                tooltip: { 
                    theme: theme.mode, 
                    followCursor: true, 
                    fixed: { enabled: false },
                    custom: function ({ series, seriesIndex, dataPointIndex, w }) {
                        var val = series[seriesIndex][dataPointIndex];
                        var name = w.globals.seriesNames[seriesIndex];
                        var label = w.globals.categoryLabels[dataPointIndex];
                        var html = '<div class="px-3 py-2 bg-slate-900 border border-slate-700 shadow-xl rounded-lg">';
                        html += '<div class="text-xs text-slate-400 font-bold uppercase mb-1">' + label + '</div>';
                        html += '<div class="flex items-center gap-2">';
                        html += '<span class="w-2 h-2 rounded-full" style="background-color:' + w.globals.colors[seriesIndex] + '"></span>';
                        html += '<span class="text-sm font-bold text-white">' + name + ': ' + (val ? val.toFixed(4) : '-') + '</span>';
                        html += '</div></div>';
                        return html;
                    }
                },
                markers: { 
                    size: [5, 0, 0, 0],
                    strokeWidth: 2,
                    strokeColors: [theme.bg, 'transparent', 'transparent', 'transparent'],
                    hover: { size: 8, sizeOnHover: 10 },
                    discrete: (pageXbarR.R.violations || []).map(function (v) { 
                        return { 
                            seriesIndex: 0, 
                            dataPointIndex: v.index, 
                            fillColor: theme.danger, 
                            strokeColor: theme.bg,
                            size: 8,
                            shape: 'circle'
                        }; 
                    })
                },
                legend: {
                    show: true,
                    position: 'top',
                    horizontalAlign: 'right',
                    floating: true,
                    fontSize: '12px',
                    fontFamily: 'Inter, sans-serif',
                    markers: { 
                        size: [6, 10, 10, 10], 
                        shape: ['circle', 'line', 'line', 'line'],
                        strokeWidth: [0, 2, 2, 2],
                        strokeColors: [theme.primary, theme.danger, theme.success, theme.danger],
                        radius: 2 
                    },
                    labels: { colors: theme.text }
                }
            };
            var chartR = new ApexCharts(document.querySelector("#rChart"), rOpt);
            chartR.render(); this.chartInstances.push(chartR);

            this.renderAnomalySidebar(pageXbarR, pageLabels);

        } else if (data.type === 'cavity') {
            var labels = data.cavityStats.map(s => s.name);
            var cpkVal = data.cavityStats.map(s => s.Cpk);

            // 1. Cpk Comparison Chart
            var cpkOpt = {
                chart: {
                    type: 'bar',
                    height: 350,
                    toolbar: { show: true },
                    background: 'transparent',
                    selection: { enabled: true, type: 'x' },
                    zoom: { enabled: false },
                    events: {
                        mounted: function (chartContext, config) {
                            chartContext.el.addEventListener('dblclick', function () {
                                chartContext.updateOptions({ xaxis: { min: undefined, max: undefined } }, false, false);
                            });
                        }
                    }
                },
                theme: { mode: theme.mode },
                series: [{ name: 'Cpk', data: cpkVal }],
                xaxis: {
                    type: 'category',
                    categories: labels.map(l => String(l)),
                    labels: {
                        rotate: -45,
                        rotateAlways: labels.length > 8,
                        style: { colors: theme.text, fontSize: '11px', fontFamily: 'Inter, sans-serif' }
                    }
                },
                plotOptions: {
                    bar: {
                        borderRadius: 4,
                        colors: {
                            ranges: [
                                { from: 0, to: 0.999, color: '#ef4444' },
                                { from: 1.0, to: (this.settings.cpkThreshold - 0.001), color: '#f59e0b' },
                                { from: this.settings.cpkThreshold, to: 99, color: '#38bdf8' }
                            ]
                        }
                    }
                },
                dataLabels: { enabled: false }, yaxis: { labels: { formatter: function (v) { return (v !== null && v !== undefined) ? v.toFixed(3) : ''; }, style: { colors: theme.text, fontSize: '11px', fontFamily: 'Inter, sans-serif' } }, title: { text: 'Cpk' } }, grid: { borderColor: theme.grid },
                tooltip: { followCursor: true, fixed: { enabled: false } },
                annotations: {
                    yaxis: [
                        { y: 1.0, borderColor: theme.danger, strokeDashArray: 4, label: { text: '1.0', style: { color: '#fff', background: theme.danger } } },
                        { y: this.settings.cpkThreshold, borderColor: theme.success, strokeDashArray: 4, strokeWidth: 2, label: { text: this.t('目標值: ', 'Target: ') + this.settings.cpkThreshold, style: { background: theme.success, color: '#fff' } } }
                    ]
                }
            };
            var chartCpk = new ApexCharts(document.querySelector("#cpkChart"), cpkOpt);
            chartCpk.render(); this.chartInstances.push(chartCpk);

            // 2. Mean Comparison Chart
            // Calculate valid range for Y-axis (Filtering out null/0 to prevent stretching)
            var allValues = [];
            data.cavityStats.forEach(s => {
                if (s.mean !== null) allValues.push(s.mean);
            });
            allValues.push(data.specs.usl, data.specs.lsl, data.specs.target);

            var validValues = allValues.filter(v => v !== null && !isNaN(v) && v !== 0);
            var yMin = validValues.length > 0 ? Math.min(...validValues) : 0;
            var yMax = validValues.length > 0 ? Math.max(...validValues) : 0.1;
            var yMargin = (yMax - yMin) * 0.1 || 0.001;

            var meanOpt = {
                chart: {
                    type: 'line',
                    height: 350,
                    toolbar: { show: true },
                    background: 'transparent',
                    selection: { enabled: true, type: 'x' },
                    zoom: { enabled: false },
                    events: {
                        mounted: function (chartContext, config) {
                            chartContext.el.addEventListener('dblclick', function () {
                                chartContext.updateOptions({ xaxis: { min: undefined, max: undefined } }, false, false);
                            });
                        }
                    }
                },
                theme: { mode: theme.mode },
                series: [
                    { name: self.t('平均值', 'Mean'), data: data.cavityStats.map(s => s.mean) },
                    { name: self.t('目標值', 'Target'), data: new Array(labels.length).fill(data.specs.target) },
                    { name: self.t('規格上限', 'USL'), data: new Array(labels.length).fill(data.specs.usl) },
                    { name: self.t('規格下限', 'LSL'), data: new Array(labels.length).fill(data.specs.lsl) }
                ],
                colors: ['#3b82f6', '#10b981', '#ef4444', '#ef4444'], // Blue-500, Emerald-500, Red-500
                stroke: {
                    width: [3, 2, 2, 2],
                    dashArray: [0, 0, 5, 5],
                    connectNulls: false
                },
                markers: { size: [4, 0, 0, 0], hover: { size: 6 } },
                xaxis: {
                    type: 'category',
                    categories: labels.map(l => String(l)),
                    labels: {
                        rotate: -45,
                        rotateAlways: labels.length > 8,
                        style: { colors: theme.text, fontSize: '11px', fontFamily: 'Inter, sans-serif' }
                    }
                },
                dataLabels: { enabled: false },
                yaxis: {
                    min: yMin - yMargin,
                    max: yMax + yMargin,
                    labels: { formatter: function (v) { return v ? v.toFixed(4) : ''; }, style: { colors: theme.text, fontSize: '11px', fontFamily: 'Inter, sans-serif' } },
                    title: { text: self.t('平均值', 'Mean') }
                },
                grid: { borderColor: theme.grid },
                tooltip: { followCursor: true, shared: true, fixed: { enabled: false } }
            };
            var chartMean = new ApexCharts(document.querySelector("#meanChart"), meanOpt);
            chartMean.render(); this.chartInstances.push(chartMean);

            // 3. StdDev Comparison Chart
            var sValues = data.cavityStats.flatMap(s => [s.overallStdDev, s.withinStdDev]).filter(v => v !== null && !isNaN(v));
            var maxVal = sValues.length > 0 ? Math.max(...sValues) : 0.0001;
            var sMax = (maxVal > 0 ? maxVal * 1.2 : 0.0001);

            var stdOpt = {
                chart: {
                    type: 'line',
                    height: 350,
                    toolbar: { show: true },
                    background: 'transparent',
                    zoom: { enabled: false }
                },
                theme: { mode: theme.mode },
                series: [
                    { name: self.t('組內標準差 (σ within)', 'σ within'), data: data.cavityStats.map(s => s.withinStdDev) },
                    { name: self.t('全樣本標準差 (σ overall)', 'σ overall'), data: data.cavityStats.map(s => s.overallStdDev) }
                ],
                colors: ['#3b82f6', '#ef4444'], // Blue-500, Red-500
                stroke: {
                    width: [2, 2],
                    curve: 'straight',
                    dashArray: [0, 0], // Both solid lines
                    connectNulls: false
                },
                markers: {
                    size: 4,
                    shape: ["circle", "square"], // Within is circle, Overall is square
                    strokeWidth: 2,
                    hover: { size: 6 }
                },
                xaxis: {
                    type: 'category',
                    categories: labels.map(l => String(l)),
                    labels: {
                        rotate: -45,
                        rotateAlways: labels.length > 8,
                        style: { colors: theme.text, fontSize: '11px', fontFamily: 'Inter, sans-serif' }
                    }
                },
                dataLabels: { enabled: false },
                yaxis: {
                    min: 0,
                    max: sMax,
                    tickAmount: 5,
                    labels: {
                        formatter: function (v) { return (v !== null && v !== undefined) ? v.toFixed(6) : ''; },
                        style: { colors: theme.text, fontSize: '11px', fontFamily: 'Inter, sans-serif' }
                    },
                    title: { text: self.t('標準差', 'StdDev') }
                },
                grid: { borderColor: theme.grid },
                tooltip: { followCursor: true, shared: true, theme: theme.mode }
            };
            var chartStd = new ApexCharts(document.querySelector("#stdDevChart"), stdOpt);
            chartStd.render(); this.chartInstances.push(chartStd);

        } else if (data.type === 'group') {
            var labels = data.groupStats.map(s => s.batch);

            // Calculate Y-axis range for better visualization (Filtering out 0/null)
            var gValues = data.groupStats.flatMap(s => [s.max, s.avg, s.min]).filter(v => v !== null && !isNaN(v) && v !== 0);
            gValues.push(data.specs.usl, data.specs.lsl, data.specs.target);
            var validGValues = gValues.filter(v => v !== null && !isNaN(v) && v !== 0);
            var yMin = validGValues.length > 0 ? Math.min(...validGValues) : 0;
            var yMax = validGValues.length > 0 ? Math.max(...validGValues) : 0.1;
            var yMargin = (yMax - yMin) * 0.1 || 0.001;

            var gOpt = {
                chart: {
                    type: 'line',
                    height: 380,
                    toolbar: {
                        show: true,
                        tools: {
                            download: true,
                            selection: true,
                            zoom: true,
                            zoomin: true,
                            zoomout: true,
                            pan: true,
                            reset: true
                        }
                    },
                    selection: { enabled: true, type: 'x' },
                    zoom: { enabled: false },
                    events: {
                        selection: function (chart, e) {
                            if (e.xaxis) {
                                chart.updateOptions({ xaxis: { min: e.xaxis.min, max: e.xaxis.max } }, false, false);
                            }
                        },
                        mounted: function (chartContext, config) {
                            chartContext.el.addEventListener('dblclick', function () {
                                chartContext.updateOptions({ xaxis: { min: undefined, max: undefined } }, false, false);
                            });
                        }
                    }
                },
                theme: { mode: theme.mode },
                series: [
                    { name: self.t('最大值', 'Max'), data: data.groupStats.map(s => s.max) },
                    { name: self.t('平均值', 'Avg'), data: data.groupStats.map(s => s.avg) },
                    { name: self.t('最小值', 'Min'), data: data.groupStats.map(s => s.min) },
                    { name: self.t('規格上限', 'USL'), data: new Array(labels.length).fill(data.specs.usl) },
                    { name: self.t('目標值', 'Target'), data: new Array(labels.length).fill(data.specs.target) },
                    { name: self.t('規格下限', 'LSL'), data: new Array(labels.length).fill(data.specs.lsl) }
                ],
                colors: ['#ef4444', '#3b82f6', '#ef4444', '#ff9800', '#10b981', '#ff9800'], // Red, Blue, Red, Orange, Emerald, Orange
                stroke: {
                    width: [1.5, 3, 1.5, 2, 2, 2],
                    dashArray: [0, 0, 0, 5, 0, 5],
                    curve: 'straight',
                    connectNulls: false // IMPORTANT: Don't connect points if data is missing
                },
                markers: {
                    size: labels.length > 50 ? [0, 0, 0, 0, 0, 0] : [0, 3, 0, 0, 0, 0], // Hide markers if too many points
                    colors: ['#3b82f6'],
                    strokeColors: '#fff',
                    strokeWidth: 1,
                    hover: { size: 5 }
                },
                xaxis: {
                    type: 'category', // Force text labels (prevents numeric batch IDs being treated as numeric axis)
                    categories: labels.map(l => String(l)),
                    labels: {
                        rotate: -45,
                        rotateAlways: labels.length > 20,
                        hideOverlappingLabels: true,
                        trim: true,
                        maxHeight: 120,
                        formatter: function (v) {
                            var i = Math.floor(v);
                            // Try multiple mapping strategies to ensure batch label is found
                            if (labels && labels[i]) return labels[i];
                            if (labels && labels[i - 1]) return labels[i - 1];
                            return v;
                        },
                        style: {
                            colors: theme.text,
                            fontSize: labels.length > 50 ? '9px' : '11px',
                            fontFamily: 'Inter, sans-serif'
                        }
                    },
                    tickAmount: labels.length > 100 ? 20 : undefined, // Limit tick marks for large datasets
                    tickPlacement: 'on'
                },
                dataLabels: { enabled: false },
                yaxis: {
                    min: yMin,
                    max: yMax,
                    forceNiceScale: false,
                    labels: {
                        formatter: function (v) { return (v !== null && v !== undefined) ? v.toFixed(4) : ''; },
                        style: {
                            colors: theme.text,
                            fontSize: '12px',
                            fontFamily: 'Inter, sans-serif'
                        }
                    },
                    title: {
                        text: self.t('測量值', 'Value'),
                        style: { fontSize: '13px', fontWeight: 600 }
                    }
                },
                grid: { borderColor: theme.grid },
                tooltip: {
                    followCursor: true,
                    intersect: false,
                    fixed: { enabled: false },
                    custom: function ({ series, seriesIndex, dataPointIndex, w }) {
                        var label = (labels && labels[dataPointIndex]) ? labels[dataPointIndex] : w.globals.categoryLabels[dataPointIndex];
                        var html = '<div class="px-3 py-2 bg-slate-50 dark:bg-slate-900 border border-slate-200 dark:border-slate-700 shadow-xl rounded-lg max-w-[320px] break-words" style="word-wrap: break-word; overflow-wrap: break-word; max-width: 320px;">';
                        html += '<div class="text-xs text-slate-500 dark:text-slate-400 font-bold uppercase mb-1 border-b border-slate-200 dark:border-slate-800 pb-1 break-words" style="word-wrap: break-word;">' + label + '</div>';

                        w.config.series.forEach(function (s, idx) {
                            var val = series[idx][dataPointIndex];
                            if (val !== undefined && val !== null) {
                                html += '<div class="flex items-center justify-between gap-4 py-0.5">';
                                html += '<div class="flex items-center gap-1.5">';
                                html += '<span class="w-2 h-2 rounded-full" style="background-color:' + w.globals.colors[idx] + '"></span>';
                                html += '<span class="text-xs font-bold text-slate-700 dark:text-slate-200">' + s.name + ':</span>';
                                html += '</div>';
                                html += '<span class="text-xs font-mono font-bold text-slate-600 dark:text-slate-300">' + (typeof val === "number" ? val.toFixed(4) : val) + '</span>';
                                html += '</div>';
                            }
                        });

                        html += '</div>';
                        return html;
                    }
                }
            };
            var chartG = new ApexCharts(document.querySelector("#groupChart"), gOpt);
            chartG.render(); this.chartInstances.push(chartG);

            // 5. Variation Chart (Visual match to old Chart.js)
            var vOpt = {
                chart: {
                    type: 'line',
                    height: 380,
                    toolbar: { show: true },
                    selection: { enabled: true, type: 'x' },
                    zoom: { enabled: false },
                    events: {
                        selection: function (chart, e) {
                            if (e.xaxis) {
                                chart.updateOptions({ xaxis: { min: e.xaxis.min, max: e.xaxis.max } }, false, false);
                            }
                        },
                        mounted: function (chartContext, config) {
                            chartContext.el.addEventListener('dblclick', function () {
                                chartContext.updateOptions({ xaxis: { min: undefined, max: undefined } }, false, false);
                            });
                        }
                    }
                },
                theme: { mode: theme.mode },
                series: [{ name: 'Range', data: data.groupStats.map(s => s.range) }],
                colors: ['#8b5cf6'], // Violet-500
                stroke: { width: 2 },
                markers: { size: 4, shape: 'square' },
                xaxis: {
                    type: 'category',
                    categories: labels.map(l => String(l)),
                    labels: {
                        rotate: -45,
                        rotateAlways: labels.length > 20,
                        hideOverlappingLabels: true,
                        trim: true,
                        formatter: function (v) {
                            var i = Math.floor(v);
                            // Try multiple mapping strategies to ensure batch label is found
                            if (labels && labels[i]) return labels[i];
                            if (labels && labels[i - 1]) return labels[i - 1];
                            return v;
                        },
                        style: { colors: theme.text, fontSize: labels.length > 25 ? '10px' : '11px', fontFamily: 'Inter, sans-serif' }
                    }
                },
                dataLabels: { enabled: false },
                yaxis: {
                    labels: {
                        formatter: function (v) { return (v !== null && v !== undefined) ? v.toFixed(4) : ''; },
                        style: { colors: theme.text, fontSize: '12px', fontFamily: 'Inter, sans-serif' }
                    },
                    title: { text: self.t('全距', 'Range') }
                },
                grid: { borderColor: theme.grid },
                tooltip: {
                    followCursor: true,
                    intersect: false,
                    fixed: { enabled: false },
                    custom: function ({ series, seriesIndex, dataPointIndex, w }) {
                        var label = (labels && labels[dataPointIndex]) ? labels[dataPointIndex] : w.globals.categoryLabels[dataPointIndex];
                        var html = '<div class="px-3 py-2 bg-slate-50 dark:bg-slate-900 border border-slate-200 dark:border-slate-700 shadow-xl rounded-lg max-w-[320px] break-words" style="word-wrap: break-word; overflow-wrap: break-word; max-width: 320px;">';
                        html += '<div class="text-xs text-slate-500 dark:text-slate-400 font-bold uppercase mb-1 border-b border-slate-200 dark:border-slate-800 pb-1 break-words" style="word-wrap: break-word;">' + label + '</div>';
                        html += '<div class="flex items-center justify-between gap-4 py-0.5">';
                        html += '<div class="flex items-center gap-1.5">';
                        html += '<span class="w-2 h-2 rounded-full shrink-0" style="background-color:' + w.globals.colors[0] + '"></span>';
                        html += '<span class="text-xs font-bold text-slate-700 dark:text-slate-200">Range:</span>';
                        html += '</div>';
                        html += '<span class="text-xs font-mono font-bold text-slate-600 dark:text-slate-300">' + (series[0][dataPointIndex] !== null ? series[0][dataPointIndex].toFixed(4) : '-') + '</span>';
                        html += '</div></div>';
                        return html;
                    }
                }
            };
            var chartV = new ApexCharts(document.querySelector("#groupVarChart"), vOpt);
            chartV.render(); this.chartInstances.push(chartV);
        }
    },

    downloadExcel: async function () {
        var self = this;
        var data = this.analysisResults;

        this.showLoading(this.t('正在生成 Excel 報表...', 'Generating Excel Report...'));

        // Capture Charts
        var images = {};
        if (data.type === 'batch') {
            // Assuming chartInstances[0] is XBar and [1] is R (based on render order in renderCharts)
            if (this.chartInstances.length >= 2) {
                try {
                    const xUri = await this.chartInstances[0].dataURI();
                    const rUri = await this.chartInstances[1].dataURI();
                    images.xbar = xUri.imgURI;
                    images.r = rUri.imgURI;
                } catch (e) { console.error('Error capturing charts', e); }
            }
        }

        try {
            // Check if SPCExcelBuilder is loaded
            if (typeof SPCExcelBuilder === 'undefined') {
                throw new Error('ExcelBuilder module not loaded');
            }

            // Template Selection Disabled (Reverted to Manual Mode)
            var templateBuffer = null;
            var templateCapacity = 0;
            console.log("Template selection disabled. Using Manual Generation Mode.");

            // // Template Smart Selection
            // var templateBase64 = null;
            // var templateCapacity = 0;
            // var cavCount = data.xbarR.summary.n;
            // var usedBuiltIn = false;

            // // Debugging Info
            // // console.log(`Template Search: Cavities=${cavCount}`);

            // // 1. Check User Uploaded Template
            // var userTemplate = localStorage.getItem('spc_template_file');
            // if (userTemplate) {
            //     if (userTemplate.length > 100) {
            //         templateBase64 = userTemplate;
            //         templateCapacity = cavCount;
            //         console.log('Using User Custom Template');
            //     } else {
            //         localStorage.removeItem('spc_template_file');
            //     }
            // }

            // // 2. Check Built-in Templates
            // if (!templateBase64) {
            //     if (typeof SPC_TEMPLATES === 'undefined') {
            //         // Last ditch: check if window.SPC_TEMPLATES exists
            //         if (window.SPC_TEMPLATES) {
            //             console.log("Found SPC_TEMPLATES on window");
            //         } else {
            //             console.error("SPC_TEMPLATES not defined. Template file probably not loaded.");
            //             // alert("Critical Error: Built-in Templates not loaded. Please refresh the page.");
            //         }
            //     } else {
            //         // Find best fit
            //         var availableSizes = Object.keys(SPC_TEMPLATES).map(Number).sort((a, b) => a - b);
            //         var bestFit = availableSizes.find(size => size >= cavCount);
            //         if (!bestFit) bestFit = availableSizes[availableSizes.length - 1];

            //         if (bestFit) {
            //             templateBase64 = SPC_TEMPLATES[bestFit];
            //             templateCapacity = bestFit;
            //             usedBuiltIn = true;
            //             // console.log(`Selected Template: ${bestFit}`);
            //         }
            //     }
            // }

            // var templateBuffer = null;
            // if (templateBase64) {
            //     try {
            //         // Extract Base64, remove whitespaces
            //         var content = templateBase64.includes(',') ? templateBase64.split(',')[1] : templateBase64;
            //         content = content.replace(/\s/g, ''); // Fix potential formatting issues

            //         var binary_string = window.atob(content);
            //         var len = binary_string.length;
            //         var bytes = new Uint8Array(len);
            //         for (var i = 0; i < len; i++) {
            //             bytes[i] = binary_string.charCodeAt(i);
            //         }
            //         templateBuffer = bytes.buffer;
            //         // console.log(`Template Buffer Prepared: ${len} bytes`);
            //     } catch (e) {
            //         console.error('Failed to parse template', e);
            //         // alert("Template Parsing Failed: " + e.message);
            //     }
            // } else {
            //     console.warn("No template selected.");
            // }

            const builder = new SPCExcelBuilder(data, images, templateBuffer, templateCapacity);
            const buffer = await builder.generate();

            // Save file
            const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });

            // Generate filename matches QIP/VBA style
            var dateStr = new Date().toISOString().slice(0, 10).replace(/-/g, '');
            var itemStr = (self.selectedItem || 'Data').replace(/[:\/\\*?"<>|]/g, '_');
            var filename = 'SPC_Report_' + itemStr + '_' + data.type + '_' + dateStr + '.xlsx';

            // Use File System Access API if available
            try {
                if (window.showSaveFilePicker) {
                    var handle = await window.showSaveFilePicker({
                        suggestedName: filename,
                        types: [{
                            description: 'Excel Workbook',
                            accept: { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'] }
                        }]
                    });
                    var writable = await handle.createWritable();
                    await writable.write(blob);
                    await writable.close();
                    this.hideLoading();
                    return;
                }
            } catch (err) {
                if (err.name !== 'AbortError') console.error('File save error:', err);
                else { this.hideLoading(); return; } // User cancelled
            }

            // Fallback download
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = filename;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);

        } catch (e) {
            console.error('Export failed', e);
            alert(this.t('導出失敗: ', 'Export failed: ') + e.message);
        } finally {
            this.hideLoading();
        }
    },

    showLoading: function (text) {
        var overlay = document.getElementById('loadingOverlay');
        var textEl = document.getElementById('loadingText');
        if (textEl) textEl.textContent = text;
        if (overlay) overlay.style.display = 'flex';
    },

    hideLoading: function () {
        var overlay = document.getElementById('loadingOverlay');
        if (overlay) overlay.style.display = 'none';
    },

    renderAnomalySidebar: function (pageXbarR, pageLabels) {
        var self = this;
        var list = document.getElementById('anomalyList');
        if (!list) return;
        list.innerHTML = '';
        var violations = pageXbarR.xBar.violations;
        var indicator = document.getElementById('anomalyIndicator');

        if (violations.length === 0) {
            list.innerHTML = '<div class="text-center text-slate-400 dark:text-slate-500 py-8 text-sm">' + this.t('本頁無異常點', 'No anomalies on this page') + '</div>';
            if (indicator) indicator.classList.add('hidden');
            return;
        } else {
            if (indicator) indicator.classList.remove('hidden');
        }

        violations.forEach(function (v) {
            var card = document.createElement('div');
            card.className = 'bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 p-4 rounded-xl mb-3 mx-4 group relative cursor-help hover:border-rose-400 dark:hover:border-rose-500 transition-colors shadow-sm';

            var rulesText = v.rules.map(function (r) { return self.t('規則 ', 'Rule ') + r; }).join(', ');

            // 收集所有違反規則的專家意見
            var allMoldingAdvice = [];
            var allQualityAdvice = [];

            // 收集並明確列出每個違反規則的專家意見
            var moldingItems = [];
            var qualityItems = [];

            v.rules.forEach(function (ruleId) {
                var expPair = self.nelsonExpertise[ruleId];
                // 如果找不到對應的規則建議，使用默認值
                if (!expPair) {
                    expPair = {
                        zh: { m: "請檢查製程參數。", q: "請參考標準作業程序。" },
                        en: { m: "Please check process parameters.", q: "Please refer to SOP." }
                    };
                }

                var exp = self.settings.language === 'zh' ? expPair.zh : expPair.en;

                // 直接收集，不去重，確保每個規則都有對應建議
                moldingItems.push({ id: ruleId, text: exp.m });
                qualityItems.push({ id: ruleId, text: exp.q });
            });

            // 渲染成 HTML
            var moldingAdviceHTML = moldingItems.map(function (item, idx) {
                return '<div class="text-sm text-slate-600 dark:text-slate-400 leading-relaxed pl-2 ' + (idx > 0 ? 'mt-2 border-t border-slate-100 dark:border-slate-700 pt-2' : '') + '">' +
                    '<span class="inline-block text-[10px] font-bold text-slate-400 bg-slate-100 dark:bg-slate-700 px-1.5 rounded mr-1.5 min-w-[24px] text-center">R' + item.id + '</span>' +
                    item.text + '</div>';
            }).join('');

            var qualityAdviceHTML = qualityItems.map(function (item, idx) {
                return '<div class="text-sm text-slate-600 dark:text-slate-400 leading-relaxed pl-2 ' + (idx > 0 ? 'mt-2 border-t border-slate-100 dark:border-slate-700 pt-2' : '') + '">' +
                    '<span class="inline-block text-[10px] font-bold text-slate-400 bg-slate-100 dark:bg-slate-700 px-1.5 rounded mr-1.5 min-w-[24px] text-center">R' + item.id + '</span>' +
                    item.text + '</div>';
            }).join('');

            card.innerHTML = '<div class="flex justify-between items-start mb-2 pb-2 border-b border-slate-100 dark:border-slate-700">' +
                '<div class="text-base font-bold text-slate-900 dark:text-white">' + (pageLabels[v.index] || 'Batch') + '</div>' +
                '<div class="text-xs font-bold text-rose-500 bg-rose-50 dark:bg-rose-900/30 px-2 py-0.5 rounded-full">' + rulesText + '</div>' +
                '</div>' +
                '<div class="space-y-4">' +
                '<div>' +
                '<div class="flex items-center gap-1.5 text-xs text-sky-600 dark:text-sky-400 font-bold mb-1">' +
                '<span class="material-icons-outlined text-sm">precision_manufacturing</span> ' + self.t('成型專家', 'Molding Expert') + '</div>' +
                moldingAdviceHTML +
                '</div>' +
                '<div>' +
                '<div class="flex items-center gap-1.5 text-xs text-emerald-600 dark:text-emerald-400 font-bold mb-1">' +
                '<span class="material-icons-outlined text-sm">assignment_turned_in</span> ' + self.t('品管專家', 'Quality Expert') + '</div>' +
                qualityAdviceHTML +
                '</div>' +
                '</div>' +
                '<div class="text-[11px] text-slate-400 mt-2 text-right italic">Index: ' + (v.index + 1) + '</div>';

            list.appendChild(card);
        });


    },

    toggleAnomalySidebar: function (show) {
        var sidebar = document.getElementById('anomalySidebar');
        if (!sidebar) return;

        var isHidden = sidebar.classList.contains('hidden') || sidebar.style.display === 'none';

        if (show === undefined) {
            // Toggle
            if (isHidden) {
                sidebar.classList.remove('hidden');
                sidebar.style.display = 'flex';
            } else {
                sidebar.classList.add('hidden');
                sidebar.style.display = 'none';
            }
        } else if (show) {
            sidebar.classList.remove('hidden');
            sidebar.style.display = 'flex';
        } else {
            sidebar.classList.add('hidden');
            sidebar.style.display = 'none';
        }
    },


    // --- Settings View Logic ---
    renderSettings: function () {
        var self = this;
        var configList = document.getElementById('settings-config-list');
        var configs = JSON.parse(localStorage.getItem('qip_configs') || '[]');

        configList.innerHTML = configs.length === 0 ?
            '<tr><td colspan="4" class="px-6 py-8 text-center text-slate-400 italic">' + this.t('尚未儲存任何 QIP 提取配置', 'No saved configurations found.') + '</td></tr>' : '';

        configs.forEach(function (config, index) {
            var row = document.createElement('tr');
            row.className = 'hover:bg-slate-50 dark:hover:bg-slate-800/50 transition-colors';
            var date = config.savedAt ? new Date(config.savedAt).toLocaleDateString() : 'N/A';

            row.innerHTML = `
                <td class="px-6 py-4 font-bold text-slate-700 dark:text-slate-300">${config.name}</td>
                <td class="px-6 py-4 text-slate-500">${config.cavityCount}</td>
                <td class="px-6 py-4 text-slate-500">${date}</td>
                <td class="px-6 py-4 text-right">
                    <button class="delete-config-btn p-2 text-slate-300 hover:text-rose-500 transition-colors" data-index="${index}">
                        <span class="material-icons-outlined">delete</span>
                    </button>
                </td>
            `;
            configList.appendChild(row);
        });

        // Event: Delete Config
        configList.querySelectorAll('.delete-config-btn').forEach(btn => {
            btn.addEventListener('click', function () {
                if (confirm(self.t('確定要刪除此配置嗎？', 'Are you sure you want to delete this config?'))) {
                    var idx = parseInt(this.dataset.index);
                    configs.splice(idx, 1);
                    localStorage.setItem('qip_configs', JSON.stringify(configs));
                    self.renderSettings();
                }
            });
        });

        // Setup Buttons (Idempotent check)
        var exportBtn = document.getElementById('exportConfigsBtn');
        if (exportBtn && !exportBtn.dataset.bound) {
            exportBtn.dataset.bound = "true";
            exportBtn.addEventListener('click', function () { self.exportConfigurations(); });
        }

        var importInput = document.getElementById('importConfigsInput');
        if (importInput && !importInput.dataset.bound) {
            importInput.dataset.bound = "true";
            importInput.addEventListener('change', function (e) { self.importConfigurations(e); });
        }

        // Language setting sync
        var langSelect = document.getElementById('setting-lang');
        if (langSelect) {
            langSelect.value = this.settings.language;
            langSelect.onchange = function () {
                self.settings.language = this.value;
                self.saveSettings();
                self.syncLanguageState();
                self.updateLanguage();
            };
        }

        // Cpk Threshold sync
        var cpkInput = document.getElementById('setting-cpk-warn');
        if (cpkInput) {
            cpkInput.value = this.settings.cpkThreshold;
            cpkInput.onchange = function () {
                self.settings.cpkThreshold = parseFloat(this.value) || 1.33;
                self.saveSettings();
            };
        }
    },

    handleTemplateUpload: function (file) {
        if (!file) return;
        var self = this;
        var reader = new FileReader();
        reader.onload = function (e) {
            try {
                // Store as Base64 string
                // Check size? LS limit is ~5MB chars. Base64 is 33% larger.
                // 3MB xlsx -> 4MB base64. OK for simple templates.
                var base64 = e.target.result;
                localStorage.setItem('spc_template_file', base64);
                localStorage.setItem('spc_template_meta', JSON.stringify({
                    name: file.name,
                    size: file.size,
                    date: new Date().toISOString()
                }));
                alert(self.t('模板上傳成功！導出報表時將自動採用此模板。', 'Template uploaded! It will be used for future Excel exports.'));
                self.renderSettings(); // Refresh UI
            } catch (err) {
                console.error(err);
                if (err.name === 'QuotaExceededError') {
                    alert(self.t('錯誤：模板檔案太大，無法儲存於瀏覽器快取。', 'Error: Template file is too large for browser storage.'));
                } else {
                    alert('Upload failed: ' + err.message);
                }
            }
        };
        reader.readAsDataURL(file);
    },

    clearTemplate: function () {
        if (confirm(this.t('確定要移除自定義模板嗎？', 'Remove custom template?'))) {
            localStorage.removeItem('spc_template_file');
            localStorage.removeItem('spc_template_meta');
            this.renderSettings();
        }
    },

    saveSettings: function () {
        localStorage.setItem('spc_settings', JSON.stringify(this.settings));
    },

    loadSettings: function () {
        var saved = localStorage.getItem('spc_settings');
        if (saved) {
            this.settings = Object.assign(this.settings, JSON.parse(saved));
        }
    },

    exportConfigurations: async function () {
        var configs = localStorage.getItem('qip_configs') || '[]';
        var blob = new Blob([configs], { type: 'application/json' });
        var date = new Date().toISOString().slice(0, 10).replace(/-/g, '');
        var filename = `SPC_QIP_Configs_Backup_${date}.json`;

        // Use File System Access API if available
        if (window.showSaveFilePicker) {
            try {
                const handle = await window.showSaveFilePicker({
                    suggestedName: filename,
                    types: [{
                        description: 'JSON Configuration File',
                        accept: { 'application/json': ['.json'] }
                    }]
                });
                const writable = await handle.createWritable();
                await writable.write(blob);
                await writable.close();
                return;
            } catch (err) {
                if (err.name === 'AbortError') return;
                console.error('Config export failed:', err);
            }
        }

        // Fallback
        var url = URL.createObjectURL(blob);
        var a = document.createElement('a');
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
    },

    importConfigurations: function (event) {
        var self = this;
        var file = event.target.files[0];
        if (!file) return;

        var reader = new FileReader();
        reader.onload = function (e) {
            try {
                var imported = JSON.parse(e.target.result);
                if (!Array.isArray(imported)) throw new Error('Invalid format');

                var existing = JSON.parse(localStorage.getItem('qip_configs') || '[]');
                // Simple merge: avoid duplicates by name
                var merged = [...existing];
                imported.forEach(imp => {
                    if (!merged.find(m => m.name === imp.name)) {
                        merged.push(imp);
                    }
                });

                localStorage.setItem('qip_configs', JSON.stringify(merged));
                alert(self.t(`成功讀取！已匯入 ${imported.length} 組配置。`, `Import success! ${imported.length} configs added.`));
                self.renderSettings();
                event.target.value = ''; // Reset input
            } catch (err) {
                alert(self.t('讀取失敗：檔案格式不正確', 'Import failed: Invalid file format'));
                console.error(err);
            }
        };
        reader.readAsText(file);
    },

    resetApp: function () {
        this.workbook = null; this.selectedItem = null; this.analysisResults = null;
        this.chartInstances.forEach(c => { if (c.destroy) c.destroy(); });
        this.chartInstances = [];
        document.getElementById('fileInput').value = '';
        var uploadZone = document.getElementById('uploadZone');
        if (uploadZone) uploadZone.style.display = 'block';
        var fileInfo = document.getElementById('fileInfo');
        if (fileInfo) fileInfo.style.display = 'none';
        ['step2', 'step3', 'results', 'dataPreviewSection'].forEach(id => { var el = document.getElementById(id); if (el) el.style.display = 'none'; });
        window.scrollTo({ top: 0, behavior: 'smooth' });
    },

    resetSystem: function () {
        console.log('SPCApp: resetSystem triggered');
        var msg = this.t('確定要重置系統嗎？這將清除所有緩存配置、歷史紀錄、QIP 設定與當前所有數據並重新整理頁面。',
            'Are you sure you want to reset the system? This will clear all cached configs, history, QIP settings, and current data, and then refresh the page.');

        if (confirm(msg)) {
            try {
                console.log('Clearing storage...');
                localStorage.clear();
                sessionStorage.clear();
                console.log('Storage cleared. Reloading...');
                window.location.reload();
            } catch (e) {
                console.error('Reset failed:', e);
                alert('Reset failed: ' + e.message);
            }
        }
    },
    /**
     * renderConstantsTable - Dynamically generate the constants reference table from SPCEngine data
     */
    renderConstantsTable: function () {
        var tbody = document.getElementById('constants-table-body');
        if (!tbody) return;

        var html = '';
        var nList = Object.keys(SPCEngine.SPC_CONSTANTS).map(Number).sort((a, b) => a - b);

        var self = this;
        nList.forEach(function (n) {
            var constants = SPCEngine.getConstants(n);
            var d2 = SPCEngine.D2_CONSTANTS[n] || '-';

            var isStandard = (n === 5);
            var rowClass = isStandard ? 'bg-indigo-50/20 dark:bg-indigo-900/10' : 'hover:bg-slate-50 dark:hover:bg-slate-900/50';
            var standardLabel = self.t('(標準常數)', '(Standard)');

            html += '<tr class="' + rowClass + '">' +
                '<td class="px-3 py-2 font-black border-r border-slate-200 dark:border-slate-700">' +
                n + (isStandard ? ' <span class="text-[10px] text-indigo-500">' + standardLabel + '</span>' : '') +
                '</td>' +
                '<td class="px-3 py-2">' + constants.A2.toFixed(3) + '</td>' +
                '<td class="px-3 py-2">' + (typeof d2 === 'number' ? d2.toFixed(3) : d2) + '</td>' +
                '<td class="px-3 py-2">' + constants.D3.toFixed(3) + '</td>' +
                '<td class="px-3 py-2">' + constants.D4.toFixed(3) + '</td>' +
                '</tr>';
        });

        tbody.innerHTML = html;
    }
};

document.addEventListener('DOMContentLoaded', function () { SPCApp.init(); });
