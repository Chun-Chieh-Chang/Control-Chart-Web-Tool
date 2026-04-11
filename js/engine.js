/**
 * SPC ENGINE - Core Statistical Calculations
 * Based on industry standard formulas for X-Bar and R charts.
 */

var SPCEngine = {
    SPC_CONSTANTS: {
        2: { A2: 1.88, D3: 0, D4: 3.267 },
        3: { A2: 1.023, D3: 0, D4: 2.574 },
        4: { A2: 0.729, D3: 0, D4: 2.282 },
        5: { A2: 0.577, D3: 0, D4: 2.114 },
        6: { A2: 0.483, D3: 0, D4: 2.004 },
        7: { A2: 0.419, D3: 0.076, D4: 1.924 },
        8: { A2: 0.373, D3: 0.136, D4: 1.864 },
        9: { A2: 0.337, D3: 0.184, D4: 1.816 },
        10: { A2: 0.308, D3: 0.223, D4: 1.777 },
        11: { A2: 0.285, D3: 0.256, D4: 1.744 },
        12: { A2: 0.266, D3: 0.283, D4: 1.717 },
        13: { A2: 0.249, D3: 0.307, D4: 1.693 },
        14: { A2: 0.235, D3: 0.328, D4: 1.672 },
        15: { A2: 0.223, D3: 0.347, D4: 1.653 },
        16: { A2: 0.212, D3: 0.363, D4: 1.637 },
        17: { A2: 0.203, D3: 0.378, D4: 1.622 },
        18: { A2: 0.194, D3: 0.391, D4: 1.608 },
        19: { A2: 0.187, D3: 0.403, D4: 1.597 },
        20: { A2: 0.180, D3: 0.415, D4: 1.585 },
        21: { A2: 0.173, D3: 0.425, D4: 1.575 },
        22: { A2: 0.167, D3: 0.434, D4: 1.566 },
        23: { A2: 0.162, D3: 0.443, D4: 1.557 },
        24: { A2: 0.157, D3: 0.451, D4: 1.548 },
        25: { A2: 0.153, D3: 0.459, D4: 1.541 },
        26: { A2: 0.149, D3: 0.466, D4: 1.534 },
        27: { A2: 0.146, D3: 0.473, D4: 1.527 },
        28: { A2: 0.143, D3: 0.479, D4: 1.521 },
        29: { A2: 0.141, D3: 0.485, D4: 1.515 },
        30: { A2: 0.138, D3: 0.491, D4: 1.509 },
        31: { A2: 0.136, D3: 0.497, D4: 1.503 },
        32: { A2: 0.134, D3: 0.502, D4: 1.498 },
        33: { A2: 0.132, D3: 0.507, D4: 1.493 },
        34: { A2: 0.130, D3: 0.512, D4: 1.488 },
        35: { A2: 0.128, D3: 0.517, D4: 1.483 },
        36: { A2: 0.126, D3: 0.521, D4: 1.479 },
        37: { A2: 0.125, D3: 0.525, D4: 1.475 },
        38: { A2: 0.123, D3: 0.529, D4: 1.471 },
        39: { A2: 0.122, D3: 0.533, D4: 1.467 },
        40: { A2: 0.120, D3: 0.537, D4: 1.463 },
        41: { A2: 0.118, D3: 0.540, D4: 1.460 },
        42: { A2: 0.117, D3: 0.544, D4: 1.456 },
        43: { A2: 0.115, D3: 0.547, D4: 1.453 },
        44: { A2: 0.114, D3: 0.551, D4: 1.449 },
        45: { A2: 0.113, D3: 0.554, D4: 1.446 },
        46: { A2: 0.111, D3: 0.557, D4: 1.443 },
        47: { A2: 0.110, D3: 0.560, D4: 1.440 },
        48: { A2: 0.109, D3: 0.563, D4: 1.437 }
    },

    // d2 constants for estimating sigma from R-bar (extended to 48 cavities)
    D2_CONSTANTS: {
        2: 1.128, 3: 1.693, 4: 2.059, 5: 2.326, 6: 2.534,
        7: 2.704, 8: 2.847, 9: 2.970, 10: 3.078, 11: 3.173,
        12: 3.258, 13: 3.336, 14: 3.407, 15: 3.472, 16: 3.532,
        17: 3.588, 18: 3.640, 19: 3.689, 20: 3.735, 21: 3.778,
        22: 3.819, 23: 3.858, 24: 3.895, 25: 3.931, 26: 3.964,
        27: 3.997, 28: 4.027, 29: 4.057, 30: 4.086, 31: 4.113,
        32: 4.139, 33: 4.165, 34: 4.189, 35: 4.213, 36: 4.236,
        37: 4.259, 38: 4.280, 39: 4.301, 40: 4.322, 41: 4.341,
        42: 4.361, 43: 4.379, 44: 4.398, 45: 4.415, 46: 4.433,
        47: 4.450, 48: 4.466
    },

    getConstants: function (n) {
        if (n < 2) return { A2: 0, D3: 0, D4: 0 };
        // Priority check for extended high-cavity values
        if (this.SPC_CONSTANTS[n]) return this.SPC_CONSTANTS[n];

        // Fallback for n > 48 (Approximation using d2)
        if (n > 48) {
            console.warn(`[SPC] n=${n} exceeds table. Using approximation.`);
            var d2 = 4.466; // Flat fallback
            return { A2: 3 / (d2 * Math.sqrt(n)), D3: 0.56, D4: 1.44 };
        }

        return this.SPC_CONSTANTS[n] || { A2: 0, D3: 0, D4: 0 };
    },

    mean: function (data) {
        var filtered = data.filter(function (v) { 
            return v !== null && v !== undefined && v !== '' && !isNaN(parseFloat(v)); 
        }).map(v => parseFloat(v));
        if (filtered.length === 0) return null;
        var sum = 0;
        for (var i = 0; i < filtered.length; i++) sum += filtered[i];
        return sum / filtered.length;
    },

    min: function (data) {
        var filtered = data.filter(function (v) { 
            return v !== null && v !== undefined && v !== '' && !isNaN(parseFloat(v)); 
        }).map(v => parseFloat(v));
        return filtered.length > 0 ? Math.min.apply(null, filtered) : null;
    },

    max: function (data) {
        var filtered = data.filter(function (v) { 
            return v !== null && v !== undefined && v !== '' && !isNaN(parseFloat(v)); 
        }).map(v => parseFloat(v));
        return filtered.length > 0 ? Math.max.apply(null, filtered) : null;
    },

    range: function (data) {
        var mx = this.max(data);
        var mn = this.min(data);
        if (mx === null || mn === null) return 0;
        var r = mx - mn;
        return isNaN(r) ? 0 : r;
    },

    stdDev: function (data) {
        var filtered = data.filter(function (v) { 
            return v !== null && v !== undefined && v !== '' && !isNaN(parseFloat(v)); 
        }).map(v => parseFloat(v));
        if (filtered.length < 2) return null;
        var avg = this.mean(filtered);
        var sumSq = 0;
        for (var i = 0; i < filtered.length; i++) {
            sumSq += Math.pow(filtered[i] - avg, 2);
        }
        return Math.sqrt(sumSq / (filtered.length - 1));
    },

    withinStdDev: function (data) {
        var filtered = data.filter(function (v) { 
            return v !== null && v !== undefined && v !== '' && !isNaN(parseFloat(v)); 
        }).map(v => parseFloat(v));
        if (filtered.length < 2) return 0;
        var mrSum = 0;
        for (var i = 1; i < filtered.length; i++) {
            mrSum += Math.abs(filtered[i] - filtered[i - 1]);
        }
        return (mrSum / (filtered.length - 1)) / 1.128;
    },

    calculateXBarRLimits: function (dataMatrix, fixedCL, fixedSigma) {
        var self = this;
        var n = dataMatrix[0] ? dataMatrix[0].length : 0;
        var constants = this.getConstants(n);
        var xBars = [];
        var ranges = [];

        for (var i = 0; i < dataMatrix.length; i++) {
            var subgroup = dataMatrix[i];
            if (!subgroup) continue;
            
            // Strict filter: Only keep real numbers
            var filtered = subgroup.filter(function (v) { 
                return v !== null && v !== undefined && v !== '' && !isNaN(parseFloat(v)); 
            }).map(function(v) { return parseFloat(v); });
            
            if (filtered.length > 0) {
                var m = self.mean(filtered);
                var r = self.range(filtered);
                
                if (!isNaN(m)) xBars.push(m);
                if (!isNaN(r)) ranges.push(r);
            }
        }

        // Use fixed limits if provided (Phase II), otherwise calculate from current data (Phase I)
        var xDoubleBar = fixedCL !== undefined ? fixedCL : (xBars.length > 0 ? this.mean(xBars) : 0);
        var rBar = ranges.length > 0 ? this.mean(ranges) : 0;
        var sigma = (fixedSigma !== undefined) ? fixedSigma : (constants.A2 * rBar / 3); 
        
        // If fixedSigma is not provided, we calculate standard Shewhart sigma: (A2 * R-bar) / 3

        var xUCL = xDoubleBar + 3 * sigma;
        var xLCL = xDoubleBar - 3 * sigma;
        var rUCL = constants.D4 * rBar;
        var rLCL = constants.D3 * rBar;

        var results = {
            type: (fixedCL !== undefined) ? 'fixed-baseline' : 'standard',
            xBar: {
                data: xBars,
                UCL: xUCL,
                CL: xDoubleBar,
                LCL: xLCL,
                sigma: sigma
            },
            R: {
                data: ranges,
                UCL: rUCL,
                CL: rBar,
                LCL: rLCL
            },
            summary: { n: n, k: xBars.length, xDoubleBar: xDoubleBar, rBar: rBar, isFixed: (fixedCL !== undefined) }
        };

        results.xBar.violations = this.checkNelsonRules(xBars, xDoubleBar, results.xBar.sigma);
        results.R.violations = this.checkRangeViolations(ranges, rUCL, rLCL);
        return results;
    },

    /**
     * checkRangeViolations - Specific violation check for R-charts (Rule 1 focus)
     */
    checkRangeViolations: function (data, ucl, lcl) {
        var violations = [];
        for (var i = 0; i < data.length; i++) {
            if (data[i] === null || isNaN(data[i])) continue;
            var rules = [];
            if (data[i] > ucl || data[i] < lcl) rules.push(1);
            if (rules.length > 0) violations.push({ index: i, rules: rules });
        }
        return violations;
    },

    /**
     * recalculateWithFixedLimits - Utility to apply Phase II monitoring
     * @param {Array} dataMatrix - The full dataset matrix
     * @param {Number} cl - Fixed Center Line 
     * @param {Number} sigma - Fixed Sigma (Standard Deviation of the mean)
     */
    recalculateWithFixedLimits: function (dataMatrix, cl, sigma) {
        return this.calculateXBarRLimits(dataMatrix, cl, sigma);
    },

    /**
     * calculateExtendedBatchLimits - For Multi-Cavity (Model C/D)
     * Incorporates between-cavity variation into limits to avoid false alarms.
     */
    calculateExtendedBatchLimits: function (dataMatrix, fixedCL, fixedSigma) {
        var standard = this.calculateXBarRLimits(dataMatrix, fixedCL);
        var n = standard.summary.n;

        // Use fixed limits if provided, otherwise estimate from data
        var xDoubleBar = standard.xBar.CL;
        var extendedSigma;

        if (fixedSigma !== undefined) {
            extendedSigma = fixedSigma;
        } else {
            // Calculate Overall Variation (sigma overall) from data
            var allValues = dataMatrix.flat().filter(function (v) { 
                return v !== null && v !== undefined && v !== '' && !isNaN(parseFloat(v)); 
            }).map(v => parseFloat(v));
            extendedSigma = this.stdDev(allValues);
        }

        // Extended Shewhart Limit: CL +/- 3 * sigmaOverall / sqrt(n_effective)
        var sigmaOfMean = extendedSigma / Math.sqrt(n || 1);
        var xUCL = xDoubleBar + (3 * sigmaOfMean);
        var xLCL = xDoubleBar - (3 * sigmaOfMean);

        standard.type = (fixedCL !== undefined) ? 'fixed-extended' : 'extended';
        standard.xBar.UCL = xUCL;
        standard.xBar.LCL = xLCL;
        standard.xBar.sigma = sigmaOfMean;

        // Re-check rules with new limits
        standard.xBar.violations = this.checkNelsonRules(standard.xBar.data, xDoubleBar, standard.xBar.sigma);
        standard.R.violations = this.checkRangeViolations(standard.R.data, standard.R.UCL, standard.R.LCL);

        return standard;
    },

    checkNelsonRules: function (data, cl, sigma) {
        var violations = [];
        if (data.length === 0 || sigma === 0) return violations;

        for (var i = 0; i < data.length; i++) {
            var rules = [];
            if (Math.abs(data[i] - cl) > 3 * sigma) rules.push(1);
            if (i >= 8) {
                var sameSide = true, side = data[i] > cl;
                for (var j = i - 8; j <= i; j++) { if ((data[j] > cl) !== side || data[j] === cl) { sameSide = false; break; } }
                if (sameSide) rules.push(2);
            }
            if (i >= 5) {
                var inc = true, dec = true;
                for (var j = i - 5; j < i; j++) { if (data[j + 1] <= data[j]) inc = false; if (data[j + 1] >= data[j]) dec = false; }
                if (inc || dec) rules.push(3);
            }
            if (i >= 13) {
                var isAlt = true;
                for (var j = i - 13; j < i; j++) { if ((data[j + 1] >= data[j] && data[j] >= data[j - 1]) || (data[j + 1] <= data[j] && data[j] <= data[j - 1])) { isAlt = false; break; } }
                if (isAlt) rules.push(4);
            }
            if (i >= 2) {
                var up = 0, lo = 0;
                for (var j = i - 2; j <= i; j++) { if (data[j] > cl + 2 * sigma) up++; if (data[j] < cl - 2 * sigma) lo++; }
                if (up >= 2 || lo >= 2) rules.push(5);
            }
            if (i >= 4) {
                var up = 0, lo = 0;
                for (var j = i - 4; j <= i; j++) { if (data[j] > cl + 1 * sigma) up++; if (data[j] < cl - 1 * sigma) lo++; }
                if (up >= 4 || lo >= 4) rules.push(6);
            }
            if (i >= 14) {
                var withinOne = true;
                for (var j = i - 14; j <= i; j++) { if (Math.abs(data[j] - cl) > sigma) { withinOne = false; break; } }
                if (withinOne) rules.push(7);
            }
            if (i >= 7) {
                var countHi = 0, countLo = 0;
                var anyInside = false;
                for (var j = i - 7; j <= i; j++) {
                    var diff = data[j] - cl;
                    if (Math.abs(diff) <= sigma) { anyInside = true; break; }
                    if (diff > 0) countHi++;
                    else countLo++;
                }
                if (!anyInside && countHi > 0 && countLo > 0) rules.push(8);
            }
            if (rules.length > 0) violations.push({ index: i, rules: rules });
        }
        return violations;
    },

    calculateProcessCapability: function (data, usl, lsl, rBar, subgroupSize) {
        var filtered = data.filter(function (v) { 
            return v !== null && v !== undefined && v !== '' && !isNaN(parseFloat(v)); 
        }).map(v => parseFloat(v));
        if (filtered.length < 2) return { Cp: 0, Cpk: 0, Pp: 0, Ppk: 0, mean: 0, withinStdDev: 0, overallStdDev: 0, count: 0 };

        var mean = this.mean(filtered);
        // Use Xbar-R formula if rBar and subgroupSize provided, otherwise fallback to I-MR
        var withinStdDev;
        if (rBar !== undefined && subgroupSize !== undefined && subgroupSize >= 2) {
            var d2 = this.D2_CONSTANTS[subgroupSize] || (subgroupSize > 25 ? 3.931 : 1.128);
            withinStdDev = rBar / d2;
        } else {
            withinStdDev = this.withinStdDev(filtered);
        }
        var overallStdDev = this.stdDev(filtered);
        
        // --- Numeric Safeguard ---
        var nUSL = parseFloat(usl);
        var nLSL = parseFloat(lsl);
        var nMean = parseFloat(mean);
        
        if (isNaN(nUSL) || isNaN(nLSL)) return { Cp: 0, Cpk: 0, Pp: 0, Ppk: 0, mean: nMean, withinStdDev: withinStdDev, overallStdDev: overallStdDev, count: filtered.length };
        
        var tolerance = nUSL - nLSL;

        var Cp = (withinStdDev > 1e-9) ? tolerance / (6 * withinStdDev) : 0;
        var Cpu = (withinStdDev > 1e-9) ? (nUSL - nMean) / (3 * withinStdDev) : 0;
        var Cpl = (withinStdDev > 1e-9) ? (nMean - nLSL) / (3 * withinStdDev) : 0;
        var Cpk = isNaN(Cpu) || isNaN(Cpl) ? 0 : Math.min(Cpu, Cpl);

        var Pp = (overallStdDev > 1e-9) ? tolerance / (6 * overallStdDev) : 0;
        var Ppu = (overallStdDev > 1e-9) ? (nUSL - nMean) / (3 * overallStdDev) : 0;
        var Ppl = (overallStdDev > 1e-9) ? (nMean - nLSL) / (3 * overallStdDev) : 0;
        var Ppk = isNaN(Ppu) || isNaN(Ppl) ? 0 : Math.min(Ppu, Ppl);

        return {
            Cp: Cp, Cpk: Cpk, Pp: Pp, Ppk: Ppk,
            mean: mean, withinStdDev: withinStdDev, overallStdDev: overallStdDev,
            min: this.min(filtered), max: this.max(filtered),
            range: this.range(filtered), count: filtered.length
        };
    },

    /**
     * analyzeCavityBalance - Professional assessment of thermal/physical balance
     * @param {Array} cavityStats - Array of {name, mean, Cpk}
     * @param {Object} specs - {usl, lsl, target}
     */
    analyzeCavityBalance: function (cavityStats, specs) {
        if (!cavityStats || cavityStats.length < 2) return null;

        var means = cavityStats.map(function (s) { return s.mean; });
        var avgOfMeans = this.mean(means);
        var maxMean = Math.max.apply(null, means);
        var minMean = Math.min.apply(null, means);
        var rangeOfMeans = maxMean - minMean;

        var tolerance = (specs.usl - specs.lsl) || 1;
        var imbalanceRatio = (rangeOfMeans / tolerance) * 100; // % of tolerance consumed by imbalance

        var status = 'Excellent';
        var color = '#10b981';
        if (imbalanceRatio > 25) { status = 'Poor'; color = '#f43f5e'; }
        else if (imbalanceRatio > 10) { status = 'Fair'; color = '#f59e0b'; }

        return {
            avgOfMeans: avgOfMeans,
            rangeOfMeans: rangeOfMeans,
            imbalanceRatio: imbalanceRatio,
            status: status,
            color: color,
            advice: status === 'Excellent' ?
                { zh: '模穴平衡良好，製程穩定。', en: 'Cavity balance is excellent, process is stable.' } :
                (status === 'Fair' ?
                    { zh: '偵測到輕微模穴不平衡，建議檢查流道平衡。', en: 'Minor cavity imbalance detected. Recommendation: check runner balance.' } :
                    { zh: '嚴重模穴不平衡！建議優先進行模具維修或熱流道調整。', en: 'Severe cavity imbalance! Priority recommendation: mold maintenance or hot runner adjustment.' })
        };
    },

    /**
     * calculateDistStats - Distribution Health (Skewness & Kurtosis)
     * @param {Array} data - Raw data points
     */
    calculateDistStats: function (data) {
        var filtered = data.filter(function (v) { 
            return v !== null && v !== undefined && v !== '' && !isNaN(parseFloat(v)); 
        }).map(v => parseFloat(v));
        var n = filtered.length;
        if (n < 3) return { skewness: 0, kurtosis: 0 };

        var mean = this.mean(filtered);
        var stdDev = this.stdDev(filtered);
        if (stdDev === 0) return { skewness: 0, kurtosis: 0 };

        var sum3 = 0, sum4 = 0;
        for (var i = 0; i < n; i++) {
            var z = (filtered[i] - mean) / stdDev;
            sum3 += Math.pow(z, 3);
            sum4 += Math.pow(z, 4);
        }

        var skewness = (n / ((n - 1) * (n - 2))) * sum3;
        var kurtosis = ((n * (n + 1)) / ((n - 1) * (n - 2) * (n - 3))) * sum4 - (3 * Math.pow(n - 1, 2)) / ((n - 2) * (n - 3));

        return { skewness: skewness, kurtosis: kurtosis };
    },

    /**
     * analyzeVarianceSource - AI Diagnosis of variation origins
     */
    analyzeVarianceSource: function (cpk, ppk, distStats) {
        if (!cpk || !ppk) return null;

        var stability = ppk / cpk;
        var primarySource = '';
        var recommendation = '';
        var color = '#6366f1';

        if (stability < 0.8) {
            primarySource = { zh: '批次間變異 (Shot-to-Shot / Equipment)', en: 'Shot-to-Shot Variation (Equipment)' };
            recommendation = { zh: '製程穩定度不足。請檢查機台參數再現性、原料穩定度或環境溫濕度波動。', en: 'Insufficient process stability. Check machine parameter repeatability, material consistency, or environmental fluctuations.' };
            color = '#f59e0b';
        } else {
            primarySource = { zh: '組內變異 (Within-Shot / Tooling)', en: 'Within-Shot Variation (Tooling)' };
            recommendation = { zh: '製程控制良好，變異主要來自模穴差異或單次注射內的波動。建議檢查模穴平衡性或澆道設計。', en: 'Good process control. Variation primarily stems from cavity differences or within-shot fluctuations. Check cavity balance or runner design.' };
            color = '#10b981';
        }

        // Add distribution warning
        var distWarning = null;
        if (Math.abs(distStats.skewness) > 1) {
            distWarning = {
                zh: '警告：數據呈現偏態分佈 (Skewed)，Cpk 數值可能存在統計偏差，建議確認是否有模穴尺寸不一或數據取樣偏誤。',
                en: 'Warning: Data is skewed. Cpk values may have statistical bias. Suggest verifying cavity dimensions or sampling bias.'
            };
        }

        return {
            stability: stability,
            source: primarySource,
            advice: recommendation,
            distWarning: distWarning,
            color: color
        };
    },

    /**
     * analyzeGroupStability - Intelligent diagnosis for Group Analysis mode
     * @param {Array} groupStats - Array of {batch, avg, max, min, range, count}
     * @param {Object} specs - {usl, lsl, target}
     */
    analyzeGroupStability: function (groupStats, specs) {
        if (!groupStats || groupStats.length < 2) return null;

        var ranges = groupStats.map(function (g) { return g.range; });
        var avgRange = this.mean(ranges);
        var maxRange = Math.max.apply(null, ranges);
        var tolerance = (specs.usl - specs.lsl) || 1;

        // Spread of ranges: consistency of variation across groups
        var rangeStdDev = this.stdDev(ranges);
        var consistencyScore = avgRange > 0 ? (rangeStdDev / avgRange) : 0;

        var status = { zh: '穩定', en: 'Stable' };
        var color = '#10b981';
        var advice = { zh: '各組變異度一致，製程控制極為穩定。', en: 'Variation across groups is consistent, process control is very stable.' };

        if (consistencyScore > 0.4) {
            status = { zh: '不穩定', en: 'Unstable' };
            color = '#f43f5e';
            advice = { zh: '組間變異波動劇烈！部分批次的內部品質差異過大，建議檢查機台保壓穩定性與模溫控制。', en: 'Intense variation between groups! Internal quality of some batches varies significantly. Check machine holding pressure and mold temperature.' };
        } else if (consistencyScore > 0.2) {
            status = { zh: '警告', en: 'Warning' };
            color = '#f59e0b';
            advice = { zh: '製程穩定度有所起伏。建議監測模具排氣狀況與原料進料壓力。', en: 'Process stability is fluctuating. Monitor mold venting and material inlet pressure.' };
        }

        // Check for specific outliers
        var outlierBatches = groupStats.filter(function (g) { return g.range > (avgRange + 2 * rangeStdDev); });
        if (outlierBatches.length > 0) {
            var batchList = outlierBatches.map(function (b) { return b.batch; }).join(', ');
            advice.zh += ' 注意：批號 ' + batchList + ' 之變異明顯高於平均。';
            advice.en += ' Note: Variation in batches ' + batchList + ' is significantly higher than average.';
        }

        return {
            status: status,
            color: color,
            avgRange: avgRange,
            maxRange: maxRange,
            consistencyScore: consistencyScore,
            advice: advice
        };
    },

    getCapabilityColor: function (cpk) {
        if (cpk >= 1.67) return { bg: '#c6efce', text: '#006100' };
        if (cpk >= 1.33) return { bg: '#c6efce', text: '#006100' };
        if (cpk >= 1.0) return { bg: '#ffeb9c', text: '#9c5700' };
        return { bg: '#ffc7ce', text: '#9c0006' };
    },

    round: function (value, decimals) {
        if (value === null || value === undefined || isNaN(value)) return '--';
        decimals = decimals || 4;
        var p = Math.pow(10, decimals);
        return Math.round(value * p) / p;
    }
};
