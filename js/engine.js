/**
 * SPC ENGINE - Core Statistical Calculations
 * Based on industry standard formulas for X-Bar and R charts.
 */

var SPCEngine = {
    SPC_CONSTANTS: {
        2: { A2: 1.88, D3: 0, D4: 3.267 },
        3: { A2: 1.023, D3: 0, D4: 2.373 },
        4: { A2: 0.729, D3: 0, D4: 2.282 },
        5: { A2: 0.577, D3: 0, D4: 2.115 },
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
        20: { A2: 0.18, D3: 0.415, D4: 1.585 },
        21: { A2: 0.173, D3: 0.425, D4: 1.575 },
        22: { A2: 0.167, D3: 0.434, D4: 1.566 },
        23: { A2: 0.162, D3: 0.443, D4: 1.557 },
        24: { A2: 0.157, D3: 0.451, D4: 1.548 },
        25: { A2: 0.153, D3: 0.459, D4: 1.541 }
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
        if (n > 25) return { A2: 3 / Math.sqrt(n), D3: 0.5, D4: 1.5 };
        return this.SPC_CONSTANTS[n];
    },

    mean: function (data) {
        var filtered = data.filter(function (v) { return v !== null && !isNaN(v); });
        if (filtered.length === 0) return 0;
        var sum = 0;
        for (var i = 0; i < filtered.length; i++) sum += filtered[i];
        return sum / filtered.length;
    },

    min: function (data) {
        var filtered = data.filter(function (v) { return v !== null && !isNaN(v); });
        return filtered.length > 0 ? Math.min.apply(null, filtered) : 0;
    },

    max: function (data) {
        var filtered = data.filter(function (v) { return v !== null && !isNaN(v); });
        return filtered.length > 0 ? Math.max.apply(null, filtered) : 0;
    },

    range: function (data) {
        return this.max(data) - this.min(data);
    },

    stdDev: function (data) {
        var filtered = data.filter(function (v) { return v !== null && !isNaN(v); });
        if (filtered.length < 2) return 0;
        var avg = this.mean(filtered);
        var sumSq = 0;
        for (var i = 0; i < filtered.length; i++) {
            sumSq += Math.pow(filtered[i] - avg, 2);
        }
        return Math.sqrt(sumSq / (filtered.length - 1));
    },

    withinStdDev: function (data) {
        var filtered = data.filter(function (v) { return v !== null && !isNaN(v); });
        if (filtered.length < 2) return 0;
        var mrSum = 0;
        for (var i = 1; i < filtered.length; i++) {
            mrSum += Math.abs(filtered[i] - filtered[i - 1]);
        }
        return (mrSum / (filtered.length - 1)) / 1.128;
    },

    calculateXBarRLimits: function (dataMatrix) {
        var self = this;
        var n = dataMatrix[0] ? dataMatrix[0].length : 0;
        var constants = this.getConstants(n);
        var xBars = [];
        var ranges = [];

        for (var i = 0; i < dataMatrix.length; i++) {
            var subgroup = dataMatrix[i];
            var filtered = subgroup.filter(function (v) { return v !== null && !isNaN(v); });
            if (filtered.length > 0) {
                xBars.push(self.mean(filtered));
                ranges.push(self.range(filtered));
            }
        }

        var xDoubleBar = this.mean(xBars);
        var rBar = this.mean(ranges);

        var xUCL = xDoubleBar + constants.A2 * rBar;
        var xLCL = xDoubleBar - constants.A2 * rBar;
        var rUCL = constants.D4 * rBar;
        var rLCL = constants.D3 * rBar;

        var results = {
            xBar: {
                data: xBars,
                UCL: xUCL,
                CL: xDoubleBar,
                LCL: xLCL,
                sigma: (xUCL - xDoubleBar) / 3
            },
            R: {
                data: ranges,
                UCL: rUCL,
                CL: rBar,
                LCL: rLCL
            },
            summary: { n: n, k: xBars.length, xDoubleBar: xDoubleBar, rBar: rBar }
        };

        results.xBar.violations = this.checkNelsonRules(xBars, xDoubleBar, results.xBar.sigma);
        return results;
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
            if (rules.length > 0) violations.push({ index: i, rules: rules });
        }
        return violations;
    },

    calculateProcessCapability: function (data, usl, lsl, rBar, subgroupSize) {
        var filtered = data.filter(function (v) { return v !== null && !isNaN(v); });
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
        var tolerance = usl - lsl;

        var Cp = withinStdDev > 0 ? tolerance / (6 * withinStdDev) : 0;
        var Cpu = withinStdDev > 0 ? (usl - mean) / (3 * withinStdDev) : 0;
        var Cpl = withinStdDev > 0 ? (mean - lsl) / (3 * withinStdDev) : 0;
        var Cpk = Math.min(Cpu, Cpl);

        var Pp = overallStdDev > 0 ? tolerance / (6 * overallStdDev) : 0;
        var Ppu = overallStdDev > 0 ? (usl - mean) / (3 * overallStdDev) : 0;
        var Ppl = overallStdDev > 0 ? (mean - lsl) / (3 * overallStdDev) : 0;
        var Ppk = Math.min(Ppu, Ppl);

        return {
            Cp: Cp, Cpk: Cpk, Pp: Pp, Ppk: Ppk,
            mean: mean, withinStdDev: withinStdDev, overallStdDev: overallStdDev,
            min: this.min(filtered), max: this.max(filtered),
            range: this.range(filtered), count: filtered.length
        };
    },

    getCapabilityColor: function (cpk) {
        if (cpk >= 1.67) return { bg: '#c6efce', text: '#006100' };
        if (cpk >= 1.33) return { bg: '#c6efce', text: '#006100' };
        if (cpk >= 1.0) return { bg: '#ffeb9c', text: '#9c5700' };
        return { bg: '#ffc7ce', text: '#9c0006' };
    },

    round: function (value, decimals) {
        decimals = decimals || 4;
        return Math.round(value * Math.pow(10, decimals)) / Math.pow(10, decimals);
    }
};
