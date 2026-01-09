// ============================================================================
// SPC Calculation Engine - Core Statistical Process Control Logic
// ============================================================================

export class SPCEngine {
    // ========================================================================
    // SPC Constants Table (A2, D3, D4) for n=2 to 32
    // Based on standard SPC factor tables
    // ========================================================================

    static SPC_CONSTANTS = {
        2: { A2: 1.88, D3: 0, D4: 3.267, d2: 1.128 },
        3: { A2: 1.023, D3: 0, D4: 2.373, d2: 1.693 },
        4: { A2: 0.729, D3: 0, D4: 2.282, d2: 2.059 },
        5: { A2: 0.577, D3: 0, D4: 2.115, d2: 2.326 },
        6: { A2: 0.483, D3: 0, D4: 2.004, d2: 2.534 },
        7: { A2: 0.419, D3: 0.076, D4: 1.924, d2: 2.704 },
        8: { A2: 0.373, D3: 0.136, D4: 1.864, d2: 2.847 },
        9: { A2: 0.337, D3: 0.184, D4: 1.816, d2: 2.970 },
        10: { A2: 0.308, D3: 0.223, D4: 1.777, d2: 3.078 },
        11: { A2: 0.285, D3: 0.256, D4: 1.744, d2: 3.173 },
        12: { A2: 0.266, D3: 0.283, D4: 1.717, d2: 3.258 },
        13: { A2: 0.249, D3: 0.307, D4: 1.693, d2: 3.336 },
        14: { A2: 0.235, D3: 0.328, D4: 1.672, d2: 3.407 },
        15: { A2: 0.223, D3: 0.347, D4: 1.653, d2: 3.472 },
        16: { A2: 0.212, D3: 0.363, D4: 1.637, d2: 3.532 },
        17: { A2: 0.203, D3: 0.378, D4: 1.622, d2: 3.588 },
        18: { A2: 0.194, D3: 0.391, D4: 1.608, d2: 3.640 },
        19: { A2: 0.187, D3: 0.403, D4: 1.597, d2: 3.689 },
        20: { A2: 0.18, D3: 0.415, D4: 1.585, d2: 3.735 },
        21: { A2: 0.173, D3: 0.425, D4: 1.575, d2: 3.778 },
        22: { A2: 0.167, D3: 0.434, D4: 1.566, d2: 3.819 },
        23: { A2: 0.162, D3: 0.443, D4: 1.557, d2: 3.858 },
        24: { A2: 0.157, D3: 0.451, D4: 1.548, d2: 3.895 },
        25: { A2: 0.153, D3: 0.459, D4: 1.541, d2: 3.931 },
        26: { A2: 0.149, D3: 0.466, D4: 1.534, d2: 3.964 },
        27: { A2: 0.145, D3: 0.473, D4: 1.527, d2: 3.997 },
        28: { A2: 0.141, D3: 0.48, D4: 1.52, d2: 4.027 },
        29: { A2: 0.138, D3: 0.486, D4: 1.514, d2: 4.057 },
        30: { A2: 0.135, D3: 0.492, D4: 1.508, d2: 4.086 },
        31: { A2: 0.132, D3: 0.498, D4: 1.502, d2: 4.113 },
        32: { A2: 0.129, D3: 0.504, D4: 1.496, d2: 4.139 }
    };

    static getConstants(n) {
        if (n < 2) return { A2: 0, D3: 0, D4: 0, d2: 1.128 };
        if (n > 32) {
            // Approximation for large n
            return {
                A2: 3 / Math.sqrt(n),
                D3: 0.5,
                D4: 1.5,
                d2: 1.128
            };
        }
        return this.SPC_CONSTANTS[n];
    }

    // ========================================================================
    // Basic Statistical Functions
    // ========================================================================

    static mean(data) {
        const filtered = data.filter(v => v !== null && !isNaN(v));
        if (filtered.length === 0) return 0;
        return filtered.reduce((sum, val) => sum + val, 0) / filtered.length;
    }

    static sum(data) {
        return data.filter(v => v !== null && !isNaN(v)).reduce((sum, val) => sum + val, 0);
    }

    static min(data) {
        const filtered = data.filter(v => v !== null && !isNaN(v));
        return filtered.length > 0 ? Math.min(...filtered) : 0;
    }

    static max(data) {
        const filtered = data.filter(v => v !== null && !isNaN(v));
        return filtered.length > 0 ? Math.max(...filtered) : 0;
    }

    static range(data) {
        return this.max(data) - this.min(data);
    }

    // Overall standard deviation (sample std dev)
    static stdDev(data) {
        const filtered = data.filter(v => v !== null && !isNaN(v));
        if (filtered.length < 2) return 0;

        const avg = this.mean(filtered);
        const sumSq = filtered.reduce((sum, val) => sum + Math.pow(val - avg, 2), 0);
        return Math.sqrt(sumSq / (filtered.length - 1));
    }

    // Within-group standard deviation using moving range method
    static withinStdDev(data) {
        const filtered = data.filter(v => v !== null && !isNaN(v));
        if (filtered.length < 2) return 0;

        // Calculate moving ranges
        let mrSum = 0;
        let mrCount = 0;

        for (let i = 1; i < filtered.length; i++) {
            mrSum += Math.abs(filtered[i] - filtered[i - 1]);
            mrCount++;
        }

        const avgMR = mrCount > 0 ? mrSum / mrCount : 0;
        const d2 = 1.128; // d2 constant for n=2

        return avgMR / d2;
    }

    // ========================================================================
    // X-Bar R Chart Calculations
    // ========================================================================

    static calculateXBarRLimits(dataMatrix) {
        const n = dataMatrix[0].length; // Number of samples per subgroup (cavities)
        const k = dataMatrix.length;     // Number of subgroups (batches)

        // Get SPC constants
        const { A2, D3, D4 } = this.getConstants(n);

        // Calculate X-bar and R for each subgroup
        const xBars = [];
        const ranges = [];

        dataMatrix.forEach(subgroup => {
            const filtered = subgroup.filter(v => v !== null && !isNaN(v));
            if (filtered.length > 0) {
                xBars.push(this.mean(filtered));
                ranges.push(this.range(filtered));
            }
        });

        // Calculate grand average (X-double-bar) and average range (R-bar)
        const xDoubleBar = this.mean(xBars);
        const rBar = this.mean(ranges);

        // Calculate control limits
        return {
            xBar: {
                data: xBars,
                UCL: xDoubleBar + A2 * rBar,
                CL: xDoubleBar,
                LCL: xDoubleBar - A2 * rBar
            },
            R: {
                data: ranges,
                UCL: D4 * rBar,
                CL: rBar,
                LCL: D3 * rBar
            },
            summary: {
                n: n,
                k: k,
                xDoubleBar: xDoubleBar,
                rBar: rBar,
                sumXBar: this.sum(xBars),
                sumR: this.sum(ranges)
            }
        };
    }

    // ========================================================================
    // Process Capability Calculations (Cp, Cpk, Pp, Ppk)
    // ========================================================================

    static calculateProcessCapability(data, usl, lsl) {
        const filtered = data.filter(v => v !== null && !isNaN(v));
        if (filtered.length < 2) {
            return { Cp: 0, Cpk: 0, Pp: 0, Ppk: 0, mean: 0, withinStdDev: 0, overallStdDev: 0 };
        }

        const mean = this.mean(filtered);
        const withinStdDev = this.withinStdDev(filtered);
        const overallStdDev = this.stdDev(filtered);
        const tolerance = usl - lsl;

        // Short-term capability (using within-group std dev)
        const Cp = tolerance / (6 * withinStdDev);
        const Cpu = (usl - mean) / (3 * withinStdDev);
        const Cpl = (mean - lsl) / (3 * withinStdDev);
        const Cpk = Math.min(Cpu, Cpl);

        // Long-term performance (using overall std dev)
        const Pp = tolerance / (6 * overallStdDev);
        const Ppu = (usl - mean) / (3 * overallStdDev);
        const Ppl = (mean - lsl) / (3 * overallStdDev);
        const Ppk = Math.min(Ppu, Ppl);

        return {
            Cp: Cp,
            Cpk: Cpk,
            Pp: Pp,
            Ppk: Ppk,
            mean: mean,
            withinStdDev: withinStdDev,
            overallStdDev: overallStdDev,
            min: this.min(filtered),
            max: this.max(filtered),
            range: this.range(filtered),
            count: filtered.length
        };
    }

    // ========================================================================
    // Out-of-Control Detection
    // ========================================================================

    static detectOutOfControl(value, ucl, lcl) {
        return value > ucl || value < lcl;
    }

    static detectOutOfControlPoints(data, ucl, lcl) {
        return data.map((value, index) => ({
            index: index,
            value: value,
            outOfControl: this.detectOutOfControl(value, ucl, lcl)
        }));
    }

    // ========================================================================
    // Capability Index Color Coding
    // ========================================================================

    static getCapabilityColor(cpk) {
        if (cpk >= 1.67) return { bg: '#c6efce', text: '#006100' }; // Green - Excellent
        if (cpk >= 1.33) return { bg: '#c6efce', text: '#006100' }; // Green - Good
        if (cpk >= 1.0) return { bg: '#ffeb9c', text: '#9c5700' };  // Yellow - Acceptable
        return { bg: '#ffc7ce', text: '#9c0006' };                  // Red - Insufficient
    }

    // ========================================================================
    // Utility Functions
    // ========================================================================

    static round(value, decimals = 4) {
        return Math.round(value * Math.pow(10, decimals)) / Math.pow(10, decimals);
    }
}
