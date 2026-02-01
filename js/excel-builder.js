/**
 * SPC Excel Builder - Generates Excel files mimicking the VBA output exactly.
 * Uses ExcelJS for advanced formatting and chart embedding.
 */
class SPCExcelBuilder {
    constructor(analysisResults, chartsImages = {}, templateBuffer = null, templateCapacity = 0) {
        this.results = analysisResults;
        this.chartsImages = chartsImages; // { xbar: 'base64...', r: 'base64...' }
        this.templateBuffer = templateBuffer;
        this.templateCapacity = templateCapacity; // The theoretical max N of the template
        this.workbook = new ExcelJS.Workbook();
        if (!this.templateBuffer) {
            this.setupWorkbook();
        }
    }

    setupWorkbook() {
        this.workbook.creator = 'Wesley Chang @ Mouldex';
        this.workbook.created = new Date();
    }

    async generate() {
        if (this.templateBuffer) {
            await this.workbook.xlsx.load(this.templateBuffer);
            if (this.results.type === 'batch') {
                await this.generateFromTemplate();
            } else {
                // Fallback for simple report if needed, or separate logic
                await this.generateFromTemplate();
            }
        } else {
            // Manual Build Mode
            if (this.results.type === 'batch') {
                await this.generateBatchReport();
            } else {
                this.generateSimpleReport();
            }
        }
        return await this.workbook.xlsx.writeBuffer();
    }

    async generateFromTemplate() {
        const data = this.results;
        const info = data.productInfo;
        const specs = data.specs;
        const totalBatches = data.batchNames.length;
        const cavCount = data.xbarR.summary.n;
        const templateN = this.templateCapacity || cavCount;
        const maxBatchesPerSheet = 25;
        const totalSheets = Math.ceil(totalBatches / maxBatchesPerSheet);

        // --- Sheet 1: Template Fill ---
        const ws = this.workbook.worksheets[0];
        if (!ws) throw new Error("Template loaded but has no sheets");

        // Rename Sheet
        ws.name = `${info.item || 'Sheet'}-001`;

        // 1. Fill Header Info
        if (info.name) ws.getCell('C2').value = info.name;
        if (info.dept) ws.getCell('O2').value = info.dept;
        const bStart = data.batchNames[0] || '';
        const s1EndIdx = Math.min(24, totalBatches - 1);
        const bEnd = data.batchNames[s1EndIdx] || '';
        ws.getCell('U2').value = `${bStart}\n到\n${bEnd}`;

        if (info.item) ws.getCell('C3').value = info.item;
        if (specs.usl) ws.getCell('I3').value = specs.usl;
        if (data.xbarR.xBar.UCL) ws.getCell('M3').value = this.round(data.xbarR.xBar.UCL);
        if (data.xbarR.R.UCL) ws.getCell('N3').value = this.round(data.xbarR.R.UCL);

        if (info.unit) ws.getCell('C4').value = info.unit;
        if (specs.target) ws.getCell('I4').value = specs.target;
        if (data.xbarR.xBar.CL) ws.getCell('M4').value = this.round(data.xbarR.xBar.CL);
        if (data.xbarR.R.CL) ws.getCell('N4').value = this.round(data.xbarR.R.CL);
        if (info.inspector) ws.getCell('Q4').value = info.inspector;

        if (specs.lsl) ws.getCell('I5').value = specs.lsl;
        if (data.xbarR.xBar.LCL) ws.getCell('M5').value = this.round(data.xbarR.xBar.LCL);

        // 2. Fill Data
        const sampleStartRow = 7;
        const sampleEndRow = sampleStartRow + templateN - 1;
        const sumRow = sampleEndRow + 1;
        const meanRow = sumRow + 1;
        const rangeRow = meanRow + 1;

        for (let i = 0; i < 25; i++) {
            const batchIdx = i;
            const colIdx = 3 + i;

            if (batchIdx > s1EndIdx) {
                ws.getCell(6, colIdx).value = null;
                for (let r = sampleStartRow; r <= rangeRow; r++) ws.getCell(r, colIdx).value = null;
                continue;
            }

            ws.getCell(6, colIdx).value = data.batchNames[batchIdx];

            for (let cav = 0; cav < templateN; cav++) {
                const r = sampleStartRow + cav;
                if (cav < cavCount) {
                    const val = data.dataMatrix[batchIdx][cav];
                    ws.getCell(r, colIdx).value = (val !== null && val !== undefined) ? val : null;
                } else {
                    ws.getCell(r, colIdx).value = null;
                }
            }

            let batSum = 0;
            let batCount = 0;
            data.dataMatrix[batchIdx].forEach(v => { if (v != null) { batSum += v; batCount++; } });
            ws.getCell(sumRow, colIdx).value = batCount > 0 ? batSum : null;

            ws.getCell(meanRow, colIdx).value = data.xbarR.xBar.data[batchIdx];
            ws.getCell(rangeRow, colIdx).value = data.xbarR.R.data[batchIdx];
        }

        const gs = data.xbarR.summary;
        ws.getCell(7, 28).value = `ΣX̄ = ${this.round(gs.xDoubleBar * gs.k)}`;
        ws.getCell(7, 28).alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };

        ws.getCell(9, 28).value = `X̿ = ${this.round(gs.xDoubleBar)}`;
        ws.getCell(9, 28).alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };

        ws.getCell(11, 28).value = `ΣR = ${this.round(gs.rBar * gs.k)}`;
        ws.getCell(11, 28).alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };

        ws.getCell(13, 28).value = `R̄ = ${this.round(gs.rBar)}`;
        ws.getCell(13, 28).alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };

        // Footer Sheet 1
        const footerRow = rangeRow + 35; // Approx
        const f2 = ws.getCell(footerRow, 28);
        // Ensure cell exists before setting value to avoid error if outside range, but ExcelJS handles this.
        f2.value = `第 1 頁，共 ${totalSheets} 頁`;

        console.log(`Template Filled (ExcelJS Mode)`);

        // --- Subsequent Sheets: Manual Build ---
        if (totalSheets > 1) {
            for (let sheetIdx = 2; sheetIdx <= totalSheets; sheetIdx++) {
                await this.buildManualSheet(sheetIdx, totalSheets);
            }
        }
    }

    // This section seems to be outside the class or part of another function.
    // The instruction was to replace methods within the class.
    // Assuming the original content had a `generateFromTemplate` method that was part of the class,
    // and the following code was part of a different function or was extraneous.
    // I will remove the extraneous code that was between the old `generate` and `generateBatchReport`
    // and insert the new `generateFromTemplate` method.

    async generateBatchReport() {
        const data = this.results;
        const totalBatches = data.batchNames.length;
        const maxBatchesPerSheet = 25;
        const totalSheets = Math.ceil(totalBatches / maxBatchesPerSheet);

        for (let sheetIdx = 1; sheetIdx <= totalSheets; sheetIdx++) {
            await this.buildManualSheet(sheetIdx, totalSheets);
        }
    }

    async buildManualSheet(sheetIdx, totalSheets) {
        const data = this.results;
        const info = data.productInfo;
        const specs = data.specs;
        const maxBatchesPerSheet = 25;
        const cavCount = data.xbarR.summary.n;

        // Constants
        const START_ROW_DATA = 6;
        const sampleStartRow = 7;
        const sampleEndRow = sampleStartRow + cavCount - 1;
        const sumRow = sampleEndRow + 1;
        const meanRow = sumRow + 1;
        const rangeRow = meanRow + 1;
        const causeStartRow = rangeRow + 1;
        const causeEndRow = causeStartRow + 4;

        // Styles
        const borderThin = { style: 'thin', color: { argb: 'FFBFBFBF' } };
        const borderMedium = { style: 'medium', color: { argb: 'FF000000' } };
        const fontNormal = { name: 'Microsoft JhengHei', size: 10 };
        const fontData = { name: 'Times New Roman', size: 10 };
        const fontBold = { name: 'Microsoft JhengHei', size: 10, bold: true };
        const fontTitle = { name: 'Microsoft JhengHei', size: 20, bold: true };
        const fontBlueBold = { name: 'Microsoft JhengHei', size: 10, bold: true, color: { argb: 'FF0000FF' } };
        const fontRed = { name: 'Microsoft JhengHei', size: 10, color: { argb: 'FFFF0000' } };
        const alignCenter = { vertical: 'middle', horizontal: 'center' };
        const applyFullBorder = (cell) => { cell.border = { top: borderThin, left: borderThin, bottom: borderThin, right: borderThin }; };

        const safeItemName = (info.item || 'Sheet').replace(/[\\/:*?"<>|\[\]]/g, '_').substring(0, 20);
        const sheetName = `${safeItemName}-${String(sheetIdx).padStart(3, '0')}`;

        const ws = this.workbook.addWorksheet(sheetName, {
            pageSetup: { paperSize: 9, orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0 }
        });

        // 1. Column Widths
        ws.getColumn(1).width = 4.5; ws.getColumn(2).width = 4.5; ws.getColumn(3).width = 6;
        for (let c = 4; c <= 28; c++) ws.getColumn(c).width = 5.2;
        ws.getColumn(28).width = 18;

        // 2. Header
        ws.mergeCells('A1:AB1');
        const titleCell = ws.getCell('A1'); titleCell.value = 'X̄ - R 管制圖'; titleCell.font = fontTitle; titleCell.alignment = alignCenter;

        // 3. Info Block
        const fillHeaderBox = (range, label, value, valBold = false) => {
            if (range[0].includes(':')) ws.mergeCells(range[0]);
            const c1 = ws.getCell(range[0].split(':')[0]); c1.value = label; c1.font = fontBold; c1.alignment = alignCenter;
            if (range[1]) {
                if (range[1].includes(':')) ws.mergeCells(range[1]);
                const c2 = ws.getCell(range[1].split(':')[0]); c2.value = value; c2.font = valBold ? fontBold : fontNormal; c2.alignment = alignCenter;
            }
        };

        fillHeaderBox(['A2:B2', 'C2:F2'], '產品名稱', info.name);
        fillHeaderBox(['G2:H2', 'I2:J2'], '規 格', '標準', true);
        fillHeaderBox(['K2:L2', 'M2'], '管制圖', 'X圖', true);
        ws.getCell('N2').value = 'R圖'; ws.getCell('N2').font = fontBold; ws.getCell('N2').alignment = alignCenter;
        fillHeaderBox(['O2:P2', 'Q2:R2'], '製造部門', info.dept || '射出課');
        ws.mergeCells('S2:T5'); ws.getCell('S2').value = '期限'; ws.getCell('S2').alignment = alignCenter; ws.getCell('S2').font = fontBold;
        ws.mergeCells('U2:X5'); // Date Range
        ws.mergeCells('Y2:AB2'); ws.getCell('Y2').value = '管制圖編號'; ws.getCell('Y2').alignment = alignCenter; ws.getCell('Y2').font = fontBold;

        fillHeaderBox(['A3:B3', 'C3:F3'], '產品料號', info.item);
        fillHeaderBox(['G3:H3', 'I3:J3'], '最大值', specs.usl);
        fillHeaderBox(['K3:L3', 'M3'], '上 限', this.round(data.xbarR.xBar.UCL));
        ws.getCell('N3').value = this.round(data.xbarR.R.UCL); ws.getCell('N3').alignment = alignCenter; ws.getCell('N3').font = fontNormal;
        fillHeaderBox(['O3:P3', 'Q3:R3'], '檢驗單位', '品管組');
        ws.mergeCells('Y3:AB5'); ws.getCell('Y3').value = sheetName; ws.getCell('Y3').alignment = alignCenter; ws.getCell('Y3').font = { ...fontBold, size: 14 };

        fillHeaderBox(['A4:B4', 'C4:F4'], '測量單位', info.unit);
        fillHeaderBox(['G4:H4', 'I4:J4'], '目標值', specs.target);
        fillHeaderBox(['K4:L4', 'M4'], '中心值', this.round(data.xbarR.xBar.CL));
        ws.getCell('N4').value = this.round(data.xbarR.R.CL); ws.getCell('N4').alignment = alignCenter; ws.getCell('N4').font = fontNormal;
        ws.mergeCells('O4:P5'); ws.getCell('O4').value = '檢驗人員'; ws.getCell('O4').alignment = alignCenter; ws.getCell('O4').font = fontBold;
        ws.mergeCells('Q4:R5'); ws.getCell('Q4').value = info.inspector || ''; ws.getCell('Q4').alignment = alignCenter; ws.getCell('Q4').font = fontNormal;

        fillHeaderBox(['A5:B5', 'C5:F5'], '管制特性', '平均值/全距');
        fillHeaderBox(['G5:H5', 'I5:J5'], '最小值', specs.lsl);
        fillHeaderBox(['K5:L5', 'M5'], '下 限', this.round(data.xbarR.xBar.LCL));
        ws.getCell('N5').value = '-'; ws.getCell('N5').alignment = alignCenter; ws.getCell('N5').font = fontNormal;

        // Date Range
        const startBIdx = (sheetIdx - 1) * maxBatchesPerSheet;
        const endBIdx = Math.min(sheetIdx * maxBatchesPerSheet, data.batchNames.length) - 1;
        const bStart = data.batchNames[startBIdx] || '';
        const bEnd = data.batchNames[endBIdx] || '';
        ws.getCell('U2').value = `${bStart} \n到\n${bEnd} `;
        ws.getCell('U2').alignment = { wrapText: true, vertical: 'middle', horizontal: 'center' }; ws.getCell('U2').font = fontNormal;

        // Header Borders
        for (let r = 2; r <= 5; r++) { for (let c = 1; c <= 28; c++) applyFullBorder(ws.getCell(r, c)); }

        // 4. Data Layout Labels
        ws.mergeCells('A6:B6'); ws.getCell('A6').value = '檢驗日期'; ws.getCell('A6').font = fontBold; ws.getCell('A6').alignment = alignCenter;
        ws.mergeCells(`A${sampleStartRow}:A${sampleEndRow}`);
        const sl = ws.getCell(`A${sampleStartRow}`); sl.value = '樣本測定值'; sl.font = fontBlueBold; sl.alignment = { textRotation: 90, vertical: 'middle', horizontal: 'center' };

        for (let i = 0; i < cavCount; i++) {
            const cell = ws.getCell(sampleStartRow + i, 2); cell.value = 'X' + (i + 1); cell.font = fontBlueBold; cell.alignment = alignCenter;
        }
        ['ΣX', 'X̄', 'R'].forEach((lbl, idx) => {
            const r = sumRow + idx; ws.mergeCells(`A${r}:B${r}`);
            const c = ws.getCell(`A${r}`); c.value = lbl; c.font = fontBold; c.alignment = alignCenter;
        });
        ws.mergeCells(`A${causeStartRow}:B${causeEndRow}`);
        const cl = ws.getCell(`A${causeStartRow}`); cl.value = '原因追查'; cl.font = { ...fontBold, color: { argb: 'FF0000FF' } }; cl.alignment = { textRotation: 90, vertical: 'middle', horizontal: 'center' };


        // 5. Fill Data & Gridlines
        for (let r = 6; r <= causeEndRow; r++) {
            for (let c = 1; c <= 28; c++) {
                const cell = ws.getCell(r, c); applyFullBorder(cell);
                if (!cell.font) cell.font = fontNormal;
                if (c <= 2 || c === 28) continue; // Skip A/B cols (labels) and AB col (summary)

                const batchOffset = c - 3; // C is col 3, so offset 0 for first batch
                const batchIndex = startBIdx + batchOffset;
                if (batchIndex > endBIdx) continue; // No more batches for this sheet/column

                if (r === 6) {
                    cell.value = data.batchNames[batchIndex]; cell.alignment = alignCenter; cell.font = fontNormal;
                } else if (r >= sampleStartRow && r <= sampleEndRow) {
                    const cavIdx = r - sampleStartRow;
                    const val = data.dataMatrix[batchIndex][cavIdx];
                    if (val !== null && val !== undefined) { cell.value = val; cell.numFmt = '0.0000'; cell.alignment = alignCenter; cell.font = fontData; }
                } else if (r === sumRow) {
                    let sum = 0, count = 0;
                    data.dataMatrix[batchIndex].forEach(v => { if (v != null) { sum += v; count++; } });
                    if (count > 0) { cell.value = sum; cell.numFmt = '0.0000'; cell.alignment = alignCenter; cell.font = fontData; }
                } else if (r === meanRow) {
                    const val = data.xbarR.xBar.data[batchIndex]; cell.value = val; cell.numFmt = '0.0000'; cell.alignment = alignCenter;
                    if (val > data.xbarR.xBar.UCL || val < data.xbarR.xBar.LCL) { cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } }; cell.font = { ...fontData, color: { argb: 'FFFF0000' } }; }
                    else { cell.font = fontData; }
                } else if (r === rangeRow) {
                    const val = data.xbarR.R.data[batchIndex]; cell.value = val; cell.numFmt = '0.0000'; cell.alignment = alignCenter;
                    if (val > data.xbarR.R.UCL) { cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } }; cell.font = { ...fontData, color: { argb: 'FFFF0000' } }; }
                    else { cell.font = fontData; }
                }
            }
        }

        // 6. Right Side Summary (Col AB) - Calculate LOCAL stats for this sheet
        // Filter dataMatrix for current sheet's batches
        let sheetSumX = 0;
        let sheetSumR = 0;
        let sheetTotalSubgroups = 0;

        for (let b = startBIdx; b <= endBIdx; b++) {
            // xBar data for this batch
            const xb = data.xbarR.xBar.data[b];
            const r = data.xbarR.R.data[b];
            if (xb !== null && r !== null) {
                sheetSumX += xb;
                sheetSumR += r;
                sheetTotalSubgroups++;
            }
        }

        // Calculate averages for this sheet
        const sheetXDoubleBar = sheetTotalSubgroups > 0 ? sheetSumX / sheetTotalSubgroups : 0;
        const sheetRBar = sheetTotalSubgroups > 0 ? sheetSumR / sheetTotalSubgroups : 0;
        const k = data.xbarR.summary.k; // Use global n/k constants

        const addSummary = (r, text) => {
            ws.mergeCells(r, 28, r + 1, 28); const cell = ws.getCell(r, 28); cell.value = text;
            cell.font = { name: 'Times New Roman', size: 10, bold: true }; cell.alignment = { vertical: 'middle', horizontal: 'left', wrapText: true }; cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDCE6F1' } };
        };

        addSummary(7, `ΣX̄ = ${this.round(sheetSumX)}`); // Sum of Averages
        addSummary(9, `X̿ = ${this.round(sheetXDoubleBar)}`); // Average of Averages
        addSummary(11, `ΣR = ${this.round(sheetSumR)}`); // Sum of Ranges
        addSummary(13, `R̄ = ${this.round(sheetRBar)}`);
        for (let r = 7; r <= 14; r++) { applyFullBorder(ws.getCell(r, 28)); }

        // 7. Outline Border
        for (let c = 1; c <= 28; c++) ws.getCell(2, c).border = { ...ws.getCell(2, c).border, top: borderMedium };
        for (let c = 1; c <= 28; c++) ws.getCell(causeEndRow, c).border = { ...ws.getCell(causeEndRow, c).border, bottom: borderMedium };
        for (let r = 2; r <= causeEndRow; r++) ws.getCell(r, 1).border = { ...ws.getCell(r, 1).border, left: borderMedium };
        for (let r = 2; r <= causeEndRow; r++) ws.getCell(r, 28).border = { ...ws.getCell(r, 28).border, right: borderMedium };
        for (let c = 1; c <= 28; c++) ws.getCell(5, c).border = { ...ws.getCell(5, c).border, bottom: borderMedium };

        // 8. Charts (Removed by user request)
        /*
        const chartRow = causeEndRow + 2;
        if (this.chartsImages.xbar) {
            const imageId = this.workbook.addImage({ base64: this.chartsImages.xbar, extension: 'png' });
            ws.addImage(imageId, { tl: { col: 0, row: chartRow }, ext: { width: 900, height: 350 } });
        }
        if (this.chartsImages.r) {
            const imageId = this.workbook.addImage({ base64: this.chartsImages.r, extension: 'png' });
            ws.addImage(imageId, { tl: { col: 0, row: chartRow + 17 }, ext: { width: 900, height: 350 } });
        }
        */

        // 9. Footer
        const footerRow = causeEndRow + 5; // Adjusted since charts are removed
        const f1 = ws.getCell(footerRow, 1); f1.value = 'MOULDEX QE20002-R01.A'; f1.font = fontBold;
        const f2 = ws.getCell(footerRow, 28); f2.value = `第 ${sheetIdx} 頁，共 ${totalSheets} 頁`; f2.font = fontBold; f2.alignment = { horizontal: 'right' };

        // 10. Auto-fit Columns (C to AB) based on content
        // Focus on adjusting width for Batch Names (Row 6) and Summary (Col 28)
        // avoiding disruption of the header layout (Rows 1-5 merged cells)
        for (let c = 3; c <= 28; c++) {
            let maxLen = 5; // Min default
            const col = ws.getColumn(c);

            // Check specific rows that drive content width
            // Row 6: Batch Names / Date Header
            const val6 = ws.getCell(6, c).value;
            if (val6) {
                // Count double-byte characters as 2 length for better fitting
                const str = val6.toString();
                let len = 0;
                for (let i = 0; i < str.length; i++) {
                    len += (str.charCodeAt(i) > 255) ? 2 : 1.1;
                }
                maxLen = Math.max(maxLen, len);
            }

            // Special Case: Column AB (28) - Summary Stats
            if (c === 28) {
                for (let r = 7; r <= 15; r++) {
                    const v = ws.getCell(r, c).value;
                    if (v) {
                        const str = v.toString();
                        let len = 0;
                        for (let i = 0; i < str.length; i++) {
                            len += (str.charCodeAt(i) > 255) ? 2 : 1.1;
                        }
                        maxLen = Math.max(maxLen, len);
                    }
                }
                maxLen += 2; // Extra padding
            }

            // Apply width
            let newWidth = maxLen + 2;
            if (newWidth < 6) newWidth = 6; // Minimum constraint
            if (newWidth > 40) newWidth = 40; // Maximum constraint

            col.width = newWidth;
        }
    }

    generateSimpleReport() {
        const ws = this.workbook.addWorksheet('Report');
        ws.getCell('A1').value = 'Detailed export only available for Batch Analysis currently.';
    }

    round(num) {
        if (typeof num !== 'number') return '-';
        return parseFloat(num.toFixed(4));
    }

    setBorder(ws, rangeStr, isOuterMedium = false) {
        // Simple internal borders helper
        // Parsed explicitly if needed, but for now relies on individual cell styling
    }

    getAllBorders() {
        return {
            top: { style: 'thin', color: { argb: 'FFBFBFBF' } },
            left: { style: 'thin', color: { argb: 'FFBFBFBF' } },
            bottom: { style: 'thin', color: { argb: 'FFBFBFBF' } },
            right: { style: 'thin', color: { argb: 'FFBFBFBF' } }
        };
    }
}

window.SPCExcelBuilder = SPCExcelBuilder;
