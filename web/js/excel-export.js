// ============================================================================
// Excel Export Module - Generate Excel with Data Tables and Chart Images
// ============================================================================

export class ExcelExport {
    constructor(analysisResults, itemName, language = 'zh') {
        this.results = analysisResults;
        this.itemName = itemName;
        this.lang = language;
    }

    async generate() {
        const workbook = new ExcelJS.Workbook();
        workbook.creator = 'Mouldex Control Chart';
        workbook.created = new Date();

        switch (this.results.type) {
            case 'batch':
                await this.generateBatchExcel(workbook);
                break;
            case 'cavity':
                await this.generateCavityExcel(workbook);
                break;
            case 'group':
                await this.generateGroupExcel(workbook);
                break;
        }

        // Download
        const buffer = await workbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = `SPC_${this.results.type}_${this.itemName}_${new Date().toISOString().split('T')[0]}.xlsx`;
        link.click();
    }

    // ========================================================================
    // Batch Analysis Excel
    // ========================================================================

    async generateBatchExcel(workbook) {
        const t = (zh, en) => this.lang === 'zh' ? zh : en;
        const data = this.results.data;

        const worksheet = workbook.addWorksheet(this.itemName);

        // Title
        worksheet.mergeCells('A1:H1');
        const titleCell = worksheet.getCell('A1');
        titleCell.value = t('X̄-R 管制圖', 'X-Bar R Control Chart');
        titleCell.font = { size: 16, bold: true, color: { argb: 'FF0A1628' } };
        titleCell.alignment = { horizontal: 'center', vertical: 'middle' };
        worksheet.getRow(1).height = 30;

        // Specifications
        worksheet.getCell('A3').value = t('規格', 'Specifications');
        worksheet.getCell('A3').font = { bold: true };
        worksheet.getCell('B3').value = 'Target';
        worksheet.getCell('C3').value = 'USL';
        worksheet.getCell('D3').value = 'LSL';

        worksheet.getCell('B4').value = data.specs.target;
        worksheet.getCell('C4').value = data.specs.usl;
        worksheet.getCell('D4').value = data.specs.lsl;

        // Control Limits
        worksheet.getCell('A6').value = t('管制界限', 'Control Limits');
        worksheet.getCell('A6').font = { bold: true };

        worksheet.getCell('A7').value = t('圖表', 'Chart');
        worksheet.getCell('B7').value = 'UCL';
        worksheet.getCell('C7').value = 'CL';
        worksheet.getCell('D7').value = 'LCL';

        worksheet.getCell('A8').value = 'X̄';
        worksheet.getCell('B8').value = data.xbarR.xBar.UCL;
        worksheet.getCell('C8').value = data.xbarR.xBar.CL;
        worksheet.getCell('D8').value = data.xbarR.xBar.LCL;

        worksheet.getCell('A9').value = 'R';
        worksheet.getCell('B9').value = data.xbarR.R.UCL;
        worksheet.getCell('C9').value = data.xbarR.R.CL;
        worksheet.getCell('D9').value = data.xbarR.R.LCL;

        // Data table
        worksheet.getCell('A11').value = t('數據', 'Data');
        worksheet.getCell('A11').font = { bold: true };

        worksheet.getCell('A12').value = t('批號', 'Batch');
        worksheet.getCell('B12').value = 'X̄';
        worksheet.getCell('C12').value = 'R';

        // Fill data
        let row = 13;
        data.batchNames.forEach((batch, index) => {
            if (index < data.xbarR.xBar.data.length) {
                worksheet.getCell(`A${row}`).value = batch;
                worksheet.getCell(`B${row}`).value = data.xbarR.xBar.data[index];
                worksheet.getCell(`C${row}`).value = data.xbarR.R.data[index];
                row++;
            }
        });

        // Style headers
        const headerStyle = {
            font: { bold: true, color: { argb: 'FFFFFFFF' } },
            fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0A1628' } },
            alignment: { horizontal: 'center', vertical: 'middle' }
        };

        ['A7', 'B7', 'C7', 'D7', 'A12', 'B12', 'C12'].forEach(cell => {
            worksheet.getCell(cell).style = headerStyle;
        });

        // Add charts as images
        await this.addChartsToWorksheet(worksheet, row + 2);

        // Auto-fit columns
        worksheet.columns.forEach(column => {
            column.width = 15;
        });
    }

    // ========================================================================
    // Cavity Analysis Excel
    // ========================================================================

    async generateCavityExcel(workbook) {
        const t = (zh, en) => this.lang === 'zh' ? zh : en;
        const data = this.results.data;

        const worksheet = workbook.addWorksheet(this.itemName);

        // Title
        worksheet.mergeCells('A1:L1');
        const titleCell = worksheet.getCell('A1');
        titleCell.value = t('模穴分析', 'Cavity Analysis');
        titleCell.font = { size: 16, bold: true, color: { argb: 'FF0A1628' } };
        titleCell.alignment = { horizontal: 'center', vertical: 'middle' };
        worksheet.getRow(1).height = 30;

        // Specifications
        worksheet.getCell('A3').value = 'Target';
        worksheet.getCell('B3').value = data.specs.target;
        worksheet.getCell('C3').value = 'USL';
        worksheet.getCell('D3').value = data.specs.usl;
        worksheet.getCell('E3').value = 'LSL';
        worksheet.getCell('F3').value = data.specs.lsl;

        // Headers
        const headers = [
            t('模穴', 'Cavity'),
            t('平均值', 'Mean'),
            t('組內標準差', 'Within StdDev'),
            t('整體標準差', 'Overall StdDev'),
            t('範圍', 'Range'),
            t('最大值', 'Max'),
            t('最小值', 'Min'),
            'Cp',
            'Cpk',
            'Pp',
            'Ppk',
            t('樣本數', 'Count')
        ];

        let col = 1;
        headers.forEach(header => {
            worksheet.getCell(5, col).value = header;
            worksheet.getCell(5, col).font = { bold: true, color: { argb: 'FFFFFFFF' } };
            worksheet.getCell(5, col).fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FF0A1628' }
            };
            col++;
        });

        // Data
        let row = 6;
        data.cavityStats.forEach(stat => {
            worksheet.getCell(row, 1).value = stat.name;
            worksheet.getCell(row, 2).value = stat.mean;
            worksheet.getCell(row, 3).value = stat.withinStdDev;
            worksheet.getCell(row, 4).value = stat.overallStdDev;
            worksheet.getCell(row, 5).value = stat.range;
            worksheet.getCell(row, 6).value = stat.max;
            worksheet.getCell(row, 7).value = stat.min;
            worksheet.getCell(row, 8).value = stat.Cp;
            worksheet.getCell(row, 9).value = stat.Cpk;
            worksheet.getCell(row, 10).value = stat.Pp;
            worksheet.getCell(row, 11).value = stat.Ppk;
            worksheet.getCell(row, 12).value = stat.count;

            // Color code Cpk
            const cpkCell = worksheet.getCell(row, 9);
            if (stat.Cpk >= 1.67) {
                cpkCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC6EFCE' } };
            } else if (stat.Cpk >= 1.33) {
                cpkCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC6EFCE' } };
            } else if (stat.Cpk >= 1.0) {
                cpkCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFEB9C' } };
            } else {
                cpkCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFC7CE' } };
            }

            row++;
        });

        // Add charts
        await this.addChartsToWorksheet(worksheet, row + 2);

        // Auto-fit
        worksheet.columns.forEach(column => {
            column.width = 14;
        });
    }

    // ========================================================================
    // Group Analysis Excel
    // ========================================================================

    async generateGroupExcel(workbook) {
        const t = (zh, en) => this.lang === 'zh' ? zh : en;
        const data = this.results.data;

        const worksheet = workbook.addWorksheet(this.itemName);

        // Title
        worksheet.mergeCells('A1:I1');
        const titleCell = worksheet.getCell('A1');
        titleCell.value = t('群組分析 (Min-Max-Avg)', 'Group Analysis (Min-Max-Avg)');
        titleCell.font = { size: 16, bold: true, color: { argb: 'FF0A1628' } };
        titleCell.alignment = { horizontal: 'center', vertical: 'middle' };
        worksheet.getRow(1).height = 30;

        // Headers
        const headers = [
            t('批號', 'Batch'),
            t('平均值', 'Avg'),
            t('最大值', 'Max'),
            t('最小值', 'Min'),
            t('全距', 'Range'),
            t('樣本數', 'n'),
            'USL',
            'Target',
            'LSL'
        ];

        let col = 1;
        headers.forEach(header => {
            worksheet.getCell(3, col).value = header;
            worksheet.getCell(3, col).font = { bold: true, color: { argb: 'FFFFFFFF' } };
            worksheet.getCell(3, col).fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FF0A1628' }
            };
            col++;
        });

        // Data
        let row = 4;
        data.groupStats.forEach(stat => {
            worksheet.getCell(row, 1).value = stat.batch;
            worksheet.getCell(row, 2).value = stat.avg;
            worksheet.getCell(row, 3).value = stat.max;
            worksheet.getCell(row, 4).value = stat.min;
            worksheet.getCell(row, 5).value = stat.range;
            worksheet.getCell(row, 6).value = stat.count;
            worksheet.getCell(row, 7).value = data.specs.usl;
            worksheet.getCell(row, 8).value = data.specs.target;
            worksheet.getCell(row, 9).value = data.specs.lsl;
            row++;
        });

        // Add charts
        await this.addChartsToWorksheet(worksheet, row + 2);

        // Auto-fit
        worksheet.columns.forEach(column => {
            column.width = 14;
        });
    }

    // ========================================================================
    // Add Charts as Images
    // ========================================================================

    async addChartsToWorksheet(worksheet, startRow) {
        const canvases = document.querySelectorAll('.chart-container canvas');

        for (let i = 0; i < canvases.length; i++) {
            const canvas = canvases[i];
            const dataUrl = canvas.toDataURL('image/png');
            const base64 = dataUrl.split(',')[1];

            // Add image to workbook
            const imageId = worksheet.workbook.addImage({
                base64: dataUrl,
                extension: 'png',
            });

            // Calculate position (approx 15 rows per chart)
            const rowPosition = startRow + (i * 25);

            // Add image to worksheet
            worksheet.addImage(imageId, {
                tl: { col: 0, row: rowPosition },
                ext: { width: 800, height: 400 }
            });
        }
    }
}
