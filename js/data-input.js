// ============================================================================
// Data Input Module - Excel File Parsing & Validation
// ============================================================================

export class DataInput {
    constructor(worksheet) {
        this.ws = worksheet;
        this.data = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
        this.parse();
    }

    parse() {
        // Row 1: Headers
        this.headers = this.data[0] || [];

        // Row 2: Specifications (Target, USL, LSL)
        this.specs = {
            target: parseFloat(this.data[1][1]) || 0,
            usl: parseFloat(this.data[1][2]) || 0,
            lsl: parseFloat(this.data[1][3]) || 0
        };

        // Identify cavity columns (columns with "穴" in header)
        this.cavityColumns = [];
        this.headers.forEach((header, index) => {
            if (typeof header === 'string' && header.includes('穴')) {
                this.cavityColumns.push({
                    index: index,
                    name: header
                });
            }
        });

        // Row 3+: Data rows
        this.dataRows = this.data.slice(2);

        // Extract batch names from column A
        this.batchNames = this.dataRows.map(row => row[0] || '').filter(name => name !== '');

        // Extract product information (if exists)
        this.productInfo = this.extractProductInfo();
    }

    extractProductInfo() {
        // Check for product name and measurement unit in specific cells
        // B5: ProductName title, B6: ProductName value
        // C5: MeasurementUnit title, C6: MeasurementUnit value
        const info = {
            productName: '',
            measurementUnit: ''
        };

        try {
            const row5 = this.data[4] || []; // Row 5 (index 4)
            const row6 = this.data[5] || []; // Row 6 (index 5)

            if (row5[1] === 'ProductName' && row6[1]) {
                info.productName = row6[1];
            }

            if (row5[2] === 'MeasurementUnit' && row6[2]) {
                info.measurementUnit = row6[2];
            }
        } catch (error) {
            console.warn('Could not extract product info:', error);
        }

        return info;
    }

    // Get all cavity data for a specific batch
    getBatchCavityData(batchIndex) {
        const row = this.dataRows[batchIndex];
        if (!row) return [];

        return this.cavityColumns.map(col => {
            const value = parseFloat(row[col.index]);
            return isNaN(value) ? null : value;
        }).filter(v => v !== null);
    }

    // Get all data for a specific cavity across all batches
    getCavityBatchData(cavityIndex) {
        const column = this.cavityColumns[cavityIndex];
        if (!column) return [];

        return this.dataRows.map(row => {
            const value = parseFloat(row[column.index]);
            return isNaN(value) ? null : value;
        }).filter(v => v !== null);
    }

    // Get complete data matrix [batch][cavity]
    getDataMatrix() {
        const matrix = [];

        this.dataRows.forEach((row, batchIdx) => {
            const batchData = [];
            this.cavityColumns.forEach(col => {
                const value = parseFloat(row[col.index]);
                batchData.push(isNaN(value) ? null : value);
            });
            matrix.push(batchData);
        });

        return matrix;
    }

    // Validation
    validate() {
        const errors = [];

        if (this.cavityColumns.length === 0) {
            errors.push('No cavity columns found (headers must contain "穴")');
        }

        if (this.batchNames.length === 0) {
            errors.push('No batch data found');
        }

        if (this.specs.usl === 0 && this.specs.lsl === 0) {
            errors.push('No specifications found (USL/LSL)');
        }

        return {
            valid: errors.length === 0,
            errors: errors
        };
    }

    // Getters
    getCavityCount() {
        return this.cavityColumns.length;
    }

    getBatchCount() {
        return this.batchNames.length;
    }

    getCavityNames() {
        return this.cavityColumns.map(col => col.name);
    }

    getSpecs() {
        return this.specs;
    }

    getProductInfo() {
        return this.productInfo;
    }
}
