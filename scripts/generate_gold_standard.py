import xlsxwriter
import os
import statistics

def generate_gold_standard_v2():
    output_path = r'c:\Users\USER\Downloads\QIP-to-ControlChart\docs\validation\Validation_Final_Gold_Standard.xlsx'
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    
    # 規格 (Spec) - 寫死以利稽核
    TARGET = 1.5000
    USL = 1.6000
    LSL = 1.4000
    
    # 數據設計 (Lot 01-11, 每批 5 穴)
    # 為了確保「標準答案」的可控性，我們使用固定數值序列 (但包含微小正態偽隨機分佈)
    # Lot 01-02: 近似 1.500
    # Lot 03-11: 近似 1.530 (觸發 Rule 2: 9點在同一側)
    raw_data = [
        [1.5020, 1.4980, 1.5010, 1.4990, 1.5000], # Lot 01 (Avg: 1.5000)
        [1.5050, 1.4950, 1.5020, 1.4980, 1.5000], # Lot 02 (Avg: 1.5000)
        [1.5320, 1.5280, 1.5310, 1.5290, 1.5300], # Lot 03 (Avg: 1.5300) -> Start Shift
        [1.5350, 1.5250, 1.5320, 1.5280, 1.5300], # Lot 04
        [1.5310, 1.5290, 1.5300, 1.5310, 1.5310], # Lot 05
        [1.5330, 1.5270, 1.5300, 1.5320, 1.5320], # Lot 06
        [1.5340, 1.5260, 1.5310, 1.5290, 1.5310], # Lot 07
        [1.5320, 1.5280, 1.5300, 1.5310, 1.5300], # Lot 08
        [1.5350, 1.5250, 1.5350, 1.5250, 1.5300], # Lot 09
        [1.5310, 1.5290, 1.5320, 1.5280, 1.5300], # Lot 10
        [1.5320, 1.5280, 1.5310, 1.5290, 1.5300], # Lot 11 (End 9th Shift Point)
    ]

    # --- 計算標準答案 ---
    all_values = [v for subgroup in raw_data for v in subgroup]
    grand_mean = sum(all_values) / len(all_values)
    
    ranges = [max(g) - min(g) for g in raw_data]
    mean_range = sum(ranges) / len(ranges)
    
    # 假設 d2 = 2.326 for n=5 (計算管制界限)
    d2 = 2.326
    sigma_within = mean_range / d2
    
    # 計算 Ppk
    std_total = statistics.stdev(all_values) # 使用樣本標準差
    ppk_u = (USL - grand_mean) / (3 * std_total)
    ppk_l = (grand_mean - LSL) / (3 * std_total)
    ppk = min(ppk_u, ppk_l)

    print("-" * 30)
    print("【標準答案對照表 / Golden Standard Answers】")
    print(f"Grand Mean (x-double-bar): {grand_mean:.6f}")
    print(f"Mean Range (R-bar): {mean_range:.6f}")
    print(f"Sigma (Total): {std_total:.6f}")
    print(f"Ppk: {ppk:.6f}")
    print(f"Rule 2 Status: Should Trigger at Lot-11 (9 points from Lot-03 to Lot-11)")
    print("-" * 30)

    # --- 產出 Excel ---
    workbook = xlsxwriter.Workbook(output_path)
    worksheet = workbook.add_worksheet('DIM-GOLD-2026')
    
    header_fmt = workbook.add_format({'bold': True, 'border': 1, 'bg_color': '#92D050', 'align': 'center'})
    data_fmt = workbook.add_format({'border': 1, 'num_format': '0.0000', 'align': 'center'})
    meta_fmt = workbook.add_format({'bold': True})

    headers = ['Target', 'USL', 'LSL', '生產批號', '1號穴', '2號穴', '3號穴', '4號穴', '5號穴']
    for i, h in enumerate(headers):
        worksheet.write(0, i, h, header_fmt)

    for i, row_values in enumerate(raw_data):
        row = i + 1
        batch_id = f'LOT-{i+1:03d}'
        
        # 規格僅在 Row 2 顯示 (官方匯出格式)
        if row == 1:
            worksheet.write(row, 0, TARGET, data_fmt)
            worksheet.write(row, 1, USL, data_fmt)
            worksheet.write(row, 2, LSL, data_fmt)
        
        worksheet.write(row, 3, batch_id, data_fmt)
        for cav_idx, val in enumerate(row_values):
            worksheet.write(row, 4 + cav_idx, val, data_fmt)

    # Metadata
    worksheet.write('A15', 'ProductName', meta_fmt)
    worksheet.write('B15', 'GOLDEN-CERT-V2')
    worksheet.write('A16', 'MeasurementUnit', meta_fmt)
    worksheet.write('B16', 'mm')

    workbook.close()
    print(f"Excel generated at: {output_path}")

if __name__ == "__main__":
    generate_gold_standard_v2()
