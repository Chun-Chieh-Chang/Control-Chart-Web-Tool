# SPC Tool Validation Execution Guide (SOP)

> [!NOTE]
> This guide is intended for the personnel performing the formal validation of the SPC Analysis Tool according to ISO 80002-2:2017.

## 1. Preparation
Before starting, ensure you have:
- [ ] A stable version of Chrome or Edge browser.
- [ ] Access to [Validation_Gold_Standard.xlsx](file:///c:/Users/USER/Downloads/QIP-to-ControlChart/docs/validation/Validation_Gold_Standard.xlsx).
- [ ] A copy of [IQ_OQ_PQ_Protocols.md](file:///c:/Users/USER/Downloads/QIP-to-ControlChart/docs/validation/IQ_OQ_PQ_Protocols.md) to record results.

## 2. Step-by-Step Execution

### Phase A: IQ (Infrastructure & Loading)
1. **Load System**: Open the tool URL.
2. **Verify Loading**: Check the UI for responsiveness.
3. **Screenshot**: Capture the browser console and the tool's main dashboard showing the version number.
4. **Record**: Mark `PASS` on IQ-01 and IQ-02 in the Protocol.

### Phase B: OQ (Functional & Language)
1. **Manual Entry**: Test the "Individual Entry" feature with a known simple set.
2. **Language Check**: Toggle the language switch at the top right.
   - *Verification*: Ensure the "Analysis Report" title and "Nelson Rules Expert Interpretation" reflect the chosen language without translation errors.
3. **Edge Case**: Use n=2 (minimum) and n=48 (maximum) to verify that constants auto-update via the dynamic table.

### Phase C: PQ (Statistical Reliability)
1. **Import Gold Standard (QIP)**: Drag and drop [Validation_Gold_Standard_QIP.xlsx](file:///c:/Users/USER/Downloads/QIP-to-ControlChart/docs/validation/Validation_Gold_Standard_QIP.xlsx) into the "Import QIP" upload zone.
2. **Range Configuration**: 
   - Ensure the system is set to read **Item Name** from Column B.
   - Set **Cavity ID Range** to `G8:V8`.
   - Set **Data Values Range** to `G12:V13`.
3. **Compare Indices**:
   - Locate the **Process Capability Indicators** card.
   - Compare the Cpk and Ppk values against the values in the Excel `EXPECTED RESULTS` notes.
4. **Verify Nelson Rules**: Ensure that for Subgroup 3 in the Excel file, the system remains stable or flags the correct rule if applicable.

## 3. Gathering Objective Evidence
**What counts as proof for auditors?**
- **Screenshots**: Always include a timestamp in your screenshots (system tray clock).
- **Log Files**: If the browser displays specific errors, save the console log.
- **Signed Protocols**: Print the MD files or sign them digitally once "Actual Results" are filled.

## 4. Final Conclusion
After all tests (IQ/OQ/PQ) are marked as `PASS`, the Validation Team lead should sign the **Validation Summary Report** stating: 
*"The software is fit for its intended use within the medical device quality system."*
