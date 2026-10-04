# ExcelTools.Data.cs

## `SetDataValidation` - `"checkbox"` case

XlDVType enum verification (via reflection against this machine's
Microsoft.Office.Interop.Excel PIA) shows 8 total validation kinds:
xlValidateInputOnly, xlValidateWholeNumber, xlValidateDecimal, xlValidateList,
xlValidateDate, xlValidateTime, xlValidateTextLength, xlValidateCustom. None of
these map to boolean-checkbox cells. The assembly does define CheckBox and
CheckBoxes types, but they are form controls (accessed via Shapes.AddFormControl),
not Data Validation options. Thus, Excel's native checkbox-cell feature (if it
exists in newer Office 365 builds) is not accessible through the Validation API
in this Interop version.
