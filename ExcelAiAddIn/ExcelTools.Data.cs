using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using Excel = Microsoft.Office.Interop.Excel;
using OfficeAi.Shared;

namespace ExcelAiAddIn
{
    public static partial class ExcelTools
    {
        private static void SetHyperlink(JsonElement op)
        {
            string address = op.GetProperty("address").GetString();
            Excel.Range range = Sheet(op).Range[address];
            if (!op.TryGetProperty("target", out var target) || target.ValueKind == JsonValueKind.Null)
            {
                foreach (Excel.Hyperlink link in range.Hyperlinks) link.Delete();
                return;
            }
            string url = target.GetString();
            Excel.Worksheet sheet = Sheet(op);
            if (url.Contains("!") && !url.StartsWith("http"))
            {
                sheet.Hyperlinks.Add(range, "", SubAddress: url);
            }
            else
            {
                sheet.Hyperlinks.Add(range, url);
            }
        }

        private static void SetNote(JsonElement op)
        {
            string address = op.GetProperty("address").GetString();
            Excel.Range cell = Sheet(op).Range[address];
            if (!op.TryGetProperty("text", out var text) || text.ValueKind == JsonValueKind.Null)
            {
                cell.Comment?.Delete();
                return;
            }
            cell.Comment?.Delete();
            cell.AddComment(text.GetString());
        }

        // PP-17: Excel rejects a name colliding with a cell address, starting
        // with a digit, or containing a space - with an unhelpful COM error.
        // Pre-check so the model gets a specific reason instead.
        private static readonly System.Text.RegularExpressions.Regex CellAddressLike =
            new System.Text.RegularExpressions.Regex(@"^\$?[A-Za-z]{1,3}\$?[0-9]+$");

        private static void ValidateDefinedNameSyntax(string name)
        {
            if (string.IsNullOrEmpty(name))
                throw new ArgumentException("add_defined_name: name cannot be empty.");
            if (char.IsDigit(name[0]))
                throw new ArgumentException("add_defined_name: name '" + name + "' cannot start with a digit.");
            if (name.IndexOf(' ') >= 0)
                throw new ArgumentException("add_defined_name: name '" + name + "' cannot contain spaces.");
            if (CellAddressLike.IsMatch(name))
                throw new ArgumentException("add_defined_name: name '" + name + "' looks like a cell address, which Excel does not allow as a defined name.");
        }

        private static bool DefinedNameExists(Excel.Names names, string name)
        {
            foreach (Excel.Name n in names)
            {
                if (string.Equals(n.Name, name, StringComparison.OrdinalIgnoreCase)) return true;
            }
            return false;
        }

        private static void AddDefinedName(JsonElement op)
        {
            string name = op.GetProperty("name").GetString();
            string reference = op.GetProperty("ref").GetString();
            bool sheetScoped = op.TryGetProperty("scope", out var sc) && sc.ValueKind == JsonValueKind.String && sc.GetString() == "sheet";
            bool overwrite = op.TryGetProperty("overwrite", out var ow) && ow.ValueKind == JsonValueKind.True;

            ValidateDefinedNameSyntax(name);

            Excel.Worksheet targetSheet = Sheet(op); // honors the existing optional "sheet" property either way
            string refersTo = reference.StartsWith("=") ? reference : "=" + reference;
            // Qualify an unqualified reference to the target sheet - otherwise a
            // workbook-scoped name resolves against whichever sheet is active
            // at evaluation time, a latent wrong-answer bug for both scopes.
            if (refersTo.IndexOf('!') < 0)
            {
                string sheetName = targetSheet.Name;
                string quotedSheet = sheetName.IndexOf(' ') >= 0 ? "'" + sheetName + "'" : sheetName;
                refersTo = "=" + quotedSheet + "!" + refersTo.Substring(1);
            }

            Excel.Names names = sheetScoped ? targetSheet.Names : Globals.ThisAddIn.Application.ActiveWorkbook.Names;
            if (!overwrite && DefinedNameExists(names, name))
                throw new ArgumentException("add_defined_name: a " + (sheetScoped ? "sheet" : "workbook") +
                    "-scoped name '" + name + "' already exists. Pass overwrite:true to replace it.");

            names.Add(name, refersTo);
        }

        private static void DeleteDefinedName(JsonElement op)
        {
            string name = op.GetProperty("name").GetString();
            bool sheetScoped = op.TryGetProperty("scope", out var sc) && sc.ValueKind == JsonValueKind.String && sc.GetString() == "sheet";
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Names names = sheetScoped ? Sheet(op).Names : wb.Names;

            if (!DefinedNameExists(names, name))
            {
                // Point the model at the other scope if the name exists there -
                // turns a dead end into a self-correcting next turn.
                string searchedScope = sheetScoped ? "sheet" : "workbook";
                foreach (Excel.Worksheet sheet in wb.Worksheets)
                {
                    if (!sheetScoped && DefinedNameExists(sheet.Names, name))
                    {
                        throw new ArgumentException("delete_defined_name: no workbook-scoped name '" + name +
                            "' found; a sheet-scoped name with this name exists on '" + sheet.Name +
                            "' - pass scope:'sheet' and sheet:'" + sheet.Name + "'.");
                    }
                }
                if (sheetScoped && DefinedNameExists(wb.Names, name))
                {
                    throw new ArgumentException("delete_defined_name: no sheet-scoped name '" + name +
                        "' found on this sheet; a workbook-scoped name with this name exists - omit scope to target it.");
                }
                throw new ArgumentException("delete_defined_name: no " + searchedScope + "-scoped name '" + name + "' found.");
            }

            names.Item(name).Delete();
        }

        private static void SetDataValidation(JsonElement op)
        {
            string range = op.GetProperty("range").GetString();
            Excel.Range target = Sheet(op).Range[range];

            if (!op.TryGetProperty("validation", out var validation) || validation.ValueKind == JsonValueKind.Null)
            {
                target.Validation.Delete();
                return;
            }

            string kind = validation.GetProperty("kind").GetString();
            target.Validation.Delete();

            switch (kind)
            {
                case "list":
                {
                    var values = new List<string>();
                    foreach (JsonElement v in validation.GetProperty("values").EnumerateArray()) values.Add(v.GetString());
                    target.Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(",", values));
                    break;
                }
                case "listRef":
                {
                    string refRange = validation.GetProperty("range").GetString();
                    target.Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, "=" + refRange);
                    break;
                }
                case "numberBetween":
                {
                    double min = validation.GetProperty("min").GetDouble();
                    double max = validation.GetProperty("max").GetDouble();
                    target.Validation.Add(Excel.XlDVType.xlValidateDecimal, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, min.ToString(), max.ToString());
                    break;
                }
                case "dateBetween":
                {
                    string start = validation.GetProperty("start").GetString();
                    string end = validation.GetProperty("end").GetString();
                    target.Validation.Add(Excel.XlDVType.xlValidateDate, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, start, end);
                    break;
                }
                case "formula":
                {
                    string formula = validation.GetProperty("formula").GetString();
                    target.Validation.Add(Excel.XlDVType.xlValidateCustom, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, formula);
                    break;
                }
                // XlDVType enum verification (via reflection against this machine's
                // Microsoft.Office.Interop.Excel PIA) shows 8 total validation kinds:
                // xlValidateInputOnly, xlValidateWholeNumber, xlValidateDecimal, xlValidateList,
                // xlValidateDate, xlValidateTime, xlValidateTextLength, xlValidateCustom. None of
                // these map to boolean-checkbox cells. The assembly does define CheckBox and
                // CheckBoxes types, but they are form controls (accessed via Shapes.AddFormControl),
                // not Data Validation options. Thus, Excel's native checkbox-cell feature (if it
                // exists in newer Office 365 builds) is not accessible through the Validation API
                // in this Interop version.
                case "checkbox":
                    throw new NotSupportedException("set_data_validation: 'checkbox' kind is not supported in this version of Excel Interop - CheckBox is a form control, not a Data Validation type.");
                default:
                    throw new ArgumentException("set_data_validation: unknown validation kind '" + kind + "'.");
            }
        }

    }
}

