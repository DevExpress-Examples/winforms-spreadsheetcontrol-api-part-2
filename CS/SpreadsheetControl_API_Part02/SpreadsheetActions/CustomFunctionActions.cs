using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
#region #usings_CFunc
using DevExpress.Spreadsheet;
using DevExpress.Spreadsheet.Functions;
using DevExpress.XtraSpreadsheet;
#endregion #usings_CFunc

namespace SpreadsheetControl_API
{
    public static class CustomFunctionActions
    {
        #region Actions
        public static Action<IWorkbook> SphereMassAction = SphereMassValue;
        #endregion

        static void SphereMassValue(IWorkbook workbook)
        {
            #region #customfunctionuse
// Create a custom function and add it to the global scope.
SphereMassFunction customFunction = new SphereMassFunction();
if (!workbook.CustomFunctions.Contains(customFunction.Name))
    workbook.CustomFunctions.Add(customFunction);
            #endregion #customfunctionuse

            workbook.BeginUpdate();
            try
            {
                Worksheet worksheet = workbook.Worksheets[0];
                worksheet.Range["A1:H1"].ColumnWidthInCharacters = 12;
                worksheet.Range["A1:H1"].Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center;

                worksheet.DefinedNames.Add("seawater", "1025");
                worksheet.DefinedNames.Add("iron", "7870");
                worksheet.DefinedNames.Add("gold", "19300");

                worksheet["A1"].Value = "Radius, m";
                worksheet["B1"].Value = "Material";
                worksheet["C1"].Value = "Mass, kg";
                worksheet["A2"].Value = 0.1;
                worksheet["B2"].Value = "";
                worksheet["C2"].Formula = "=SPHEREMASS(A2)";
                worksheet["C2"].NumberFormat = "#.##";
                worksheet["A3"].Value = 0.1;
                worksheet["B3"].Value = "Seawater";
                worksheet["C3"].Formula = "=SPHEREMASS(A3,seawater)";
                worksheet["C3"].NumberFormat = "#.##";
                worksheet["A4"].Value = 0.1;
                worksheet["B4"].Value = "Iron";
                worksheet["C4"].Formula = "=SPHEREMASS(A4,iron)";
                worksheet["C4"].NumberFormat = "#.##";
                worksheet["A5"].Value = 0.1;
                worksheet["B5"].Value = "Gold";
                worksheet["C5"].Formula = "=SPHEREMASS(A5,gold)";
                worksheet["C5"].NumberFormat = "#.##";
            }
            finally
            {
                workbook.EndUpdate();
            }

        }

    }


#region #customfunctiondef
// Inheritance from Object is required for automatic VB.NET conversion
public class SphereMassFunction : Object, ICustomFunction
{
    const string functionName = "SPHEREMASS";
    readonly ParameterInfo[] functionParameters;
 
    public SphereMassFunction()
    {   
        // Missing optional parameters do not result in an error message.
        this.functionParameters = new ParameterInfo[] { new ParameterInfo(ParameterType.Value, ParameterAttributes.Required), 
            new ParameterInfo(ParameterType.Value, ParameterAttributes.Optional)};
    }

    public string Name { get { return functionName; } }
    ParameterInfo[] IFunction.Parameters { get { return functionParameters; } }
    ParameterType IFunction.ReturnType { get { return ParameterType.Value; } }
    // Reevaluate cells on every recalculation.
    bool IFunction.Volatile { get { return true; } }

    ParameterValue IFunction.Evaluate(IList<ParameterValue> parameters, EvaluationContext context)
    {
        double radius;
        double density = 1000;
        ParameterValue radiusParameter;
        ParameterValue densityParameter;

        if (parameters.Count == 2)
        {
            densityParameter = parameters[1];
            if (densityParameter.IsError)
                return densityParameter;
            else 
                density = densityParameter.NumericValue;                
        }
                    
        radiusParameter = parameters[0];
        if (radiusParameter.IsError)
            return radiusParameter;
        else
            radius = radiusParameter.NumericValue;

        return (4 * Math.PI) / 3 * Math.Pow(radius,3) * density;

    }
    string IFunction.GetName(CultureInfo culture)
    {
        return functionName;
    }
}
#endregion #customfunctiondef
}
