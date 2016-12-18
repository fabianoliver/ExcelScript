using ExcelDna.Integration;
using ExcelDna.Registration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using Excel = NetOffice.ExcelApi;

namespace ExcelScript.Registration
{
    // based on https://github.com/Excel-DNA/Registration/blob/master/Source/ExcelDna.Registration.VisualBasic/RangeParameterConversion.vb
    public static class RangeParameterConversion
    {
        public static IEnumerable<ExcelFunctionRegistration> UpdateRegistrationForRangeParameters(this IEnumerable<ExcelFunctionRegistration> reg)
        {
            return reg.Select(UpdateAttributesForRangeParameters);
        }

        public static Expression<Func<object, Excel.Range>> ParameterConversion(Type paramType, ExcelParameterRegistration paramRegistration)
        {
            if (paramType == typeof(Excel.Range))
                return (Expression<Func<object, Excel.Range>>)((object input) => ReferenceToRange(input));
            else
                return null;
        }

        private static ExcelFunctionRegistration UpdateAttributesForRangeParameters(ExcelFunctionRegistration reg)
        {
            var rangeParams = reg.FunctionLambda.Parameters.Select((x, i) => new { _Parameter = x, _Index = i })
                .Where(x => x._Parameter.Type == typeof(Excel.Range));
             
            bool hasRangeParam = false;
            foreach(var param in rangeParams)
            {
                reg.ParameterRegistrations[param._Index].ArgumentAttribute.AllowReference = true;
                hasRangeParam = true;
            }

            if (hasRangeParam)
                reg.FunctionAttribute.IsMacroType = true;

            return reg;
        }

        private static Excel.Range ReferenceToRange(object input)
        {
            ExcelReference xlRef = (ExcelReference)input;
            return ExcelScriptAddin.FromExcelReferenceToRange(xlRef);
        }
    }
}
