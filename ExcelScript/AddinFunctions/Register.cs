using ExcelDna.Integration;
using ExcelDna.Registration;
using ExcelScript.Registration;
using ObjectStorage.Abstractions;
using ScriptingAbstractions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Runtime.InteropServices;
using Excel = NetOffice.ExcelApi;

namespace ExcelScript
{
    public partial class ExcelScriptAddin
    {
        private static readonly IDictionary<string, int> _RegisteredFunctions = new Dictionary<string, int>();  // key: name of registered function, value: script hash

        [ExcelFunction(Name = FunctionPrefix + nameof(Register), IsMacroType = true, Description = "Registers this script as a function with Excel, so it can be called like a normal formula", IsVolatile = false)]
        [SuppressInDialog]
        public static object Register(
        [ExcelArgument("Handle to the script which shall be registered as a UDF")] string ScriptHandle)
        {
            IStoredObject<IScript> storedScript;
            IScript script = GetFromStoreOrThrow<IScript>(ScriptHandle, out storedScript);
            string FunctionName = storedScript.Name;
            int scriptHash = script.GetHashCode();

            if (_RegisteredFunctions.ContainsKey(FunctionName) && _RegisteredFunctions[FunctionName] == scriptHash)
                return $"Function {FunctionName} with same signature already registered.";
          
            var reg = ToExcelFunctionRegistration(script, FunctionName);
            ProcessAndRegisterFunction(reg);

            _RegisteredFunctions[FunctionName] = scriptHash;

            ExcelDna.IntelliSense.IntelliSenseServer.Refresh();

            return "Registered.";
        }

        [ExcelFunction]
        public static object __DEBUG__TEXT(Excel.Range r)
        {
            return r.Address;
        }

        private static ExcelFunctionRegistration ToExcelFunctionRegistration(IScript script, string FunctionName)
        {
            var paramExprs = script
                        .Parameters
                        .Select(x => Expression.Parameter(GetExcelRegistrationTypeFor(x), x.Name))
                        .ToList();

            var methodInfo = typeof(ExcelScriptAddin).GetMethod(nameof(InternalRun), BindingFlags.Static | BindingFlags.NonPublic);

            var paramsToObj = paramExprs.Select(x => Expression.Convert(x, typeof(object))); // cast parameter to object, otherwise won't match function signature of ExcelScriptAddin.Run
            var paramsArray = Expression.NewArrayInit(typeof(object), paramsToObj.ToArray());
            var methodArgs = new Expression[] { Expression.Constant(script, typeof(IScript)), paramsArray };

            LambdaExpression lambdaExpression = Expression.Lambda(Expression.Call(methodInfo, methodArgs), FunctionName, paramExprs);

            ExcelFunctionAttribute excelFunctionAttribute = new ExcelFunctionAttribute() { Name = FunctionName };

            if (!String.IsNullOrEmpty(script.Description))
                excelFunctionAttribute.Description = script.Description;

            IEnumerable<ExcelParameterRegistration> paramRegistrations = script
                .Parameters
                .Select((IParameter x) =>
                {
                    var argumentAttribute = new ExcelArgumentAttribute() { Name = x.Name };
                    var parameterRegistration = new ExcelParameterRegistration(argumentAttribute);

                    if (x.Type == typeof(Excel.Range) || x.Type == typeof(ExcelReference))
                        argumentAttribute.AllowReference = true;

                    if (x.IsOptional)
                    {
                        var optionalAttribute = new OptionalAttribute();
                        parameterRegistration.CustomAttributes.Add(optionalAttribute);

                        var defaultValueAttribute = new DefaultParameterValueAttribute(x.DefaultValue);
                        parameterRegistration.CustomAttributes.Add(defaultValueAttribute);
                    }

                    if (!String.IsNullOrEmpty(x.Description))
                    {
                        argumentAttribute.Description = x.Description;
                    }

                    return parameterRegistration;
                });

            var reg = new ExcelFunctionRegistration(lambdaExpression, excelFunctionAttribute, paramRegistrations);
            return reg;
        }


        private static Type GetExcelRegistrationTypeFor(IParameter parameter)
        {
            if (parameter.IsOptional || !parameter.Type.IsValueType)
                return parameter.Type;
            else
                // If we have a non-optional value-type argument, we need to wrap it in Nullable<>
                // to make sure Excel can actually pass an ExcelMissing or ExcelEmpty argument to the parameter
                // (which will then be picked up by .Run(), which in turn will throw an Exception for missing parameters)
                return typeof(Nullable<>).MakeGenericType(parameter.Type);
        }

        private static object InternalRun(IScript script, params object[] ParameterValues)
        {
            if (script == null)
                throw new ArgumentNullException(nameof(script));

            if (ParameterValues == null)
                throw new ArgumentNullException(nameof(ParameterValues));

            // At this point, ParameterValues should have gone through the type converters, so now we can check for type consistency too
            ParameterValues = CheckParameterConsistency(script.Parameters, true, ParameterValues);

            var parameterValues = script.Parameters.Select((_param, _i) => SelectParameterValueFor(_param, ParameterValues[_i])).ToArray();
            var result = script.Run(parameterValues);

            if (result.IsSuccess)
            {
                return result.ReturnValue;
            }
            else
            {
                throw new AggregateException("There were errors running the script", result.Errors);
            }
        }

        private static IParameterValue SelectParameterValueFor(IParameter parameter, object ValueCandidate)
        {
            return parameter.WithValue(ValueCandidate);
        }
    }
}
