using ExcelDna.Integration;
using ExcelScript.Internal;
using ScriptingAbstractions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Caching;
using Excel = NetOffice.ExcelApi;

namespace ExcelScript
{
    public partial class ExcelScriptAddin
    {
        private static readonly MemoryCache _ExecutiongCache = MemoryCache.Default;

        // todo:
        // consider setting up this function dynamically to allow for more params in newer excel versions (no big problem though, user could just input an array instead)
        [ExcelFunction(Name = ExcelScriptAddin.FunctionPrefix + nameof(Run), Description = "Executes the script defined by the given Script Handle", IsVolatile = false, IsMacroType = true)]
        public static object Run(
            [ExcelArgument(Name = "ScriptHandle", Description = "A stored handled to the script which shall be executed")] string ScriptHandle,
            [ExcelArgument(Name = "Parameter1Value...", Description = "Values to be passed as parameters to the script. These parameters must, in type and count, match the parameter definitions that are expected/defined by the script stored in ScriptHandle.", AllowReference = true)] object Parameter1Value = null,
            [ExcelArgument(Name = "", Description = "2nd value to be passed as parameters to the script. These parameters must, in type and count, match the parameter definitions that are expected/defined by the script stored in ScriptHandle.", AllowReference = true)] object Parameter2Value = null,
            [ExcelArgument(Name = "", Description = "3rd value to be passed as parameters to the script. These parameters must, in type and count, match the parameter definitions that are expected/defined by the script stored in ScriptHandle.", AllowReference = true)] object Parameter3Value = null,
            [ExcelArgument(Name = "", Description = "4th value to be passed as parameters to the script. These parameters must, in type and count, match the parameter definitions that are expected/defined by the script stored in ScriptHandle.", AllowReference = true)] object Parameter4Value = null,
            [ExcelArgument(Name = "", Description = "5th value to be passed as parameters to the script. These parameters must, in type and count, match the parameter definitions that are expected/defined by the script stored in ScriptHandle.", AllowReference = true)] object Parameter5Value = null,
            [ExcelArgument(Name = "", Description = "6th value to be passed as parameters to the script. These parameters must, in type and count, match the parameter definitions that are expected/defined by the script stored in ScriptHandle.", AllowReference = true)] object Parameter6Value = null,
            [ExcelArgument(Name = "", Description = "7th value to be passed as parameters to the script. These parameters must, in type and count, match the parameter definitions that are expected/defined by the script stored in ScriptHandle.", AllowReference = true)] object Parameter7Value = null,
            [ExcelArgument(Name = "", Description = "8th value to be passed as parameters to the script. These parameters must, in type and count, match the parameter definitions that are expected/defined by the script stored in ScriptHandle.", AllowReference = true)] object Parameter8Value = null,
            [ExcelArgument(Name = "", Description = "9th value to be passed as parameters to the script. These parameters must, in type and count, match the parameter definitions that are expected/defined by the script stored in ScriptHandle.", AllowReference = true)] object Parameter9Value = null,
            [ExcelArgument(Name = "", Description = "10th value to be passed as parameters to the script. These parameters must, in type and count, match the parameter definitions that are expected/defined by the script stored in ScriptHandle.", AllowReference = true)] object Parameter10Value = null,
            [ExcelArgument(Name = "", Description = "11th value to be passed as parameters to the script. These parameters must, in type and count, match the parameter definitions that are expected/defined by the script stored in ScriptHandle.", AllowReference = true)] object Parameter11Value = null,
            [ExcelArgument(Name = "", Description = "12th value to be passed as parameters to the script. These parameters must, in type and count, match the parameter definitions that are expected/defined by the script stored in ScriptHandle.", AllowReference = true)] object Parameter12Value = null,
            [ExcelArgument(Name = "", Description = "13th value to be passed as parameters to the script. These parameters must, in type and count, match the parameter definitions that are expected/defined by the script stored in ScriptHandle.", AllowReference = true)] object Parameter13Value = null,
            [ExcelArgument(Name = "", Description = "14th value to be passed as parameters to the script. These parameters must, in type and count, match the parameter definitions that are expected/defined by the script stored in ScriptHandle.", AllowReference = true)] object Parameter14Value = null,
            [ExcelArgument(Name = "", Description = "15th value to be passed as parameters to the script. These parameters must, in type and count, match the parameter definitions that are expected/defined by the script stored in ScriptHandle.", AllowReference = true)] object Parameter15Value = null,
            [ExcelArgument(Name = "", Description = "16th value to be passed as parameters to the script. These parameters must, in type and count, match the parameter definitions that are expected/defined by the script stored in ScriptHandle.", AllowReference = true)] object Parameter16Value = null,
            [ExcelArgument(Name = "", Description = "17th value to be passed as parameters to the script. These parameters must, in type and count, match the parameter definitions that are expected/defined by the script stored in ScriptHandle.", AllowReference = true)] object Parameter17Value = null,
            [ExcelArgument(Name = "", Description = "18th value to be passed as parameters to the script. These parameters must, in type and count, match the parameter definitions that are expected/defined by the script stored in ScriptHandle.", AllowReference = true)] object Parameter18Value = null,
            [ExcelArgument(Name = "", Description = "19th value to be passed as parameters to the script. These parameters must, in type and count, match the parameter definitions that are expected/defined by the script stored in ScriptHandle.", AllowReference = true)] object Parameter19Value = null,
            [ExcelArgument(Name = "", Description = "20th value to be passed as parameters to the script. These parameters must, in type and count, match the parameter definitions that are expected/defined by the script stored in ScriptHandle.", AllowReference = true)] object Parameter20Value = null,
            [ExcelArgument(Name = "", Description = "21st value to be passed as parameters to the script. These parameters must, in type and count, match the parameter definitions that are expected/defined by the script stored in ScriptHandle.", AllowReference = true)] object Parameter21Value = null,
            [ExcelArgument(Name = "", Description = "22nd value to be passed as parameters to the script. These parameters must, in type and count, match the parameter definitions that are expected/defined by the script stored in ScriptHandle.", AllowReference = true)] object Parameter22Value = null,
            [ExcelArgument(Name = "", Description = "23th value to be passed as parameters to the script. These parameters must, in type and count, match the parameter definitions that are expected/defined by the script stored in ScriptHandle.", AllowReference = true)] object Parameter23Value = null,
            [ExcelArgument(Name = "", Description = "24th value to be passed as parameters to the script. These parameters must, in type and count, match the parameter definitions that are expected/defined by the script stored in ScriptHandle.", AllowReference = true)] object Parameter24Value = null,
            [ExcelArgument(Name = "", Description = "25th value to be passed as parameters to the script. These parameters must, in type and count, match the parameter definitions that are expected/defined by the script stored in ScriptHandle.", AllowReference = true)] object Parameter25Value = null,
            [ExcelArgument(Name = "", Description = "26th value to be passed as parameters to the script. These parameters must, in type and count, match the parameter definitions that are expected/defined by the script stored in ScriptHandle.", AllowReference = true)] object Parameter26Value = null,
            [ExcelArgument(Name = "", Description = "27th value to be passed as parameters to the script. These parameters must, in type and count, match the parameter definitions that are expected/defined by the script stored in ScriptHandle.", AllowReference = true)] object Parameter27Value = null,
            [ExcelArgument(Name = "", Description = "28th value to be passed as parameters to the script. These parameters must, in type and count, match the parameter definitions that are expected/defined by the script stored in ScriptHandle.", AllowReference = true)] object Parameter28Value = null
            )
        {
            if (DirtyRangeFlagger.IsRecalculatingDirtyCells)
            {
                XlCall.Excel(XlCall.xlfVolatile, true);
                return "Script will be executed on next recalculation cycle";
            }
            else
            {
                XlCall.Excel(XlCall.xlfVolatile, false);
            }

            var script = GetFromStoreOrThrow<IScript>(ScriptHandle);

            // at this point, we can't check for type consistency yet as the input parameters have not gone through the registered converters yet. Type checking is done in the function lambda -> InternalRun.
            var ParameterValues = CheckParameterConsistency(script.Parameters, false, Parameter1Value, Parameter2Value, Parameter3Value, Parameter4Value, Parameter5Value, Parameter6Value, Parameter7Value, Parameter8Value, Parameter9Value, Parameter10Value, Parameter11Value, Parameter12Value, Parameter13Value, Parameter14Value, Parameter15Value, Parameter16Value, Parameter17Value, Parameter18Value, Parameter19Value, Parameter20Value, Parameter21Value, Parameter22Value, Parameter23Value, Parameter24Value, Parameter25Value, Parameter26Value, Parameter27Value, Parameter28Value);
            var @delegate = GetOrCreateDelegateFor(script);

            var result = @delegate.DynamicInvoke(ParameterValues);
            return result;
        }

        private static Delegate GetOrCreateDelegateFor(IScript script)
        {
            var hash = script.GetHashCode().ToString();

            var cachedDelegate = _ExecutiongCache[hash] as Delegate;

            if (cachedDelegate != null)
                return cachedDelegate;

            var registration = ToExcelFunctionRegistration(script, "INTERNAL-RUN-CALL");
            var processedRegistration = ProcessFunction(registration);

            var @delegate = processedRegistration.FunctionLambda.Compile();
            _ExecutiongCache.Add(new CacheItem(hash, @delegate), new CacheItemPolicy { AbsoluteExpiration = DateTime.UtcNow + TimeSpan.FromMinutes(5) });
            return @delegate;
        }


        /// <summary>
        /// Ensures
        /// 1) All required values are given
        /// 2) Not more parameters than defined are given
        /// and returns an array of values (with result.Count == definitions.Count), where missing optional values have been filled with their default values.
        /// This function assumes that null values are missing values.
        /// </summary>
        /// <param name="definitions">Parameter definitions</param>
        /// <param name="EnsureTypeConsistency">If true, the function will throw an ArgumentException if any of the given parameters were of a non-compatible type</param>
        /// <param name="values">Values passed to an Excel function</param>
        /// <returns></returns>
        /// <exception cref="ArgumentException">Thrown when the input arguments are invalid</exception>
        private static object[] CheckParameterConsistency(IEnumerable<IParameter> definitions, bool EnsureTypeConsistency, params object[] values)
        {
            var result = new List<object>();

            int i;
            for (i = 0; i < definitions.Count(); i++)
            {
                var definition = definitions.ElementAt(i);
                object value = (i >= values.Count()) ? null : values.ElementAt(i);

                if (value == null)
                {
                    if (definition.IsOptional)
                    {
                        result.Add(definition.DefaultValue);
                    }
                    else
                    {
                        throw new ArgumentException($"Non-optional Parameter {definition.Name} was not supplied.", definition.Name);
                    }

                }
                else
                {
                    Type valueType = value.GetType();

                    if (EnsureTypeConsistency && !definition.Type.IsAssignableFrom(valueType) && definition.Type != typeof(Excel.Range))
                    {
                        // the user specified some parameter with a wrong type
                        throw new ArgumentException($"Parameter {definition.Name} expects input of type {definition.Type}; you specified an incompatible/unconvertable type {valueType} ({value}). Please use input of correct type.", definition.Name);
                    }
                    else
                    {
                        if (!IsRangeType(definition.Type) && value != null && IsRangeType(value.GetType()))
                        {
                            // The parameter expects a non-range, but the input value is a range. So we will try to convert the input value to the expected parameter type - if possible
                            object innerValue;

                            if (value is Excel.Range)
                                innerValue = ((Excel.Range)value).Value;
                            else if (value is ExcelReference)
                                innerValue = ((ExcelReference)value).GetValue();
                            else
                                throw new InvalidOperationException($"value type {value.GetType().Name} not supported for range/value conversion");

                            // todo: what happens if parameter requested is object[,], but user passes only a single value?
                            // we should find a general conversion for that as well
                            result.Add(innerValue);
                        }
                        else
                        {
                            result.Add(value);
                        }

                    }
                }
            }

            var TooManyParametersGiven = values.Skip(i).Where(x => x != null);

            if (TooManyParametersGiven.Any())
                throw new ArgumentException($"The function you are calling only takes {definitions.Count()} parameters; you specified {TooManyParametersGiven.Count()} too many: " + String.Join(", ", TooManyParametersGiven.Select(x => "\"" + x + "\"")));

            // In case of arrays: Replace ExcelEmpty by null
            result = result.Select(x => (x is Array) ? (object)ConvertAll<object, object>((Array)x, y => (y is ExcelEmpty) ? null : y) : x).ToList();

            return result.ToArray();
        }

        private static bool IsRangeType(Type type)
        {
            return (type == typeof(Excel.Range) || type == typeof(ExcelReference));
        }

        // From http://stackoverflow.com/questions/3867961/c-altering-values-for-every-item-in-an-array
        static Array ConvertAll<TSource, TResult>(Array source,
                                      Converter<TSource, TResult> projection)
        {
            if (!typeof(TSource).IsAssignableFrom(source.GetType().GetElementType()))
            {
                throw new ArgumentException();
            }
            var dims = Enumerable.Range(0, source.Rank)
                .Select(dim => new
                {
                    lower = source.GetLowerBound(dim),
                    upper = source.GetUpperBound(dim)
                });
            var result = Array.CreateInstance(typeof(TResult),
                dims.Select(dim => 1 + dim.upper - dim.lower).ToArray(),
                dims.Select(dim => dim.lower).ToArray());
            var indices = dims
                .Select(dim => Enumerable.Range(dim.lower, 1 + dim.upper - dim.lower))
                .Aggregate(
                    (IEnumerable<IEnumerable<int>>)null,
                    (total, current) => total != null
                        ? total.SelectMany(
                            item => current,
                            (existing, item) => existing.Concat(new[] { item }))
                        : current.Select(item => (IEnumerable<int>)new[] { item }))
                .Select(index => index.ToArray());
            foreach (var index in indices)
            {
                var value = (TSource)source.GetValue(index);
                result.SetValue(projection(value), index);
            }
            return result;
        }

    }
}
