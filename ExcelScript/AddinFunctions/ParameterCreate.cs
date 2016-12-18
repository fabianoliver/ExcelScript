using ExcelDna.Integration;
using ExcelScript.Internal;
using ExcelScript.Registration;
using ObjectStorage.Abstractions;
using ScriptingAbstractions;
using System;

namespace ExcelScript
{
    public partial class ExcelScriptAddin
    {
        [ExcelFunction(Name = FunctionPrefix + "Parameter.Create", Description = "Creates a definition for an input parameter for a user script. This can then be passed on to ExcelSCript.Create() ParameterHandles.", IsVolatile = false)]
        [SuppressInDialog]
        [ManageDirtyFlags]
        public static object Parameter_Create(
            [ExcelArgument(Description = "Name under which the parameter will be usable in scripts")] string ParameterName,
            [ExcelArgument(Description = "Type of the parameter")] string ParameterType,
            [ExcelArgument(Description = "If true, this parameter can be ommitted when calling the script, and will then default to its defaultValue (next parameter)")] bool isOptional = false,
            [ExcelArgument(Description = "Default value if the parameter is ommitted - only specify this when isOptional (previous parameter) = true")] object defaultValue = null,
            string Description = null,
            string HandleName = null)
        {
            if (String.IsNullOrWhiteSpace(ParameterName))
                throw new ArgumentException("Parameter name was empty", nameof(ParameterName));

            if (!Microsoft.CodeAnalysis.CSharp.SyntaxFacts.IsValidIdentifier(ParameterName))
                throw new ArgumentException("Parameter name was not in a valid format; must be in the format of a valid C# identifier", nameof(ParameterName));

            if (!isOptional && defaultValue != null)
                throw new ArgumentException("Please omit the default value when specifying the parameter as non-optional (isOptional = false).", nameof(defaultValue));

            if (HandleName == null)
                HandleName = ParameterName;

            var parameter = m_ParameterFactory.Create();

            parameter.Name = ParameterName;
            parameter.Type = ToType(ParameterType);
            parameter.IsOptional = isOptional;
            parameter.Description = Description;

            if (isOptional)
                parameter.DefaultValue = defaultValue;

            IStoredObject<IParameter> storedParameter;
            string result = AddOrUpdateInStoreOrThrow<IParameter>(HandleName, parameter, out storedParameter, () => TryDispose(parameter));

            return result;
        }
    }
}
