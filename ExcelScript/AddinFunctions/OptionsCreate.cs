using ExcelDna.Integration;
using ExcelScript.Internal;
using ExcelScript.Registration;
using System;

namespace ExcelScript
{
    public partial class ExcelScriptAddin
    {
        [ExcelFunction(Name = FunctionPrefix + "Options.Create", Description = "Can be used to create/set options for user scripts. The result handle of this can be passed to the ExcelScript.Create() function's OptionsHandle parameter", IsVolatile = false)]
        [SuppressInDialog]
        [ManageDirtyFlags]
        public static object OptionsCreate(
            [ExcelArgument(Description = "Name of the object handle to be created")] string HandleName,
            [ExcelArgument(Description = "Return type of the script. Object by default.")] string ReturnType = "object",
            [ExcelArgument(Description = "Hosting type for script execution. Possible values: Shared = Same for all scripts (default). Individual = Seperate AppDomain for each script. Global = Same AppDomain as ExcelScript-Addin.")] string HostingType = "Shared"
            )
        {
            Type _ReturnType = ToType(ReturnType);
            ScriptingAbstractions.HostingType _HostingType = ToHostingType(HostingType);

            var options = new XlScriptOptions(_ReturnType, _HostingType);
            return AddOrUpdateInStoreOrThrow<XlScriptOptions>(HandleName, options, () => TryDispose(options));
        }
    }
}
