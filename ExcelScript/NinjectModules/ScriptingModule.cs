using ExcelDna.Integration;
using ExcelScript.Internal;
using Ninject.Modules;
using ObjectStorage;
using ObjectStorage.Abstractions;
using RoslynScriptGlobals;
using RoslynScripting.Factory;
using ScriptingAbstractions;
using ScriptingAbstractions.Factory;
using System;
using Excel = NetOffice.ExcelApi;

namespace ExcelScript.NinjectModules
{
    public class ScriptingModule : NinjectModule
    {
        public override void Load()
        {
            Bind<IScriptFactory>().To<ScriptFactory>().InSingletonScope();
            Bind<IParameterFactory>().To<ParameterFactory>().InSingletonScope();
            Bind<IObjectStore>().To<ObjectStore>().InSingletonScope();
            Bind<Func<Excel.Application>>().ToConstant<Func<Excel.Application>>(GetApplication).WhenInjectedInto<GlobalsFactory>();
            Bind<IGlobalsFactory<Globals>>().To<GlobalsFactory>().InSingletonScope();
            Bind<DirtyRangeFlagger>().ToSelf().InSingletonScope();
        }

        private static Excel.Application GetApplication()
        {
            return new Excel.Application(null, ExcelDnaUtil.Application);
        }
    }
}
