using ScriptingAbstractions;
using System;
using Excel = NetOffice.ExcelApi;

namespace RoslynScriptGlobals
{
    public class GlobalsFactory : MarshalByRefObject, IGlobalsFactory<Globals>
    {
        public Func<Excel.Application> _ApplicationFactory;

        public GlobalsFactory(Func<Excel.Application> ApplicationFactory)
        {
            this._ApplicationFactory = ApplicationFactory;
        }

        public GlobalsFactory()
        {
        }

        public Globals Create(AppDomain ExecutingDomain)
        {
            return new Globals(_ApplicationFactory);
        }

        object IGlobalsFactory.Create(AppDomain ExecutingDomain)
        {
            return Create(ExecutingDomain);
        }


          
    }
}
