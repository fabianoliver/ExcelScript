using ScriptingAbstractions;
using System;
using System.Collections.Generic;
using System.Linq;
using Excel = NetOffice.ExcelApi;

namespace RoslynScriptGlobals
{
    public class Globals : IScriptGlobals, IDisposable
    {
        private readonly Lazy<Excel.Application> _Application;

        public IDictionary<string, object> Parameters { get; private set; }

        public Excel.Application Application { get { return GetApplication(); } }
        public Excel.Worksheet ActiveSheet { get { return GetActiveSheet(); } }
        public Excel.Workbook ActiveWorkbook { get { return GetActiveWorkbook(); } }
       
        public Globals(Func<Excel.Application> ApplicationFactory)
        {
            this.Parameters = new Dictionary<string, object>();
            this._Application = new Lazy<Excel.Application>(() => ApplicationFactory());
        }     
        
        private Excel.Worksheet GetActiveSheet()
        {
            var app = _Application.Value;
            var sheet = app.ActiveSheet;
            var result = (Excel.Worksheet)sheet;
            return result;
        }

        private Excel.Application GetApplication()
        {
            var result = _Application.Value;
            return result;
        }

        private Excel.Workbook GetActiveWorkbook()
        {
            var app = _Application.Value;
            var result = app.ActiveWorkbook;

            return result;
        }

        void IDisposable.Dispose()
        {
            foreach(var parameter in this.Parameters.Values)
            {
                var disposable = parameter as IDisposable;

                if (disposable != null)
                    disposable.Dispose();
            }

            if(_Application.IsValueCreated)
            {
                _Application.Value?.Dispose();
            }
        }

    }
}
