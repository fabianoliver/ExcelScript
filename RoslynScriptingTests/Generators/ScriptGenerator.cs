using RoslynScripting;
using RoslynScripting.Factory;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RoslynScriptingTests.Generators
{
    public static class ScriptGenerator
    {
        public static Script<TestGlobals> Generate(int seed = 1)
        {
            var globalsFactory = new TestGlobalsFactory();
            var script = new ScriptFactory().Create<TestGlobals>((x) => new TestGlobalsFactory().Create(x));

            script.ReturnType = typeof(int);
            script.Code = $"1+{seed}";
            script.Description = $"Description of script {seed}";
   
            for(int i = 1; i <= 2 + (seed%2); i++)
            {
                var parameter = ParameterGenerator.Generate(i);
                script.Parameters.Add(parameter);
            }

            return (Script<TestGlobals>)script;
        }
    }
}
