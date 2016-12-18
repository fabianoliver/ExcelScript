using ScriptingAbstractions;
using ScriptingAbstractions.Factory;

namespace RoslynScripting.Factory
{
    public class ParameterFactory : IParameterFactory
    {
        public IParameter Create()
        {
            var parameter = new Parameter();
            return parameter;
        }
    }
}
