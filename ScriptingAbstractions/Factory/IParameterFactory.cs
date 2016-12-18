using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Remoting.Lifetime;
using System.Text;
using System.Threading.Tasks;

namespace ScriptingAbstractions.Factory
{
    public interface IParameterFactory
    {
        IParameter Create();
    }
}
