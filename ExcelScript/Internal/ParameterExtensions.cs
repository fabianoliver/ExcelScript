using ExcelScript.Internal;
using ScriptingAbstractions;
using ScriptingAbstractions.Factory;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelScript
{
    internal static class ParameterExtensions
    {
        private static readonly IParameterValueFactory m_ParameterValueFactory = new ParameterValueFactory();

        public static IParameterValue WithValue(this IParameter parameter, object Value)
        {
            return m_ParameterValueFactory.CreateFor(parameter, Value);
        }
    }
}
