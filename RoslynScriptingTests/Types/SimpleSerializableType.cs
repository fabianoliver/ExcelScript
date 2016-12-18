using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RoslynScriptingTests.Types
{
    [Serializable]
    public class SimpleSerializableType
    {
        public string Value1 { get; set; }
        public int Value2 { get; set; }

        public override int GetHashCode()
        {
            int hash;

            unchecked
            {
                hash = 27;
                hash = (hash * 23) + ((Value1 == null) ? 0 : Value1.GetHashCode());
                hash = (hash * 23) + Value2.GetHashCode();
            }

            return hash;
        }
    }
}
