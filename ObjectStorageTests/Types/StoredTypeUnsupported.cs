using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ObjectStorageTests.Types
{
    /// <summary>
    /// This type should not be supported, as it does not override GetHashCode()
    /// </summary>
    public class StoredTypeUnsupported : IStoredType
    {
        public string Value1 { get; set; }
        public int Value2 { get; set; }

        public StoredTypeUnsupported(string Value1, int Value2)
        {
            this.Value1 = Value1;
            this.Value2 = Value2;
        }

    }
}
