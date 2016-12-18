using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ObjectStorageTests.Types
{
    public class StoredTypeHashed : IStoredType
    {
        public string Value1 { get; set; }
        public int Value2 { get; set; }

        public StoredTypeHashed(string Value1, int Value2)
        {
            this.Value1 = Value1;
            this.Value2 = Value2;
        }

        public override int GetHashCode()
        {
            int hash = 13;
            hash = (hash * 7) + Value1.GetHashCode();
            hash = (hash * 7) + Value2.GetHashCode();
            return hash;
        }
    }
}
