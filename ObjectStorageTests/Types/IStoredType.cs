using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ObjectStorageTests.Types
{
    public interface IStoredType
    {
        string Value1 { get; set; }
        int Value2 { get; set; }
    }
}
