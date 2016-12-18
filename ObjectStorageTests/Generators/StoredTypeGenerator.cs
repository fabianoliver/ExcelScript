using ObjectStorageTests.Types;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ObjectStorageTests.Generators
{
    public class StoredTypeGenerator
    {
        public enum ImplementationMethod
        {
            Hash,
            NotifyPropertyChanged,
            NoneUnsupported
        }

        public static IStoredType Generate(int seed = 1, ImplementationMethod implementationMethod = ImplementationMethod.Hash)
        {
            string value1 = String.Format("Stored Type #{0}", seed);
            int value2 = seed;

            switch (implementationMethod)
            {
                case ImplementationMethod.Hash:
                    return new StoredTypeNotifyPropertyChanged(value1, value2);
                case ImplementationMethod.NotifyPropertyChanged:
                    return new StoredTypeNotifyPropertyChanged(value1, value2);
                case ImplementationMethod.NoneUnsupported:
                    return new StoredTypeUnsupported(value1, value2);
                default:
                    throw new NotImplementedException();
            }
        }
    }
}
