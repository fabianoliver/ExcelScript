using Microsoft.VisualStudio.TestTools.UnitTesting;
using ObjectStorageTests.Generators;
using ObjectStorageTests.Types;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ObjectStorageTests
{
    [TestClass]
    public class ObjectStoreTests_NotifyPropertyChanged : ObjectStoreTestBase
    {
        private Func<int, IStoredType> m_generator = (i) => StoredTypeGenerator.Generate(i, StoredTypeGenerator.ImplementationMethod.NotifyPropertyChanged);
        protected override Func<int, IStoredType> ValidObjectFactory
        {
            get
            {
                return m_generator;
            }
        }
    }
}
