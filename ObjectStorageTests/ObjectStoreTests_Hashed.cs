using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ObjectStorageTests.Types;
using ObjectStorageTests.Generators;

namespace ObjectStorageTests
{
    [TestClass]
    public class ObjectStoreTests_Hashed : ObjectStoreTestBase
    {
        private Func<int, IStoredType> m_generator = (i) => StoredTypeGenerator.Generate(i, StoredTypeGenerator.ImplementationMethod.Hash);
        protected override Func<int, IStoredType> ValidObjectFactory
        {
            get
            {
                return m_generator;
            }
        }
    }
}
