using Microsoft.VisualStudio.TestTools.UnitTesting;
using ObjectStorage;
using ObjectStorage.Abstractions;
using ObjectStorageTests.Generators;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ObjectStorageTests
{
    [TestClass]
    public class ObjectStoreTests_Mixed
    {
        [TestMethod]
        public void TestReplaceAndVersioning_Mixed()
        {
            var store = new ObjectStore();

            string key = "MyObject1";
            var obj1 = StoredTypeGenerator.Generate(1, StoredTypeGenerator.ImplementationMethod.Hash);
            var obj2 = StoredTypeGenerator.Generate(1, StoredTypeGenerator.ImplementationMethod.NotifyPropertyChanged);

            ObjectStoreTestBase.CheckedAddToStore(store, key, obj1);

            {
                IStoredObject result;
                bool success = store.GetByName(key, out result);
                int version = result.Version;
                int expectedVersion = 1;

                Assert.AreEqual(expectedVersion, version, $"Expected version to be {expectedVersion}, but was {version} instead");
            }

            // Replace the object
            ObjectStoreTestBase.CheckedAddToStore(store, key, obj2);

            {
                IStoredObject result;
                bool success = store.GetByName(key, out result);
                int version = result.Version;
                int expectedVersion = 2;

                Assert.AreEqual(expectedVersion, version, $"Expected version to be {expectedVersion}, but was {version} instead");
            }
        }
    }
}
