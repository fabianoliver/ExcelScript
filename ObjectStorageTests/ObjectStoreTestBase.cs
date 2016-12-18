using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ObjectStorage;
using ObjectStorageTests.Generators;
using ObjectStorage.Abstractions;
using ObjectStorageTests.Types;

namespace ObjectStorageTests
{
    public abstract class ObjectStoreTestBase
    {
        // Factory that generates IStoredType objects that are assumed to be valid & manageable by the object store. For example, StoreTypeHashed or StoreTypeNotifyPropertyChanged objects.
        protected abstract Func<int, IStoredType> ValidObjectFactory { get; }

        private Func<int, IStoredType> m_invalidObjGenerator = (i) => StoredTypeGenerator.Generate(i, StoredTypeGenerator.ImplementationMethod.NoneUnsupported);

        // Factory that generates IStoredType objects that are assumed NOT to be valid & manageable by the object store. For example, StoreTypeUnsupported objects.
        protected virtual Func<int, IStoredType> InvalidObjectFactory
        {
            get
            {
                return m_invalidObjGenerator;
            }
        }

        [TestMethod]
        // Adds an object to the store and retrieves it again
        public void TestAddAndRetrieve()
        {
            var store = new ObjectStore();

            string key = "MyObject1";
            var obj = ValidObjectFactory(1);

            CheckedAddToStore(store, key, obj);
        }

        [TestMethod]
        // Adds an object to the store, then removes it again
        public void TestRemove()
        {
            var store = new ObjectStore();

            string key = "MyObject1";
            var obj = ValidObjectFactory(1);

            bool added = store.AddOrUpdate(key, obj);
            Assert.IsTrue(added, "Failed to add object to the object store");

            bool exists = store.Exists(key);
            Assert.IsTrue(exists, "Object was not found in the object store");

            bool removed = store.Remove(key);
            Assert.IsTrue(removed, "Object could not be removed from the object store");

            bool doesNotExist = !store.Exists(key);
            Assert.IsTrue(doesNotExist, "Object was still found in the object store, even though it should have been removed");
        }

        [TestMethod]
        // Replaces an object in a store with another one
        public void TestReplace()
        {
            var store = new ObjectStore();

            string key = "MyObject1";
            var obj1 = ValidObjectFactory(1);
            var obj2 = ValidObjectFactory(2);

            CheckedAddToStore(store, key, obj1);
            CheckedAddToStore(store, key, obj2);
        }
        
        [TestMethod]
        // Adds an object to the store, then changes one of its properties.
        // Checks that the version is correctly being increased through this operation.
        public void TestVersioning()
        {
            var store = new ObjectStore();

            string key = "MyObject1";
            var obj = ValidObjectFactory(1);

            CheckedAddToStore(store, key, obj);

            {
                IStoredObject result;
                bool success = store.GetByName(key, out result);
                int version = result.Version;
                int expectedVersion = 1;

                Assert.AreEqual(expectedVersion, version, $"Expected version to be {expectedVersion}, but was {version} instead");
            }

            // modify the object
            obj.Value1 = obj.Value1 + " - modified";

            {
                IStoredObject result;
                bool success = store.GetByName(key, out result);
                int version = result.Version;
                int expectedVersion = 2;

                Assert.AreEqual(expectedVersion, version, $"Expected version to be {expectedVersion}, but was {version} instead");
            }
        }

        [TestMethod]
        // Adds an object to the store, and replaces it with a different (different values) object.
        // Checks version is correctly increased.
        public void TestReplaceAndVersioning()
        {
            var store = new ObjectStore();

            string key = "MyObject1";
            var obj1 = ValidObjectFactory(1);
            var obj2 = ValidObjectFactory(2);

            CheckedAddToStore(store, key, obj1);

            {
                IStoredObject result;
                bool success = store.GetByName(key, out result);
                int version = result.Version;
                int expectedVersion = 1;

                Assert.AreEqual(expectedVersion, version, $"Expected version to be {expectedVersion}, but was {version} instead");
            }

            // Replace the object
            CheckedAddToStore(store, key, obj2);

            {
                IStoredObject result;
                bool success = store.GetByName(key, out result);
                int version = result.Version;
                int expectedVersion = 2;

                Assert.AreEqual(expectedVersion, version, $"Expected version to be {expectedVersion}, but was {version} instead");
            }
        }

        [TestMethod]
        // Adds an object to the store, and replaces it with a different object (reference) with the same values (!)
        // by contract, versioning should still be increased through this operation
        public void TestReplaceWithDuplicateAndVersioning()
        {
            var store = new ObjectStore();

            string key = "MyObject1";
            var obj1 = ValidObjectFactory(1);
            var obj2 = ValidObjectFactory(2);

            CheckedAddToStore(store, key, obj1);

            {
                IStoredObject result;
                bool success = store.GetByName(key, out result);
                int version = result.Version;
                int expectedVersion = 1;

                Assert.AreEqual(expectedVersion, version, $"Expected version to be {expectedVersion}, but was {version} instead");
            }

            // Replace the object
            CheckedAddToStore(store, key, obj2);

            {
                IStoredObject result;
                bool success = store.GetByName(key, out result);
                int version = result.Version;
                int expectedVersion = 2;

                Assert.AreEqual(expectedVersion, version, $"Expected version to be {expectedVersion}, but was {version} instead");
            }
        }

        [TestMethod]
        [ExpectedException(typeof(InvalidOperationException))]
        // Tries adding an unsuppirted type to the store,
        // and ensures the correct exception is being thrown because of that.
        public void TestUnsupportedType()
        {
            var store = new ObjectStore();

            string key = "MyObject1";
            var obj = InvalidObjectFactory(1);

            // This should throw an InvalidOperationException
            store.AddOrUpdate<IStoredType>(key, obj);
        }

        public static void CheckedAddToStore(ObjectStore store, string key, IStoredType obj)
        {
            bool added = store.AddOrUpdate(key, obj);
            Assert.IsTrue(added, "Failed to add object to the object store");

            bool exists = store.Exists(key);
            Assert.IsTrue(exists, "Object was not found in the object store");

            IStoredObject<IStoredType> retrievedObj = null;
            bool retrievedSuccess = store.GetByName<IStoredType>(key, out retrievedObj);
            Assert.IsTrue(retrievedSuccess, "Object could not be retrieved from the object store");

            bool retrievedObjectEqualsOriginal = retrievedObj.Object.Equals(obj);
            Assert.IsTrue(retrievedObjectEqualsOriginal, "Retrieved object does not equal original object");
        }
    }
}
