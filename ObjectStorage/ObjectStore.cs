using ObjectStorage.Abstractions;
using System;
using System.Collections.Concurrent;
using System.Linq;
using System.Reflection;

namespace ObjectStorage
{
    public class ObjectStore : IObjectStore
    {
        private readonly ConcurrentDictionary<string, IStoredObject> _StoredObjects = new ConcurrentDictionary<string, IStoredObject>();

        public bool GetByName(string name, out IStoredObject result)
        {
            IStoredObject<object> temp_result = null;
            bool success = GetByName<object>(name, out temp_result);

            result = temp_result as IStoredObject;
            return success;
        }


        public bool GetByName<T>(string name, out IStoredObject<T> result)
            where T : class
        {
            IStoredObject return_value = null;

            if (!_StoredObjects.TryGetValue(name, out return_value))
            {
                result = null;
                return false;
            }
            else
            {
                if(return_value is IStoredObject<T>)
                {
                    result = (IStoredObject<T>)return_value;
                    return true;
                } else if(return_value.Object is T && return_value is IUnregisterAware)
                {
                    result = Wrap<T>(name, (T)return_value.Object, ((IUnregisterAware)return_value).OnUnregister);
                    return true;
                } else
                {
                    result = null;
                    return false;
                }
            }
        }

        public bool Exists(string name)
        {
            return _StoredObjects.ContainsKey(name);
        }

        public bool Remove(string name)
        {
            IStoredObject _unused;
            return Remove(name, out _unused);
        }

        public bool Remove(string name, out IStoredObject RemovedItem)
        {
            bool success =  _StoredObjects.TryRemove(name, out RemovedItem);
            TryDispose(RemovedItem);
            return success;
        }

        // MethodInfo of AddOrUpdate<T>(string name, T Object, out IStoredObject<T> StoredObject, Action OnUnregistered)
        private static readonly MethodInfo _GenericAddOrUpdateMi = typeof(ObjectStore).GetMethods()
                         .Where(m => m.Name == nameof(AddOrUpdate))
                         .Select(m => new
                         {
                             Method = m,
                             Params = m.GetParameters(),
                             GenericArgs = m.GetGenericArguments()
                         })
                         .Where(x => x.Params.Length == 4
                                     && x.GenericArgs.Length == 1)
                         .Select(x => x.Method)
                         .Single();

        public bool AddOrUpdate(string name, object Object, Type objectType, out IStoredObject StoredObject, Action OnUnregistered = null)
        {
            var parameters = new object[] { name, Object, null, OnUnregistered };
            var mi = _GenericAddOrUpdateMi.MakeGenericMethod(objectType);
            bool result = (bool)mi.Invoke(this, parameters);
            StoredObject = (IStoredObject)parameters[2]; // the out IStoredObject<T> parameter
            return result;
        }

        public bool AddOrUpdate<T>(string name, T Object, out IStoredObject<T> StoredObject, Action OnUnregistered = null)
            where T : class
        {
           StoredObject =  (IStoredObject<T>)_StoredObjects.AddOrUpdate(
               name,

               // AddValueFactory
               (_name) =>
               {
                   return Wrap(name, Object, OnUnregistered);
               },

               // UpdateValueFactory
               (_name, _oldValue) =>
               {
                    // When updating, we must ensure to carry on the version of the previously existing object
                    var result = Wrap(name, Object, _oldValue, OnUnregistered);
                   TryDispose(_oldValue);
                   return result;
               }
           );

            return true;
        }

        public bool AddOrUpdate<T>(string name, T Object, Action OnUnregistered = null)
            where T : class
        {
            IStoredObject<T> _unused;
            return AddOrUpdate(name, Object, out _unused, OnUnregistered);
        }

        private void TryDispose(IStoredObject storedObject)
        {
            IDisposable disposable = storedObject as IDisposable;
            if (disposable != null)
                disposable.Dispose();
        }

        private IStoredObject<T> Wrap<T>(string Name, T Item, Action OnUnregistered)
            where T : class
        {
            return new StoredObject<T>(Name, Item, OnUnregistered);
        }

        private IStoredObject<T> Wrap<T>(string Name, T Item, IStoredObject oldItem, Action OnUnregistered)
           where T : class
        {
            return new StoredObject<T>(Name, Item, oldItem, OnUnregistered);
        }
    }
}
