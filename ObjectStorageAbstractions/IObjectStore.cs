using System;

namespace ObjectStorage.Abstractions
{
    public interface IObjectStore
    {
        /// <summary>
        /// Gets an object by name
        /// </summary>
        /// <param name="name">Name of the object</param>
        /// <param name="result">Result of the operation</param>
        /// <returns>True if an object with the given name was found in the obejct store. False otherwise.</returns>
        bool GetByName(string name, out IStoredObject result);

        /// <summary>
        /// Gets an object by name
        /// </summary>
        /// <param name="name">Name of the object</param>
        /// <param name="result">Result of the operation</param>
        /// <returns>True if an object with the given name was found in the obejct store. False otherwise.</returns>
        bool GetByName<T>(string name, out IStoredObject<T> result) where T : class;

        /// <summary>
        /// Checks if an object with the given name exists in the store.
        /// Use the result of this operation with care in multi-threaded scenarios, as the object may be removed before the next call.
        /// </summary>
        /// <param name="name">Name of the Object</param>
        /// <returns>True if the object store contains an item with the given name. False otherwise.</returns>
        bool Exists(string name);

        /// <summary>
        /// Removes an object with the given name from the object store
        /// </summary>
        /// <param name="name">Name of the object which is to be removed</param>
        /// <returns>True if the item was successfully removed. False otherwise (e.g. because no item with the name existed)</returns>
        bool Remove(string name);

        /// <summary>
        /// Removes an object with the given name from the object store
        /// </summary>
        /// <param name="name">Name of the object which is to be removed</param>
        /// <param name="RemovedItem">Object which was removed from the store</param>
        /// <returns>True if the item was successfully removed. False otherwise (e.g. because no item with the name existed)</returns>
        bool Remove(string name, out IStoredObject RemovedItem);

        /// <summary>
        /// Adds an object with the given name to the object store, or updates the record if one with the same name already exists
        /// </summary>
        /// <param name="name">Name under which the given OBject shall be added</param>
        /// <param name="Object">Object which shall be added to the object store</param>
        /// <param name="OnUnregistered">Callback to execute when the object is remoed from the store (explicitly or because its replaced by a new object)</param>
        /// <returns>True if the object was added or updated to the store</returns>
        bool AddOrUpdate<T>(string name, T Object, Action OnUnregistered = null) where T : class;

        /// <summary>
        /// Adds an object with the given name to the object store, or updates the record if one with the same name already exists
        /// </summary>
        /// <param name="name">Name under which the given OBject shall be added</param>
        /// <param name="Object">Object which shall be added to the object store</param>
        /// <param name="OnUnregistered">Callback to execute when the object is remoed from the store (explicitly or because its replaced by a new object)</param>
        /// <param name="StoredObject">Returns the IStoredObject that was stored</param>
        /// <returns>True if the object was added or updated to the store</returns>
        bool AddOrUpdate<T>(string name, T Object, out IStoredObject<T> StoredObject, Action OnUnregistered = null) where T : class;

        /// <summary>
        /// Adds an object with the given name to the object store, or updates the record if one with the same name already exists
        /// </summary>
        /// <param name="name">Name under which the given OBject shall be added</param>
        /// <param name="Object">Object which shall be added to the object store</param>
        /// <param name="objectType">Runtime type of the object under which it shall be stored (will be stored as an IStoredObject<$objectType>)</param>
        /// <param name="OnUnregistered">Callback to execute when the object is remoed from the store (explicitly or because its replaced by a new object)</param>
        /// <param name="StoredObject">Returns the IStoredObject that was stored. This will actually be an IStoredObject<$objectType>.</param>
        /// <returns>True if the object was added or updated to the store</returns>
        bool AddOrUpdate(string name, object Object, Type objectType, out IStoredObject StoredObject, Action OnUnregistered = null);
    }
}
