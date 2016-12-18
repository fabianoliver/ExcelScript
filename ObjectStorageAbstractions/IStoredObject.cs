namespace ObjectStorage.Abstractions
{
    /// <summary>
    /// Represents a stored object; allows to retrieve the object itself, as well as meta-information about it
    /// </summary>
    public interface IStoredObject
    {
        /// <summary>
        /// Name under which the object has been stored
        /// </summary>
        string Name { get; }

        /// <summary>
        /// The actual stored object
        /// </summary>
        object Object { get; }

        /// <summary>
        /// Version of this object
        /// </summary>
        int Version { get; }
    }

    /// <summary>
    /// Represents a stored object; allows to retrieve the object itself, as well as meta-information about it
    /// </summary>
    public interface IStoredObject<out T> : IStoredObject
        where T : class
    {
        /// <summary>
        /// The actual stored object
        /// </summary>
        new T Object { get; }
    }


}
