using ExcelScript.CommonUtilities;
using ObjectStorage.Abstractions;
using ObjectStorage.VersionProviders;
using System;
using System.Xml;
using System.Xml.Schema;
using System.Xml.Serialization;

namespace ObjectStorage
{
    internal interface IUnregisterAware
    {
        Action OnUnregister { get; }
    }

    /// <summary>
    /// Implementation of IStoredObject<T>.
    /// For Versioning to work correctly, objets stored with this class should propertly implement GetHashCode() for value equality
    /// </summary>
    /// <typeparam name="T">Type of the stored object</typeparam>
    public class StoredObject<T> : IStoredObject<T>, IDisposable, IXmlSerializable, IUnregisterAware
        where T : class
    {
        public Action OnUnregister { get; private set; }

        public virtual string Name { get; protected set; }

        private T m_Object;
        public virtual T Object
        {
            get
            {
                return m_Object;
            }
            protected set
            {
                if (m_Object != value)
                {
                    OnSetObject(m_Object, value);
                    m_Object = value;
                }
            }
        }

        protected IVersionProvider m_VersionProvider;

        public virtual int Version
        {
            get
            {
                return m_VersionProvider.GetVersionFor(this.Object);
            }
        }

        // IStoredObject
        Object IStoredObject.Object
        {
            get
            {
                return Object;
            }
        }

        // IStoredObject<T>
        T IStoredObject<T>.Object
        {
            get
            {
                return Object;
            }
        }

        protected void OnSetObject(T oldValue, T newValue)
        {
            IDisposable oldProviderDisposable = this.m_VersionProvider as IDisposable;
            
            this.m_VersionProvider = CreateVersionProviderFor(newValue, oldValue);

            if (oldProviderDisposable != null)
                oldProviderDisposable.Dispose();
        }

        public void Dispose()
        {
            if (OnUnregister != null)
                OnUnregister();

            IDisposable disposable = this.m_VersionProvider as IDisposable;

            if (disposable != null)
                disposable.Dispose();       
        }

        public StoredObject(string Name, T Object, Action OnUnregister)
            : this(Name, Object, null, OnUnregister)
        {
        }

        // This constructor is needed for Deserialization
        private StoredObject()
            : this(null, null, null, null)
        {
        }

        /// <summary>
        /// This constructor can be called when creating a new StoredObject which is intended to replace the OldObject in the store.
        /// The constructor will ensure the new object's versioning carries forward the old version's versioning.
        /// </summary>
        /// <param name="Name">Name under which the object is added to the store</param>
        /// <param name="Object">The actual (new) object</param>
        /// <param name="OldObject">The old object which is being replaced</param>
        internal StoredObject(string Name, T Object, IStoredObject OldObject, Action OnUnregister)
        {
            this.Name = Name;
            this.OnUnregister = OnUnregister;

            int initial_version = (OldObject == null) ? 0 : OldObject.Version;
            this.m_VersionProvider = new ConstantVersionProvider(initial_version); // we'll start with a ConstantVersionProvider to supply the initial version. Each time this.Objec is set (such as later in this constructor), this.m_VersionProvider will be replaced with a provider that actually suits the runtime type anyway.

            this.Object = Object;
        }


        private IVersionProvider CreateVersionProviderFor(T newValue, T oldValue)
        {
            int initial_version = (m_VersionProvider.GetVersionFor(oldValue)+1);

            if (newValue == null)
                return new NullObjectVersionProvider(initial_version);

            Type runtimeType = newValue.GetType();

            if (HashVersionProvider.CanHandle(runtimeType))
                return new HashVersionProvider(newValue, initial_version);
            else if (NotifyPropertyChangedVersionProvider.CanHandle(runtimeType))
                return new NotifyPropertyChangedVersionProvider(newValue, initial_version);
            else
                throw new InvalidOperationException(String.Format("No IVersionProvider found that can handle type '{0}'", typeof(T).FullName));
        }

        #region IXmlSerializable

        public XmlSchema GetSchema()
        {
            return null;
        }

        public void WriteXml(XmlWriter writer)
        {
            writer.Write(nameof(Name), Name);
            writer.Write(nameof(Version), Version);
            writer.Write(nameof(Object), (object)Object);
        }

        public void ReadXml(XmlReader reader)
        {
            reader.MoveToCustomStart();

            string _Name;
            int _Version;
            object _Obj;

            reader.Read(nameof(Name), out _Name);
            reader.Read(nameof(Version), out _Version);
            reader.Read(nameof(Object), out _Obj);

            this.Name = _Name;
            this.Object = (T)_Obj;
            // no need to set version

            reader.ReadEndElement();
        }

        #endregion

    }
}
