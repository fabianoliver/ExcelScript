using ObjectStorage.Abstractions;
using System;
using System.ComponentModel;

namespace ObjectStorage
{
    public class WeaklyStoredObject<T> : StoredObject<T>, IStoredObject<T>
        where T : class, INotifyPropertyChanged
    {
        private WeakReference<T> m_Object;

        
        public override T Object
        {
            get
            {
                T result = default(T);
                bool success = m_Object.TryGetTarget(out result);

                if (!success)
                    return default(T);

                return result;
            }
            protected set
            {
                if (m_Object != value)
                {
                    OnSetObject(Object, value);
                    m_Object = new WeakReference<T>(value);
                }
            }

        }
        
        public WeaklyStoredObject(string Name, T Object, Action OnUnregister)
            : base(Name, Object, OnUnregister)
        {
        }

    }
}
