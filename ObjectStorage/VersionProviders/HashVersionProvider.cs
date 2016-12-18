using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ObjectStorage.VersionProviders
{
    internal class HashVersionProvider : IVersionProvider
    {
        private int m_Version = 1;
        private int m_lastHash = -1;

        public int GetVersionFor(object obj)
        {
            int new_hash = GetHashFor(obj);

            if(new_hash != m_lastHash)
            {
                m_lastHash = new_hash;
                return ++m_Version;
            } else
            {
                return m_Version;
            }
        }

        public static bool CanHandle(Type t)
        {
            return (t.GetMethod("GetHashCode").DeclaringType == t);
        }

        private static int GetHashFor(object obj)
        {
            if (obj == null)
                return -1;
            else
                return obj.GetHashCode();
        }

        /// <param name="InitialObject">Initial value. The value of this will not be stored in this class / will not store any reference to this</param>
        public HashVersionProvider(object InitialObject)
        {
            this.m_lastHash = GetHashFor(InitialObject);
        }

        public HashVersionProvider(object InitialObject, int initial_version)
            : this(InitialObject)
        {
            this.m_Version = initial_version;
        }
    }
}
