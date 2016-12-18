using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ObjectStorage.VersionProviders
{
    internal class NullObjectVersionProvider : IVersionProvider
    {
        private int m_Version;

        public NullObjectVersionProvider(int initial_version)
        {
            this.m_Version = initial_version;
        }

        public int GetVersionFor(object obj)
        {
            if (obj != null)
                throw new InvalidOperationException($"NullObjectVersionProvider may only be called with Null-values; Non-null parameter was {obj}");

            return m_Version;
        }
    }
}
