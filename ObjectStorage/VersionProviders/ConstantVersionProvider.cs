using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ObjectStorage.VersionProviders
{
    internal class ConstantVersionProvider : IVersionProvider
    {
        private int m_Version;

        public ConstantVersionProvider(int Version)
        {
            this.m_Version = Version;
        }

        public int GetVersionFor(object obj)
        {
            return this.m_Version;
        }
    }
}
