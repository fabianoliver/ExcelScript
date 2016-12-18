using ScriptingAbstractions;
using System;

namespace ExcelScript
{
    public class XlScriptOptions
    {
        public Type ReturnType { get; private set; }
        public HostingType HostingType { get; private set; }


        public XlScriptOptions()
        {
            this.ReturnType = typeof(Object);
            this.HostingType = HostingType.SharedSandboxAppDomain;
        }

        public XlScriptOptions(Type ReturnType, HostingType HostingType)
        {
            this.ReturnType = ReturnType;
            this.HostingType = HostingType;
        }

        public static XlScriptOptions Default
        {
            get
            {
                return new XlScriptOptions();
            }
        }

        public override int GetHashCode()
        {
            unchecked
            {
                int hash = (int)2166136261;

                hash = (hash * 16777619) ^ ReturnType.GetHashCode();
                hash = (hash * 16777619) ^ HostingType.GetHashCode();

                return hash;
            }
        }
    }
}
