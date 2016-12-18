using ExcelScript.CommonUtilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Xml;
using System.Xml.Schema;
using System.Xml.Serialization;

namespace ScriptingAbstractions
{
    public enum HostingType
    {
        /// <summary>
        /// Run in the same AppDomain that the main Addin (ExcelScript) is hosted in
        /// </summary>
        GlobalAppDomain,

        /// <summary>
        /// Run in a domain seperate from the main Addin, but shared by all scripts
        /// </summary>
        SharedSandboxAppDomain,

        /// <summary>
        /// Create a seperate app domain for each script
        /// </summary>
        IndividualScriptAppDomain
    }

    [Serializable]
    public class ScriptingOptions : MarshalByRefObject, IXmlSerializable
    {
        public virtual List<Assembly> References { get; protected set; } = new List<Assembly>();
        public virtual List<string> Imports { get; protected set; } = new List<string>();
        public HostingType HostingType { get; set; } = HostingType.SharedSandboxAppDomain;
        /*
        [SecurityPermissionAttribute(SecurityAction.Demand, Flags = SecurityPermissionFlag.Infrastructure)]
        public override object InitializeLifetimeService()
        {
            return null;
        }*/

        public override int GetHashCode()
        {
            unchecked
            {
                int hash = (int)2166136261;

                foreach (var item in References.Cast<object>().Concat(Imports.Cast<object>()))
                    hash = (hash * 16777619) ^ item.GetHashCode();

                int hostingType = (int)this.HostingType;
                hash = (hash * 16777619) ^ hostingType.GetHashCode();

                return hash;
            }
        }

        #region IXmlSerializable

        public XmlSchema GetSchema()
        {
            return null;
        }

        public void WriteXml(XmlWriter writer)
        {
            writer.WriteEnumerable(nameof(References), References, x => x.FullName);
            writer.WriteEnumerable(nameof(Imports), Imports);
            writer.Write(nameof(HostingType), HostingType);
        }

        public void ReadXml(XmlReader reader)
        {
            reader.MoveToCustomStart();

            this.References = reader.ReadEnumerable<Assembly>(nameof(References), x =>
            {
                string value = (string)x;
                Assembly assembly = Assembly.Load(value);
                return assembly;
            }).ToList();

            this.Imports = reader.ReadEnumerable<string>(nameof(Imports), x => (string)x).ToList();

            object _hostingType;
            reader.Read(nameof(HostingType), out _hostingType);
            this.HostingType = (HostingType)_hostingType;

            reader.ReadEndElement();
        }

        #endregion
    }
}
