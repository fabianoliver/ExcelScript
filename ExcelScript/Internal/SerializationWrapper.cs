using ExcelScript.CommonUtilities;
using System.Xml;
using System.Xml.Schema;
using System.Xml.Serialization;

namespace ExcelScript.Internal
{
    public class SerializationWrapper : IXmlSerializable
    {
        public object ContainedObject { get; set; }

        #region IXmlSerializable

        public XmlSchema GetSchema()
        {
            return null;
        }

        public void WriteXml(XmlWriter writer)
        {
            writer.Write(nameof(ContainedObject), ContainedObject);
        }

        public void ReadXml(XmlReader reader)
        {
            reader.MoveToCustomStart();

            object _ContainedObject;

            reader.Read(nameof(ContainedObject), out _ContainedObject);

            this.ContainedObject = _ContainedObject;

            reader.ReadEndElement();
        }

        #endregion


    }
}
