using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Serialization;

namespace ExcelScript.CommonUtilities
{
    public static class XmlSerializationExtensions
    {
        #region Write

        public static void Write(this XmlWriter writer, string localName, string value)
        {
            writer.WriteStartElement(localName);
            writer.WriteAttributeString("Value", value);
            writer.WriteFullEndElement();
        }

        public static void Write(this XmlWriter writer, string localName, bool value)
        {
            writer.Write(localName, Convert.ToString(value));
        }

        public static void Write(this XmlWriter writer, string localName, int value)
        {
            writer.Write(localName, Convert.ToString(value));
        }

        public static void Write(this XmlWriter writer, string localName, float value)
        {
            writer.Write(localName, Convert.ToString(value));
        }

        public static void Write(this XmlWriter writer, string localName, UInt16 value)
        {
            writer.Write(localName, Convert.ToString(value));
        }

        public static void Write(this XmlWriter writer, string localName, UInt32 value)
        {
            writer.Write(localName, Convert.ToString(value));
        }

        public static void Write(this XmlWriter writer, string localName, UInt64 value)
        {
            writer.Write(localName, Convert.ToString(value));
        }

        public static void Write(this XmlWriter writer, string localName, Int16 value)
        {
            writer.Write(localName, Convert.ToString(value));
        }

        public static void Write(this XmlWriter writer, string localName, Int64 value)
        {
            writer.Write(localName, Convert.ToString(value));
        }

        public static void Write(this XmlWriter writer, string localName, Decimal value)
        {
            writer.Write(localName, Convert.ToString(value));
        }

        public static void Write(this XmlWriter writer, string localName, Byte value)
        {
            writer.Write(localName, Convert.ToString(value));
        }

        public static void Write(this XmlWriter writer, string localName, Type value)
        {
            string strValue = (value == null) ? String.Empty : value.AssemblyQualifiedName;
            writer.Write(localName, strValue);
        }

        public static void Write(this XmlWriter writer, string localName, object value)
        {
            WriteXmlObjectProperty(writer, localName, value);
        }

        public static void WriteEnumerable<TItems, TSerializationProperty>(this XmlWriter writer, string localName, IEnumerable<TItems> Items, Func<TItems, TSerializationProperty> SerializationPropertySelector)
        {
            writer.WriteStartElement(localName);

            if (Items != null)
            {
                foreach (var item in Items)
                {
                    var itemProperty = SerializationPropertySelector(item);
                    writer.Write("Item", itemProperty);
                }
            }

            writer.WriteFullEndElement();
        }

        public static void WriteEnumerable(this XmlWriter writer, string localName, IEnumerable Items)
        {
            WriteEnumerable(writer, localName, Items.Cast<object>(), x => x);
            /*
            writer.WriteStartElement(localName);

            if (Items != null)
            {
                foreach (var item in Items)
                {
                    writer.Write("Item", item);
                }
            }


            writer.WriteFullEndElement();
            */
        }

        #endregion


        #region Read

        public static void MoveToCustomStart(this XmlReader reader)
        {
            reader.MoveToContent();
            reader.ReadStartElement();
        }

        private static string ReadStr(this XmlReader reader, string localName)
        {
            // Moves to the element
            if (!reader.IsStartElement(localName))
                throw new FormatException($"Expected xml node {localName}");
 
            // Get the serialized type
            string value = reader.GetAttribute("Value");
            reader.Read();
            reader.ReadEndElement();

            return value;
        }

        public static void Read(this XmlReader reader, string localName, out string value)
        {

            value = reader.ReadStr(localName);
        }

        public static void Read(this XmlReader reader, string localName, out bool value)
        {
            value = Convert.ToBoolean(reader.ReadStr(localName));
        }

        public static void Read(this XmlReader reader, string localName, out int value)
        {
            value = Convert.ToInt32(reader.ReadStr(localName));
        }

        public static void Read(this XmlReader reader, string localName, out float value)
        {
            value = Convert.ToSingle(reader.ReadStr(localName));
        }

        public static void Read(this XmlReader reader, string localName, out UInt16 value)
        {
            value = Convert.ToUInt16(reader.ReadStr(localName));
        }

        public static void Read(this XmlReader reader, string localName, out UInt32 value)
        {
            value = Convert.ToUInt32(reader.ReadStr(localName));
        }

        public static void Read(this XmlReader reader, string localName, out UInt64 value)
        {
            value = Convert.ToUInt64(reader.ReadStr(localName));
        }

        public static void Read(this XmlReader reader, string localName, out Int16 value)
        {
            value = Convert.ToInt16(reader.ReadStr(localName));
        }

        public static void Read(this XmlReader reader, string localName, out Decimal value)
        {
            value = Convert.ToDecimal(reader.ReadStr(localName));
        }

        public static void Read(this XmlReader reader, string localName, out Byte value)
        {
            value = Convert.ToByte(reader.ReadStr(localName));
        }

        public static void Read(this XmlReader reader, string localName, out Type value)
        {
            string typeName = reader.ReadStr(localName);
            value = String.IsNullOrWhiteSpace(typeName) ? null : Type.GetType(typeName);
        }

        public static void Read(this XmlReader reader, string localName, out object value)
        {
            ReadXmlObjectProperty(reader, localName, out value);
        }

        public static IEnumerable<TItems> ReadEnumerable<TItems>(this XmlReader reader, string localName, Func<object,TItems> Converter)
        {
            var result = new List<TItems>();

            if (!reader.IsStartElement(localName))
                throw new FormatException($"Expected xml node {localName}");

            reader.ReadStartElement();

            while (reader.IsStartElement("Item"))
            {
                object item;
                reader.Read("Item", out item);

                TItems restoredItem = Converter(item);

                result.Add(restoredItem);
            }

            reader.ReadEndElement();

            return result;
        }

        public static IEnumerable ReadEnumerable(this XmlReader reader, string localName)
        {
            return ReadEnumerable<object>(reader, localName, x => x);
            /*
            var result = new List<object>();

            if (!reader.IsStartElement(localName))
                throw new FormatException($"Expected xml node {localName}");

            reader.ReadStartElement();

            while (reader.IsStartElement("Item"))
            {
                object item;
                reader.Read("Item", out item);

                result.Add(item);
            }

            reader.ReadEndElement();

            return result;*/
        }

        #endregion

        // From http://stackoverflow.com/questions/9497310/how-to-serialize-property-of-type-object-with-xmlserializer
        private static bool ReadXmlObjectProperty(XmlReader reader, string name, out object value)
        {
            value = null;

            // Moves to the element
            while (!reader.IsStartElement(name))
            {
                return false;
            }
            // Get the serialized type
            string typeName = reader.GetAttribute("Type");

            Boolean isEmptyElement = reader.IsEmptyElement;
            reader.ReadStartElement();
            if (!isEmptyElement)
            {
                Type type = Type.GetType(typeName);

                if (type != null)
                {
                    // Deserialize it
                    XmlSerializer serializer = new XmlSerializer(type);
                    value = serializer.Deserialize(reader);
                }
                else
                {
                    // Type not found within this namespace: get the raw string!
                    string xmlTypeName = typeName.Substring(typeName.LastIndexOf('.') + 1);
                    value = reader.ReadElementString(xmlTypeName);
                }
                reader.ReadEndElement();
            }

            return true;
        }
        private static void WriteXmlObjectProperty(XmlWriter writer, string name, object value)
        {
            if (value != null)
            {
                Type valueType = value.GetType();
                writer.WriteStartElement(name);
                writer.WriteAttributeString("Type", valueType.AssemblyQualifiedName);
                writer.WriteRaw(ToXmlString(value, valueType));
                writer.WriteFullEndElement();
            }
        }

        private static string ToXmlString(object item, Type type)
        {
            XmlWriterSettings settings = new XmlWriterSettings();
            settings.Encoding = Encoding.ASCII;
            settings.Indent = true;
            settings.OmitXmlDeclaration = true;
            settings.NamespaceHandling = NamespaceHandling.OmitDuplicates;

            using (StringWriter textWriter = new StringWriter())
            using (XmlWriter xmlWriter = XmlWriter.Create(textWriter, settings))
            {
                XmlSerializer serializer = new XmlSerializer(type);
                serializer.Serialize(xmlWriter, item, new XmlSerializerNamespaces(new[] { XmlQualifiedName.Empty }));
                return textWriter.ToString();
            }
        }

    }
}
