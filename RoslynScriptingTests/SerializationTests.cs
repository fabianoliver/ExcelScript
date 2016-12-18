using Microsoft.VisualStudio.TestTools.UnitTesting;
using RoslynScripting;
using RoslynScriptingTests.Generators;
using System.IO;
using System.Xml.Serialization;

namespace RoslynScriptingTests
{
    [TestClass]
    public class SerializationTests
    {
        [TestMethod]
        public void SerializeDeserialize_Parameter()
        {
            var parameter = ParameterGenerator.Generate(1);

            string xml = XmlSerialize<Parameter>(parameter);
            var deserialized = XmlDeserialize<Parameter>(xml);

            Assert.IsNotNull(deserialized, "Deserialized object was null");

            int hash1 = parameter.GetHashCode();
            int hash2 = deserialized.GetHashCode();

            Assert.AreEqual(hash1, hash2, "Hash of original object differs from deserialized object");
        }

        [TestMethod]
        public void SerializeDeserialize_Script()
        {
            var script = ScriptGenerator.Generate(1);

            string xml = XmlSerialize<Script<TestGlobals>>(script);
            var deserialized = XmlDeserialize<Script<TestGlobals>>(xml);

            Assert.IsNotNull(deserialized, "Deserialized object was null");

            int hash1 = script.GetHashCode();
            int hash2 = deserialized.GetHashCode();

            Assert.AreEqual(hash1, hash2, "Hash of original object differs from deserialized object");
        }

        private static string XmlSerialize<T>(T obj)
        {
            XmlSerializer serializer = new XmlSerializer(typeof(T));
            using (var writer = new StringWriter())
            {
                serializer.Serialize(writer, obj);
                return writer.ToString();
            }
        }

        private static T XmlDeserialize<T>(string xml)
        {
            XmlSerializer deserializer = new XmlSerializer(typeof(T));
            TextReader reader = new StringReader(xml);
            object obj = deserializer.Deserialize(reader);
            T result = (T)obj;
            reader.Close();

            return result;
        }
    }
}
