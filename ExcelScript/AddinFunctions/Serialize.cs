using ExcelDna.Integration;
using ExcelScript.Internal;
using ExcelScript.Registration;
using ObjectStorage.Abstractions;
using System;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using System.Xml.Serialization;

namespace ExcelScript
{
    public partial class ExcelScriptAddin
    {
        [ExcelFunction(Name = FunctionPrefix + nameof(Serialize), IsVolatile = false, Description = "Serializes the object into a format which can later be re-loaded (with ExcelScript.Deserialize)")]
        [SuppressInDialog]
        public static object Serialize(
            [ExcelArgument(Description = "Handle to the object which shall be serialized")] string Handle,
            [ExcelArgument(Description = "If true, compressed the output. Signitifantly reduces the size of the output, but obfuscates it (makes it unreadable to humans)")] bool Compress = false,
            [ExcelArgument(Description = "Optional. If a file path is given, the deserialized data will be saved to a file to the given target destination")] string OutputFilePath = null)
        {
            return InternalSerialize(Handle, Compress, OutputFilePath);
        }

        private static object InternalSerialize(string HandleName, bool Compress = false, string OutputFilePath = null)
        {
            HandleName = HandleNames.GetNameFrom(HandleName);
            IStoredObject storedObj;

            if (!m_ObjectStore.GetByName(HandleName, out storedObj))
                throw new Exception($"No object named {HandleName} found");

            bool WriteToFile = !String.IsNullOrWhiteSpace(OutputFilePath);

            var obj = new SerializationWrapper { ContainedObject = storedObj };
            var result = XmlSerialize<SerializationWrapper>(obj);

            if (Compress)
            {
                var compressed = Zip(result);
                XDocument doc = new XDocument(new XElement("CompressedData", new XAttribute("Algorithm", "gzip"), Convert.ToBase64String(compressed)));

                if (WriteToFile)
                {
                    OutputFilePath = Environment.ExpandEnvironmentVariables(OutputFilePath);

                    using (var fileWriter = new FileStream(OutputFilePath, FileMode.Create))
                    {
                        doc.Save(fileWriter);
                    }
                    return $"Written to '{OutputFilePath}' in compressed format.";
                }
                else
                {
                    using (var stringWriter = new StringWriter())
                    {
                        doc.Save(stringWriter);
                        return stringWriter.ToString();
                    }
                }
            }
            else
            {
                if (WriteToFile)
                {
                    File.WriteAllLines(OutputFilePath, new string[] { result });
                    return $"Written to '{OutputFilePath}' in uncompressed XML format.";
                }
                else
                {
                    return result;
                }
            }
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


        // From http://stackoverflow.com/questions/7343465/compression-decompression-string-with-c-sharp
        private static byte[] Zip(string str)
        {
            var bytes = Encoding.UTF8.GetBytes(str);

            using (var msi = new MemoryStream(bytes))
            using (var mso = new MemoryStream())
            {
                using (var gs = new GZipStream(mso, CompressionMode.Compress))
                {
                    CopyTo(msi, gs);
                }

                return mso.ToArray();
            }
        }
    }
}
