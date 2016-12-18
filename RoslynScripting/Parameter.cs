using ExcelScript.CommonUtilities;
using RoslynScripting.Internal;
using ScriptingAbstractions;
using System;
using System.Xml;
using System.Xml.Schema;
using System.Xml.Serialization;

namespace RoslynScripting
{
    [Serializable]
    public class Parameter : IXmlSerializable, IParameter
    {
        [HashValue]
        public string Name { get; set; }

        [HashValue]
        public Type Type { get; set; }

        [HashValue]
        public bool IsOptional { get; set; }

        [HashValue]
        public object DefaultValue { get; set; }

        [HashValue]
        public string Description { get; set; }

        /// <summary>
        /// Checks consistency of a potential parameter value with its parmater definition. Throws an InvalidOperationException if incompatible.
        /// </summary>
        /// <param name="parameter">Parameter definition that defines constraints</param>
        /// <param name="Value">Value candidate for the parameter</param>
        /// <exception cref="InvalidOperationException">Thrown if the parameter value is incompatible with the parameter definition</exception>
        public static void EnsureValueIsCompatible(IParameter parameter, object Value)
        {
            if (Value == null)
            {
                if (parameter.Type.IsValueType)
                    throw new InvalidOperationException($"Value was null, but {parameter.Type.Name} is a value type");
            }
            else
            {
                Type ValueType = Value.GetType();

                if (!parameter.Type.IsAssignableFrom(ValueType))
                {
                    throw new InvalidOperationException($"Expected type {parameter.Type.Name}, but the assigned value was of type {ValueType.Name}");
                }
            }
        }

        public override int GetHashCode()
        {
            return HashHelper.HashOfAnnotated<Parameter>(this);
        }

        public override string ToString()
        {
            return $"Paramter ({Type.Name} {Name})";
        }

        #region IXmlSerializable

        public XmlSchema GetSchema() {
            return null;
        }

        public void WriteXml(XmlWriter writer)
        {
            writer.Write(nameof(Name), Name);
            writer.Write(nameof(Type), Type);
            writer.Write(nameof(IsOptional), IsOptional);
            writer.Write(nameof(DefaultValue), DefaultValue);
            writer.Write(nameof(Description), Description);
        }

        public void ReadXml(XmlReader reader)
        {
            reader.MoveToCustomStart();

            string _Name;
            Type _Type;
            bool _IsOptional;
            object _DefaultValue;
            string _Description;

            reader.Read(nameof(Name), out _Name);
            reader.Read(nameof(Type), out _Type);
            reader.Read(nameof(IsOptional), out _IsOptional);
            reader.Read(nameof(DefaultValue), out _DefaultValue);
            reader.Read(nameof(Description), out _Description);

            this.Name = _Name;
            this.Type = _Type;
            this.IsOptional = _IsOptional;
            this.DefaultValue = _DefaultValue;
            this.Description = _Description;

            reader.ReadEndElement();
        }

        #endregion
    }
}
