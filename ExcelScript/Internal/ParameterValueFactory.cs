using ExcelDna.Integration;
using RoslynScripting;
using ScriptingAbstractions;
using ScriptingAbstractions.Factory;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = NetOffice.ExcelApi;

namespace ExcelScript.Internal
{
    internal class ConverterInfo
    {
        public readonly TransferableValueToOriginalValueConverter transferableValueToOriginalValueConverter;
        public readonly OriginalValueToTransferableValueConverter originalValueToTransferableValueConverter;
        public readonly Type TransferableValueType;

        public ConverterInfo(TransferableValueToOriginalValueConverter transferableValueToOriginalValueConverter, OriginalValueToTransferableValueConverter originalValueToTransferableValueConverter, Type TransferableValueType)
        {
            this.transferableValueToOriginalValueConverter = transferableValueToOriginalValueConverter;
            this.originalValueToTransferableValueConverter = originalValueToTransferableValueConverter;
            this.TransferableValueType = TransferableValueType;
        }

    }

    public class ParameterValueFactory : IParameterValueFactory
    {
        private static readonly IDictionary<Type, ConverterInfo> Converters = CreateConverters();

        public IParameterValue CreateFor(IParameter Parameter, object WithValue)
        {
            var converter = GetMarshalConverterFor(Parameter);
            IParameterValue result;

            if(converter == null)
            {
                result = new ParameterValue(Parameter, WithValue);
            } else
            {
                result = new ParameterValue(Parameter, WithValue, converter.TransferableValueType, converter.originalValueToTransferableValueConverter, converter.transferableValueToOriginalValueConverter);
            }

            return result;
        }



        /// <summary>
        /// Gets the ConverterInfo needed to convert for marsalling the given parameter.
        /// If the type needs conversion but no appropriate converter is registered, throws an exception
        /// </summary>
        /// <param name="Parameter"></param>
        /// <returns></returns>
        private static ConverterInfo GetMarshalConverterFor(IParameter Parameter)
        {
            if (!NeedsMarshalConverter(Parameter))
                return null;

            if (!Converters.ContainsKey(Parameter.Type))
                throw new InvalidOperationException($"Parameter value needs conversion, however did not find any converter to convert parameter type {Parameter.Type.Name} into a marshallable value");

            var result = Converters[Parameter.Type];
            return result;
        }


        private static bool NeedsMarshalConverter(IParameter Parameter)
        {
            if (Parameter.Type.IsValueType)
                return false;  // marshals by value

            if (Parameter.Type == typeof(string))
                return false; // marshals by bleed

            if (typeof(MarshalByRefObject).IsAssignableFrom(Parameter.Type))
                return false; // marshals by ref

            if (Parameter.Type.IsSerializable)
                return false; // marshal by value - todo: double check this ALWAYS works with any .IsSerializable type

            return true;
        }

        private static IDictionary<Type, ConverterInfo> CreateConverters()
        {
            var result = new Dictionary<Type, ConverterInfo>();

            var rangeConverter = new ConverterInfo(TransferableValueToRange, RangeToTransferableValue, typeof(string));
            result.Add(typeof(Excel.Range), rangeConverter);
            return result;
        }


        #region Converter Functions

        // this will be executing on the host's appdomain
        private static Excel.Range TransferableValueToRange(object transferableValue)
        {
            string address = (string)transferableValue;
            var application = new Excel.Application(null, ExcelDnaUtil.Application);
            var range = application.Range(address);

            range.OnDispose += (args) =>
            {
                range.Application.Dispose();
            };

            return range;
        }

        private static string RangeToTransferableValue(object OriginalValue)
        {
            var range = (Excel.Range)OriginalValue;
            string result = range.Address(true, true, NetOffice.ExcelApi.Enums.XlReferenceStyle.xlA1, true);
            return result;
        }

        #endregion
    }
}
