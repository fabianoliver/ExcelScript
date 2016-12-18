using ScriptingAbstractions;
using System;

namespace RoslynScripting
{
    public class ParameterValue : MarshalByRefObject, IParameterValue
    {
        public IParameter Parameter { get; private set; }
        public Type TransferableValueType { get; private set; }

        public readonly object Value;
        private readonly TransferableValueToOriginalValueConverter Converter;
        private Lazy<object> TransferableValueGetter;
        private bool AllowNullTransferableValues;
       
        public ParameterValue(IParameter Parameter, object Value)
            : this(Parameter, Value, Parameter.Type, x => x, x => x, true)
        {
        }

        public ParameterValue(IParameter Parameter, object Value, Type TransferableType, OriginalValueToTransferableValueConverter originalValueToTransferableValueConverter, TransferableValueToOriginalValueConverter transferableValueToOriginalValueConverter, bool AllowNullTransferableValues = false)
        {
            if (Parameter == null)
                throw new ArgumentNullException(nameof(Parameter));

            if (TransferableType == null)
                throw new ArgumentNullException(nameof(TransferableType));

            if (originalValueToTransferableValueConverter == null)
                throw new ArgumentNullException(nameof(originalValueToTransferableValueConverter));

            if (transferableValueToOriginalValueConverter == null)
                throw new ArgumentNullException(nameof(transferableValueToOriginalValueConverter));

            if ((!originalValueToTransferableValueConverter.Target?.GetType()?.IsSerializable) ?? false)
                throw new ArgumentException("The converter's target type must be serializable in order to use it accross AppDomains. This error often occurs if you reference any out-of-scope variables in your delegate, if its declaring type is not serializable.", nameof(originalValueToTransferableValueConverter));

            if ((!transferableValueToOriginalValueConverter.Target?.GetType()?.IsSerializable) ?? false)
                throw new ArgumentException("The converter's target type must be serializable in order to use it accross AppDomains. This error often occurs if you reference any out-of-scope variables in your delegate, if its declaring type is not serializable.", nameof(transferableValueToOriginalValueConverter));

            this.Value = Value;
            this.TransferableValueType = TransferableType;
            this.AllowNullTransferableValues = AllowNullTransferableValues;
            this.Parameter = Parameter;
            this.Converter = transferableValueToOriginalValueConverter;
            this.TransferableValueGetter = new Lazy<object>(() => originalValueToTransferableValueConverter(this.Value));
        }
        /// <summary>
        /// Gets the original value of the parameter
        /// </summary>
        public object GetOriginalValue()
        {
            RoslynScripting.Parameter.EnsureValueIsCompatible(this.Parameter, this.Value);
            return Value;
        }

        /// <summary>
        /// Gets the value of the object in a format that can be marshalled in between AppDomains.
        /// The returned object will be of type <see cref="TransferableValueType"/>, which in turn will be a type
        /// that is either marshallable by ference (MarshalByRefObject), by value (serializable), or by bleed (string, int etc.)
        /// </summary>
        public object GetTransferableValue()
        {
            var value = TransferableValueGetter.Value;

            if(value != null)
            {
                if (!TransferableValueType.IsAssignableFrom(value.GetType()))
                    throw new InvalidCastException($"The transferable value was found to be of type {value.GetType().Name}, which is incompatible with expected type {TransferableValueType.Name}");
            } else if(!AllowNullTransferableValues)
            {
                throw new InvalidOperationException("The transferable value was null, which is invalid");
            }

            return value;
        }

        /// <summary>
        /// Gets a delegate that must by contract execute in the AppDomain of the caller of this function.
        /// Invoking the returned delegate will convert a value of type <see cref="TransferableValueType"/> to one of type <see cref="IParameter.Type"/>
        /// </summary>
        /// <returns></returns>
        public TransferableValueToOriginalValueConverter GetTransferableValueToOriginalValueConverter()
        {
            return this.Converter;
        }

    }
}
