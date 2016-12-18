using System;

namespace ScriptingAbstractions
{
    /// <summary>
    /// Converts a transferblale value created by <see cref="IParameterValue.GetTransferableValue"/> into a copy of <see cref="IParameterValue.GetOriginalValue"/>
    /// </summary>
    /// <param name="TransferableValue"></param>
    /// <returns>A value of type <see cref="IParameter.Type"/> representing a copy of the original object</returns>
    public delegate object TransferableValueToOriginalValueConverter(object TransferableValue);

    /// <summary>
    /// Converts a value into a transferable value, which must be of type <see cref="IParameterValue.TransferableValueType"/>
    /// <param name="OriginalValue">the value which is to be converted</param>
    /// <returns>A value of <see cref="IParameterValue.TransferableValueType"/>, which must be marshallable by reference, value or bleed</returns>
    public delegate object OriginalValueToTransferableValueConverter(object OriginalValue);



    /// <summary>
    /// Represents the value of a parameter.
    /// This may need some explaining: Why all the hassle with transferable values, converters and such?
    /// Well, because effectively, we are likely to instantiate implementations of IParameterValue on a certain appdomain, but consume its values in another.
    /// This would cause issues if the value is not marshallable (by type/reference/bleed). In that case, we need to pass over a converted value which IS marshallable,
    /// and then on the other app domain get a converter to convert this back to the original type - where "converting back" essentially means "creating a copy of the object in the new app domain"
    /// </summary>
    public interface IParameterValue
    {

        Type TransferableValueType { get; }
        IParameter Parameter { get; }


        /// <summary>
        /// Gets the original value of the parameter
        /// </summary>
        object GetOriginalValue();

        /// <summary>
        /// Gets the value of the object in a format that can be marshalled in between AppDomains.
        /// The returned object will be of type <see cref="TransferableValueType"/>, which in turn will be a type
        /// that is either marshallable by ference (MarshalByRefObject), by value (serializable), or by bleed (string, int etc.)
        /// </summary>
        object GetTransferableValue();

        /// <summary>
        /// Gets a delegate that must by contract execute in the AppDomain of the caller of this function.
        /// Invoking the returned delegate will convert a value of type <see cref="TransferableValueType"/> to one of type <see cref="IParameter.Type"/>
        /// </summary>
        /// <returns></returns>
        TransferableValueToOriginalValueConverter GetTransferableValueToOriginalValueConverter();
    }

}
