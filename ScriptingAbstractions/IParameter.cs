using System;

namespace ScriptingAbstractions
{
    public interface IParameter : IHasDescription
    {
        /// <summary>
        /// Name of the parameter
        /// </summary>
        string Name { get; set; }

        /// <summary>
        /// Type the value of this parameter must habe
        /// </summary>
        Type Type { get; set; }

        /// <summary>
        /// Defines whether this parameter can be ommited, in which case the <see cref="DefaultValue"/> is returned
        /// </summary>
        bool IsOptional { get; set; }

        /// <summary>
        /// Default value, in case the parameter is optional (see <see cref="IsOptional"/>). By contract, returns value of type <see cref="Type"/>
        /// </summary>
        object DefaultValue { get; set; }
    }
}
