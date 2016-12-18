using RoslynScripting;
using RoslynScriptingTests.Types;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RoslynScriptingTests.Generators
{
    public static class ParameterGenerator
    {
        public static Parameter Generate(object DefaultValue, int seed = 1)
        {
            var result = new Parameter
            {
                Name = $"Parameter{seed}",
                IsOptional = seed % 2 == 0,
                DefaultValue = DefaultValue,
                Description = $"Parameter Description {seed} with come <xml ? @ .x-unfriendly{Environment.NewLine} signs\r\n"
            };

            return result;
        }

        public static Parameter Generate(int seed = 1)
        {
            object DefaultValue = new SimpleSerializableType
            {
                Value1 = $"Simple type {seed}",
                Value2 = seed
            };
            return Generate(DefaultValue, seed);
        }
    }
}
