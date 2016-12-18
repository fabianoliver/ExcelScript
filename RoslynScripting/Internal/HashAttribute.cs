using System;

namespace RoslynScripting.Internal
{
    [AttributeUsage(AttributeTargets.Field | AttributeTargets.Property)]
    internal abstract class HashAttribute : Attribute
    {
        public abstract void AddTo(ref int hash, object value);
    }
}
