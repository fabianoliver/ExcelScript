using System.Collections;

namespace RoslynScripting.Internal
{
    internal class HashEnumerable : HashAttribute
    {
        public override void AddTo(ref int hash, object value)
        {
            IEnumerable enumerable = value as IEnumerable;

            unchecked
            {
                if (enumerable == null)
                {
                    hash = (hash * 23) + HashHelper.HashOf(null);
                }
                else
                {
                    foreach (var item in enumerable)
                    {
                        hash = (hash * 23) + HashHelper.HashOf(item);
                    }
                }
            }
            
        }
    }
}
