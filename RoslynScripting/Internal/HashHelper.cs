using System.Linq;
using System.Reflection;

namespace RoslynScripting.Internal
{
    internal static class HashHelper
    {
        public static int HashOf(object obj)
        {
            if (obj == null)
                return 0;
            else
                return obj.GetHashCode();
        }

        // todo: maybe add caching here, build a dynamic method to do this so don't have to re-query all attributes/properties again every time
        public static int HashOfAnnotated<T>(T obj)
        {
            var properties = typeof(T).GetProperties(BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic)
                .Where(prop => prop.IsDefined(typeof(HashAttribute), false));

            unchecked
            {
                int hash = 27;

                foreach(var property in properties)
                {
                    var attribute = (HashAttribute)property.GetCustomAttributes(typeof(HashAttribute), false).Single();
                    object value = property.GetValue(obj);
                    attribute.AddTo(ref hash, value);
                }

                return hash;
            }
        }
    }
}
