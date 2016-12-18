namespace RoslynScripting.Internal
{
    internal class HashValueAttribute : HashAttribute
    {
        public override void AddTo(ref int hash, object value)
        {
            hash = (hash*23) + HashHelper.HashOf(value);
        }
    }
}
