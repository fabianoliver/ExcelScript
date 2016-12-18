using ObjectStorage.Abstractions;
using System;
using System.Text.RegularExpressions;

namespace ExcelScript
{
    public static class HandleNames
    {
        private static readonly Regex HandleRegex = new Regex("^(.*)(:[0-9]*)$");
        private static readonly Regex InvalidChars = new Regex("[^0-9a-zA-Z_]");  // only allow alphabetic characters, letters and underscores

        public static string ToHandle(IStoredObject storedObject)
        {
            return $"{storedObject.Name}:{storedObject.Version}";
        }

        /// <summary>
        /// Returns the Handle name from the given Handle
        /// e.g. Handle = Test:123, return value = Test
        /// Removes invalid characters.
        /// </summary>
        public static string GetNameFrom(string Text)
        {
            // Remove the :12345 at the end, if there is any
            var match = HandleRegex.Match(Text);

            if (match.Success)
                Text = match.Groups[1].Value;

            // Remove al invalid chars
            Text = InvalidChars.Replace(Text, String.Empty); // Remove invalid chars

            return Text;
        }
    }
}
