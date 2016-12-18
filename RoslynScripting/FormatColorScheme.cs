using System;
using System.Drawing;
using System.Linq;

namespace RoslynScripting
{
    [Serializable]
    public struct FormatColorScheme
    {
        public readonly Color Keyword;
        public readonly Color ClassName;
        public readonly Color Text;
        public readonly Color StringVerbatim;
        public readonly Color EnumName;
        public readonly Color Comment;
        public readonly Color Number;
        public readonly Color Punctuation;
        public readonly Color StructName;
        public readonly Color Operator;
        public readonly Color Identifier;
        public readonly Color PreprocessorKeyword;
        public readonly Color Unknown;

        public FormatColorScheme(Color Keyword, Color ClassName, Color Text, Color StringVerbatim, Color EnumName, Color Comment, Color Number, Color Punctuation, Color StructName, Color Operator, Color Identifier, Color PreprocessorKeyword, Color Unknown)
        {
            this.Keyword = Keyword;
            this.ClassName = ClassName;
            this.Text = Text;
            this.StringVerbatim = StringVerbatim;
            this.EnumName = EnumName;
            this.Comment = Comment;
            this.Number = Number;
            this.Punctuation = Punctuation;
            this.StructName = StructName;
            this.Operator = Operator;
            this.Identifier = Identifier;
            this.PreprocessorKeyword = PreprocessorKeyword;
            this.Unknown = Unknown;
       
        }

        internal int GetColorIndexForKeyword(string keyword)
        {
            switch (keyword)
            {
                case "keyword":
                    return 1;
                case "class name":
                    return 2;
                case "string":
                    return 3;
                case "string - verbatim":
                    return 4;
                case "enum name":
                    return 5;
                case "comment":
                    return 6;
                case "number":
                    return 7;
                case "punctuation":
                    return 8;
                case "struct name":
                    return 9;
                case "operator":
                    return 10;
                case "identifier":
                    return 11;
                case "preprocessor keyword":
                    return 12;
                default:
                    return 13;
            }
        }

        internal Color GetColorForKeyword(string keyword)
        {
            var index = GetColorIndexForKeyword(keyword) - 1;
            return GetColorsInRtfOrder()[index];
        }

        internal Color[] GetColorsInRtfOrder()
        {
            return new Color[] { Keyword, ClassName, Text, StringVerbatim, EnumName, Comment, Number, Punctuation, StructName, Operator, Identifier, PreprocessorKeyword, Unknown };
        }

        public readonly static FormatColorScheme LightTheme = new FormatColorScheme(Color.Blue, Color.LightSeaGreen, Color.DarkRed, Color.DarkRed, Color.LightSeaGreen, Color.Green, Color.Black, Color.Black, Color.LightSeaGreen, Color.Black, Color.Black, Color.Black, Color.Black);
    }

}
