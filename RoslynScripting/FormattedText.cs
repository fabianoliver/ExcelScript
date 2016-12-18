using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RoslynScripting
{
    [Serializable]
    public struct TextSpan
    {
        public readonly int Start;
        public readonly int End;
        public readonly int Length;

        public TextSpan(int start, int length)
            : this(start, length, start+length)
        {
        }

        public static TextSpan FromBounds(int start, int end)
        {
            var result = new TextSpan(start, end-start, end);
            return result;
        }

        private TextSpan(int start, int length, int end)
        {
            this.Start = start;
            this.Length = length;
            this.End = end;
        }
    }

    [Serializable]
    public class TextFormat
    {
        public readonly Color TextColor;

        public TextFormat(Color TextColor)
        {
            this.TextColor = TextColor;
        }
    }

    [Serializable]
    public class FormattedTextLinePart
    {
        /// <summary>
        /// Line number of the text
        /// </summary>
        public FormattedTextLine Line { get; internal set; }

        /// <summary>
        /// Position of the text relative to its line
        /// </summary>
        public readonly TextSpan LineSpan;

        /// <summary>
        /// The actual text
        /// </summary>
        public readonly string Text;

        /// <summary>
        ///  Format of the text
        /// </summary>
        public readonly TextFormat TextFormat;

        public FormattedTextLinePart(string Text, TextSpan LineSpan, TextFormat TextFormat)
        {
            this.Text = Text;
            this.LineSpan = LineSpan;
            this.TextFormat = TextFormat;
        }

        public override string ToString()
        {
            return Text;
        }
    }

    [Serializable]
    public class FormattedTextLine
    {
        /// <summary>
        /// The text this line is part of
        /// </summary>
        public FormattedText FormattedText { get; internal set; }

        public string Text
        {
            get
            {
                return String.Join(String.Empty, Parts.Select(x => x.Text));
            }
        }

        /// <summary>
        /// Number of the line
        /// </summary>
        public readonly int LineNumber;

        /// <summary>
        /// Formatted parts that make up this line
        /// </summary>
        public readonly IReadOnlyCollection<FormattedTextLinePart> Parts;
        private readonly IList<FormattedTextLinePart> _Parts = new List<FormattedTextLinePart>();

        public FormattedTextLine()
        {
            this.Parts = new ReadOnlyCollection<FormattedTextLinePart>(_Parts);
        }

        public void AddPart(FormattedTextLinePart Part)
        {
            Part.Line = this;
            _Parts.Add(Part);
        }

        public FormattedTextLinePart AppendText(string Text, TextFormat Format)
        {
            int SpanStart = _Parts.Any() ? _Parts.Last().LineSpan.End : 0;
            int SpanEnd = SpanStart + Text.Length;

            var LineSpan = TextSpan.FromBounds(SpanStart, SpanEnd);
            var LinePart = new FormattedTextLinePart(Text, LineSpan, Format);

            AddPart(LinePart);
            return LinePart;
        }

        public override string ToString()
        {
            return Text;
        }
    }

    [Serializable]
    public class FormattedText
    {
        /// <summary>
        /// The lines that make up this text
        /// </summary>
        public readonly IReadOnlyCollection<FormattedTextLine> TextLines;
        private readonly IList<FormattedTextLine> _TextLines = new List<FormattedTextLine>();

        public FormattedText()
        {
            this.TextLines = new ReadOnlyCollection<FormattedTextLine>(_TextLines);
        }

        public void AddLine(FormattedTextLine Line)
        {
            Line.FormattedText = this;
            this._TextLines.Add(Line);
        }

        public FormattedTextLine AppendLine()
        {
            var line = new FormattedTextLine();
            AddLine(line);
            return line;
        }
    }
}
