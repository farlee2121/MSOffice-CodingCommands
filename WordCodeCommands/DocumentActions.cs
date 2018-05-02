using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace WordCodeCommands
{
    class DocumentActions
    {
        const string CommentStart = "<--";
        const string CommentEnd = "-->";

        public void CommentLine(Range range)
        {
            //https://social.msdn.microsoft.com/Forums/vstudio/en-US/f87bf140-932e-4de0-ac6c-b30ec06534e4/get-cursor-position-in-current-line-and-get-current-paragraph?forum=vsto
            
            string trailingWhitespace = GetTrailingWhitespace(range.Text);

            range.Text = string.Format($"{CommentStart}{range.Text.Trim()}{CommentEnd}{trailingWhitespace}");
        }

        private string GetTrailingWhitespace(string text)
        {
            Regex trailingWhitespaceRegex = new Regex(@"\s+$");

            string trailingWhitespace = trailingWhitespaceRegex.Match(text).Value;
            return trailingWhitespace;
        }

        public void UncommentLine(Range range)
        {
            if (range.Text.Contains(CommentStart))
            {
                // get first comment start
                int commentStartIndex = range.Text.IndexOf(CommentStart);
                range.Text = range.Text.Remove(commentStartIndex, CommentStart.Length);
            }

            // should I regex this to ensure I only remove when the comment end is at the end of the string?
            // Answer: no maybe they add text after the comment, we still want to remove the comment marker
            if (range.Text.Contains(CommentEnd))
            {
                // get last comment end
                // only removing one instance to allow nested comments
                int commentEndIndex = range.Text.LastIndexOf(CommentEnd);
                range.Text = range.Text.Remove(commentEndIndex, CommentEnd.Length);
            }
        }
    }
}
