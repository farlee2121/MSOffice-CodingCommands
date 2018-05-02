using Gma.System.MouseKeyHook;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordCodeCommands
{
    public interface IWordCommandsConfig
    {
        string CommentStartMarker();
        string CommentEndMarker();
        WdColor CommentColor();
        Combination CommentKeyBinding();
        Combination UncommentKeyBinding();
    }
    class WordCommandsConfig : IWordCommandsConfig
    {
        public string CommentStartMarker()
        {
            return "<--";
        }

        public string CommentEndMarker()
        {
            return "-->";
        }

        public WdColor CommentColor()
        {
            return WdColor.wdColorGreen;
        }

        // Can I make this clearly extensible for custom actions?
        public Combination CommentKeyBinding()
        {
            return Combination.TriggeredBy(System.Windows.Forms.Keys.Divide)
                    .With(System.Windows.Forms.Keys.Control);
        }

        public Combination UncommentKeyBinding()
        {
            return Combination.TriggeredBy(System.Windows.Forms.Keys.Divide)
                    .With(System.Windows.Forms.Keys.Shift)
                    .With(System.Windows.Forms.Keys.Control);
        }
        
    }
}
