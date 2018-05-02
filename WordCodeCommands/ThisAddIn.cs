using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using Gma.System.MouseKeyHook;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;

namespace WordCodeCommands
{
    public partial class ThisAddIn
    {
        IKeyboardMouseEvents keyboardHook;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            
            DocumentActions documentActions = new DocumentActions();
            Dictionary<Combination, System.Action> bindings = new Dictionary<Combination, System.Action>()
            {
                {
                    Combination.TriggeredBy(System.Windows.Forms.Keys.Divide)
                    .With(System.Windows.Forms.Keys.Control),
                    // how do I want to handle ab
                    () =>{documentActions.CommentLine(GetCurrentRange(Application)); }
                },
                {
                    Combination.TriggeredBy(System.Windows.Forms.Keys.Divide)
                    .With(System.Windows.Forms.Keys.Shift)
                    .With(System.Windows.Forms.Keys.Control),
                    () =>{documentActions.UncommentLine(GetCurrentRange(Application)); }
                },
            };

            keyboardHook = Gma.System.MouseKeyHook.Hook.GlobalEvents();
            keyboardHook.OnCombination(bindings);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            keyboardHook.Dispose();
        }

        private Range GetCurrentRange(Word.Application application)
        {
            return application.Selection.Paragraphs[1].Range;
        }
        
        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
