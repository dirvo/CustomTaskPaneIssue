using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;

namespace VstoWord2010Addin
{
    public partial class ThisAddIn
    {
        private void ThisAddInStartup(object sender, EventArgs e)
        {
            AddTaskPane();

            Application.DocumentOpen += ApplicationDocumentOpen;
            ((Word.ApplicationEvents4_Event)Application).NewDocument += ApplicationNewDocument;
        }

        private void AddTaskPane()
        {
            if (this.CustomTaskPanes.Any(ctp => ctp.Window == Application.ActiveWindow))
                return; // prevent adding multiple panes on recycled windows

            var taskPane = this.CustomTaskPanes.Add(new UserControl1(), "TEST");
            taskPane.Visible = true;
            taskPane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionFloating;
            taskPane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionLeft;
        }

        private void ApplicationDocumentOpen(Word.Document doc)
        {
            AddTaskPane();
        }

        private void ApplicationNewDocument(Word.Document doc)
        {
            AddTaskPane();
        }


        private void ThisAddInShutdown(object sender, EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += ThisAddInStartup;
            this.Shutdown += ThisAddInShutdown;
        }

        #endregion
    }
}
