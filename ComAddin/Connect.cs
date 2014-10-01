using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using Extensibility;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;

namespace ComAddin
{
    [ComVisible(true)]
    [ProgId("ComAddin.Connect")]
    [Guid("9103146C-A423-4E15-BFD0-BB03A5D9EBDB")]
    public class Connect : StandardOleMarshalObject,
        IDTExtensibility2,
        ICustomTaskPaneConsumer
    {
        public Application Application { get; set; }

        public ICTPFactory TaskPaneFactory { get; set; }

        public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
        {
            Application = (Application) application;

            Application.DocumentOpen += ApplicationDocumentOpen;
            ((ApplicationEvents4_Event)Application).NewDocument += ApplicationNewDocument;
        }

        private void AddTaskPane(Window window)
        {
            var taskPane = window != null
                ? TaskPaneFactory.CreateCTP("ComAddin.UserControl", "TEST", window)
                : TaskPaneFactory.CreateCTP("ComAddin.UserControl", "TEST");
            taskPane.Visible = true;
            taskPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionLeft;
        }

        private void ApplicationDocumentOpen(Document doc)
        {
            AddTaskPane(doc.ActiveWindow);
        }

        private void ApplicationNewDocument(Document doc)
        {
            AddTaskPane(doc.ActiveWindow);
        }

        public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
        {
        }

        public void OnAddInsUpdate(ref Array custom)
        {
        }

        public void OnStartupComplete(ref Array custom)
        {
        }

        public void OnBeginShutdown(ref Array custom)
        {
        }

        public void CTPFactoryAvailable(ICTPFactory ctpFactoryInst)
        {
            TaskPaneFactory = ctpFactoryInst;

            AddTaskPane(Application.ActiveWindow);
        }
    }
}
