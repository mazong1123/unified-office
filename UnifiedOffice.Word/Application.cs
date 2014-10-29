using InteropWord = Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;

namespace UnifiedOffice.Word
{
    /// <summary>
    /// Represents the Microsoft Office Word application.
    /// </summary>
    public class Application
    {
        private InteropWord.Application app = null;
        private Document activeDocument = null;
        private IList<Document> documents = new List<Document>();

        public Application()
        {
            this.app = new InteropWord.Application();
        }

        public bool Visible
        {
            get
            {
                return this.app.Visible;
            }

            set
            {
                this.app.Visible = value;
            }
        }

        public Document ActiveDocument
        {
            get 
            {
                return this.activeDocument;
            }
        }

        public IList<Document> Documents
        {
            get 
            {
                return this.documents;
            }
        }

        public WdAlertLevel DisplayAlerts
        {
            get
            {
                var displayAlerts = this.app.DisplayAlerts;
                return (WdAlertLevel)displayAlerts;
            }

            set
            {
                this.app.DisplayAlerts = (InteropWord.WdAlertLevel)value;
            }
        }

        public Document AddDocument(string templateName = "", bool openAsTemplate = false, WdNewDocumentType documentType = WdNewDocumentType.wdNewBlankDocument, bool visible = true)
        {
            InteropWord.Document newInteropDocument = this.app.Documents.Add(templateName, openAsTemplate, documentType, visible);
            Document newDocument = new Document(newInteropDocument);
            this.documents.Add(newDocument);

            if (visible)
            {
                this.activeDocument = this.GetUODocument(this.app.ActiveDocument);
            }

            return newDocument;
        }

        public Document OpenDocument(string fileName, bool isReadOnly = true, bool isConfirmConvention = false, bool isVisible = true)
        {
            InteropWord.Document openedInteropDocument = this.app.Documents.Open(fileName, 
                isConfirmConvention, 
                isReadOnly, 
                Type.Missing, 
                Type.Missing,
                Type.Missing,
                Type.Missing,
                Type.Missing,
                Type.Missing,
                Type.Missing,
                Type.Missing,
                isVisible);
            Document openedDocument = this.GetUODocument(openedInteropDocument);

            if (isVisible)
            {
                this.activeDocument = this.GetUODocument(this.app.ActiveDocument);
            }

            return openedDocument;
        }

        public void Quit()
        {
            foreach (Document doc in this.documents)
            {
                doc.Close();
            }

            if (this.app != null)
            {
                ((InteropWord._Application)this.app).Quit();

                Marshal.ReleaseComObject(this.app);
                this.app = null;
            }
        }

        private Document GetUODocument(InteropWord.Document interopDocument)
        {
            Document foundDocument = null;
            foreach (var d in this.documents)
            {
                if (d.InteropEquals(interopDocument))
                {
                    foundDocument = d;
                    break;
                }
            }

            if (foundDocument == null)
            {
                foundDocument = new Document(interopDocument);
                this.documents.Add(foundDocument);
            }

            return foundDocument;
        }
    }
}
