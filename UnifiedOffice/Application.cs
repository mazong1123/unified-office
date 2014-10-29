using InteropWord = Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace UnifiedOffice.Word
{
    public class Application
    {
        private InteropWord.Application app = null;
        private Document activeDocument = null;
        private IList<Document> documents = new List<Document>();

        public Application()
        {
            this.app = new InteropWord.Application();
            
            //this.activeDocument = new Document(this.app.ActiveDocument);
            /*foreach (var interopDoc in this.app.Documents)
            {
                Document doc = new Document((InteropWord.Document)interopDoc);
                this.documents.Add(doc);
            }*/
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
            InteropWord.Document newInteropDocument = this.app.Documents.Add();
            Document newDocument = new Document(newInteropDocument);
            this.documents.Add(newDocument);

            this.activeDocument = this.GetUODocument(this.app.ActiveDocument);

            return newDocument;
        }

        public Document OpenDocument(string fileName)
        {
            InteropWord.Document openedInteropDocument = this.app.Documents.Open(fileName);
            Document openedDocument = this.GetUODocument(openedInteropDocument);

            this.activeDocument = this.GetUODocument(this.app.ActiveDocument);

            return openedDocument;
        }

        public void Quit()
        {
            if (this.app != null)
            {
                ((InteropWord._Application)this.app).Quit();
                this.app = null;
            }
        }

        public Document GetUODocument(InteropWord.Document interopDocument)
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
