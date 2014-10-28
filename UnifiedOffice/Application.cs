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

        public Document AddDocument()
        {
            InteropWord.Document newInteropDocument = this.app.Documents.Add();
            Document newDocument = new Document(newInteropDocument);
            this.documents.Add(newDocument);
            this.activeDocument = newDocument;

            return newDocument;
        }

        public void Quit()
        {
            if (this.app != null)
            {
                ((InteropWord._Application)this.app).Quit();
                this.app = null;
            }
        }
    }
}
