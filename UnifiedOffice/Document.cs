using InteropWord = Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace UnifiedOffice.Word
{
    public class Document
    {
        private InteropWord.Document document = null;
        private Guid id = Guid.Empty;

        public Document()
        {
            this.document = new InteropWord.Document();
            this.id = Guid.NewGuid();
        }

        internal Document(InteropWord.Document interopDocument)
        {
            this.document = interopDocument;
            this.id = Guid.NewGuid();
        }

        public Guid Id
        {
            get 
            {
                return this.id;
            }
        }

        public void SaveAs(string fileName="", WdSaveFormat format = WdSaveFormat.wdFormatDocument)
        {
            this.document.SaveAs2(fileName, format);
        }

        public bool Equals(Document document)
        {
            if (this.Id.Equals(document.id))
            {
                return true;
            }

            return false;
        }

        public bool InteropEquals(InteropWord.Document interopDocument)
        {
            if (this.document.Equals(interopDocument))
            {
                return true;
            }

            return false;
        }
    }
}
