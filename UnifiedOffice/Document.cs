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
    }
}
