using InteropWord = Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

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

        public void SaveAs(string fileName = "", WdSaveFormat format = WdSaveFormat.wdFormatDocument)
        {
            this.document.SaveAs2(fileName, format);
        }

        public void SaveToImages(string directory)
        {
            foreach (Microsoft.Office.Interop.Word.Window window in this.document.Windows)
            {
                foreach (Microsoft.Office.Interop.Word.Pane pane in window.Panes)
                {
                    for (var i = 1; i <= pane.Pages.Count; i++)
                    {
                        var bits = pane.Pages[i].EnhMetaFileBits;
                        var target = directory + string.Format(@"\{0}_image.doc", i);
                        using (var ms = new MemoryStream((byte[])(bits)))
                        {
                            var image = System.Drawing.Image.FromStream(ms);
                            var pngTarget = Path.ChangeExtension(target, "png");
                            image.Save(pngTarget, System.Drawing.Imaging.ImageFormat.Png);
                        }
                    }
                }
            }
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

        public void Close()
        {
            ((InteropWord._Document)this.document).Close();
        }
    }
}
