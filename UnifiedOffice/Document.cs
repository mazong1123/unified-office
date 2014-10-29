using InteropWord = Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Runtime.InteropServices;

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
            InteropWord.Windows windows = this.document.Windows;
            int windowCount = windows.Count;
            for (var i = 1; i <= windowCount; i++)
            {
                InteropWord.Window win = windows[i];

                InteropWord.Panes panes = win.Panes;
                int paneCount = panes.Count;
                for (var j = 1; j <= paneCount; j++)
                {
                    InteropWord.Pane pane = panes[j];
                    var pages = pane.Pages;
                    var pageCount = pages.Count;
                    for (var k = 1; k <= pageCount; k++)
                    {
                        var p = pages[i];
                        var bits = p.EnhMetaFileBits;
                        var target = directory + string.Format(@"\{0}_image.doc", k);
                        using (var ms = new MemoryStream((byte[])(bits)))
                        {
                            var image = System.Drawing.Image.FromStream(ms);
                            var pngTarget = Path.ChangeExtension(target, "png");
                            image.Save(pngTarget, System.Drawing.Imaging.ImageFormat.Png);
                        }

                        Marshal.ReleaseComObject(p);
                        p = null;
                    }

                    Marshal.ReleaseComObject(pages);
                    pages = null;

                    Marshal.ReleaseComObject(pane);
                    pane = null;
                }

                Marshal.ReleaseComObject(panes);
                panes = null;

                Marshal.ReleaseComObject(win);
                win = null;
            }

            Marshal.ReleaseComObject(windows);
            windows = null;
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
