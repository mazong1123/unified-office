﻿using InteropWord = Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Runtime.InteropServices;
using System.Drawing.Imaging;

namespace UnifiedOffice.Word
{
    /// <summary>
    /// Represents a Microsoft Office Word document.
    /// </summary>
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

        /// <summary>
        /// Save word document to images. The type of output images is png.
        /// </summary>
        /// <param name="directory">Specifies the directory for output images.</param>
        public void SaveToImages(string directory)
        {
            InteropWord.Windows windows = this.document.Windows;
            int windowCount = windows.Count;
            for (var i = 1; i <= windowCount; i++)
            {
                InteropWord.Window win = windows[i];
                InteropWord.View windowsView = win.View;

                // Pages can only be retrieved in print layout view.
                windowsView.Type = InteropWord.WdViewType.wdPrintView;

                InteropWord.Panes panes = win.Panes;
                int paneCount = panes.Count;
                for (var j = 1; j <= paneCount; j++)
                {
                    InteropWord.Pane pane = panes[j];
                    var pages = pane.Pages;
                    var pageCount = pages.Count;
                    for (var k = 1; k <= pageCount;)
                    {
                        InteropWord.Page p = null;

                        try {
                            p = pages[k];
                        }
                        catch
                        {
                            // pages[k] sometimes throws exception: 'System.Runtime.InteropServices.COMException: The requested member of the collection does not exist'.
                            // This is a workaround for this issue.
                            continue;
                        }

                        var bits = p.EnhMetaFileBits;
                        var target = directory + string.Format(@"\{0}_image.doc", k);
                        using (var ms = new MemoryStream((byte[])(bits)))
                        {
                            var image = System.Drawing.Image.FromStream(ms);
                            var imageTarget = Path.ChangeExtension(target, "png");
                            image.Save(imageTarget, ImageFormat.Png);
                        }

                        Marshal.ReleaseComObject(p);
                        p = null;

                        k++;
                    }

                    Marshal.ReleaseComObject(pages);
                    pages = null;

                    Marshal.ReleaseComObject(windowsView);
                    windowsView = null;

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
            if (this.document != null)
            {
                ((InteropWord._Document)this.document).Close();

                Marshal.ReleaseComObject(this.document);
                this.document = null;
                this.id = Guid.Empty;
            }
        }
    }
}
