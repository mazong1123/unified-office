using InteropPowerPoint = Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;

namespace UnifiedOffice.PowerPoint
{
    /// <summary>
    /// Represents the Microsoft Office PowerPoint application.
    /// </summary>
    public class Application
    {
        private InteropPowerPoint.Application app = null;
        private IList<Presentation> presentations = new List<Presentation>();

        public Application()
        {
            this.app = new InteropPowerPoint.Application();
        }

        public Presentation OpenPresentation(string fileName, bool isReadOnly = true, bool isConfirmConvention = false, bool isVisible = true)
        {
            var isReadOnlyMsi = Microsoft.Office.Core.MsoTriState.msoTrue;
            if (!isReadOnly)
            {
                isReadOnlyMsi = Microsoft.Office.Core.MsoTriState.msoFalse;
            }

            var isVisibleMsi = Microsoft.Office.Core.MsoTriState.msoTrue;
            if (!isVisible)
            {
                isVisibleMsi = Microsoft.Office.Core.MsoTriState.msoFalse;
            }

            InteropPowerPoint.Presentation openedInteropPresentation = this.app.Presentations.Open(fileName, isReadOnlyMsi, Microsoft.Office.Core.MsoTriState.msoFalse, isVisibleMsi);

            Presentation openedPresentation = this.GetUOPresentation(openedInteropPresentation);

            return openedPresentation;
        }

        public void Quit()
        {
            if (this.app != null)
            {
                ((InteropPowerPoint._Application)this.app).Quit();

                Marshal.ReleaseComObject(this.app);
                this.app = null;
            }
        }

        private Presentation GetUOPresentation(InteropPowerPoint.Presentation interopPresentation)
        {
            Presentation foundPresentation = null;
            foreach (var p in this.presentations)
            {
                if (p.InteropEquals(interopPresentation))
                {
                    foundPresentation = p;
                    break;
                }
            }

            if (foundPresentation == null)
            {
                foundPresentation = new Presentation(interopPresentation);
                this.presentations.Add(foundPresentation);
            }

            return foundPresentation;
        }
    }
}
