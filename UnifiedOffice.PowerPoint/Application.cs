using InteropPowerPoint = Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace UnifiedOffice.PowerPoint
{
    /// <summary>
    /// Represents the Microsoft Office PowerPoint application.
    /// </summary>
    public class Application
    {
        private InteropPowerPoint.Application app = null;

        public Application()
        {
            this.app = new InteropPowerPoint.Application();
        }
    }
}
