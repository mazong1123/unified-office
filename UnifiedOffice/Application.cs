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

        public Application()
        {
            this.app = new InteropWord.Application();
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
