using Word = Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace UnifiedOffice.Word
{
    public class Application
    {
        private Word.Application app = null;
        private WdAlertLevel displayAlerts = WdAlertLevel.wdAlertsAll;

        public Application()
        {
            this.app = new Word.Application();
        }

        public WdAlertLevel DisplayAlerts
        {
            get
            {
                return this.displayAlerts;
            }

            set
            {
                this.displayAlerts = value;
            }
        }

        public void Quit()
        {
            if (this.app != null)
            {
                this.app.Quit();
                this.app = null;
            }
        }
    }
}
