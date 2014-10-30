using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace UnifiedOffice.PowerPoint.Tests
{
    [TestClass]
    public class ApplicationTest
    {
        Application pptApp = null;

        [TestInitialize]
        public void Setup()
        {
            this.pptApp = new Application();
        }

        [TestCleanup]
        public void Cleanup()
        {
            this.pptApp.Quit();

            this.pptApp = null;
        }
    }
}
