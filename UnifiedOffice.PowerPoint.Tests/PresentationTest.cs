using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace UnifiedOffice.PowerPoint.Tests
{
    [TestClass]
    public class PresentationTest
    {
        Application pptApp = null;

        [TestInitialize]
        public void Setup()
        {
            this.pptApp = new Application();

            string projectDirectory = this.GetProjectDirectory();
            DirectoryInfo projectDirectoryInfo = new DirectoryInfo(projectDirectory + @"\misc\images");

            foreach (FileInfo file in projectDirectoryInfo.GetFiles())
            {
                file.Delete();
            }

            foreach (DirectoryInfo dir in projectDirectoryInfo.GetDirectories())
            {
                dir.Delete(true);
            }
        }

        [TestCleanup]
        public void Cleanup()
        {
            this.pptApp.Quit();

            this.pptApp = null;
        }

        [TestMethod]
        public void Test_SaveToImages()
        {
            // Prepare

            // Act
            string projectDirectory = this.GetProjectDirectory();
            Presentation openedPPT = this.pptApp.OpenPresentation(projectDirectory + @"\misc\1.pptx", true, false, false);

            openedPPT.SaveToImages(projectDirectory + @"\misc\images\");

            var files = Directory.GetFiles(projectDirectory + @"\misc\images\");

            // Assert
            Assert.AreEqual(1, files.Count());
        }

        [TestMethod]
        public void Test_SaveToImages2()
        {
            // Prepare

            // Act
            string projectDirectory = this.GetProjectDirectory();
            Presentation openedPPT = this.pptApp.OpenPresentation(projectDirectory + @"\misc\2.ppt", true, false, false);

            openedPPT.SaveToImages(projectDirectory + @"\misc\images\");

            var files = Directory.GetFiles(projectDirectory + @"\misc\images\");

            // Assert
            Assert.AreEqual(37, files.Count());
        }

        private string GetProjectDirectory()
        {
            string projectDirectory = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName;

            return projectDirectory;
        }
    }
}
