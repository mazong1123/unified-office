using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;

namespace UnifiedOffice.Word.Tests
{
    [TestClass]
    public class ApplicationTest
    {
        Application wordApp = null;

        [TestInitialize]
        public void Setup()
        {
            wordApp = new Application();
        }

        [TestCleanup]
        public void Cleanup()
        {
            wordApp.Quit();

            wordApp = null;

            string projectDirectory = this.GetProjectDirectory();
            DirectoryInfo projectDirectoryInfo = new DirectoryInfo(projectDirectory);

            foreach (FileInfo file in projectDirectoryInfo.GetFiles())
            {
                file.Delete();
            }

            foreach (DirectoryInfo dir in projectDirectoryInfo.GetDirectories())
            {
                dir.Delete(true);
            }
        }

        [TestMethod]
        public void Test_AddDocumentWithDefaultParams()
        {
            // Act
            Document addedDocument = this.wordApp.AddDocument();

            // Assert
            Assert.AreEqual(1, this.wordApp.Documents.Count);
            Assert.AreEqual(this.wordApp.ActiveDocument.Id, addedDocument.Id);
        }

        [TestMethod]
        public void Test_OpenDocumentWithOnlyFileName()
        { 
            // Prepare
            this.wordApp.DisplayAlerts = WdAlertLevel.wdAlertsNone;

            // Act
            string projectDirectory = this.GetProjectDirectory();
            Document openedDocument = this.wordApp.OpenDocument(projectDirectory + @"\misc\test.doc");

            // Assert
            Assert.AreEqual(wordApp.ActiveDocument.Id, openedDocument.Id);
        }

        private string GetProjectDirectory()
        {
            string projectDirectory = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName;

            return projectDirectory;
        }
    }
}
