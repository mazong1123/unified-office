using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace UnifiedOffice.Word.Tests
{
    [TestClass]
    public class DocumentTest
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
        }

        [TestMethod]
        public void Test_SaveAsDocumentWithWordDocxFormat()
        {
            // Prepare
            this.wordApp.DisplayAlerts = WdAlertLevel.wdAlertsNone;

            // Act
            string projectDirectory = this.GetProjectDirectory();
            Document openedDocument = this.wordApp.OpenDocument(projectDirectory + @"\misc\test.doc");

            openedDocument.SaveAs(projectDirectory + @"\misc\test-word.docx", WdSaveFormat.wdFormatDocumentDefault);

            Document savedDocument = this.wordApp.OpenDocument(projectDirectory + @"\misc\test-word.docx");

            // Assert
            Assert.IsNotNull(savedDocument);
        }

        [TestMethod]
        public void Test_SaveAsDocumentWithPDFFormat()
        {
            // Prepare
            this.wordApp.DisplayAlerts = WdAlertLevel.wdAlertsNone;

            // Act
            string projectDirectory = this.GetProjectDirectory();
            Document openedDocument = this.wordApp.OpenDocument(projectDirectory + @"\misc\test.doc");

            openedDocument.SaveAs(projectDirectory + @"\misc\test-pdf.pdf", WdSaveFormat.wdFormatPDF);

            Document savedDocument = this.wordApp.OpenDocument(projectDirectory + @"\misc\test-pdf.pdf");

            // Assert
            Assert.IsNotNull(savedDocument);
        }

        [TestMethod]
        public void Test_SaveToImages()
        {
            // Prepare
            this.wordApp.DisplayAlerts = WdAlertLevel.wdAlertsNone;

            // Act
            string projectDirectory = this.GetProjectDirectory();
            Document openedDocument = this.wordApp.OpenDocument(projectDirectory + @"\misc\test.doc");

            openedDocument.SaveToImages(projectDirectory + @"\misc\images\");

            var files = Directory.GetFiles(projectDirectory + @"\misc\images\");

            // Assert
            Assert.AreEqual(2, files.Count());
        }

        private string GetProjectDirectory()
        {
            string projectDirectory = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName;

            return projectDirectory;
        }
    }
}
