using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace UnifiedOffice.Word.Tests
{
    [TestClass]
    public class ApplicationTest
    {
        [TestMethod]
        public void Test_ApplicationInitialization()
        {
            Application wordApp = new Application();

            Assert.IsTrue(true);
        }

        [TestMethod]
        public void Test_AddDocumentWithDefaultParams()
        {
            // Prepare
            Application wordApp = new Application();

            // Act
            Document addedDocument = wordApp.AddDocument();

            // Assert
            Assert.AreEqual(1, wordApp.Documents.Count);
            Assert.AreEqual(wordApp.ActiveDocument.Id, addedDocument.Id);
        }
    }
}
