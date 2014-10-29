using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace UnifiedOffice.Word
{
    /// <summary>
    /// Specifies the type of new document to create.
    /// </summary>
    public enum WdNewDocumentType
    {
        /// <summary>
        /// Blank document.
        /// </summary>
        wdNewBlankDocument = 0,

        /// <summary>
        /// Web page.
        /// </summary>
        wdNewWebPage = 1,

        /// <summary>
        /// E-mail message.
        /// </summary>
        wdNewEmailMessage = 2,

        /// <summary>
        /// Frameset.
        /// </summary>
        wdNewFrameset = 3,

        /// <summary>
        /// XML document.
        /// </summary>
        wdNewXMLDocument = 4,
    }
}
