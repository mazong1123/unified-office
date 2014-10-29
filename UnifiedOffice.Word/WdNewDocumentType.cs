using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace UnifiedOffice.Word
{
    // Summary:
    //     Specifies the type of new document to create.
    public enum WdNewDocumentType
    {
        // Summary:
        //     Blank document.
        wdNewBlankDocument = 0,
        //
        // Summary:
        //     Web page.
        wdNewWebPage = 1,
        //
        // Summary:
        //     E-mail message.
        wdNewEmailMessage = 2,
        //
        // Summary:
        //     Frameset.
        wdNewFrameset = 3,
        //
        // Summary:
        //     XML document.
        wdNewXMLDocument = 4,
    }
}
