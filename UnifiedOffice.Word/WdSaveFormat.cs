using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace UnifiedOffice.Word
{
    // Summary:
    //     Specifies the format to use when saving a document.
    public enum WdSaveFormat
    {
        // Summary:
        //     Microsoft Word 97 document format.
        wdFormatDocument97 = 0,
        //
        // Summary:
        //     Microsoft Word format.
        wdFormatDocument = 0,
        //
        // Summary:
        //     Word 97 template format.
        wdFormatTemplate97 = 1,
        //
        // Summary:
        //     Microsoft Word template format.
        wdFormatTemplate = 1,
        //
        // Summary:
        //     Microsoft Windows text format.
        wdFormatText = 2,
        //
        // Summary:
        //     Microsoft Windows text format with line breaks preserved.
        wdFormatTextLineBreaks = 3,
        //
        // Summary:
        //     Microsoft DOS text format.
        wdFormatDOSText = 4,
        //
        // Summary:
        //     Microsoft DOS text with line breaks preserved.
        wdFormatDOSTextLineBreaks = 5,
        //
        // Summary:
        //     Rich text format (RTF).
        wdFormatRTF = 6,
        //
        // Summary:
        //     Unicode text format.
        wdFormatUnicodeText = 7,
        //
        // Summary:
        //     Encoded text format.
        wdFormatEncodedText = 7,
        //
        // Summary:
        //     Standard HTML format.
        wdFormatHTML = 8,
        //
        // Summary:
        //     Web archive format.
        wdFormatWebArchive = 9,
        //
        // Summary:
        //     Filtered HTML format.
        wdFormatFilteredHTML = 10,
        //
        // Summary:
        //     Extensible Markup Language (XML) format.
        wdFormatXML = 11,
        //
        // Summary:
        //     XML document format.
        wdFormatXMLDocument = 12,
        //
        // Summary:
        //     XML template format with macros enabled.
        wdFormatXMLDocumentMacroEnabled = 13,
        //
        // Summary:
        //     XML template format.
        wdFormatXMLTemplate = 14,
        //
        // Summary:
        //     XML template format with macros enabled.
        wdFormatXMLTemplateMacroEnabled = 15,
        //
        // Summary:
        //     Word default document file format. For Microsoft Office Word 2007, this is
        //     the DOCX format.
        wdFormatDocumentDefault = 16,
        //
        // Summary:
        //     PDF format.
        wdFormatPDF = 17,
        //
        // Summary:
        //     XPS format.
        wdFormatXPS = 18,
        //
        // Summary:
        //     Reserved for internal use.
        wdFormatFlatXML = 19,
        //
        // Summary:
        //     Reserved for internal use.
        wdFormatFlatXMLMacroEnabled = 20,
        //
        // Summary:
        //     Reserved for internal use.
        wdFormatFlatXMLTemplate = 21,
        //
        // Summary:
        //     Reserved for internal use.
        wdFormatFlatXMLTemplateMacroEnabled = 22,
        //
        wdFormatOpenDocumentText = 23,
    }
}
