using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace UnifiedOffice.Word
{
    /// <summary>
    /// Specifies the format to use when saving a document.
    /// </summary>
    public enum WdSaveFormat
    {
        /// <summary>
        /// Microsoft Word 97 document format.
        /// </summary>
        wdFormatDocument97 = 0,

        /// <summary>
        /// Microsoft Word format.
        /// </summary>
        wdFormatDocument = 0,

        /// <summary>
        /// Word 97 template format.
        /// </summary>
        wdFormatTemplate97 = 1,

        /// <summary>
        /// Microsoft Word template format.
        /// </summary>
        wdFormatTemplate = 1,

        /// <summary>
        /// Microsoft Windows text format.
        /// </summary>
        wdFormatText = 2,

        /// <summary>
        /// Microsoft Windows text format with line breaks preserved.
        /// </summary>
        wdFormatTextLineBreaks = 3,

        /// <summary>
        /// Microsoft DOS text format.
        /// </summary>
        wdFormatDOSText = 4,

        /// <summary>
        /// Microsoft DOS text with line breaks preserved.
        /// </summary>
        wdFormatDOSTextLineBreaks = 5,

        /// <summary>
        /// Rich text format (RTF).
        /// </summary>
        wdFormatRTF = 6,

        /// <summary>
        /// Unicode text format.
        /// </summary>
        wdFormatUnicodeText = 7,

        /// <summary>
        /// Encoded text format.
        /// </summary>
        wdFormatEncodedText = 7,

        /// <summary>
        /// Standard HTML format.
        /// </summary>
        wdFormatHTML = 8,

        /// <summary>
        /// Web archive format.
        /// </summary>
        wdFormatWebArchive = 9,

        /// <summary>
        /// Filtered HTML format.
        /// </summary>
        wdFormatFilteredHTML = 10,

        /// <summary>
        /// Extensible Markup Language (XML) format.
        /// </summary>
        wdFormatXML = 11,

        /// <summary>
        /// XML document format.
        /// </summary>
        wdFormatXMLDocument = 12,

        /// <summary>
        /// XML template format with macros enabled.
        /// </summary>
        wdFormatXMLDocumentMacroEnabled = 13,

        /// <summary>
        /// XML template format.
        /// </summary>
        wdFormatXMLTemplate = 14,

        /// <summary>
        ///  XML template format with macros enabled.
        /// </summary>
        wdFormatXMLTemplateMacroEnabled = 15,

        /// <summary>
        /// Word default document file format. For Microsoft Office Word 2007, this is
        /// the DOCX format.
        /// </summary>
        wdFormatDocumentDefault = 16,

        /// <summary>
        /// PDF format.
        /// </summary>
        wdFormatPDF = 17,

        /// <summary>
        /// XPS format.
        /// </summary>
        wdFormatXPS = 18,

        /// <summary>
        /// Reserved for internal use.
        /// </summary>
        wdFormatFlatXML = 19,

        /// <summary>
        /// Reserved for internal use.
        /// </summary>
        wdFormatFlatXMLMacroEnabled = 20,

        /// <summary>
        /// Reserved for internal use.
        /// </summary>
        wdFormatFlatXMLTemplate = 21,

        /// <summary>
        /// Reserved for internal use.
        /// </summary>
        wdFormatFlatXMLTemplateMacroEnabled = 22,
        
        /// <summary>
        /// Reserved for internal use.
        /// </summary>
        wdFormatOpenDocumentText = 23,
    }
}
