using System;
using System.Diagnostics;
using MigraDoc.Rendering;

namespace HelloMigraDoc
{
    /// <summary>
    /// This sample shows some features of MigraDoc.
    /// </summary>
    class Program
    {
        static void Main(string[] args)
        {
            // Create a MigraDoc document.
            var document = Documents.CreateDocument();

            //var ddl = MigraDoc.DocumentObjectModel.IO.DdlWriter.WriteToString(document);
            MigraDoc.DocumentObjectModel.IO.DdlWriter.WriteToFile(document, "MigraDoc.mdddl");

            var renderer = new PdfDocumentRenderer(true);
            renderer.Document = document;

            renderer.RenderDocument();

            // Save the document...
#if DEBUG
            var filename = Guid.NewGuid().ToString("N").ToUpper() + ".pdf";
#else
            var filename = "HelloMigraDoc.pdf";
#endif
            renderer.PdfDocument.Save(filename);
            // ...and start a viewer.
            Process.Start(filename);
        }
    }
}
