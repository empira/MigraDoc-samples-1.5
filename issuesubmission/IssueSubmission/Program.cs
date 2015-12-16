using System;
using System.Collections.Generic;
using System.Diagnostics;
#if GDI
//using System.Drawing;
#endif
#if WPF
using System.Windows.Media.Imaging;
#endif
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Shapes;
using MigraDoc.Rendering;

namespace BugSubmission
{
    class Program
    {
        static void Main()
        {
            // Create a new MigraDoc document
            Document document = new Document();
            //document.UseCmykColor = true;

            // Add a section to the document
            Section section = document.AddSection();

            // Add a paragraph to the section
            Paragraph paragraph = section.AddParagraph();

            paragraph.Format.Font.Color = Color.FromCmyk(100, 30, 20, 50);

            // Add some text to the paragraph
            paragraph.AddFormattedText("Hello, World!", TextFormat.Bold);

#if GDI
            // Using GDI-specific routines.
            // Make sure to use "#if GDI" for any usings you add for platform-specific code.
            {
            }
#endif

#if WPF
            // Using WPF-specific routines.
            // Make sure to use "#if GDI" for any usings you add for platform-specific code.
            {
            }
#endif

            // Create a renderer for the MigraDoc document.
            PdfDocumentRenderer pdfRenderer = new PdfDocumentRenderer(true);

            // Associate the MigraDoc document with a renderer
            pdfRenderer.Document = document;

            // Layout and render document to PDF
            pdfRenderer.RenderDocument();

            // Save the document...
            const string filename = "HelloWorld.pdf";
            pdfRenderer.PdfDocument.Save(filename);
            // ...and start a viewer.
            Process.Start(filename);
        }
    }
}
