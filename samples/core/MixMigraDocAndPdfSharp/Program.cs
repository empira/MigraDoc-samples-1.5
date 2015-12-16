using System;
using System.Diagnostics;
using System.Globalization;
using MigraDoc.DocumentObjectModel;
using MigraDoc.Rendering;
using PdfSharp.Drawing;
using PdfSharp.Pdf;

namespace MixMigraDocAndPdfSharp
{
    /// <summary>
    /// This sample demonstrates how to mix MigraDoc and PDFsharp.
    /// </summary>
    class Program
    {
        static void Main(string[] args)
        {
            var now = DateTime.Now;
            var filename = "MixMigraDocAndPdfSharp.pdf";
#if DEBUG
            filename = Guid.NewGuid().ToString("D").ToUpper() + ".pdf";
#endif
            var document = new PdfDocument();
            document.Info.Title = "Mixing MigraDoc and PDFsharp Sample";
            document.Info.Author = "Stefan Lange";
            document.Info.Subject = "Created with code snippets that show the use of MigraDoc with PDFsharp";
            document.Info.Keywords = "PDFsharp, MigraDoc";

            SamplePage1(document);

            SamplePage2(document);

            Debug.WriteLine("seconds=" + (DateTime.Now - now).TotalSeconds.ToString(CultureInfo.InvariantCulture));

            // Save the document...
            document.Save(filename);
            // ...and start a viewer.
            Process.Start(filename);
        }

        /// <summary>
        /// Renders a single paragraph.
        /// </summary>
        static void SamplePage1(PdfDocument document)
        {
            var page = document.AddPage();
            var gfx = XGraphics.FromPdfPage(page);
            // HACK²
            gfx.MUH = PdfFontEncoding.Unicode;

            var font = new XFont("Segoe UI", 13, XFontStyle.Bold);

            gfx.DrawString("The following paragraph was rendered using MigraDoc:", font, XBrushes.Black,
                new XRect(100, 100, page.Width - 200, 300), XStringFormats.Center);

            // You always need a MigraDoc document for rendering.
            var doc = new Document();
            var sec = doc.AddSection();
            // Add a single paragraph with some text and format information.
            var para = sec.AddParagraph();
            para.Format.Alignment = ParagraphAlignment.Justify;
            para.Format.Font.Name = "Times New Roman";
            para.Format.Font.Size = 12;
            para.Format.Font.Color = Colors.DarkGray;
            para.Format.Font.Color = Colors.DarkGray;
            para.AddText("Duisism odigna acipsum delesenisl ");
            para.AddFormattedText("ullum in velenit", TextFormat.Bold);
            para.AddText(" ipit iurero dolum zzriliquisis nit wis dolore vel et nonsequipit, velendigna " +
                "auguercilit lor se dipisl duismod tatem zzrit at laore magna feummod oloborting ea con vel " +
                "essit augiati onsequat luptat nos diatum vel ullum illummy nonsent nit ipis et nonsequis " +
                "niation utpat. Odolobor augait et non etueril landre min ut ulla feugiam commodo lortie ex " +
                "essent augait el ing eumsan hendre feugait prat augiatem amconul laoreet. ≤≥≈≠");
            para.Format.Borders.Distance = "5pt";
            para.Format.Borders.Color = Colors.Gold;

            // Create a renderer and prepare (=layout) the document.
            var docRenderer = new DocumentRenderer(doc);
            docRenderer.PrepareDocument();

            // Render the paragraph. You can render tables or shapes the same way.
            docRenderer.RenderObject(gfx, XUnit.FromCentimeter(5), XUnit.FromCentimeter(10), "12cm", para);
        }

        /// <summary>
        /// Renders a whole MigraDoc document scaled to a single PDF page.
        /// </summary>
        static void SamplePage2(PdfDocument document)
        {
            var page = document.AddPage();
            var gfx = XGraphics.FromPdfPage(page);
            // HACK²
            gfx.MUH = PdfFontEncoding.Unicode;

            // Create document from HelloMigraDoc sample.
            var doc = HelloMigraDoc.Documents.CreateDocument();

            // Create a renderer and prepare (=layout) the document.
            var docRenderer = new DocumentRenderer(doc);
            docRenderer.PrepareDocument();

            // For clarity we use point as unit of measure in this sample.
            // A4 is the standard letter size in Germany (21cm x 29.7cm).
            var a4Rect = new XRect(0, 0, A4Width, A4Height);

            var pageCount = docRenderer.FormattedDocument.PageCount;
            for (var idx = 0; idx < pageCount; idx++)
            {
                var rect = GetRect(idx);

                // Use BeginContainer / EndContainer for simplicity only. You can naturally use your own transformations.
                var container = gfx.BeginContainer(rect, a4Rect, XGraphicsUnit.Point);

                // Draw page border for better visual representation.
                gfx.DrawRectangle(XPens.LightGray, a4Rect);

                // Render the page. Note that page numbers start with 1.
                docRenderer.RenderPage(gfx, idx + 1);

                // Note: The outline and the hyperlinks (table of contents) do not work in the produced PDF document.

                // Pop the previous graphical state.
                gfx.EndContainer(container);
            }
        }

        /// <summary>
        /// Calculates thumb rectangle.
        /// </summary>
        static XRect GetRect(int index)
        {
            var rect = new XRect(0, 0, A4Width / 3 * 0.9, A4Height / 3 * 0.9)
            {
                X = (index % 3) * A4Width / 3 + A4Width * 0.05 / 3,
                Y = (index / 3) * A4Height / 3 + A4Height * 0.05 / 3
            };
            return rect;
        }
        static readonly double A4Width = XUnit.FromCentimeter(21).Point;
        static readonly double A4Height = XUnit.FromCentimeter(29.7).Point;
    }
}
