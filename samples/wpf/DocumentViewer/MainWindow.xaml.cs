using System;
using System.Diagnostics;
using System.IO;
using System.Windows;
using Microsoft.Win32;
using MigraDoc.Rendering;
using MigraDoc.DocumentObjectModel.IO;

namespace DocumentViewer
{
    /// <summary>
    /// Interaction logic for Window1.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();

            // Create a new MigraDoc document
            var document = SampleDocuments.CreateSample1();

            // HACK
            var ddl = DdlWriter.WriteToString(document);
            Preview.Ddl = ddl;
        }

        private void Sample1Click(object sender, RoutedEventArgs e)
        {
            var document = SampleDocuments.CreateSample1();
            Preview.Ddl = DdlWriter.WriteToString(document);
        }

        private void Sample2Click(object sender, RoutedEventArgs e)
        {
            Directory.SetCurrentDirectory(GetProgramDirectory());
            var document = SampleDocuments.CreateSample2();
            Preview.Ddl = DdlWriter.WriteToString(document);
        }

        private void CloseClick(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void CreatePdfClick(object sender, RoutedEventArgs e)
        {
            var printer = new PdfDocumentRenderer();
            printer.DocumentRenderer = Preview.Renderer;
            printer.Document = Preview.Document;
            printer.RenderDocument();
            Preview.Document.BindToRenderer(null);
            printer.Save("test.pdf");

            Process.Start("test.pdf");
        }

        private void OpenDdlClick(object sender, RoutedEventArgs e)
        {
            try
            {
                var dialog = new OpenFileDialog
                {
                    CheckFileExists = true,
                    CheckPathExists = true,
                    Filter = "MigraDoc DDL (*.mdddl)|*.mdddl|All Files (*.*)|*.*",
                    FilterIndex = 1,
                    InitialDirectory = Path.GetFullPath(Path.Combine(GetProgramDirectory(), "..\\..\\..\\..\\assets\\ddl"))
                };
                //dialog.RestoreDirectory = true;
                if (dialog.ShowDialog() == true)
                {
                    var document = DdlReader.DocumentFromFile(dialog.FileName);
                    var folder = Path.GetDirectoryName(dialog.FileName);
                    Environment.CurrentDirectory = folder;
                    var ddl = DdlWriter.WriteToString(document);
                    Preview.Ddl = ddl;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, Title);
                Preview.Ddl = null; // TODO has no effect
            }
            finally
            {
                //if (dialog != null)
                //  dialog.Dispose();
            }
            //UpdateStatusBar();
        }

        private string GetProgramDirectory()
        {
            var assembly = System.Reflection.Assembly.GetExecutingAssembly();
            return Path.GetDirectoryName(assembly.Location);
        }
    }
}
