using iTextSharp.text;
using iTextSharp.text.pdf;
using Microsoft.Win32;
using System.IO;
using System.Windows;

namespace PdfTools
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();

        }

        //========================================
        //
        //  
        //========================================
        private void BSelectSource_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "PDF files (*.pdf)|*.pdf|All files (*.*)|*.*";
            if (openFileDialog.ShowDialog() == true)
            {
                CSourceFile.Text = openFileDialog.FileName;
                BSplitFile.IsEnabled = true;
            }
        }

        //========================================
        //
        //  
        //========================================
        private void BSplitFile_Click(object sender, RoutedEventArgs e)
        {
            int numberOfPages;
            if (!int.TryParse(CInterval.Text, out numberOfPages) || numberOfPages <= 0)
            {
                MessageBox.Show(Application.Current.MainWindow, "Enter number of pages per file", "Invalid page number", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return;
            }

            int suffix               = 1;
            string outputFolderPath  = System.IO.Path.GetDirectoryName(CSourceFile.Text);
            string sourcePdfFilePath = CSourceFile.Text;

            PdfReader reader   = new PdfReader(sourcePdfFilePath);
            string pdfFileName = System.IO.Path.GetFileNameWithoutExtension(sourcePdfFilePath); 

            for (int pageNumber = 1; pageNumber <= reader.NumberOfPages; pageNumber += numberOfPages)
            {
                string outpuPdfFileName = string.Format(pdfFileName + "-{0}", suffix.ToString().PadLeft(3, '0'));
                ExtractAndSaveFile(sourcePdfFilePath, outpuPdfFileName, outputFolderPath, pageNumber, numberOfPages);
                suffix++;
            }
            MessageBox.Show(Application.Current.MainWindow, "Done!", "Split Done", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        //========================================
        //
        //  
        //========================================
        private void ExtractAndSaveFile(string sourcePdfFilePath, string outpuPdfFileName, string outputPath, int startPage, int numberOfPages)
        {
            using (PdfReader reader = new PdfReader(sourcePdfFilePath))
            {
                string outputFilename = System.IO.Path.Combine(outputPath, outpuPdfFileName + ".pdf");
                Document document     = new Document();
                PdfCopy partial       = new PdfCopy(document, new FileStream(outputFilename, FileMode.Create));
                document.Open();

                for (int pagenumber = startPage; pagenumber < (startPage + numberOfPages); pagenumber++)
                {
                    if (reader.NumberOfPages >= pagenumber)
                    {
                        partial.AddPage(partial.GetImportedPage(reader, pagenumber));
                    }
                    else
                    {
                        break;
                    }
                }
                document.Close();
            }
        }
    }
}