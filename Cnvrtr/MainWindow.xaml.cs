using System;
using System.Windows;
using System.Windows.Controls;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using Microsoft.Win32;
using iTextSharp.text.pdf;
using iTextSharp.text;
using Microsoft.Office.Interop.Word;

namespace Cnvrtr
{
    public partial class MainWindow : System.Windows.Window
    {
        public ObservableCollection<FileItem> SelectedFiles { get; set; }

        public MainWindow()
        {
            InitializeComponent();
            DataContext = this;
            SelectedFiles = new ObservableCollection<FileItem>();
        }

        private void BTN_Select(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Multiselect = true;

            if (openFileDialog.ShowDialog() == true)
            {
                string[] filePaths = openFileDialog.FileNames;
                AddFile(filePaths);
            }
        }

        private void AddFile(string[] filePaths)
        {
            foreach (string filePath in filePaths)
            {
                FileInfo fileInfo = new FileInfo(filePath);
                SelectedFiles.Add(new FileItem
                {
                    Name = fileInfo.Name,
                    Size = fileInfo.Length,
                    Path = filePath
                });
            }
        }

        private void BTN_Save(object sender, RoutedEventArgs e)
        {
            if (fileList.AlternationCount == 0)
            {
                MessageBox.Show("No file selected.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            if (ComboBox.SelectedItem == null)
            {
                MessageBox.Show("No format selected.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            string selectedFormat = ((ComboBoxItem)ComboBox.SelectedItem).Content.ToString().ToLower();

            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string saveFolderPath = Path.Combine(desktopPath, "SavedFiles", $"{DateTime.Now:yyyyMMdd_HHmmss}");
            Directory.CreateDirectory(saveFolderPath);

            foreach (FileItem fileItem in SelectedFiles)
            {
                string sourceFilePath = fileItem.Path;
                string destinationFilePath = Path.Combine(saveFolderPath, Path.ChangeExtension(fileItem.Name, selectedFormat));

                if (selectedFormat == "pdf")
                {
                    ConvertToPdf(sourceFilePath, destinationFilePath);
                }
                else if (selectedFormat == "docx")
                {
                    ConvertToDocx(sourceFilePath, destinationFilePath);
                }
            }

            OpenFolder(saveFolderPath);
        }

        private void ConvertToPdf(string sourceFilePath, string destinationFilePath)
        {
            using (FileStream sourceStream = new FileStream(sourceFilePath, FileMode.Open, FileAccess.Read))
            {
                using (iTextSharp.text.Document pdfDocument = new iTextSharp.text.Document())
                {
                    iTextSharp.text.Font font = FontFactory.GetFont("Courier New", 12);

                    PdfWriter pdfWriter = PdfWriter.GetInstance(pdfDocument, new FileStream(destinationFilePath, FileMode.Create));
                    pdfDocument.Open();

                    using (StreamReader reader = new StreamReader(sourceStream))
                    {
                        string line;
                        while ((line = reader.ReadLine()) != null)
                        {
                            pdfDocument.Add(new iTextSharp.text.Paragraph(line, font));
                        }
                    }

                    pdfDocument.Close();
                }
            }
        }

        private void ConvertToDocx(string sourceFilePath, string destinationFilePath)
        {
            Microsoft.Office.Interop.Word.Application wordApplication = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document wordDocument = null;

            try
            {
                wordDocument = wordApplication.Documents.Open(sourceFilePath);

                object fontName = "Courier New";
                object fontSize = 12;

                wordDocument.SaveAs2(destinationFilePath, WdSaveFormat.wdFormatXMLDocument);
            }
            finally
            {
                wordDocument?.Close();
                wordApplication.Quit();

                ReleaseObject(wordDocument);
                ReleaseObject(wordApplication);
            }
        }

        private void ReleaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Error releasing object: " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        private void OpenFolder(string folderPath)
        {
            Process.Start("explorer.exe", folderPath);
        }

        private void ListView_MouseClick(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (fileList.SelectedItem is FileItem selectedItem)
            {
                Process.Start(selectedItem.Path);
            }
        }

        public class FileItem
        {
            public string Name { get; set; }
            public long Size { get; set; }
            public string Path { get; internal set; }
        }
    }
}
