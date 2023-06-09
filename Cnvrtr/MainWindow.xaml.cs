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

        private void BTN_Select(object sender, RoutedEventArgs e) //SELECT
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Multiselect = true;

            if (openFileDialog.ShowDialog() == true)
            {
                string[] filePaths = openFileDialog.FileNames;
                AddFile(filePaths);
            }
        }

        private void AddFile(string[] filePaths) //ADD
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

        private void BTN_Save(object sender, RoutedEventArgs e) //SAVE
        {
            if (SelectedFiles.Count == 0)
            {
                ShowErrorMessage("No file selected.");
                return;
            }

            if (fileList.SelectedItem == null)
            {
                ShowErrorMessage("No file selected.");
                return;
            }

            if (ComboBox.SelectedItem == null)
            {
                ShowErrorMessage("No format selected.");
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
                    if (!ConvertToDocx(sourceFilePath, destinationFilePath))
                    {
                        ShowErrorMessage("Failed to convert to DOCX.");
                        return;
                    }
                }
            }

            OpenFolder(saveFolderPath);
        }

        private void ConvertToPdf(string sourceFilePath, string destinationFilePath) //PDF
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

        private bool ConvertToDocx(string sourceFilePath, string destinationFilePath) //DOCX
        {
            Microsoft.Office.Interop.Word.Application wordApplication = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document wordDocument = null;
            bool success = false;

            try
            {
                wordDocument = wordApplication.Documents.Open(sourceFilePath);

                object fontName = "Courier New";
                object fontSize = 12;

                wordDocument.SaveAs2(destinationFilePath, WdSaveFormat.wdFormatXMLDocument);

                success = true;
            }
            catch (Exception ex)
            {
                ShowErrorMessage("Error converting to DOCX: " + ex.Message);
                success = false;
            }
            finally
            {
                wordDocument?.Close();
                wordApplication.Quit();

                ReleaseObject(wordDocument);
                ReleaseObject(wordApplication);
            }

            return success;
        }

        private void ShowErrorMessage(string message) //Exception while releasing the object
        {
            MessageBox.Show(message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }

        private void ReleaseObject(object obj) //Perform garbage collection
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                ShowErrorMessage("Error releasing object: " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        private void OpenFolder(string folderPath) // Open the folder in File Explorer
        {
            Process.Start("explorer.exe", folderPath);
        }

        private void ListView_MouseClick(object sender, System.Windows.Input.MouseButtonEventArgs e) //Mouse click on listview
        {
            if (fileList.SelectedItem is FileItem selectedItem)
            {
                Process.Start(selectedItem.Path);
            }
        }

        public class FileItem // Represents file list
        {
            public string Name { get; set; }
            public long Size { get; set; }
            public string Path { get; internal set; }
        }
    }
}
