using Microsoft.Win32;
using System.IO;
using System.Windows;
using System.Windows.Documents;
using Spire.Doc;
using Spire.Doc.Documents;
using System.Windows.Controls;

namespace Word
{
    public partial class Redactor : Window
    {
        public Redactor()
        {
            InitializeComponent();
            this.WindowStartupLocation = WindowStartupLocation.CenterScreen;
            this.MinWidth = 560;
            this.MinHeight = 600;
        }

        private void Save_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Rich Text Format (*.rtf)|*.rtf|Word Document (*.docx)|*.docx|All files (*.*)|*.*";
            if (saveFileDialog.ShowDialog() == true)
            {
                string fileExtension = Path.GetExtension(saveFileDialog.FileName).ToLower();
                if (fileExtension == ".rtf")
                {
                    SaveAsRtf(saveFileDialog.FileName);
                }
                else if (fileExtension == ".docx")
                {
                    SaveAsDocx(saveFileDialog.FileName);
                }
                else
                {
                    MessageBox.Show("Unsupported file format", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void SaveAsRtf(string fileName)
        {
            using (FileStream fileStream = new FileStream(fileName, FileMode.Create))
            {
                System.Windows.Documents.TextRange range = new System.Windows.Documents.TextRange(RichTextBox.Document.ContentStart, RichTextBox.Document.ContentEnd);
                range.Save(fileStream, DataFormats.Rtf);
            }
            MessageBox.Show("Файл успешно сохранен как RTF!", "Сохранение", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void SaveAsDocx(string fileName)
        {
            // Создание нового документа
            Document document = new Document();

            // Создание нового раздела и добавление его в документ
            Spire.Doc.Section section = document.AddSection();

            // Создание нового абзаца
            Spire.Doc.Documents.Paragraph paragraph = section.AddParagraph();

            // Получение текста из RichTextBox
            System.Windows.Documents.TextRange textRange = new System.Windows.Documents.TextRange(RichTextBox.Document.ContentStart, RichTextBox.Document.ContentEnd);

            // Добавление текста в абзац
            paragraph.AppendText(textRange.Text);

            // Сохранение документа
            document.SaveToFile(fileName, FileFormat.Docx);

            MessageBox.Show("Файл успешно сохранен как DOCX!", "Сохранение", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            SendWindow sending = new SendWindow();
            sending.Show();
        }
    }
}