using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Win32;

namespace TryToWork
{
    public class MW :MainWindow
    {
        public Word.Application word_app = new Word.Application();
        public Word.Document wordDoc;

        public void Save(Word.Document word_doc, Word.Application word_app)
        {
            var saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.ShowDialog();
            string filename = saveFileDialog1.FileName;

            try
            {
                word_doc.SaveAs(filename);
                word_doc.Close();
                MessageBox.Show("файл сохранен");
                word_app.Visible = true;

            }
            catch { MessageBox.Show("произошла ошибка"); }
        }

        public string Open(Word.Application word_app)
        {
            var openFileDialog1 = new OpenFileDialog();
            openFileDialog1.ShowDialog();

            string filename = openFileDialog1.FileName;
            return filename;
        }

        public void KolontitulV(Word.Document word_doc)
        {
            foreach (Word.Section section in word_doc.Sections)
            {
                Word.Range headerRange = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                headerRange.Fields.Add(headerRange, Word.WdFieldType.wdFieldPage);
                headerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                headerRange.Font.ColorIndex = Word.WdColorIndex.wdBlue;
                headerRange.Font.Size = 10;
                headerRange.Text = "Верхний колонтитул" + Environment.NewLine + "www.CSharpCoderR.com";
            }
        }

        public void addText( Word.Document word_doc, string Text)
        {
            var para = word_doc.Paragraphs.Add();

            object style_name = "Заголовок 1";
            para.Range.set_Style(ref style_name);
            para.Range.Text += "Кривая хризантемы";
            para.Range.InsertParagraphAfter();

            para.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            //para.Range.Text += richTextBox.Selection.Text;//ПРОБЛЕМА С РИЧЕМ
            para.Range.Text += Text;
            para.Range.InsertParagraphAfter();

            para.Range.Font.Italic = -1;
            para.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            // para.Range.Text += richTextBox.Selection.Text;
            para.Range.Text += Text;
            para.Range.InsertParagraphAfter();



            para.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
            // para.Range.Text += richTextBox.Selection.Text;
            para.Range.Text += Text;
            para.Range.InsertParagraphAfter();
        }



    }
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, RoutedEventArgs e)
        {
            MW Docum = new MW();
            Docum.wordDoc = Docum.word_app.Documents.Open(Docum.Open(Docum.word_app));
            Docum.addText(Docum.wordDoc, /*richTextBox.Selection.Text*/"sdfghj");
            Docum.KolontitulV(Docum.wordDoc);
            Docum.Save(Docum.wordDoc, Docum.word_app);

        }


        private void button_Click(object sender, RoutedEventArgs e)
        {
            string filename = OPEN();


            var word_app = new Word.Application();
            word_app.Visible = false;//отображение ворда во время работы кода

            var wordDoc = word_app.Documents.Open(filename);

            ReplaceWordStub("{}", textBox.Text, wordDoc);
            SAVE(wordDoc, word_app);
        }


        private void button1_Click(object sender, RoutedEventArgs e)
        {
            var word_app = new Word.Application
            {
                Visible = false
            };

            // Создаем документ Word.
            object missing = Type.Missing;

            var word_doc = word_app.Documents.Add();

            // Создаем абзац заголовка.
            var para = word_doc.Paragraphs.Add(ref missing);

            object style_name = "Заголовок 1";
            para.Range.set_Style(ref style_name);
            para.Range.Text += "Кривая хризантемы";
            para.Range.InsertParagraphAfter();

            

            //Добавление верхнего колонтитула
            foreach (Microsoft.Office.Interop.Word.Section section in word_doc.Sections)
            {
                Microsoft.Office.Interop.Word.Range headerRange =
                section.Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                headerRange.Fields.Add(headerRange, Microsoft.Office.Interop.Word.WdFieldType.wdFieldPage);
                headerRange.ParagraphFormat.Alignment =
                Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                headerRange.Font.ColorIndex = Microsoft.Office.Interop.Word.WdColorIndex.wdBlue;
                headerRange.Font.Size = 10;
                headerRange.Text = "Верхний колонтитул" + Environment.NewLine + "www.CSharpCoderR.com";
            }
            

            para.Range.Font.Size = 13;
            para.Range.Font.Bold = -1;

            richTextBox.SelectAll();
            string myText = richTextBox.Selection.Text;

            para.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            para.Range.Text += richTextBox.Selection.Text;
            para.Range.InsertParagraphAfter();

            para.Range.Font.Italic = -1;
            para.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            para.Range.Text += richTextBox.Selection.Text;
            para.Range.InsertParagraphAfter();



            para.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
            para.Range.Text += richTextBox.Selection.Text;
            para.Range.InsertParagraphAfter();


            //richTextBox1.Selection.Text = para.Range.Text;

            //Set myRange = ActiveDocument.Words(1)
            //para.Range.Words(1) = "Dear ";

            

            SAVE(word_doc, word_app);
        }
        private void SAVE(Word.Document word_doc, Word.Application word_app)
        {
            var saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.ShowDialog();

            string filename = saveFileDialog1.FileName;

            try
            {
                word_doc.SaveAs(filename);
                MessageBox.Show("файл сохранен");
                word_app.Visible = true;
            }
            catch { MessageBox.Show("произошла ошибка"); }
        }

        private string OPEN()
        {
            var openFileDialog1 = new OpenFileDialog();
            openFileDialog1.ShowDialog();
            
            string filename = openFileDialog1.FileName;
            return filename;
        }

        private void ReplaceWordStub(string stubToReplace, string text, Word.Document wordDocument)
        {
            var range = wordDocument.Content;
            range.Find.ClearFormatting();
            range.Find.Execute(FindText: stubToReplace, ReplaceWith: text);

        }

       
    }

    
   

    
}
