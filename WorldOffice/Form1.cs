using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WorldOffice
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        Microsoft.Office.Interop.Word.Application application;
        Microsoft.Office.Interop.Word.Document document;
        Microsoft.Office.Interop.Word.Paragraph paragraphone1, paragraphone2;
        Microsoft.Office.Interop.Word.Range rng;
        Microsoft.Office.Interop.Word.Range rngg;
        String surname, name, patronymic;
        String Boss = "Иванова И.И.";

        private void button1_Click(object sender, EventArgs e)
        {
            application = new Microsoft.Office.Interop.Word.Application();
            try
            {
                document = application.Application.Documents.Add();
                document.Paragraphs.SpaceAfter = 0;
                document.Paragraphs.SpaceBefore = 0;
                document.Paragraphs.Space1();

                surname = textBox1.Text;
                name = textBox2.Text;
                patronymic = textBox3.Text;

                document.Content.Paragraphs.Add(Type.Missing);

                document.Paragraphs[1].Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                document.Paragraphs[1].Range.Font.Name = "Times New Roman";
                document.Paragraphs[1].Range.Font.Size = 16;
                document.Paragraphs[1].Range.Font.Bold = 1;

                document.Paragraphs[1].Range.Text = "Общество с ограниченной ответственностью"+"\r\n"+"\"Интеграл\"" + "\r\n" + "\r\n";

                rng = document.Paragraphs[1].Range;
                rng.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                rng = document.Paragraphs[2].Range;
                rng.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;

                paragraphone2 = document.Content.Paragraphs.Add(Type.Missing);

                paragraphone2.Range.Font.Name = "Times New Roman";
                paragraphone2.Range.Font.Size = 12;
                paragraphone2.Range.Font.Italic = 1;
                paragraphone2.Range.Text = "Начальнику отдела по информационной поддержке" + "\r\n" + "и развитию"+ "\r\n" + "Петрову С.Е." + "\r\n" + "От начальника отделения" + "\r\n" + Boss + "\r\n" + "\r\n";

                rng = document.Range(document.Sentences[3].Start, document.Sentences[7].End);
                rng.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphRight;

                paragraphone2 = document.Content.Paragraphs.Add(Type.Missing);

                paragraphone2.Range.Font.Name = "Times New Roman";
                paragraphone2.Range.Font.Size = 12;
                paragraphone2.Range.Font.Bold = 1;

                paragraphone2.Range.Text = "Служебная записка" + "\r\n" + "по привлечению к работе по замене оборудования" + "\r\n" + "в выходной день" + "\r\n" + "\r\n";

                rng = document.Range(document.Sentences[8].Start, document.Sentences[10].End);
                rng.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;

                paragraphone2 = document.Content.Paragraphs.Add(Type.Missing);
                paragraphone2.Range.Font.Name = "Times New Roman";
                paragraphone2.Range.Font.Size = 12;

                paragraphone2.Range.Text = "В связи с производственной необходимостью прошу привлечь" + " " + surname + " " + name + " " + patronymic + "\r\n" + "\r\n";
                rng = document.Range(document.Sentences[11].Start, document.Sentences[11].End);
                rng.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft;

                paragraphone2 = document.Content.Paragraphs.Add(Type.Missing);

                paragraphone2.Range.Font.Name = "Times New Roman";
                paragraphone2.Range.Font.Size = 12;

                paragraphone2.Range.Text = "К сверхурочной работе:" + "\r\n" + "\r\n";
                rng = document.Range(document.Sentences[12].Start, document.Sentences[12].End);
                rng.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft;
                application.Visible = true;
            }
            
            catch (Exception error)
            {
                document.Close();
                application.Quit();
                document = null;
                application = null;
                MessageBox.Show(error.Message, "WordOffice", MessageBoxButtons.OK, MessageBoxIcon.Information);

                return;
            }
        }
    }
}
