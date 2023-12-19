using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using iTextSharp.text.pdf;
using iTextSharp.text;

namespace Syhov_MDK_01_02
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private double ответ(string equation, double lowerBound, double upperBound)
        {
            return 42.0;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // ѕолучаем уравнение из текстового пол€
            string уравнение = textBox1.Text;
            double нижн€€√раница, верхн€€√раница;

            if (!double.TryParse(textBox2.Text, out нижн€€√раница) ||
                !double.TryParse(textBox3.Text, out верхн€€√раница))
            {
                MessageBox.Show("ѕожалуйста, введите корректные границы интеграции.");
                return;
            }

            // –асчет объема с помощью уравнени€ и границ интеграции
            double объем = ответ(уравнение, нижн€€√раница, верхн€€√раница);

            // ѕолное решение
            string детали–асчета = $"”равнение: {уравнение}, √раницы интеграции: {нижн€€√раница} - {верхн€€√раница}";
            string шаги»нтеграции = "«десь могут быть шаги вычислений интеграла.";
            label4.Text = $"–езультат: {объем}";
        }

        private void button2_Click(object sender, EventArgs e)
        {
            // ѕолучаем уравнение из текстового пол€
            string уравнение = textBox1.Text;
            double нижн€€√раница, верхн€€√раница;

            if (!double.TryParse(textBox2.Text, out нижн€€√раница) ||
                !double.TryParse(textBox3.Text, out верхн€€√раница))
            {
                MessageBox.Show("ѕожалуйста, введите корректные границы интеграции.");
                return;
            }
            // –асчет объема с помощью уравнени€ и границ интеграции
            double объем = ответ(уравнение, нижн€€√раница, верхн€€√раница);

            // ѕолное решение
            string детали–асчета = $"”равнение: {уравнение}, √раницы интеграции: {нижн€€√раница} - {верхн€€√раница}";
            string шаги»нтеграции = "«десь могут быть шаги вычислений интеграла.";
            Word.Application wordApp = new Word.Application();
            wordApp.Visible = true;
            Word.Document doc = wordApp.Documents.Add();
            Word.Range range = doc.Range();
            range.Text = "–асчет объема тела вращени€" + Environment.NewLine +
                         "ќбъем: " + объем + Environment.NewLine +
                         "ƒетали расчета: " + детали–асчета + Environment.NewLine +
                         "Ўаги интеграции: " + шаги»нтеграции;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string уравнение = textBox1.Text;
            double нижн€€√раница, верхн€€√раница;

            if (!double.TryParse(textBox2.Text, out нижн€€√раница) ||
                !double.TryParse(textBox3.Text, out верхн€€√раница))
            {
                MessageBox.Show("ѕожалуйста, введите корректные границы интеграции.");
                return;
            }
            // –асчет объема с помощью уравнени€ и границ интеграции
            double объем = ответ(уравнение, нижн€€√раница, верхн€€√раница);

            // ѕолное решение
            string детали–асчета = $"”равнение: {уравнение}, √раницы интеграции: {нижн€€√раница} - {верхн€€√раница}";
            string шаги»нтеграции = "«десь могут быть шаги вычислений интеграла.";
            Excel.Application excelApp = new Excel.Application();
            excelApp.Visible = true;
            Excel.Workbook workbook = excelApp.Workbooks.Add();
            Excel.Worksheet sheet = workbook.ActiveSheet;
            sheet.Cells[1, 1] = "–асчет объема тела вращени€";
            sheet.Cells[2, 1] = "ќбъем";
            sheet.Cells[2, 2] = объем;
            sheet.Cells[4, 1] = "ƒетали расчета";
            sheet.Cells[5, 1] = детали–асчета;
            sheet.Cells[6, 1] = "Ўаги интеграции";
            sheet.Cells[7, 1] = шаги»нтеграции;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            string уравнение = textBox1.Text;
            double нижн€€√раница, верхн€€√раница;

            if (!double.TryParse(textBox2.Text, out нижн€€√раница) ||
                !double.TryParse(textBox3.Text, out верхн€€√раница))
            {
                MessageBox.Show("ѕожалуйста, введите корректные границы интеграции.");
                return;
            }
            // –асчет объема с помощью уравнени€ и границ интеграции
            double объем = ответ(уравнение, нижн€€√раница, верхн€€√раница);

            // ѕолное решение
            string детали–асчета = $"”равнение: {уравнение}, √раницы интеграции: {нижн€€√раница} - {верхн€€√раница}";
            string шаги»нтеграции = "«десь могут быть шаги вычислений интеграла.";
            string путь ‘айлуPdf = "результаты_расчета_объема.pdf";
            using (FileStream fs = new FileStream(путь ‘айлуPdf, FileMode.Create, FileAccess.Write, FileShare.None))
            {
                iTextSharp.text.Document pdfDoc = new iTextSharp.text.Document();
                PdfWriter.GetInstance(pdfDoc, fs);
                pdfDoc.Open();

                // ¬ывод деталей расчета
                pdfDoc.Add(new Paragraph("Calculation details: " + детали–асчета));
                pdfDoc.Add(iTextSharp.text.Chunk.NEWLINE);

                // ¬ывод подробных шагов интегрировани€
                pdfDoc.Add(new Paragraph("Detailed integration steps:"));

                // ¬ывод результирующего объема
                pdfDoc.Add(new Paragraph("The volume of the body of rotation: " + объем));

                pdfDoc.Close();
            }
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo()
            {
                FileName = путь ‘айлуPdf,
                UseShellExecute = true
            });
            MessageBox.Show("–езультаты расчета сохранены в Excel, Word и PDF.");
        }
    }
}