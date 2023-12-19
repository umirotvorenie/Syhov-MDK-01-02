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

        private double �����(string equation, double lowerBound, double upperBound)
        {
            return 42.0;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // �������� ��������� �� ���������� ����
            string ��������� = textBox1.Text;
            double �������������, ��������������;

            if (!double.TryParse(textBox2.Text, out �������������) ||
                !double.TryParse(textBox3.Text, out ��������������))
            {
                MessageBox.Show("����������, ������� ���������� ������� ����������.");
                return;
            }

            // ������ ������ � ������� ��������� � ������ ����������
            double ����� = �����(���������, �������������, ��������������);

            // ������ �������
            string ������������� = $"���������: {���������}, ������� ����������: {�������������} - {��������������}";
            string �������������� = "����� ����� ���� ���� ���������� ���������.";
            label4.Text = $"���������: {�����}";
        }

        private void button2_Click(object sender, EventArgs e)
        {
            // �������� ��������� �� ���������� ����
            string ��������� = textBox1.Text;
            double �������������, ��������������;

            if (!double.TryParse(textBox2.Text, out �������������) ||
                !double.TryParse(textBox3.Text, out ��������������))
            {
                MessageBox.Show("����������, ������� ���������� ������� ����������.");
                return;
            }
            // ������ ������ � ������� ��������� � ������ ����������
            double ����� = �����(���������, �������������, ��������������);

            // ������ �������
            string ������������� = $"���������: {���������}, ������� ����������: {�������������} - {��������������}";
            string �������������� = "����� ����� ���� ���� ���������� ���������.";
            Word.Application wordApp = new Word.Application();
            wordApp.Visible = true;
            Word.Document doc = wordApp.Documents.Add();
            Word.Range range = doc.Range();
            range.Text = "������ ������ ���� ��������" + Environment.NewLine +
                         "�����: " + ����� + Environment.NewLine +
                         "������ �������: " + ������������� + Environment.NewLine +
                         "���� ����������: " + ��������������;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string ��������� = textBox1.Text;
            double �������������, ��������������;

            if (!double.TryParse(textBox2.Text, out �������������) ||
                !double.TryParse(textBox3.Text, out ��������������))
            {
                MessageBox.Show("����������, ������� ���������� ������� ����������.");
                return;
            }
            // ������ ������ � ������� ��������� � ������ ����������
            double ����� = �����(���������, �������������, ��������������);

            // ������ �������
            string ������������� = $"���������: {���������}, ������� ����������: {�������������} - {��������������}";
            string �������������� = "����� ����� ���� ���� ���������� ���������.";
            Excel.Application excelApp = new Excel.Application();
            excelApp.Visible = true;
            Excel.Workbook workbook = excelApp.Workbooks.Add();
            Excel.Worksheet sheet = workbook.ActiveSheet;
            sheet.Cells[1, 1] = "������ ������ ���� ��������";
            sheet.Cells[2, 1] = "�����";
            sheet.Cells[2, 2] = �����;
            sheet.Cells[4, 1] = "������ �������";
            sheet.Cells[5, 1] = �������������;
            sheet.Cells[6, 1] = "���� ����������";
            sheet.Cells[7, 1] = ��������������;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            string ��������� = textBox1.Text;
            double �������������, ��������������;

            if (!double.TryParse(textBox2.Text, out �������������) ||
                !double.TryParse(textBox3.Text, out ��������������))
            {
                MessageBox.Show("����������, ������� ���������� ������� ����������.");
                return;
            }
            // ������ ������ � ������� ��������� � ������ ����������
            double ����� = �����(���������, �������������, ��������������);

            // ������ �������
            string ������������� = $"���������: {���������}, ������� ����������: {�������������} - {��������������}";
            string �������������� = "����� ����� ���� ���� ���������� ���������.";
            string ����������Pdf = "����������_�������_������.pdf";
            using (FileStream fs = new FileStream(����������Pdf, FileMode.Create, FileAccess.Write, FileShare.None))
            {
                iTextSharp.text.Document pdfDoc = new iTextSharp.text.Document();
                PdfWriter.GetInstance(pdfDoc, fs);
                pdfDoc.Open();

                // ����� ������� �������
                pdfDoc.Add(new Paragraph("Calculation details: " + �������������));
                pdfDoc.Add(iTextSharp.text.Chunk.NEWLINE);

                // ����� ��������� ����� ��������������
                pdfDoc.Add(new Paragraph("Detailed integration steps:"));

                // ����� ��������������� ������
                pdfDoc.Add(new Paragraph("The volume of the body of rotation: " + �����));

                pdfDoc.Close();
            }
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo()
            {
                FileName = ����������Pdf,
                UseShellExecute = true
            });
            MessageBox.Show("���������� ������� ��������� � Excel, Word � PDF.");
        }
    }
}