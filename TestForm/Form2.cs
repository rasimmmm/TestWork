using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TestForm
{
    public partial class Form2 : Form
    {
        public Form2(Dictionary<string, int[]> d, TimeSpan ts)
        {
            InitializeComponent();
            int i = 1;
            foreach (KeyValuePair<string, int[]> tmp in d)
                dataGridView1.Rows.Add(i++, tmp.Key, tmp.Value[0],
                                                   tmp.Value[1],
                                                   tmp.Value[2]);
            string elapsedTime = String.Format("{0:00}:{1:00}:{2:00}.{3:00}",
                                                ts.Hours, ts.Minutes, ts.Seconds,
                                                ts.Milliseconds / 10);
            label1.Text = "Время выполнения: " + elapsedTime;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            create_word(dataGridView1);
        }

        static void create_word(DataGridView dataGridView)
        {
            var Wordapp = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document doc = Wordapp.Documents.Add();
            Microsoft.Office.Interop.Word.Range range = doc.Range();
            try
            {
                int countAll=0, countAsks=0, countRKK=0;
                for (int i = 1; i < dataGridView.RowCount; i++)
                {
                    countRKK = countRKK+ Convert.ToInt32(dataGridView.Rows[i - 1].Cells[2].Value);
                    countAsks = countAsks + Convert.ToInt32(dataGridView.Rows[i - 1].Cells[3].Value);
                    countAll = countAll + Convert.ToInt32(dataGridView.Rows[i - 1].Cells[4].Value);
                }
                range.Text = "Справка о неисполненных документах и обращениях граждан \n";
                range.ParagraphFormat.Alignment= Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                range.Bold = 1;
                range.Font.Name = "Arial";
                range.Font.Size = 14;
                range.EndOf();
                range.Text = "Не исполнено в срок " + countAll + " документов, из них: \n";
                range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft;
                range.Bold = 0;
                range.Font.Name = "Arial";
                range.Font.Size = 10;
                range.EndOf();
                range.Text = "- количество неисполненных входящих документов: "+ countRKK + "; \n";
                range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft;
                range.Bold = 0;
                range.Font.Name = "Arial";
                range.Font.Size = 10;
                range.EndOf();
                range.Text = "- количество неисполненных письменных обращений граждан: "+ countAsks + ". \n";
                range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft;
                range.Bold = 0;
                range.Font.Name = "Arial";
                range.Font.Size = 10;
                range.EndOf();
                range.Text = "Сортировка: по общему количеству документов \n";
                range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft;
                range.Bold = 0;
                range.Font.Name = "Arial";
                range.Font.Size = 10;
                range.EndOf();

                Microsoft.Office.Interop.Word.Table table = doc.Tables.Add(range, dataGridView.RowCount, dataGridView.ColumnCount);
                table.Borders.Enable = 1;
                table.Cell(1, 1).Range.Text = "№ п.п.";
                table.Cell(1, 2).Range.Text = "Ответственный исполнитель"; 
                table.Cell(1, 3).Range.Text = "Количество неисполненных входящих документов";
                table.Cell(1, 4).Range.Text = "Количество неисполненных письменных обращений граждан";
                table.Cell(1, 5).Range.Text = "Общее количество документов и обращений";
                table.Range.Bold = 0;
                table.Range.Font.Name = "Arial";
                table.Range.Font.Size = 10;
                table.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                //table.Columns[0].
                for (int i = 1; i < dataGridView.RowCount; i++)
                {
                    table.Cell(i + 1, 1).Range.Text = dataGridView.Rows[i - 1].Cells[0].Value.ToString();
                    table.Cell(i + 1, 2).Range.Text = dataGridView.Rows[i - 1].Cells[1].Value.ToString();
                    table.Cell(i + 1, 3).Range.Text = dataGridView.Rows[i - 1].Cells[2].Value.ToString();
                    table.Cell(i + 1, 4).Range.Text = dataGridView.Rows[i - 1].Cells[3].Value.ToString();
                    table.Cell(i + 1, 5).Range.Text = dataGridView.Rows[i - 1].Cells[4].Value.ToString();
                }
                var Paragraph = Wordapp.ActiveDocument.Paragraphs.Add();
                var rangeEnd = Paragraph.Range;
                DateTime dt = DateTime.Now;
                string curDate = dt.ToShortDateString();
                rangeEnd.Text = "\n Дата составление справки: " + curDate;
                rangeEnd.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft;
                rangeEnd.Bold = 0;
                rangeEnd.Font.Name = "Arial";
                rangeEnd.Font.Size = 10;
                Wordapp.Visible = true;
            }
            catch { }
        }
    }
}
