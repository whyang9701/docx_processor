using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace docx_processor
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string filepath = @"..\..\template\template1.docx";
            string filepath2 = @"..\..\export\export1.docx";
            using (WordprocessingDocument doc = WordprocessingDocument.CreateFromTemplate(filepath))
            {
                Dictionary<string, string> dic = new Dictionary<string, string>();
                dic.Add("aaa", "apple");
                dic.Add("bbb", "Banana");
                dic.Add("ccc", "California");
                dic.Add("ddd", "dog");

                DocProcessor.ReplaceTags(doc.MainDocumentPart.Document.Body, dic);
                doc.SaveAs(filepath2).Close();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string filepath = @"..\..\template\template1.docx";
            string filepath2 = @"..\..\export\export2.docx";
            using (WordprocessingDocument doc = WordprocessingDocument.CreateFromTemplate(filepath))
            {
                Body body = doc.MainDocumentPart.Document.Body;
                Table table = body.Elements<Table>().ElementAt(0);
                DocProcessor.TableCellMerge(table, 0, 0, 3, 0);
                DocProcessor.TableCellMerge(table, 0, 1, 0, 3);
                DocProcessor.TableCellMerge(table, 4, 4, 6, 6);

                doc.SaveAs(filepath2).Close();

            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string filepath = @"..\..\template\template2.docx";
            string filepath2 = @"..\..\export\export3.docx";
            using (WordprocessingDocument doc = WordprocessingDocument.CreateFromTemplate(filepath))
            {
                Body body = doc.MainDocumentPart.Document.Body;
                Table table = body.Elements<Table>().ElementAt(1);
                DocProcessor.DeleteTableColumn(table, 3);
                table = body.Elements<Table>().ElementAt(2);
                DocProcessor.DeleteTableRow(table, 3);


                doc.SaveAs(filepath2).Close();

            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            string filepath = @"..\..\template\template2.docx";
            string filepath2 = @"..\..\export\export4.docx";
            using (WordprocessingDocument doc = WordprocessingDocument.CreateFromTemplate(filepath))
            {
                Body body = doc.MainDocumentPart.Document.Body;
                Table table = body.Elements<Table>().ElementAt(1);
                DataTable dt = new DataTable();
                dt.Columns.AddRange(new DataColumn[] { new DataColumn("column1"), new DataColumn("column2"), new DataColumn("column3"), new DataColumn("column4") });
                DataRow dr= dt.NewRow();
                dr["column1"] = "google";
                dr["column2"] = "facebook";
                dr["column3"] = "yahoo";
                dr["column4"] = "apple";
                dt.Rows.Add(dr);
                DocProcessor.TableRowInsert(table, dt);
                table = body.Elements<Table>().ElementAt(2);
                dt = new DataTable();
                dt.Columns.AddRange(new DataColumn[] { new DataColumn("column1"), new DataColumn("column2"), new DataColumn("column3"), new DataColumn("column4") });
                 dr = dt.NewRow();
                dr["column1"] = "google";
                dr["column2"] = "facebook";
                dr["column3"] = "yahoo";
                dr["column4"] = "apple";
                dt.Rows.Add(dr);
                DocProcessor.TableColumnInsert(table, dt);
                doc.SaveAs(filepath2).Close();

            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            string filepath = @"..\..\template\template2.docx";
            string filepath2 = @"..\..\export\export5.docx";
            using (WordprocessingDocument doc = WordprocessingDocument.CreateFromTemplate(filepath))
            {
                Body body = doc.MainDocumentPart.Document.Body;
                Table table = body.Elements<Table>().ElementAt(1);
                DocProcessor.DuplicateElement(table);
                table = body.Elements<Table>().ElementAt(3);
                DocProcessor.DuplicateElement(table, new Paragraph(new Run(new Break() { Type = BreakValues.Page })));
                doc.SaveAs(filepath2).Close();


            }
        }
    }
}
