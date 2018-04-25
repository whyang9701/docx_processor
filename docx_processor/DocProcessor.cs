using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using System.Text.RegularExpressions;
using System.Data.SqlClient;
using System.Data;

namespace docx_processor
{
    class DocProcessor
    {
        public static void DuplicateElement(OpenXmlElement element)
        {
            OpenXmlElement element2 = element.CloneNode(true);
            element.InsertAfterSelf(element2);
        }

        public static void DuplicateElement(OpenXmlElement element, OpenXmlElement saparator)
        {
            OpenXmlElement element2 = element.CloneNode(true);
            element.InsertAfterSelf(saparator);
            saparator.InsertAfterSelf(element2);
        }

        //向右插入資料並保持原格式
        public static void TableColumnInsert(Table tb, DataTable data)
        {
            int widthCount = 0;
            TableGrid tg = tb.Elements<TableGrid>().First();
            for (int j = 0; j < data.Rows.Count; j++)
            {
                DataRow dr = data.Rows[j];
                GridColumn gc = (GridColumn)tg.Elements<GridColumn>().Last().CloneNode(true);
                widthCount += Convert.ToInt32(gc.Width.Value);
                tg.AppendChild<GridColumn>(gc);
                for (int i = 0; i < data.Columns.Count; i++)
                {
                    var row = tb.Elements<TableRow>().ElementAt(i);
                    var newCell = (TableCell)(row.Elements<TableCell>().LastOrDefault<TableCell>() ?? new TableCell()).CloneNode(true);

                    var newParagraph = newCell.Elements<Paragraph>().FirstOrDefault() ?? new Paragraph();

                    Run newRun = (Run)(newParagraph.Elements<Run>().FirstOrDefault<Run>() ?? new Run()).CloneNode(true);

                    Text text = newRun.Elements<Text>().FirstOrDefault() ?? new Text();

                    if (newParagraph.Parent == null)
                    {
                        newCell.AppendChild<Paragraph>(newParagraph);
                    }

                    if (text.Parent == null)
                    {
                        newRun.AppendChild<Text>(text);
                    }

                    text.Text = dr[i].ToString();
                    newParagraph.RemoveAllChildren<Run>();
                    newParagraph.AppendChild<Run>(newRun);
                    row.AppendChild<TableCell>(newCell);
                }

            }

            //設定新的寬度
            var tbW = tb.Elements<TableProperties>().First<TableProperties>().Elements<TableWidth>().First<TableWidth>();
            tbW.Width.Value = (Convert.ToInt32(tbW.Width.Value) + widthCount).ToString();

        }

        //刪除指定欄
        public static void DeleteTableColumn(Table tb, int index)
        {
            //delect gridColumn
            GetElement<GridColumn>(GetElement<TableGrid>(tb, 0), index).Remove();
            foreach (TableRow tr in tb.Elements<TableRow>())
            {
                GetElement<TableCell>(tr, index).Remove();
            }
        }

        //取得指定子元素
        public static T GetElement<T>(OpenXmlElement element, int index) where T : OpenXmlElement
        {
            return element.Elements<T>().ElementAtOrDefault(index);
        }

        //向下插入資料並保持原格式
        public static void TableRowInsert(Table tb, DataTable data)
        {
            foreach (DataRow dr in data.Rows)
            {
                TableRow tr = (TableRow)tb.Elements<TableRow>().Last().CloneNode(true);

                for (int i = 0; i < data.Columns.Count; i++)
                {

                    TableCell tc = tr.Elements<TableCell>().ElementAt(i);

                    Paragraph p = tc.Elements<Paragraph>().FirstOrDefault() ?? new Paragraph();
                    if (p.Parent == null)
                    {
                        tc.AppendChild<Paragraph>(p);
                    }
                    Run r = (Run)(p.Elements<Run>().FirstOrDefault() ?? new Run()).CloneNode(true);
                    Text text = r.Elements<Text>().FirstOrDefault() ?? new Text();
                    if (text.Parent == null)
                    {
                        r.AppendChild<Text>(text);
                    }
                    text.Text = dr[i].ToString();
                    p.RemoveAllChildren<Run>();
                    p.AppendChild<Run>(r);
                }
                tb.AppendChild<TableRow>(tr);
            }

        }

        //刪除指定列
        public static void DeleteTableRow(Table tb, int index)
        {
            tb.Elements<TableRow>().ElementAt(index).Remove();
        }

        //合併儲存格
        public static void TableCellMerge(Table tb, int x1, int y1, int x2, int y2)
        {
            if (x1 == x2 && y1 == y2)
            {
                return;

            }
            int minX = Math.Min(x1, x2);
            int maxX = Math.Max(x1, x2);
            int minY = Math.Min(y1, y2);
            int maxY = Math.Max(y1, y2);
            int mergeColumnCount = maxX - minX + 1;
            TableRow tr;
            TableCell tc;
            TableCellProperties tcpr;
            GridSpan gs;
            VerticalMerge vm;
            //horizontal merge
            for (int j = minY; j <= maxY; j++)
            {
                tr = tb.Elements<TableRow>().ElementAt<TableRow>(j);
                tc = tr.Elements<TableCell>().ElementAt<TableCell>(minX);
                tcpr = tc.Elements<TableCellProperties>().FirstOrDefault();
                gs = new GridSpan() { Val = mergeColumnCount };
                if (tcpr != null)
                {
                    tcpr.AppendChild<GridSpan>(gs);
                }
                else
                {
                    tc.AppendChild<TableCellProperties>(tcpr);
                    tcpr.AppendChild<GridSpan>(gs);
                }

                for (int i = minX + 1; i <= maxX; i++)
                {
                    tr.Elements<TableCell>().ElementAt<TableCell>(minX + 1).Remove();
                }

            }
            //vertical merge
            if (maxY != minY)
            {
                tr = tb.Elements<TableRow>().ElementAt<TableRow>(minY);
                tc = tr.Elements<TableCell>().ElementAt<TableCell>(minX);
                tcpr = tc.Elements<TableCellProperties>().FirstOrDefault();
                vm = new VerticalMerge() { Val = MergedCellValues.Restart };
                if (tcpr != null)
                {
                    tcpr.AppendChild<VerticalMerge>(vm);
                }
                else
                {
                    tcpr.AppendChild<VerticalMerge>(vm);
                    tc.AppendChild<TableCellProperties>(tcpr);
                }

                for (int j = minY + 1; j <= maxY; j++)
                {
                    tr = tb.Elements<TableRow>().ElementAt<TableRow>(j);
                    tc = tr.Elements<TableCell>().ElementAt<TableCell>(minX);
                    tcpr = tc.Elements<TableCellProperties>().FirstOrDefault();

                    vm = new VerticalMerge() { Val = MergedCellValues.Continue };
                    if (tcpr != null)
                    {
                        tcpr.AppendChild<VerticalMerge>(vm);

                    }
                    else
                    {
                        tcpr.AppendChild<VerticalMerge>(vm);
                        tc.AppendChild<TableCellProperties>(tcpr);
                    }
                }
            }
        }
        public static void ReplaceTags(Body body, Dictionary<string, string> dic)
        {
            var tables = body.Elements<Table>();
            var paragraphs = body.Elements<Paragraph>();

            foreach (var pair in dic)
            {
                foreach (Paragraph p in paragraphs)
                {
                    Regex regex = new Regex(string.Format("{{{0}}}", pair.Key));
                    if (regex.IsMatch(p.InnerText))
                    {
                        string newText = regex.Replace(p.InnerText, pair.Value);

                        Run r = (Run)p.Elements<Run>().First().CloneNode(true);
                        Text text = r.Elements<Text>().First();
                        text.Text = (newText);


                        p.RemoveAllChildren<Run>();
                        p.AppendChild<Run>(r);

                    }
                }

                foreach (Table t in tables)
                {
                    foreach (TableRow tr in t.Elements<TableRow>())
                    {
                        foreach (TableCell tc in tr.Elements<TableCell>())
                        {
                            foreach (Paragraph p in tc.Elements<Paragraph>())
                            {
                                Regex regex = new Regex(string.Format("{{{0}}}", pair.Key));
                                if (regex.IsMatch(p.InnerText))
                                {
                                    string newText = regex.Replace(p.InnerText, pair.Value);

                                    Run r = (Run)p.Elements<Run>().First().CloneNode(true);
                                    Text text = r.Elements<Text>().First();
                                    text.Text = (newText);


                                    p.RemoveAllChildren<Run>();
                                    p.AppendChild<Run>(r);

                                }
                            }
                        }
                    }
                }
            }
        }
    }
}
