using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing.Charts;
using Aspose.Words.Tables;
using NetCoreWeb.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using Chart = Aspose.Words.Drawing.Charts.Chart;
using ChartType = Aspose.Words.Drawing.Charts.ChartType;

namespace NetCoreWeb.Common
{
    public class WordHelper
    {
        #region 导出Word
        public static string ModelToWord(Model_Car car, string path, int num)
        {
            try
            {
                Document doc = new Document(path);
                DocumentBuilder builder = new DocumentBuilder(doc);

                foreach (System.Reflection.PropertyInfo p in car.GetType().GetProperties())
                {
                    builder.MoveToBookmark(p.Name);
                    builder.Write(p.GetValue(car, null).ToString());
                }
                doc.Save(System.AppDomain.CurrentDomain.BaseDirectory + string.Format("OutFile/car{0}协议书By书签.doc", num));
                return "OK";
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        #endregion

        /// <summary>
        /// 将DataTable转换为Word
        /// </summary>
        /// <param name="inputdt"></param>
        /// <param name="docfullpath"></param>
        public static void DtToWord(DataTable inputdt, string docfullpath)
        {

            for (int i = 0; i < inputdt.Rows.Count; i++)
            {
                Document doc = new Document(docfullpath);
                DocumentBuilder builder = new DocumentBuilder(doc);
                for (int j = 0; j < inputdt.Columns.Count; j++)
                {
                    string Cellvalue = inputdt.Rows[i][j].ToString().Trim();
                    builder.MoveToBookmark(inputdt.Columns[j].ColumnName);
                    builder.Write(Cellvalue);
                }
                doc.Save(System.AppDomain.CurrentDomain.BaseDirectory + string.Format("OutFile/car{0}协议书By书签.doc", i));

            }

        }

        /// <summary>
        /// 根据datatable获得列名
        /// </summary>
        /// <param name="dt">表对象</param>
        /// <returns>返回结果的数据列数组</returns>
        public static string[] GetColumnsByDataTable(DataTable dt)
        {
            string[] strColumns = null;


            if (dt.Columns.Count > 0)
            {
                int columnNum = 0;
                columnNum = dt.Columns.Count;
                strColumns = new string[columnNum];
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    strColumns[i] = dt.Columns[i].ColumnName;
                }
            }
            return strColumns;
        }

        public static void GenerateWordDynamically() {

            // 创建一个新的 Word 文档
            Document doc = new Document();

            // 添加标题
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            builder.Font.Size = 18;
            builder.Font.Bold = true;
            builder.Writeln("销售报表");

            // 添加日期
            builder.ParagraphFormat.ClearFormatting();
            builder.Font.Size = 12;
            builder.Font.Bold = false;
            builder.Writeln("报告日期：" + DateTime.Now.ToShortDateString());

            #region 添加表格
            // 开始画Table
            builder.InsertBreak(BreakType.SectionBreakNewPage);
            Table table = builder.StartTable();

            // 添加表头行
            Row headerRow = new Row(doc); // 创建新的行
            table.AppendChild(headerRow); // 将新行添加到表格中
            // 添加表头单元格
            headerRow.AppendChild(new Cell(doc));
            headerRow.AppendChild(new Cell(doc));
            headerRow.AppendChild(new Cell(doc));

            AddCellWithText(headerRow.Cells[0], "产品名称");
            AddCellWithText(headerRow.Cells[1], "销售数量");
            AddCellWithText(headerRow.Cells[2], "销售金额");


            // 添加数据行
            string[] products = { "产品A", "产品B", "产品C" };
            int[] quantities = { 100, 150, 80 };
            decimal[] amounts = { 5000.0m, 7500.0m, 4000.0m };

            for (int i = 0; i < products.Length; i++)
            {
                Row dataRow = new Row(doc);
                table.AppendChild(dataRow); // 将数据行添加到表格中
                //添加数据行单元格
                dataRow.AppendChild(new Cell(doc));
                dataRow.AppendChild(new Cell(doc));
                dataRow.AppendChild(new Cell(doc));
                AddCellWithText(dataRow.Cells[0], products[i]);
                AddCellWithText(dataRow.Cells[1], quantities[i].ToString());
                AddCellWithText(dataRow.Cells[2], amounts[i].ToString("C")); // 格式化为货币
            }

            //在添加表格数据之后设置表格格式
            //Aspose.Words 要求在为表格设置格式之前，至少要添加一行数据到表格中。
            table.PreferredWidth = PreferredWidth.FromPercent(100);
            table.AllowAutoFit = false;
            // 设置单元格内文字的对齐方式
            foreach (Cell cell in table.FirstRow.Cells)
            {
                cell.CellFormat.HorizontalMerge = CellMerge.None; // 防止单元格合并
                cell.Paragraphs[0].ParagraphFormat.Alignment = ParagraphAlignment.Center; // 文字水平居中对齐
            }
            builder.EndTable();
            #endregion

            #region 添加折线图
            // 添加折线图
            builder.InsertBreak(BreakType.SectionBreakNewPage);
            Aspose.Words.Drawing.Shape lineChartShape = builder.InsertChart(ChartType.Line, 400, 300);
            Chart lineChart = lineChartShape.Chart;
            ChartSeriesCollection seriesCollection = lineChart.Series;
            seriesCollection.Clear();//清除默认 
            lineChart.Title.Text = "产品销售趋势";
            lineChart.AxisX.Title.Text = "产品类别";
            lineChart.AxisY.Title.Text = "销售数量";

            ChartSeries lineSeries = lineChart.Series.Add("销售数量", new string[] { products[0], products[1], products[2] }, new double[] { quantities[0], quantities[1], quantities[2] });
            lineSeries.Marker.Symbol = MarkerSymbol.Circle;
            lineSeries.Marker.Size = 10;
            #endregion
            #region 添加柱状图
            // 添加柱状图
            builder.InsertBreak(BreakType.SectionBreakNewPage);
            Aspose.Words.Drawing.Shape barChartShape = builder.InsertChart(ChartType.Bar, 400, 300);
            Chart barChart = barChartShape.Chart;
            barChart.Series.Clear();
            barChart.Title.Text = "产品销售数量";
            barChart.AxisX.Title.Text = "产品";
            barChart.AxisY.Title.Text = "销售数量";

            // 创建柱状图数据系列
            ChartSeries barSeries = barChart.Series.Add("销售数量", new string[] { "产品A", "产品B", "产品C" }, new double[] { quantities[0], quantities[1], quantities[2] });
            barSeries.DataLabels.ShowValue = true; // 显示数据标签 
            #endregion
            #region 添加饼图
            builder.InsertBreak(BreakType.SectionBreakNewPage);
            Aspose.Words.Drawing.Shape pieChartShape = builder.InsertChart(ChartType.Pie, 400, 300);
            Chart pieChart = pieChartShape.Chart;
            pieChart.Series.Clear();
            pieChart.Title.Text = "销售额分布";

            // 创建饼图数据系列
            ChartSeries pieSeries = pieChart.Series.Add("销售金额", new string[] { "产品A", "产品B", "产品C" }, new double[] { Convert.ToDouble(amounts[0]), Convert.ToDouble(amounts[1]), Convert.ToDouble(amounts[2]) });
            //pieSeries.Bubble3D = true;

            #endregion
            #region 添加散点图
            // 添加散点图
            builder.InsertBreak(BreakType.SectionBreakNewPage);
            Aspose.Words.Drawing.Shape scatterChartShape = builder.InsertChart(ChartType.Scatter, 400, 300);
            Chart scatterChart = scatterChartShape.Chart;
            scatterChart.Series.Clear();//清除默认
            scatterChart.Title.Text = "销售数量与销售金额关系";
            scatterChart.AxisX.Title.Text = "销售数量";
            scatterChart.AxisY.Title.Text = "销售金额";

            //ChartAxis categoryAxis = scatterChart.AxisX;
            //categoryAxis.MajorTickMark = AxisTickMark.Outside; 
            //categoryAxis.MinorTickMark = AxisTickMark.None;  

            // 创建散点图数据系列
            ChartSeries scatterSeries = scatterChart.Series.Add("销售数据", new double[] { quantities[0], quantities[1], quantities[2] }, new double[] { Convert.ToDouble(amounts[0]), Convert.ToDouble(amounts[1]), Convert.ToDouble(amounts[2]) });

            scatterSeries.Marker.Symbol = MarkerSymbol.Circle; // 设置标记类型为圆圈
            scatterSeries.Marker.Size = 10; // 设置标记大小

            #endregion
            // 保存文档
            doc.Save(System.AppDomain.CurrentDomain.BaseDirectory + string.Format("OutFile/SalesReportWithCharts{0}.docx", DateTime.Now.ToString("yyyyMMddHHmmss")));
            Console.WriteLine("报表生成成功！");
        }
        private static void AddCellWithText(Cell cell, string text)
        {
            Paragraph cellParagraph = new Paragraph(cell.Document);
            Run cellRun = new Run(cell.Document);
            cellRun.Text = text;
            cellParagraph.Runs.Add(cellRun);
            cell.AppendChild(cellParagraph);
        }
    }
}
