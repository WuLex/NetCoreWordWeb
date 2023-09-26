using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;

namespace NetCoreWeb.Common
{
    public class ExcelHelper
    {
        #region 数据导入

        /// <summary>
        /// 将Excel文件转换为DataTable
        /// </summary>
        /// <param name="Path"></param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        public static DataTable ExeclToDataTable(string Path)
        {
            try
            {
                DataTable dt = new DataTable();

                // 加载Excel工作簿
                Aspose.Cells.Workbook workbook = new Workbook(Path);
                //workbook.Open(Path); //已过时
                //Worksheets wsts = workbook.Worksheets;  //已过时
                //Worksheet sheet = workbook.Worksheets["New Worksheet1"];
                WorksheetCollection wsts = workbook.Worksheets;

                // 遍历工作表
                for (int i = 0; i < wsts.Count; i++)
                {
                    Worksheet wst = wsts[i];
                    int MaxR = wst.Cells.MaxRow;
                    int MaxC = wst.Cells.MaxColumn;
                    if (MaxR > 0 && MaxC > 0)
                    {
                        dt = wst.Cells.ExportDataTableAsString(0, 0, MaxR + 1, MaxC + 1, true);
                    }
                }
                return dt;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        public static DataSet ExeclToDataSet(string Path)
        {
            try
            {
                DataTable dt = new DataTable();
                //Aspose.Cells.Workbook workbook = new Workbook();
                //workbook.Open(Path);

                Aspose.Cells.Workbook workbook = new Workbook(Path);
                WorksheetCollection wsts = workbook.Worksheets;

                for (int i = 0; i < wsts.Count; i++)
                {
                    Worksheet wst = wsts[i];
                    int MaxR = wst.Cells.MaxRow;
                    int MaxC = wst.Cells.MaxColumn;
                    if (MaxR > 0 && MaxC > 0)
                    {
                        dt = wst.Cells.ExportDataTableAsString(0, 0, MaxR + 1, MaxC + 1, true);
                    }
                }

                //SqlDataAdapter adapter = null;
                DataSet ds = new DataSet();
                ds.Tables.Add(dt);
                //adapter.Fill(dt);
                return ds;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        #endregion
    }
}
