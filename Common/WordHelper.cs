using Aspose.Words;
using NetCoreWeb.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;

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

    }
}
