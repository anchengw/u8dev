using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Linq;
using System.Text;

using Excel = Microsoft.Office.Interop.Excel;

namespace u8dev
{
    public class excelReader
    {
        public static string connectionString = "";//数据库连接字符串
        #region OleDb读取Excel
        /// <summary>
        /// 将 Excel 文件转成 DataTable 后,再把 DataTable中的数据写入表Products
        /// </summary>
        /// <param name="serverMapPathExcelAndFileName"></param>
        /// <param name="excelFileRid"></param>
        /// <returns></returns>
        public static int WriteExcelToDataBase(string excelFileName)
        {
            int rowsCount = 0;
            OleDbConnection objConn = new System.Data.OleDb.OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + excelFileName + ";" + "Extended Properties=Excel 8.0;");
            objConn.Open();
            try
            {
                DataTable schemaTable = objConn.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables, null);
                string sheetName = string.Empty;
                for (int j = 0; j < schemaTable.Rows.Count; j++)
                {
                    sheetName = schemaTable.Rows[j][2].ToString().Trim();//获取 Excel 的表名，默认值是sheet1 
                    DataTable excelDataTable = ExcelToDataTable(excelFileName, sheetName, true);
                    if (excelDataTable.Columns.Count > 1)
                    {
                        SqlBulkCopy sqlbulkcopy = new SqlBulkCopy(connectionString, SqlBulkCopyOptions.UseInternalTransaction);
                        sqlbulkcopy.DestinationTableName = "Products";//数据库中的表名


                        sqlbulkcopy.WriteToServer(excelDataTable);
                        sqlbulkcopy.Close();
                    }
                }
            }
            catch (SqlException ex)
            {
                throw ex;
            }
            finally
            {
                objConn.Close();
                objConn.Dispose();
            }
            return rowsCount;
        }
        /// <summary>
        /// 读取Excel
        /// </summary>
        /// <param name="Path"></param>
        /// <param name="tableName"></param>
        /// <returns></returns>
        public static DataSet ExcelToDS(string Path, string sheetName)
        {
            string strConn = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + Path + ";" + "Extended Properties=Excel 8.0;";  //HDR=Yes;
            DataSet ds = null;
            using (OleDbConnection conn = new OleDbConnection(strConn))
            {
                OleDbDataAdapter myCommand = null;
                try
                {
                    conn.Open();
                    string strExcel = "";
                    strExcel = "select * from [" + sheetName + "$]";
                    myCommand = new OleDbDataAdapter(strExcel, strConn);
                    ds = new DataSet();
                    myCommand.Fill(ds, "table1");
                }
                catch (SqlException ex)
                {
                    throw ex;
                }
                finally
                {
                    myCommand.Dispose();
                    conn.Close();
                }
                return ds;
            }
        }
        /// <summary>
        /// 将 Excel 文件转成 DataTable
        /// </summary>
        /// <param name="serverMapPathExcel">Excel文件及其路径</param>
        /// <param name="strSheetName">工作表名,如:Sheet1</param>
        /// <param name="isTitleOrDataOfFirstRow">True 第一行是标题,False 第一行是数据</param>
        /// <returns>DataTable</returns>
        public static DataTable ExcelToDataTable(string serverMapPathExcel, string strSheetName, bool isTitleOrDataOfFirstRow)
        {


            string HDR = string.Empty;//如果第一行是数据而不是标题的话, 应该写: "HDR=No;"
            if (isTitleOrDataOfFirstRow)
            {
                HDR = "YES";//第一行是标题
            }
            else
            {
                HDR = "NO";//第一行是数据
            }
            //源的定义 
            string strConn = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + serverMapPathExcel + ";" + "Extended Properties='Excel 8.0;HDR=" + HDR + ";IMEX=1';";
            //Sql语句
            //string strExcel = string.Format("select * from [{0}$]", strSheetName); 这是一种方法
            string strExcel = "select * from   [" + strSheetName + "]";
            //定义存放的数据表
            DataSet ds = new DataSet();
            //连接数据源
            using (OleDbConnection conn = new OleDbConnection(strConn))
            {
                try
                {
                    conn.Open();
                    //适配到数据源
                    OleDbDataAdapter adapter = new OleDbDataAdapter(strExcel, strConn);

                    adapter.Fill(ds, strSheetName);
                }
                catch (System.Data.SqlClient.SqlException ex)
                {
                    throw ex;
                }
                finally
                {
                    conn.Close();
                    conn.Dispose();
                }
            }
            return ds.Tables[strSheetName];
        }
        public static DataSet ExcelToDS(string Path)
        {
            string strConn = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + Path + ";" + "Extended Properties=Excel 8.0;";
            OleDbConnection conn = new OleDbConnection(strConn);
            conn.Open();
            string strExcel = "";
            OleDbDataAdapter myCommand = null;
            DataSet ds = null;
            strExcel = "select * from [sheet1$]";
            myCommand = new OleDbDataAdapter(strExcel, strConn);
            ds = new DataSet();
            myCommand.Fill(ds, "table1");
            return ds;
        }
        public static DataTable ExcelToTable(string excelFilename, bool isTitleOrDataOfFirstRow)
        {
            string HDR = string.Empty;//如果第一行是数据而不是标题的话, 应该写: "HDR=No;"
            if (isTitleOrDataOfFirstRow)
            {
                HDR = "YES";//第一行是标题
            }
            else
            {
                HDR = "NO";//第一行是数据
            }
            string connectionString = string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Jet OLEDB:Engine Type=35;Extended Properties=Excel 8.0;HDR=\"{1}\";Persist Security Info=False", excelFilename,HDR);
            DataSet ds = new DataSet();
            string tableName;
            using (System.Data.OleDb.OleDbConnection connection = new System.Data.OleDb.OleDbConnection(connectionString))
            {
                connection.Open();
                DataTable table = connection.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables, null);
                tableName = table.Rows[0]["Table_Name"].ToString();
                string strExcel = "select * from " + "[" + tableName + "]";
                OleDbDataAdapter adapter = new OleDbDataAdapter(strExcel, connectionString);
                adapter.Fill(ds, tableName);
                connection.Close();
            }
            return ds.Tables[tableName];
        }
        #endregion
        #region Microsoft.Office.Interop.Excel 操作excel
        public static System.Data.DataTable getExcelToTable(string filename)
        {
            object missing = System.Reflection.Missing.Value;
            Excel.Application myExcel = new Excel.Application();//lauch excel application
            if (myExcel == null)
            {
                return null;//打开EXCEL应用失败 
            }
            else
            {
                myExcel.Visible = false;
                myExcel.UserControl = true;
                //以只读的形式打开EXCEL文件
                Excel.Workbook myBook = myExcel.Application.Workbooks.Open(filename, missing, true, missing, missing, missing,
                     missing, missing, missing, true, missing, missing, missing, missing, missing);

                if (myBook != null)   //打开成功
                {
                    myExcel.Visible = false;
                    //取得第一个工作薄
                    Excel.Worksheet mySheet = (Excel.Worksheet)myBook.Worksheets[1];  //得到工作表
                    //取得总记录行数   (包括标题列)
                    //int rowsint = mySheet.UsedRange.Cells.Rows.Count; //得到行数
                    //int columnsint = mySheet.UsedRange.Cells.Columns.Count;//得到列数
                    System.Data.DataTable dt = new System.Data.DataTable();
                    for (int j = 1; j <= mySheet.Cells.CurrentRegion.Columns.Count; j++)
                    {
                        string colName = ((Excel.Range)mySheet.Cells[1, j]).Text.ToString();
                        dt.Columns.Add(colName);
                    }
                    for (int i = 2; i <= mySheet.Cells.CurrentRegion.Rows.Count; i++)   //i=2 第一行作列标题 把工作表导入DataTable中
                    {
                        DataRow myRow = dt.NewRow();
                        for (int j = 1; j <= mySheet.Cells.CurrentRegion.Columns.Count; j++)
                        {
                            Excel.Range temp = (Excel.Range)mySheet.Cells[i, j];
                            string strValue = temp.Text.ToString();
                            myRow[j - 1] = strValue;
                        }
                        dt.Rows.Add(myRow);
                    }
                    myExcel.Quit();  //退出Excel文件
                    myExcel = null;
                    System.Diagnostics.Process[] procs = System.Diagnostics.Process.GetProcessesByName("excel");
                    foreach (System.Diagnostics.Process pro in procs)
                    {
                        pro.Kill();//杀掉进程
                    }
                    GC.Collect();
                    return dt;
                }
            }
            //打开不成功
            return null;
        }
        ///<summary>
        /// 获取指定文件的指定单元格内容
        ///</summary>
        /// <param name="fileName">文件路径</param>
        /// <param name="row">行号</param>
        /// <param name="column">列号</param>
        /// <returns>返回单元指定单元格内容</returns>
        public static string getExcelOneCell(string fileName, int row, int column)
        {
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook wbook = app.Workbooks.Open(fileName, Type.Missing, Type.Missing,
                 Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                 Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                 Type.Missing, Type.Missing);
            Microsoft.Office.Interop.Excel.Worksheet workSheet = (Microsoft.Office.Interop.Excel.Worksheet)wbook.Worksheets[1];
            string temp = ((Microsoft.Office.Interop.Excel.Range)workSheet.Cells[row, column]).Text.ToString();
            wbook.Close(false, fileName, false);
            app.Quit();
            NAR(app);
            NAR(wbook);
            NAR(workSheet);
            return temp;
        }
        //此函数用来释放对象的相关资源
        private static void NAR(Object o)
        {
            try
            {
                //使用此方法，来释放引用某些资源的基础 COM 对象。 这里的o就是要释放的对象
                System.Runtime.InteropServices.Marshal.ReleaseComObject(o);
            }
            catch { }
            finally
            {
                o = null; GC.Collect();
            }
        }
        private void OpenExcel(string strFileName)
        {
            object missing = System.Reflection.Missing.Value;
            Excel.Application excelApp = new Excel.ApplicationClass();//lauch excel application
            if (excelApp == null)
            {
                return;//创建EXCEL应用失败!
            }
            else
            {
                excelApp.Visible = false;
                excelApp.UserControl = true;
                //以只读的形式打开EXCEL文件
                Excel.Workbook wb = excelApp.Application.Workbooks.Open(strFileName, missing, true, missing, missing, missing,
                missing, missing, missing, true, missing, missing, missing, missing, missing);
                //取得第一个工作薄
                Excel.Worksheet ws = (Excel.Worksheet)wb.Worksheets.get_Item(1);
                //取得总记录行数   (包括标题列)
                int rowsint = ws.UsedRange.Cells.Rows.Count; //得到行数
                //int columnsint = mySheet.UsedRange.Cells.Columns.Count;//得到列数

                //取得数据范围区域 (不包括标题列) 
                Excel.Range rng1 = ws.Cells.get_Range("B2", "B" + rowsint);   //item
                Excel.Range rng2 = ws.Cells.get_Range("K2", "K" + rowsint); //Customer
                object[,] arryItem = (object[,])rng1.Value2;   //get range's value
                object[,] arryCus = (object[,])rng2.Value2;
                //将新值赋给一个数组
                string[,] arry = new string[rowsint - 1, 2];
                for (int i = 1; i <= rowsint - 1; i++)
                {
                    //Item_Code列
                    arry[i - 1, 0] = arryItem[i, 1].ToString();
                    //Customer_Name列
                    arry[i - 1, 1] = arryCus[i, 1].ToString();
                }
                //MessageBox.Show(arry[0, 0] + " / " + arry[0, 1] + "#" + arry[rowsint - 2, 0] + " / " + arry[rowsint - 2, 1]);
            }
            excelApp.Quit();
            excelApp = null;
            System.Diagnostics.Process[] procs = System.Diagnostics.Process.GetProcessesByName("excel");

            foreach (System.Diagnostics.Process pro in procs)
            {
                pro.Kill();//没有更好的方法,只有杀掉进程
            }
            GC.Collect();
        }
        #endregion
    }
}
