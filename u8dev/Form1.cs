using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
//需要添加以下命名空间
using UFIDA.U8.MomServiceCommon;
using UFIDA.U8.U8MOMAPIFramework;
using UFIDA.U8.U8APIFramework;
using UFIDA.U8.U8APIFramework.Meta;
using UFIDA.U8.U8APIFramework.Parameter;
using MSXML2;
using System.Data.OleDb;
using Excel = Microsoft.Office.Interop.Excel;
using System.Collections;
using System.Data.SqlClient;

namespace u8dev
{
    public partial class Form1 : Form
    {
        string impUrl = "http://u8vmtest/u8eai/import.asp";
        U8Login.clsLogin u8Login;
        DataTable excelDt = null;
        string connStr = "";
        Dictionary<string,string> autoidList = new Dictionary<string, string>();
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            u8Login = new U8Login.clsLogin();
            String sSubId = "AS";
            String sAccID = "(default)@001";
            String sYear = "2018";
            String sUserID = "demo";
            String sPassword = "wangan";
            String sDate = "2018-08-03";
            String sServer = "u8vmtest";
            String sSerial = "";
            if (!u8Login.Login(ref sSubId, ref sAccID, ref sYear, ref sUserID, ref sPassword, ref sDate, ref
            sServer, ref sSerial))
            {
                MessageBox.Show("登陆失败，原因：" + u8Login.ShareString);
                Marshal.FinalReleaseComObject(u8Login);
                return;
            }

        }

        private void button2_Click(object sender, EventArgs e)
        {
            U8EnvContext envContext = new U8EnvContext();
            envContext.U8Login = u8Login;
            U8ApiAddress myApiAddress = new U8ApiAddress("装载单据的地址标识");
            //构造APIBroker 
            U8ApiBroker broker = new U8ApiBroker(myApiAddress, envContext); //API参数赋值 
            broker.AssignNormalValue("参数名", "参数值");
            //调用LOAD接口API 
            if (!broker.Invoke())
            {
                //错误处理 
                Exception apiEx = broker.GetException();
                if (apiEx != null)
                {
                    if (apiEx is MomSysException)
                    {
                        MomSysException sysEx = apiEx as MomSysException;
                        Console.WriteLine("系统异常：" + sysEx.Message);
                        //todo:异常处理 
                    }
                    else if (apiEx is MomBizException)
                    {
                        MomBizException bizEx = apiEx as MomBizException;
                        Console.WriteLine("API异常：" + bizEx.Message);
                        //todo:异常处理 
                    }
                }
                //结束本次调用，释放API资源 
                broker.Release();
                return;
            }
            //获取表头或表体的BO对象，如果要取原始的XMLDOM对象结果，请使用GetResult(参数名) 
            BusinessObject DomRet = broker.GetBoParam("表头或表体参数名");
            //修改获取的BO对象，对需要更改的字段重新赋值 
            DomRet[0]["字段名"] = "新的字段值";
            //获取普通返回值 
            System.String result = broker.GetReturnValue() as System.String; //获取out/inout参数值
            //结束本次调用，释放API资源 
            broker.Release();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("<ufinterface sender=\"001\" receiver=\"u8\" roottag=\"department\" docid=\" \" proc=\"Query \" codeexchanged=\"n \">")
                .Append("<department importfile=\" \" exportfile=\" \" code=\"011 \" bincrementout=\"n \">")
                .Append("<field display=\"部门编码 \" name=\"cDepCode \" operation=\" =\" value=\"1 \" logic=\" \"> ")
                .Append("</department>")
                .Append("</ufinterface>");
            System.Xml.XmlDocument dom = new System.Xml.XmlDocument();
            dom.LoadXml(sb.ToString());
            MSXML2.XMLHTTPClass xmlHttp = new MSXML2.XMLHTTPClass();
            xmlHttp.open("POST", "http://localhost:8080/U8EAI/import.asp", false, null, null);
            xmlHttp.send(dom.OuterXml);
            String responseXml = xmlHttp.responseText;
            MessageBox.Show(responseXml);
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(xmlHttp);       //COM释放
        }

        private void button4_Click(object sender, EventArgs e)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("<ufinterface sender=\"001 \" receiver=\"u8 \" roottag=\"department \" docid=\" \" proc=\"Query \" codeexchanged=\"n \">")
                .Append("<department>")
                .Append("</department>")
                .Append("</ufinterface>");
            System.Xml.XmlDocument dom = new System.Xml.XmlDocument();
            dom.LoadXml(sb.ToString());
            MSXML2.XMLHTTPClass xmlHttp = new MSXML2.XMLHTTPClass();
            xmlHttp.open("POST", "http://localhost:8080/U8EAI/import.asp", false, null, null);
            xmlHttp.send(dom.OuterXml);
            String responseXml = xmlHttp.responseText;
            MessageBox.Show(responseXml);
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(xmlHttp);       //COM释放
        }
        //U8类库：U8Distribute.dll(U8SOFT\EAI\)   Interop.U8Distribute.dll(U8SOFT\Interop\)
        //ProgID：U8Distribute.iDistribute
        //方法; String Process(String RequestXml)
        private void button5_Click(object sender, EventArgs e)
        {
            string requesXml = "";
            U8Distribute.iDistributeClass eaiBroker = new U8Distribute.iDistributeClass();
            String responseXml = eaiBroker.Process(requesXml);

            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(eaiBroker);

        }

        /// <summary>
        /// 销售订单：装载单据
        /// </summary>
        public void LoadSaleOrder(string vouchID)
        {
            try
            {
                //第二步：构造环境上下文对象，传入login，并按需设置其它上下文参数
                U8EnvContext envContext = new U8EnvContext();
                envContext.U8Login = u8Login;

                //销售所有接口均支持内部独立事务和外部事务，默认内部事务
                //如果是外部事务，则需要传递ADO.Connection对象，并将IsIndependenceTransaction属性设置为false
                //envContext.BizDbConnection = new ADO.Connection();
                //envContext.IsIndependenceTransaction = false;

                //设置上下文参数
                envContext.SetApiContext("VoucherType", 12); //上下文数据类型：int，含义：单据类型：12

                //第三步：设置API地址标识(Url)
                //当前API：装载单据的地址标识为：U8API/SaleOrder/Load
                U8ApiAddress myApiAddress = new U8ApiAddress("U8API/SaleOrder/Load");

                //第四步：构造APIBroker
                U8ApiBroker broker = new U8ApiBroker(myApiAddress, envContext);

                //第五步：API参数赋值

                //给普通参数VouchID赋值。此参数的数据类型为string，此参数按值传递，表示单据号
                broker.AssignNormalValue("VouchID", vouchID);

                //给普通参数blnAuth赋值。此参数的数据类型为System.Boolean，此参数按值传递，表示是否控制权限：true
                broker.AssignNormalValue("blnAuth", true);

                //第六步：调用API
                if (!broker.Invoke())
                {
                    //错误处理
                    Exception apiEx = broker.GetException();
                    if (apiEx != null)
                    {
                        if (apiEx is MomSysException)
                        {
                            MomSysException sysEx = apiEx as MomSysException;
                            Console.WriteLine("系统异常：" + sysEx.Message);
                            //todo:异常处理
                        }
                        else if (apiEx is MomBizException)
                        {
                            MomBizException bizEx = apiEx as MomBizException;
                            Console.WriteLine("API异常：" + bizEx.Message);
                            //todo:异常处理
                        }
                        //异常原因
                        String exReason = broker.GetExceptionString();
                        if (exReason.Length != 0)
                        {
                            Console.WriteLine("异常原因：" + exReason);
                        }
                    }
                    //结束本次调用，释放API资源
                    broker.Release();
                    //return null;
                }

                //第七步：获取返回结果

                //获取返回值
                //获取普通返回值。此返回值数据类型为System.String，此参数按值传递，表示成功为空串
                System.String result = broker.GetReturnValue() as System.String;
                //throw new Exception(result);
                if (!string.IsNullOrEmpty(result))
                    throw new Exception(result);
                //获取out/inout参数值

                //out参数domHead为BO对象(表头)，此BO对象的业务类型为销售订单。BO参数均按引用传递，具体请参考服务接口定义
                //如果要取原始的XMLDOM对象结果，请使用GetResult("domHead") as MSXML2.DOMDocument
                BusinessObject domHeadRet = broker.GetBoParam("domHead");
                Console.WriteLine("BO对象(表头)行数为：" + domHeadRet.RowCount); //获取BO对象(表头)的行数
                //获取BO对象(表头)各字段的值。字段定义详见API服务接口定义

                //MSXML2.DOMDocument domHead = broker.GetResult("domHead") as MSXML2.DOMDocument;

                //out参数domBody为BO对象(表体)，此BO对象的业务类型为销售订单。BO参数均按引用传递，具体请参考服务接口定义
                //如果要取原始的XMLDOM对象结果，请使用GetResult("domBody") as MSXML2.DOMDocument
                BusinessObject domBodyRet = broker.GetBoParam("domBody");
                Console.WriteLine("BO对象(表体)行数为：" + domBodyRet.RowCount); //获取BO对象(表体)的行数
                //获取BO对象(表体)各字段的值。以下代码示例只取第一行。字段定义详见API服务接口定义

                //MSXML2.DOMDocument domBody = broker.GetResult("domBody") as MSXML2.DOMDocument;


                //结束本次调用，释放API资源
                broker.Release();
                /*
                SO_SOMain so_somain = EntityConvert.ToSO_SOMain(domHeadRet[0]);
                SO_SODetails so_sodetail;
                for (int i = 0; i < domBodyRet.RowCount; i++)
                {
                    so_sodetail = EntityConvert.ToSO_SODetails(domBodyRet[i]);
                    so_sodetail.cSOCode = so_somain.cSOCode;
                    so_somain.List.Add(so_sodetail);
                }
                return so_somain;
                */
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);

            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Excel文件(*.xls;*.xlsx)|*.xls;*.xlsx|所有文件|*.*";
            ofd.ValidateNames = true;
            ofd.CheckPathExists = true;
            ofd.CheckFileExists = true;
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                string strFileName = ofd.FileName;
                //dataGridView1.DataSource = excelReader.getExcelToTable(strFileName);
                excelDt = excelReader.ExcelToTable(strFileName, true);
                int index = excelDt.Rows.Count - 1;
                string heji = excelDt.Rows[index]["单据日期"].ToString();
                if (heji.Equals("合计"))
                    excelDt.Rows.RemoveAt(index);
                dataGridView1.DataSource = excelDt;

                //其他代码
            }
        }

        private void dataGridView1_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
        {
            if (e.RowIndex != -1 && e.RowIndex <= dataGridView1.Rows.Count - 1)
            {
                DataGridViewRow dgrSingle = dataGridView1.Rows[e.RowIndex];
                try
                {
                    if (string.IsNullOrEmpty(dgrSingle.Cells["单价"].Value.ToString()) || string.IsNullOrEmpty(dgrSingle.Cells["金额"].Value.ToString()) || string.IsNullOrEmpty(dgrSingle.Cells["单据号"].Value.ToString()) || string.IsNullOrEmpty(dgrSingle.Cells["存货编码"].Value.ToString()))
                    {
                        dgrSingle.DefaultCellStyle.ForeColor = Color.Orange;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }
        /// <summary>
        /// 检查数据
        /// </summary>
        private bool checkData()
        {
            try
            {
                foreach (DataRow dr in excelDt.Rows)
                {
                    if (string.IsNullOrEmpty(dr["单价"].ToString()) || string.IsNullOrEmpty(dr["金额"].ToString()) || string.IsNullOrEmpty(dr["单据号"].ToString()) || string.IsNullOrEmpty(dr["存货编码"].ToString()))
                    {
                        return false;
                    }
                }
            }
            catch(Exception e)
            {
                return false;
            }
            return true;
        }
        /// <summary>
        /// 提交数据库
        /// </summary>
        public void commitDatabase()
        {
            string autoid = "";
            Double price, jine, num;
            if (checkData())
            {               
                autoidList.Clear();
                foreach (DataRow dr in excelDt.Rows)
                {
                    string sqlstr = @"select autoid from rdrecord01 as RdRecord inner join rdrecords01 as RdRecords on rdrecord.id=rdrecords.id  where RdRecord.cCode = {0} and RdRecords.cInvCode = '{1}';";
                    string upsql = @"Update rdrecords01 set iUnitCost = {0}, faCost = {0}, iPrice = {1}, iAPrice = {1} Where autoid = {2}";

                    sqlstr = string.Format(sqlstr, dr["单据号"].ToString(), dr["存货编码"].ToString());
                    autoid = (Sqlhelper.DbHelperSQL.GetSingle(sqlstr)).ToString();
                    price = Convert.ToDouble(dr["单价"]);
                    num = Convert.ToDouble(dr["数量"]);
                    jine = price * num ;                   
                    if(string.IsNullOrEmpty(autoid))
                    {
                        break;
                    }
                    upsql = string.Format(upsql, price, jine, autoid);
                    autoidList.Add(autoid,upsql);
                }
                using (SqlConnection conn = new SqlConnection(@"Data Source=u8testserver;Initial Catalog = UFDATA_003_2016; User ID = sa; Password=jiwu@2017"))
                {
                    conn.Open();
                    SqlCommand cmd = new SqlCommand();
                    cmd.Connection = conn;
                    SqlTransaction tx = conn.BeginTransaction();
                    cmd.Transaction = tx;
                    try
                    {
                        foreach (var item in autoidList)
                        {
                            string strsql = item.Value;
                            int id = int.Parse(item.Key);
                            if (strsql.Trim().Length > 1)
                            {
                                cmd.CommandText = strsql;
                                cmd.ExecuteNonQuery();
                                //执行存储过程
                                SqlCommand sqlCmd = new SqlCommand("Pu_WBRkdCostPrice", conn);
                                sqlCmd.CommandType = CommandType.StoredProcedure;
                                /*
                                sqlCmd.Parameters.Add(new SqlParameter("@sRdsID", SqlDbType.Int));
                                sqlCmd.Parameters["@sRdsID"].Value = id;

                                sqlCmd.Parameters.Add(new SqlParameter("@QuanPoint", SqlDbType.Int));
                                sqlCmd.Parameters["@QuanPoint"].Value = 2;

                                sqlCmd.Parameters.Add(new SqlParameter("@PricePoint", SqlDbType.Int));
                                sqlCmd.Parameters["@PricePoint"].Value = 2;

                                sqlCmd.Parameters.Add(new SqlParameter("@NumPoint", SqlDbType.Int));
                                sqlCmd.Parameters["@NumPoint"].Value = 2;

                                sqlCmd.Parameters.Add(new SqlParameter("@iError", SqlDbType.Int));
                                sqlCmd.Parameters["@iError"].Value = DBNull.Value;
                                */
                                SqlParameter[] param = new SqlParameter[]
                                {
                                   new SqlParameter("@sRdsID",id),
                                   new SqlParameter("@QuanPoint",2),
                                   new SqlParameter("@PricePoint",2),
                                   new SqlParameter("@NumPoint", 2),
                                   new SqlParameter("@iError",SqlDbType.Int,4)
                                };
                                param[4].Value = DBNull.Value;
                                param[4].Direction = ParameterDirection.Output;
                                //param[4].Direction = ParameterDirection.ReturnValue;
                                foreach (SqlParameter parameter in param)
                                {
                                    sqlCmd.Parameters.Add(parameter);
                                }
                                sqlCmd.Transaction = tx;
                                sqlCmd.ExecuteNonQuery();
                                object obj = sqlCmd.Parameters["@iError"].Value;
                                if (obj.ToString() == "-1") //出错
                                {
                                    throw new Exception("存储过程Pu_WBRkdCostPrice出错!");
                                }
                            }
                        }
                        tx.Commit();
                    }
                    catch (System.Data.SqlClient.SqlException E)
                    {
                        tx.Rollback();
                        throw new Exception(E.Message);
                    }
                }
                
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            commitDatabase();
        }
    }
}
