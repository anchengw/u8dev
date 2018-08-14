using System;
using System.Collections.Generic;
using System.Text;
using System.Data.SqlClient;
using System.Data;

namespace SqlDBhelper
{
    /// 用法
     /* using (DBHelper db = new DBHelper(Config.ConnStr))
        {
        }
     */

public class DBHelper : IDisposable
    {
        #region private
        private bool m_AlreadyDispose = false;
        private int m_CommandTimeout = 30;
        private string m_ConnStr;
        private SqlConnection m_Connection;
        private SqlCommand m_Command;
        #endregion
        #region 属性
        /// <summary>
        　　　　/// 数据库连接字符串
        　　　　/// </summary>
        public string ConnStr
        {
            set { m_ConnStr = value; }
            get { return m_ConnStr; }
        }
        /// <summary>
        　　　　/// 执行时间
        　　　　/// </summary>
        public int CommandTimeout
        {
            set { m_CommandTimeout = value; }
            get { return m_CommandTimeout; }
        }
        #endregion
        #region DBHelper
        /// <summary>
        　　　　/// 构造函数
        　　　　/// </summary>
        　　　　/// <param name="connStr">数据库连接字符串</param>
        public DBHelper(string connStr)
        {
            m_ConnStr = connStr;
            Initialization();
        }
        /// <summary>
        　　　　/// 构造函数
        　　　　/// </summary>
        　　　　/// <param name="connStr">数据库连接字符串</param>
        　　　　/// <param name="commandTimeout">执行时间</param>
        public DBHelper(string connStr, int commandTimeout)
        {
            m_ConnStr = connStr;
            m_CommandTimeout = commandTimeout;
            Initialization();
        }
        /// <summary>
        　　　　/// 初始化函数
        　　　　/// </summary>
        protected void Initialization()
        {
            try
            {
                m_Connection = new SqlConnection(m_ConnStr);
                if (m_Connection.State == ConnectionState.Closed)
                    m_Connection.Open();
                m_Command = new SqlCommand();
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message.ToString());
            }
        }
        #endregion
        #region Dispose
        /// <summary>
        　　　　/// 析构函数
        　　　　/// </summary>
        ~DBHelper()
        {
            Dispose();
        }
        /// <summary>
        　　　　///　释放资源
        　　　　/// </summary>
        　　　　/// <param name="isDisposing">标志</param>
        protected virtual void Dispose(bool isDisposing)
        {
            if (m_AlreadyDispose) return;
            if (isDisposing)
            {
                if (m_Command != null)
                {
                    m_Command.Cancel();
                    m_Command.Dispose();
                }
                if (m_Connection != null)
                {
                    try
                    {
                        if (m_Connection.State != ConnectionState.Closed)
                            m_Connection.Close();
                        m_Connection.Dispose();
                    }
                    catch (Exception ex)
                    {
                        throw new Exception(ex.ToString());
                    }
                    finally
                    {
                        m_Connection = null;
                    }
                }
            }
            m_AlreadyDispose = true;//已经进行的处理
        }
        /// <summary>
        　　　　/// 释放资源
        　　　　/// </summary>
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        #endregion
        #region
        #endregion
        #region ExecuteNonQuery
        public int ExecuteNonQuery(string cmdText)
        {
            try
            {
                m_Command = new SqlCommand(cmdText, m_Connection);
                m_Command.CommandTimeout = m_CommandTimeout;
                int iRet = m_Command.ExecuteNonQuery();
                return iRet;
            }
            catch (Exception ex)
            {
                //Loger.Debug(ex.ToString(),@"C:\sql.txt");
                throw new Exception(ex.ToString());
            }
            finally
            {
                if (m_Command != null)
                    m_Command.Dispose();
            }
        }
        public int ExecuteNonQuery(string cmdText, SqlParameter[] para)
        {
            if (para == null)
            {
                return ExecuteNonQuery(cmdText);
            }
            try
            {
                m_Command = new SqlCommand(cmdText, m_Connection);
                m_Command.CommandTimeout = m_CommandTimeout;
                for (int i = 0; i < para.Length; i++)
                    m_Command.Parameters.Add(para[i]);
                int iRet = m_Command.ExecuteNonQuery();
                return iRet;
            }
            catch (Exception ex)
            {
                //Loger.Debug(ex.ToString(), @"C:\sql.txt");
                throw new Exception(ex.ToString());
            }
            finally
            {
                if (m_Command != null)
                    m_Command.Dispose();
                m_Command.Parameters.Clear();
            }
        }
        public int ExecuteNonQuery(string cmdText, SqlParameter[] para, bool isStoreProdure)
        {
            if (!isStoreProdure)
            {
                return ExecuteNonQuery(cmdText, para);
            }
            try
            {
                m_Command = new SqlCommand(cmdText, m_Connection);
                m_Command.CommandTimeout = m_CommandTimeout;
                m_Command.CommandType = CommandType.StoredProcedure;
                if (para != null)
                {
                    for (int i = 0; i < para.Length; i++)
                        m_Command.Parameters.Add(para[i]);
                }
                int iRet = m_Command.ExecuteNonQuery();
                return iRet;
            }
            catch (Exception ex)
            {
                //Loger.Debug(ex.ToString(), @"C:\sql.txt");
                throw new Exception(ex.ToString());
            }
            finally
            {
                if (m_Command != null)
                    m_Command.Dispose();
                m_Command.Parameters.Clear();
            }
        }
        #endregion
        #region ExecuteTransaction
        public bool ExecuteTransaction(string[] cmdText)
        {
            SqlTransaction trans = m_Connection.BeginTransaction();
            try
            {
                m_Command = new SqlCommand();
                m_Command.Connection = m_Connection;
                m_Command.CommandTimeout = m_CommandTimeout;
                m_Command.Transaction = trans;
                for (int i = 0; i < cmdText.Length; i++)
                {
                    if (cmdText[i] != null && cmdText[i] != string.Empty)
                    {
                        m_Command.CommandText = cmdText[i];
                        m_Command.ExecuteNonQuery();
                    }
                }
                trans.Commit();
                return true;
            }
            catch (Exception ex)
            {
                trans.Rollback();
                throw new Exception(ex.ToString());
            }
            finally
            {
                if (m_Command != null)
                    m_Command.Dispose();
                trans.Dispose();
            }
        }
        public bool ExecuteTransaction(string[] cmdText, SqlParameter[] para)
        {
            if (para == null)
                return ExecuteTransaction(cmdText);
            SqlTransaction trans = m_Connection.BeginTransaction();
            try
            {
                m_Command = new SqlCommand();
                m_Command.Connection = m_Connection;
                m_Command.CommandTimeout = m_CommandTimeout;
                m_Command.Transaction = trans;
                for (int i = 0; i < para.Length; i++)
                    m_Command.Parameters.Add(para[i]);
                for (int i = 0; i < cmdText.Length; i++)
                {
                    if (cmdText[i] != null && cmdText[i] != string.Empty)
                    {
                        m_Command.CommandText = cmdText[i];
                        m_Command.ExecuteNonQuery();
                    }
                }
                trans.Commit();
                return true;
            }
            catch (Exception ex)
            {
                trans.Rollback();
                throw new Exception(ex.ToString());
            }
            finally
            {
                if (m_Command != null)
                    m_Command.Dispose();
                trans.Dispose();
            }
        }
        #endregion
        #region ExecuteScalar
        public object ExecuteScalar(string cmdText)
        {
            try
            {
                m_Command = new SqlCommand(cmdText, m_Connection);
                m_Command.CommandTimeout = m_CommandTimeout;
                object obj = m_Command.ExecuteScalar();
                if (object.Equals(obj, null) || object.Equals(obj, DBNull.Value))
                {
                    obj = null;
                }
                return obj;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.ToString());
            }
            finally
            {
                if (m_Command != null)
                    m_Command.Dispose();
            }
        }
        public object ExecuteScalar(string cmdText, SqlParameter[] para)
        {
            if (para == null)
                return ExecuteScalar(cmdText);
            try
            {
                m_Command = new SqlCommand(cmdText, m_Connection);
                m_Command.CommandTimeout = m_CommandTimeout;
                for (int i = 0; i < para.Length; i++)
                    m_Command.Parameters.Add(para[i]);
                object obj = m_Command.ExecuteScalar();
                if (object.Equals(obj, null) || object.Equals(obj, DBNull.Value))
                    obj = null;
                return obj;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.ToString());
            }
            finally
            {
                if (m_Command != null)
                    m_Command.Dispose();
                m_Command.Parameters.Clear();
            }
        }
        public object ExecuteScalar(string cmdText, SqlParameter[] para, bool isStoreProdure)
        {
            if (!isStoreProdure)
                return ExecuteScalar(cmdText, para);
            try
            {
                m_Command = new SqlCommand(cmdText, m_Connection);
                m_Command.CommandTimeout = m_CommandTimeout;
                m_Command.CommandType = CommandType.StoredProcedure;
                if (para != null)
                    for (int i = 0; i < para.Length; i++)
                        m_Command.Parameters.Add(para[i]);
                object obj = m_Command.ExecuteScalar();
                if (object.Equals(obj, null) || object.Equals(obj, DBNull.Value))
                    obj = null;
                return obj;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.ToString());
            }
            finally
            {
                if (m_Command != null)
                {
                    m_Command.Dispose();
                }
                m_Command.Parameters.Clear();
            }
        }
        #endregion
        #region ExecuteDataTable
        public DataTable ExecuteDataTable(string tableName, string cmdText)
        {
            try
            {
                DataTable myTable = new DataTable(tableName);
                m_Command = new SqlCommand(cmdText, m_Connection);
                m_Command.CommandTimeout = m_CommandTimeout;
                SqlDataAdapter da = new SqlDataAdapter(m_Command);
                da.Fill(myTable);
                da.Dispose();
                return myTable;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.ToString());
            }
            finally
            {
                if (m_Command != null)
                    m_Command.Dispose();
            }
        }
        public DataTable ExecuteDataTable(string tableName, string cmdText, SqlParameter[] para)
        {
            if (para == null)
                return ExecuteDataTable(tableName, cmdText);
            try
            {
                DataTable myTable = new DataTable(tableName);
                m_Command = new SqlCommand(cmdText, m_Connection);
                m_Command.CommandTimeout = m_CommandTimeout;
                for (int i = 0; i < para.Length; i++)
                    m_Command.Parameters.Add(para[i]);
                SqlDataAdapter da = new SqlDataAdapter(m_Command);
                da.Fill(myTable);
                da.Dispose();
                return myTable;
            }
            catch (Exception ex)
            {
                //Loger.Debug(ex.ToString(), @"C:\sql.txt");
                throw new Exception(ex.ToString());
            }
            finally
            {
                if (m_Command != null)
                    m_Command.Dispose();
                m_Command.Parameters.Clear();
            }
        }
        public DataTable ExecuteDataTable(string tableName, string cmdText, SqlParameter[] para, bool isStoreProdure)
        {
            if (!isStoreProdure)
            {
                return ExecuteDataTable(tableName, cmdText, para);
            }
            try
            {
                DataTable myTable = new DataTable(tableName);
                m_Command = new SqlCommand(cmdText, m_Connection);
                m_Command.CommandTimeout = m_CommandTimeout;
                m_Command.CommandType = CommandType.StoredProcedure;
                if (para != null)
                    for (int i = 0; i < para.Length; i++)
                        m_Command.Parameters.Add(para[i]);
                SqlDataAdapter da = new SqlDataAdapter(m_Command);
                da.Fill(myTable);
                da.Dispose();
                return myTable;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.ToString());
            }
            finally
            {
                if (m_Command != null)
                    m_Command.Dispose();
                m_Command.Parameters.Clear();
            }

        }
        #endregion
        #region ExecuteDataReader
        public SqlDataReader ExecuteDataReader(string cmdText)
        {
            try
            {
                m_Command = new SqlCommand(cmdText, m_Connection);
                m_Command.CommandTimeout = m_CommandTimeout;
                SqlDataReader reader = m_Command.ExecuteReader();
                return reader;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.ToString());
            }
            finally
            {
                if (m_Command != null)
                    m_Command.Dispose();
            }
        }
        public SqlDataReader ExecuteDataReader(string cmdText, SqlParameter[] para)
        {
            if (para == null)
            {
                return ExecuteDataReader(cmdText);
            }
            try
            {
                m_Command = new SqlCommand(cmdText, m_Connection);
                m_Command.CommandTimeout = m_CommandTimeout;
                for (int i = 0; i < para.Length; i++)
                    m_Command.Parameters.Add(para[i]);
                SqlDataReader reader = m_Command.ExecuteReader();
                return reader;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.ToString());
            }
            finally
            {
                if (m_Command != null)
                    m_Command.Dispose();
                m_Command.Parameters.Clear();
            }
        }
        public SqlDataReader ExecuteDataReader(string cmdText, SqlParameter[] para, bool isStoreProdure)
        {
            if (!isStoreProdure)
            {
                return ExecuteDataReader(cmdText, para);
            }
            try
            {
                m_Command = new SqlCommand(cmdText, m_Connection);
                m_Command.CommandTimeout = m_CommandTimeout;
                m_Command.CommandType = CommandType.StoredProcedure;
                if (para != null)
                    for (int i = 0; i < para.Length; i++)
                        m_Command.Parameters.Add(para[i]);
                SqlDataReader reader = m_Command.ExecuteReader();
                return reader;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.ToString());
            }
            finally
            {
                if (m_Command != null)
                    m_Command.Dispose();
                m_Command.Parameters.Clear();
            }
        }
        #endregion
        #region　Static
        public static SqlParameter MakeInParam(string paraName, SqlDbType paraType, object value)
        {
            SqlParameter para = new SqlParameter(paraName, paraType);
            if (Object.Equals(value, null) || Object.Equals(value, DBNull.Value) || value.ToString().Trim() == string.Empty)
                para.Value = DBNull.Value;
            else
                para.Value = value;
            return para;
        }
        public static SqlParameter MakeInParam(string paraName, SqlDbType paraType, int len, object value)
        {
            SqlParameter para = new SqlParameter(paraName, paraType, len);
            if (Object.Equals(value, null) || Object.Equals(value, DBNull.Value) || value.ToString().Trim() == string.Empty)
                para.Value = DBNull.Value;
            else
                para.Value = value;
            return para;
        }
        #endregion
    }
}