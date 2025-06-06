﻿using System;
using System.Collections.Generic;
using System.Text;
using System.Data.SqlClient;
using System.Data;
using System.Collections;
using System.Configuration;
namespace TimerProjectByWindowsService.Common
{

    /// <summary>
    /// SqlHelper类是专门提供给广大用户用于高性能、可升级和最佳练习的sql数据操作
    /// </summary>
    public abstract class SQLHelper
    {

        //数据库连接字符串
        //public static readonly string Con = System.Configuration.ConfigurationManager.AppSettings["SQLServer"];
        //public static readonly string Conn = System.Configuration.ConfigurationManager.ConnectionStrings["DBConn"].ConnectionString;
        // 用于缓存参数的HASH表
        private static Hashtable parmCache = Hashtable.Synchronized(new Hashtable());

        /// <summary>
        ///  给定连接的数据库用假设参数执行一个sql命令（不返回数据集）
        /// </summary>
        /// <param name="connectionString">一个有效的连接字符串</param>
        /// <param name="commandType">命令类型(存储过程, 文本, 等等)</param>
        /// <param name="commandText">存储过程名称或者sql命令语句</param>
        /// <param name="commandParameters">执行命令所用参数的集合</param>
        /// <returns>执行命令所影响的行数</returns>
        public static int ExecuteNonQuery(string connectionString, CommandType cmdType, string cmdText, params SqlParameter[] commandParameters)
        {

            SqlCommand cmd = new SqlCommand();
            cmd.CommandTimeout = 2000000;
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                try
                {
                    PrepareCommand(cmd, conn, null, cmdType, cmdText, commandParameters);
                    int val = cmd.ExecuteNonQuery();
                    cmd.Parameters.Clear();
                    return val;
                }
                catch (Exception Ex)
                {
                    return 0;
                }
                finally
                {
                    cmd.Parameters.Clear();
                }

            }
        }
        //5_1_a_s_p_x.c_o_m

        /// <summary>
        /// 用现有的数据库连接执行一个sql命令（不返回数据集）
        /// </summary>
        /// <param name="conn">一个现有的数据库连接</param>
        /// <param name="commandType">命令类型(存储过程, 文本, 等等)</param>
        /// <param name="commandText">存储过程名称或者sql命令语句</param>
        /// <param name="commandParameters">执行命令所用参数的集合</param>
        /// <returns>执行命令所影响的行数</returns>
        public static int ExecuteNonQuery(SqlConnection connection, CommandType cmdType, string cmdText, params SqlParameter[] commandParameters)
        {

            SqlCommand cmd = new SqlCommand();
            cmd.CommandTimeout = 2000000;
            PrepareCommand(cmd, connection, null, cmdType, cmdText, commandParameters);
            int val = cmd.ExecuteNonQuery();
            cmd.Parameters.Clear();
            return val;
        }

        /// <summary>
        /// 用现有的数据库连接执行一个返回数据集的sql命令
        /// </summary>
        /// <param name="connection">一个现有的数据库连接</param>
        /// <param name="cmdType">命令类型(存储过程, 文本, 等等)</param>
        /// <param name="cmdText">存储过程名称或者sql命令语句</param>
        /// <param name="commandParameters">执行命令所用参数的集合</param>
        /// <returns>包含结果的数据集</returns>
        public static DataSet ExecuteDataSet(string connection, CommandType cmdType, string cmdText, params SqlParameter[] commandParameters)
        {
            SqlCommand cmd = new SqlCommand();
            using (SqlConnection conn = new SqlConnection(connection))
            {
                PrepareCommand(cmd, conn, null, cmdType, cmdText, commandParameters);
                SqlDataAdapter adt = new SqlDataAdapter();

                adt.SelectCommand = cmd;
                adt.SelectCommand.CommandTimeout = 2000000;
                SqlCommandBuilder cbd = new SqlCommandBuilder(adt);
                DataSet ds = new DataSet();
                try
                {
                    adt.Fill(ds);
                    return ds;
                }
                catch (Exception Ex)
                {
                    return null;
                }
                finally
                {
                    cmd.Parameters.Clear();
                }
            }

        }

        /// <summary>
        /// 用现有的数据库连接执行一个返回数据表的sql命令
        /// </summary>
        /// <param name="connection">一个现有的数据库连接</param>
        /// <param name="cmdType">命令类型(存储过程, 文本, 等等)</param>
        /// <param name="cmdText">存储过程名称或者sql命令语句</param>
        /// <param name="commandParameters">执行命令所用参数的集合</param>
        /// <returns>包含结果的数据表</returns>
        public static DataTable ExecuteDataTable(SqlConnection connection, CommandType cmdType, string cmdText, params SqlParameter[] commandParameters)
        {
            SqlCommand cmd = new SqlCommand();
            PrepareCommand(cmd, connection, null, cmdType, cmdText, commandParameters);
            SqlDataAdapter adt = new SqlDataAdapter();
            adt.SelectCommand = cmd;
            adt.SelectCommand.CommandTimeout = 2000000;
            //adt.
            SqlCommandBuilder cbd = new SqlCommandBuilder(adt);

            DataSet ds = new DataSet();
            try
            {
                adt.Fill(ds);
                return ds.Tables[0];
            }
            catch
            {
                return new DataTable();
            }
            finally
            {
                cmd.Parameters.Clear();
            }


        }



        /// <summary>
        /// 用现有的数据库连接执行一个返回数据表的sql命令
        /// </summary>
        /// <param name="connection">一个现有的数据库连接</param>
        /// <param name="cmdType">命令类型(存储过程, 文本, 等等)</param>
        /// <param name="cmdText">存储过程名称或者sql命令语句</param>
        /// <param name="commandParameters">执行命令所用参数的集合</param>
        /// <returns>包含结果的数据表</returns>
        public static DataTable ExecuteDataTable(string connection, CommandType cmdType, string cmdText, params SqlParameter[] commandParameters)
        {
            SqlCommand cmd = new SqlCommand();
            using (SqlConnection conn = new SqlConnection(connection))
            {
                PrepareCommand(cmd, conn, null, cmdType, cmdText, commandParameters);
                SqlDataAdapter adt = new SqlDataAdapter();
                adt.SelectCommand = cmd;
                adt.SelectCommand.CommandTimeout = 2000000;
                SqlCommandBuilder cbd = new SqlCommandBuilder(adt);

                DataSet ds = new DataSet();
                try
                {
                    adt.Fill(ds);
                    return ds.Tables[0];
                }
                catch (Exception Ex)
                {
                    return new DataTable();
                }
                finally
                {
                    cmd.Parameters.Clear();
                }
            }
        }

        /// <summary>
        ///使用现有的SQL事务执行一个sql命令（不返回数据集）
        /// </summary>
        /// <remarks>
        ///举例:  
        ///  int result = ExecuteNonQuery(connString, CommandType.StoredProcedure, "PublishOrders", new SqlParameter("@prodid", 24));
        /// </remarks>
        /// <param name="trans">一个现有的事务</param>
        /// <param name="commandType">命令类型(存储过程, 文本, 等等)</param>
        /// <param name="commandText">存储过程名称或者sql命令语句</param>
        /// <param name="commandParameters">执行命令所用参数的集合</param>
        /// <returns>执行命令所影响的行数</returns>
        public static int ExecuteNonQuery(SqlTransaction trans, CommandType cmdType, string cmdText, params SqlParameter[] commandParameters)
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandTimeout = 2000000;
            try
            {
                PrepareCommand(cmd, trans.Connection, trans, cmdType, cmdText, commandParameters);
                int val = cmd.ExecuteNonQuery();
                return val;
            }
            catch (Exception Ex)
            {
                return -1;
            }
            finally
            {
                cmd.Parameters.Clear();
            }
        }

        /// <summary>
        /// 用执行的数据库连接执行一个返回数据集的sql命令
        /// </summary>
        /// <remarks>
        /// 举例:  
        ///  SqlDataReader r = ExecuteReader(connString, CommandType.StoredProcedure, "PublishOrders", new SqlParameter("@prodid", 24));
        /// </remarks>
        /// <param name="connectionString">一个有效的连接字符串</param>
        /// <param name="commandType">命令类型(存储过程, 文本, 等等)</param>
        /// <param name="commandText">存储过程名称或者sql命令语句</param>
        /// <param name="commandParameters">执行命令所用参数的集合</param>
        /// <returns>包含结果的读取器</returns>
        public static SqlDataReader ExecuteReader(string connectionString, CommandType cmdType, string cmdText, params SqlParameter[] commandParameters)
        {
            //创建一个SqlCommand对象
            SqlCommand cmd = new SqlCommand();
            cmd.CommandTimeout = 2000000;
            //创建一个SqlConnection对象
            SqlConnection conn = new SqlConnection(connectionString);

            //在这里我们用一个try/catch结构执行sql文本命令/存储过程，因为如果这个方法产生一个异常我们要关闭连接，因为没有读取器存在，
            //因此commandBehaviour.CloseConnection 就不会执行
            try
            {
                //调用 PrepareCommand 方法，对 SqlCommand 对象设置参数
                PrepareCommand(cmd, conn, null, cmdType, cmdText, commandParameters);
                //调用 SqlCommand  的 ExecuteReader 方法
                SqlDataReader reader = cmd.ExecuteReader(CommandBehavior.CloseConnection);
                //清除参数
                return reader;
            }
            catch
            {
                //关闭连接，抛出异常
                conn.Close();
                throw;
            }
            finally
            {
                cmd.Parameters.Clear();
            }
        }

        /// <summary>
        /// 用指定的数据库连接字符串执行一个命令并返回一个数据集的第一列
        /// </summary>
        /// <remarks>
        ///例如:  
        ///  Object obj = ExecuteScalar(connString, CommandType.StoredProcedure, "PublishOrders", new SqlParameter("@prodid", 24));
        /// </remarks>
        ///<param name="connectionString">一个有效的连接字符串</param>
        /// <param name="commandType">命令类型(存储过程, 文本, 等等)</param>
        /// <param name="commandText">存储过程名称或者sql命令语句</param>
        /// <param name="commandParameters">执行命令所用参数的集合</param>
        /// <returns>用 Convert.To{Type}把类型转换为想要的 </returns>
        public static object ExecuteScalar(string connectionString, CommandType cmdType, string cmdText, params SqlParameter[] commandParameters)
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandTimeout = 200000000;
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                PrepareCommand(cmd, connection, null, cmdType, cmdText, commandParameters);
                try
                {
                    object val = cmd.ExecuteScalar();
                    //cmd.Parameters.Clear();
                    return val;
                }
                catch (Exception Ex)
                {
                    return Ex.Message;
                }
                finally
                {
                    cmd.Parameters.Clear();
                }
            }
        }

        /// <summary>
        /// 用指定的数据库连接执行一个命令并返回一个数据集的第一列
        /// </summary>
        /// <remarks>
        /// 例如:  
        ///  Object obj = ExecuteScalar(connString, CommandType.StoredProcedure, "PublishOrders", new SqlParameter("@prodid", 24));
        /// </remarks>
        /// <param name="conn">一个存在的数据库连接</param>
        /// <param name="commandType">命令类型(存储过程, 文本, 等等)</param>
        /// <param name="commandText">存储过程名称或者sql命令语句</param>
        /// <param name="commandParameters">执行命令所用参数的集合</param>
        /// <returns>用 Convert.To{Type}把类型转换为想要的 </returns>
        public static object ExecuteScalar(SqlConnection connection, CommandType cmdType, string cmdText, params SqlParameter[] commandParameters)
        {

            SqlCommand cmd = new SqlCommand();
            cmd.CommandTimeout = 200000000;
            try
            {
                PrepareCommand(cmd, connection, null, cmdType, cmdText, commandParameters);
                object val = cmd.ExecuteScalar();
                return val;
            }
            catch (Exception Ex)
            {
                return Ex.Message;
            }
            finally
            {
                cmd.Parameters.Clear();
            }
        }


        /// <summary>
        /// 将参数集合添加到缓存
        /// </summary>
        /// <param name="cacheKey">添加到缓存的变量</param>
        /// <param name="cmdParms">一个将要添加到缓存的sql参数集合</param>
        public static void CacheParameters(string cacheKey, params SqlParameter[] commandParameters)
        {
            parmCache[cacheKey] = commandParameters;
        }

        /// <summary>
        /// 找回缓存参数集合
        /// </summary>
        /// <param name="cacheKey">用于找回参数的关键字</param>
        /// <returns>缓存的参数集合</returns>
        public static SqlParameter[] GetCachedParameters(string cacheKey)
        {
            SqlParameter[] cachedParms = (SqlParameter[])parmCache[cacheKey];

            if (cachedParms == null)
                return null;

            SqlParameter[] clonedParms = new SqlParameter[cachedParms.Length];

            for (int i = 0, j = cachedParms.Length; i < j; i++)
                clonedParms[i] = (SqlParameter)((ICloneable)cachedParms[i]).Clone();

            return clonedParms;
        }

        /// <summary>
        /// 准备执行一个命令
        /// </summary>
        /// <param name="cmd">sql命令</param>
        /// <param name="conn">Sql连接</param>
        /// <param name="trans">Sql事务</param>
        /// <param name="cmdType">命令类型例如 存储过程或者文本</param>
        /// <param name="cmdText">命令文本,例如：Select * from Products</param>
        /// <param name="cmdParms">执行命令的参数</param>
        private static void PrepareCommand(SqlCommand cmd, SqlConnection conn, SqlTransaction trans, CommandType cmdType, string cmdText, SqlParameter[] cmdParms)
        {

            if (conn.State != ConnectionState.Open)
                conn.Open();

            cmd.Connection = conn;
            cmd.CommandText = cmdText;

            if (trans != null)
                cmd.Transaction = trans;

            cmd.CommandType = cmdType;

            if (cmdParms != null)
            {
                foreach (SqlParameter parm in cmdParms)
                    cmd.Parameters.Add(parm);
            }
        }
        public static string TestDB(string dbserver, string username, string password, string dbName, string successMsg, string failMsg)
        {
            SqlConnection sqlC = new SqlConnection("server=" + dbserver + ";user id=" + username + ";password=" + password + ";database=" + dbName);
            return TestDB(sqlC, successMsg, failMsg);
        }
        public static string TestDB(SqlConnection SC, string successMsg, string failMsg)
        {
            try
            {
                SC.Open();
                return successMsg;
            }
            catch (Exception Ex)
            {
                return failMsg + "\r\n错误信息:" + Ex.Message;
            }
            finally
            {
                SC.Close();
                SC.Dispose();
            }
        }
        public static string TestDB(string dbserver, string username, string password, string dbName)
        {
            return TestDB(dbserver, username, password, dbName, "连接测试成功", "连接测试失败");
        }
        public static string TestDB(SqlConnection SC)
        {
            return TestDB(SC, "连接测试成功", "连接测试失败");
        }
    }
}
