﻿using System;
using System.Data.SqlClient;
using System.Data;
using System.ComponentModel;
using System.Collections.Generic;
using System.Reflection;
using System.Linq;
using System.Configuration;
using System.Text;
using log4net;


namespace SQLExtend
{
    public class ADODB
    {
        private static ILog log = LogManager.GetLogger("logger");

        /// <summary>
        /// Insert
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="obj"></param>
        /// <returns></returns>
        public static int Insert<T>(object obj) where T : new()
        {
            string connString = ConfigurationManager.ConnectionStrings["LocalDB"].ToString();           
            string table = TypeDescriptor.GetClassName(obj).Split('.')[1];  // 以ClassName作為TableName

            using (SqlConnection conn = new SqlConnection(connString))
            {
                using (SqlCommand cmd = new SqlCommand())
                {
                    try
                    {
                        cmd.Connection = conn;
                        conn.Open();

                        List<PropertyInfo> properties = typeof(T).GetProperties().ToList();
                        bool FirstColFlag = true;
                        string Cols = "(";
                        string Values = "(";
                        string paras = "";
                        foreach (var property in properties)
                        {
                            // 若有提供成員變數值，則Insert包含該欄位(以該成員變數作為欄位名稱)
                            if (property.GetValue(obj) != null)
                            {
                                if (FirstColFlag)
                                {
                                    Cols = string.Concat(Cols, "[", property.Name, "]");
                                    Values = string.Concat(Values, "@", property.Name); // 這邊可以改用官方成員，取得是@或:  或用Spring切換DB
                                    cmd.Parameters.Add(new SqlParameter("@" + property.Name, property.GetValue(obj)));
                                    if (property.PropertyType == typeof(string))
                                    {
                                        paras = string.Concat(paras, property.Name, ": ", "'", property.GetValue(obj), "'");
                                    }
                                    else
                                    {
                                        paras = string.Concat(paras, property.Name, ": ", property.GetValue(obj));
                                    }
                                    FirstColFlag = false;
                                }
                                else
                                {
                                    Cols = string.Concat(Cols, ", [", property.Name, "]");
                                    Values = string.Concat(Values, ", @", property.Name); // 這邊可以改用官方成員，取得是@或:  或用Spring切換DB                
                                    cmd.Parameters.Add(new SqlParameter("@" + property.Name, property.GetValue(obj)));

                                    if (property.PropertyType == typeof(string))
                                    {
                                        paras = string.Concat(paras, ", ", property.Name, ": ", "'", property.GetValue(obj), "'");
                                    }
                                    else
                                    {
                                        paras = string.Concat(paras, ", ", property.Name, ": ", property.GetValue(obj));
                                    }
                                }
                            }
                            else
                            {
                                switch (property.Name)
                                {
                                    case "sys_createdate":
                                        if (FirstColFlag)
                                        {
                                            Cols = string.Concat(Cols, "[", property.Name, "]");
                                            Values = string.Concat(Values, @"GETDATE()");
                                            FirstColFlag = false;
                                        }
                                        else
                                        {
                                            Cols = string.Concat(Cols, ", [", property.Name, "]");
                                            Values = string.Concat(Values, ", GETDATE()");
                                        }
                                        break;
                                    case "sys_updatedate":
                                        if (FirstColFlag)
                                        {
                                            Cols = string.Concat(Cols, "[", property.Name, "]");
                                            Values = string.Concat(Values, @"GETDATE()");
                                            FirstColFlag = false;
                                        }
                                        else
                                        {
                                            Cols = string.Concat(Cols, ", [", property.Name, "]");
                                            Values = string.Concat(Values, ", GETDATE()");
                                        }
                                        break;
                                    case "sys_createuser":
                                        if (FirstColFlag)
                                        {
                                            Cols = string.Concat(Cols, "[", property.Name, "]");
                                            Values = string.Concat(Values, "'SYSTEM'");
                                            FirstColFlag = false;
                                        }
                                        else
                                        {
                                            Cols = string.Concat(Cols, ", [", property.Name, "]");
                                            Values = string.Concat(Values, ", 'SYSTEM'");
                                        }
                                        break;
                                    case "sys_updateuser":
                                        if (FirstColFlag)
                                        {
                                            Cols = string.Concat(Cols, "[", property.Name, "]");
                                            Values = string.Concat(Values, "'SYSTEM'");
                                            FirstColFlag = false;
                                        }
                                        else
                                        {
                                            Cols = string.Concat(Cols, ", [", property.Name, "]");
                                            Values = string.Concat(Values, ", 'SYSTEM'");
                                        }
                                        break;
                                    default:
                                        break;
                                }
                            }

                        }
                        Cols = string.Concat(Cols, ")");
                        Values = string.Concat(Values, ")");

                        /*** 慘痛教訓，這邊這個sql必須一氣呵成串起來，否則會有SQL語法錯誤的問題 ***/
                        string sql = string.Concat("INSERT INTO [", table, "]", Cols, " VALUES", Values);
                        log.Info(string.Concat("\n\t", sql, "\nParamaters: { ", paras, " }"));

                        cmd.CommandText = sql;
                        return cmd.ExecuteNonQuery();
                    }
                    catch (Exception ex)
                    {
                        log.Error(string.Concat("\n", ex.ToString()));
                        throw ex;
                    }
                    finally
                    {
                        if (conn.State == ConnectionState.Open)
                        {
                            conn.Close();
                        }
                    }
                }
            }
        }


        /// <summary>
        /// Update
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="obj"></param>
        /// <returns></returns>
        /// <history>
        /// 2021/09/23  Chris Liao  Create  僅建立，功能尚未測試
        /// </history>
        public static int Update<T>(object obj, string whereSql, SqlParameter wherePara) where T : new()
        {
            string connString = ConfigurationManager.ConnectionStrings["LocalDB"].ToString();
            string table = TypeDescriptor.GetClassName(obj).Split('.')[1];  // 以ClassName作為TableName
            
            using (SqlConnection conn = new SqlConnection(connString))
            {
                using (SqlCommand cmd = new SqlCommand())
                {
                    try
                    {
                        cmd.Connection = conn;
                        conn.Open();
 
                        List<PropertyInfo> properties = typeof(T).GetProperties().ToList();
                        bool FirstColFlag = true;
                        string Cols = "";
                        string paras = "";
                        foreach (var property in properties)
                        {
                            // 若有提供成員變數值，則Insert包含該欄位(以該成員變數作為欄位名稱)
                            if (property.GetValue(obj) != null)
                            {
                                if (FirstColFlag)
                                {
                                    Cols = string.Concat(Cols, property.Name, " = @", property.Name);
                                    cmd.Parameters.Add(new SqlParameter("@" + property.Name, property.GetValue(obj)));
                                    if (property.PropertyType == typeof(string))
                                    {
                                        paras = string.Concat(paras, property.Name, ": ", "'", property.GetValue(obj), "'");
                                    }
                                    else
                                    {
                                        paras = string.Concat(paras, property.Name, ": ", property.GetValue(obj));
                                    }
                                    FirstColFlag = false;
                                }
                                else
                                {
                                    Cols = string.Concat(Cols, ", ", property.Name, " = @", property.Name);        
                                    cmd.Parameters.Add(new SqlParameter("@" + property.Name, property.GetValue(obj)));

                                    if (property.PropertyType == typeof(string))
                                    {
                                        paras = string.Concat(paras, ", ", property.Name, ": ", "'", property.GetValue(obj), "'");
                                    }
                                    else
                                    {
                                        paras = string.Concat(paras, ", ", property.Name, ": ", property.GetValue(obj));
                                    }
                                }
                            }
                            else
                            {
                                switch (property.Name)
                                {
                                    case "sys_updatedate":
                                        if (FirstColFlag)
                                        {
                                            Cols = string.Concat(Cols, property.Name, " = GETDATE()");
                                            FirstColFlag = false;
                                        }
                                        else
                                        {
                                            Cols = string.Concat(Cols, ", ", property.Name, " = GETDATE()");
                                        }
                                        break;
                                    case "sys_updateuser":
                                        if (FirstColFlag)
                                        {
                                            Cols = string.Concat(Cols, property.Name, " = 'SYSTEM'");
                                            FirstColFlag = false;
                                        }
                                        else
                                        {
                                            Cols = string.Concat(Cols, ", ", property.Name, " = 'SYSTEM'");
                                        }
                                        break;
                                    default:
                                        break;
                                }
                            }

                        }
                        Cols = string.Concat(Cols);

                        whereSql = whereSql.Substring(whereSql.IndexOf(" ") + 1);   // WHERE XX=XX .... 第一個空白肯定是WHERE後
                        cmd.Parameters.Add(wherePara);  // Add傳入的WHERE Parameters

                        /*** 慘痛教訓，這邊這個sql必須一氣呵成串起來，否則會有SQL語法錯誤的問題 ***/
                        string sql = string.Concat("UPDATE [", table, "] SET ", Cols, " WHERE 1=1 ", (string.IsNullOrEmpty(whereSql) ? "" : "AND " + whereSql));
                        log.Info(string.Concat("\n\t", sql, "\nParamaters: { ", paras, " }"));

                        cmd.CommandText = sql;
                        return cmd.ExecuteNonQuery();
                    }
                    catch (Exception ex)
                    {
                        log.Error(string.Concat("\n", ex.ToString()));
                        throw ex;
                    }
                    finally
                    {
                        if (conn.State == ConnectionState.Open)
                        {
                            conn.Close();
                        }
                    }
                }
            }
        }
    }

}
