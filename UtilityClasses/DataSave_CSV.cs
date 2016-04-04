using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;

namespace UtilityClasses
{

    /********************************************************************************
    ** Author： Xiaokai Zh
    ** Created：2016-03-30
    ** Desc：数据保存到CSV 的通用类 
    *********************************************************************************/

    public static class DataSave_CSV
    {
        /// <summary>
        /// 数据保存到CSV List To CSV 
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="listData"></param>
        /// <param name="filePathName"></param>
        /// <param name="columnName"></param>
        /// <param name="encoding"></param>
        /// <returns></returns>
        public static bool SaveInCSV<T>(List<T> listData, string filePathName, List<string> columnName = null, Encoding encoding = null) where T : class
        {
            bool isSavaSuccess = false;
            Encoding realEncoding = encoding ?? Encoding.UTF8;
            FileInfo fi = new FileInfo(filePathName);
            if (!fi.Directory.Exists)
            {
                fi.Directory.Create();
            }
            FileStream fs = new FileStream(filePathName, System.IO.FileMode.Create, System.IO.FileAccess.Write);
            StreamWriter sw = new StreamWriter(fs, realEncoding);
            string data = "";
            
            //写出列名称
            List<PropertyInfo> pList = new List<PropertyInfo>();
            Array.ForEach<PropertyInfo>(typeof(T).GetProperties(), p => pList.Add(p));

            if (columnName == null || columnName.Count != pList.Count)
            {
                for (int i = 0; i < pList.Count; i++)
                {
                    data += pList[i].Name.ToString();
                    if (i < pList.Count - 1)
                    {
                        data += ",";
                    }
                }
            }
            else
            {
                for (int i = 0; i < columnName.Count; i++)
                {
                    data += columnName[i].ToString();
                    if (i < columnName.Count - 1)
                    {
                        data += ",";
                    }
                }
            }

            sw.WriteLine(data);

            //写出各行数据
            for (int i = 0; i < listData.Count; i++)
            {
                data = "";
                for (int j = 0; j < pList.Count; j++)
                {
                    Type type = listData[i].GetType();

                    string str = type.GetProperty(pList[j].Name).GetValue(listData[i], null).ToString();

                    str = str.Replace("\"", "\"\"");//替换英文冒号 英文冒号需要换成两个冒号
                    if (str.Contains(',') || str.Contains('"')
                        || str.Contains('\r') || str.Contains('\n')) //含逗号 冒号 换行符的需要放到引号中
                    {
                        str = string.Format("\"{0}\"", str);
                    }

                    data += str;
                    if (j < pList.Count - 1)
                    {
                        data += ",";
                    }
                }
                sw.WriteLine(data);
            }
            sw.Close();
            fs.Close();
            return isSavaSuccess;
        }

        /// <summary>
        /// 数据保存到CSV DataTable To CSV 
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="filePathName"></param>
        /// <param name="IsAppend"></param>
        /// <param name="encoding"></param>
        /// <returns></returns>
        public static bool SaveInCSV(DataTable dt, string filePathName, bool IsAppend, Encoding encoding = null)
        {
            bool isSavaSuccess = false;
            Encoding realEncoding = encoding ?? Encoding.UTF8;
            FileInfo fi = new FileInfo(filePathName);
            if (!fi.Directory.Exists)
            {
                fi.Directory.Create();
            }
            FileStream fs = new FileStream(filePathName, System.IO.FileMode.Create, System.IO.FileAccess.Write);
            StreamWriter sw = new StreamWriter(fs, realEncoding);
            string data = "";
            //写出列名称
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                data += dt.Columns[i].ColumnName.ToString();
                if (i < dt.Columns.Count - 1)
                {
                    data += ",";
                }
            }
            sw.WriteLine(data);
            //写出各行数据
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                data = "";
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    string str = dt.Rows[i][j].ToString();
                    str = str.Replace("\"", "\"\"");//替换英文冒号 英文冒号需要换成两个冒号
                    if (str.Contains(',') || str.Contains('"')
                        || str.Contains('\r') || str.Contains('\n')) //含逗号 冒号 换行符的需要放到引号中
                    {
                        str = string.Format("\"{0}\"", str);
                    }

                    data += str;
                    if (j < dt.Columns.Count - 1)
                    {
                        data += ",";
                    }
                }
                sw.WriteLine(data);
            }
            sw.Close();
            fs.Close();
            return isSavaSuccess;
        }

    }
}
