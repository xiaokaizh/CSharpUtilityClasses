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
    ** Desc：数据保存到Txt 的通用类 
    **
    *********************************************************************************/

    public static class DataSave_Txt
    {
        /// <summary>
        /// 数据保存到Text List To txt 
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="data"></param>
        /// <param name="filePathName"></param>
        /// <param name="path"></param>
        /// <param name="encoding"></param>
        /// <returns></returns>
        public static bool SaveInTxt<T>(List<T> data, string filePathName,bool IsAppend,Encoding encoding=null) where T : class
        {
            Encoding realEncoding = encoding ?? Encoding.UTF8;
            bool isSavaSuccess = false;
            if (data == null||data.Count == 0)
            {
                return isSavaSuccess;
            }

            try
            {
                //List 转为 DataTable
                List<PropertyInfo> pList = new List<PropertyInfo>();
                DataTable dt = new DataTable();
                Array.ForEach<PropertyInfo>(typeof(T).GetProperties(), p => { pList.Add(p); dt.Columns.Add(p.Name, p.PropertyType); });
                foreach (var item in data)
                {
                    DataRow dr = dt.NewRow();
                    pList.ForEach(p => dr[p.Name] = p.GetValue(item, null));
                    dt.Rows.Add(dr);
                }
                WriteInStream(dt, filePathName, IsAppend, realEncoding);
                isSavaSuccess = true;
            }
            catch (Exception e)
            {
                throw e;
            }
            return isSavaSuccess;
        }

        /// <summary>
        /// 数据保存到Text DataTable To txt
        /// </summary>
        /// <param name="data"></param>
        /// <param name="filePathName"></param>
        /// <param name="IsAppend"></param>
        /// <param name="encoding"></param>
        /// <returns></returns>
        public static bool SaveInTxt(DataTable data, string filePathName, bool IsAppend, Encoding encoding = null)
        {
            Encoding realEncoding = encoding ?? Encoding.UTF8;
            bool isSavaSuccess = false;
            if (data == null||data.Rows.Count==0)
            {
                return isSavaSuccess;
            }
            try
            {
                WriteInStream(data, filePathName, IsAppend, realEncoding);
                isSavaSuccess = true;
            }
            catch (Exception e)
            {               
                throw e;
            }
            return isSavaSuccess;
        }

        private static void WriteInStream(DataTable dt, string filePath, bool IsAppend, Encoding encoding)
        {
            //写入 写入txt的内容可以直接复制到Excel中
            using (StreamWriter sw = new StreamWriter(filePath, IsAppend, encoding))
            {
                //获取表头
                string dtHeader = string.Empty;
                foreach (DataColumn dtColum in dt.Columns)
                {
                    dtHeader = dtHeader + "\t" + dtColum.ColumnName;
                }
                sw.WriteLine(dtHeader);

                //每一行数据写入
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    string everyTxtRow = string.Empty;
                    object[] s = dt.Rows[i].ItemArray;
                    foreach (var item in s)
                    {
                        everyTxtRow = everyTxtRow + "\t" + item.ToString();
                    }
                    sw.WriteLine(everyTxtRow);
                }
                sw.Close();
            }
        }
    }
}
