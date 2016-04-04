using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Text;

namespace UtilityClasses
{

    /********************************************************************************
    ** Author： Xiaokai Zh
    ** Created：2016-03-30
    ** Desc：数据保存到Excel 的通用类
    *********************************************************************************/

    public static class DataSave_Excel
    {
        /// <summary>
        /// 数据保存到Excel List=>
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="data"></param>
        /// <param name="filePathName"></param>
        /// <param name="columnName"></param>
        /// <returns></returns>
        public static bool SaveInExcel<T>(List<T> data,string filePathName,List<string> columnName=null) where T:class
        {
            bool isSuccessSave = false;
            //return isSuccessSave;
            Application xlApp = ExcelInit();
            if (data==null||data.Count==0||xlApp == null)
            {
                xlApp.Quit();
                GC.Collect();//强行销毁
                return isSuccessSave;
            }
            Microsoft.Office.Interop.Excel.Worksheet worksheet = xlApp.Workbooks[1].Worksheets[1];

            
            #region 对Excel第一行命名 表头生成
            List<PropertyInfo> pList = new List<PropertyInfo>();
            Array.ForEach<PropertyInfo>(typeof(T).GetProperties(), p => pList.Add(p));
            if (columnName == null || columnName.Count != pList.Count)
            {
                for (int i = 1; i < pList.Count + 1; i++)
                {
                    worksheet.Cells[1, i] = pList[i - 1].Name;
                }
            }
            else
            {
                for (int i = 1; i < pList.Count + 1; i++)
                {
                    worksheet.Cells[1, i] = columnName[i - 1];
                }
            }
            #endregion

            //写入数值
            for (int r = 0; r < data.Count; r++)
            {
                for(int pi=0;pi<pList.Count;pi++)
                {
                    Type type=data[r].GetType();
                    worksheet.Cells[r + 2, pi + 1] = type.GetProperty(pList[pi].Name).GetValue(data[r], null).ToString();
                }              
            }

            //保存 强行销毁 
            xlApp.Workbooks[1].Saved = true;
            xlApp.Workbooks[1].SaveCopyAs(filePathName);
            xlApp.Quit();
            GC.Collect();
            return isSuccessSave ;

        }

        /// <summary>
        /// 数据保存到Excel DataTable=>
        /// </summary>
        /// <param name="data"></param>
        /// <param name="filePathName"></param>
        /// <param name="columnName"></param>
        /// <returns></returns>
        public static bool SaveInExcel(System.Data.DataTable data, string filePathName, List<string> columnName = null)
        {
            bool isSuccessSave = false;
            return isSuccessSave;
        }


        private static Application ExcelInit()
        {
            Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            xlApp.Workbooks.Add(Microsoft.Office.Interop.Excel.XlWBATemplate.xlWBATWorksheet);
            return xlApp;
        }

    }
}
