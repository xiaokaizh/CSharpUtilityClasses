using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using UtilityClasses;

namespace TestUtility
{
    class Program
    {
        //C:\Users\Administrator\Desktop\日志查询_2016_04_01.txt
        static void Main(string[] args)
        {
            List<Person> p = new List<Person>() { 
                new Person() { Name = "赵晓凯" ,Age=23,Gender="男"},
                new Person() { Name = "黎明",Age=46 ,Gender="男"},
                new Person() { Name="刘德华",Age=25,Gender="男"}};
            //DataTable dt=TDDataConvert.ToDataTable<Person>(p);
            string filePathName = "D:\\test\\text.csv";
            List<string> columnName = new List<string>() { "姓名", "年龄","性别" };
            //DataSave_Txt.SaveInTxt<Person>(p, filePathName, true);
            DataSave_CSV.SaveInCSV<Person>(p, filePathName, columnName);
        }


    }

    class Person
    {
        string name;

        public string Name
        {
            get { return name; }
            set { name = value; }
        }

        private int age;

        public int Age
        {
            get { return age; }
            set { age = value; }
        }

        private string gender;

        public string Gender
        {
            get { return gender; }
            set { gender = value; }
        }
    }
}
