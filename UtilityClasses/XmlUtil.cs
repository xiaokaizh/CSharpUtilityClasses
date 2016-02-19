using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Serialization;

namespace UtilityClasses
{
    public static class XmlUtil
    {
        

        #region 反序列化
        /// <summary>
        /// 反序列化
        /// </summary>
        /// <param name="type">类型</param>
        /// <param name="xml">XML字符串</param>
        /// <returns></returns>
        public static object Deserialize(Type type, string xml)
        {
            try
            {
                using (StringReader sr = new StringReader(xml))
                {
                    XmlSerializer xmldes = new XmlSerializer(type);
                    return xmldes.Deserialize(sr);
                }
            }
            catch (Exception e)
            {
                return null;
            }
        }
        /// <summary>
        /// 反序列化
        /// </summary>
        /// <param name="type"></param>
        /// <param name="xml"></param>
        /// <returns></returns>
        public static object Deserialize(Type type, Stream stream)
        {
            XmlSerializer xmldes = new XmlSerializer(type);
            return xmldes.Deserialize(stream);
        }
        #endregion


        #region 序列化
        /// <summary>
        /// 序列化
        /// </summary>
        /// <param name="type">类型</param>
        /// <param name="obj">对象</param>
        /// <returns></returns>
        public static string Serializer(Type type, object obj)
        {
            MemoryStream Stream = new MemoryStream();
            XmlSerializer xml = new XmlSerializer(type);
            try
            {
                //序列化对象
                xml.Serialize(Stream, obj);
            }
            catch (InvalidOperationException)
            {
                throw;
            }
            Stream.Position = 0;
            StreamReader sr = new StreamReader(Stream);
            string str = sr.ReadToEnd();

            sr.Dispose();
            Stream.Dispose();

            return str;
        }

        #endregion

        #region 手动生成xml
        /*
<chart>
  <categories />
  <series>
    <name>China</name>
    <data>
      <point>1</point>
      <point>2</point>
      <point>3</point>
    </data>
  </series>
  <series>
    <name>America</name>
    <data>
      <point>5</point>
      <point>6</point>
      <point>3</point>
    </data>
  </series>
  <series>
    <name>France</name>
    <data>
      <point>4</point>
      <point>8</point>
      <point>3</point>
    </data>
  </series>
</chart> 
         
         */
        
            //List<string> xAxisList = new List<string>()
            //{
            //    "Apples","Pears","Oranges"
            //};

            //List<ChartDataTemplate> data = new List<ChartDataTemplate>() { 
            //new ChartDataTemplate(){name="China",point= new long[]{1,2,3}},
            //new ChartDataTemplate(){name="America",point=new long[]{5,6,3}},
            //new ChartDataTemplate(){name="France",point=new long[]{4,8,3}},
            //};

            //XmlDocument doc = new XmlDocument();
            ////主内容
            //XmlElement root = doc.CreateElement("chart");
            //doc.AppendChild(root);

            ////生成横坐标信息集合
            //XmlElement child0 = doc.CreateElement("categories");
            //root.AppendChild(child0);

            //for (int i = 0; i < xAxisList.Count; i++)
            //{
            //    XmlElement child00 = doc.CreateElement("item");
            //    child00.InnerText = xAxisList[i];
            //    child0.AppendChild(child00);
            //}

            ////生成数据部分
            //for (int i = 0; i < data.Count; i++)
            //{
            //    XmlElement childData = doc.CreateElement("series");
            //    root.AppendChild(childData);

            //    XmlElement child1 = doc.CreateElement("name");
            //    child1.InnerText = data[i].name;
            //    childData.AppendChild(child1);

            //    XmlElement child2 = doc.CreateElement("data");
            //    childData.AppendChild(child2);

            //    for (int j = 0; j < data[i].point.Count(); j++)
            //    {
            //        XmlElement child3 = doc.CreateElement("point");
            //        child3.InnerText = data[i].point[j].ToString();
            //        child2.AppendChild(child3);
            //    }
            //}

        #endregion
    }
}
