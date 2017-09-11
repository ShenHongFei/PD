using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Xml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using PaperFormatDetection.Tools;
using PaperFormatDetection.Frame;
namespace PaperFormatDetection.Format
{
    public class Formula : ModuleFormat
    {
        public Formula(List<Module> modList, PageLocator locator, int masterType)
            : base(modList, locator, masterType)
        {

        }

        public override void getStyle(WordprocessingDocument doc, String fileName)
        {
            string xmlFullPath = fileName + "\\Formula.xml";
            CreateXmlFile(xmlFullPath);
            getFormulaXml(doc, xmlFullPath);
            //Console.ReadLine();
        }

        public static void CreateXmlFile(string xmlPath)
        {
            XmlDocument xmlDocx = new XmlDocument();
            //创建类型声明节点  
            XmlNode node = xmlDocx.CreateXmlDeclaration("1.0", "utf-8", "");
            xmlDocx.AppendChild(node);
            //创建根节点  
            XmlNode root = xmlDocx.CreateElement("Formula");
            XmlElement xe1 = xmlDocx.CreateElement("spErroInfo");
            xe1.SetAttribute("name", "特殊错误信息");
            XmlElement xe2 = xmlDocx.CreateElement("partName");
            xe2.SetAttribute("name", "提示名称");
            XmlElement xe3 = xmlDocx.CreateElement("Text");
            xe3.InnerText = "-----------------公式-----------------";
            xe2.AppendChild(xe3);
            root.AppendChild(xe1);
            root.AppendChild(xe2);
            xmlDocx.AppendChild(root);
            try
            {
                xmlDocx.Save(xmlPath);
            }
            catch (Exception e)
            {
                //显示错误信息  
                Console.WriteLine(e.Message);
            }
        }
        public static void getFormulaXml(WordprocessingDocument doc, String xmlFullPath)
        {
            Body body = doc.MainDocumentPart.Document.Body;
            IEnumerable<Paragraph> paras = body.Elements<Paragraph>();
            XmlDocument xmlDocx = new XmlDocument();
            xmlDocx.Load(xmlFullPath);
            XmlNode root = xmlDocx.SelectSingleNode("Formula/spErroInfo");
            List<Paragraph> pList = toList(paras);
            int count = -1;
            List<int> iList = Tool.getTitlePosition(doc);
            string chapter = "";
            string last_chapter = null;
            foreach (Paragraph p in paras)
            {
                Run r = p.GetFirstChild<Run>();
                if (r != null)
                {
                    EmbeddedObject Ob = null;
                    OfficeMath oMath = null;

                    count++;
                    Ob = r.GetFirstChild<EmbeddedObject>();
                    oMath = r.GetFirstChild<OfficeMath>();
                    if (Ob != null || oMath != null)
                    {

                        if (pList != null && count < pList.Count - 1)
                        {

                            ParagraphProperties pPr = p.GetFirstChild<ParagraphProperties>();
                            List<int> listchapter = Tool.getTitlePosition(doc);
                            chapter = Chapter(listchapter, count, body);
                            int num_incap = 1;

                            if (chapter == last_chapter)
                            {
                                num_incap++;

                            }
                            else if (chapter != last_chapter)
                            {
                                num_incap = 1;
                                last_chapter = chapter;
                            }
                            //公式编号
                            string number = getFullText(p);
                            string Show_number = number.Trim();
                            bool Isformule = false;
                            int FirstNumber = 0;
                            bool sameP = false;
                            bool havePicturename = false;
                            for (int a = 0; a < number.Length; a++)
                            {
                                if (number[a] == '图')
                                {
                                    sameP = true;
                                }
                            }
                            for (int i = 0; i < number.Length; i++)
                            {
                                if (number[i] >= 48 && number[i] <= 58 && FirstNumber == 0 && sameP == false)
                                {
                                    
                                    Isformule = true;
                                    if (i != 0)
                                        FirstNumber = i;
                                }
                                else if (number[i] >= 48 && number[i] <= 58 && FirstNumber != 0 && sameP == false)
                                {
                                    /*if (number[i - 1] != '(')
                                    {
                                        XmlElement xml = xmlDocx.CreateElement("Text");
                                        xml.InnerText = "公式编号格式不规范，建议改为（M.N）：{" + Show_number + "}";
                                        root.AppendChild(xml);
                                    }*/
                                }
                                else
                                {
                                    Paragraph Nextline = pList[count + 1];
                                    string NL = getFullText(Nextline);
                                    for (int j = 0; j < NL.Length; j++)
                                    {
                                        if ((NL[j] == '图' && j == 0 )|| (NL[j] == 'F' && j == 0 ))
                                        {
                                            havePicturename = true;
                                            /*Isformule = true;
                                            XmlElement xml = xmlDocx.CreateElement("Text");
                                            xml.InnerText = "公式缺少编号：{" + chapter + "}";
                                            root.AppendChild(xml);*/
                                        }
                                        if ((NL[j] == '图' && j != 0)|| (NL[j] == 'F' && j != 0 ))
                                        {
                                            havePicturename = true;
                                            /*if (NL[j - 1] != ' ')
                                            {
                                                Isformule = true;
                                                XmlElement xml = xmlDocx.CreateElement("Text");
                                                xml.InnerText = "公式缺少编号：{" + chapter + "}";
                                                root.AppendChild(xml);
                                            }*/
                                        }
                                    }
                                    if (havePicturename == false)
                                    {
                                        Isformule = true;
                                        XmlElement xml = xmlDocx.CreateElement("Text");
                                        xml.InnerText = "公式缺少编号：{" + chapter + "}";
                                        root.AppendChild(xml);
                                    }
                                }


                            }

                            //右对齐
                            if (pPr != null && Isformule == true)
                            {
                                if (pPr.GetFirstChild<Justification>() != null)
                                {
                                    if (pPr.GetFirstChild<Justification>().Val != null)
                                    {
                                        if (pPr.GetFirstChild<Justification>().Val != "right")
                                        {
                                            if (Show_number != null)
                                            {
                                                XmlElement xml = xmlDocx.CreateElement("Text");
                                                xml.InnerText = "公式应整行右对齐：{" + Show_number + "}";
                                                root.AppendChild(xml);
                                            }
                                            if (Show_number == null)
                                            { }
                                        }

                                    }
                                }
                            }

                            //公式前后行距
                            if (pPr != null && Isformule == true)
                            {
                                if (pPr.GetFirstChild<SpacingBetweenLines>() != null)
                                {
                                    if (pPr.GetFirstChild<SpacingBetweenLines>().Before != null && pPr.GetFirstChild<SpacingBetweenLines>().Before != "156")
                                    {
                                        if (Show_number != null)
                                        {
                                            XmlElement xml = xmlDocx.CreateElement("Text");
                                            xml.InnerText = "公式段前间距错误，应为0.5倍行距：{" + "}";
                                            root.AppendChild(xml);
                                        }
                                        if (Show_number == null)
                                        { }
                                    }
                                    if (pPr.GetFirstChild<SpacingBetweenLines>().After != null && pPr.GetFirstChild<SpacingBetweenLines>().After != "156")
                                    {
                                        if (Show_number != null)
                                        {
                                            XmlElement xml = xmlDocx.CreateElement("Text");
                                            xml.InnerText = "公式段前间距错误，应为0.5倍行距：{" + "}";
                                            root.AppendChild(xml);
                                        }
                                        if (Show_number == null)
                                        { }
                                    }
                                }
                            }
                        }
                    }

                }
            }
            xmlDocx.Save(xmlFullPath);
        }
        public static List<Paragraph> toList(IEnumerable<Paragraph> paras)
        {
            List<Paragraph> list = new List<Paragraph>();
            foreach (Paragraph p in paras)
            {
                Paragraph t = p as Paragraph;
                if (t != null)
                {
                    list.Add(t);
                }
            }
            return list;
        }
        public static String getFullText(Paragraph p)
        {
            String text = "";
            IEnumerable<Run> list = p.Elements<Run>();
            foreach (Run r in list)
            {
                Text pText = r.GetFirstChild<Text>();
                if (pText != null)
                {
                    text += pText.Text;
                }
            }
            return text;
        }

        public static string Chapter(List<int> titlePosition, int location, Body body)
        {
            string chapter = "";
            int titlelocation = -1;
            int i = 0;
            if (titlePosition.Count != 0)
            {
                for (i = 0; titlePosition[i] < location; i++)
                {
                    titlelocation = i;
                    if (i == titlePosition.Count - 1)
                        break;
                }
            }
            Paragraph p = null;
            if (titlelocation >= 0)
            {
                if (titlePosition[titlelocation] - 1 >= 0)
                {
                    p = (Paragraph)body.ChildElements.GetItem(titlePosition[titlelocation] - 1);
                }
            }
            if (p != null)
            {
                chapter = Tool.getFullText(p);
            }
            return chapter;
        }
        public static int Item(List<int> titlePosition, int location, Body body)
        {
            int item = 0;
            int titlelocation = -1;
            int i = 0;
            if (titlePosition.Count != 0)
            {
                for (i = 0; titlePosition[i] < location; i++)
                {
                    titlelocation = i;
                    if (i == titlePosition.Count - 1)
                        break;
                }
            }
            Paragraph p = null;
            if (titlelocation >= 0)
            {
                if (titlePosition[titlelocation] - 1 >= 0)
                {
                    p = (Paragraph)body.ChildElements.GetItem(titlePosition[titlelocation] - 1);
                }
            }
            if (p != null)
            {

            }
            return item;
        }
    }
}
