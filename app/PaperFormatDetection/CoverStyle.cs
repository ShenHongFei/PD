

using System;
using System.Linq;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Xml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using PaperFormatDetection.Tools;
using PaperFormatDetection.Frame;

using System.IO;
using DocumentFormat.OpenXml;

namespace PaperFormatDetection.Format
{
    public class CoverStyle : ModuleFormat
    {
        String[] charList = { "a", "an", "the", "of", "in", "on", "from", "and", "for", "is", "are", "with", "A", "An", "The", "Of", "In", "On", "From", "And", "For", "Is", "Are", "With" };
        string CNname = null;
        //构造函数
        public CoverStyle(List<Module> modList, PageLocator locator,int masterType)
            : base(modList, locator,masterType)
        {

        }
        public override void getStyle(WordprocessingDocument doc, String fileName)
        {
            string xmlFullPath = fileName + "\\CoverStyle.xml";//xml模板文件保存路径
            CreateXmlFile(xmlFullPath);
            getHeadingXML(doc, xmlFullPath);
            getCNTitleXML(doc, xmlFullPath);
            getENTitleXML(doc, xmlFullPath);
            //getStudentInfoXml(doc), xmlFullPath);
            getCNLogoXML(doc, xmlFullPath);
        }

        private void CreateXmlFile(string xmlPath)
        {
            XmlDocument xmlDoc = new XmlDocument();
            //创建类型声明节点  
            XmlNode node = xmlDoc.CreateXmlDeclaration("1.0", "utf-8", "");
            xmlDoc.AppendChild(node);
            //创建根节点  
            XmlNode root = xmlDoc.CreateElement("CoverStyle");

            XmlElement xe1 = xmlDoc.CreateElement("Title");
            xe1.SetAttribute("name", "标题");
            XmlElement xe2 = xmlDoc.CreateElement("StudentInfo");
            xe2.SetAttribute("name", "学生信息");
            XmlElement xe3 = xmlDoc.CreateElement("CoverLogo");
            xe3.SetAttribute("name", "封面标注");
            XmlElement xe4 = xmlDoc.CreateElement("spErroInfo");
            xe4.SetAttribute("name", "特殊错误信息");
            XmlElement xe5 = xmlDoc.CreateElement("partName");
            xe5.SetAttribute("name", "提示名称");
            XmlElement xesum = xmlDoc.CreateElement("Text");
            xesum.InnerText = "红蚂蚁实验室提供技术支持";
            root.AppendChild(xesum);
            XmlElement xe6 = xmlDoc.CreateElement("Text");
            xe6.InnerText = "-----------------封面-----------------";
            root.AppendChild(xe1);
            root.AppendChild(xe2);
            root.AppendChild(xe3);
            root.AppendChild(xe4);
            root.AppendChild(xe5);
            xe5.AppendChild(xe6);
            xmlDoc.AppendChild(root);
            try
            {
                xmlDoc.Save(xmlPath);
            }
            catch (Exception e)
            {
                //显示错误信息  
                Console.WriteLine(e.Message);
            }
        }
  

        private void getHeadingXML(WordprocessingDocument doc,string xmlFullPath)
        {
            Body body = doc.MainDocumentPart.Document.Body;
            MainDocumentPart mainPart = doc.MainDocumentPart;
            IEnumerable<Paragraph> paras = body.Elements<Paragraph>();
            XmlDocument xml = new XmlDocument();
            xml.Load(xmlFullPath);
            XmlNode errInfor = xml.SelectSingleNode("CoverStyle/spErroInfo");
            bool flag1 = true;
            bool flag2 = true;
            bool flag3 = true;
      
            int count = 0;

            foreach (Paragraph p in paras)
            {
                Run r = p.GetFirstChild<Run>();
                if (r == null)
                    continue;
                String fullText = Tool.getFullText(p).Trim();
                if (fullText != "")
                    count++;
                if (count == 1 && fullText != "")
                {

                    ParagraphProperties prp = p.GetFirstChild<ParagraphProperties>();
                    IEnumerable<Run> rs = p.Elements<Run>();

                    if (prp != null)
                    {
                        if (prp.GetFirstChild<Justification>() != null)
                        {
                            if (prp.GetFirstChild<Justification>().Val.ToString().ToLower() != "center")
                            {
                                XmlElement error = xml.CreateElement("Text");
                                error.InnerText = "封面大标题未居中";
                                errInfor.AppendChild(error);
                                prp.GetFirstChild<Justification>().Val = JustificationValues.Center;
                            }
                        }
                        else
                        {
                            Justification justification = new Justification() { Val = JustificationValues.Center };
                            prp.Append(justification);
                        }
                    }

                   // string sentence2 = Regex.Replace(fullText, @"\s*", "");

                   
                       // if (sentence2 != "大连理工大学本科毕业设计（论文）")
                       // {

                            if (p.Elements<Run>().Count() == 1)
                            {
                                p.GetFirstChild<Run>().GetFirstChild<Text>().Text = "大连理工大学本科毕业设计（论文）";
                            }
                            else
                            {
                                IEnumerable<Run> runs = p.Elements<Run>();
                                int num = 0;
                                foreach (Run rr in runs)
                                {
                                    num++;
                                    if (num != 1)
                                        if (rr.GetFirstChild<Text>() != null)
                                        {
                                            if (rr.GetFirstChild<Text>().Text != null)
                                            {
                                                rr.GetFirstChild<Text>().Text = null;
                                            }
                                        }
                                }

                                p.GetFirstChild<Run>().GetFirstChild<Text>().Text = "大连理工大学本科毕业设计（论文）";
                            }


                      //  }
                    
                 
                    if (rs != null)
                    {




                        foreach (Run rr in rs)
                        {
                            if (rr != null)
                            {


                                RunProperties rpr = rr.GetFirstChild<RunProperties>();
                                if (rpr != null)
                                {
                                    if (rpr.GetFirstChild<RunFonts>() != null)
                                    {
                                        if (rpr.GetFirstChild<RunFonts>().Ascii != null)
                                        {
                                            if (rpr.GetFirstChild<RunFonts>().Ascii != "宋体")
                                            {
                                                if (rpr.GetFirstChild<RunFonts>().Hint != "eastAsia")
                                                {
                                                    flag2 = false;
                                                }
                                            }
                                        }

                                        rpr.GetFirstChild<RunFonts>().Ascii = "宋体";
                                        rpr.GetFirstChild<RunFonts>().HighAnsi = "宋体";
                                        rpr.GetFirstChild<RunFonts>().ComplexScript = "宋体";
                                        rpr.GetFirstChild<RunFonts>().EastAsia = "宋体";

                                    }

                                    else
                                    {
                                        RunFonts runfont = new RunFonts() {/* Hint = FontTypeHintValues.EastAsia, */Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "宋体", EastAsia = "宋体" };
                                        rpr.Append(runfont);
                                    }

                                    if (rpr.GetFirstChild<FontSize>() != null)
                                    {
                                        if (rpr.GetFirstChild<FontSize>().Val != null)
                                        {
                                            if (rpr.GetFirstChild<FontSize>().Val != "48")
                                            {
                                                flag1 = false;
                                                rpr.GetFirstChild<FontSize>().Val = "48";
                                            }
                                        }
                                    }

                                    else
                                    {
                                        FontSize fontSize1 = new FontSize() { Val = "48" };
                                        rpr.Append(fontSize1);
                                    }

                                    if (rpr.GetFirstChild<Bold>() == null)
                                    {
                                        if (prp != null)
                                        {
                                            if (prp.GetFirstChild<ParagraphStyleId>() == null)
                                            {
                                                flag3 = false;
                                            }
                                        }
                                       
                                        Bold bold = new Bold() { Val = true };
                                        rpr.Append(bold);
                                    }
                                    else
                                    {
                                        if (rpr.GetFirstChild<Bold>().Val == null || rpr.GetFirstChild<Bold>().Val != true)
                                        {
                                       
                                            rpr.GetFirstChild<Bold>().Val = true;
                                        }
                                    }

                                }
                            }
                        }


                    }
                    if (!flag2)
                    {
                        XmlElement error = xml.CreateElement("Text");
                        error.InnerText = "封面大标题字体错误，应为宋体";
                        errInfor.AppendChild(error);
                    }
                    if (!flag3)
                    {
                        XmlElement error = xml.CreateElement("Text");
                        error.InnerText = "封面大标题未加粗";
                        errInfor.AppendChild(error);
                    }
                    if (!flag1)
                    {
                        XmlElement error = xml.CreateElement("Text");
                        error.InnerText = "封面大标题字号错误，应为小一";
                        errInfor.AppendChild(error);
                    }
                    xml.Save(xmlFullPath);
                    break;
                }
            }


        }





        private void getCNTitleXML(WordprocessingDocument doc,string xmlFullPath)
        {
            MainDocumentPart mainPart = doc.MainDocumentPart;
            Body body = mainPart.Document.Body;
            IEnumerable<Paragraph> paras = body.Elements<Paragraph>();
            ParagraphProperties pPr = null;
            //XmlDocument xml = new XmlDocument();
            XmlDocument xml = new XmlDocument();
            xml.Load(xmlFullPath);
            XmlNode errInfor = xml.SelectSingleNode("CoverStyle/spErroInfo");
            bool flag1 = true;
            bool flag2 = true;
            bool flag3 = true;
           // bool flag4 = true;
           
            string fulltext = "";
            int count = 0;
            List<Paragraph> plist = toList(paras);
            int number = -1;
            bool twolines = false;

            foreach (Paragraph p in paras)
            {
                number++;
                Run r = p.GetFirstChild<Run>();
                if (r == null)
                    continue;
                fulltext = Tool.getFullText(p).Trim();


                if (fulltext != "")
                {
                    count++;
                }
                if (count == 2)
                {

                    if (hasChinese(Tool.getFullText(plist[number + 1])))
                    {

                        fulltext = Tool.getFullText(p).Trim() + Tool.getFullText(plist[number + 1]).Trim();

                        twolines = true;
                    }
                    else
                    {
                        fulltext = Tool.getFullText(p).Trim();
                    }


                   /* if (Tool.GetHanNumFromString(fulltext) > 20)
                    {
                       
                        addComment(doc, p, "中文标题应少于20字");

                    }*/

                    List<Paragraph> pp = new List<Paragraph>();

                    if (twolines)
                    {

                        pp.Add(p);
                        pp.Add(plist[number + 1]);
                    }
                    else
                    {
                        pp.Add(p);
                    }

                    foreach (Paragraph singlep in pp)
                    {

                        pPr = singlep.GetFirstChild<ParagraphProperties>();

                        if (pPr != null)
                        {
                            if (pPr.GetFirstChild<Justification>() != null)
                            {
                                if (pPr.GetFirstChild<Justification>().Val.ToString() != "center")
                                {
                                    XmlElement error = xml.CreateElement("Text");
                                error.InnerText = "中文标题未居中";
                                errInfor.AppendChild(error);
                                    pPr.GetFirstChild<Justification>().Val = JustificationValues.Center;

                                }
                            }
                            else
                            {
                                Justification justification = new Justification() { Val = JustificationValues.Center };
                                pPr.Append(justification);
                            }

                            if (pPr.GetFirstChild<SpacingBetweenLines>() != null)
                            {
                                if (pPr.GetFirstChild<SpacingBetweenLines>().Line != null)
                                {
                                    if (pPr.GetFirstChild<SpacingBetweenLines>().Line.Value != "300")
                                    {
                                         XmlElement error = xml.CreateElement("Text");
                                error.InnerText = "中文标题行距错误，应为1.25";
                                errInfor.AppendChild(error);
                                        pPr.GetFirstChild<SpacingBetweenLines>().Line.Value = "300";
                                    }
                                }
                                else
                                {
                                    SpacingBetweenLines spacebetweenlines = new SpacingBetweenLines() { Line = "300" };
                                    pPr.Append(spacebetweenlines);
                                }
                                if (pPr.GetFirstChild<SpacingBetweenLines>().BeforeLines != null && pPr.GetFirstChild<SpacingBetweenLines>().Before != null)
                                {
                                    if (pPr.GetFirstChild<SpacingBetweenLines>().BeforeLines != 0 || pPr.GetFirstChild<SpacingBetweenLines>().Before.Value != "0")
                                    {

                                        if (pPr.GetFirstChild<SpacingBetweenLines>().BeforeLines != null)
                                            pPr.GetFirstChild<SpacingBetweenLines>().BeforeLines = 0;

                                        if (pPr.GetFirstChild<SpacingBetweenLines>().Before != null)
                                            pPr.GetFirstChild<SpacingBetweenLines>().Before.Value = "0";

                                    }
                                }

                                if (pPr.GetFirstChild<SpacingBetweenLines>().AfterLines != null && pPr.GetFirstChild<SpacingBetweenLines>().After != null)
                                {
                                    if (pPr.GetFirstChild<SpacingBetweenLines>().AfterLines != 0 || pPr.GetFirstChild<SpacingBetweenLines>().After.Value != "0")
                                    {
                                        if (pPr.GetFirstChild<SpacingBetweenLines>().AfterLines != null)
                                            pPr.GetFirstChild<SpacingBetweenLines>().AfterLines = 0;

                                        if (pPr.GetFirstChild<SpacingBetweenLines>().After != null)
                                            pPr.GetFirstChild<SpacingBetweenLines>().After.Value = "0";

                                    }
                                }

                            }
                        }

                        IEnumerable<Run> runs = singlep.Elements<Run>();
                        if (runs != null)
                        {

                            RunProperties rPr = null;

                            foreach (Run rr in runs)
                            {
                                if (rr != null)
                                {
                                    rPr = rr.GetFirstChild<RunProperties>();
                                    if (rPr != null)
                                    {
                                        if (rPr.GetFirstChild<RunFonts>() != null)
                                        {
                                            if (rPr.GetFirstChild<RunFonts>().Ascii != null)
                                            {
                                                if (rPr.GetFirstChild<RunFonts>().Ascii != "华文细黑")
                                                {
                                                    flag1 = false;
                                                    rPr.GetFirstChild<RunFonts>().Ascii = "华文细黑";
                                                    rPr.GetFirstChild<RunFonts>().HighAnsi = "华文细黑";
                                                    rPr.GetFirstChild<RunFonts>().ComplexScript = "华文细黑";
                                                    rPr.GetFirstChild<RunFonts>().EastAsia = "华文细黑";

                                                }
                                            }
                                        }

                                        else
                                        {
                                            RunFonts runfont = new RunFonts() {/* Hint = FontTypeHintValues.EastAsia, */Ascii = "华文细黑", HighAnsi = "华文细黑", ComplexScript = "华文细黑", EastAsia = "华文细黑" };
                                            rPr.Append(runfont);
                                        }




                                        if (rPr.GetFirstChild<FontSize>() != null)
                                        {
                                            if (rPr.GetFirstChild<FontSize>().Val != null)
                                            {
                                                if (rPr.GetFirstChild<FontSize>().Val != "44")
                                                 flag2 = false;
                                                { rPr.GetFirstChild<FontSize>().Val = "44"; }
                                            }
                                        }
                                        else
                                        {

                                            FontSize fontSize1 = new FontSize() { Val = "44" };
                                            rPr.Append(fontSize1);
                                        }


                                        if (rPr.GetFirstChild<Bold>() == null)
                                        {
                                            if (pPr != null)
                                            {
                                                if (pPr.GetFirstChild<ParagraphStyleId>() == null)
                                                {
                                                    flag3 = false;
                                                }
                                            }
                                            Bold bold = new Bold() { Val = true };
                                            rPr.Append(bold);
                                        }
                                        else
                                        {
                                            if (rPr.GetFirstChild<Bold>().Val == null || rPr.GetFirstChild<Bold>().Val != true)
                                            {
                                                
                                                rPr.GetFirstChild<Bold>().Val = true;
                                            }
                                        }
                                    }
                                }
                            }


                        }
                    }
                    if (!flag1)
                    {
                        XmlElement error = xml.CreateElement("Text");
                        error.InnerText = "封面中文标题字体错误，应为华文细黑";
                        errInfor.AppendChild(error);
                    }
                    if (!flag2)
                    {
                        XmlElement error = xml.CreateElement("Text");
                        error.InnerText = "封面中文标题字号错误，应为二号";
                        errInfor.AppendChild(error);
                    }
                    break;
                }

            }
            xml.Save(xmlFullPath);

        }






        private void getENTitleXML(WordprocessingDocument doc, string xmlFullPath)
        {
            Body body = doc.MainDocumentPart.Document.Body;
            MainDocumentPart mainPart = doc.MainDocumentPart;
            IEnumerable<Paragraph> paras = body.Elements<Paragraph>();
            ParagraphProperties pPr = null;
            //XmlDocument xmlDoc = new XmlDocument();
            XmlDocument xml = new XmlDocument();
            xml.Load(xmlFullPath);
            XmlNode errInfor = xml.SelectSingleNode("CoverStyle/spErroInfo");
            bool flag1 = true;
            bool flag2 = true;
            bool flag3 = true;
            bool flag4 = true;
            bool flag5 = true;
            // bool flag6 = true;

            bool twolines = false;
            int number = -1;
            List<Paragraph> plist = toList(paras);


            foreach (Paragraph p in paras)
            {
                number++;
                Run r = p.GetFirstChild<Run>();
                if (r == null)
                    continue;
                String fullText = "";
                if (r != null || p.GetFirstChild<Hyperlink>() != null)
                {
                    fullText = Tool.getFullText(p).Trim();
                    //Console.WriteLine(fullText);
                    if (fullText != "" && !hasChinese(fullText))
                    {

                        if (!hasChinese(Tool.getFullText(plist[number + 1])))
                        {

                            twolines = true;
                        }
                        else
                        {
                            //fullText = Tool.getFullText(p).Trim();
                        }
                    }
                }

                List<Paragraph> pp = new List<Paragraph>();

                if (twolines)
                {
                    //Console.WriteLine("1234567");

                    pp.Add(p);
                    pp.Add(plist[number + 1]);
                }
                else if (fullText != "" && !hasChinese(fullText))
                {
                    pp.Add(p);
                }
                else
                {
                    continue;
                }

                foreach (Paragraph singlep in pp)
                {
                    //flag= true;
                    fullText = Tool.getFullText(singlep).Trim();
                    //Console.WriteLine(fullText);


                    if (fullText != "")
                    {
                        if (!hasChinese(fullText))
                        {
                            String[] strList = fullText.Trim().Split(' ');

                            pPr = singlep.GetFirstChild<ParagraphProperties>();
                            IEnumerable<Run> pRunList = singlep.Elements<Run>();
                            string ChangedText = "";
                            for (int i = 0; i < strList.Length; i++)
                            {

                                bool strFalg = true;
                                for (int j = 0; j < charList.Length; j++)
                                {
                                    if (strList[i] == charList[j])
                                    {
                                        strFalg = false;
                                    }

                                    if (Regex.IsMatch(strList[i].Substring(0, 1), "[A-Z]") || Regex.IsMatch(strList[i].Substring(0, 1), "[a-z]"))
                                    {
                                        if (strFalg)
                                        {
                                            if (strList[i].Length > 1)
                                            {
                                                if (!Regex.IsMatch(strList[i].Substring(0, 1), "[A-Z]"))
                                                {
                                                    flag1 = false;
                                                    string FirstLetter = strList[i].Substring(0, 1).ToUpper();
                                                    strList[i] = FirstLetter + strList[i].Substring(1);
                                                }
                                            }
                                        }
                                        else
                                        {
                                            if (strList[i].Length > 1)
                                            {
                                                if (!Regex.IsMatch(strList[i].Substring(0, 1), "[a-z]"))
                                                {
                                                    flag2 = false;
                                                    string FirstLetter = strList[i].Substring(0, 1).ToLower();
                                                    strList[i] = FirstLetter + strList[i].Substring(1);
                                                }
                                            }
                                        }
                                    }




                                        }
                                ChangedText += strList[i];
                                if (i != strList.Length - 1)
                                    ChangedText += " ";
                                    }
                                


                                if (singlep.Elements<Run>().Count() == 1)
                                {
                                    singlep.GetFirstChild<Run>().GetFirstChild<Text>().Text = ChangedText;
                                }
                                else
                                {
                                    IEnumerable<Run> runs = singlep.Elements<Run>();
                                    int num = 0;
                                    foreach (Run rr in runs)
                                    {

                                        num++;
                                        if (num != 1)
                                        {
                                            if (rr.GetFirstChild<Text>() != null)
                                                rr.GetFirstChild<Text>().Text = null;
                                        }
                                    }

                                    singlep.GetFirstChild<Run>().GetFirstChild<Text>().Text = ChangedText;
                                }



                                if (pPr != null)
                                {
                                    if (pPr.GetFirstChild<Justification>() != null)
                                    {
                                        if (pPr.GetFirstChild<Justification>().Val.ToString() != "center")
                                        {
                                            XmlElement error = xml.CreateElement("Text");
                                            error.InnerText = "封面英文标题未居中";
                                            errInfor.AppendChild(error);
                                            pPr.GetFirstChild<Justification>().Val = JustificationValues.Center;

                                        }
                                    }
                                    else
                                    {
                                        Justification justification = new Justification() { Val = JustificationValues.Center };
                                        pPr.Append(justification);
                                    }

                                    if (pPr.GetFirstChild<SpacingBetweenLines>() != null)
                                    {
                                        if (pPr.GetFirstChild<SpacingBetweenLines>().Line != null)
                                        {
                                            if (pPr.GetFirstChild<SpacingBetweenLines>().Line.Value != "240")
                                            {
                                                XmlElement error = xml.CreateElement("Text");
                                                error.InnerText = "封面英文标题行距应为1.0";
                                                errInfor.AppendChild(error);
                                                pPr.GetFirstChild<SpacingBetweenLines>().Line.Value = "240";
                                            }
                                        }
                                        else
                                        {
                                            SpacingBetweenLines spacebetweenlines = new SpacingBetweenLines() { Line = "240" };
                                            pPr.Append(spacebetweenlines);
                                        }
                                        if (pPr.GetFirstChild<SpacingBetweenLines>().BeforeLines != null && pPr.GetFirstChild<SpacingBetweenLines>().Before != null)
                                        {
                                            if (pPr.GetFirstChild<SpacingBetweenLines>().BeforeLines != 0 || pPr.GetFirstChild<SpacingBetweenLines>().Before.Value != "0")
                                            {

                                                if (pPr.GetFirstChild<SpacingBetweenLines>().BeforeLines != null)
                                                    pPr.GetFirstChild<SpacingBetweenLines>().BeforeLines = 0;

                                                if (pPr.GetFirstChild<SpacingBetweenLines>().Before != null)
                                                    pPr.GetFirstChild<SpacingBetweenLines>().Before.Value = "0";



                                            }
                                        }

                                        if (pPr.GetFirstChild<SpacingBetweenLines>().AfterLines != null && pPr.GetFirstChild<SpacingBetweenLines>().After != null)
                                        {
                                            if (pPr.GetFirstChild<SpacingBetweenLines>().AfterLines != 0 || pPr.GetFirstChild<SpacingBetweenLines>().After.Value != "0")
                                            {
                                                if (pPr.GetFirstChild<SpacingBetweenLines>().AfterLines != null)
                                                    pPr.GetFirstChild<SpacingBetweenLines>().AfterLines = 0;

                                                if (pPr.GetFirstChild<SpacingBetweenLines>().After != null)
                                                    pPr.GetFirstChild<SpacingBetweenLines>().After.Value = "0";


                                            }
                                        }
                                    }
                                }
                                if (pRunList != null)
                                {
                                    foreach (Run pr in pRunList)
                                    {
                                        if (pr != null)
                                        {
                                            RunProperties Rrpr = pr.GetFirstChild<RunProperties>();
                                            if (Rrpr != null)
                                            {
                                                if (Rrpr.GetFirstChild<RunFonts>() != null)
                                                {
                                                    if (Rrpr.GetFirstChild<RunFonts>().Ascii != null)
                                                    {
                                                        if (Rrpr.GetFirstChild<RunFonts>().Ascii != "Times New Roman")
                                                        {
                                                            flag3 = false;
                                                            Rrpr.GetFirstChild<RunFonts>().Ascii = "Times New Roman";
                                                            Rrpr.GetFirstChild<RunFonts>().HighAnsi = "Times New Roman";
                                                            Rrpr.GetFirstChild<RunFonts>().ComplexScript = "Times New Roman";
                                                            Rrpr.GetFirstChild<RunFonts>().EastAsia = "Times New Roman";

                                                        }
                                                    }
                                                }
                                                else
                                                {
                                                    RunFonts runfonts = new RunFonts()
                                                    {
                                                        Ascii = "Times New Roman",
                                                        HighAnsi = "Times New Roman",
                                                        ComplexScript = "Times New Roman",
                                                        EastAsia = "Times New Roman"
                                                    };
                                                    Rrpr.Append(runfonts);
                                                }

                                                if (Rrpr.GetFirstChild<FontSize>() != null)
                                                {
                                                    if (Rrpr.GetFirstChild<FontSize>().Val != null)
                                                    {
                                                        if (Rrpr.GetFirstChild<FontSize>().Val != "32")
                                                        {
                                                            flag4 = false;
                                                            Rrpr.GetFirstChild<FontSize>().Val = "32";
                                                        }
                                                    }
                                                }
                                                else
                                                {

                                                    FontSize fontSize1 = new FontSize() { Val = "32" };
                                                    Rrpr.Append(fontSize1);
                                                }
                                                if (Rrpr.GetFirstChild<Bold>() == null)
                                                {
                                                    if (pPr != null)
                                                    {
                                                        if (pPr.GetFirstChild<ParagraphStyleId>() == null)
                                                        {
                                                            flag5 = false;
                                                        }
                                                    }

                                                    Bold bold = new Bold() { Val = true };
                                                    Rrpr.Append(bold);
                                                }
                                                else
                                                {
                                                    if (Rrpr.GetFirstChild<Bold>().Val == null || Rrpr.GetFirstChild<Bold>().Val != true)
                                                    {

                                                        Rrpr.GetFirstChild<Bold>().Val = true;
                                                    }
                                                }


                                            }
                                            else
                                            {
                                                RunProperties runp = new RunProperties();
                                                RunFonts runFonts1 = new RunFonts()
                                                    {
                                                        Ascii = "Times New Roman",
                                                        HighAnsi = "Times New Roman",
                                                        ComplexScript = "Times New Roman",
                                                        EastAsia = "Times New Roman"
                                                    };
                                                FontSize fontSize1 = new FontSize() { Val = "32" };
                                                Bold bold1 = new Bold() { Val = true };

                                                runp.Append(runFonts1);
                                                runp.Append(fontSize1);
                                                runp.Append(bold1);
                                                pr.Append(runp);
                                            }

                                        }
                                    }

                                }
                            }
                        }
                    }

                    if (!flag1)
                    {
                        XmlElement error = xml.CreateElement("Text");
                        error.InnerText = "封面英文标题实词首字母应大写";
                        errInfor.AppendChild(error);
                    }
                    if (!flag2)
                    {
                        XmlElement error = xml.CreateElement("Text");
                        error.InnerText = "封面英文标题虚词首字母应小写";
                        errInfor.AppendChild(error);
                    }
                    if (!flag3)
                    {
                        XmlElement error = xml.CreateElement("Text");
                        error.InnerText = "封面英文标题字体错误，应为Times New Roman";
                        errInfor.AppendChild(error);
                    }
                    if (!flag4)
                    {
                        XmlElement error = xml.CreateElement("Text");
                        error.InnerText = "封面英文标题字号错误，应为三号";
                        errInfor.AppendChild(error);
                    }
                    if (!flag4)
                    {
                        XmlElement error = xml.CreateElement("Text");
                        error.InnerText = "封面英文标题未加粗";
                        errInfor.AppendChild(error);
                    }



                    break;

                }
                xml.Save(xmlFullPath);
            }
        


        /*private void getStudentInfoXml(WordprocessingDocument doc)
        {
            MainDocumentPart mainPart = doc.MainDocumentPart;
            Body body = mainPart.Document.Body;
            IEnumerable<Paragraph> paras = body.Elements<Paragraph>();
            XmlDocument xml = new XmlDocument();
            /*xml.Load(xmlFullPath);
            XmlNode root = xml.SelectSingleNode("CoverStyle/CoverLogo");
            XmlNode sproot = xml.SelectSingleNode("CoverStyle/spErroInfo");*/
           /* int count = 0;
            bool flag = false;
            foreach (Paragraph p in paras)
            {
                if (p.InnerText.Trim().Length > 0)
                    count++;
                if (count > 20)
                    break;
                string sentence = Regex.Replace(p.InnerText, @"\s*", "");
                if (sentence.IndexOf("完成日期") > 0)
                    break;
                Run r = p.GetFirstChild<Run>();
                if (r == null)
                    continue;

                int tempa = Tool.getFullText(p).Length;
                if (tempa > 7 && flag == true)
                {
                    if (p.InnerText.IndexOf("学科") != -1 || p.InnerText.IndexOf("工") != -1)
                    {
                        if (masterType == 0)
                        {
                            if (Tool.getFullText(p).Substring(0, 6) != "学科、 专业")
                            {

                                if (p.Elements<Run>().Count() == 1)
                                {
                                    p.GetFirstChild<Run>().GetFirstChild<Text>().Text = "学科、 专业：";
                                }
                                else
                                {
                                    IEnumerable<Run> runs = p.Elements<Run>();
                                    int num = 0;
                                    foreach (Run rr in runs)
                                    {

                                        num++;
                                        if (num != 1 && rr.GetFirstChild<RunProperties>().GetFirstChild<Underline>() == null)
                                        {

                                            rr.GetFirstChild<Text>().Text = null;
                                        }
                                    }

                                    p.GetFirstChild<Run>().GetFirstChild<Text>().Text = "学科、 专业：";
                                }

                            }
                        }
                        if (masterType == 1)
                        {
                            if (p.InnerText.Substring(0, 7) != "工 程 领 域")
                            {
                                if (p.Elements<Run>().Count() == 1)
                                {
                                    p.GetFirstChild<Run>().GetFirstChild<Text>().Text = "工 程 领 域:";
                                }
                                else
                                {
                                    IEnumerable<Run> runs = p.Elements<Run>();
                                    int num = 0;
                                    foreach (Run rr in runs)
                                    {

                                        num++;
                                        if (num != 1 && rr.GetFirstChild<RunProperties>().GetFirstChild<Underline>() == null)
                                        {

                                            rr.GetFirstChild<Text>().Text = null;
                                        }
                                    }

                                    p.GetFirstChild<Run>().GetFirstChild<Text>().Text = "工 程 领 域:";
                                }
                            }
                        }
                        flag = false;
                    }
                }

                if (p.InnerText.IndexOf("作 者 姓 名") != -1)
                {
                    flag = true;

                }
            }
            // xml.Save(xmlFullPath);
        }*/



        private void getCNLogoXML(WordprocessingDocument doc,string xmlFullPath)
        {

            MainDocumentPart mainPart = doc.MainDocumentPart;
            Body body = mainPart.Document.Body;
            IEnumerable<Paragraph> paras = body.Elements<Paragraph>();
            //XmlDocument xml = new XmlDocument();
        XmlDocument xml = new XmlDocument();
            xml.Load(xmlFullPath);
            XmlNode errInfor = xml.SelectSingleNode("CoverStyle/spErroInfo");
            bool flag1 = true;
            bool flag2 = true;
            bool flag3 = true;
            //bool flag4 = true;
          
            int count = 0;

            bool flagen = false;
            string fulltext = "";
            foreach (Paragraph p in paras)
            {
                Run r = p.GetFirstChild<Run>();
                if (r == null)
                    continue;
                fulltext = Tool.getFullText(p).Trim();
                if (fulltext != "")
                    count++;
                if (count > 0 && fulltext == "大连理工大学")
                {
                    flagen = true;
                    ParagraphProperties pPr = p.GetFirstChild<ParagraphProperties>();
                    if (pPr != null)
                    {
                        /*if (pPr.GetFirstChild<Indentation>() != null &&
                            pPr.GetFirstChild<Indentation>() != null)
                        {
                            if (pPr.GetFirstChild<Indentation>().FirstLine != null &&
                                pPr.GetFirstChild<Indentation>().FirstLineChars != null)
                            {
                                if (pPr.GetFirstChild<Indentation>().FirstLine == "3420" &&
                                    pPr.GetFirstChild<Indentation>().FirstLineChars == "950")
                                { }
                            }
                        }
                        else if (pPr.GetFirstChild<Justification>() != null)
                        {
                            if (pPr.GetFirstChild<Justification>().Val.ToString().ToLower() != "center")
                            {
                                pPr.GetFirstChild<Justification>().Val = JustificationValues.Center;

                            }
                        }
                        else
                        {
                            Justification justification = new Justification() { Val = JustificationValues.Center };
                            pPr.Append(justification);
                        }*/

                        if (pPr.GetFirstChild<Justification>() == null)
                        {
                            if (pPr.GetFirstChild<Indentation>() != null)
                            {
                                //if (pPr.GetFirstChild<Indentation>().FirstLine != null &&
                                //pPr.GetFirstChild<Indentation>().FirstLineChars != null)
                                //{
                                if (pPr.GetFirstChild<Indentation>().FirstLine == "3420" &&
                                    pPr.GetFirstChild<Indentation>().FirstLineChars == "950")
                                { }

                                else if (pPr.GetFirstChild<Indentation>().FirstLine == "3240" &&
                                    pPr.GetFirstChild<Indentation>().FirstLineChars == "900")
                                { }
                                // }
                            }

                            else
                            {
                                Justification justification = new Justification() { Val = JustificationValues.Center };
                                pPr.Append(justification);
                            }
                        }
                        else
                        {
                            if (pPr.GetFirstChild<Justification>().Val.ToString().ToLower() != "center")
                            {
                                XmlElement error = xml.CreateElement("Text");
                                error.InnerText = "封面中文logo未居中";
                                errInfor.AppendChild(error);
                                pPr.GetFirstChild<Justification>().Val = JustificationValues.Center;

                            }
                        }

                    }


                    IEnumerable<Run> runss = p.Elements<Run>();



                    foreach (Run rr in runss)
                    {
                        if (rr.GetFirstChild<RunProperties>() != null)
                        {
                            if (rr.GetFirstChild<RunProperties>().GetFirstChild<RunFonts>() != null)
                            {
                                if (rr.GetFirstChild<RunProperties>().GetFirstChild<RunFonts>().Ascii != null)
                                {
                                    if (rr.GetFirstChild<RunProperties>().GetFirstChild<RunFonts>().Ascii != "华文行楷")
                                        flag3 = false;
                                }
                              
                                
                                rr.GetFirstChild<RunProperties>().GetFirstChild<RunFonts>().Ascii = "华文行楷";
                                rr.GetFirstChild<RunProperties>().GetFirstChild<RunFonts>().HighAnsi = "华文行楷";
                                rr.GetFirstChild<RunProperties>().GetFirstChild<RunFonts>().ComplexScript = "华文行楷";
                                rr.GetFirstChild<RunProperties>().GetFirstChild<RunFonts>().EastAsia = "华文行楷";
                            }

                            else
                            {
                                RunFonts runfonts = new RunFonts()
                                {
                                    Ascii = "华文行楷",
                                    HighAnsi = "华文行楷",
                                    ComplexScript = "华文行楷",
                                    EastAsia = "华文行楷"
                                };
                                rr.GetFirstChild<RunProperties>().Append(runfonts);
                            }
                        }
                    }


                    foreach (Run rr in runss)
                    {
                        RunProperties Rrpr = rr.GetFirstChild<RunProperties>();
                        if (Rrpr != null)
                        {
                            if (Rrpr.GetFirstChild<FontSize>() != null)
                            {
                                if (Rrpr.GetFirstChild<FontSize>().Val != null)
                                {
                                    if (Rrpr.GetFirstChild<FontSize>().Val != "36")
                                    {
                                        flag1 = false;
                                        Rrpr.GetFirstChild<FontSize>().Val = "36";
                                    }
                                }
                            }
                            else
                            {

                                FontSize fontSize1 = new FontSize() { Val = "36" };
                                Rrpr.Append(fontSize1);
                            }
                        }
                    }


                    if (!flag3)
                    {
                        XmlElement error = xml.CreateElement("Text");
                        error.InnerText = "封面中文logo字体错误，应为华文行楷";
                        errInfor.AppendChild(error);
                    }
                    if (!flag1)
                    {
                        XmlElement error = xml.CreateElement("Text");
                        error.InnerText = "封面中文logo字号错误，应为小二号";
                        errInfor.AppendChild(error);
                    }
                
                     xml.Save(xmlFullPath);
                    continue;
                }

                if (flagen && fulltext != "")
                {
                    //Console.WriteLine("pppp");
                    ParagraphProperties ppr = p.GetFirstChild<ParagraphProperties>();
                    if (ppr != null)
                    {
                        /*if (ppr.GetFirstChild<Indentation>() != null &&
                           ppr.GetFirstChild<Indentation>() != null)
                        {

                            if (ppr.GetFirstChild<Indentation>().FirstLine != null &&
                               ppr.GetFirstChild<Indentation>().FirstLineChars != null)
                            {
                                if (ppr.GetFirstChild<Indentation>().FirstLine == "2880" &&
                                   ppr.GetFirstChild<Indentation>().FirstLineChars == "1200")
                                { }
                            }


                        }
                        else if (ppr.GetFirstChild<Justification>() != null)
                        {

                            if (ppr.GetFirstChild<Justification>().Val.ToString().ToLower() != "center")
                            {

                                ppr.GetFirstChild<Justification>().Val = JustificationValues.Center;
                            }
                        }
                        else
                        {

                            Justification justification = new Justification() { Val = JustificationValues.Center };
                            ppr.Append(justification);
                        }*/

                        if (ppr.GetFirstChild<Justification>() == null)
                        {
                            if (ppr.GetFirstChild<Indentation>() != null)
                            {
                                //if (pPr.GetFirstChild<Indentation>().FirstLine != null &&
                                //pPr.GetFirstChild<Indentation>().FirstLineChars != null)
                                //{
                                if (ppr.GetFirstChild<Indentation>().FirstLine == "2880" &&
                                    ppr.GetFirstChild<Indentation>().FirstLineChars == "1200")
                                { }

                                else if (ppr.GetFirstChild<Indentation>().FirstLine == "2760" &&
                                    ppr.GetFirstChild<Indentation>().FirstLineChars == "1150")
                                { }
                                // }
                            }

                            else
                            {
                                Justification justification = new Justification() { Val = JustificationValues.Center };
                                ppr.Append(justification);
                            }
                        }
                        else
                        {
                            if (ppr.GetFirstChild<Justification>().Val.ToString().ToLower() != "center")
                            {
                                XmlElement error = xml.CreateElement("Text");
                                error.InnerText = "封面英文logo未居中";
                                errInfor.AppendChild(error);
                                ppr.GetFirstChild<Justification>().Val = JustificationValues.Center;

                            }
                        }
                    }

                    IEnumerable<Run> runs = p.Elements<Run>();
                    if (!Tool.correctfonts(p, doc, "Times New Roman", "Times New Roman"))
                    {
                        XmlElement error = xml.CreateElement("Text");
                        error.InnerText = "封面英文logo字体错误，应为Times New Roman";
                        errInfor.AppendChild(error);

                        foreach (Run rr in runs)
                        {
                            if (rr.GetFirstChild<RunProperties>() != null)
                            {
                                if (rr.GetFirstChild<RunProperties>().GetFirstChild<RunFonts>() != null)
                                {
                                    rr.GetFirstChild<RunProperties>().GetFirstChild<RunFonts>().Ascii = "Times New Roman";
                                    rr.GetFirstChild<RunProperties>().GetFirstChild<RunFonts>().HighAnsi = "Times New Roman";
                                    rr.GetFirstChild<RunProperties>().GetFirstChild<RunFonts>().ComplexScript = "Times New Roman";
                                    rr.GetFirstChild<RunProperties>().GetFirstChild<RunFonts>().EastAsia = "Times New Roman";
                                }

                                else
                                {
                                    RunFonts runfonts = new RunFonts()
                                    {
                                        Ascii = "Times New Roman",
                                        HighAnsi = "Times New Roman",
                                        ComplexScript = "Times New Roman",
                                        EastAsia = "Times New Roman"
                                    };
                                    rr.GetFirstChild<RunProperties>().Append(runfonts);
                                }
                            }

                        }
                    }

                    foreach (Run rr in runs)
                    {
                        RunProperties Rrpr = rr.GetFirstChild<RunProperties>();
                        if (Rrpr != null)
                        {
                            if (Rrpr.GetFirstChild<FontSize>() != null)
                            {
                                if (Rrpr.GetFirstChild<FontSize>().Val != null)
                                {
                                    if (Rrpr.GetFirstChild<FontSize>().Val != "24")
                                    {
                                        flag2 = false;
                                        Rrpr.GetFirstChild<FontSize>().Val = "24";
                                    }
                                }
                            }
                            else
                            {

                                FontSize fontSize1 = new FontSize() { Val = "24" };
                                Rrpr.Append(fontSize1);
                            }
                        }
                    }


                    if (!flag2)
                    {
                        XmlElement error = xml.CreateElement("Text");
                        error.InnerText = "封面英文logo字号错误，应为小四号";
                        errInfor.AppendChild(error);
                    }
                      xml.Save(xmlFullPath);
                    break;
                }
            }
        }




        private void Second_Page(WordprocessingDocument doc, String xmlFullPath)
        {
            Body body = doc.MainDocumentPart.Document.Body;
            MainDocumentPart mainPart = doc.MainDocumentPart;
            IEnumerable<Paragraph> paras = body.Elements<Paragraph>();
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(xmlFullPath);
            XmlNode root = xmlDoc.SelectSingleNode("CoverStyle/Second");
            XmlNode spRoot = xmlDoc.SelectSingleNode("CoverStyle/spErroInfo");
            int count = -1;
            bool find_declare = false;
            List<Paragraph> pList = toList(paras);
            foreach (Paragraph p in paras)
            {
                Run r = p.GetFirstChild<Run>();
                String fullText = "";
                count++;
                if (r != null || p.GetFirstChild<Hyperlink>() != null)
                {
                    fullText = Tool.getFullText(p).Trim();
                }
                if (fullText != "" && count <= 40)
                {

                    if (fullText == "大连理工大学学位论文独创性声明")
                    {
                        find_declare = true;
                        Paragraph before_declare = pList[count - 1];
                        Paragraph after_declare = pList[count + 1];
                        string bd = getFullText(before_declare);
                        string ad = getFullText(after_declare);
                        string bd2 = getFullText(pList[count - 2]);
                        string ad2 = getFullText(pList[count + 2]);
                        if (bd != "")
                        {
                            XmlElement xe1 = xmlDoc.CreateElement("Text");
                            xe1.InnerText = "独创性声明页首应空一行";
                            spRoot.AppendChild(xe1);
                        }
                        if (ad != "")
                        {
                            XmlElement xe1 = xmlDoc.CreateElement("Text");
                            xe1.InnerText = "独创性声明题目与正文间应空一行";
                            spRoot.AppendChild(xe1);
                        }
                        if (bd != "" && bd2 != "")
                        {
                            XmlElement xe1 = xmlDoc.CreateElement("Text");
                            xe1.InnerText = "独创性声明页首应只有一个空行";
                            spRoot.AppendChild(xe1);
                        }
                        if (ad != "" && ad2 != "")
                        {
                            XmlElement xe1 = xmlDoc.CreateElement("Text");
                            xe1.InnerText = "独创性声明题目与正文间应只有一个空行";
                            spRoot.AppendChild(xe1);
                        }
                    }
                }
                if (count > 40 && find_declare == false)
                {
                    XmlElement xe1 = xmlDoc.CreateElement("Text");
                    xe1.InnerText = "论文的第二页应为大连理工大学学位论文独创性声明，疑似缺失";
                    spRoot.AppendChild(xe1);
                }
            }

            xmlDoc.Save(xmlFullPath);

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

        //判断封面字体是否正确

        public static bool correctCoverfonts(Paragraph p, WordprocessingDocument doc)
        {
            IEnumerable<Style> style = doc.MainDocumentPart.StyleDefinitionsPart.Styles.Elements<Style>();
            //段落style id

            //正文style
            Style Normalst = null;
            foreach (Style s in style)
            {
                if (s.StyleName.Val == "Normal")
                {
                    Normalst = s;
                    break;
                }
            }
            //正文样式字体 
            string Normalfonts = null;
            if (Normalst != null)
            {
                if (Normalst.StyleRunProperties != null)
                {
                    if (Normalst.StyleRunProperties.RunFonts != null)
                    {
                        if (Normalst.StyleRunProperties.RunFonts.Ascii != null)
                        {
                            Normalfonts = Normalst.StyleRunProperties.RunFonts.Ascii.ToString();
                        }
                    }
                }
            }
            IEnumerable<Run> run = p.Elements<Run>();
            foreach (Run r in run)
            {
                if (r.InnerText != null)
                {
                    //rstyleid
                    string rstyleid = null;
                    //rBaseonstyleid
                    string rBasestyleid = null;
                    //pstyleid
                    string pstyleid = "";
                    //段落style
                    Style pstyle = null;
                    //Baseonstyle
                    Style Basestyle = null;
                    //rfonts
                    string rfonts = null;
                    //pfonts
                    string Pfonts = null;
                    //Pbasefonts
                    string Basefonts = null;
                    //rBaseonfonts
                    string rBasefonts = null;
                    if (r.RunProperties != null)
                    {
                        if (r.RunProperties.RunStyle != null)
                        {
                            if (r.RunProperties.RunStyle.Val != null)
                            {
                                rstyleid = r.RunProperties.RunStyle.Val.ToString();
                            }
                        }
                    }
                    //rstyle
                    Style rstyle = null;
                    //rBaseonstyle
                    Style rBasestyle = null;
                    if (rstyleid != null)
                    {
                        foreach (Style s in style)
                        {
                            if (s.StyleId == rstyleid)
                            {
                                rstyle = s;
                                break;
                            }
                        }
                    }

                    if (rstyle != null)
                    {
                        if (rstyle.StyleRunProperties != null)
                        {
                            if (rstyle.StyleRunProperties.RunFonts != null)
                            {
                                if (rstyle.StyleRunProperties.RunFonts.Ascii != null)
                                {
                                    rfonts = rstyle.StyleRunProperties.RunFonts.Ascii.ToString();
                                }
                            }
                        }
                        else if (rstyle.BasedOn != null)
                        {
                            if (rstyle.BasedOn.Val != null)
                            {
                                rBasestyleid = rstyle.BasedOn.Val;
                            }
                            if (rBasestyleid != null)
                            {
                                foreach (Style s in style)
                                {
                                    if (s.StyleId == rBasestyleid)
                                    {
                                        rBasestyle = s;
                                    }
                                }
                            }
                            if (rBasestyle != null)
                            {
                                if (rBasestyle.StyleRunProperties != null)
                                {
                                    if (rBasestyle.StyleRunProperties.RunFonts != null)
                                    {
                                        if (rBasestyle.StyleRunProperties.RunFonts.Ascii != null)
                                        {
                                            rBasefonts = rBasestyle.StyleRunProperties.RunFonts.Ascii.ToString();
                                        }
                                    }
                                }
                            }
                        }
                    }
                    if (p.GetFirstChild<ParagraphProperties>() != null)
                    {
                        if (p.GetFirstChild<ParagraphProperties>().ParagraphStyleId != null)
                        {
                            if (p.GetFirstChild<ParagraphProperties>().ParagraphStyleId.Val != null)
                            {
                                pstyleid = p.GetFirstChild<ParagraphProperties>().ParagraphStyleId.Val;
                            }
                        }
                    }

                    if (pstyleid != null)
                    {
                        foreach (Style s in style)
                        {
                            if (s.StyleId == pstyleid)
                            {
                                pstyle = s;
                                break;
                            }
                        }
                    }

                    if (pstyle != null)
                    {
                        if (pstyle.StyleRunProperties != null)
                        {
                            if (pstyle.StyleRunProperties.RunFonts != null)
                            {
                                if (pstyle.StyleRunProperties.RunFonts.Ascii != null)
                                {
                                    if (pstyle.StyleRunProperties.RunFonts.Ascii.ToString() != null)
                                    {
                                        Pfonts = pstyle.StyleRunProperties.RunFonts.Ascii.ToString();
                                    }
                                }
                            }
                        }
                        else if (pstyle.BasedOn != null)
                        {
                            string Basestyleid = null;//Baseonstyleid
                            if (pstyle.BasedOn.Val != null)
                            {
                                Basestyleid = pstyle.BasedOn.Val;
                            }

                            if (Basestyleid != null)
                            {
                                foreach (Style s in style)
                                {
                                    if (s.StyleId == Basestyleid)
                                    {
                                        Basestyle = s;
                                        break;
                                    }
                                }
                            }
                            if (Basestyle != null)
                            {
                                if (Basestyle.StyleRunProperties != null)
                                {
                                    if (Basestyle.StyleRunProperties.RunFonts != null)
                                    {
                                        if (Basestyle.StyleRunProperties.RunFonts.Ascii != null)
                                        {
                                            if (Basestyle.StyleRunProperties.RunFonts.Ascii.ToString() != null)
                                            {
                                                Basefonts = Basestyle.StyleRunProperties.RunFonts.Ascii.ToString();
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }



                    RunProperties rpr = r.RunProperties;
                    if (rpr != null)
                    {
                        if (rpr.RunFonts != null)
                        {
                            if (rpr.RunFonts.Ascii != null)
                            {
                                if (rpr.RunFonts.Ascii == "华文行楷" || rpr.RunFonts.Ascii == "Times New Roman")
                                {
                                }
                                else
                                {
                                    if (rpr.RunFonts.Ascii.ToString()[0] >= 'A' && rpr.RunFonts.Ascii.ToString()[0] <= 'Z')
                                    {
                                        if ((int)r.GetFirstChild<Text>().InnerText[0] > 127)//是汉字，就正确
                                        {

                                        }
                                        else
                                        {
                                            return false;
                                        }
                                    }
                                    else
                                    {
                                        return false;
                                    }
                                }
                            }
                            else if (rfonts != null)
                            {
                                if (rfonts == "华文行楷" || rfonts == "Times New Roman")
                                {
                                }
                                else
                                {
                                    if (rfonts[0] >= 'A' && rfonts[0] <= 'Z')
                                    {
                                        if ((int)r.GetFirstChild<Text>().InnerText[0] > 127)//是汉字，就正确
                                        {

                                        }
                                        else
                                        {
                                            return false;
                                        }
                                    }
                                    else
                                    {
                                        return false;
                                    }
                                }
                            }
                            else if (rBasefonts != null)
                            {
                                if (rBasefonts == "华文行楷" || rBasefonts == "Times New Roman")
                                {
                                }
                                else
                                {
                                    if (rBasefonts[0] >= 'A' && rBasefonts[0] <= 'Z')
                                    {
                                        if ((int)r.GetFirstChild<Text>().InnerText[0] > 127)//是汉字，就正确
                                        {

                                        }
                                        else
                                        {
                                            return false;
                                        }
                                    }
                                    else
                                    {
                                        return false;
                                    }
                                }
                            }
                            else if (Pfonts != null)
                            {
                                if (Pfonts == "华文行楷" || Pfonts == "Times New Roman")
                                {
                                }
                                else
                                {
                                    if (Pfonts[0] >= 'A' && Pfonts[0] <= 'Z')
                                    {
                                        if ((int)r.GetFirstChild<Text>().InnerText[0] > 127)//是汉字，就正确
                                        {

                                        }
                                        else
                                        {
                                            return false;
                                        }
                                    }
                                    else
                                    {
                                        return false;
                                    }
                                }
                            }
                            else if (Basefonts != null)
                            {
                                if (Basefonts == "华文行楷" || Basefonts == "Times New Roman")
                                {
                                }
                                else
                                {
                                    if (Basefonts[0] >= 'A' && Basefonts[0] <= 'Z')
                                    {
                                        if ((int)r.GetFirstChild<Text>().InnerText[0] > 127)//是汉字，就正确
                                        {

                                        }
                                        else
                                        {
                                            return false;
                                        }
                                    }
                                    else
                                    {
                                        return false;
                                    }
                                }
                            }
                            else if (Normalfonts != null)
                            {
                                if (Normalfonts == "华文行楷" || Normalfonts == "Times New Roman")
                                {
                                }
                                else
                                {
                                    if (Normalfonts[0] >= 'A' && Normalfonts[0] <= 'Z')
                                    {
                                        if ((int)r.GetFirstChild<Text>().InnerText[0] > 127)//是汉字，就正确
                                        {

                                        }
                                        else
                                        {
                                            return false;
                                        }
                                    }
                                    else
                                    {
                                        return false;
                                    }
                                }
                            }
                        }
                        else
                        {
                            if (rfonts != null)
                            {
                                if (rfonts == "华文行楷" || rfonts == "Times New Roman")
                                {
                                }
                                else
                                {
                                    if (rfonts[0] >= 'A' && rfonts[0] <= 'Z')
                                    {
                                        if ((int)r.GetFirstChild<Text>().InnerText[0] > 127)//是汉字，就正确
                                        {

                                        }
                                        else
                                        {
                                            return false;
                                        }
                                    }
                                    else
                                    {
                                        return false;
                                    }
                                }
                            }
                            else if (rBasefonts != null)
                            {
                                if (rBasefonts == "华文行楷" || rBasefonts == "Times New Roman")
                                {
                                }
                                else
                                {
                                    if (rBasefonts[0] >= 'A' && rBasefonts[0] <= 'Z')
                                    {
                                        if ((int)r.GetFirstChild<Text>().InnerText[0] > 127)//是汉字，就正确
                                        {

                                        }
                                        else
                                        {
                                            return false;
                                        }
                                    }
                                    else
                                    {
                                        return false;
                                    }
                                }
                            }
                            else if (Pfonts != null)
                            {
                                if (Pfonts == "华文行楷" || Pfonts == "Times New Roman")
                                {
                                }
                                else
                                {
                                    if (Pfonts[0] >= 'A' && Pfonts[0] <= 'Z')
                                    {
                                        if ((int)r.GetFirstChild<Text>().InnerText[0] > 127)//是汉字，就正确
                                        {

                                        }
                                        else
                                        {
                                            return false;
                                        }
                                    }
                                    else
                                    {
                                        return false;
                                    }
                                }
                            }
                            else if (Basefonts != null)
                            {
                                if (Basefonts == "华文行楷" || Basefonts == "Times New Roman")
                                {
                                }
                                else
                                {
                                    if (Basefonts[0] >= 'A' && Basefonts[0] <= 'Z')
                                    {
                                        if ((int)r.GetFirstChild<Text>().InnerText[0] > 127)//是汉字，就正确
                                        {

                                        }
                                        else
                                        {
                                            return false;
                                        }
                                    }
                                    else
                                    {
                                        return false;
                                    }
                                }
                            }
                            else if (Normalfonts != null)
                            {
                                if (Normalfonts == "华文行楷" || Normalfonts == "Times New Roman")
                                {
                                }
                                else
                                {
                                    if (Normalfonts[0] >= 'A' && Normalfonts[0] <= 'Z')
                                    {
                                        if ((int)r.GetFirstChild<Text>().InnerText[0] > 127)//是汉字，就正确
                                        {

                                        }
                                        else
                                        {
                                            return false;
                                        }
                                    }
                                    else
                                    {
                                        return false;
                                    }
                                }
                            }
                        }
                    }
                }
            }
            return true;

        }


        private bool hasChinese(string str)
        {
            return Regex.IsMatch(str, @"[\u4e00-\u9fa5]");
        }

        public static void addComment(WordprocessingDocument document, Paragraph p, string comment)
        {
            Comments comments;
            int id = 1;
            if (document.MainDocumentPart.GetPartsCountOfType<WordprocessingCommentsPart>() > 0)
            {
                comments =
                    document.MainDocumentPart.WordprocessingCommentsPart.Comments;
                if (comments.HasChildren)
                {
                    if (comments.Descendants<Comment>().Select(e => Convert.ToInt32(e.Id.Value)).Max<int>() == 10)
                        id = comments.Descendants<Comment>().Select(e => Convert.ToInt32(e.Id.Value)).Max<int>();
                    else
                        id = comments.Descendants<Comment>().Select(e => Convert.ToInt32(e.Id.Value)).Max<int>();
                    id = id + 1;
                }
            }
            else
            {
                WordprocessingCommentsPart commentPart =
                            document.MainDocumentPart.AddNewPart<WordprocessingCommentsPart>();
                commentPart.Comments = new Comments();
                comments = commentPart.Comments;
            }
            Paragraph p1 = new Paragraph(new Run(new Text(comment)));
            Comment cmt = new Comment()
            {
                Id = id.ToString(),
                Author = "Red Ant",
                Initials = "yxy",
                Date = DateTime.Now.AddHours(8)
            };
            cmt.AppendChild(p1);
            comments.AppendChild(cmt);
            comments.Save();
            /************4/16新加*/
            if (p.Elements<Run>().Count<Run>() == 0)
            {
                p.AppendChild<Run>(new Run(new Text("")));
            }
            /***************/
            p.InsertBefore(new CommentRangeStart() { Id = id.ToString() }, p.GetFirstChild<Run>());

            var cmtEnd = p.InsertAfter(new CommentRangeEnd() { Id = id.ToString() }, p.Elements<Run>().Last());

            p.InsertAfter(new Run(new CommentReference() { Id = id.ToString() }), cmtEnd);

        }

    }

}

