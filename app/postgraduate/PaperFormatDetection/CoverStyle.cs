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
    public class CoverStyle : ModuleFormat
    {
        String[] charList = { "a", "an", "the", "of", "in", "on", "from", "and", "for", "is", "are", "with" };
        string CNname = null;
        //构造函数
        public CoverStyle(List<Module> modList, PageLocator locator, int masterType)
            : base(modList, locator, masterType)
        {

        }
        public override void getStyle(WordprocessingDocument doc, String fileName)
        {
            string xmlFullPath = fileName + "\\CoverStyle.xml";//xml模板文件保存路径
            CreateXmlFile(xmlFullPath);
            getHeadingXML(doc, xmlFullPath);
            getCNTitleXML(doc, xmlFullPath);
            getENTitleXML(doc, xmlFullPath);
            getStudentInfoXml(doc, xmlFullPath);
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

        private void getHeadingXML(WordprocessingDocument doc, String xmlFullPath)
        {
            Body body = doc.MainDocumentPart.Document.Body;
            MainDocumentPart mainPart = doc.MainDocumentPart;
            IEnumerable<Paragraph> paras = body.Elements<Paragraph>();
            XmlDocument xml = new XmlDocument();
            xml.Load(xmlFullPath);
            XmlNode errInfor = xml.SelectSingleNode("CoverStyle/spErroInfo");
            int count = 0;

            foreach (Paragraph p in paras)
            {
                Run r = p.GetFirstChild<Run>();
                if (r == null)
                    continue;
                String fullText = Tool.getFullText(p).Trim();
                if (fullText != "")
                    count++;
                if (count == 1)
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
                            }
                        }
                    }

                    string sentence2 = Regex.Replace(fullText, @"\s*", "");

                    if (masterType == 0)
                    {
                        if (sentence2 != "硕士学位论文")
                        {
                            XmlElement error = xml.CreateElement("Text");
                            error.InnerText = "封面大标题错误，应该为 硕士学位论文";
                            errInfor.AppendChild(error);
                        }
                    }
                    else if (masterType == 1)
                    {
                        if (sentence2 != "专业学位硕士学位论文")
                        {
                            XmlElement error = xml.CreateElement("Text");
                            error.InnerText = "封面大标题错误，应该为 专业学位硕士学位论文";
                            errInfor.AppendChild(error);

                        }
                    }
                    if (rs != null)
                    {
                        bool flag1 = true;
                        bool flag2 = true;
                        bool flag3 = true;



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
                                                    flag1 = false;
                                                }
                                            }
                                        }
                                    }

                                    if (rpr.GetFirstChild<FontSize>() != null)
                                    {
                                        if (rpr.GetFirstChild<FontSize>().Val != null)
                                        {
                                            if (rpr.GetFirstChild<FontSize>().Val != "48")
                                            {
                                                flag2 = false;
                                            }
                                        }
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
                                    }
                                }
                            }
                        }

                        if (!flag1)
                        {
                            XmlElement error = xml.CreateElement("Text");
                            error.InnerText = "封面大标题字体错误，应为宋体";
                            errInfor.AppendChild(error);
                        }

                        if (!flag2)
                        {
                            XmlElement error = xml.CreateElement("Text");
                            error.InnerText = "封面大标题字号错误，应为小一号";
                            errInfor.AppendChild(error);
                        }

                        if (!flag3)
                        {
                            XmlElement error = xml.CreateElement("Text");
                            error.InnerText = "封面大标题未加粗";
                            errInfor.AppendChild(error);
                        }
                    }

                    xml.Save(xmlFullPath);
                    break;
                }
            }


        }





        private void getCNTitleXML(WordprocessingDocument doc, String xmlFullPath)
        {
            MainDocumentPart mainPart = doc.MainDocumentPart;
            Body body = mainPart.Document.Body;
            IEnumerable<Paragraph> paras = body.Elements<Paragraph>();
            ParagraphProperties pPr = null;
            XmlDocument xml = new XmlDocument();
            xml.Load(xmlFullPath);
            XmlNode root = xml.SelectSingleNode("CoverStyle/Title");
            XmlNode sproot = xml.SelectSingleNode("CoverStyle/spErroInfo");
            string fulltext = "";
            int count = 0;

            foreach (Paragraph p in paras)
            {
                Run r = p.GetFirstChild<Run>();

                if (r != null || p.GetFirstChild<Hyperlink>() != null)
                {
                    fulltext = Tool.getFullText(p).Trim();
                }
                if (fulltext != "")
                {
                    count++;
                }
                if (count == 2)
                {
                    if (Tool.GetHanNumFromString(fulltext) > 20)
                    {
                        XmlElement xe = xml.CreateElement("Text");
                        xe.InnerText = "中文标题应少于20字";
                        sproot.AppendChild(xe);
                    }

                    pPr = p.GetFirstChild<ParagraphProperties>();

                    if (pPr != null)
                    {
                        if (pPr.GetFirstChild<Justification>() != null)
                        {
                            if (pPr.GetFirstChild<Justification>().Val.ToString() != "center")
                            {
                                XmlElement xe = xml.CreateElement("Text");

                                xe.InnerText = "中文标题应居中";

                                sproot.AppendChild(xe);
                            }
                        }

                        if (pPr.GetFirstChild<SpacingBetweenLines>() != null)
                        {
                            if (pPr.GetFirstChild<SpacingBetweenLines>().Line != null)
                            {
                                if (pPr.GetFirstChild<SpacingBetweenLines>().Line.Value != "300")
                                {
                                    XmlElement xe = xml.CreateElement("Text");

                                    xe.InnerText = "中文标题行距应为1.25";

                                    sproot.AppendChild(xe);
                                }
                            }

                            if (pPr.GetFirstChild<SpacingBetweenLines>().Before != null)
                            {
                                if (pPr.GetFirstChild<SpacingBetweenLines>().Before.Value != "0")
                                {
                                    XmlElement xe = xml.CreateElement("Text");

                                    xe.InnerText = "中文标题段前间距应为0行";

                                    sproot.AppendChild(xe);
                                }
                            }

                            if (pPr.GetFirstChild<SpacingBetweenLines>().After != null)
                            {
                                if (pPr.GetFirstChild<SpacingBetweenLines>().After.Value != "0")
                                {
                                    XmlElement xe = xml.CreateElement("Text");

                                    xe.InnerText = "中文标题段后间距应为0行";

                                    sproot.AppendChild(xe);
                                }
                            }
                        }

                        IEnumerable<Run> runs = p.Elements<Run>();
                        if (runs != null)
                        {
                            bool flag1 = true;
                            bool flag2 = true;
                            bool flag3 = true;
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
                                                }
                                            }
                                        }

                                        if (rPr.GetFirstChild<FontSize>() != null)
                                        {
                                            if (rPr.GetFirstChild<FontSize>().Val != null)
                                            {
                                                if (rPr.GetFirstChild<FontSize>().Val != "44")
                                                    flag2 = false;
                                            }
                                        }

                                        if (rPr.GetFirstChild<Bold>() == null)
                                        {
                                            if (pPr != null)
                                            {
                                                if (pPr.GetFirstChild<ParagraphStyleId>() == null)
                                                    flag3 = false;
                                            }
                                        }
                                    }
                                }
                            }

                            if (!flag1)
                            {
                                XmlElement xe1 = xml.CreateElement("Text");
                                xe1.InnerText = "封面中文标题字体错误，应为华文细黑";
                                sproot.AppendChild(xe1);
                            }
                            if (!flag2)
                            {
                                XmlElement xe1 = xml.CreateElement("Text");
                                xe1.InnerText = "封面中文标题字体字号错误，应为二号";
                                sproot.AppendChild(xe1);
                            }
                            if (!flag3)
                            {
                                XmlElement xe1 = xml.CreateElement("Text");
                                xe1.InnerText = "封面中文标题未加粗";
                                sproot.AppendChild(xe1);
                            }

                        }
                    }




                    break;
                }
            }

            xml.Save(xmlFullPath);
        }









        private void getENTitleXML(WordprocessingDocument doc, String xmlFullPath)
        {
            Body body = doc.MainDocumentPart.Document.Body;
            MainDocumentPart mainPart = doc.MainDocumentPart;
            IEnumerable<Paragraph> paras = body.Elements<Paragraph>();
            ParagraphProperties pPr = null;
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(xmlFullPath);
            XmlNode root = xmlDoc.SelectSingleNode("CoverStyle/Title");
            XmlNode spRoot = xmlDoc.SelectSingleNode("CoverStyle/spErroInfo");
            int count = 0;
            foreach (Paragraph p in paras)
            {
                Run r = p.GetFirstChild<Run>();
                String fullText = "";
                if (r != null || p.GetFirstChild<Hyperlink>() != null)
                {
                    fullText = Tool.getFullText(p).Trim();
                }
                /*if (fullText != "")
                {
                    count++;
                }*/
                if (fullText != "")
                {
                    if (!hasChinese(fullText))
                    {
                        String[] strList = fullText.Trim().Split(' ');
                        XmlElement xesub = xmlDoc.CreateElement("ENTitle");
                        xesub.SetAttribute("name", "论文英文标题");
                        pPr = p.GetFirstChild<ParagraphProperties>();
                        IEnumerable<Run> pRunList = p.Elements<Run>();
                        for (int i = 0; i < strList.Length; i++)
                        {
                            bool strFalg = true;
                            for (int j = 0; j < charList.Length; j++)
                            {
                                if (strList[i] == charList[j])
                                {
                                    strFalg = false;
                                }
                            }
                            if (strFalg)
                            {
                                if (strList[i].Length > 1)
                                {
                                    if (!Regex.IsMatch(strList[i].Substring(0, 1), "[A-Z]"))
                                    {
                                        XmlElement xe1 = xmlDoc.CreateElement("Text");
                                        xe1.InnerText = "封面英文标题实词首字母未大写：" + "{" + strList[i] + "}";
                                        spRoot.AppendChild(xe1);
                                    }
                                }
                            }
                        }
                        if (pPr != null)
                        {
                            if (pPr.GetFirstChild<Justification>() != null)
                            {
                                if (pPr.GetFirstChild<Justification>().Val.ToString().ToLower() != "center")
                                {
                                    XmlElement xe1 = xmlDoc.CreateElement("Text");
                                    xe1.InnerText = "封面英文标题未居中";
                                    spRoot.AppendChild(xe1);
                                }

                            }
                            if (pPr.GetFirstChild<SpacingBetweenLines>() != null)
                            {
                                if (pPr.GetFirstChild<SpacingBetweenLines>().Line != null)
                                {
                                    if (pPr.GetFirstChild<SpacingBetweenLines>().Line.Value != "300")
                                    {
                                        XmlElement xe1 = xmlDoc.CreateElement("Text");
                                        xe1.InnerText = "封面英文标题行距错误，应为多倍行距1.25";
                                        spRoot.AppendChild(xe1);
                                    }
                                }
                                if (pPr.GetFirstChild<SpacingBetweenLines>().Before != null)
                                {
                                    if (pPr.GetFirstChild<SpacingBetweenLines>().Before.Value != "0")
                                    {
                                        XmlElement xe1 = xmlDoc.CreateElement("Text");
                                        xe1.InnerText = "封面英文标题段前间距错误，应为0行";
                                        spRoot.AppendChild(xe1);
                                    }
                                }
                                if (pPr.GetFirstChild<SpacingBetweenLines>().After != null)
                                {
                                    if (pPr.GetFirstChild<SpacingBetweenLines>().After.Value != "0")
                                    {
                                        XmlElement xe1 = xmlDoc.CreateElement("Text");
                                        xe1.InnerText = "封面英文标题段后间距错误，应为0行";
                                        spRoot.AppendChild(xe1);
                                    }
                                }
                            }
                        }
                        if (pRunList != null)
                        {
                            bool flag1 = true;
                            bool flag2 = true;
                            bool flag3 = true;
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
                                                    flag1 = false;
                                                }
                                            }
                                        }
                                        if (Rrpr.GetFirstChild<FontSize>() != null)
                                        {
                                            if (Rrpr.GetFirstChild<FontSize>().Val != null)
                                            {
                                                if (Rrpr.GetFirstChild<FontSize>().Val != "32")
                                                {
                                                    flag2 = false;
                                                }
                                            }
                                        }
                                        if (Rrpr.GetFirstChild<Bold>() == null)
                                        {
                                            if (pPr != null)
                                            {
                                                if (pPr.GetFirstChild<ParagraphStyleId>() == null)
                                                {
                                                    flag3 = false;
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            if (!flag1)
                            {
                                XmlElement xe1 = xmlDoc.CreateElement("Text");
                                xe1.InnerText = "封面英文标题字体错误，应为“Times New Roman”";
                                spRoot.AppendChild(xe1);
                            }
                            if (!flag2)
                            {
                                XmlElement xe1 = xmlDoc.CreateElement("Text");
                                xe1.InnerText = "封面英文标题字号错误，应为三号";
                                spRoot.AppendChild(xe1);
                            }
                            if (!flag3)
                            {
                                XmlElement xe1 = xmlDoc.CreateElement("Text");
                                xe1.InnerText = "封面英文标题未加粗";
                                spRoot.AppendChild(xe1);
                            }
                        }
                        root.AppendChild(xesub);
                        xmlDoc.Save(xmlFullPath);
                        break;
                    }
                }
            }
        }

        private void getStudentInfoXml(WordprocessingDocument doc, String xmlFullPath)
        {
            MainDocumentPart mainPart = doc.MainDocumentPart;
            Body body = mainPart.Document.Body;
            IEnumerable<Paragraph> paras = body.Elements<Paragraph>();
            XmlDocument xml = new XmlDocument();
            xml.Load(xmlFullPath);
            XmlNode root = xml.SelectSingleNode("CoverStyle/CoverLogo");
            XmlNode sproot = xml.SelectSingleNode("CoverStyle/spErroInfo");
            int count = 0;
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
                                XmlElement xe1 = xml.CreateElement("Text");
                                xe1.InnerText = "统招、单考硕士、高校教师在职申请硕士学位、同等学历硕士研究生信息的第二个标题应为“学科、 专业”";
                                sproot.AppendChild(xe1);
                            }
                        }
                        if (masterType == 1)
                        {
                            if (p.InnerText.Substring(0, 7) != "工 程 领 域")
                            {
                                XmlElement xe1 = xml.CreateElement("Text");
                                xe1.InnerText = "工程硕士、MBA、EMBA、MPA的研究生信息的第二个标题应为“工 程 领 域”";
                                sproot.AppendChild(xe1);
                            }
                        }
                        flag = false;
                    }
                }

                if (p.InnerText.IndexOf("作 者 姓 名") != -1)
                {
                    flag = true;
                    ParagraphProperties pPr = p.GetFirstChild<ParagraphProperties>();
                    RunProperties rPr = r.GetFirstChild<RunProperties>();
                    bool rfflag = false;
                    bool fzflag = false;
                    bool bflag = false;
                    if (rPr != null)
                    {
                        //字体
                        if (rPr.GetFirstChild<RunFonts>() != null)
                        {
                            if (rPr.GetFirstChild<RunFonts>().Ascii != null)
                            {
                                rfflag = true;
                                XmlElement xe1 = xml.CreateElement("Fonts");
                                xe1.InnerText = rPr.GetFirstChild<RunFonts>().Ascii;
                                root.AppendChild(xe1);

                            }
                        }
                        //字号
                        if (rPr.GetFirstChild<FontSize>() != null)
                        {
                            if (rPr.GetFirstChild<FontSize>().Val != null)
                            {
                                fzflag = true;
                                XmlElement xe1 = xml.CreateElement("size");
                                xe1.InnerText = rPr.GetFirstChild<FontSize>().Val.Value;
                                root.AppendChild(xe1);
                            }
                        }
                        if (rPr.GetFirstChild<Bold>() != null)
                        {
                            bflag = true;
                            XmlElement xe1 = xml.CreateElement("bold");
                            xe1.InnerText = "true";
                            root.AppendChild(xe1);
                        }
                    }
                    if (pPr != null)
                    {
                        //对齐方式
                        if (pPr.GetFirstChild<Justification>() != null)
                        {
                            XmlElement xe1 = xml.CreateElement("jc");
                            xe1.InnerText = pPr.GetFirstChild<Justification>().Val.Value.ToString().ToLower();
                            root.AppendChild(xe1);
                        }
                    }
                    if (pPr.GetFirstChild<ParagraphStyleId>() != null)
                    {
                        StyleDefinitionsPart stylePart = mainPart.StyleDefinitionsPart;
                        Styles styles = stylePart.Styles;
                        var t = styles.ChildElements;
                        foreach (var s in t)
                        {
                            Style m = new Style();
                            if (s.GetType().Equals(m.GetType()))
                            {
                                m = (Style)s;
                                StyleRunProperties srPr = m.StyleRunProperties;
                                StyleParagraphProperties spPr = m.StyleParagraphProperties;
                                if (m.StyleId.ToString() == pPr.GetFirstChild<ParagraphStyleId>().Val)
                                {
                                    if (srPr != null)
                                    {
                                        if (srPr.RunFonts != null && !rfflag)
                                        {
                                            if (srPr.RunFonts.EastAsia != null)
                                            {
                                                XmlElement xe1 = xml.CreateElement("Fonts");
                                                xe1.InnerText = srPr.RunFonts.EastAsia;
                                                root.AppendChild(xe1);
                                            }
                                            else if (srPr.RunFonts.Ascii != null)
                                            {
                                                XmlElement xe1 = xml.CreateElement("Fonts");
                                                xe1.InnerText = srPr.RunFonts.Ascii;
                                                root.AppendChild(xe1);
                                            }
                                            else if (srPr.RunFonts.HighAnsi != null)
                                            {
                                                XmlElement xe1 = xml.CreateElement("Fonts");
                                                xe1.InnerText = srPr.RunFonts.HighAnsi;
                                                root.AppendChild(xe1);
                                            }
                                        }
                                        else if (!rfflag)
                                        {
                                            XmlElement xe1 = xml.CreateElement("Fonts");
                                            xe1.InnerText = "宋体";
                                            root.AppendChild(xe1);
                                        }
                                        if (srPr.FontSize != null && !fzflag)
                                        {
                                            XmlElement xe1 = xml.CreateElement("size");
                                            xe1.InnerText = srPr.FontSize.Val;
                                            root.AppendChild(xe1);
                                        }
                                        else if (!fzflag)
                                        {
                                            XmlElement xe1 = xml.CreateElement("size");
                                            xe1.InnerText = "30";
                                            root.AppendChild(xe1);
                                        }

                                        if (srPr.Bold != null && !bflag)
                                        {
                                            XmlElement xe1 = xml.CreateElement("bold");
                                            xe1.InnerText = "true";
                                            root.AppendChild(xe1);
                                        }
                                        else
                                        {
                                            XmlElement xe1 = xml.CreateElement("bold");
                                            xe1.InnerText = "false";
                                            root.AppendChild(xe1);
                                        }
                                    }
                                    if (spPr != null)
                                    {
                                        if (spPr.Justification != null)
                                        {
                                            XmlElement xe1 = xml.CreateElement("jc");
                                            xe1.InnerText = spPr.Justification.Val.ToString().ToLower();
                                            root.AppendChild(xe1);
                                        }
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        if (!rfflag)
                        {
                            XmlElement xe1 = xml.CreateElement("Fonts");
                            xe1.InnerText = "宋体";
                            root.AppendChild(xe1);
                        }
                        if (!fzflag)
                        {
                            XmlElement xe1 = xml.CreateElement("size");
                            xe1.InnerText = "30";
                            root.AppendChild(xe1);
                        }
                    }
                }
            }
            xml.Save(xmlFullPath);
        }



        private void getCNLogoXML(WordprocessingDocument doc, String xmlFullPath)
        {

            MainDocumentPart mainPart = doc.MainDocumentPart;
            Body body = mainPart.Document.Body;
            IEnumerable<Paragraph> paras = body.Elements<Paragraph>();
            XmlDocument xml = new XmlDocument();
            xml.Load(xmlFullPath);
            XmlNode root = xml.SelectSingleNode("CoverStyle/CoverLogo");
            XmlNode sproot = xml.SelectSingleNode("CoverStyle/spErroInfo");
            int count = 0;
            bool flaga = false;
            bool flagb = false;
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
                        if (pPr.GetFirstChild<Justification>() != null)
                        {
                            if (pPr.GetFirstChild<Justification>().Val.ToString().ToLower() != "center")
                            {
                                XmlElement xe = xml.CreateElement("text");
                                xe.InnerText = "底部中文logo未居中";
                                sproot.AppendChild(xe);
                            }
                        }
                    }

                    IEnumerable<Run> runs = p.Elements<Run>();
                    if (runs != null)
                    {
                        bool flag1 = true;
                        bool flag2 = true;

                        foreach (Run rr in runs)
                        {
                            if (rr != null)
                            {
                                RunProperties rPr = rr.GetFirstChild<RunProperties>();
                                if (rPr != null)
                                {
                                    if (rPr.GetFirstChild<RunFonts>() != null)
                                    {
                                        if (rPr.GetFirstChild<RunFonts>().Ascii != null)
                                        {
                                            if (rPr.GetFirstChild<RunFonts>().Ascii != "华文行楷" && rPr.GetFirstChild<RunFonts>().Ascii != "Times New Roman")
                                                flag1 = false;
                                        }

                                        else if (rPr.GetFirstChild<RunFonts>().Hint != null)
                                        {
                                            if (rPr.GetFirstChild<RunFonts>().Hint != "eastAsia")
                                                flag1 = false;
                                        }
                                    }

                                    if (rPr.GetFirstChild<FontSize>() != null)
                                    {
                                        if (rPr.GetFirstChild<FontSize>().Val != null)
                                        {
                                            if (rPr.GetFirstChild<FontSize>().Val != "36")
                                                flag2 = false;
                                        }
                                    }
                                }
                            }
                        }

                        if (!flag1)
                        {
                            flaga = true;
                            XmlElement xe = xml.CreateElement("Text");
                            xe.InnerText = "封面底部的中文校名字体错误，应为华文行楷";
                            sproot.AppendChild(xe);
                        }

                        if (!flag2)
                        {
                            flagb = true;
                            XmlElement xe = xml.CreateElement("Text");
                            xe.InnerText = "封面底部的中文校名字号错误，应为小二号";
                            sproot.AppendChild(xe);
                        }
                    }

                    String id = Tool.getPargraphStyleId(p);
                    if (id != "")
                    {
                        StyleDefinitionsPart stylePart = mainPart.StyleDefinitionsPart;
                        Styles styles = stylePart.Styles;
                        var t = styles.ChildElements;
                        foreach (var s in t)
                        {
                            Style m = new Style();
                            if (s.GetType().Equals(m.GetType()))
                            {
                                m = (Style)s;

                                if (m.StyleId.ToString() == id)
                                {
                                    StyleRunProperties srpr = m.StyleRunProperties;

                                    if (srpr != null)
                                    {
                                        if (srpr.RunFonts != null && !flaga)
                                        {
                                            if (srpr.RunFonts.Ascii != null)
                                            {
                                                if (srpr.RunFonts.Ascii != "华文行楷")
                                                {

                                                    XmlElement xe1 = xml.CreateElement("Text");
                                                    xe1.InnerText = "封面底部的中文校名字体错误，应为华文行楷";
                                                    sproot.AppendChild(xe1);
                                                }
                                            }
                                        }

                                        if (srpr.FontSize != null && !flagb)
                                        {
                                            if (srpr.FontSize.Val != null)
                                            {
                                                if (srpr.FontSize.Val != "36")
                                                {
                                                    XmlElement xe1 = xml.CreateElement("Text");
                                                    xe1.InnerText = "封面底部的中文校名字号错误，应为小二号";
                                                    sproot.AppendChild(xe1);
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }

                    xml.Save(xmlFullPath);
                    continue;
                }

                if (flagen)
                {
                    ParagraphProperties ppr = p.GetFirstChild<ParagraphProperties>();
                    if (ppr != null)
                    {
                        if (ppr.GetFirstChild<Justification>() != null)
                        {
                            if (ppr.GetFirstChild<Justification>().Val.ToString().ToLower() != "center")
                            {
                                XmlElement xe = xml.CreateElement("Text");
                                xe.InnerText = "封面英文校名未居中";
                                sproot.AppendChild(xe);
                            }
                        }
                    }

                    IEnumerable<Run> runs = p.Elements<Run>();
                    if (runs != null)
                    {
                        bool flag1 = true;
                        bool flag2 = true;
                        foreach (Run rr in runs)
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
                                            if (rpr.GetFirstChild<RunFonts>().Ascii != "Times New Roman")
                                                flag1 = false;
                                        }
                                    }

                                    if (Tool.correctsize(p, doc, "24") == false)
                                    {
                                        flag2 = false;
                                    }
                                }
                            }
                        }

                        if (!flag1)
                        {
                            XmlElement xe = xml.CreateElement("Text");
                            xe.InnerText = "封面英文校名字体错误，应为Times New Roman";
                            sproot.AppendChild(xe);

                        }

                        if (!flag2)
                        {
                            XmlElement xe = xml.CreateElement("Text");
                            xe.InnerText = "封面英文校名字号错误，应为小四号";
                            sproot.AppendChild(xe);
                        }
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
    }

}

