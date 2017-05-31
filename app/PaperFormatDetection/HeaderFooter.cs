using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using PaperFormatDetection.Tools;
using PaperFormatDetection.Frame;
using System.Text.RegularExpressions;

namespace PaperFormatDetection.Format
{
    public class HeaderFooter : ModuleFormat
    {
        //构造函数
        public HeaderFooter(List<Module> modList, PageLocator locator, int masterType)
            : base(modList, locator, masterType)
        {

        }
        public override void getStyle(WordprocessingDocument doc, String fileName)
        {
            string xmlFullPath = fileName + "\\HeaderFooter.xml";//xml模板文件保存路径
            CreateXmlFile(xmlFullPath);
            GetXml(doc, xmlFullPath);
        }
        private void CreateXmlFile(string xmlFullPath)
        {
            XmlDocument xmlDoc = new XmlDocument();
            XmlNode node = xmlDoc.CreateXmlDeclaration("1.0", "utf-8", "");
            xmlDoc.AppendChild(node);
            XmlNode root = xmlDoc.CreateElement("HeaderFooter");
            XmlElement xe1 = xmlDoc.CreateElement("HeaderStyle");
            xe1.SetAttribute("name", "页眉");
            XmlElement xe2 = xmlDoc.CreateElement("FooterStyle");
            xe2.SetAttribute("name", "页脚");
            XmlElement xe3 = xmlDoc.CreateElement("spErroInfo");
            xe3.SetAttribute("name", "特殊错误信息");
            XmlElement xe4 = xmlDoc.CreateElement("partName");
            xe4.SetAttribute("name", "提示名称");
            XmlElement xe5 = xmlDoc.CreateElement("Text");
            xe5.InnerText = "-----------------页眉页脚-----------------";
            xe4.AppendChild(xe5);
            root.AppendChild(xe4);
            root.AppendChild(xe1);
            root.AppendChild(xe2);
            root.AppendChild(xe3);
            xmlDoc.AppendChild(root);
            try
            {
                xmlDoc.Save(xmlFullPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
        private void GetXml(WordprocessingDocument docx, string xmlFullPath)
        {
            XmlDocument xmlDocx = new XmlDocument();
            xmlDocx.Load(xmlFullPath);
            XmlNode root = xmlDocx.SelectSingleNode("HeaderFooter/spErroInfo");
            Body body = docx.MainDocumentPart.Document.Body;
            IEnumerable<Paragraph> paras = body.Elements<Paragraph>();
            MainDocumentPart Mpart = docx.MainDocumentPart;
            List<SectionProperties> list = secPrList(body);         //页眉页脚组成链表
            List<int> intlist = secPrListInt(body);                  //位置

            //检测第一部分test
            firstSection(list, intlist, docx, xmlFullPath, root, xmlDocx);
            //其他部分
            //获取中文标题
            string name = "";                                       //中文题目
            int count = 0;
            foreach (Paragraph p in paras)
            {
                if (p.InnerText.Trim().Length > 0)
                {
                    count++;
                }
                if (count == 2)
                {
                    name = p.InnerText;
                    break;
                }
            }
            Console.WriteLine(name);
            otherSectionHeader(list, intlist, docx, xmlFullPath, root, xmlDocx, name);
            lastSectionHeader(docx, root, xmlDocx, name);
            xmlDocx.Save(xmlFullPath);
        }
        //获取所有章节属性SecPr的位置，用list保存
        static private List<int> secPrListInt(Body body)
        {
            List<int> list = new List<int>(20);
            int l = body.ChildElements.Count();
            for (int i = 0; i < l; i++)
            {
                if (body.ChildElements.GetItem(i).GetType().ToString() == "DocumentFormat.OpenXml.Wordprocessing.Paragraph")
                {
                    Paragraph p = (Paragraph)body.ChildElements.GetItem(i);
                    if (p.ParagraphProperties != null)
                    {
                        if (p.ParagraphProperties.SectionProperties != null)
                        {
                            list.Add(i);
                        }
                    }
                }
            }
            if (body.ChildElements.GetItem(l - 1).GetType().ToString() == "DocumentFormat.OpenXml.Wordprocessing.Paragraph")
            {
                if (((Paragraph)body.ChildElements.GetItem(l - 1)).ParagraphProperties != null)
                {
                    if (((Paragraph)body.ChildElements.GetItem(l - 1)).ParagraphProperties.SectionProperties != null)
                    {
                        list.Add(l - 1);
                    }
                }

            }
            return list;
        }
        static private List<SectionProperties> secPrList(Body body)
        {
            List<SectionProperties> list = new List<SectionProperties>(20);
            int l = body.ChildElements.Count();
            for (int i = 0; i < l; i++)
            {
                if (body.ChildElements.GetItem(i).GetType().ToString() == "DocumentFormat.OpenXml.Wordprocessing.Paragraph")
                {
                    Paragraph p = (Paragraph)body.ChildElements.GetItem(i);
                    if (p.ParagraphProperties != null)
                    {
                        if (p.ParagraphProperties.SectionProperties != null)
                        {
                            list.Add(p.ParagraphProperties.SectionProperties);
                        }
                    }
                }
            }
            if (body.ChildElements.GetItem(l - 1).GetType().ToString() == "DocumentFormat.OpenXml.Wordprocessing.Paragraph")
            {
                if (((Paragraph)body.ChildElements.GetItem(l - 1)).ParagraphProperties != null)
                {
                    if (((Paragraph)body.ChildElements.GetItem(l - 1)).ParagraphProperties.SectionProperties != null)
                    {
                        list.Add(((Paragraph)body.ChildElements.GetItem(l - 1)).ParagraphProperties.SectionProperties);
                    }
                }

            }
            return list;
        }
        /*
        function:the first page and the second should have no header
        params: list:the list of sectPrs
                intlist:the location of sectPrs in body
                wordpro:
                docxpath:
                root:xml root
                xmlDocx
        return:void
        */
        static private void firstSection(List<SectionProperties> list, List<int> intlist, WordprocessingDocument wordpro, string docxPath, XmlNode root, XmlDocument xmlDocx)
        {
            MainDocumentPart Mpart = wordpro.MainDocumentPart;
            SectionProperties s = null;
            if (list.Count == 0)
            {
                return;
            }
            s = list[0];
            TitlePage tp = s.GetFirstChild<TitlePage>();
            int location = intlist[0];
            IEnumerable<HeaderReference> headrs = s.Elements<HeaderReference>();
            HeaderReference headerfirst = null;
            HeaderReference headereven = null;
            HeaderReference headerdefault = null;
            FooterReference footerfirst = null;
            FooterReference footereven = null;
            FooterReference footerdefault = null;
            IEnumerable<FooterReference> footrs = s.Elements<FooterReference>();
            foreach (HeaderReference headr in headrs)
            {
                if (headr.Type == HeaderFooterValues.First)
                {
                    headerfirst = headr;
                }
                if (headr.Type == HeaderFooterValues.Even)
                {
                    headereven = headr;
                }
                if (headr.Type == HeaderFooterValues.Default)
                {
                    headerdefault = headr;
                }
            }
            foreach (FooterReference footr in footrs)
            {
                if (footr.Type == HeaderFooterValues.First)
                {
                    footerfirst = footr;
                }
                if (footr.Type == HeaderFooterValues.Even)
                {
                    footereven = footr;
                }
                if (footr.Type == HeaderFooterValues.Default)
                {
                    footerdefault = footr;
                }
            }

            //如果设置了首页不同
            #region
            if (tp != null)
            {
                if (headerfirst != null)
                {
                    string ID = headerfirst.Id.ToString();
                    HeaderPart hp = (HeaderPart)Mpart.GetPartById(ID);
                    Header header = hp.Header;
                    Paragraph p = header.GetFirstChild<Paragraph>();
                    if (header.InnerText != null)
                    {
                        if (header.InnerText.Trim().Length > 0)
                        {
                            p.RemoveAllChildren<Run>();
                            XmlElement xml = xmlDocx.CreateElement("Text");
                            xml.InnerText = "封面应无页眉";
                            root.AppendChild(xml);
                        }
                    }
                }
                if (footerfirst != null)
                {
                    string ID = footerfirst.Id.ToString();
                    FooterPart fp = (FooterPart)Mpart.GetPartById(ID);
                    Footer footer = fp.Footer;
                    Paragraph p = footer.GetFirstChild<Paragraph>();
                    if (footer.InnerText != null)
                    {
                        if (footer.InnerText.Trim().Length > 0)
                        {
                            p.RemoveAllChildren<Run>();
                            XmlElement xml = xmlDocx.CreateElement("Text");
                            xml.InnerText = "封面应无页脚";
                            root.AppendChild(xml);
                        }
                    }
                }
            }
            #endregion
            //若没有设置首页不同
            #region
            else
            {
                if (headerdefault != null)
                {
                    string ID = headerdefault.Id.ToString();
                    HeaderPart hp = (HeaderPart)Mpart.GetPartById(ID);
                    Header header = hp.Header;
                    Paragraph p = header.GetFirstChild<Paragraph>();
                    if (header.InnerText != null)
                    {
                        if (header.InnerText.Trim().Length > 0)
                        {
                            p.RemoveAllChildren<Run>();
                            XmlElement xml = xmlDocx.CreateElement("Text");
                            xml.InnerText = "封面应无页眉";
                            root.AppendChild(xml);
                        }
                    }
                }
                if (footerdefault != null)
                {
                    string ID = footerdefault.Id.ToString();
                    FooterPart fp = (FooterPart)Mpart.GetPartById(ID);
                    Footer footer = fp.Footer;
                    Paragraph p = footer.GetFirstChild<Paragraph>();
                    if (footer.InnerText != null)
                    {
                        if (footer.InnerText.Trim().Length > 0)
                        {
                            p.RemoveAllChildren<Run>();
                            XmlElement xml = xmlDocx.CreateElement("Text");
                            xml.InnerText = "封面应无页脚";
                            root.AppendChild(xml);
                        }
                    }
                }
                if (headereven != null)
                {
                    string ID = headereven.Id.ToString();
                    HeaderPart hp = (HeaderPart)Mpart.GetPartById(ID);
                    Header header = hp.Header;
                    Paragraph p = header.GetFirstChild<Paragraph>();
                    if (header.InnerText != null)
                    {
                        if (header.InnerText.Trim().Length > 0)
                        {
                            p.RemoveAllChildren<Run>();
                            XmlElement xml = xmlDocx.CreateElement("Text");
                            xml.InnerText = "封面应无页眉";
                            root.AppendChild(xml);
                        }
                    }
                }
                if (footereven != null)
                {
                    string ID = footereven.Id.ToString();
                    FooterPart fp = (FooterPart)Mpart.GetPartById(ID);
                    Footer footer = fp.Footer;
                    Paragraph p = footer.GetFirstChild<Paragraph>();
                    if (footer.InnerText != null)
                    {
                        if (footer.InnerText.Trim().Length > 0)
                        {
                            p.RemoveAllChildren<Run>();
                            XmlElement xml = xmlDocx.CreateElement("Text");
                            xml.InnerText = "封面应无页脚";
                            root.AppendChild(xml);
                        }
                    }
                }
            }
            #endregion
        }
        /*其他部分*/
        private void otherSectionHeader(List<SectionProperties> list, List<int> intlist, WordprocessingDocument wordpro, string docxPath, XmlNode root, XmlDocument xmlDocx, string name)
        {
            MainDocumentPart Mpart = wordpro.MainDocumentPart;
            HeaderReference headerfirst = null;
            HeaderReference headereven = null;
            HeaderReference headerdefault = null;
            if (list.Count == 0)
                return;
            SectionProperties s = list[0];
            IEnumerable<HeaderReference> headrs = s.Elements<HeaderReference>();
            for (int i = 2; i <= list.Count<SectionProperties>(); i++)
            {
                s = list[i - 1];
                List<int> chapterTitle = Tool.getTitlePosition(wordpro);
                string chapter = Tool.Chapter(chapterTitle, intlist[i - 1], wordpro.MainDocumentPart.Document.Body);
                //所在章
                TitlePage tp = s.GetFirstChild<TitlePage>();

                headrs = s.Elements<HeaderReference>();
                foreach (HeaderReference headr in headrs)
                {
                    if (headr.Type == HeaderFooterValues.First)
                    {
                        headerfirst = headr;
                    }
                    if (headr.Type == HeaderFooterValues.Even)
                    {
                        headereven = headr;
                    }
                    if (headr.Type == HeaderFooterValues.Default)
                    {
                        headerdefault = headr;
                    }
                }
                if (tp != null)//设置首页不同
                {
                    s.RemoveChild<TitlePage>(tp);
                    XmlElement xml = xmlDocx.CreateElement("Text");
                    xml.InnerText = "页眉命名不规范,不应该设置首页不同||" + chapter;
                    root.AppendChild(xml);
                }
                else
                {
                    //奇数页 
                    #region
                    if (headerdefault != null)
                    {
                        string ID = headerdefault.Id.ToString();
                        HeaderPart hp = (HeaderPart)Mpart.GetPartById(ID);
                        Header header = hp.Header;
                        Paragraph p = header.GetFirstChild<Paragraph>();
                        if (header != null && p != null)
                        {
                            if (header.InnerText != null)
                            {
                                string headername = header.InnerText.Trim();
                                //字体
                                if (!Tool.correctfonts(p, wordpro, "宋体", "Times New Roman"))
                                {
                                    Tool.change_rfonts(p, "宋体");
                                    XmlElement xml = xmlDocx.CreateElement("Text");
                                    xml.InnerText = "页眉字体应为宋体||" + chapter;
                                    root.AppendChild(xml);
                                }
                                //字号
                                if (!Tool.correctsize(p, wordpro, "21"))
                                {
                                    Tool.change_fontsize(p, "21");
                                    XmlElement xml = xmlDocx.CreateElement("Text");
                                    xml.InnerText = "页眉字号应为五号||" + chapter;
                                    root.AppendChild(xml);
                                }
                                //居中
                                if (!JustificationCenter(p, Mpart))
                                {
                                    Tool.change_center(p);
                                    XmlElement xml = xmlDocx.CreateElement("Text");
                                    xml.InnerText = "页眉应居中||" + chapter;
                                    root.AppendChild(xml);
                                }
                                if (headername != name)
                                {
                                    p.RemoveAllChildren<Run>();
                                    Tool.GenerateRun(p, name);
                                    Tool.change_center(p);
                                    Tool.change_rfonts(p, "宋体");
                                    Tool.change_rfonts(p, "21");
                                    XmlElement xml = xmlDocx.CreateElement("Text");
                                    xml.InnerText = "页眉命名不规范应为：标题||" + chapter;
                                    root.AppendChild(xml);
                                }
                                Tool.remmovejiachu(p);
                            }
                            else
                            {
                                Tool.GenerateRun(p, name);
                                Tool.change_center(p);
                                Tool.change_rfonts(p, "宋体");
                                Tool.change_rfonts(p, "21");
                                XmlElement xml = xmlDocx.CreateElement("Text");
                                xml.InnerText = "缺少页眉||" + chapter;
                                root.AppendChild(xml);
                            }
                            Tool.remmovejiachu(p);
                        }

                    }

                    #endregion
                    //偶数页
                    #region
                    if (headereven != null)
                    {
                        if (headereven.Id != null)
                        {
                            string ID1 = headereven.Id.ToString();
                            HeaderPart hp = (HeaderPart)Mpart.GetPartById(ID1);
                            Header header = hp.Header;
                            Paragraph p1 = header.GetFirstChild<Paragraph>();
                            if (header != null && p1 != null)
                            {
                                if (header.InnerText != null)
                                {
                                    //字体
                                    if (!Tool.correctfonts(p1, wordpro, "宋体", "Times New Roman"))
                                    {
                                        Tool.change_rfonts(p1, "宋体");
                                        XmlElement xml = xmlDocx.CreateElement("Text");
                                        xml.InnerText = "字体应为宋体||" + chapter;
                                        root.AppendChild(xml);
                                    }
                                    //字号
                                    if (!Tool.correctsize(p1, wordpro, "21"))
                                    {
                                        Tool.change_fontsize(p1, "21");
                                        XmlElement xml = xmlDocx.CreateElement("Text");
                                        xml.InnerText = "偶数页页眉字号应为五号||" + chapter;
                                        root.AppendChild(xml);
                                    }
                                    //居中
                                    if (!JustificationCenter(p1, Mpart))
                                    {
                                        Tool.change_center(p1);
                                        XmlElement xml = xmlDocx.CreateElement("Text");
                                        xml.InnerText = "偶数页页眉应居中||" + chapter;
                                        root.AppendChild(xml);
                                    }
                                    string headername = header.GetFirstChild<Paragraph>().InnerText.Trim();
                                    //合乎规范
                                    if (headername != name)
                                    {
                                        p1.RemoveAllChildren<Run>();
                                        Tool.GenerateRun(p1, name);
                                        Tool.change_rfonts(p1, "宋体");
                                        Tool.change_fontsize(p1, "21");
                                        Tool.change_center(p1);
                                        XmlElement xml = xmlDocx.CreateElement("Text");
                                        xml.InnerText = "页眉命名不规范，应为论文中文题目||" + chapter;
                                        root.AppendChild(xml);
                                    }
                                }
                                else
                                {
                                    p1.RemoveAllChildren<Run>();
                                    Tool.GenerateRun(p1, name);
                                    Tool.change_rfonts(p1, "宋体");
                                    Tool.change_fontsize(p1, "21");
                                    Tool.change_center(p1);
                                    XmlElement xml = xmlDocx.CreateElement("Text");
                                    xml.InnerText = "缺少偶数页页眉||" + chapter;
                                    root.AppendChild(xml);
                                }
                                Tool.remmovejiachu(p1);
                            }

                        }

                    }
                    #endregion
                }
            }
        }
        private void lastSectionHeader(WordprocessingDocument docx, XmlNode root, XmlDocument xmlDocx, string name)
        {
            SectionProperties scetpr = docx.MainDocumentPart.Document.Body.GetFirstChild<SectionProperties>();
            HeaderReference headerfirst = null;
            HeaderReference headereven = null;
            HeaderReference headerdefault = null;
            IEnumerable<HeaderReference> headrs = scetpr.Elements<HeaderReference>();
            foreach (HeaderReference headr in headrs)
            {
                if (headr.Type == HeaderFooterValues.First)
                {
                    headerfirst = headr;
                }
                if (headr.Type == HeaderFooterValues.Even)
                {
                    headereven = headr;
                }
                if (headr.Type == HeaderFooterValues.Default)
                {
                    headerdefault = headr;
                }
            }
            TitlePage tp = scetpr.GetFirstChild<TitlePage>();
            if (tp != null)//设置首页不同
            {
                scetpr.RemoveChild<TitlePage>(tp);
                XmlElement xml = xmlDocx.CreateElement("Text");
                xml.InnerText = "最后一节命名不规范,不应该设置首页不同||";
                root.AppendChild(xml);
            }
            if (headerdefault != null)
            {
                string ID = headerdefault.Id.ToString();
                HeaderPart hp = (HeaderPart)docx.MainDocumentPart.GetPartById(ID);
                Header header = hp.Header;
                Paragraph p = header.GetFirstChild<Paragraph>();
                string headername = header.GetFirstChild<Paragraph>().InnerText.Trim();
                if (headername != name)
                {
                    p.RemoveAllChildren<Run>();
                    Tool.GenerateRun(p, name);
                    Tool.change_rfonts(p, "宋体");
                    Tool.change_fontsize(p, "21");
                    Tool.change_center(p);
                    XmlElement xml = xmlDocx.CreateElement("Text");
                    xml.InnerText = "最后一节页眉命名不规范应为：标题";
                    root.AppendChild(xml);
                }
                Tool.remmovejiachu(p);
            }

            if (headereven != null)
            {
                string ID = headereven.Id.ToString();
                HeaderPart hp = (HeaderPart)docx.MainDocumentPart.GetPartById(ID);
                Header header = hp.Header;
                Paragraph p = header.GetFirstChild<Paragraph>();
                string headername = header.GetFirstChild<Paragraph>().InnerText.Trim();
                if (headername != name)
                {
                    p.RemoveAllChildren<Run>();
                    Tool.GenerateRun(p, name);
                    Tool.change_rfonts(p, "宋体");
                    Tool.change_fontsize(p, "21");
                    Tool.change_center(p);
                    XmlElement xml = xmlDocx.CreateElement("Text");
                    xml.InnerText = "最后一节页眉命名不规范应为：标题";
                    root.AppendChild(xml);
                }
                Tool.remmovejiachu(p);
            }

        }


        /*
        function:whether No.2 page before the first sectionPr
        是否有独创性声明,ture有
        params:
               location:the location of first sectionPr in body
               body:    body
        return:
               true: yes,it is before
               flase: no*/
        static bool no2PageInfirstSection(int location, Body body)
        {
            while (location > 0)
            {
                if (body.ChildElements.GetItem(location - 1).GetType().ToString() != "DocumentFormat.OpenXml.Wordprocessing.Paragraph")
                {
                    //continue;
                }
                else
                {
                    Paragraph p = (Paragraph)body.ChildElements.GetItem(location - 1);
                    if (p.InnerText.IndexOf("大连理工大学学位论文独创性声明") != -1)
                    {
                        return true;
                    }
                }
                location--;
            }
            return false;
        }
        /*
        function：judge a paragraph's position center
        判断是否居中
        params: p:paragraph
                Mpart:MainDocumentPart
        return:
                bool
        */
        public static bool JustificationCenter(Paragraph p, MainDocumentPart Mpart)
        {
            Justification jc = null;
            ParagraphStyleId pid = null;
            if (p.ParagraphProperties != null)
            {
                if ((jc = p.ParagraphProperties.Justification) != null)
                {
                    if (jc.Val != JustificationValues.Center)
                    { return false; }
                }
                if (jc != null)
                {
                    if ((pid = p.ParagraphProperties.ParagraphStyleId) != null)
                    {
                        Styles styles = Mpart.StyleDefinitionsPart.Styles;
                        Style style = null;
                        IEnumerable<Style> stys = styles.OfType<Style>();
                        foreach (Style sty in stys)
                        {
                            if (sty.StyleId.ToString() == pid.ToString())
                            {
                                style = sty;
                                break;
                            }
                        }
                        if (style != null)
                        {
                            if (style.StyleParagraphProperties != null)
                            {
                                if ((jc = style.StyleParagraphProperties.Justification) != null)
                                {
                                    if (jc.Val != JustificationValues.Center)
                                    { return false; }
                                }
                            }
                        }
                    }
                }
            }
            return true;
        }

    }
}
