using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Xml;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using PaperFormatDetection.Tools;
using PaperFormatDetection.Frame;
using System.IO;

namespace PaperFormatDetection.Format
{
    /* 目录样式检测类 */
    public class Contents : ModuleFormat
    {   
        /* 构造函数 */
        public Contents(List<Module> modList, PageLocator locator, int masterType) : base(modList, locator, masterType)
        {

        }

        /* 继承自ModuleFormat中的getStyle方法 */
        public override void getStyle(WordprocessingDocument doc, String fileName)
        {
            string xmlFullPath = fileName + "\\Contents.xml";//xml模板文件保存路径
            CreateXmlFile(xmlFullPath);
            pageNum =  3;
            getContentsXml(doc, xmlFullPath);//检测目录样式
        }

        /* 创建用于保存检测结果的XM文件 */
        private void CreateXmlFile(string xmlPath)
        {
            XmlDocument xmlDoc = new XmlDocument();
            //创建类型声明节点  
            XmlNode node = xmlDoc.CreateXmlDeclaration("1.0", "utf-8", "");
            xmlDoc.AppendChild(node);
            //创建根节点  
            XmlNode root = xmlDoc.CreateElement("Contents");
            //以下为结点创建
            XmlElement xe1 = xmlDoc.CreateElement("ContentsTitle");
            xe1.SetAttribute("name", "目录标题");
            XmlElement xe2 = xmlDoc.CreateElement("ContentsText");
            xe2.SetAttribute("name", "目录内容");
            XmlElement xe3 = xmlDoc.CreateElement("spErroInfo");
            xe3.SetAttribute("name", "特殊错误信息");
            XmlElement xe4 = xmlDoc.CreateElement("partName");
            xe4.SetAttribute("name", "提示名称");
            XmlElement xe5 = xmlDoc.CreateElement("Text");
            xe5.InnerText = "-----------------目录-----------------";
            xe4.AppendChild(xe5);
            root.AppendChild(xe1);
            root.AppendChild(xe2);
            root.AppendChild(xe3);
            root.AppendChild(xe4);
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

        /* 目录检测 */
        private void getContentsXml(WordprocessingDocument doc, String xmlFullPath)
        {
            Body body = doc.MainDocumentPart.Document.Body;
            MainDocumentPart mainPart = doc.MainDocumentPart;
            IEnumerable<Paragraph> paras = body.Elements<Paragraph>();
            ParagraphProperties pPr = null;
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(xmlFullPath);
            XmlNode root = xmlDoc.SelectSingleNode("Contents/ContentsTitle");
            XmlNode spRoot = xmlDoc.SelectSingleNode("Contents/spErroInfo");
            bool isContents = false;
            bool overFour = false;
            String curCount = "";
            Paragraph temppara = null;

            List<String> titleList = new List<string>();
            foreach (Paragraph p in paras)
            {
                Run r = p.GetFirstChild<Run>();
                String fullText = "";
                if (r != null)
                {
                    fullText = Tool.getFullText(p).Trim();
                } 
                //是否到了目录标题
                if (fullText.Replace(" ", "") == "目录" && !isContents)
                {
                    isContents = true;
                    pageNum = this.getPageNum(pageNum, fullText);
                    IEnumerable<Run> pRunList = p.Elements<Run>();
                    int spaceCount = Tool.getSpaceCount(fullText);
                    //空格判断
                    if (spaceCount != 4)
                    {
                        XmlElement xe = xmlDoc.CreateElement("Text");
                        xe.InnerText = this.addPageInfo(pageNum) + "目录标题“目录”两字之间应有4个空格";
                        spRoot.AppendChild(xe);
                        /*更改目录标题*/
                        if (p.Elements<Run>().Count() == 1)
                            p.GetFirstChild<Run>().GetFirstChild<Text>().Text = "目    录";
                        else
                        {
                            IEnumerable<Run> runs = p.Elements<Run>();
                            int num = 0;
                            foreach (Run rr in runs)
                            {
                                num++;
                                if (num != 1)
                                    rr.GetFirstChild<Text>().Text = null;
                            }
                            p.GetFirstChild<Run>().GetFirstChild<Text>().Text = "目    录";
                        }
                    }
                    pPr = p.GetFirstChild<ParagraphProperties>();
                    if (pPr != null)
                    {
                        if (pPr.GetFirstChild<Justification>() != null)
                        {
                            if (pPr.GetFirstChild<Justification>().Val.ToString().ToLower() != "center")
                            {
                                XmlElement xe1 = xmlDoc.CreateElement("Text");
                                xe1.InnerText = this.addPageInfo(pageNum) + "目录标题未居中";
                                spRoot.AppendChild(xe1);
                                Tool.change_center(p);
                            }
                        }
                        else
                        {
                            XmlElement xe1 = xmlDoc.CreateElement("Text");
                            xe1.InnerText = this.addPageInfo(pageNum) + "目录标题未居中";
                            spRoot.AppendChild(xe1);
                            Tool.change_center(p);
                        }
                    }
                    //取出段落所有run，判断有文本的run的设置时候正确
                    if (pRunList != null)
                    {
                        bool flag1 = true;
                        bool flag2 = true;
                        foreach (Run pr in pRunList)
                        {
                            Text t = pr.GetFirstChild<Text>();
                            if (t != null)
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
                                                if (Rrpr.GetFirstChild<RunFonts>().Ascii != "黑体")
                                                {
                                                    flag1 = false;
                                                }
                                            }
                                        }
                                        if (Rrpr.GetFirstChild<FontSize>() != null)
                                        {
                                            if (Rrpr.GetFirstChild<FontSize>().Val != null)
                                            {
                                                if (Rrpr.GetFirstChild<FontSize>().Val != "30")
                                                {
                                                    flag2 = false;
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            
                        }
                        if (!flag1)
                        {
                            XmlElement xe1 = xmlDoc.CreateElement("Text");
                            xe1.InnerText = this.addPageInfo(pageNum) + "目录标题字体错误，应为黑体";
                            spRoot.AppendChild(xe1);
                            Tool.change_rfonts(p, "黑体");
                        }
                        if (!flag2)
                        {
                            XmlElement xe1 = xmlDoc.CreateElement("Text");
                            xe1.InnerText = this.addPageInfo(pageNum) + "目录标题字号错误，应为小三号";
                            spRoot.AppendChild(xe1);
                            Tool.change_fontsize(p, "30");
                        }
                    }
                    continue;
                }
                if (isContents)
                {
                    /*除目录标题，格式刷字号、字体*/
                    Tool.change_fontsize(p, "24");
                    IEnumerable<Run> runlist = p.Elements<Run>();
                    foreach (Run rt in runlist)
                    {
                        if (rt.RunProperties != null)
                        {
                            if (rt.RunProperties.GetFirstChild<RunFonts>() != null)
                            {
                                //字体改，此处要分清要改的是其中的属性，还是其中的内容
                                rt.RunProperties.GetFirstChild<RunFonts>().Hint = FontTypeHintValues.EastAsia;
                                rt.RunProperties.GetFirstChild<RunFonts>().Ascii = "Times New Roman";
                                rt.RunProperties.GetFirstChild<RunFonts>().EastAsia = "宋体";
                            }
                            else
                            {
                                //添加标签
                                RunFonts rfont = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Times New Roman", EastAsia = "宋体" };
                                rt.RunProperties.AppendChild<RunFonts>(rfont);
                            }
                        }
                    }

                    if (fullText.Replace(" ", "") == "结论")
                    {

                    }
                    Hyperlink h = p.GetFirstChild<Hyperlink>();//取这个标签
                    Regex titleOne = new Regex("[1-9]");
                    IEnumerable<Run> runList = p.Elements<Run>();
                    bool flag = false;
                    if (runList != null)
                    {
                        foreach (Run rr in runList)
                        {
                            //这个是为了处理未更新域的目录的
                            //更新域过的目录会有一个Hyperlinkd 的标签，没有跟新过的会有一个FieldChar标签
                            if (rr.GetFirstChild<FieldChar>() != null)
                            {
                                flag = true;
                            }
                        }
                    }
                    //判断是更新过的   
                    if (h != null)
                    {
                        IEnumerable<Run> pRunList = h.Elements<Run>();
                        String str = getHyperlinkFullText(h);
                        pageNum = this.getPageNum(str);
                        bool flag1 = true;
                        bool flag2 = true;
                        foreach (Run pr in pRunList)
                        {
                            if (pr != null)
                            {
                                Text t = pr.GetFirstChild<Text>();
                                RunProperties Rrpr = pr.GetFirstChild<RunProperties>();
                                if (Rrpr != null)
                                {
                                    if (Rrpr.GetFirstChild<RunFonts>() != null)
                                    {
                                        if (Rrpr.GetFirstChild<RunFonts>().Ascii != null)
                                        {
                                            if (Rrpr.GetFirstChild<RunFonts>().Ascii != "宋体" && Rrpr.GetFirstChild<RunFonts>().Ascii != "Times New Roman")
                                            {
                                                if (t != null)
                                                {
                                                    if(t.Text.Replace(" ", "") != "")
                                                    {
                                                        flag1 = false;
                                                    }
                                                }                                               
                                            }
                                        }
                                    }
                                    if (Rrpr.GetFirstChild<FontSize>() != null)
                                    {
                                        if (Rrpr.GetFirstChild<FontSize>().Val != null)
                                        {
                                            if (Rrpr.GetFirstChild<FontSize>().Val != "24")
                                            {
                                                flag2 = false;
                                            }
                                        }
                                    }
                                }
                                
                            }
                        }
                        if (!flag1)
                        {
                            XmlElement xe1 = xmlDoc.CreateElement("Text");
                            xe1.InnerText = this.addPageInfo(pageNum) + "目录中该章节标题字体错误,中文应为“宋体”英文为“Times New Roman”：{" + str+"}";
                            spRoot.AppendChild(xe1);
                        }
                        if (!flag2)
                        {
                            XmlElement xe1 = xmlDoc.CreateElement("Text");
                            xe1.InnerText = this.addPageInfo(pageNum) + "目录中该章节标题字号错误，应为小四号：{" + str + "}";
                            spRoot.AppendChild(xe1);
                            //Tool.change_fontsize(p, "24");
                        }
                        //通过.来判断是否超过3级
                        if (str.Split('.').Length > 4)
                        {
                            overFour = true;
                        }
                        if (str.Length > 1)
                        {
                            int spaceP = str.Trim().IndexOf(" ");
                            Match m = titleOne.Match(str.Substring(0, 1));
                            if (m.Success)
                            {
                                if (getSpaceCount(str.Trim().Substring(0, spaceP+3)) != 2)
                                {
                                    XmlElement xe1 = xmlDoc.CreateElement("Text");
                                    xe1.InnerText = this.addPageInfo(pageNum) + "目录中该章节标题与序号之间应空两个空格：" + "{" + str + "}";
                                    spRoot.AppendChild(xe1);
                                    Tool.addComment(doc, p, "目录中该章节标题与序号之间应空两个空格");
                                }
                                else
                                {
                                    if (str.Trim().Length > spaceP+5)
                                    {
                                        if (getSpaceCount(str.Trim().Substring(0, spaceP + 4)) > 2)
                                        {
                                            XmlElement xe1 = xmlDoc.CreateElement("Text");
                                            xe1.InnerText = this.addPageInfo(pageNum) + "目录中该章节标题与序号之间应空两个空格:" + "{" + str + "}";
                                            spRoot.AppendChild(xe1);
                                            Tool.addComment(doc, p, "目录中该章节标题与序号之间应空两个空格");
                                        }
                                    }                                   
                                }
                            }
                        }                       
                        xmlDoc.Save(xmlFullPath);
                    }
                    else if (flag)
                    {
                        //上面的是跟新过的，这个部分是没跟新过的，检测方法一样
                        IEnumerable<Run> pRunList = p.Elements<Run>();
                        String str = Tool.getFullText(p);
                        pageNum = this.getPageNum(str);
                        bool flag1 = true;
                        bool flag2 = true;
                        foreach (Run pr in pRunList)
                        {
                            if (pr != null)
                            {
                                Text t = pr.GetFirstChild<Text>();
                                RunProperties Rrpr = pr.GetFirstChild<RunProperties>();
                                if (Rrpr != null)
                                {
                                    if (Rrpr.GetFirstChild<RunFonts>() != null)
                                    {
                                        if (Rrpr.GetFirstChild<RunFonts>().Ascii != null)
                                        {
                                            if (Rrpr.GetFirstChild<RunFonts>().Ascii != "宋体" && Rrpr.GetFirstChild<RunFonts>().Ascii != "Times New Roman")
                                            {
                                                 if (t != null)
                                                {
                                                    if(t.Text.Replace(" ", "") != "")
                                                    {

                                                        flag1 = false;
                                                    }
                                                } 
                                            }
                                        }
                                    }
                                    if (Rrpr.GetFirstChild<FontSize>() != null)
                                    {
                                        if (Rrpr.GetFirstChild<FontSize>().Val != null)
                                        {
                                            if (Rrpr.GetFirstChild<FontSize>().Val != "24")
                                            {
                                                flag2 = false;
                                            }
                                        }
                                    }
                                }

                            }
                        }
                        //不一致时flag为假，向xml中写入错误提示
                        if (!flag1)
                        {
                            XmlElement xe1 = xmlDoc.CreateElement("Text");
                            xe1.InnerText = this.addPageInfo(pageNum) + "目录中该章节标题字体错误,中文应为“宋体”英文为“Times New Roman”：{" + str + "}";
                            spRoot.AppendChild(xe1);
                        }
                        if (!flag2)
                        {
                            XmlElement xe1 = xmlDoc.CreateElement("Text");
                            xe1.InnerText = this.addPageInfo(pageNum) + "目录中该章节标题字号错误，应为小四号：{" + str + "}";
                            spRoot.AppendChild(xe1);
                        }
                        if (str.Trim().Length > 7)
                        {
                            if(!str.Trim().Substring(0, 7).Contains(" "))
                            {
                                if (str.Trim().Substring(0, 7).Split('.').Length > 4)
                                {
                                    overFour = true;
                                    Tool.addComment(doc, p, "目录章节最多不应超过3级");
                                }
                            }
                        }
                        
                        if (str.Length > 1)
                        {
                            int spaceP = str.Trim().IndexOf(" ");
                            Match m = titleOne.Match(str.Substring(0, 1));
                            if (m.Success)
                            {
                                curCount = str.Substring(0, 1);
                                if (str.Substring(0, 1) != curCount)
                                {
                                    titleList.Clear();
                                }
                                titleList.Add(str);
                                if (str.Substring(0, 1) != curCount)
                                {

                                }
                                if (getSpaceCount(str.Trim().Substring(0, spaceP + 3)) != 2)
                                {
                                    XmlElement xe1 = xmlDoc.CreateElement("Text");
                                    xe1.InnerText = this.addPageInfo(pageNum) + "目录中该章节标题与序号之间应空两个空格：" + "{" + str + "}";
                                    spRoot.AppendChild(xe1);
                                    Tool.addComment(doc, p, "目录中该章节标题与序号之间应空两个空格");
                                }
                                else
                                {
                                    if (str.Trim().Length > spaceP + 5)
                                    {
                                        if (getSpaceCount(str.Trim().Substring(0, spaceP + 4)) > 2)
                                        {
                                            XmlElement xe1 = xmlDoc.CreateElement("Text");
                                            xe1.InnerText = this.addPageInfo(pageNum) + "目录中该章节标题与序号之间应空两个空格:" + "{" + str + "}";
                                            spRoot.AppendChild(xe1);
                                            Tool.addComment(doc, p, "目录中该章节标题与序号之间应空两个空格");
                                        }
                                    }
                                }
                            }
                        }
                        xmlDoc.Save(xmlFullPath);

                    }
                    else
                    {
                        break;
                    }               
                }
                temppara = p;
            }
            if (overFour)
            {
                XmlElement xe1 = xmlDoc.CreateElement("Text");
                xe1.InnerText = this.addPageInfo(pageNum) +　"目录章节最多不应超过3级";
                spRoot.AppendChild(xe1);
                xmlDoc.Save(xmlFullPath);
            }
        }
        
        /** 
         * 该方法用于获取目录类型的段落中的目录文本，
         * 目录类型为Hyperlink时，使用该方法 
         */
        private static String getHyperlinkFullText(Hyperlink p)
        {
            String text = "";
            IEnumerable<Run> list = p.Elements<Run>();
            TabChar pTab = new TabChar();
            foreach (Run r in list)
            {
                Text pText = r.GetFirstChild<Text>();
                if (pText != null)
                {
                    text += pText.Text;
                }
                if (r.LastChild.GetType() == pTab.GetType())
                {
                    text += '\t';
                }
            }
            return text;
        }
        

        /* 获取空格数 */
        private static int getSpaceCount(String str)
        {
            List<int> lst = new List<int>();
            char[] chr = str.ToCharArray();
            int iSpace = 0;
            foreach (char c in chr)
            {
                if (char.IsWhiteSpace(c))
                {
                    iSpace++;
                }
            }
            return iSpace;
        }

    }
}
