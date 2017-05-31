using System;
using System.Collections.Generic;
using System.Xml;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using PaperFormatDetection.Tools;
using PaperFormatDetection.Frame;

namespace PaperFormatDetection.Format
{
    public class Abstract : ModuleFormat
    {

        //关键词屏蔽是调用，包含这些字符串的直接过滤
        String[] keyWordList = { "关键词：", "关键词", "：" };
        
        /* 构造函数 */
        public Abstract(List<Module> modList, PageLocator locator, int masterType) : base(modList, locator, masterType)
        {
            
        }
        
        /* 继自ModuleFormat中的getStyle方法 */
        public override void getStyle(WordprocessingDocument doc, String fileName)
        {
            string xmlFullPath = fileName + "\\Abstract.xml";//xml模板文件保存路径
            CreateXmlFile(xmlFullPath);
            pageNum = 3;
            getCNAbstractXml(doc, xmlFullPath);//中文摘要检测
            getENAbstractTitleXml(doc, xmlFullPath);//英文摘要上方论文英文题目检测
            getENAbstractXml(doc, xmlFullPath);//英文摘要检测
        }

        private void CreateXmlFile(string xmlPath)
        {
            XmlDocument xmlDoc = new XmlDocument();
            //创建类型声明节点  
            XmlNode node = xmlDoc.CreateXmlDeclaration("1.0", "utf-8", "");
            xmlDoc.AppendChild(node);
            //创建根节点  
            XmlNode root = xmlDoc.CreateElement("Abstract");
            XmlElement xe7 = xmlDoc.CreateElement("spErroInfo");
            xe7.SetAttribute("name", "特殊错误信息");
            XmlElement xe8 = xmlDoc.CreateElement("partName");
            xe8.SetAttribute("name", "提示名称");
            XmlElement xe9 = xmlDoc.CreateElement("Text");
            xe9.InnerText = "-----------------摘要-----------------";
            xe8.AppendChild(xe9);
            root.AppendChild(xe7);
            root.AppendChild(xe8);
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

        /* 获取关键词与摘要正文的空行数 */
        static int countSpaceLine(WordprocessingDocument doc)
        {
            Body body = doc.MainDocumentPart.Document.Body;
            MainDocumentPart mainPart = doc.MainDocumentPart;
            IEnumerable<Paragraph> paras = body.Elements<Paragraph>();         
            int keyWordPosition = -1;//获取关键词位置
            List<Paragraph> pList = Tool.toList(paras);
            int spaceCount = 0;//空行数
            foreach (Paragraph p in paras)
            {
                keyWordPosition++;
                Run r = p.GetFirstChild<Run>();
                String fullText = "";
                if (r != null)
                {
                    fullText = Tool.getFullText(p).Trim();
                }
                if (fullText.Replace(" ", "").Length > 4)
                {
                    //判断该段落是否是关键词
                    if (fullText.Replace(" ", "").Substring(0, 4) == "关键词：")
                    {
                        if(pList.Count>0 && keyWordPosition > 0)
                        {
                            for (int i= keyWordPosition-1; i < keyWordPosition; i--)
                            {                             
                                Paragraph temp = pList[i];                                
                                if (temp.GetFirstChild<Run>() == null)
                                {
                                    //run为空，空行数加一
                                    spaceCount++;
                                }
                                else
                                {
                                    break;
                                }
                            }                            
                        }
                        break;                       
                    }
                }
            }
            return spaceCount;
        }
        static int countSpaceLineEN(WordprocessingDocument doc)
        {
            Body body = doc.MainDocumentPart.Document.Body;
            MainDocumentPart mainPart = doc.MainDocumentPart;
            IEnumerable<Paragraph> paras = body.Elements<Paragraph>();
            int keyWordPosition = -1;//获取关键词位置
            List<Paragraph> pList = Tool.toList(paras);
            int spaceCount = 0;//空行数
            foreach (Paragraph p in paras)
            {
                keyWordPosition++;
                Run r = p.GetFirstChild<Run>();
                String fullText = "";
                if (r != null)
                {
                    fullText = Tool.getFullText(p).Trim();
                }
                if (fullText.Replace(" ", "").Length > 4)
                {
                    //判断该段落是否是关键词
                    if (fullText.Replace(" ", "") == "abstract")
                    {
                        if (pList.Count > 0 && keyWordPosition > 0)
                        {
                            for (int i = keyWordPosition - 1; i < keyWordPosition; i--)
                            {
                                Paragraph temp = pList[i];
                                if (temp.GetFirstChild<Run>() == null)
                                {
                                    //run为空，空行数加一
                                    spaceCount++;
                                }
                                else
                                {
                                    break;
                                }
                            }
                        }
                        break;
                    }
                }
            }
            return spaceCount;
        }
        /**
        对比中文摘要
        */
        private void getCNAbstractXml(WordprocessingDocument doc, String xmlFullPath)
        {
            Body body = doc.MainDocumentPart.Document.Body;
            MainDocumentPart mainPart = doc.MainDocumentPart;
            IEnumerable<Paragraph> paras = body.Elements<Paragraph>();
            ParagraphProperties pPr = null;
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(xmlFullPath);
            XmlNode root = xmlDoc.SelectSingleNode("Abstract/CNAbstractTitle");
            XmlNode spRoot = xmlDoc.SelectSingleNode("Abstract/spErroInfo");
            bool isAbsTittle = false;//标志用来判断是否为摘要段落
            List<int> strCount = new List<int>();
            int ss = countSpaceLine(doc);    //正文与关键词空行数 
            
            List<int> intlist = Tool.secPrListInt(doc.MainDocumentPart.Document.Body);
            List<SectionProperties> list = Tool.secPrList(doc.MainDocumentPart.Document.Body);
            int count = -1;
            foreach (Paragraph p in paras)
            {
                count++;
                Run r = p.GetFirstChild<Run>();
                String fullText = "";
                if (r != null)
                {
                    //run不为空，获取该段落的全部文本
                    fullText = Tool.getFullText(p).Trim();
                }
                if (fullText.Replace(" ", "") == "摘要" && !isAbsTittle)
                {
                    //遇到摘要段落，标志致为true
                    isAbsTittle = true;
                    IEnumerable<Run> pRunList = p.Elements<Run>();
                    //获取空格数
                    int spaceCount = Tool.getSpaceCount(fullText);
                    //获取页码
                    pageNum = getPageNum(pageNum, fullText);
                    if (spaceCount != 4)
                    {
                        //空格数不为4的提示
                        XmlElement xe1 = xmlDoc.CreateElement("Text");
                        xe1.InnerText = addPageInfo(pageNum) + "中文摘要标题“摘要”两字之间应有4个空格";
                        spRoot.AppendChild(xe1);
                        /*修改摘要标题*/
                        if(p.Elements<Run>().Count() == 1)
                            p.GetFirstChild<Run>().GetFirstChild<Text>().Text = "摘    要";
                        else
                        {
                            IEnumerable<Run> runs = p.Elements<Run>();
                            int num = 0;
                            foreach(Run rr in runs)
                            {
                                num++;
                                if (num != 1)
                                    rr.GetFirstChild<Text>().Text = null;
                            }
                            p.GetFirstChild<Run>().GetFirstChild<Text>().Text = "摘    要";
                        }
                    }                    
                    pPr = p.GetFirstChild<ParagraphProperties>();
                    //居中判断
                    if (pPr != null)
                    {
                        if (pPr.GetFirstChild<Justification>() != null)
                        {
                            if (pPr.GetFirstChild<Justification>().Val.ToString().ToLower() != "center")
                            {
                                XmlElement xe1 = xmlDoc.CreateElement("Text");
                                xe1.InnerText = addPageInfo(pageNum) + "中文摘要标题未居中";
                                spRoot.AppendChild(xe1);
                                Tool.change_center(p);
                            }
                        }
                        
                        if (pPr.GetFirstChild<SpacingBetweenLines>() != null)
                        {
                            //行距判断
                            if (pPr.GetFirstChild<SpacingBetweenLines>().Line != null)
                            {
                                if (pPr.GetFirstChild<SpacingBetweenLines>().Line.Value != "360")
                                {
                                    XmlElement xe1 = xmlDoc.CreateElement("Text");
                                    xe1.InnerText = addPageInfo(pageNum) + "中文摘要标题行距错误，应为1.5倍行距";
                                    spRoot.AppendChild(xe1);
                                    Tool.changeSpacingLine(p, 1.5);
                                }
                            }

                            //段前间距判断
                            if (!Tool.correctSpacingBetweenLines_Be(p, count, intlist, list, doc, "0", false))
                            {
                                XmlElement xe1 = xmlDoc.CreateElement("Text");
                                xe1.InnerText = addPageInfo(pageNum) + "中文摘要标题段前间距错误，应为0行";
                                spRoot.AppendChild(xe1);
                                Tool.changeSpacingBefore(p, 0);
                            }

                            //段后间距判断
                            if (!Tool.correctSpacingBetweenLines_Af(p, count, intlist, list, doc, "100", false))
                            {
                                XmlElement xe1 = xmlDoc.CreateElement("Text");
                                xe1.InnerText = addPageInfo(pageNum) + "中文摘要标题段后间距错误，应为11磅";
                                spRoot.AppendChild(xe1);
                                Tool.changeSpacingAfter_point(p, 11);
                            }
                        }
                    }
                    //这一部分是将一个段落的run全部取出，判断每一个run里的设置是否都是一致的
                    if (pRunList != null)
                    {
                        bool flag1 = true;
                        bool flag2 = true;
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
                                            if (Rrpr.GetFirstChild<RunFonts>().Ascii != "黑体" && Rrpr.GetFirstChild<RunFonts>().Ascii != "Times New Roman")
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
                        if (!flag1)
                        {
                            XmlElement xe1 = xmlDoc.CreateElement("Text");
                            xe1.InnerText = addPageInfo(pageNum) + "中文摘要标题字体错误，应为黑体";
                            spRoot.AppendChild(xe1);
                            Tool.change_rfonts(p, "黑体");
                        }
                        if (!flag2)
                        {
                            XmlElement xe1 = xmlDoc.CreateElement("Text");
                            xe1.InnerText = addPageInfo(pageNum) + "中文摘要标题字号错误，应为小三号";
                            spRoot.AppendChild(xe1);
                            Tool.change_fontsize(p, "30");
                        }
                    }
                    xmlDoc.Save(xmlFullPath);
                    continue;
                }
                //遇到关键词，摘要正文结束
                if (fullText.Replace(" ", "").Length > 4)
                {
                    if (fullText.Replace(" ", "").Substring(0, 4) == "关键词：")
                    {
                        isAbsTittle = false;
                    }
                }
                //isAbsTittle为true表示接下来的段落是摘要正文
                if (isAbsTittle)
                {
                    IEnumerable<Run> pRunList = p.Elements<Run>();
                    String fullTextt = Tool.getFullText(p).Trim();
                    strCount.Add(Tool.GetHanNumFromString(fullTextt));
                    //所在页码
                    pageNum = getPageNum(pageNum, fullTextt);

                    /*更改缩进*/
                    Tool.changeIntent(p, 2);
                    Tool.deleteSpacingChar(p);
                    /*更改段前段后间距*/
                    Tool.changeSpacingBefore(p, 0);
                    Tool.changeSpacingAfter(p, 0);
                    /*更改字体、字号*/
                    Tool.change_rfonts(p, "宋体");
                    Tool.change_fontsize(p, "24");
                    
                    ParagraphProperties prpt = p.GetFirstChild<ParagraphProperties>();
                    if (prpt != null)
                    {
                        /*
                         缩进检测
                         1.FirstLine为空时，应有四个空格。
                         2.FirstLine=“0”，且FirstLineChars == "200"时为两缩进。
                         3.没有Indentation标签
                        */
                        ParagraphMarkRunProperties prPr = prpt.GetFirstChild<ParagraphMarkRunProperties>();
                        if (prpt.GetFirstChild<Indentation>() != null)
                        {
                            if (prpt.GetFirstChild<Indentation>() != null)
                            {
                                if (prpt.GetFirstChild<Indentation>().FirstLine == null || prpt.GetFirstChild<Indentation>().FirstLine.Value == "0")
                                {
                                    if (Tool.getFullText(p).Length > 4 && Tool.getFullText(p).Substring(0, 3) == "    ")
                                    {
                                    }
                                    else
                                    {
                                        if (fullTextt.Length > 5)
                                        {
                                            XmlElement xe1 = xmlDoc.CreateElement("Text");
                                            xe1.InnerText = addPageInfo(pageNum) + "中文摘要正文段落首行缩进错误，应为两个汉字的缩进量:" + "{" + fullTextt.Substring(0, 5) + "......}";
                                            spRoot.AppendChild(xe1);
                                        }
                                        else
                                        {
                                            XmlElement xe1 = xmlDoc.CreateElement("Text");
                                            xe1.InnerText = addPageInfo(pageNum) + "中文摘要正文段落首行缩进错误，应为两个汉字的缩进量:" + "{" + fullTextt + "......}";
                                            spRoot.AppendChild(xe1);
                                        }
                                    }
                                }
                                else
                                {
                                    if (prpt.GetFirstChild<Indentation>().FirstLineChars == "200")
                                    {
                                    }
                                    else
                                    {
                                        if (fullTextt.Length > 5)
                                        {
                                            XmlElement xe1 = xmlDoc.CreateElement("Text");
                                            xe1.InnerText = addPageInfo(pageNum) + "中文摘要正文段落首行缩进错误，应为两个汉字的缩进量:" + "{" + fullTextt.Substring(0, 5) + "......}";
                                            spRoot.AppendChild(xe1);
                                        }
                                        else
                                        {
                                            XmlElement xe1 = xmlDoc.CreateElement("Text");
                                            xe1.InnerText = addPageInfo(pageNum) + "中文摘要正文段落首行缩进错误，应为两个汉字的缩进量:" + "{" + fullTextt + "......}";
                                            spRoot.AppendChild(xe1);
                                        }
                                    }
                                }
                            }
                            else
                            {
                                if (fullTextt.Length > 5)
                                {
                                    XmlElement xe1 = xmlDoc.CreateElement("Text");
                                    xe1.InnerText = addPageInfo(pageNum) + "中文摘要正文段落首行缩进错误，应为两个汉字的缩进量:" + "{" + fullTextt.Substring(0, 5) + "......}";
                                    spRoot.AppendChild(xe1);
                                }
                                else
                                {
                                    XmlElement xe1 = xmlDoc.CreateElement("Text");
                                    xe1.InnerText = addPageInfo(pageNum) + "中文摘要正文段落首行缩进错误，应为两个汉字的缩进量:" + "{" + fullTextt + "......}";
                                    spRoot.AppendChild(xe1);
                                }
                            }
                            /*
                         * 这里判断首行缩进有3种：
                         * 1.FirstLine值为480，段前没有人为的空格，设置正确
                         * 2.FirstLine值为240，段前有一个空格，设置正确
                         * 3.FirstLine值为0，段前两个空格，设置正确
                            */
                            /*if (prpt.GetFirstChild<Indentation>().FirstLine == "480")
                            {
                                if (Tool.getFullText(p).Length > 1)
                                {
                                    if (Tool.getFullText(p).Substring(0, 1) == " ")
                                    {
                                        if (fullTextt.Length > 5)
                                        {
                                            XmlElement xe1 = xmlDoc.CreateElement("Text");
                                            xe1.InnerText = addPageInfo(pageNum) + "中文摘要正文段落首行缩进错误，应为两个汉字的缩进量:" +  "{" + fullTextt.Substring(0, 5) + "......}";
                                            spRoot.AppendChild(xe1);
                                        }
                                        else
                                        {
                                            XmlElement xe1 = xmlDoc.CreateElement("Text");
                                            xe1.InnerText = addPageInfo(pageNum) + "中文摘要正文段落首行缩进错误，应为两个汉字的缩进量:" +  "{" + fullTextt + "......}";
                                            spRoot.AppendChild(xe1);
                                        }
                                    }
                                }
                            }
                            else if (prpt.GetFirstChild<Indentation>().FirstLine == "240")
                            {
                                if (Tool.getFullText(p).Length > 1)
                                {
                                    if (Tool.getFullText(p).Substring(0, 1) == " ")
                                    {
                                    }
                                    else
                                    {
                                        if (fullTextt.Length > 5)
                                        {
                                            XmlElement xe1 = xmlDoc.CreateElement("Text");
                                            xe1.InnerText = addPageInfo(pageNum) + "中文摘要正文段落首行缩进错误，应为两个汉字的缩进量:" + "{" + fullTextt.Substring(0, 5) + "······}";
                                            spRoot.AppendChild(xe1);
                                        }
                                        else
                                        {
                                            XmlElement xe1 = xmlDoc.CreateElement("Text");
                                            xe1.InnerText = addPageInfo(pageNum) + "中文摘要正文段落首行缩进错误，应为两个汉字的缩进量:" + "{" + fullTextt + "······}";
                                            spRoot.AppendChild(xe1);
                                        }
                                    }
                                }
                            }
                            else
                            {
                                if (Tool.getFullText(p).Length > 2)
                                {
                                    if (Tool.getFullText(p).Substring(0, 2) == "  ")
                                    {
                                        break;
                                    }
                                    else
                                    {
                                        if (fullTextt.Length > 5)
                                        {
                                            XmlElement xe1 = xmlDoc.CreateElement("Text");
                                            xe1.InnerText = addPageInfo(pageNum) + "中文摘要正文段落首行缩进错误，应为两个汉字的缩进量:" +  "{" + fullTextt.Substring(0, 5) + "······}";
                                            spRoot.AppendChild(xe1);
                                        }
                                        else
                                        {
                                            XmlElement xe1 = xmlDoc.CreateElement("Text");
                                            xe1.InnerText = addPageInfo(pageNum) + "中文摘要正文段落首行缩进错误，应为两个汉字的缩进量:" + "{" + fullTextt + "······}";
                                            spRoot.AppendChild(xe1);
                                        }
                                    }
                                }
                            }*/
                        }
                        if (prpt.GetFirstChild<SpacingBetweenLines>() != null)
                        {
                            /*
                            基本上所有的行距判断都是一样的
                            */
                            if (!Tool.correctSpacingBetweenLines_line(p, count, intlist, list, doc, "1.25", false))
                            {
                                if (fullTextt.Length > 5)
                                {
                                    XmlElement xe1 = xmlDoc.CreateElement("Text");
                                    xe1.InnerText = addPageInfo(pageNum) + "中文摘要正文行距错误，应为多倍行距1.25:" + "{" + fullTextt.Substring(0, 5) + "······}";
                                    spRoot.AppendChild(xe1);
                                    Tool.changeSpacingLine(p, 1.25);
                                }
                                else
                                {
                                    XmlElement xe1 = xmlDoc.CreateElement("Text");
                                    xe1.InnerText = addPageInfo(pageNum) + "中文摘要正文行距错误，应为多倍行距1.25:" + "{" + fullTextt + "······}";
                                    spRoot.AppendChild(xe1);
                                    Tool.changeSpacingLine(p, 1.25);
                                }
                            }
                            /*
                            if (prpt.GetFirstChild<SpacingBetweenLines>().Line != null)
                            {
                                if (prpt.GetFirstChild<SpacingBetweenLines>().Line != "300")
                                {
                                    if (fullTextt != null)
                                    {
                                        if (fullTextt.Length > 5)
                                        {
                                            XmlElement xe1 = xmlDoc.CreateElement("Text");
                                            xe1.InnerText = addPageInfo(pageNum) + "中文摘要正文行距错误，应为多倍行距1.25:" + "{" + fullTextt.Substring(0, 5) + "······}";
                                            spRoot.AppendChild(xe1);
                                            Tool.changeSpacingLine(p, 1.25);
                                        }
                                        else
                                        {
                                            XmlElement xe1 = xmlDoc.CreateElement("Text");
                                            xe1.InnerText = addPageInfo(pageNum) + "中文摘要正文行距错误，应为多倍行距1.25:" + "{" + fullTextt + "······}";
                                            spRoot.AppendChild(xe1);
                                            Tool.changeSpacingLine(p, 1.25);
                                        }
                                    }
                                }
                            }
                            */
                            if (!Tool.correctSpacingBetweenLines_Be(p, count, intlist, list, doc, "0", false))
                            {
                                XmlElement xe1 = xmlDoc.CreateElement("Text");
                                xe1.InnerText = addPageInfo(pageNum) + "中文摘要正文段前间距错误，应为0行:" + "{" + fullTextt.Substring(0, 5) + "······}";
                                spRoot.AppendChild(xe1);
                            }
                            if (!Tool.correctSpacingBetweenLines_Af(p, count, intlist, list, doc, "0", false))
                            {
                                XmlElement xe1 = xmlDoc.CreateElement("Text");
                                xe1.InnerText = addPageInfo(pageNum) + "中文摘要正文段前后距错误，应为0行:" + "{" + fullTextt + "······}";
                                spRoot.AppendChild(xe1);
                            }
                        }
                        if (prPr != null)
                        {
                            /*
                            这里是直接判断的段落属性
                            */
                            if (prPr.GetFirstChild<RunFonts>() != null)
                            {
                                if (prPr.GetFirstChild<RunFonts>().Ascii != null)
                                {
                                    if (prPr.GetFirstChild<RunFonts>().Ascii != "宋体")
                                    {
                                        if (fullTextt.Trim().Length > 5)
                                        {
                                            XmlElement xe1 = xmlDoc.CreateElement("Text");
                                            xe1.InnerText = addPageInfo(pageNum) + "中文摘要正文字体错误，应为宋体:" + "{" + fullTextt.Substring(0, 5) + "······}";
                                            spRoot.AppendChild(xe1);
                                        }
                                        else
                                        {
                                            XmlElement xe1 = xmlDoc.CreateElement("Text");
                                            xe1.InnerText = addPageInfo(pageNum) + "中文摘要正文字体错误，应为宋体:" + "{" + fullTextt + "······}";
                                            spRoot.AppendChild(xe1);
                                        }
                                    }
                                }
                            }
                            if (prPr.GetFirstChild<FontSize>() != null)
                            {
                                if (prPr.GetFirstChild<FontSize>().Val != "24")
                                {
                                    if (fullTextt.Trim().Length > 5)
                                    {
                                        prPr.GetFirstChild<FontSize>().Val = "24";
                                        XmlElement xe1 = xmlDoc.CreateElement("Text");
                                        xe1.InnerText = addPageInfo(pageNum) + "中文摘要正文字号错误，应为小四号:"  + "{" + fullTextt.Substring(0, 5) + "······}";
                                        spRoot.AppendChild(xe1);
                                    }
                                    else
                                    {
                                        prPr.GetFirstChild<FontSize>().Val = "24";
                                        XmlElement xe1 = xmlDoc.CreateElement("Text");
                                        xe1.InnerText = addPageInfo(pageNum) + "中文摘要正文字号错误，应为小四号:"  + "{" + fullTextt + "······}";
                                        spRoot.AppendChild(xe1);
                                    }
                                }
                            }
                        }
                    }
                }
                //关键词检测开始
                if (fullText.Replace(" ", "").Length > 4)
                {
                    if (fullText.IndexOf("关键词") == 0)
                    {
                        ParagraphProperties prpt = p.GetFirstChild<ParagraphProperties>();
                        int pageNumBegin = getPageNum(2,"摘要");
                        int pageNumEnd = getPageNum(pageNumBegin,"关键词：");
                        //所在页码
                        pageNum = getPageNum(pageNum, "关键词：");
                        //判断摘要字数
                        if (pageNumEnd != -1 && pageNumBegin != -1 && pageNumEnd != pageNumBegin)
                        {
                            XmlElement xe1 = xmlDoc.CreateElement("Text");
                            xe1.InnerText = addPageInfo(pageNum) + "中文摘要篇幅超出一页";
                            spRoot.AppendChild(xe1);
                            Tool.addComment(doc, p, "中文摘要篇幅超出一页");
                        }
                        //直接用段落属性判断缩进
                        if (prpt != null)
                        {
                            if (prpt.GetFirstChild<Indentation>() != null)
                            {
                                if (prpt.GetFirstChild<Indentation>().FirstLine != null || Tool.getFullText(p).Substring(0, 1) == " ")
                                {
                                    XmlElement xe1 = xmlDoc.CreateElement("Text");
                                    xe1.InnerText = addPageInfo(pageNum) + "中文摘要关键词缩进错误，不应缩进:" + "{关键词：......}";
                                    spRoot.AppendChild(xe1);
                                    Tool.changeIntent(p, 0);
                                    Tool.deleteSpacingChar(p);
                                }
                            }
                        }
                        //用Contains方法判断是否使用启发分隔符
                        if (!fullText.Contains("；"))
                        {
                            XmlElement xe1 = xmlDoc.CreateElement("Text");
                            xe1.InnerText = addPageInfo(pageNum) + "中文摘要关键词应使用中文分号";
                            spRoot.AppendChild(xe1);
                            Tool.addComment(doc, p, "中文摘要关键词应使用中文分号");
                        }
                        /*
                      这一部分是将关键词内容完整取出，然后靠用分号分卡是每一个关键词
                      */
                        IEnumerable<Run> krList = p.Elements<Run>();
                        String fullKeyWords = Tool.getFullText(p);
                        String[] array = fullKeyWords.Split('；');
                        String[] array1 = fullKeyWords.Split(';');
                        bool totalChn = true;
                        if (fullKeyWords.Contains(";"))
                        {
                            totalChn = false;
                            XmlElement xe1 = xmlDoc.CreateElement("Text");
                            xe1.InnerText = addPageInfo(pageNum) + "中文摘要关键词应使用中文分号";
                            spRoot.AppendChild(xe1);
                            Tool.addComment(doc, p, "中文摘要关键词应使用中文分号");
                        }
                        int i = 0;
                        if (totalChn == false)
                        {
                            foreach (string str in array1)
                            {
                                i++;
                            }
                        }
                        else
                        {
                            foreach (string str in array)
                            {
                                i++;
                            }
                        }
                        if (i < 2 || i > 5)
                        {
                            XmlElement xe1 = xmlDoc.CreateElement("Text");
                            xe1.InnerText = addPageInfo(pageNum) + "中文摘要关键词应不少于3个且不多于5个";
                            spRoot.AppendChild(xe1);
                            Tool.addComment(doc, p, "中文摘要关键词应不少于3个且不多于5个");
                        }
                        foreach (String str in array)
                        {
                            if (str.Length > 1)
                            {
                                if (str.Substring(0, 1) == " ")
                                {
                                    XmlElement xe1 = xmlDoc.CreateElement("Text");
                                    xe1.InnerText = addPageInfo(pageNum) + "中文摘要关键词前不应空格：{" + str + "}";
                                    spRoot.AppendChild(xe1);
                                    Tool.addComment(doc, p, "中文摘要关键词前不应空格");
                                }
                            }
                        }
                        foreach (Run kr in krList)
                        {
                            RunProperties krpr = kr.GetFirstChild<RunProperties>();
                            Text krText = kr.GetFirstChild<Text>();
                            if (krText != null)
                            {
                                String str = krText.Text.ToString();
                                if (krpr != null)
                                {
                                    if (str.Replace(" ", "") != "")
                                    {
                                        if (str != "关键词：")
                                        {
                                            if (krpr.GetFirstChild<Bold>() != null && krpr.GetFirstChild<Bold>().Val != "0")
                                            {
                                                XmlElement xe1 = xmlDoc.CreateElement("Text");
                                                xe1.InnerText = addPageInfo(pageNum) + "中文摘要关键词内容不应加粗:" + "{" + str + "}";
                                                spRoot.AppendChild(xe1);

                                                if (krpr.GetFirstChild<BoldComplexScript>() != null)
                                                {
                                                    krpr.RemoveAllChildren<Bold>();
                                                    krpr.RemoveAllChildren<BoldComplexScript>();
                                                    Bold b = new Bold() { Val = false };
                                                    BoldComplexScript bcs = new BoldComplexScript();
                                                    krpr.AppendChild<Bold>(b);
                                                    krpr.AppendChild<BoldComplexScript>(bcs);
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        if (Tool.correctfonts(p, doc, "黑体", "仿宋_GB2312") == false)
                        {
                            XmlElement xe1 = xmlDoc.CreateElement("Text");
                            xe1.InnerText = addPageInfo(pageNum) + "中文摘要关键词字体错误，“关键词”应为黑体，关键词内容应为仿宋_GB2312";
                            spRoot.AppendChild(xe1);
                            IEnumerable<Run> runs = p.Elements<Run>();
                            int maohao = 0;
                            foreach (Run rr in runs)
                            {
                                maohao++;
                                string t = "";
                                if (rr.GetFirstChild<Text>() != null)
                                    t = rr.GetFirstChild<Text>().Text;
                                if (t.Contains("："))
                                    break;
                            }
                            foreach (Run rr in runs)
                            {
                                if ((maohao--) != 0)
                                {
                                    if (rr.RunProperties != null)
                                    {
                                        if (rr.RunProperties.GetFirstChild<RunFonts>() != null)
                                        {
                                            //字体改，此处要分清要改的是其中的属性，还是其中的内容
                                            rr.RunProperties.GetFirstChild<RunFonts>().Hint = FontTypeHintValues.EastAsia;
                                            rr.RunProperties.GetFirstChild<RunFonts>().Ascii = "黑体";
                                            rr.RunProperties.GetFirstChild<RunFonts>().EastAsia = "黑体";
                                        }
                                        else
                                        {
                                            //添加标签
                                            RunFonts rfont = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "黑体", HighAnsi = "黑体", EastAsia = "黑体" };
                                            rr.RunProperties.AppendChild<RunFonts>(rfont);
                                        }
                                    }
                                }
                                else
                                {
                                    if (rr.RunProperties.GetFirstChild<RunFonts>() != null)
                                    {
                                        if (rr.RunProperties.GetFirstChild<RunFonts>() != null)
                                        {
                                            //字体改，此处要分清要改的是其中的属性，还是其中的内容
                                            rr.RunProperties.GetFirstChild<RunFonts>().Hint = FontTypeHintValues.EastAsia;
                                            rr.RunProperties.GetFirstChild<RunFonts>().Ascii = "仿宋_GB2312";
                                            rr.RunProperties.GetFirstChild<RunFonts>().EastAsia = "仿宋_GB2312";
                                        }
                                        else
                                        {
                                            //添加标签
                                            RunFonts rfont = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "仿宋_GB2312", HighAnsi = "仿宋_GB2312", EastAsia = "仿宋_GB2312" };
                                            r.RunProperties.AppendChild<RunFonts>(rfont);
                                        }
                                    }
                                }
                            }
                        }
                        /*
                        这里是查看是否使用了样式
                        */
                        /*foreach (Run kr in krList)
                        {
                            RunProperties krpr = kr.GetFirstChild<RunProperties>();
                            Text krText = kr.GetFirstChild<Text>();
                            if (krText != null)
                            {                               
                                String str = krText.Text.ToString();
                                if (flag)
                                {
                                    if (krpr != null)
                                    {
                                        if (krpr.GetFirstChild<RunStyle>() != null)
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
                                                    //此处获取样式ID，并且和style.xml中的样式id一致的
                                                    if (m.StyleId.ToString() == krpr.GetFirstChild<RunStyle>().Val)
                                                    {
                                                        if (srPr != null)
                                                        {
                                                            if (srPr.RunFonts != null)
                                                            {
                                                                if (srPr.RunFonts.Ascii != "黑体")
                                                                {
                                                                    XmlElement xe1 = xmlDoc.CreateElement("Text");
                                                                    xe1.InnerText = addPageInfo(pageNum) + "中文摘要关键词字体错误，错误部分为：" + "{" + str + "}";
                                                                    spRoot.AppendChild(xe1);
                                                                }
                                                            }
                                                            if (srPr.FontSize != null)
                                                            {
                                                                if (srPr.FontSize.Val != "24")
                                                                {
                                                                    XmlElement xe1 = xmlDoc.CreateElement("Text");
                                                                    xe1.InnerText = addPageInfo(pageNum) + "中文摘要关键词字号错误，错误部分为：" + "{" + str + "}";
                                                                    spRoot.AppendChild(xe1);
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                                //flag用来刚表示关键词段落是否结束
                                if (flag)
                                {
                                    if (str.Contains("关键词"))
                                    {
                                        flag = false;
                                        continue;
                                    }
                                    else
                                    {
                                        continue;
                                    }
                                }
                                if (krpr != null)
                                {
                                    if (krpr.GetFirstChild<Bold>() != null)
                                    {
                                        if (str.Replace(" ", "") != "" && str != "：")
                                        {
                                            XmlElement xe1 = xmlDoc.CreateElement("Text");
                                            xe1.InnerText = addPageInfo(pageNum) + "中文摘要关键词内容不应加粗：" + "{" + str + "}";
                                            spRoot.AppendChild(xe1);
                                        }

                                    }
                                    /*
                                   这里也是将整个段落的run取出，然后判断run里的text是否是关键词，还是关键词的内容，然后再判断字体字号
                                   */
                                    /*if (krpr.GetFirstChild<RunFonts>() != null)
                                    {
                                        if (krpr.GetFirstChild<RunFonts>().Ascii != null)
                                        {
                                            if (krpr.GetFirstChild<RunFonts>().Ascii != "仿宋_GB2312" && krpr.GetFirstChild<RunFonts>().Ascii != "仿宋")
                                            {
                                                //这里是为了防止检测到‘关键词：’这几个字，因为只检测的是关键词内容
                                                if (!str.Contains("关键词") && !str.Contains("关键词：") && str != "：")
                                                {
                                                    if (str.Replace(" ", "") != "")
                                                    {
                                                        XmlElement xe1 = xmlDoc.CreateElement("Text");
                                                        xe1.InnerText = addPageInfo(pageNum) + "中文摘要关键词内容字体错误，应为仿宋，错误部分为：" + "{" + str + "}";
                                                        spRoot.AppendChild(xe1);
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }

                        }*/
                        if (ss != 1)
                        {
                            XmlElement xe1 = xmlDoc.CreateElement("Text");
                            xe1.InnerText = addPageInfo(pageNum) + "中文摘要关键词与正文之间应有且只有一行空行";
                            spRoot.AppendChild(xe1);
                        }
                        xmlDoc.Save(xmlFullPath);
                        break;
                    }
                }
            }
        }
        /*
        英文摘要的论文英文题目部分检测
        */
        private void getENAbstractTitleXml(WordprocessingDocument doc, String xmlFullPath)
        {
            Body body = doc.MainDocumentPart.Document.Body;
            MainDocumentPart mainPart = doc.MainDocumentPart;
            IEnumerable<Paragraph> paras = body.Elements<Paragraph>();
            ParagraphProperties pPr = null;
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(xmlFullPath);
            XmlNode spRoot = xmlDoc.SelectSingleNode("Abstract/spErroInfo");
            int count = 0;
            int position = 0;
            bool hasAbstract1 = false;
            bool hasAbstract2 = false;
            String enTitle = "";

            int count2 = -1;
            List<int> intlist = Tool.secPrListInt(doc.MainDocumentPart.Document.Body);
            List<SectionProperties> list = Tool.secPrList(doc.MainDocumentPart.Document.Body);
            foreach (Paragraph p in paras)
            {
                count2++;
                Run r = p.GetFirstChild<Run>();
                if (r == null) continue;
                String fullText = Tool.getFullText(p).Trim();
                if (fullText != null)
                {
                    position++;
                }
                //遇到Abstract，则hasAbstract标志位真，用于后续的正文检测
                if (fullText.Replace(" ", "") == "Abstract")
                {
                    hasAbstract1 = true;
                    break;
                }
                //两个判断因为有的用户或全部大写，或大小写不规范，导致检测不到后续正文
                else if (fullText.Replace(" ", "").ToLower() == "abstract")
                {
                    hasAbstract2 = true;
                    break;
                }
            }
            //标志为假的提示
            if (!hasAbstract1)
            {
                XmlElement xe1 = xmlDoc.CreateElement("Text");
                xe1.InnerText = "英文摘要缺少'Abstract'字样或书写错误";
                spRoot.AppendChild(xe1);
                xmlDoc.Save(xmlFullPath);
            }
            foreach (Paragraph p in paras)
            {
                Run r = p.GetFirstChild<Run>();
                if (r == null) continue;
                String fullText = Tool.getFullText(p).Trim();
                //页码信息
                pageNum = getPageNum(pageNum, fullText);
                bool enAbtittle = false;
                if (fullText != null)
                {
                    count++;
                }
                if (count == 3)
                {
                    enTitle = fullText.Trim(); 
                }
                if (count == position - 1 && (hasAbstract1 || hasAbstract2))
                {
                    //英文摘要上方应该是论文英文名字，这里就是去上一个有文本的段落判断，enTitle就是论文英文名
                    if (fullText.Trim() != enTitle)
                    {
                        XmlElement xe1 = xmlDoc.CreateElement("Text");
                        xe1.InnerText = addPageInfo(pageNum) + "英文摘要上方缺少论文英文题目，或者与封面的英文标题书写不一致";
                        spRoot.AppendChild(xe1);
                        xmlDoc.Save(xmlFullPath);
                       
                        /*找到下一段落*/
                        Paragraph pt = p.NextSibling<Paragraph>();
                        Run rt = pt.GetFirstChild<Run>();
                        while(pt == null || rt == null)
                        {
                            pt = pt.NextSibling<Paragraph>();
                            rt = pt.GetFirstChild<Run>();
                        }
                        Tool.addComment(doc, pt, "英文摘要上方缺少论文英文题目，或者与封面的英文标题书写不一致");
                        
                    }
                    else
                    {
                        //剩下的部分的检测个中文基本一样
                        IEnumerable<Run> pRunList = p.Elements<Run>();
                        pPr = p.GetFirstChild<ParagraphProperties>();
                        enAbtittle = true;
                        bool flagA = false;
                        bool flagB = false;
                        if (pPr != null)
                        {
                            ParagraphMarkRunProperties prPr = pPr.GetFirstChild<ParagraphMarkRunProperties>();
                            //居中
                            if (pPr.GetFirstChild<Justification>() != null)
                            {
                                if (pPr.GetFirstChild<Justification>().Val.Value.ToString().ToLower() != "center")
                                {
                                    XmlElement xe1 = xmlDoc.CreateElement("Text");
                                    xe1.InnerText = addPageInfo(pageNum) + "英文摘要部分论文英文题目未居中";
                                    spRoot.AppendChild(xe1);
                                    Tool.change_center(p);
                                }
                            }
                            else
                            {
                                XmlElement xe1 = xmlDoc.CreateElement("Text");
                                xe1.InnerText = addPageInfo(pageNum) + "英文摘要部分论文英文题目未居中";
                                spRoot.AppendChild(xe1);
                                Tool.change_center(p);
                            }
                            //行距判断
                            if (!Tool.correctSpacingBetweenLines_line(p, count2, intlist, list, doc, "1.25", false))
                            {
                                XmlElement xe1 = xmlDoc.CreateElement("Text");
                                xe1.InnerText = addPageInfo(pageNum) + "英文摘要部分论文英文题目行距错误，应为多倍行距1.25倍行距";
                                spRoot.AppendChild(xe1);
                                Tool.changeSpacingLine(p, 1.25);
                            }
                            //段前间距判断
                            if (!Tool.correctSpacingBetweenLines_Be(p, count2, intlist, list, doc, "0", false))
                            {
                                XmlElement xe1 = xmlDoc.CreateElement("Text");
                                xe1.InnerText = addPageInfo(pageNum) + "英文摘要部分论文英文题目段前间距错误，应为0";
                                spRoot.AppendChild(xe1);
                                Tool.changeSpacingBefore(p, 0);
                            }

                            //段后间距判断
                            if (!Tool.correctSpacingBetweenLines_Af(p, count2, intlist, list, doc, "0", false))
                            {
                                XmlElement xe1 = xmlDoc.CreateElement("Text");
                                xe1.InnerText = addPageInfo(pageNum) + "英文摘要部分论文英文题目段后间距错误，应为0行";
                                spRoot.AppendChild(xe1);
                                Tool.changeSpacingAfter(p, 0);
                            }

                            if (prPr != null)
                            {
                                if (prPr.GetFirstChild<Bold>() == null || prPr.GetFirstChild<Bold>().Val == "0")
                                {
                                    XmlElement xe1 = xmlDoc.CreateElement("Text");
                                    xe1.InnerText = addPageInfo(pageNum) + "英文摘要部分论文英文题目未加粗";
                                }
                            }
                            /*更改加粗*/
                            foreach(Run rr in pRunList)
                            {
                                if (rr.GetFirstChild<RunProperties>() != null)
                                {
                                    RunProperties rp = rr.GetFirstChild<RunProperties>();
                                    if (rp.GetFirstChild<Bold>() == null)
                                    {
                                        if (rp.GetFirstChild<BoldComplexScript>() == null)
                                        {
                                            Bold b = new Bold();
                                            BoldComplexScript bcs = new BoldComplexScript() { Val = false };
                                            rp.AppendChild<Bold>(b);
                                            rp.AppendChild<BoldComplexScript>(bcs);
                                        }
                                    }
                                    else
                                    {
                                        if (rp.GetFirstChild<BoldComplexScript>() != null)
                                        {
                                            rp.RemoveAllChildren<Bold>();
                                            rp.RemoveAllChildren<BoldComplexScript>();
                                            Bold b = new Bold();
                                            BoldComplexScript bcs = new BoldComplexScript() { Val = false };
                                            rp.AppendChild<Bold>(b);
                                            rp.AppendChild<BoldComplexScript>(bcs);
                                        }
                                    }
                                }
                            }
                            /*更改加粗结束*/
                        }
                        if (pRunList != null)
                        {
                            bool flag1 = true;
                            bool flag2 = true;
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
                                                if (Rrpr.GetFirstChild<FontSize>().Val != "30")
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
                                flagA = true;
                                XmlElement xe1 = xmlDoc.CreateElement("Text");
                                xe1.InnerText = "英文摘要部分论文英文题目字体错误，应为Times New Roman";
                                spRoot.AppendChild(xe1);
                                Tool.change_rfonts(p, "Times New Roman");
                            }
                            if (!flag2)
                            {
                                flagB = true;
                                XmlElement xe1 = xmlDoc.CreateElement("Text");
                                xe1.InnerText = "英文摘要部分论文英文题目字号错误，应为小三号";
                                spRoot.AppendChild(xe1);
                                Tool.change_fontsize(p, "30");
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
                                    StyleRunProperties srPr = m.StyleRunProperties;
                                    StyleParagraphProperties spPr = m.StyleParagraphProperties;
                                    if (m.StyleId.ToString() == id)
                                    {
                                        if (srPr != null)
                                        {
                                            if (srPr.RunFonts != null && !flagA)
                                            {
                                                if (srPr.RunFonts.Ascii != null)
                                                {
                                                    if (srPr.RunFonts.Ascii != "Times New Roman")
                                                    {
                                                        XmlElement xe1 = xmlDoc.CreateElement("Text");
                                                        xe1.InnerText = "英文摘要部分论文英文题目字体错误，应为Times New Roman";
                                                        spRoot.AppendChild(xe1);
                                                        Tool.change_rfonts(p, "Times New Roman");
                                                    }
                                                }
                                            }
                                            if (srPr.FontSize != null && !flagB)
                                            {
                                                if (srPr.FontSize.Val != null)
                                                {
                                                    if (srPr.FontSize.Val != "30")
                                                    {
                                                        XmlElement xe1 = xmlDoc.CreateElement("Text");
                                                        xe1.InnerText = "英文摘要部分论文英文题目字号错误，应为小三号";
                                                        spRoot.AppendChild(xe1);
                                                        Tool.change_fontsize(p, "30");
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        xmlDoc.Save(xmlFullPath);
                        continue;
                    }
                    if (enAbtittle)
                    {
                        Paragraph pt = p.NextSibling<Paragraph>();
                        String text = Tool.getFullText(pt).Trim();
                        if (text != null)
                        {
                            XmlElement xe1 = xmlDoc.CreateElement("Text");
                            xe1.InnerText = "英文摘要部分论文英文题目与“Abstract”之间应空一行";
                            spRoot.AppendChild(xe1);
                            xmlDoc.Save(xmlFullPath);
                            Tool.addComment(doc, p, "英文摘要部分论文英文题目与“Abstract”之间应空一行");
                        }
                    }
                }
                
            }
        }
        /*
       英文摘要的论文部分检测（其实代码和中文摘要的差不多）
       */
        private void getENAbstractXml(WordprocessingDocument doc, String xmlFullPath)
        {
            Body body = doc.MainDocumentPart.Document.Body;
            MainDocumentPart mainPart = doc.MainDocumentPart;
            IEnumerable<Paragraph> paras = body.Elements<Paragraph>();
            ParagraphProperties pPr = null;
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(xmlFullPath);
            XmlNode root = xmlDoc.SelectSingleNode("Abstract/CNAbstractTitle");
            XmlNode spRoot = xmlDoc.SelectSingleNode("Abstract/spErroInfo");
            bool isAbsTittle = false;
            int ss = countSpaceLineEN(doc);

            List<int> intlist = Tool.secPrListInt(body);//索引，用于端前后间距判断传参
            List<SectionProperties> list = Tool.secPrList(body);
            int count = -1;
            foreach (Paragraph p in paras)
            {
                count++;
                Run r = p.GetFirstChild<Run>();
                if (r == null) continue;
                String fullText = Tool.getFullText(p).Trim();
                //所在页码
                pageNum = getPageNum(pageNum, fullText);
                if (fullText.Replace(" ", "").ToLower() == "abstract")
                {
                    isAbsTittle = true;
                    IEnumerable<Run> pRunList = p.Elements<Run>();
                    if (pPr != null)
                    {
                        //对齐方式
                        if (pPr.GetFirstChild<Justification>() != null)
                        {
                            XmlElement xe1 = xmlDoc.CreateElement("jc");
                            xe1.InnerText = pPr.GetFirstChild<Justification>().Val.Value.ToString().ToLower();
                            root.AppendChild(xe1);
                        }
                    }
                    pPr = p.GetFirstChild<ParagraphProperties>();
                    //居中判断
                    if (pPr != null)
                    {
                        if (pPr.GetFirstChild<Justification>() != null)
                        {
                            if (pPr.GetFirstChild<Justification>().Val.ToString().ToLower() != "center")
                            {
                                XmlElement xe1 = xmlDoc.CreateElement("Text");
                                xe1.InnerText = addPageInfo(pageNum) + "英文摘要标题未居中";
                                spRoot.AppendChild(xe1);
                                Tool.change_center(p);
                            }
                        }
                        else
                        {
                            XmlElement xe1 = xmlDoc.CreateElement("Text");
                            xe1.InnerText = addPageInfo(pageNum) + "英文摘要部分论文英文题目未居中";
                            spRoot.AppendChild(xe1);
                            Tool.change_center(p);
                        }
                        //行距判断
                        if (!Tool.correctSpacingBetweenLines_line(p, count, intlist, list, doc, "1.5", false))
                        {
                            XmlElement xe1 = xmlDoc.CreateElement("Text");
                            xe1.InnerText = addPageInfo(pageNum) + "英文摘要标题行距错误，应为1.5倍行距";
                            spRoot.AppendChild(xe1);
                            Tool.changeSpacingLine(p, 1.5);
                        }

                        //段前间距判断
                        if (!Tool.correctSpacingBetweenLines_Be(p, count, intlist, list, doc, "0", false))
                        {
                            XmlElement xe1 = xmlDoc.CreateElement("Text");
                            xe1.InnerText = addPageInfo(pageNum) + "英文摘要标题段前间距错误，应为0行";
                            spRoot.AppendChild(xe1);
                            Tool.changeSpacingBefore(p, 0);
                        }

                        //段后间距判断
                        if (!Tool.correctSpacingBetweenLines_Af(p, count, intlist, list, doc, "11", true))
                        {
                            XmlElement xe1 = xmlDoc.CreateElement("Text");
                            xe1.InnerText = addPageInfo(pageNum) + "英文摘要标题段后间距错误，应为11磅";
                            spRoot.AppendChild(xe1);
                            Tool.changeSpacingAfter_point(p, 11);
                        }
                    }
                    //这一部分是将一个段落的run全部取出，判断每一个run里的设置是否都是一致的
                    if (pRunList != null)
                    {
                        bool flag1 = true;
                        bool flag2 = true;
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
                                                Tool.change_rfonts(p, "Times New Roman");
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
                                                Tool.change_fontsize(p, "30");
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        if (!flag1)
                        {
                            XmlElement xe1 = xmlDoc.CreateElement("Text");
                            xe1.InnerText = addPageInfo(pageNum) + "英文摘要“Abstract”字体错误，应为Times New Roman";
                            spRoot.AppendChild(xe1);
                            Tool.change_rfonts(p, "Times New Roman");
                        }
                        if (!flag2)
                        {
                            XmlElement xe1 = xmlDoc.CreateElement("Text");
                            xe1.InnerText = addPageInfo(pageNum) +"英文摘要“Abstract”字号错误，应为小三";
                            spRoot.AppendChild(xe1);
                            Tool.change_fontsize(p, "30");
                        }
                    }
                    xmlDoc.Save(xmlFullPath);
                    continue;
                }
                if(fullText.Trim().Replace(" ", "").Length > 8)
                {
                    if (fullText.Trim().ToLower().Replace(" ", "").Substring(0, 8) == "keywords")
                    {
                        isAbsTittle = false;
                    }
                }                
                if (isAbsTittle)
                {
                    IEnumerable<Run> pRunList = p.Elements<Run>();
                    String fullTextt = Tool.getFullText(p).Trim();

                    /*更改缩进*/
                    Tool.changeIntent(p, 2);
                    Tool.deleteSpacingChar(p);
                    /*更改段前段后间距*/
                    Tool.changeSpacingBefore(p, 0);
                    Tool.changeSpacingAfter(p, 0);
                    /*更改字体、字号*/
                    Tool.change_rfonts(p, "Times New Roman");
                    Tool.change_fontsize(p, "24");

                    ParagraphProperties prpt = p.GetFirstChild<ParagraphProperties>();
                    if (prpt != null)
                    {
                        ParagraphMarkRunProperties prPr = prpt.GetFirstChild<ParagraphMarkRunProperties>();
                        
                        if (prpt.GetFirstChild<Indentation>() != null)
                        {
                            if (prpt.GetFirstChild<Indentation>().FirstLine == null || prpt.GetFirstChild<Indentation>().FirstLine.Value == "0")
                            {
                                if (Tool.getFullText(p).Length > 4 && Tool.getFullText(p).Substring(0, 3) == "    ")
                                {
                                }
                                else
                                {
                                    if (fullTextt.Length > 5)
                                    {
                                        XmlElement xe1 = xmlDoc.CreateElement("Text");
                                        xe1.InnerText = addPageInfo(pageNum) + "英文摘要正文段落首行缩进错误，应为两个汉字的缩进量:" + "{" + fullTextt.Substring(0, 5) + "......}";
                                        spRoot.AppendChild(xe1);
                                    }
                                    else
                                    {
                                        XmlElement xe1 = xmlDoc.CreateElement("Text");
                                        xe1.InnerText = addPageInfo(pageNum) + "英文摘要正文段落首行缩进错误，应为两个汉字的缩进量:" + "{" + fullTextt + "......}";
                                        spRoot.AppendChild(xe1);
                                    }
                                }
                            }
                            else
                            {
                                if (prpt.GetFirstChild<Indentation>().FirstLineChars == "200")
                                {
                                }
                                else
                                {
                                    if (fullTextt.Length > 5)
                                    {
                                        XmlElement xe1 = xmlDoc.CreateElement("Text");
                                        xe1.InnerText = addPageInfo(pageNum) + "英文摘要正文段落首行缩进错误，应为两个汉字的缩进量:" + "{" + fullTextt.Substring(0, 5) + "......}";
                                        spRoot.AppendChild(xe1);
                                    }
                                    else
                                    {
                                        XmlElement xe1 = xmlDoc.CreateElement("Text");
                                        xe1.InnerText = addPageInfo(pageNum) + "英文摘要正文段落首行缩进错误，应为两个汉字的缩进量:" + "{" + fullTextt + "......}";
                                        spRoot.AppendChild(xe1);
                                    }
                                }
                            }
                            /*
                            if (prpt.GetFirstChild<Indentation>().FirstLine == "480")
                            {
                                if (Tool.getFullText(p).Length > 0 && Tool.getFullText(p).Replace(" ","")!="")
                                {
                                    if (Tool.getFullText(p).Substring(0, 1) == " ")
                                    {
                                        if (fullTextt.Length > 20)
                                        {
                                            XmlElement xe1 = xmlDoc.CreateElement("Text");
                                            xe1.InnerText = addPageInfo(pageNum) + "英文摘要正文部分缩进错误，应为两个汉字的缩进量:" + "{" + fullTextt.Substring(0, 20) + "······}";
                                            spRoot.AppendChild(xe1);
                                        }
                                        else
                                        {
                                            XmlElement xe1 = xmlDoc.CreateElement("Text");
                                            xe1.InnerText = addPageInfo(pageNum) +"英文摘要正文部分缩进错误，应为两个汉字的缩进量:" + "{" + fullTextt + "······}";
                                            spRoot.AppendChild(xe1);
                                        }

                                    }
                                }
                            }
                            else if (prpt.GetFirstChild<Indentation>().FirstLine == "240")
                            {
                                if (Tool.getFullText(p).Length > 1 && Tool.getFullText(p).Replace(" ", "") != "")
                                {
                                    if (Tool.getFullText(p).Substring(0, 1) == " ")
                                    {
                                    }
                                    else
                                    {
                                        if (fullTextt.Length >20)
                                        {
                                            XmlElement xe1 = xmlDoc.CreateElement("Text");
                                            xe1.InnerText = addPageInfo(pageNum) + "英文摘要正文段落首行缩进错误，应为两个汉字的缩进量:" + "{" + fullTextt.Substring(0, 20) + "······}";
                                            spRoot.AppendChild(xe1);
                                        }
                                        else
                                        {
                                            XmlElement xe1 = xmlDoc.CreateElement("Text");
                                            xe1.InnerText = addPageInfo(pageNum) + "英文摘要正文段落首行缩进错误，应为两个汉字的缩进量:" + "{" + fullTextt + "······}";
                                            spRoot.AppendChild(xe1);
                                        }
                                    }
                                }
                            }
                            else
                            {
                                if (Tool.getFullText(p).Length > 2 && Tool.getFullText(p).Replace(" ", "") != "")
                                {
                                    if (Tool.getFullText(p).Substring(0, 2) == "  " )
                                    {
                                        
                                    }
                                    else if (Tool.getFullText(p).Substring(0, 2) == null)
                                    { }
                                    else
                                    {

                                        if (fullTextt.Length > 20)
                                        {
                                            XmlElement xe1 = xmlDoc.CreateElement("Text");
                                            xe1.InnerText = addPageInfo(pageNum) + "英文摘要正文段落首行缩进错误，应为两个汉字的缩进量:" + "{" + fullTextt.Substring(0, 20) + "······}";
                                            spRoot.AppendChild(xe1);
                                        }
                                        else
                                        {
                                            XmlElement xe1 = xmlDoc.CreateElement("Text");
                                            xe1.InnerText = addPageInfo(pageNum) + "英文摘要正文段落首行缩进错误，应为两个汉字的缩进量:" + "{" + fullTextt + "······}";
                                            spRoot.AppendChild(xe1);
                                        }
                                    }
                                }
                            }
                            */
                        }
                        if (!Tool.correctSpacingBetweenLines_line(p, count, intlist, list, doc, "1.25", false))
                        {
                            if (fullTextt.Length > 5)
                            {
                                XmlElement xe1 = xmlDoc.CreateElement("Text");
                                xe1.InnerText = addPageInfo(pageNum) + "英文摘要正文行距错误，应为多倍行距1.25:" + "{" + fullTextt.Substring(0, 5) + "······}";
                                spRoot.AppendChild(xe1);
                                Tool.changeSpacingLine(p, 1.25);
                            }
                            else
                            {
                                XmlElement xe1 = xmlDoc.CreateElement("Text");
                                xe1.InnerText = addPageInfo(pageNum) + "英文摘要正文行距错误，应为多倍行距1.25:" + "{" + fullTextt + "······}";
                                spRoot.AppendChild(xe1);
                                Tool.changeSpacingLine(p, 1.25);
                            }
                        }
                        if (!Tool.correctSpacingBetweenLines_Be(p, count, intlist, list, doc, "0", false))
                        {
                            if (fullTextt.Length > 20)
                            {
                                XmlElement xe1 = xmlDoc.CreateElement("Text");
                                xe1.InnerText = addPageInfo(pageNum) + "英文摘要正文段前间距错误，应为0行:" + "{" + fullTextt.Substring(0, 20) + "······}";
                                spRoot.AppendChild(xe1);
                            }
                            else
                            {
                                XmlElement xe1 = xmlDoc.CreateElement("Text");
                                xe1.InnerText = addPageInfo(pageNum) + "英文摘要正文段前间距错误，应为0行:" + "{" + fullTextt + "······}";
                                spRoot.AppendChild(xe1);
                            }
                        }
                        if (!Tool.correctSpacingBetweenLines_Af(p, count, intlist, list, doc, "0", false))
                        {
                            if (fullTextt.Length > 20)
                            {
                                XmlElement xe1 = xmlDoc.CreateElement("Text");
                                xe1.InnerText = addPageInfo(pageNum) + "英文摘要正文段后间距错误，应为0行:" + "{" + fullTextt.Substring(0, 20) + "······}";
                                spRoot.AppendChild(xe1);
                            }
                            else
                            {
                                XmlElement xe1 = xmlDoc.CreateElement("Text");
                                xe1.InnerText = addPageInfo(pageNum) + "英文摘要正文段后间距错误，应为0行:" + "{" + fullTextt + "······}";
                                spRoot.AppendChild(xe1);
                            }
                        }
                        
                        if (prPr != null)
                        {
                            if (prPr.GetFirstChild<RunFonts>() != null)
                            {
                                if (prPr.GetFirstChild<RunFonts>().Ascii != null)
                                {
                                                                     
                                    if (prPr.GetFirstChild<RunFonts>().Ascii != "Times New Roman")
                                    {
                                        if (fullTextt.Length > 20)
                                        {
                                            XmlElement xe1 = xmlDoc.CreateElement("Text");
                                            xe1.InnerText = addPageInfo(pageNum) + "英文摘要正文字体错误，应为Times New Roman:" + "{" + fullTextt.Substring(0, 20) + "······}";
                                            spRoot.AppendChild(xe1);
                                        }
                                        else
                                        {
                                            XmlElement xe1 = xmlDoc.CreateElement("Text");
                                            xe1.InnerText = addPageInfo(pageNum) + "英文摘要正文字体错误，应为Times New Roman:" + "{" + fullTextt + "······}";
                                            spRoot.AppendChild(xe1);
                                        }
                                    }
                                }
                            }
                            if (prPr.GetFirstChild<FontSize>() != null)
                            {
                                if (prPr.GetFirstChild<FontSize>().Val != "24")
                                {
                                    if (fullTextt.Length > 20)
                                    {
                                        XmlElement xe1 = xmlDoc.CreateElement("Text");
                                        xe1.InnerText = addPageInfo(pageNum) + "英文摘要正文字号错误，应为小四号:" + "{" + fullTextt.Substring(0, 20) + "······}";
                                        spRoot.AppendChild(xe1);
                                    }
                                    else
                                    {
                                        XmlElement xe1 = xmlDoc.CreateElement("Text");
                                        xe1.InnerText = addPageInfo(pageNum) + "英文摘要正文字号错误，应为小四号:" + "{" + fullTextt + "······}";
                                        spRoot.AppendChild(xe1);
                                    }
                                }
                            }
                        }
                    }
                }

                //关键词
                if (fullText.Trim().Replace(" ", "").Length > 8)
                {
                    if(fullText.Trim().ToLower().Replace(" ","").Substring(0,8)=="keywords")
                    {
                        IEnumerable<Run> krList = p.Elements<Run>();
                        String fullKeyWords = Tool.getFullText(p);
                        String[] array = fullKeyWords.Split('；');
                        String[] array1 = fullKeyWords.Split(';');
                        ParagraphProperties prpt = p.GetFirstChild<ParagraphProperties>();

                        bool totalEN = true;
                        if (!fullKeyWords.Contains("；"))
                        {
                            XmlElement xe1 = xmlDoc.CreateElement("Text");
                            xe1.InnerText = addPageInfo(pageNum) + "英文摘要关键词应使用中文分号";
                            spRoot.AppendChild(xe1);
                            Tool.addComment(doc, p, "英文摘要关键词应使用中文分号");
                        }
                        if (fullKeyWords.Contains(";"))
                        {
                            totalEN = false;
                            XmlElement xe1 = xmlDoc.CreateElement("Text");
                            xe1.InnerText = addPageInfo(pageNum) + "英文摘要关键词应使用中文分号";
                            spRoot.AppendChild(xe1);
                            Tool.addComment(doc, p, "英文摘要关键词应使用中文分号");
                        }
                        int i = 0;
                        if (totalEN == false)
                        {
                            foreach (string str in array1)
                            {
                                i++;
                            }
                        }
                        else
                        {
                            foreach (string str in array)
                            {
                                i++;
                            }
                        }
                        if (i < 2 || i > 5)
                        {
                            XmlElement xe1 = xmlDoc.CreateElement("Text");
                            xe1.InnerText = addPageInfo(pageNum) + "英文摘要关键词应不少于3个且不多于5个";
                            spRoot.AppendChild(xe1);
                            Tool.addComment(doc, p, "英文摘要关键词应不少于3个且不多于5个");
                        }
                        if (prpt != null)
                        {
                            if (prpt.GetFirstChild<Indentation>() != null)
                            {
                                if (prpt.GetFirstChild<Indentation>().FirstLine != null || Tool.getFullText(p).Substring(0, 1) == " ")
                                {
                                    XmlElement xe1 = xmlDoc.CreateElement("Text");
                                    xe1.InnerText = addPageInfo(pageNum) + "英文摘要关键词缩进错误，不应缩进:" + "{关键词：......}";
                                    spRoot.AppendChild(xe1);
                                    Tool.changeIntent(p, 0);
                                    Tool.deleteSpacingChar(p);
                                }
                            }
                        }
                        foreach (String str in array)
                        {
                            if (str.Length > 1)
                            {
                                if (str.Substring(0, 1) == " ")
                                {
                                    XmlElement xe1 = xmlDoc.CreateElement("Text");
                                    xe1.InnerText = addPageInfo(pageNum) + "英文摘要关键词前不应空格:{" + str + "}";
                                    spRoot.AppendChild(xe1);
                                    Tool.addComment(doc, p, "英文摘要关键词前不应空格");
                                }
                            }
                        }
                        foreach (String str in array1)
                        {
                            if (str.Length > 1)
                            {
                                if (str.Substring(0, 1) == " ")
                                {
                                    XmlElement xe1 = xmlDoc.CreateElement("Text");
                                    xe1.InnerText = addPageInfo(pageNum) + "英文摘要关键词前不应空格:{" + str + "}";
                                    spRoot.AppendChild(xe1);
                                    Tool.addComment(doc, p, "英文摘要关键词前不应空格");
                                }
                            }
                        }
                        bool flag = true;
                        foreach (Run kr in krList)
                        {
                            RunProperties krpr = kr.GetFirstChild<RunProperties>();
                            Text krText = kr.GetFirstChild<Text>();
                            if (krText != null)
                            {
                                String str = krText.Text.ToString();
                                if (flag)
                                {
                                    if (krpr != null)
                                    {
                                        if (krpr.GetFirstChild<RunStyle>() != null)
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
                                                    if (m.StyleId.ToString() == krpr.GetFirstChild<RunStyle>().Val)
                                                    {
                                                        if (srPr != null)
                                                        {
                                                            if (srPr.RunFonts != null)
                                                            {
                                                                if (srPr.RunFonts.Ascii != "Times New Roman")
                                                                {
                                                                    if (str != "Key Words")
                                                                    {
                                                                        XmlElement xe1 = xmlDoc.CreateElement("Text");
                                                                        xe1.InnerText = addPageInfo(pageNum) + "英文摘要关键词字体错误，错误部分为:" + "{" + str + "}";
                                                                        spRoot.AppendChild(xe1);
                                                                        Tool.change_rfonts(p, "Times New Roman");
                                                                    }
                                                                }
                                                            }
                                                            if (srPr.FontSize != null)
                                                            {
                                                                if (srPr.FontSize.Val != "24")
                                                                {
                                                                    XmlElement xe1 = xmlDoc.CreateElement("Text");
                                                                    xe1.InnerText = addPageInfo(pageNum) + "英文摘要关键词字号错误，错误部分为:" + "{" + str + "}";
                                                                    spRoot.AppendChild(xe1);
                                                                    Tool.change_fontsize(p,"24");
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                                if (flag)
                                {
                                    if (str.Contains("："))
                                    {
                                        flag = false;
                                        continue;
                                    }
                                    else
                                    {
                                        continue;
                                    }
                                }
                                if (krpr != null)
                                {
                                    if (str.Replace(" ", "") != "")
                                    {
                                        if (str != "Key Words")
                                        {
                                            if (krpr.GetFirstChild<Bold>() != null && krpr.GetFirstChild<Bold>().Val != "0")
                                            {
                                                XmlElement xe1 = xmlDoc.CreateElement("Text");
                                                xe1.InnerText = addPageInfo(pageNum) + "英文摘要关键词内容不应加粗:" + "{" + str + "}";
                                                spRoot.AppendChild(xe1);

                                                if (krpr.GetFirstChild<BoldComplexScript>() != null)
                                                {
                                                    krpr.RemoveAllChildren<Bold>();
                                                    krpr.RemoveAllChildren<BoldComplexScript>();
                                                    Bold b = new Bold() { Val = false };
                                                    BoldComplexScript bcs = new BoldComplexScript();
                                                    krpr.AppendChild<Bold>(b);
                                                    krpr.AppendChild<BoldComplexScript>(bcs);
                                                }
                                            }
                                        }
                                    }



                                }
                            }
                        }
                        if (ss != 1)
                        {
                            XmlElement xe1 = xmlDoc.CreateElement("Text");
                            xe1.InnerText = addPageInfo(pageNum) + "英文摘要关键词与正文之间应有且只有一行空行";
                            spRoot.AppendChild(xe1);
                        }
                        xmlDoc.Save(xmlFullPath);
                        break;
                    }
                }
            }
        }
    }
}