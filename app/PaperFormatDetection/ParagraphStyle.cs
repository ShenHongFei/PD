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
    public class ParagraphStyle : ModuleFormat
    {
        //这个列表用来保存目录
        private List<String> content = new List<string>();
        /**************/
        //屏蔽代码的关键词
        /*private string[] codeKeyWord = { "//", "While", "while", "{", "}", "return", "errors", "class", "void", "try","catch","if","else","String","int","long","double", "private",
            "public","static","false","true","namespace","continue","break"
        };*/
        /***************/
        private bool codeFlag = true;
        //构造函数
        public ParagraphStyle(List<Module> modList, PageLocator locator, int masterType)
            : base(modList, locator, masterType)
        {
            
            codeFlag = true;
        }

        public override void getStyle(WordprocessingDocument doc, String fileName)
        {
            string xmlFullPath = fileName + "\\ParagraphStyle.xml";//xml模板文件保存路径
            content = getContent(doc);
            createXmlFile(xmlFullPath);
            pageNum = 5;//正文页码查询起始位置
            getTitleXml(doc, xmlFullPath); //章节标题检测
            pageNum = 5;//移动正文页码查询起始位置
            getTextStyle(doc, xmlFullPath);//正文检测
        }

        private void createXmlFile(string xmlPath)
        {
            XmlDocument xmlDoc = new XmlDocument();
            //创建类型声明节点  
            XmlNode node = xmlDoc.CreateXmlDeclaration("1.0", "utf-8", "");
            xmlDoc.AppendChild(node);
            //创建根节点  
            XmlNode root = xmlDoc.CreateElement("ParagraphStyle");
            XmlElement xe1 = xmlDoc.CreateElement("ParagraphText");
            xe1.SetAttribute("name", "段落正文");
            XmlElement xe2 = xmlDoc.CreateElement("spErroInfo");
            xe2.SetAttribute("name", "特殊错误信息");
            XmlElement xe3 = xmlDoc.CreateElement("partName");
            xe3.SetAttribute("name", "提示名称");
            XmlElement xe4 = xmlDoc.CreateElement("Text");
            xe4.InnerText = "-----------------正文-----------------";
            xe3.AppendChild(xe4);
            root.AppendChild(xe1);
            root.AppendChild(xe2);
            root.AppendChild(xe3);
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

        /**
         * 该方式用来判断段落是否是标题，是的则保存为"当前标题"，否则则认为是正文
         * 放入正文检测方法中（图，表名以及特殊章节名称也会过滤）
         */
        public void getTextStyle(WordprocessingDocument doc, String xmlFullPath)
        {
            Body body = doc.MainDocumentPart.Document.Body;
            MainDocumentPart mainPart = doc.MainDocumentPart;
            IEnumerable<Paragraph> paras = body.Elements<Paragraph>();
            string curTitle = "";
            string text = null;
            /***************start*********************/
            string text2 = null;
            bool ckwxFlag = false;//参考文献等几个特殊标题标志
            bool pText = false;//正文标志
            Paragraph temp_p = new Paragraph();
            int count = -1;
            foreach (var tp in body.ChildElements)
            {
                //遍历整个body，每一个加一
                count++;
                if (tp.GetType() == temp_p.GetType())
                {
                    Paragraph p = (Paragraph)tp;
                    Run r = p.GetFirstChild<Run>();
                    IEnumerable<Run> runList = p.Elements<Run>();
                    bool spFlag = false;
                    //正则用来过滤这些特殊标题
                    String[] reg = { "摘要", "Abstract", "目录" };
                    Regex titleOne = new Regex("[1-9]");
                    //********寻找标题
                    if (r != null)
                    {
                        //text = Tool.getFullText(p);
                        text = p.InnerText.Trim();
                        if (text != null)
                        {
                            for (int i = 0; i < reg.Length; i++)
                            {
                                if (reg[i] == text.Replace(" ", ""))
                                {
                                    spFlag = true;
                                }
                            }
                        }
                        if (spFlag)
                        {
                            continue;
                        }
                        text2 = Regex.Replace(text, @"\s*", "");
                        if (text2 == "引言")
                        {
                            pText = true;
                            curTitle = text;
                            continue;
                        }
                        if (text2 == "参考文献")
                        {
                            ckwxFlag = true;
                            continue;
                        }
                        if (text2 == "致谢")
                        {
                            ckwxFlag = false;
                        }

                        Match match1 = Regex.Match(text, @"[\u4E00-\u9FA5][1-9][0-9]*\.[1-9][0-9]*");//过滤中文图名表名
                        if (match1.Success)
                        {
                            continue;
                        }
                        match1 = Regex.Match(text, @"Tab.\ *[1-9][0-9]*\.[1-9][0-9]*");//过滤英文表名
                        if (match1.Success)
                        {
                            continue;
                        }
                        match1 = Regex.Match(text, @"Fig.\*[1-9][0-9]*\.[1-9][0-9]*");//过滤英文图名
                        if (match1.Success)
                        {
                            continue;
                        }
                        //过滤代码

                        if (text2.IndexOf("大连理工大学学位论文版权使用授权书") == 0)
                        {
                            return;
                        }
                        /*****************end************************/
                        //以上都是特殊标题的过滤
                        if (text.Trim().Length > 3)
                        {
                            if (isTitle(text))
                            {
                                //段落前几个字有.d 一般都是标题，直接过滤
                                if (p.GetFirstChild<BookmarkStart>() != null || Tool.getFullText(p).Trim().Substring(0, 3).Split('.').Length > 1)
                                {
                                    pText = true;
                                    curTitle = text;
                                }
                                else
                                {
                                    //判断是否是章标题的正则判断
                                    if (titleOne.Match(text.Trim().Substring(0, 1)).Success)
                                    {
                                        pText = true;
                                        curTitle = text;
                                    }
                                }
                                continue;
                            }
                            if (p.GetFirstChild<BookmarkStart>() != null && p.GetFirstChild<BookmarkEnd>() != null)
                            {
                                //标题的段落一般都会有BookmarkStart这个标签，所以用这个来过滤标题
                                bool flag = false;
                                if (text.Trim().Length < 30)
                                {
                                    if (titleOne.Match(text.Trim().Substring(0, 1)).Success)
                                    {
                                        if (runList != null)
                                        {
                                            foreach (Run rr in runList)
                                            {
                                                if (rr.GetFirstChild<FieldChar>() != null)
                                                {
                                                    flag = true;
                                                }
                                            }
                                        }
                                        if (flag)
                                        {
                                            continue;
                                        }
                                        pText = true;
                                        curTitle = text;
                                    }
                                    continue;
                                }
                            }
                        }
                    }
                    // 遇到正文，并且不是特殊标题
                    if (pText && !ckwxFlag)
                    {
                        var tempBe = body.ChildElements.GetItem(count - 2);
                        var tempAf = body.ChildElements.GetItem(count);
                        Table tab = new Table();
                        // 这一块的判断方法是这样，遇到有图的段落，则图的下一段是图名，不检测，遇到表的段落，则表的上一段是表名，不检测
                        if (tempAf.GetType() == tab.GetType())
                        {
                            continue;
                        }
                        Run tr = tempBe.GetFirstChild<Run>();
                        if (tr != null)
                        {
                            Drawing td = tr.GetFirstChild<Drawing>();
                            Picture pic = tr.GetFirstChild<Picture>();
                            EmbeddedObject ob = tr.GetFirstChild<EmbeddedObject>();
                            if (td != null || pic != null || ob != null)
                            {
                                continue;
                            }
                        }
                        if (r != null)
                        {
                            Drawing td1 = r.GetFirstChild<Drawing>();
                            Picture pic1 = r.GetFirstChild<Picture>();
                            EmbeddedObject ob1 = r.GetFirstChild<EmbeddedObject>();
                            if (td1 != null || pic1 != null || ob1 != null)
                            {
                                continue;
                            }
                        }
                        if (text.Trim().Length > 1)
                        {
                            if (text.Trim().Substring(0, 1) == "图")
                            {
                                if (tempBe.GetType() == tab.GetType())
                                {
                                    continue;
                                }
                            }
                            // 这个是用来排除公式，公式的标签就是DocumentFormat.OpenXml.Math.OfficeMath这个
                            if (p.GetFirstChild<DocumentFormat.OpenXml.Math.OfficeMath>() != null)
                            {
                                continue;
                            }
                            else
                            {
                                if (codeFlag)
                                {
                                    bool isCode = IsNumAndEnCh(text2);/**************改了IsNumAndEnCh函数和变量text2**************************************/
                                    //bool iscode1 = false;
                                    if (isCode)
                                    {
                                        continue;
                                    }
                                    else
                                    {
                                        checkParagraph(p, curTitle, xmlFullPath, doc);
                                    }
                                }
                                else
                                {
                                    checkParagraph(p, curTitle, xmlFullPath, doc);
                                }
                            }
                        }
                    }
                }
            }
        }

        /* 过滤后的正文，就传到这里进行检测 */
        public void checkParagraph(Paragraph p, String curTitle, String xmlFullPath, WordprocessingDocument doc)
        {
            Run r = p.GetFirstChild<Run>();
            IEnumerable<Run> pRunList = p.Elements<Run>();
            ParagraphProperties pPr = null;
            if (r != null)
            {
                pPr = p.GetFirstChild<ParagraphProperties>();
            }
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(xmlFullPath);
            XmlNode root = xmlDoc.SelectSingleNode("ParagraphStyle/spErroInfo");
            bool flag1 = false;
            bool flag2 = false;
            bool flag3 = false;
            bool isNumId1 = false;//手动输入的项目符号标记
            bool isNumId2 = false;//自动生成的项目符号标记
            bool isPic = false;

            string paraText = p.InnerText.Trim();
            pageNum = getPageNum(pageNum, paraText);

            if (paraText.Length > 3)
            {
                if (paraText.Substring(0, 1) == "（")
                {
                    isNumId1 = true;
                }
                //有可能图名过滤的不干净，这里做个疑似判断
                if (paraText.Substring(0, 1) == "图" && paraText.Substring(0, 4).Split('.').Length > 1 && paraText.Length < 20)
                {
                    isPic = true;
                    string erroInfo;
                    if (paraText.Length > 10)
                    {
                        erroInfo = this.addPageInfo(pageNum) + "该段落疑似为图名，请检查图名与图片之间是否多添加了空行或使用了不规范的图片附注：{" + curTitle + "  " + '“' + paraText.Substring(0, 10) + "······" + '”' + "}" + "（" + "疑似" + "）";
                    }
                    else
                    {
                        erroInfo = this.addPageInfo(pageNum) + "该段落疑似为图名，请检查图名与图片之间是否多添加了空行或使用了不规范的图片附注：{" + curTitle + "  " + '“' + paraText + '”' + "}" + "（" + "疑似" + "）";
                    }
                    XmlElement xe1 = xmlDoc.CreateElement("Text");
                    xe1.InnerText = erroInfo;
                    root.AppendChild(xe1);
                    //return;
                }
            }
            if (pPr != null)
            {
                if (pPr.GetFirstChild<NumberingProperties>() != null)
                {
                    if (pPr.GetFirstChild<NumberingProperties>().NumberingId != null)
                    {
                        isNumId2 = true;
                    }
                }
            }
            //对于手动输入的项目符号的检测（就是自己输入的‘（）’，不是word生成的）
            if (isNumId1)
            {
                int position = paraText.IndexOf("）");
                if (position > 0 && paraText.Length > position + 3)
                {
                    String str = paraText.Substring(position, 3);
                    if (Tool.getSpaceCount(str) != 1)
                    {
                        string erroInfo;
                        if (paraText.Length > 10)
                        {
                            erroInfo = this.addPageInfo(pageNum) + "此段落项目编号与正文之间应有且只有一个空格：{" + curTitle + "  " + '“' + paraText.Substring(0, 10) + "······" + '”' + "}";
                        }
                        else
                        {
                            erroInfo = this.addPageInfo(pageNum) + "此段落项目编号与正文之间应有且只有一个空格：{" + curTitle + "  " + '“' + paraText + '”' + "}";
                        }
                        XmlElement xe1 = xmlDoc.CreateElement("Text");
                        xe1.InnerText = erroInfo;
                        root.AppendChild(xe1);
                    }
                }
            }
            //不是图名的段落，检测段落的所有有文本的run，防止中间某个字体设置错误          
            if (pRunList != null && !isPic)
            {
                if (!Tool.correctfonts(p, doc, "宋体", "Times New Roman"))
                {
                    string erroInfo = this.addPageInfo(pageNum) + "此段落字体错误，应为宋体：{" + curTitle + "  " + '“' + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "······" + '”' + "}";
                    XmlElement xe1 = xmlDoc.CreateElement("Text");
                    xe1.InnerText = erroInfo;
                    root.AppendChild(xe1);
                }
                if (!Tool.correctsize(p, doc, "24"))
                {
                    string erroInfo = this.addPageInfo(pageNum) + "此段落字号错误，应为小四：{" + curTitle + "  " + '“' + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + '”' + "}";
                    XmlElement xe1 = xmlDoc.CreateElement("Text");
                    xe1.InnerText = erroInfo;
                    root.AppendChild(xe1);
                }
            }
            //缩进的检测和摘要部分差不多，FirstLine480的查看是否有空格，240的要有一个空格，
            //0的有两个，480就是两个字符的缩进量
            if (pPr != null && !isPic)
            {
                if (pPr.GetFirstChild<Indentation>() != null && !isNumId2)
                {
                    if (pPr.GetFirstChild<Indentation>().FirstLine != null)
                    {
                        if (pPr.GetFirstChild<Indentation>().FirstLine.Value != "480")
                        {
                            flag3 = true;
                            string erroInfo = "";
                            if (pPr.GetFirstChild<Indentation>().FirstLine.Value == "0")
                            {
                                if (paraText.Length > 5)
                                {
                                    if (Tool.getSpaceCount(paraText.Substring(0, 4)) != 4 && Tool.getSpaceCount(paraText.Substring(0, 4)) != 0)
                                    {
                                        
                                        if (paraText.Length > 10)
                                        {
                                            erroInfo = this.addPageInfo(pageNum) + "此段落缩进错误，应为首行缩进2字：{" + curTitle + "  " + '“' + paraText.Substring(0, 10) + "······" + '”' + "}";
                                        }
                                        else
                                        {
                                            erroInfo = this.addPageInfo(pageNum) + "此段落缩进错误，应为首行缩进2字：{" + curTitle + "  " + '“' + paraText + '”' + "}";
                                        }
                                        XmlElement xe1 = xmlDoc.CreateElement("Text");
                                        xe1.InnerText = erroInfo;
                                        root.AppendChild(xe1);
                                    }
                                }
                            }
                            else if (pPr.GetFirstChild<Indentation>().FirstLine.Value == "240")
                            {
                                flag3 = true;
                                if (paraText.Length > 5)
                                {
                                    if (paraText.Substring(0, 1) != "  " && Tool.getSpaceCount(paraText.Substring(0, 4)) != 0)
                                    {
                                        
                                        if (paraText.Length > 10)
                                        {
                                            erroInfo = this.addPageInfo(pageNum) + "此段落缩进错误，应为首行缩进2字：{" + curTitle + "  " + '“' + paraText.Substring(0, 10) + "······" + '”' + "}";
                                        }
                                        else
                                        {
                                            erroInfo = this.addPageInfo(pageNum) + "此段落缩进错误，应为首行缩进2字：{" + curTitle + "  " + '“' + paraText + '”' + "}";
                                        }
                                        XmlElement xe1 = xmlDoc.CreateElement("Text");
                                        xe1.InnerText = erroInfo;
                                        root.AppendChild(xe1);
                                    }
                                }
                            }
                            else
                            {
                                flag3 = true;
                                if (paraText.Length > 5 && Tool.getSpaceCount(paraText.Substring(0, 4)) != 0)
                                {
                                    if (paraText.Length > 10)
                                    {
                                        erroInfo = this.addPageInfo(pageNum) + "此段落缩进错误，应为首行缩进2字：{" + curTitle + "  " + '“' + paraText.Substring(0, 10) + "······" + '”' + "}";
                                    }
                                    else
                                    {
                                        erroInfo = this.addPageInfo(pageNum) + "此段落缩进错误，应为首行缩进2字：{" + curTitle + "  " + '“' + paraText + '”' + "}";
                                    }
                                    XmlElement xe1 = xmlDoc.CreateElement("Text");
                                    xe1.InnerText = erroInfo;
                                    root.AppendChild(xe1);
                                }                               
                            }
                        }
                        else
                        {
                            flag3 = true;
                            string erroInfo = "";
                            if (paraText.Length > 5 && Tool.getSpaceCount(paraText.Substring(0, 4)) != 0)
                            {
                                if (paraText.Substring(0, 1) == " ")
                                {
                                    if (paraText.Length > 10)
                                    {
                                        erroInfo = this.addPageInfo(pageNum) + "此段落缩进错误，应为首行缩进2字：{" + curTitle + "  " + '“' + paraText.Substring(0, 10) + "······" + '”' + "}";
                                    }
                                    else
                                    {
                                        erroInfo = this.addPageInfo(pageNum) + "此段落缩进错误，应为首行缩进2字：{" + curTitle + "  " + '“' + paraText + '”' + "}";
                                    }
                                    XmlElement xe1 = xmlDoc.CreateElement("Text");
                                    xe1.InnerText = erroInfo;
                                    root.AppendChild(xe1);
                                }
                            }
                        }
                    }
                }
                //段间距的判断
                if (pPr.GetFirstChild<SpacingBetweenLines>() != null && !isPic)
                {
                    if (pPr.GetFirstChild<SpacingBetweenLines>().BeforeLines != null)
                    {
                        if (pPr.GetFirstChild<SpacingBetweenLines>().BeforeLines.Value != 0)
                        {
                            string erroInfo;
                            if (paraText.Length > 10)
                            {
                                erroInfo = this.addPageInfo(pageNum) + "此段落段前间距错误，应为段前0行：{" + curTitle + "  " + '“' + paraText.Substring(0, 10) + "······" + '”' + "}";
                            }
                            else
                            {
                                erroInfo = this.addPageInfo(pageNum) + "此段落段前间距错误，应为段前0行：{" + curTitle + "  " + '“' + paraText + '”' + "}";
                            }
                            XmlElement xe1 = xmlDoc.CreateElement("Text");
                            xe1.InnerText = erroInfo;
                            root.AppendChild(xe1);
                        }

                    }
                    if (pPr.GetFirstChild<SpacingBetweenLines>().Line != null)
                    {
                        if (pPr.GetFirstChild<SpacingBetweenLines>().Line.Value != "300")
                        {
                            string erroInfo;
                            if (paraText.Length > 10)
                            {
                                erroInfo = this.addPageInfo(pageNum) + "此段落行间距错误，应为多倍行距1.25：{" + curTitle + "  " + '“' + paraText.Substring(0, 10) + "······" + '”' + "}";
                            }
                            else
                            {
                                erroInfo = this.addPageInfo(pageNum) + "此段落行间距错误，应为多倍行距1.25：{" + curTitle + "  " + '“' + paraText + '”' + "}";
                            }
                            XmlElement xe1 = xmlDoc.CreateElement("Text");
                            xe1.InnerText = erroInfo;
                            root.AppendChild(xe1);
                        }
                    }
                    if (pPr.GetFirstChild<SpacingBetweenLines>().AfterLines != null)
                    {
                        if (pPr.GetFirstChild<SpacingBetweenLines>().AfterLines.Value != 0)
                        {
                            string erroInfo;
                            if (paraText.Length > 10)
                            {
                                erroInfo = this.addPageInfo(pageNum) + "此段落段后间距错误，应为段后0行：{" + curTitle + "  " + '“' + paraText.Substring(0, 10) + "······" + '”' + "}";
                            }
                            else
                            {
                                erroInfo = this.addPageInfo(pageNum) + "此段落段后间距错误，应为段后0行：{" + curTitle + "  " + '“' + paraText + '”' + "}";
                            }
                            XmlElement xe1 = xmlDoc.CreateElement("Text");
                            xe1.InnerText = erroInfo;
                            root.AppendChild(xe1);
                        }
                    }
                }
            }
            //段落使用样式的检测，也是一样的，段落属性优先于样式，段落属性如果设置了字号为三号，样式也设置了字号为四号，
            //那么word里这个段落字号应该是三号，因为段落属性优先使用
            String id = Tool.getPargraphStyleId(p);
            if (id != "" && !isPic)
            {
                MainDocumentPart mainPart = doc.MainDocumentPart;
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
                            if (spPr != null)
                            {
                                if (spPr.Indentation != null && !flag3 && !isNumId2)
                                {
                                    if (spPr.Indentation.FirstLine != null)
                                    {
                                        if (spPr.Indentation.FirstLine.Value != "480")
                                        {
                                            if (pPr != null)
                                            {
                                                if (pPr.GetFirstChild<NumberingProperties>() != null)
                                                {
                                                    if (paraText.Substring(0, 1) == " ")
                                                    {
                                                        string erroInfo;
                                                        /*if (paraText.Length > 10)
                                                        {
                                                            erroInfo = this.addPageInfo(pageNum) + "此段落缩进错误，应为首行缩进2字：{" + curTitle + "  " + '“' + paraText.Substring(0, 10) + "······" + '”' + "}";
                                                        }
                                                        else
                                                        {
                                                            erroInfo = this.addPageInfo(pageNum) + "此段落缩进错误，应为首行缩进2字：{" + curTitle + "  " + '“' + paraText + '”' + "}";
                                                        }
                                                        XmlElement xe1 = xmlDoc.CreateElement("Text");
                                                        xe1.InnerText = erroInfo;
                                                        root.AppendChild(xe1);*/
                                                    }
                                                }
                                                else
                                                {
                                                    string erroInfo;
                                                    /*if (paraText.Length > 10)
                                                    {
                                                        erroInfo = this.addPageInfo(pageNum) + "此段落缩进错误，应为首行缩进2字：{" + curTitle + "  " + '“' + paraText.Substring(0, 10) + "······" + '”' + "}";
                                                    }
                                                    else
                                                    {
                                                        erroInfo = this.addPageInfo(pageNum) + "此段落缩进错误，应为首行缩进2字：{" + curTitle + "  " + '“' + paraText + '”' + "}";
                                                    }
                                                    XmlElement xe1 = xmlDoc.CreateElement("Text");
                                                    xe1.InnerText = erroInfo;
                                                    root.AppendChild(xe1);*/
                                                }
                                            }
                                        }
                                    }

                                }
                                else if (!flag3)
                                {
                                    string erroInfo;
                                    /*if (paraText.Length > 10)
                                    {
                                        erroInfo = this.addPageInfo(pageNum) + "此段落缩进错误，应为首行缩进2字：{" + curTitle + "  " + '“' + paraText.Substring(0, 10) + "······" + '”' + "}";
                                    }
                                    else
                                    {
                                        erroInfo = this.addPageInfo(pageNum) + "此段落缩进错误，应为首行缩进2字：{" + curTitle + "  " + '“' + paraText + '”' + "}";
                                    }
                                    XmlElement xe1 = xmlDoc.CreateElement("Text");
                                    xe1.InnerText = erroInfo;
                                    root.AppendChild(xe1);*/
                                }
                            }
                        }
                    }
                }
            }
            /* 既没有使用样式，也没有设置缩进的时候，可能使用了空格 */
            /*else if (!flag3 && !isPic && !isNumId2)
            {
                if (Tool.getFullText(p).Length > 5)
                {
                    string erroInfo;
                    if (Tool.getSpaceCount(paraText.Substring(0, 4)) != 4)
                    {
                        if (paraText.Length > 10)
                        {
                            erroInfo = this.addPageInfo(pageNum) + "此段落缩进错误，应为首行缩进2字：{" + curTitle + "  " + '“' + paraText.Substring(0, 10) + "······" + '”' + "}";
                        }
                        else
                        {
                            erroInfo = this.addPageInfo(pageNum) + "此段落缩进错误，应为首行缩进2字：{" + curTitle + "  " + '“' + paraText + '”' + "}";
                        }
                        XmlElement xe1 = xmlDoc.CreateElement("Text");
                        xe1.InnerText = erroInfo;
                        root.AppendChild(xe1);
                    }
                }
            }*/
            xmlDoc.Save(xmlFullPath);
        }

        /* 正文标题部分检测函数 */
        private void getTitleXml(WordprocessingDocument docx, string xmlFullPath)
        {
            const int RegexNumber = 13;
            const int propertiesNumber = 7;
            // const int level = 7;
            string[,] regulation;
            regulation = new string[7, propertiesNumber]{
                {"item","字体","字号大小","标题位置","段前间距","段后间距","行间距"},
                {"大连理工大学学位论文版权使用授权书","黑体","","center","0","220","360"},
                {"Abstract","Cambria","30","center","0","240","360"},
                {"一级","黑体","30","left","0","240","360"},
                {"二级","黑体","28","left","120","0","360"},
                {"三级","黑体","24","left","120","0","360"},
                {"引言、参考文献","黑体","30","center","0","240","360"}};
            string[,] result;
            result = new string[7, propertiesNumber]{
                {"item","字体","字号大小","标题位置","段前间距","段后间距","行间距"},
                {"大连理工大学学位论文版权使用授权书","黑体","三号","居中","0","段后12磅","1.5倍行间距"},
                {"Abstract","Cambria","小三","居中","段前0行","段后1行","1.5倍行间距"},
                {"一级","黑体","小三","靠左","段前0行","段后1行","1.5倍行间距"},
                {"二级","黑体","四号","靠左","段前0.5行","0","1.5倍行间距"},
                {"三级","黑体","小四","靠左","段前0.5行","0","1.5倍行间距"},
                {"引言、参考文献","黑体","小三","居中","0","段后1行","1.5倍行间距"}};//
            //after240,afterline100
            Regex[] reg;
            reg = new Regex[RegexNumber];
            reg[0] = new Regex(@"摘,,,,要");//6
            reg[1] = new Regex(@"Abstract");//2
            reg[2] = new Regex(@"引,,,,言");//6
            reg[3] = new Regex(@"结,,,,论");//6
            reg[4] = new Regex(@"参,考,文,献");//6
            reg[5] = new Regex(@"攻读硕士学位期间发表学术论文情况");//6
            reg[6] = new Regex(@"致,,,,谢");//6
            reg[7] = new Regex(@"附录[A-Z],,");//6
            // reg[8] = new Regex(@",,");//
            reg[10] = new Regex(@"[1-9][0-9]*");//3
            reg[9] = new Regex(@"[1-9][0-9]*\.[1-9][0-9]*");//4
            reg[8] = new Regex(@"[1-9][0-9]*\.[1-9][0-9]*\.[1-9][0-9]*");//5
            reg[12] = new Regex(@"\,\,");
            // reg[11] = new Regex(@"大连理工大学学位论文版权使用授权书");
            XmlDocument xmlDocx = new XmlDocument();
            xmlDocx.Load(xmlFullPath);
            XmlNode root = xmlDocx.SelectSingleNode("ParagraphStyle/spErroInfo");
            Body body = docx.MainDocumentPart.Document.Body;
            IEnumerable<Paragraph> paragraph = body.Elements<Paragraph>();
            StyleDefinitionsPart sDp = docx.MainDocumentPart.StyleDefinitionsPart;
            Styles styles = sDp.Styles;
            var t = styles.ChildElements;
            foreach (Paragraph p in paragraph)
            {
                bool flag = false;
                //flag = new bool[RegexNumber];
                string sentence = null;
                ParagraphProperties pPr = p.GetFirstChild<ParagraphProperties>();
                BookmarkStart bS = p.GetFirstChild<BookmarkStart>();
                if (bS == null)
                    continue;
                IEnumerable<Run> run = p.Elements<Run>();
                sentence = p.InnerText.Trim();
                if (sentence == null || sentence.Length > 30)//过滤标题
                    continue;
                string sentence2 = Regex.Replace(sentence, @"\s*", "");
                if (sentence2.IndexOf("目录") == 0 || sentence2.IndexOf("大连理工大学学位论文版权使用授权书") == 0)//过滤目录,授权书
                    continue;

                pageNum = this.getPageNum(pageNum, sentence);
                sentence = sentence.Replace(" ", ",");
                //Console.WriteLine(sentence);
                bool title = true;
                for (int i = 0; i < 11; i++)
                {
                    title = true;
                    Match m = reg[i].Match(sentence);
                    if (m.Success == true)
                    {
                        sentence = sentence.Replace(',', ' ');
                        int index = -1;
                        switch (i)
                        {
                            case 0:
                            case 2:
                            case 3:
                            case 4:
                            case 5:
                            case 6:
                            case 7:
                                index = 6; break;
                            case 1: index = 2; break;
                            case 10:
                                index = 3;
                                if (m.Index != 0)
                                {
                                    title = false;
                                    break;
                                }
                                sentence = sentence.Replace(' ', ',');
                                Match m2 = reg[12].Match(sentence);
                                if (!m2.Success)
                                {
                                    sentence = sentence.Replace(',', ' ');
                                    XmlElement Xml = xmlDocx.CreateElement("Text");
                                    sentence = sentence.Replace(",", "");
                                    Xml.InnerText = "该一级标题序号后缺少两个空格：" +
                                            "{" + sentence + "}";
                                    root.AppendChild(Xml);
                                }
                                else
                                {
                                    if (m2.Index != m.Index + m.Value.Length)
                                    {
                                        sentence = sentence.Replace(',', ' ');
                                        XmlElement Xml = xmlDocx.CreateElement("Text");
                                        sentence = sentence.Replace(",", "");
                                        Xml.InnerText = "该一级标题序号后缺少两个空格：" +
                                                "{" + sentence + "}";
                                        root.AppendChild(Xml);
                                    }
                                }
                                break;
                            case 9:
                                index = 4;
                                if (m.Index != 0)
                                {
                                    title = false;
                                    break;
                                }
                                sentence = sentence.Replace(' ', ',');
                                Match m3 = reg[12].Match(sentence);
                                if (!m3.Success)
                                {
                                    sentence = sentence.Replace(',', ' ');
                                    XmlElement Xml = xmlDocx.CreateElement("Text");
                                    sentence = sentence.Replace(",", "");
                                    Xml.InnerText = "该二级标题序号后缺少两个空格：" +
                                            "{" + sentence + "}";
                                    root.AppendChild(Xml);
                                }
                                else
                                {
                                    if (m3.Index != m.Index + m.Value.Length)
                                    {
                                        sentence = sentence.Replace(',', ' ');
                                        XmlElement Xml = xmlDocx.CreateElement("Text");
                                        sentence = sentence.Replace(",", "");
                                        Xml.InnerText = "该二级标题序号后缺少两个空格：" +
                                                "{" + sentence + "}";
                                        root.AppendChild(Xml);
                                    }
                                }
                                break;
                            case 8:
                                index = 5;
                                sentence = sentence.Replace(' ', ',');
                                if (m.Index != 0)
                                {
                                    title = false;
                                    break;
                                }
                                Match m4 = reg[12].Match(sentence);
                                if (!m4.Success)
                                {
                                    sentence = sentence.Replace(',', ' ');
                                    XmlElement Xml = xmlDocx.CreateElement("Text");
                                    sentence = sentence.Replace(",", "");
                                    Xml.InnerText = "该三级标题序号后缺少两个空格：" +
                                            "{" + sentence + "}";
                                    root.AppendChild(Xml);
                                }
                                else
                                {
                                    if (m4.Index != m.Index + m.Value.Length)
                                    {
                                        sentence = sentence.Replace(',', ' ');
                                        XmlElement Xml = xmlDocx.CreateElement("Text");
                                        sentence = sentence.Replace(",", "");
                                        Xml.InnerText = "该三级标题序号后缺少两个空格：" +
                                                "{" + sentence + "}";
                                        root.AppendChild(Xml);
                                    }
                                }
                                break;
                            case 12: index = -1; break;
                        }
                        if (title == false)
                        {
                            break;
                        }
                        if (index == -1)
                        {
                            XmlElement Xml = xmlDocx.CreateElement("Text");
                            sentence = sentence.Replace(",", "");
                            Xml.InnerText = "该标题缺少内容：" +
                                    "{" + sentence + "}";
                            root.AppendChild(Xml);
                            continue;
                        }
                        flag = true;
                        sentence = sentence.Replace(",", "");
                        if (index == 3 || index == 4 || index == 5)
                        {
                            string number = sentence.Substring(0, m.Length);
                            if (!correctNumberingFontsInTitle(p, number, docx))
                            {
                                XmlElement Xml = xmlDocx.CreateElement("Text");
                                Xml.InnerText = this.addPageInfo(pageNum) + "该标题编号字体错误，应为Cambria" + ":" +
                                        "{" + sentence + "}";
                                root.AppendChild(Xml);
                            }
                            if (!Tool.correctfonts(p, docx, regulation[index, 1], "Cambria"))//看序号后的标题
                            {
                                XmlElement Xml = xmlDocx.CreateElement("Text");
                                Xml.InnerText = this.addPageInfo(pageNum) + "该标题编号后的内容字体错误，应为" + result[index, 1] + ":" +
                                        "{" + sentence + "}";
                                root.AppendChild(Xml);
                            }
                        }
                        else if (index != 2)
                        {
                            if (!Tool.correctfonts(p, docx, regulation[index, 1], "Cambria"))
                            {
                                XmlElement Xml = xmlDocx.CreateElement("Text");
                                Xml.InnerText = this.addPageInfo(pageNum) + "该标题字体错误，应为" + result[index, 1] + ":" +
                                        "{" + sentence + "}";
                                root.AppendChild(Xml);
                            }
                        }
                        else if (index == 2)
                        {
                            if (!Tool.correctfonts(p, docx, "宋体", regulation[index, 1]))
                            {
                                XmlElement Xml = xmlDocx.CreateElement("Text");
                                Xml.InnerText = this.addPageInfo(pageNum) + "该标题字体错误，应为" + result[index, 1] + ":" +
                                        "{" + sentence + "}";
                                root.AppendChild(Xml);
                            }
                        }
                        if (!Tool.correctsize(p, docx, regulation[index, 2]))
                        {
                            XmlElement Xml = xmlDocx.CreateElement("Text");
                            Xml.InnerText = this.addPageInfo(pageNum) + "该标题字号错误，应为" + result[index, 2] + ":" +
                                    "{" + sentence + "}";
                            root.AppendChild(Xml);
                        }
                        if (!Tool.correctJustification(p, docx, regulation[index, 3]))
                        {
                            XmlElement Xml = xmlDocx.CreateElement("Text");
                            Xml.InnerText = this.addPageInfo(pageNum) + "该标题位置错误，应为" + result[index, 3] + ":" +
                                    "{" + sentence + "}";
                            root.AppendChild(Xml);
                        }
                        if (!Tool.correctSpacingBetweenLines_Be(p, docx, regulation[index, 4]))
                        {
                            XmlElement Xml = xmlDocx.CreateElement("Text");
                            Xml.InnerText = this.addPageInfo(pageNum) + "该标题段前间距错误，应为" + result[index, 4] + ":" +
                                    "{" + sentence + "}";
                            root.AppendChild(Xml);
                        }
                        //段后间距
                        bool haveSpacing = true;

                        if (pPr.GetFirstChild<SpacingBetweenLines>() != null)
                        {
                            /*if (pPr.GetFirstChild<SpacingBetweenLines>().After != null && pPr.GetFirstChild<SpacingBetweenLines>().After != "312")
                            {
                                XmlElement Xml = xmlDocx.CreateElement("Text");
                                Xml.InnerText = this.addPageInfo(pageNum) + "该标题段前后距错误，应为1.5倍行距，段后间距一行:" + "{" + sentence + "}";
                                root.AppendChild(Xml);
                            }*/

                        }
                        else if (pPr.GetFirstChild<SpacingBetweenLines>() == null)
                        {
                            haveSpacing = false;
                        }
                        if (!Tool.correctSpacingBetweenLines_Af(p, docx, regulation[index, 5]))
                        {
                            /*if (haveSpacing == false)
                            {
                                XmlElement Xml = xmlDocx.CreateElement("Text");
                                Xml.InnerText = this.addPageInfo(pageNum) + "该标题段后间距错误，应为1.5倍行距，段后间距一行:" +
                                        "{" + sentence + "}";
                                root.AppendChild(Xml);
                            }*/
                        }
                        if (!Tool.correctSpacingBetweenLines_line(p, docx, regulation[index, 6]))
                        {
                            XmlElement Xml = xmlDocx.CreateElement("Text");
                            Xml.InnerText = this.addPageInfo(pageNum) + "该标题行间距错误，应为" + result[index, 5] + ":" +
                                    "{" + sentence + "}";
                            root.AppendChild(Xml);
                        }
                        break;
                    }
                }
                if (flag == false && title)
                {
                    if (sentence.Length == 0)
                    {
                       
                    }
                    else
                    {
                        //检测是否是自动编号
                        if (p.ParagraphProperties != null)
                        {
                            if (p.ParagraphProperties.NumberingProperties == null)
                            {
                                XmlElement Xml = xmlDocx.CreateElement("Text");
                                sentence = sentence.Replace(",", "");
                                Xml.InnerText = this.addPageInfo(pageNum) + "该标题格式错误，可能是忘记编号：" +
                                        "{" + sentence + "}";
                                root.AppendChild(Xml);
                                continue;
                            }
                            else
                            {
                                string numberingId = p.ParagraphProperties.NumberingProperties.NumberingId.Val;
                                string ilvl = p.ParagraphProperties.NumberingProperties.NumberingLevelReference.Val;
                                NumberingDefinitionsPart numberingDefinitionsPart1 = docx.MainDocumentPart.NumberingDefinitionsPart;
                                Numbering numbering1 = numberingDefinitionsPart1.Numbering;
                                IEnumerable<NumberingInstance> nums = numbering1.Elements<NumberingInstance>();
                                IEnumerable<AbstractNum> abstractNums = numbering1.Elements<AbstractNum>();
                                foreach (NumberingInstance num in nums)
                                {
                                    if (num.NumberID == numberingId)
                                    {
                                        Int32 abstractNumId1 = num.AbstractNumId.Val;
                                        foreach (AbstractNum abstractNum in abstractNums)
                                        {
                                            if (abstractNum.AbstractNumberId == abstractNumId1)
                                            {
                                                IEnumerable<Level> levels = abstractNum.Elements<Level>();
                                                foreach (Level level in levels)
                                                {
                                                    if (level.LevelIndex == ilvl)
                                                    {
                                                        string levelText1 = level.LevelText.Val;
                                                        Match match1 = Regex.Match(levelText1, @"\%[0-9]");
                                                        if (match1.Success)
                                                        {
                                                            flag = true;
                                                        }
                                                        Match match2 = Regex.Match(levelText1, @"\%[0-9]+\.\%[0-9]");
                                                        if (match2.Success)
                                                        {
                                                            flag = true;
                                                        }
                                                        Match match3 = Regex.Match(levelText1, @"\%[0-9]+\.\%[0-9]+\.\%[0-3]");
                                                        if (match3.Success)
                                                        {
                                                            flag = true;
                                                        }
                                                        break;
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                if (!title)
                {
                    XmlElement Xml = xmlDocx.CreateElement("Text");
                    sentence = sentence.Replace(",", "");
                    Xml.InnerText = this.addPageInfo(pageNum) + "该标题格式错误，可能多加了空格或其他符号" +
                            "{" + sentence + "}";
                    root.AppendChild(Xml);
                    continue;
                }
            }//对比
            xmlDocx.Save(xmlFullPath);
        }
        public List<String> getContent(WordprocessingDocument doc)
        {
            Body body = doc.MainDocumentPart.Document.Body;
            MainDocumentPart mainPart = doc.MainDocumentPart;
            IEnumerable<Paragraph> paras = body.Elements<Paragraph>();
            List<String> contens = new List<string>();
            bool isContents = false;
            foreach (Paragraph p in paras)
            {
                Run r = p.GetFirstChild<Run>();
                String fullText = "";
                if (r != null)
                {
                    fullText = Tool.getFullText(p).Trim();
                }
                if (fullText.Replace(" ", "") == "目录" && !isContents)
                {
                    isContents = true;
                    continue;
                }
                if (isContents)
                {
                    Hyperlink h = p.GetFirstChild<Hyperlink>();
                    if (h != null)
                    {
                        contens.Add(getHyperlinkFullText(h));
                    }
                }
            }
            return contens;
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

        private bool isTitle(String str)
        {
            foreach (String s in content)
            {
                if (s.Contains(str.Trim()))
                {
                    return true;
                }
            }
            return false;
        }

        private static bool IsNumAndEnCh(string str)
        {
            Match match = Regex.Match(str, @"[\u4e00-\u9fa5]");
            int index = str.IndexOf("//") != -1 ? str.IndexOf("//") : str.IndexOf("/*");
            if (str.Length > 40)
            {
                return false;
            }
            else
            {
                if (match.Success)
                {
                    if (index == -1)
                    {
                        return false;
                    }
                    else
                    {
                        return index > match.Index;
                    }
                }
                else
                {
                    return true;
                }
            }
        }

        private int getPosition(Body body, Paragraph p)
        {
            int count = 0;
            foreach (var t in body.ChildElements)
            {
                count++;
                if (t.GetType() == p.GetType())
                {
                    if (t.GetFirstChild<Run>() != null)
                    {
                        if (Tool.getFullText((Paragraph)t) == Tool.getFullText(p))
                        {
                            return count;
                        }
                    }
                }
            }
            return 0;
        }
        private static bool correctNumberingFontsInTitle(Paragraph p, string number, WordprocessingDocument doc)
        {
            //string s = p.InnerText.Trim();
            IEnumerable<Style> style = doc.MainDocumentPart.StyleDefinitionsPart.Styles.Elements<Style>();
            //正文style(default (Default Style))
            Style Normalst = null;
            foreach (Style s in style)
            {
                if (s.Type == StyleValues.Paragraph && s.Default == true)
                {
                    Normalst = s;
                    break;
                }
            }
            //正文样式字体(default (Default Style))
            string Normalfonts = null;
            string CNNormalfonts = null;
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
                        else if (Normalst.StyleRunProperties.RunFonts.AsciiTheme != null)
                        {
                            Normalfonts = Normalst.StyleRunProperties.RunFonts.Ascii.ToString();
                        }
                        if (Normalst.StyleRunProperties.RunFonts.EastAsia != null)
                        {
                            CNNormalfonts = Normalst.StyleRunProperties.RunFonts.EastAsia;
                        }
                        else if (Normalst.StyleRunProperties.RunFonts.EastAsiaTheme != null)
                        {
                            CNNormalfonts = Normalst.StyleRunProperties.RunFonts.EastAsiaTheme;
                        }
                    }
                }
            }
            //defaults
            string Defaultsfonts = null;
            string CNDefaultsfonts = null;
            if (doc.MainDocumentPart.StyleDefinitionsPart.Styles.DocDefaults != null)
            {
                if (doc.MainDocumentPart.StyleDefinitionsPart.Styles.DocDefaults.RunPropertiesDefault != null)
                {
                    if (doc.MainDocumentPart.StyleDefinitionsPart.Styles.DocDefaults.RunPropertiesDefault.RunPropertiesBaseStyle != null)
                    {
                        if (doc.MainDocumentPart.StyleDefinitionsPart.Styles.DocDefaults.RunPropertiesDefault.RunPropertiesBaseStyle.RunFonts != null)
                        {
                            if (doc.MainDocumentPart.StyleDefinitionsPart.Styles.DocDefaults.RunPropertiesDefault.RunPropertiesBaseStyle.RunFonts.Ascii != null)
                            {
                                Defaultsfonts = doc.MainDocumentPart.StyleDefinitionsPart.Styles.DocDefaults.RunPropertiesDefault.RunPropertiesBaseStyle.RunFonts.Ascii;
                            }
                            else if (doc.MainDocumentPart.StyleDefinitionsPart.Styles.DocDefaults.RunPropertiesDefault.RunPropertiesBaseStyle.RunFonts.AsciiTheme != null)
                            {
                                Defaultsfonts = doc.MainDocumentPart.StyleDefinitionsPart.Styles.DocDefaults.RunPropertiesDefault.RunPropertiesBaseStyle.RunFonts.AsciiTheme;
                            }

                            if (doc.MainDocumentPart.StyleDefinitionsPart.Styles.DocDefaults.RunPropertiesDefault.RunPropertiesBaseStyle.RunFonts.EastAsia != null)
                            {
                                CNDefaultsfonts = doc.MainDocumentPart.StyleDefinitionsPart.Styles.DocDefaults.RunPropertiesDefault.RunPropertiesBaseStyle.RunFonts.EastAsia;
                            }
                            else if (doc.MainDocumentPart.StyleDefinitionsPart.Styles.DocDefaults.RunPropertiesDefault.RunPropertiesBaseStyle.RunFonts.EastAsiaTheme != null)
                            {
                                CNDefaultsfonts = doc.MainDocumentPart.StyleDefinitionsPart.Styles.DocDefaults.RunPropertiesDefault.RunPropertiesBaseStyle.RunFonts.EastAsiaTheme;
                            }
                        }
                    }
                }
            }
            //pstyleid
            string pstyleid = null;
            string pbasestyleid = null;
            //段落style
            Style pstyle = null;
            string pstylefonts = null;
            string CNpstylefonts = null;
            // Style pbasestyle = null;
            string pbasestylefonts = null;
            string CNpbasestylefonts = null;
            if (p.GetFirstChild<ParagraphProperties>() != null)
            {
                if (p.GetFirstChild<ParagraphProperties>().ParagraphStyleId != null)
                {
                    if (p.GetFirstChild<ParagraphProperties>().ParagraphStyleId.Val != null)
                    {
                        pstyleid = p.GetFirstChild<ParagraphProperties>().ParagraphStyleId.Val;
                        foreach (Style s in style)
                        {
                            if (s.StyleId == pstyleid)
                            {
                                pstyle = s;
                                if (pstyle.StyleRunProperties != null)
                                {
                                    if (pstyle.StyleRunProperties.RunFonts != null)
                                    {
                                        if (pstyle.StyleRunProperties.RunFonts.Ascii != null)
                                        {
                                            if (pstyle.StyleRunProperties.RunFonts.Ascii != null)
                                            {
                                                pstylefonts = pstyle.StyleRunProperties.RunFonts.Ascii;
                                            }
                                            else if (pstyle.StyleRunProperties.RunFonts.AsciiTheme != null)
                                            {
                                                pstylefonts = pstyle.StyleRunProperties.RunFonts.AsciiTheme;
                                            }
                                            if (pstyle.StyleRunProperties.RunFonts.EastAsia != null)
                                            {
                                                CNpstylefonts = pstyle.StyleRunProperties.RunFonts.EastAsia;
                                            }
                                            else if (pstyle.StyleRunProperties.RunFonts.EastAsiaTheme != null)
                                            {
                                                CNpstylefonts = pstyle.StyleRunProperties.RunFonts.EastAsiaTheme;
                                            }
                                        }
                                    }
                                }
                                break;
                            }
                        }
                    }
                }
            }
            //pstyle-basedon
            if (pstyle != null)
            {
                while (pstyle.BasedOn != null && (pbasestylefonts == null || CNpbasestylefonts == null))
                {
                    if (pstyle.BasedOn.Val != null)
                    {
                        pbasestyleid = pstyle.BasedOn.Val;
                    }
                    if (pbasestyleid != null)
                    {
                        foreach (Style s in style)
                        {
                            if (s.StyleId == pbasestyleid)
                            {
                                pstyle = s;
                                if (s.StyleRunProperties != null)
                                {
                                    if (s.StyleRunProperties.RunFonts != null)
                                    {
                                        if (s.StyleRunProperties.RunFonts.Ascii != null)
                                        {
                                            if (s.StyleRunProperties.RunFonts.Ascii != null && pbasestylefonts == null)
                                            {
                                                pbasestylefonts = s.StyleRunProperties.RunFonts.Ascii.ToString();
                                            }
                                            else if (s.StyleRunProperties.RunFonts.AsciiTheme != null && pbasestylefonts == null)
                                            {
                                                pbasestylefonts = s.StyleRunProperties.RunFonts.AsciiTheme;
                                            }
                                            if (s.StyleRunProperties.RunFonts.EastAsia != null && CNpbasestylefonts == null)
                                            {
                                                CNpbasestylefonts = s.StyleRunProperties.RunFonts.EastAsia;
                                            }
                                            else if (s.StyleRunProperties.RunFonts.EastAsiaTheme != null && CNpbasestylefonts == null)
                                            {
                                                CNpbasestylefonts = s.StyleRunProperties.RunFonts.EastAsiaTheme;
                                            }
                                        }
                                    }
                                }
                                break;
                            }
                        }
                    }
                }
            }
            IEnumerable<Run> run = p.Elements<Run>();
            foreach (Run r in run)
            {
                string rtext = r.InnerText.Replace(' ', '\0');
                if (rtext.Length != 0)
                {
                    if (number.IndexOf(rtext) == -1)
                    {
                        continue;
                    }
                    //rstyleid
                    string rstyleid = null;
                    //rBaseonstyleid
                    string rBasestyleid = null;
                    //rfonts
                    string rfonts = null;
                    string CNrfonts = null;
                    //rstylefonts
                    string rstylefonts = null;
                    string CNrstylefonts = null;
                    //rBaseonfonts
                    string rBasefonts = null;
                    string CNrBasefonts = null;
                    //rstyle
                    Style rstyle = null;
                    //rBaseonstyle
                    Style rBasestyle = null;
                    if (r.RunProperties != null)
                    {
                        if (r.RunProperties.RunStyle != null)
                        {
                            if (r.RunProperties.RunStyle.Val != null)
                            {
                                rstyleid = r.RunProperties.RunStyle.Val.ToString();
                                foreach (Style s in style)
                                {
                                    if (s.StyleId == rstyleid)
                                    {
                                        rstyle = s;
                                        if (rstyle.StyleRunProperties != null)
                                        {
                                            if (rstyle.StyleRunProperties.RunFonts != null)
                                            {
                                                if (rstyle.StyleRunProperties.RunFonts.Ascii != null)
                                                {
                                                    rstylefonts = rstyle.StyleRunProperties.RunFonts.Ascii;
                                                }
                                                else if (rstyle.StyleRunProperties.RunFonts.AsciiTheme != null)
                                                {
                                                    rstylefonts = rstyle.StyleRunProperties.RunFonts.AsciiTheme;
                                                }
                                                if (rstyle.StyleRunProperties.RunFonts.EastAsia != null)
                                                {
                                                    CNrstylefonts = rstyle.StyleRunProperties.RunFonts.EastAsia;
                                                }
                                                else if (rstyle.StyleRunProperties.RunFonts.EastAsiaTheme != null)
                                                {
                                                    CNrstylefonts = rstyle.StyleRunProperties.RunFonts.EastAsiaTheme;
                                                }
                                            }
                                        }
                                        break;
                                    }
                                }
                            }
                        }
                    }

                    if (rstyle != null)
                    {
                        if (rstyle.BasedOn != null)
                        {
                            if (rstyle.BasedOn.Val != null)
                            {
                                rBasestyleid = rstyle.BasedOn.Val;
                                foreach (Style s in style)
                                {
                                    if (s.StyleId == rBasestyleid)
                                    {
                                        rBasestyle = s;
                                        if (rBasestyle.StyleRunProperties != null)
                                        {
                                            if (rBasestyle.StyleRunProperties.RunFonts != null)
                                            {
                                                if (rBasestyle.StyleRunProperties.RunFonts.Ascii != null)
                                                {
                                                    rBasefonts = rBasestyle.StyleRunProperties.RunFonts.Ascii;
                                                }
                                                else if (rBasestyle.StyleRunProperties.RunFonts.AsciiTheme != null)
                                                {
                                                    rBasefonts = rBasestyle.StyleRunProperties.RunFonts.AsciiTheme;
                                                }
                                                if (rBasestyle.StyleRunProperties.RunFonts.EastAsia != null)
                                                {
                                                    CNrBasefonts = rBasestyle.StyleRunProperties.RunFonts.EastAsia;
                                                }
                                                else if (rBasestyle.StyleRunProperties.RunFonts.EastAsiaTheme != null)
                                                {
                                                    CNrBasefonts = rBasestyle.StyleRunProperties.RunFonts.EastAsiaTheme;
                                                }
                                            }
                                        }
                                        break;
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
                                rfonts = rpr.RunFonts.Ascii;
                            }
                            else if (rpr.RunFonts.AsciiTheme != null)
                            {
                                rfonts = rpr.RunFonts.AsciiTheme;
                            }
                            if (rpr.RunFonts.EastAsia != null)
                            {
                                CNrfonts = rpr.RunFonts.EastAsia;
                            }
                            else if (rpr.RunFonts.EastAsiaTheme != null)
                            {
                                CNrfonts = rpr.RunFonts.EastAsiaTheme;
                            }
                        }
                    }
                    string ENfonts = "Cambria";
                    if (rfonts != null)
                    {
                        if (rfonts != ENfonts)
                        {
                            return false;
                        }
                        else
                        {
                            return true;
                        }
                    }
                    else if (rstylefonts != null)
                    {
                        if (rstylefonts != ENfonts)
                        {
                            return false;
                        }
                        else
                        {
                            return true;
                        }
                    }
                    else if (rBasefonts != null)
                    {
                        if (rBasefonts != ENfonts)
                        {
                            return false;
                        }
                        else
                        {
                            return true;
                        }
                    }
                    else if (pstylefonts != null)
                    {
                        if (pstylefonts != ENfonts)
                        {
                            return false;
                        }
                        else
                        {
                            return true;
                        }
                    }
                    else if (pbasestylefonts != null)
                    {
                        if (pbasestylefonts != ENfonts)
                        {
                            return false;
                        }
                        else
                        {
                            return true;
                        }
                    }
                    else if (Normalfonts != null)
                    {
                        if (Normalfonts != ENfonts)
                        {
                            return false;
                        }
                        else
                        {
                            return true;
                        }
                    }
                    else if (Defaultsfonts != null)
                    {
                        if (Defaultsfonts != ENfonts)
                        {
                            return false;
                        }
                        else
                        {
                            return true;
                        }
                    }
                    return true;
                }
            }
            return true;
        }
    }
}
