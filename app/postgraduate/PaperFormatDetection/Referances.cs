using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using PaperFormatDetection.Frame;
using PaperFormatDetection.Tools;
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Xml;

namespace PaperFormatDetection.Format
{
    class References : ModuleFormat
    {
        //参考文献类型
        private enum RefTypes
        {
            M,//普通图书,
            J, //期刊
            C, //论文集、会议录
            D, //学位论文
            P, //专利文献
            P_OL, //专利文献包含链接地址, 标志为"[P/OL]"
            R, //科技报告
            N, //报纸
            S, //标准
            G, //汇编
            J_OL, //电子文献一种,标志为"[J/OL]"
            EB_OL, //电子文献一种,电子公告，标志为"[EB/OL]"
            C_OL, //电子文献一种,标志为"[C/OL]"
            M_OL, //电子文献一种,标志为"[M/OL]",
            None //不是参考文献的类型
        };
        /* 构造函数 */
        public References(List<Module> modList, PageLocator locator, int masterType)
            : base(modList, locator, masterType)
        {

        }

        /* 继自ModuleFormat中的getStyle方法 */
        public override void getStyle(WordprocessingDocument doc, String fileName)
        {
            string xmlFullPath = fileName + "\\References.xml";//xml模板文件保存路径
            CreateXmlFile(xmlFullPath);
            pageNum = 1;
            getReferencesStyle(doc, xmlFullPath);

        }

        private void CreateXmlFile(string xmlPath)
        {
            XmlDocument xmlDoc = new XmlDocument();
            //创建类型声明节点  
            XmlNode node = xmlDoc.CreateXmlDeclaration("1.0", "utf-8", "");
            xmlDoc.AppendChild(node);
            //创建根节点  
            XmlNode root = xmlDoc.CreateElement("References");
            //以下为结点创建
            XmlElement title = xmlDoc.CreateElement("ReferencesTitle");
            title.SetAttribute("name", "参考文献标题");
            XmlElement err = xmlDoc.CreateElement("spErroInfo");
            err.SetAttribute("name", "特殊错误信息");
            XmlElement part = xmlDoc.CreateElement("partName");
            part.SetAttribute("name", "提示名称");
            XmlElement partText = xmlDoc.CreateElement("Text");
            partText.InnerText = "-----------------参考文献-----------------";
            part.AppendChild(partText);
            root.AppendChild(title);
            root.AppendChild(err);
            root.AppendChild(part);
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

        /* 参考文献检测，主要处理过程 */
        private void getReferencesStyle(WordprocessingDocument doc, String xmlFullPath)
        {

            Body body = doc.MainDocumentPart.Document.Body;
            IEnumerable<Paragraph> paras = body.Elements<Paragraph>();
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(xmlFullPath);
            XmlNode errRoot = xmlDoc.SelectSingleNode("References/spErroInfo");

            //标记参考文献的开始
            bool isRefBegin = false;
            //标记参考文献的结束
            bool isRefEnd = false;
            //标记编号开始
            bool NumberingStart = false;
            int countRef = 0;
            int countRef_J = 0;
            int countCnRef = 0;
            //
            string numberingId = null;
            string ilvl = null;
            //标记电子文献段落是否可能分成两段
            bool twopara = false;
            //遍历每一个Paragraph
            foreach (Paragraph para in paras)
            {
                Run r = para.GetFirstChild<Run>();
                if (r == null) continue;
                string fullText = para.InnerText;
                if (fullText.Trim().Length == 0)
                {
                    continue;//无内容
                }
                //判断参考文献检测起始位置，检测参考文献标题
                if (fullText.Replace(" ", "").Equals("参考文献"))
                {
                    isRefBegin = true;
                    checkReferenceTitle(para, doc, fullText, xmlDoc, errRoot);
                    continue;
                }
                //判断参考文献检测结束
                if (isRefBegin && (fullText.Replace(" ", "").IndexOf("附录") != -1 ||
                    fullText.Replace(" ", "").Equals("攻读硕士学位期间发表学术论文情况") || fullText.Replace(" ", "").Equals("致谢")))
                {
                    isRefEnd = true;
                    //检测参考文献总数量
                    if (countRef < 20)
                    {
                        XmlElement errText = xmlDoc.CreateElement("Text");
                        errText.InnerText = addPageInfo(pageNum) + "参考文献数量少于20篇";
                        errRoot.AppendChild(errText);
                    }
                    //检测期刊数量
                    if (countRef_J < 10)
                    {
                        XmlElement errText = xmlDoc.CreateElement("Text");
                        errText.InnerText = addPageInfo(pageNum) + "期刊数量少于10篇";
                        errRoot.AppendChild(errText);
                    }
                    //检测外文参考文献数量
                    if (countRef - countCnRef < 3)
                    {
                        XmlElement errText = xmlDoc.CreateElement("Text");
                        errText.InnerText = addPageInfo(pageNum) + "外文参考文献数量少于3篇";
                        errRoot.AppendChild(errText);
                    }
                    break;
                }

                //检测参考文献内容的每个Paragraph
                if (isRefBegin == true && isRefEnd == false)
                {
                    //页码定位
                    pageNum = this.getPageNum(pageNum, fullText);
                    if (twopara == true)
                    {
                        checkURL(para, fullText, twopara, xmlDoc, errRoot);
                        twopara = false;
                    }
                    else
                    {
                        countRef++;

                        /* XmlElement xe = xmlDoc.CreateElement("Text");
                         xe.InnerText = fullText+ '['+countRef+']';
                         errRoot.AppendChild(xe);*/

                        //期刊数目统计
                        RefTypes refRype = getRefType(fullText);
                        if (refRype == RefTypes.J)
                        {
                            countRef_J++;
                        }
                        //中文参考文献统计
                        bool isCnRef = hasChinese(fullText);
                        if (isCnRef)
                        {
                            countCnRef++;
                        }
                        if (para.ParagraphProperties != null)
                        {
                            if (para.ParagraphProperties.NumberingProperties != null)
                            {
                                if (para.ParagraphProperties.NumberingProperties.NumberingId != null)
                                {
                                    if (NumberingStart == false || numberingId != para.ParagraphProperties.NumberingProperties.NumberingId.Val || ilvl != para.ParagraphProperties.NumberingProperties.NumberingLevelReference.Val)
                                    {
                                        //参考文献序号检测
                                        isRefNumberingCorrect(fullText, para, doc.MainDocumentPart, countRef, xmlDoc, errRoot);
                                        //编号寻找到编号最开始的段落
                                        NumberingStart = true;
                                    }
                                    numberingId = para.ParagraphProperties.NumberingProperties.NumberingId.Val;
                                    ilvl = para.ParagraphProperties.NumberingProperties.NumberingLevelReference.Val;
                                }
                            }
                            else
                            {
                                //遇到不自动编号的段落将NumberingStart置为false
                                NumberingStart = false;
                                isRefNumberingCorrect(fullText, countRef, xmlDoc, errRoot);
                            }
                        }
                        twopara = checkReferenceParqagraph(para, xmlDoc, errRoot, doc, isCnRef, twopara, refRype);
                    }
                }
            }
            isNumberingCorrectinContents(countRef, doc, xmlDoc, errRoot);
            try
            {
                xmlDoc.Save(xmlFullPath);
            }
            catch (Exception e)
            {
                //显示错误信息  
                Console.WriteLine(e.Message);
            }
        }

        /* 检测参考文献标题 */
        private void checkReferenceTitle(Paragraph para, WordprocessingDocument doc, string paraText, XmlDocument xmlDoc, XmlNode errRoot)
        {

            IEnumerable<Run> pRunList = para.Elements<Run>();
            int spaceCount = Tool.getSpaceCount(paraText);
            //空格判断
            if (spaceCount != 3)
            {
                XmlElement xe = xmlDoc.CreateElement("Text");
                xe.InnerText = "参考文献标题“参考文献”四个字之间应有3个空格";
                errRoot.AppendChild(xe);
            }
            ParagraphProperties pPr = para.GetFirstChild<ParagraphProperties>();
            if (pPr != null)
            {
                if (pPr.GetFirstChild<Justification>() != null)
                {
                    if (pPr.GetFirstChild<Justification>().Val.ToString().ToLower() != "center")
                    {
                        XmlElement errText = xmlDoc.CreateElement("Text");
                        errText.InnerText = "参考文献标题未居中";
                        errRoot.AppendChild(errText);
                    }

                }
            }
            if (!Tool.correctfonts(para, doc, "黑体", "Times New Roman"))
            {
                XmlElement errText = xmlDoc.CreateElement("Text");
                errText.InnerText = "参考文献标题字体错误，应为黑体";
                errRoot.AppendChild(errText);
            }
            if (!Tool.correctsize(para, doc, "30"))
            {
                XmlElement errText = xmlDoc.CreateElement("Text");
                errText.InnerText = "参考文献标题字号错误，应为小三号";
                errRoot.AppendChild(errText);
            }
        }

        /* 检测参考文献内容的Paragraph */
        private bool checkReferenceParqagraph(Paragraph para, XmlDocument xmlDoc, XmlNode errRoot, WordprocessingDocument doc, bool isCnRef, bool twopara, RefTypes refType)
        {
            string paraText = para.InnerText;
            Run run = para.GetFirstChild<Run>();
            IEnumerable<Run> pRunList = para.Elements<Run>();
            ParagraphProperties pPr = null;
            if (run != null)
            {
                pPr = para.GetFirstChild<ParagraphProperties>();
            }
            if (pRunList != null)
            {
                //段前、段后间距和行距的检测
                if (pPr.GetFirstChild<SpacingBetweenLines>() != null)
                {
                    //段前间距
                    if (pPr.GetFirstChild<SpacingBetweenLines>().BeforeLines != null)
                    {
                        if (pPr.GetFirstChild<SpacingBetweenLines>().BeforeLines.Value != 0)
                        {
                            XmlElement errText = xmlDoc.CreateElement("Text");
                            errText.InnerText = addPageInfo(pageNum) + "此条参考文献段前间距错误，应为段前0行：" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                            errRoot.AppendChild(errText);
                        }

                    }
                    //行距
                    if (pPr.GetFirstChild<SpacingBetweenLines>().Line != null)
                    {
                        if (pPr.GetFirstChild<SpacingBetweenLines>().Line.Value != "300")
                        {
                            XmlElement errText = xmlDoc.CreateElement("Text");
                            errText.InnerText = addPageInfo(pageNum) + "此条参考文献行间距错误，应为多倍行距1.25：" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                            errRoot.AppendChild(errText);
                        }
                    }
                    //段后间距
                    if (pPr.GetFirstChild<SpacingBetweenLines>().AfterLines != null)
                    {
                        if (pPr.GetFirstChild<SpacingBetweenLines>().AfterLines.Value != 0)
                        {
                            XmlElement errText = xmlDoc.CreateElement("Text");
                            errText.InnerText = addPageInfo(pageNum) + "此条参考文献段后间距错误，应为段后0行：" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                            errRoot.AppendChild(errText);
                        }
                    }
                }
                //字体检测
                if (!Tool.correctfonts(para, doc, "宋体", "Times New Roman"))
                {
                    XmlElement errText = xmlDoc.CreateElement("Text");
                    errText.InnerText = addPageInfo(pageNum) + "此条参考文献字体错误，应为中文宋体，英文Times New Roman：" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                    errRoot.AppendChild(errText);
                }
                //字号检测
                if (!Tool.correctsize(para, doc, "21"))
                {
                    XmlElement errText = xmlDoc.CreateElement("Text");
                    errText.InnerText = addPageInfo(pageNum) + "此条参考文献字号错误,应为五号字：" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                    errRoot.AppendChild(errText);
                }
                //参考文献类型检测，酌情添加对应类型的处理函数

                switch (refType)
                {

                    case RefTypes.M: return checkRefType_M(para, paraText, xmlDoc, errRoot, isCnRef); //普通图书,
                    case RefTypes.J: return checkRefType_J(para, paraText, xmlDoc, errRoot, isCnRef);//期刊
                    case RefTypes.C: return checkRefType_C(para, paraText, xmlDoc, errRoot, isCnRef); //论文集、会议录
                    case RefTypes.D: return checkRefType_D(para, paraText, xmlDoc, errRoot, isCnRef); //学位论文
                    case RefTypes.P: return checkRefType_P(para, paraText, xmlDoc, errRoot, isCnRef);//专利文献
                    case RefTypes.P_OL: break;//专利文献包含链接地址, 标志为"[P/OL]"
                    case RefTypes.R: break;//科技报告
                    case RefTypes.N: break;//报纸
                    case RefTypes.S: break;//标准
                    case RefTypes.G: break;//汇编
                    case RefTypes.J_OL: return checkRefType_online(para, paraText, xmlDoc, errRoot, isCnRef, twopara);//电子文献一种,标志为"[J/OL]"
                    case RefTypes.EB_OL: return checkRefType_online(para, paraText, xmlDoc, errRoot, isCnRef, twopara);//电子文献一种,电子公告，标志为"[EB/OL]"
                    case RefTypes.C_OL: return checkRefType_online(para, paraText, xmlDoc, errRoot, isCnRef, twopara); //电子文献一种,标志为"[C/OL]"
                    case RefTypes.M_OL: return checkRefType_online(para, paraText, xmlDoc, errRoot, isCnRef, twopara);//电子文献一种,标志为"[M/OL]",
                    case RefTypes.None: return false; //不是参考文献的类型 
                }
            }
            return false;
        }

        /*
          * 普通图书类型检测方法
          * 标志为"M"
          */
        private bool checkRefType_M(Paragraph para, string paraText, XmlDocument xmlDoc, XmlNode errRoot, bool isCnRef)
        {

            if (isCnRef == true)
            {
                string[] textArr = Regex.Split(paraText, @"\[\w*\]");//用中括号分割参考文献条目
                //多做测试这两条
                string refTextBefore = textArr[0];
                string refTextAfter = textArr[1];
                int PLength = paraText.Length;
                //是否以“年份.”结尾
                if (paraText[PLength - 1] != '.')
                {
                    XmlElement errText = xmlDoc.CreateElement("Text");
                    errText.InnerText = addPageInfo(pageNum) + "图书类参考文献应以英文'.'结尾：" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                    errRoot.AppendChild(errText);
                }
                bool IsY = true;
                for (int IsYear = PLength - 5; IsYear <= PLength - 2; IsYear++)
                {
                    if (paraText[IsYear] < 48 || paraText[IsYear] > 57)
                    {
                        IsY = false;
                    }
                }
                if (IsY == false)
                {
                    XmlElement errText = xmlDoc.CreateElement("Text");
                    errText.InnerText = addPageInfo(pageNum) + "图书类参考文献结尾应以“年份.”结尾：" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                    errRoot.AppendChild(errText);
                }
                //是否标注出版社所在城市
                bool HaveC = false;
                for (int HaveCity = 0; HaveCity < paraText.Length; HaveCity++)
                {
                    if (paraText[HaveCity] == ':')
                    {
                        if (HaveCity != 0)
                        {
                            HaveC = true;
                        }
                    }
                }
                if (HaveC == false)
                {
                    XmlElement errText = xmlDoc.CreateElement("Text");
                    errText.InnerText = addPageInfo(pageNum) + "图书类参考文献的文献类型与出版社之间应标注出版社所在城市，格式为“城市：出版社”：" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                    errRoot.AppendChild(errText);
                }
                //是否缺标点
                bool HaveCommas = false;
                for (int HaveComma = 0; HaveComma < paraText.Length; HaveComma++)
                {
                    if (paraText[HaveComma] == ',')
                    {
                        if (HaveComma != 0)
                        {
                            HaveCommas = true;
                        }
                    }
                }
                if (HaveCommas == false)
                {
                    XmlElement errText = xmlDoc.CreateElement("Text");
                    errText.InnerText = addPageInfo(pageNum) + "图书类参考文献缺少标点，在出版社与年份之间应有“,”：" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                    errRoot.AppendChild(errText);
                }
                bool HaveDots = false;
                for (int HaveDot = 0; HaveDot < paraText.Length; HaveDot++)
                {
                    if (paraText[HaveDot] == ',')
                    {
                        if (HaveDot != 0)
                        {
                            HaveDots = true;
                        }
                    }
                }
                if (HaveDots == false)
                {
                    XmlElement errText = xmlDoc.CreateElement("Text");
                    errText.InnerText = addPageInfo(pageNum) + "图书类参考文献缺少标点，在作者与书名之间应有“.”：" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                    errRoot.AppendChild(errText);
                }
            }
            return false;
        }

        /*
         * 期刊类型检测检测方法
         * 标志为"J"
         */
        public bool checkRefType_J(Paragraph para, string paraText, XmlDocument xmlDoc, XmlNode errRoot, bool isCnRef)
        {
            string[] textArr = Regex.Split(paraText, @"\[\w*\]");//用中括号分割参考文献条目
            //多做测试这两条
            string refTextBefore = textArr[0];
            string refTextAfter = textArr[1];
            string txt = Tool.test();
            if (isCnRef == true)
            {
                //页码
                int colonPosition = 0;
                for (int a = paraText.Length - 1; a >= 0; a--)
                {
                    if (colonPosition == 0)
                    {
                        if (paraText[a] == ':')
                        {
                            colonPosition = a;
                        }
                    }
                }
                bool findlines = false;
                if (colonPosition != 0)
                {
                    for (int findline = colonPosition; findline < paraText.Length - 1; findline++)
                    {
                        if (paraText[findline] == '~')
                        {
                            findlines = true;
                        }
                    }

                    if (findlines == true)
                    {
                        XmlElement errText = xmlDoc.CreateElement("Text");
                        errText.InnerText = addPageInfo(pageNum) + "期刊类参考文献页码间应用“-”标识而非“~”：" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                        errRoot.AppendChild(errText);
                    }
                }
                if (colonPosition == 0)
                {
                    XmlElement errText = xmlDoc.CreateElement("Text");
                    errText.InnerText = addPageInfo(pageNum) + "期刊类参考文献应以“:页码范围”结尾”：" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                    errRoot.AppendChild(errText);
                }

                if (colonPosition != 0)
                {
                    //期号
                    bool findissue = false;
                    int findleftbrackets = 0;
                    if (paraText[colonPosition - 1] == ')')
                    {
                        findissue = true;

                        for (int a = colonPosition - 1; a >= 0; a--)
                        {
                            if (findleftbrackets == 0)
                            {
                                if (paraText[a] == '(')
                                {
                                    findleftbrackets = a;
                                }
                            }
                        }
                        int count = paraText.Length;
                        string issue = paraText.Substring(findleftbrackets + 1, (colonPosition - findleftbrackets) - 2);
                        bool Issueisnumber = true;
                        for (int a = 0; a < issue.Length; a++)
                        {
                            if (issue[a] < 48 || issue[a] > 57)
                            {
                                Issueisnumber = false;
                                XmlElement errText = xmlDoc.CreateElement("Text");
                                errText.InnerText = addPageInfo(pageNum) + "期刊类参考文献页码前应有期号，格式为“(number)”：" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                                errRoot.AppendChild(errText);
                            }

                        }
                        if (Issueisnumber == true)
                        {
                            int Issue = Convert.ToInt32(issue);
                            if (Issue > 20)
                            {
                                XmlElement errText = xmlDoc.CreateElement("Text");
                                errText.InnerText = addPageInfo(pageNum) + "警告：期刊类参考文献期号一般不超过20，此条参考文献期号过大：" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                                errRoot.AppendChild(errText);
                            }
                        }
                    }
                    if (paraText[colonPosition - 1] != ')')
                    {
                        XmlElement errText = xmlDoc.CreateElement("Text");
                        errText.InnerText = addPageInfo(pageNum) + "期刊类参考文献页码前应有期号，格式为“(number)”：" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                        errRoot.AppendChild(errText);
                    }

                    //卷号
                    if (findissue == true)
                    {
                        int reelposition = 0;
                        bool findreel = true;
                        for (int a = findleftbrackets; a >= 0; a--)
                        {

                            if (reelposition == 0)
                            {
                                if (paraText[a] == ',')
                                {
                                    reelposition = a;
                                }
                            }
                        }
                        bool Output = false;
                        for (int Isnumber = reelposition + 1; Isnumber < findleftbrackets; Isnumber++)
                        {
                            if (paraText[Isnumber] > 57 || paraText[Isnumber] < 48)
                            {
                                if (paraText[Isnumber] == ' ')
                                {
                                    XmlElement errText = xmlDoc.CreateElement("Text");
                                    errText.InnerText = addPageInfo(pageNum) + "期刊类参考文献期号卷号处不应有空格：" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                                    errRoot.AppendChild(errText);
                                }
                                else
                                {
                                    findreel = false;
                                    if (Output == false)
                                    {
                                        XmlElement errText = xmlDoc.CreateElement("Text");
                                        errText.InnerText = addPageInfo(pageNum) + "期刊类参考文献期号前应有卷号，格式为“,卷号(期号)”：" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                                        errRoot.AppendChild(errText);
                                        Output = true;
                                    }
                                }
                            }
                        }
                        bool AllNumber = false;
                        string reel = paraText.Substring(reelposition + 1, (findleftbrackets - 1) - reelposition);
                        if (reel != "")
                        {
                            AllNumber = true;
                        }

                        for (int a = 0; a < reel.Length; a++)
                        {
                            if (reel[a] < 48 || reel[a] > 57)
                            {
                                AllNumber = false;

                            }
                        }
                        if (findreel == true && AllNumber)
                        {


                            int Reel = Convert.ToInt32(reel);
                            bool findYear = false;
                            int yearposition = 0;
                            for (int a = reelposition; a > 0; a--)
                            {
                                if (yearposition == 0)
                                {
                                    if (paraText[a] == ',')
                                    {
                                        yearposition = a;
                                        findYear = true;
                                    }
                                }
                            }
                            for (int Isyear = yearposition + 1; Isyear < reelposition; Isyear++)
                            {
                                if (paraText[Isyear] < 48 || paraText[Isyear] > 57)
                                {
                                    findYear = false;
                                    XmlElement errText = xmlDoc.CreateElement("Text");
                                    errText.InnerText = addPageInfo(pageNum) + "期刊类参考文献卷号前应标明年份，格式为“,年份,卷号(期号)”：" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                                    errRoot.AppendChild(errText);
                                }
                            }
                            if (findYear == true && yearposition != reelposition)
                            {

                                string year = paraText.Substring(yearposition + 1, (reelposition - 1) - yearposition);
                                int FindFirstComma = 0;
                                bool FindPublishing = false;
                                int Firsty = 0;
                                for (int a = 0; a < refTextAfter.Length; a++)
                                {
                                    if (refTextAfter[a] == ',' && FindFirstComma == 0)
                                    {
                                        FindFirstComma = a;
                                        FindPublishing = true;
                                    }
                                }
                                if (FindPublishing == true)
                                {
                                    string Firstyear = null;
                                    bool findyear = false;
                                    string PublishingHouse = refTextAfter.Substring(0, FindFirstComma - 1);
                                    int FindInTxt = txt.IndexOf(PublishingHouse);
                                    for (int GetFirstYear = FindInTxt; GetFirstYear < txt.Length; GetFirstYear++)
                                    {
                                        if (txt[GetFirstYear] >= 48 && txt[GetFirstYear] <= 57 && findyear == false)
                                        {
                                            findyear = true;
                                            Firstyear = txt.Substring(GetFirstYear, 4);
                                        }
                                    }
                                    for (int a = 0; a < 4; a++)
                                    {
                                        if (Firstyear[a] > 57 || Firstyear[a] < 48)
                                        {
                                            findyear = false;
                                        }
                                    }
                                    if (findyear == true)
                                    {
                                        Firsty = Convert.ToInt32(Firstyear);
                                    }
                                    if (Firsty == Reel)
                                    {

                                    }
                                    if ((Firsty - Reel) <= 10)
                                    {
                                        XmlElement errText = xmlDoc.CreateElement("Text");
                                        errText.InnerText = addPageInfo(pageNum) + "警告：卷号与出版社创刊年份不符：" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                                        errRoot.AppendChild(errText);
                                    }
                                    if ((Firsty - Reel) > 10)
                                    {
                                        XmlElement errText = xmlDoc.CreateElement("Text");
                                        errText.InnerText = addPageInfo(pageNum) + "6)	卷号与出版社标注的卷号相差过大：" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                                        errRoot.AppendChild(errText);
                                    }
                                }
                            }

                        }

                    }
                }


            }
            return false;
        }

        /*
         * 论文集、会议录类型检测方法
         * 标志为"C"
         */
        private bool checkRefType_C(Paragraph para, string paraText, XmlDocument xmlDoc, XmlNode errRoot, bool isCnRef)
        {
            Match match1 = Regex.Match(paraText, @"\[\d+\]");
            if (match1.Success)
            {
                paraText = paraText.Remove(0, match1.Length);
            }
            string[] textArr = Regex.Split(paraText, @"\[\w*\]");//用中括号分割参考文献条目
            int index = textArr[0].IndexOf('.') == -1 ? textArr[0].IndexOf('．') : textArr[0].IndexOf('.');
            if (index <= 0)
            {
                XmlElement xml = xmlDoc.CreateElement("Text");
                xml.InnerText = "作者后应带标点“.”" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                errRoot.AppendChild(xml);
            }
            else
            {
                //超出三位作者加“等”
                string[] autornames = Regex.Split(textArr[0].Substring(0, index), @",");
                if (autornames.Length > 3)
                {
                    if (isCnRef)
                    {
                        if (textArr[0].Substring(0, index).IndexOf('等') == -1)
                        {
                            XmlElement xml = xmlDoc.CreateElement("Text");
                            xml.InnerText = "超出三位作者加“等”" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                            errRoot.AppendChild(xml);
                        }
                    }
                    else
                    {
                        if (textArr[0].Substring(0, index).IndexOf("et al") == -1 && textArr[0].Substring(0, index).IndexOf("etc") == -1)
                        {
                            XmlElement xml = xmlDoc.CreateElement("Text");
                            xml.InnerText = "超出三位作者加et al”" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                            errRoot.AppendChild(xml);
                        }
                    }
                }
                //作者不超过三个时应全部列出
                if (autornames.Length <= 2)
                {
                    if (textArr[0].IndexOf('等') != -1)
                    {
                        XmlElement xml = xmlDoc.CreateElement("Text");
                        xml.InnerText = "作者不超过三个时应全部列出" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                        errRoot.AppendChild(xml);
                    }
                }
            }
            if (isCnRef)//中文类型
            {

                if (textArr.Length == 2)
                {
                    int indexcolon = -1;
                    int indexcomma = -1;
                    //出版地后冒号检测
                    if ((indexcolon = textArr[1].IndexOf(':')) <= 2 && textArr[1].IndexOf('：') <= 2)
                    {
                        XmlElement xml = xmlDoc.CreateElement("Text");
                        xml.InnerText = "缺少出版社（者）或者出版社前缺少标点符号“：”" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                        errRoot.AppendChild(xml);
                    }
                    //年份检测
                    Match match = Regex.Match(textArr[1], @"[1-2][0-9][0-9][0-9]");
                    if (match.Success)
                    {
                        int year = Convert.ToInt32(match.Value);
                        DateTime now = DateTime.Now;
                        if (year > now.Year)
                        {
                            XmlElement xml = xmlDoc.CreateElement("Text");
                            xml.InnerText = "超出当前年份，不合法" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                            errRoot.AppendChild(xml);
                        }
                        //出版社后逗号检测
                        indexcomma = textArr[1].IndexOf(',') == -1 ? textArr[1].IndexOf('，') : textArr[1].IndexOf(',');
                        if (indexcomma != match.Index - 1 && indexcomma != match.Index - 2)
                        {
                            XmlElement xml = xmlDoc.CreateElement("Text");
                            xml.InnerText = "日期前缺少标点符号“，”" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                            errRoot.AppendChild(xml);
                        }
                    }
                    else
                    {
                        XmlElement xml = xmlDoc.CreateElement("Text");
                        xml.InnerText = "年份不合法" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                        errRoot.AppendChild(xml);
                    }
                    //句末英文句号
                    if (paraText.Trim()[paraText.Trim().Length - 1] != '.')
                    {
                        XmlElement xml = xmlDoc.CreateElement("Text");
                        xml.InnerText = "建议参考文献的结尾用“.”号" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                        errRoot.AppendChild(xml);
                    }
                }
                else if (textArr.Length == 3)//出版者不详  s.n.(英文)
                {
                    if (paraText.IndexOf("出版者不详") == -1 && paraText.IndexOf("出版地不详") == -1)
                    {
                        XmlElement xml = xmlDoc.CreateElement("Text");
                        xml.InnerText = "多出一个“[]”或者“[]”内的内容不规范”" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                        errRoot.AppendChild(xml);
                    }
                    int indexcolon = -1;
                    int indexcomma = -1;
                    //出版地后冒号检测
                    if ((indexcolon = textArr[1].IndexOf(':')) <= 2 && textArr[1].IndexOf('：') <= 2)
                    {
                        XmlElement xml = xmlDoc.CreateElement("Text");
                        xml.InnerText = "缺少出版社或者缺少标点符号“：”" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                        errRoot.AppendChild(xml);
                    }

                    Match match = Regex.Match(textArr[2], @"[1-2][0-9][0-9][0-9]");
                    if (match.Success)
                    {
                        int year = Convert.ToInt32(match.Value);
                        DateTime now = DateTime.Now;
                        if (year > now.Year)
                        {
                            XmlElement xml = xmlDoc.CreateElement("Text");
                            xml.InnerText = "超出当前年份，不合法" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                            errRoot.AppendChild(xml);
                        }
                        //出版社后逗号检测
                        indexcomma = textArr[2].IndexOf(',') == -1 ? textArr[2].IndexOf('，') : textArr[2].IndexOf(',');
                        if (indexcomma != match.Index - 1 && indexcomma != match.Index - 2)
                        {
                            XmlElement xml = xmlDoc.CreateElement("Text");
                            xml.InnerText = "日期前缺少标点符号“，”" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                            errRoot.AppendChild(xml);
                        }
                    }
                    else
                    {
                        XmlElement xml = xmlDoc.CreateElement("Text");
                        xml.InnerText = "年份不合法" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                        errRoot.AppendChild(xml);
                    }
                    //句末句号
                    if (paraText.Trim()[paraText.Trim().Length - 1] != '.')
                    {
                        XmlElement xml = xmlDoc.CreateElement("Text");
                        xml.InnerText = "建议参考文献的结尾用“.”号" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                        errRoot.AppendChild(xml);
                    }
                }
                else
                {
                    XmlElement xml = xmlDoc.CreateElement("Text");
                    xml.InnerText = "多出一个“[]”" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                    errRoot.AppendChild(xml);
                    //句末句号
                    if (paraText.Trim()[paraText.Trim().Length - 1] != '.')
                    {
                        xml = xmlDoc.CreateElement("Text");
                        xml.InnerText = "建议参考文献的结尾用“.”号" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                        errRoot.AppendChild(xml);
                    }
                }
            }
            else//英文类型
            {
                if (textArr.Length == 2)
                {
                    int indexcolon = -1;
                    int indexcomma = -1;
                    //出版地后冒号检测
                    if ((indexcolon = textArr[1].IndexOf(':')) <= 2 && textArr[1].IndexOf('：') <= 2)
                    {
                        XmlElement xml = xmlDoc.CreateElement("Text");
                        xml.InnerText = "缺少出版社（者）或者出版社前缺少标点符号“：”" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                        errRoot.AppendChild(xml);
                    }
                    //年份检测
                    Match match = Regex.Match(textArr[1], @"[1-2][0-9][0-9][0-9]");
                    if (match.Success)
                    {
                        int year = Convert.ToInt32(match.Value);
                        DateTime now = DateTime.Now;
                        if (year > now.Year)
                        {
                            XmlElement xml = xmlDoc.CreateElement("Text");
                            xml.InnerText = "超出当前年份，不合法" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                            errRoot.AppendChild(xml);
                        }
                        //出版社后逗号检测
                        indexcomma = textArr[1].IndexOf(',') == -1 ? textArr[1].IndexOf('，') : textArr[1].IndexOf(',');
                        if (indexcomma != match.Index - 1 && indexcomma != match.Index - 2)
                        {
                            XmlElement xml = xmlDoc.CreateElement("Text");
                            xml.InnerText = "日期前缺少标点符号“，”" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                            errRoot.AppendChild(xml);
                        }
                    }
                    else
                    {
                        XmlElement xml = xmlDoc.CreateElement("Text");
                        xml.InnerText = "年份不合法" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                        errRoot.AppendChild(xml);
                    }
                    //句末英文句号
                    if (paraText.Trim()[paraText.Trim().Length - 1] != '.')
                    {
                        XmlElement xml = xmlDoc.CreateElement("Text");
                        xml.InnerText = "建议参考文献的结尾用“.”号" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                        errRoot.AppendChild(xml);
                    }
                }
                else if (textArr.Length == 3)//s.n.(英文)
                {
                    if (paraText.IndexOf("s.n.") == -1)
                    {
                        XmlElement xml = xmlDoc.CreateElement("Text");
                        xml.InnerText = "多出一个“[]”" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                        errRoot.AppendChild(xml);
                    }
                    int indexcolon = -1;
                    int indexcomma = -1;
                    //出版地后冒号检测
                    if ((indexcolon = textArr[2].IndexOf(':')) <= 2 && textArr[1].IndexOf('：') <= 2)
                    {
                        XmlElement xml = xmlDoc.CreateElement("Text");
                        xml.InnerText = "缺少出版社（者）或者出版社前缺少标点符号“：”" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                        errRoot.AppendChild(xml);
                    }
                    //年份检测
                    Match match = Regex.Match(textArr[2], @"[1-2][0-9][0-9][0-9]");
                    if (match.Success)
                    {
                        int year = Convert.ToInt32(match.Value);
                        DateTime now = DateTime.Now;
                        if (year > now.Year)
                        {
                            XmlElement xml = xmlDoc.CreateElement("Text");
                            xml.InnerText = "超出当前年份，不合法" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                            errRoot.AppendChild(xml);
                        }
                        //出版社后逗号检测
                        indexcomma = textArr[2].IndexOf(',') == -1 ? textArr[2].IndexOf('，') : textArr[2].IndexOf(',');
                        if (indexcomma != match.Index - 1 && indexcomma != match.Index - 2)
                        {
                            XmlElement xml = xmlDoc.CreateElement("Text");
                            xml.InnerText = "日期前缺少标点符号“，”" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                            errRoot.AppendChild(xml);
                        }
                    }
                    else
                    {
                        XmlElement xml = xmlDoc.CreateElement("Text");
                        xml.InnerText = "年份不合法" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                        errRoot.AppendChild(xml);
                    }
                    //句末英文句号
                    if (paraText.Trim()[paraText.Trim().Length - 1] != '.')
                    {
                        XmlElement xml = xmlDoc.CreateElement("Text");
                        xml.InnerText = "建议参考文献的结尾用“.”号" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                        errRoot.AppendChild(xml);
                    }

                }
                else
                {
                    XmlElement xml = xmlDoc.CreateElement("Text");
                    xml.InnerText = "多出一个“[]”" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                    errRoot.AppendChild(xml);
                    //句末句号
                    if (paraText.Trim()[paraText.Trim().Length - 1] != '.')
                    {
                        xml = xmlDoc.CreateElement("Text");
                        xml.InnerText = "建议参考文献的结尾用“.”号" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                        errRoot.AppendChild(xml);
                    }
                }
            }
            return false;
        }

        /*
         * 学位论文类型检测方法
         * 标志为"D"
         */
        private bool checkRefType_D(Paragraph para, string paraText, XmlDocument xmlDoc, XmlNode errRoot, bool isCnRef)
        {
            Match match1 = Regex.Match(paraText, @"\[\d+\]");
            if (match1.Success)
            {
                paraText = paraText.Remove(0, match1.Length);
            }
            string[] textArr = Regex.Split(paraText, @"\[\w*\]");//用中括号分割参考文献条目
            int index = textArr[0].IndexOf('.');
            if (index <= 0)
            {
                XmlElement xml = xmlDoc.CreateElement("Text");
                xml.InnerText = "作者后应带标点“.”" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                errRoot.AppendChild(xml);
            }
            else
            {
                //超出三位作者加“等”
                string[] autornames = Regex.Split(textArr[0].Substring(0, index), @",");
                if (autornames.Length > 3)
                {
                    if (isCnRef)
                    {
                        if (textArr[0].Substring(0, index).IndexOf('等') == -1)
                        {
                            XmlElement xml = xmlDoc.CreateElement("Text");
                            xml.InnerText = "超出三位作者加“等”" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                            errRoot.AppendChild(xml);
                        }
                    }
                    else
                    {
                        if (textArr[0].Substring(0, index).IndexOf("et al") == -1 && textArr[0].Substring(0, index).IndexOf("etc") == -1)
                        {
                            XmlElement xml = xmlDoc.CreateElement("Text");
                            xml.InnerText = "超出三位作者加et al”" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                            errRoot.AppendChild(xml);
                        }
                    }
                }
                //作者不超过三个时应全部列出
                if (autornames.Length <= 2)
                {
                    if (textArr[0].IndexOf('等') != -1)
                    {
                        XmlElement xml = xmlDoc.CreateElement("Text");
                        xml.InnerText = "作者不超过三个时应全部列出" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                        errRoot.AppendChild(xml);
                    }
                }
            }
            if (isCnRef)//中文
            {
                if (textArr.Length == 2)
                {
                    int indexcolon = -1;
                    int indexcomma = -1;
                    //学院院系前缺少标点符号
                    if ((indexcolon = textArr[1].IndexOf(':')) <= 2 && textArr[1].IndexOf('：') <= 2)
                    {
                        XmlElement xml = xmlDoc.CreateElement("Text");
                        xml.InnerText = "学院院系前缺少标点符号“：”" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                        errRoot.AppendChild(xml);
                    }
                    //年份检测
                    Match match = Regex.Match(textArr[1], @"[1-2][0-9][0-9][0-9]");
                    if (match.Success)
                    {
                        int year = Convert.ToInt32(match.Value);
                        DateTime now = DateTime.Now;
                        if (year > now.Year)
                        {
                            XmlElement xml = xmlDoc.CreateElement("Text");
                            xml.InnerText = "超出当前年份，不合法" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                            errRoot.AppendChild(xml);
                        }
                        //出版社后逗号检测
                        indexcomma = textArr[1].IndexOf(',') == -1 ? textArr[1].IndexOf('，') : textArr[1].IndexOf(',');
                        if (indexcomma != match.Index - 1 && indexcomma != match.Index - 2)
                        {
                            XmlElement xml = xmlDoc.CreateElement("Text");
                            xml.InnerText = "日期前缺少标点符号“，”" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                            errRoot.AppendChild(xml);
                        }
                        //学院，系不能缺少
                        string school = textArr[1].Substring(indexcolon + 1);
                        if (school.IndexOf("学院") == -1 && school.IndexOf("系") == -1)
                        {
                            XmlElement xml = xmlDoc.CreateElement("Text");
                            xml.InnerText = "不能缺少院系" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                            errRoot.AppendChild(xml);
                        }
                    }
                    else
                    {
                        XmlElement xml = xmlDoc.CreateElement("Text");
                        xml.InnerText = "年份不合法" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                        errRoot.AppendChild(xml);
                    }
                    //句末句号
                    if (paraText.Trim()[paraText.Trim().Length - 1] != '.')
                    {
                        XmlElement xml = xmlDoc.CreateElement("Text");
                        xml.InnerText = "建议参考文献的结尾用“.”号" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                        errRoot.AppendChild(xml);
                    }
                }
                else
                {
                    XmlElement xml = xmlDoc.CreateElement("Text");
                    xml.InnerText = "多“[]”" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                    errRoot.AppendChild(xml);
                }
            }
            else//英文
            {
                int indexcolon = -1;
                int indexcomma = -1;
                //出版地后冒号检测
                if ((indexcolon = textArr[1].IndexOf(':')) <= 2 && textArr[1].IndexOf('：') <= 2)
                {
                    XmlElement xml = xmlDoc.CreateElement("Text");
                    xml.InnerText = "缺少出版社（者）或者出版社前缺少标点符号“：”" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                }
                //年份检测
                Match match = Regex.Match(textArr[1], @"[1-2][0-9][0-9][0-9]");
                if (match.Success)
                {
                    int year = Convert.ToInt32(match.Value);
                    DateTime now = DateTime.Now;
                    if (year > now.Year)
                    {
                        XmlElement xml = xmlDoc.CreateElement("Text");
                        xml.InnerText = "超出当前年份，不合法" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                    }
                    //出版社后逗号检测
                    indexcomma = textArr[1].IndexOf(',') == -1 ? textArr[1].IndexOf('，') : textArr[1].IndexOf(',');
                    if (indexcomma != match.Index - 1 && indexcomma != match.Index - 2)
                    {
                        XmlElement xml = xmlDoc.CreateElement("Text");
                        xml.InnerText = "日期前缺少标点符号“，”" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                    }
                    //学院，系不能缺少
                    /* string school = textArr[1].Substring(indexcolon + 1);
                     if (school.IndexOf("学院") == -1)
                     {
                         XmlElement xml = xmlDoc.CreateElement("Text");
                         xml.InnerText = "院系不能缺少" + paraText.Substring(0, 10);
                         errRoot.AppendChild(xml);
                     }*/
                }
                else
                {
                    XmlElement xml = xmlDoc.CreateElement("Text");
                    xml.InnerText = "年份不合法" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                }
                //句末英文句号
                if (paraText.Trim()[paraText.Trim().Length - 1] != '.')
                {
                    XmlElement xml = xmlDoc.CreateElement("Text");
                    xml.InnerText = "建议参考文献的结尾用“.”号" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                    errRoot.AppendChild(xml);
                }
            }
            return false;
        }

        /*
         * 专利文献类型检测方法
         * 标志为"P"
         */
        private bool checkRefType_P(Paragraph para, string paraText, XmlDocument xmlDoc, XmlNode errRoot, bool isCnRef)
        {
            return false;
        }
        private bool checkRefType_online(Paragraph para, string paraText, XmlDocument xmlDoc, XmlNode errRoot, bool isCnRef, bool twopara)
        {
            Match match1 = Regex.Match(paraText, @"\[\d+\]");
            if (match1.Success)
            {
                paraText = paraText.Remove(0, match1.Length);
            }
            string[] textArr = Regex.Split(paraText, @"\[\w*\]");//用中括号分割参考文献条目
            if (inTwoParagraph(para, paraText))//可能分成两段
            {
                twopara = true;//将twopara置为true              
            }
            int index = textArr[0].IndexOf('.');
            if (index <= 0)
            {
                XmlElement xml = xmlDoc.CreateElement("Text");
                xml.InnerText = "作者后应带标点“.”" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                errRoot.AppendChild(xml);
            }
            else
            {
                //超出三位作者加“等”
                string[] autornames = Regex.Split(textArr[0].Substring(0, index), @",");
                if (autornames.Length > 3)
                {
                    if (isCnRef)
                    {
                        if (textArr[0].Substring(0, index).IndexOf('等') == -1)
                        {
                            XmlElement xml = xmlDoc.CreateElement("Text");
                            xml.InnerText = "超出三位作者加“等”" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                            errRoot.AppendChild(xml);
                        }
                    }
                    else
                    {
                        if (textArr[0].Substring(0, index).IndexOf("et al") == -1 && textArr[0].Substring(0, index).IndexOf("etc") == -1)
                        {
                            XmlElement xml = xmlDoc.CreateElement("Text");
                            xml.InnerText = "超出三位作者加et al”" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                            errRoot.AppendChild(xml);
                        }
                    }
                }
                //作者不超过三个时应全部列出
                if (autornames.Length <= 2)
                {
                    if (textArr[0].IndexOf('等') != -1)
                    {
                        XmlElement xml = xmlDoc.CreateElement("Text");
                        xml.InnerText = "作者不超过三个时应全部列出" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                        errRoot.AppendChild(xml);
                    }
                }
                if (!isCnRef)
                {
                    //英文作者首字母大写
                    for (int i = 0; i < autornames.Length; i++)
                    {
                        //找到第一个字母位置
                        Match match = Regex.Match(autornames[i], @"\w");
                        if (match.Index != -1)
                        {
                            if (autornames[i][match.Index] <= 'A' || autornames[i][match.Index] >= 'Z')
                            {
                                XmlElement xml = xmlDoc.CreateElement("Text");
                                xml.InnerText = "英文作者首字母应大写”" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}" + autornames[i].Substring(match.Index);
                                errRoot.AppendChild(xml);
                            }
                        }
                        else
                        {
                            XmlElement xml = xmlDoc.CreateElement("Text");
                            xml.InnerText = "缺少作者”" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}" + autornames[i];
                            errRoot.AppendChild(xml);
                        }
                    }
                }
            }
            return twopara;
        }
        //分成两段了
        private bool inTwoParagraph(Paragraph p, string paraText)
        {
            if (paraText.IndexOf("http://www.") >= 0 || paraText.IndexOf("www.") >= 0)
            {
                return false;
            }
            else
            {
                return true;
            }
        }
        private void checkURL(Paragraph p, string paraText, bool twopara, XmlDocument xmlDoc, XmlNode errRoot)
        {
            twopara = false;
            if (paraText.IndexOf("http://www.") == -1 && paraText.IndexOf("www.") == -1)
            {
                XmlElement xml = xmlDoc.CreateElement("Text");
                xml.InnerText = "OnLine类型参考文献应包含网址" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "。。。的上一段}";
                errRoot.AppendChild(xml);
            }
        }

        /* 检测是否以“[数字]”开头 */
        //此函数错误，因为编号不一定在str里
        /*  private bool isRefStart(string str)
          {
              return Regex.IsMatch(str, @"^\[([0-9]*[1-9][0-9]*)\]");
          }
    */

        /* 检测参考文献序号是否正确 */
        private void isRefNumberingCorrect(string paraText, Paragraph p, MainDocumentPart Mpart, int count, XmlDocument xmlDoc, XmlNode errRoot)
        {
            if (p.ParagraphProperties.NumberingProperties.NumberingId != null)
            {
                string numberingId = p.ParagraphProperties.NumberingProperties.NumberingId.Val;
                string ilvl = p.ParagraphProperties.NumberingProperties.NumberingLevelReference.Val;
                NumberingDefinitionsPart numberingDefinitionsPart1 = Mpart.NumberingDefinitionsPart;
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
                                        Match match = Regex.Match(levelText1, @"\[\%[0-9]+\]");
                                        if (match.Success == false)
                                        {
                                            XmlElement xml = xmlDoc.CreateElement("Text");
                                            xml.InnerText = "建议此条参考文献编号应为：[" + count + "]:" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + '}';
                                            errRoot.AppendChild(xml);
                                        }
                                        if (level.StartNumberingValue != null)
                                        {
                                            if (level.StartNumberingValue.Val != Convert.ToUInt32(count))
                                            {
                                                XmlElement xml = xmlDoc.CreateElement("Text");
                                                xml.InnerText = "参考文献编号错误，应为：[" + count + "]:" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + '}';
                                                errRoot.AppendChild(xml);
                                            }
                                        }
                                        break;
                                    }
                                }
                                return;
                            }
                        }
                    }
                }
            }
        }
        private void isRefNumberingCorrect(string paraText, int count, XmlDocument xmlDoc, XmlNode errRoot)
        {
            if (paraText.IndexOf(Convert.ToString(count)) == -1)
            {
                XmlElement xml = xmlDoc.CreateElement("Text");
                xml.InnerText = "此条参考文献忘记编号或编号错误，应为：[" + count + "]:" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + '}';
                errRoot.AppendChild(xml);
            }
            else
            {
                if (paraText.Trim().IndexOf('[' + Convert.ToString(count) + ']') != 0)
                {
                    XmlElement xml = xmlDoc.CreateElement("Text");
                    xml.InnerText = "建议此条参考文献编号应为：[" + count + "]:" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + '}';
                    errRoot.AppendChild(xml);
                }
            }
        }
        /* 检测字符串是否包含中文 */
        private bool hasChinese(string str)
        {
            return Regex.IsMatch(str, @"[\u4e00-\u9fa5]");
        }

        /* 获取参考文献类型 */
        private RefTypes getRefType(string paraText)
        {

            Match match = Regex.Match(paraText, @"\[[A-Z](/OL)?\]");
            if (match.Success)
            {
                string type = match.Groups[0].Value;
                string typenormal = null;
                typenormal = type.Substring(1, type.Length - 2);
                switch (typenormal)
                {
                    case "M": Console.WriteLine("boom!"); return RefTypes.M; //普通图书,
                    case "J": return RefTypes.J; //期刊
                    case "C": return RefTypes.C; //论文集、会议录
                    case "D": return RefTypes.D; //学位论文
                    case "P": return RefTypes.P; //专利文献
                    case "P_OL": return RefTypes.P_OL; //专利文献包含链接地址, 标志为"[P/OL]"
                    case "R": return RefTypes.R; //科技报告
                    case "N": return RefTypes.N; //报纸
                    case "S": return RefTypes.S; //标准
                    case "G": return RefTypes.G; //汇编
                    case "J/OL": return RefTypes.J_OL; //电子文献一种,标志为"[J/OL]"
                    case "EB/OL": return RefTypes.EB_OL; //电子文献一种,电子公告，标志为"[EB/OL]"
                    case "C/OL": return RefTypes.C_OL; //电子文献一种,标志为"[C/OL]"
                    case "M/OL": return RefTypes.M_OL; //电子文献一种,标志为"[M/OL]",
                    default: return RefTypes.None; //不是参考文献的类型
                }
            }
            return RefTypes.None;
        }
        private void isNumberingCorrectinContents(int RefCount, WordprocessingDocument doc, XmlDocument xmlDoc, XmlNode errRoot)
        {
            IEnumerable<Paragraph> paras = doc.MainDocumentPart.Document.Body.Elements<Paragraph>();
            List<int> list = new List<int>();
            int maxnumber = 0;
            int location = -1;
            foreach (Paragraph para in paras)
            {
                location++;
                string runText = null;
                IEnumerable<Run> runs = para.Elements<Run>();
                foreach (Run run in runs)
                {
                    if (run.RunProperties != null)
                    {
                        if (run.RunProperties.VerticalTextAlignment != null)
                        {
                            runText += run.InnerText;
                            Match match = Regex.Match(runText, @"\[\d+\-*\d*\]");
                            if (match.Success)
                            {
                                int index = match.Value.IndexOf('-');
                                if (index == -1)
                                {
                                    if (Convert.ToInt16(runText.Substring(match.Index + 1, match.Length - 2)) > maxnumber + 1)
                                    {
                                        List<int> listchapter = Tool.getTitlePosition(doc);
                                        string chapter = Chapter(listchapter, location, doc.MainDocumentPart.Document.Body);
                                        XmlElement xml = xmlDoc.CreateElement("Text");
                                        xml.InnerText = "此段落参考文献角标错误，建议为：[" + (maxnumber + 1) + "]:" + "{" + para.InnerText.Substring(0, para.InnerText.Length < 20 ? para.InnerText.Length : 20) + '}' + chapter;
                                        errRoot.AppendChild(xml);
                                        maxnumber++;
                                    }
                                    else if (Convert.ToInt16(runText.Substring(match.Index + 1, match.Length - 2)) == maxnumber + 1)
                                    {
                                        maxnumber++;
                                    }
                                    if (Convert.ToInt16(runText.Substring(match.Index + 1, match.Length - 2)) > RefCount)
                                    {
                                        List<int> listchapter = Tool.getTitlePosition(doc);
                                        string chapter = Chapter(listchapter, location, doc.MainDocumentPart.Document.Body);
                                        XmlElement xml = xmlDoc.CreateElement("Text");
                                        xml.InnerText = "此段落参考文献角标超过总参考文献数目" + "{" + para.InnerText.Substring(0, para.InnerText.Length < 20 ? para.InnerText.Length : 20) + '}' + chapter;
                                        errRoot.AppendChild(xml);
                                    }
                                }
                                else
                                {
                                    //[m-n]
                                    //m
                                    string number1 = match.Value.Substring(1, index - 1);
                                    if (Convert.ToInt16(number1) > maxnumber + 1)
                                    {
                                        List<int> listchapter = Tool.getTitlePosition(doc);
                                        string chapter = Chapter(listchapter, location, doc.MainDocumentPart.Document.Body);
                                        XmlElement xml = xmlDoc.CreateElement("Text");
                                        xml.InnerText = "此段落参考文献角标错误，建议为：[" + (maxnumber + 1) + "-*]:" + "{" + para.InnerText.Substring(0, para.InnerText.Length < 20 ? para.InnerText.Length : 20) + '}' + chapter;
                                        errRoot.AppendChild(xml);
                                        maxnumber = maxnumber + 2;
                                        continue;
                                    }
                                    else if (Convert.ToInt16(number1) == maxnumber + 1)
                                    {
                                        maxnumber++;
                                    }
                                    if (Convert.ToInt16(number1) > RefCount)
                                    {
                                        List<int> listchapter = Tool.getTitlePosition(doc);
                                        string chapter = Chapter(listchapter, location, doc.MainDocumentPart.Document.Body);
                                        XmlElement xml = xmlDoc.CreateElement("Text");
                                        xml.InnerText = "此段落参考文献角标超过总参考文献数目" + "{" + para.InnerText.Substring(0, para.InnerText.Length < 20 ? para.InnerText.Length : 20) + '}' + chapter;
                                        errRoot.AppendChild(xml);
                                        continue;
                                    }
                                    //n
                                    maxnumber = Convert.ToInt16(match.Value.Substring(index + 1, match.Length - (index + 2)));
                                    if (maxnumber > RefCount)
                                    {
                                        List<int> listchapter = Tool.getTitlePosition(doc);
                                        string chapter = Chapter(listchapter, location, doc.MainDocumentPart.Document.Body);
                                        XmlElement xml = xmlDoc.CreateElement("Text");
                                        xml.InnerText = "此段落参考文献角标超过总参考文献数目" + "{" + para.InnerText.Substring(0, para.InnerText.Length < 20 ? para.InnerText.Length : 20) + '}' + chapter;
                                        errRoot.AppendChild(xml);
                                    }
                                }
                                runText = null;
                            }
                        }
                    }
                }
            }
            if (maxnumber < RefCount - 1)
            {
                XmlElement xml = xmlDoc.CreateElement("Text");
                xml.InnerText = "正文中缺少参考文献角标，请补全";
                errRoot.AppendChild(xml);
            }
        }
        static string Chapter(List<int> titlePosition, int location, Body body)
        {
            string chapter = "";
            int titlelocation = 0;
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
    }
}
