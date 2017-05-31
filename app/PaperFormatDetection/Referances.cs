using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Linq;
using PaperFormatDetection.Frame;
using PaperFormatDetection.Tools;
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Xml;
using System.IO;
using DocumentFormat.OpenXml;

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
            getReferencesStyle(doc,fileName);

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
        private void getReferencesStyle(WordprocessingDocument doc,string fileName)
        {

            Body body = doc.MainDocumentPart.Document.Body;
            IEnumerable<Paragraph> paras = body.Elements<Paragraph>();
            XmlDocument xmlDoc = new XmlDocument();
            string xmlFullPath = fileName + "\\References.xml";//xml模板文件保存路径
           xmlDoc.Load(xmlFullPath);
            XmlNode errRoot = xmlDoc.SelectSingleNode("References/spErroInfo");
            List<Paragraph> plist = toList(paras);
            int number = -1;
            Paragraph referenceTitle=new Paragraph();
            bool NewItem = false;

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
            bool two = false;
            //遍历每一个Paragraph
            foreach (Paragraph para in paras)
            {
                number++;
                Run r = para.GetFirstChild<Run>();
                if (r == null) continue;
                string fullText = para.InnerText;
                if (fullText.Trim().Length == 0)
                {

                    continue;//无内容
                }
                //判断参考文献检测起始位置，检测参考文献标题
                if (fullText.Replace(" ", "").Equals("参考文献") || fullText.Replace(" ", "").Equals("考文献"))
                {
                    referenceTitle = para;
                    isRefBegin = true;
                    checkReferenceTitle(para, doc, fullText,xmlDoc,errRoot);
                   
                    continue;
                }
                //判断参考文献检测结束
                if (isRefBegin && (fullText.Replace(" ", "").IndexOf("附录") != -1 ||
                    fullText.Replace(" ", "").Equals("攻读硕士学位期间发表学术论文情况") || fullText.Replace(" ", "").Equals("致谢") || fullText.Trim().Equals("致谢") || fullText.Replace(" ", "").Equals("致谢")))
                {
                    isRefEnd = true;
                    //检测参考文献总数量
                    if (countRef < 20)
                    {

                        XmlElement errText = xmlDoc.CreateElement("Text");
                        errText.InnerText = "参考文献数量少于20篇";
                        errRoot.AppendChild(errText);
                        addComment(doc,referenceTitle , "参考文献数量少于20篇");
                    }
                    //检测期刊数量
                    if (countRef_J < 10)
                    {
                        XmlElement errText = xmlDoc.CreateElement("Text");
                        errText.InnerText =   "期刊数量少于10篇";
                        errRoot.AppendChild(errText);
                        addComment(doc, referenceTitle, "期刊数量少于10篇");
                    }
                    //检测外文参考文献数量
                    if (countRef - countCnRef < 3)
                    {
                        XmlElement errText = xmlDoc.CreateElement("Text");
                        errText.InnerText = "外文参考文献数量少于3篇";
                        errRoot.AppendChild(errText);
                       
                        addComment(doc, referenceTitle, "外文参考文献数量少于3篇");
                        
                    }
                    break;
                }

                //检测参考文献内容的每个Paragraph
                if (isRefBegin == true && isRefEnd == false)
                {
                    NewItem = false;
                    if (number + 1 < plist.Count())
                    {
                        if (plist[number + 1].ParagraphProperties != null)
                        {
                            if (!two && plist[number + 1].ParagraphProperties.NumberingProperties == null &&
                                plist[number + 1].InnerText.IndexOf("[" + (countRef + 2).ToString() + "]") == -1)
                            {
                                /*Match match1 = Regex.Match(para.InnerText, @"\[[0-9](/OL)?\]");
                                Match match2 = Regex.Match(para.InnerText, @"\[(1|2|3|4)[0-9](/OL)?\]");
                                if (match1.Success)
                                    NewItem = true;
                                else if (match2.Success)
                                    NewItem = true;
                                else if (plist[number + 1].InnerText.IndexOf("[]") != -1)
                                    NewItem = true;
                                else { }*/
                                if (para.InnerText.Trim()[para.InnerText.Trim().Length - 1] != '.')
                                { }
                                else
                                {

                                    NewItem = true;
                                }


                                if (!NewItem)
                                {
                                    fullText = para.InnerText + plist[number + 1].InnerText;
                                    two = true;
                                }
                            }
                        }
                    }
                    

                    
                    if (twopara)
                    {
                        
                        two = false;
                        twopara = false;
                        continue;
                    }
                    else
                    {
                        countRef++;

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
                                        //isRefNumberingCorrect(fullText, para, doc.MainDocumentPart, countRef, xmlDoc, errRoot);
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
                                isRefNumberingCorrect(para, countRef);
                            }
                        }
                        if (two)
                        {
                            checkReferenceParqagraph(para, plist[number + 1], doc, isCnRef, twopara, refRype,xmlDoc,errRoot);
                            twopara = true;
                        }
                        else
                        {
                             checkReferenceParqagraph(para, null, doc, isCnRef, twopara, refRype,xmlDoc,errRoot);
                             twopara = false;
                        }
                    }
                }
            }
            isNumberingCorrectinContents(countRef, doc);
            xmlDoc.Save(xmlFullPath);
          
        }

        /* 检测参考文献标题 */
        private void checkReferenceTitle(Paragraph para, WordprocessingDocument doc, string paraText,XmlDocument xmlDoc,XmlNode errRoot)
        {

            bool flag1 = true;
            bool flag2 = true;
            IEnumerable<Run> pRunList = para.Elements<Run>();
            int spaceCount = Tool.getSpaceCount(paraText);
            //空格判断
            if (spaceCount != 3||!paraText.Trim().Equals("参考文献"))
            {
                XmlElement errText = xmlDoc.CreateElement("Text");
                errText.InnerText =  "参考文献标题四字及间距有误";
                errRoot.AppendChild(errText);
                
                if (para.Elements<Run>().Count() == 1)
                {
                    para.GetFirstChild<Run>().GetFirstChild<Text>().Text = "参 考 文 献";
                }
                else
                {
                    IEnumerable<Run> runs = para.Elements<Run>();
                    int num = 0;
                    foreach (Run rr in runs)
                    {

                        num++;
                        if (num != 1)
                        {
                            if(rr.GetFirstChild<Text>()!=null)
                        
                            rr.GetFirstChild<Text>().Text = null;
                        }
                    }
                    if (para.GetFirstChild<Run>().GetFirstChild<Text>() != null)
                        para.GetFirstChild<Run>().GetFirstChild<Text>().Text = "参 考 文 献";

                    else
                    {
                        Text text1 = new Text();
                        text1.Text = "参 考 文 献";
                        para.GetFirstChild<Run>().Append(text1);
                    }
                        

                    
                }

            }
            ParagraphProperties pPr = para.GetFirstChild<ParagraphProperties>();
            if (pPr != null)
            {
                if (pPr.GetFirstChild<Justification>() != null)
                {
                    if (pPr.GetFirstChild<Indentation>() != null &&
                           pPr.GetFirstChild<Indentation>() != null)
                    {
                        if (pPr.GetFirstChild<Indentation>().FirstLine != null &&
                           pPr.GetFirstChild<Indentation>().FirstLineChars != null)
                        {
                            if (pPr.GetFirstChild<Indentation>().FirstLine == "3392" &&
                               pPr.GetFirstChild<Indentation>().FirstLineChars == "1060")
                            { }
                        }
                    }
                    else if (pPr.GetFirstChild<Justification>().Val.ToString().ToLower() != "center")
                    {
                      
                        pPr.GetFirstChild<Justification>().Val = JustificationValues.Center;
                        XmlElement errText = xmlDoc.CreateElement("Text");
                        errText.InnerText =  "参考文献标题未居中";
                        errRoot.AppendChild(errText);
                    }

                }
                else
                {
                    Justification justification = new Justification() { Val = JustificationValues.Center };
                    pPr.Append(justification);
                }
            }
           
            foreach (Run r in pRunList)
            {
                if (r != null)
                {


                    RunProperties rpr = r.GetFirstChild<RunProperties>();
                    if (rpr != null)
                    {
                        if (rpr.GetFirstChild<RunFonts>() != null)
                        {
                            if (rpr.GetFirstChild<RunFonts>().Ascii != null)
                            {
                                if (rpr.GetFirstChild<RunFonts>().Ascii != "黑体")
                                {
                                    flag2 = false;
                                }
                            }

                            rpr.GetFirstChild<RunFonts>().Ascii = "黑体";
                            rpr.GetFirstChild<RunFonts>().HighAnsi = "黑体";
                            rpr.GetFirstChild<RunFonts>().ComplexScript = "黑体";
                            rpr.GetFirstChild<RunFonts>().EastAsia = "黑体";

                        }

                        else
                        {
                            RunFonts runfont = new RunFonts() { Ascii = "黑体", HighAnsi = "黑体", ComplexScript = "黑体", EastAsia = "黑体" };
                            rpr.Append(runfont);
                        }

                        if (rpr.GetFirstChild<FontSize>() != null)
                        {
                            if (rpr.GetFirstChild<FontSize>().Val != null)
                            {
                                if (rpr.GetFirstChild<FontSize>().Val != "30")
                                {
                                    flag1 = false;
                                    rpr.GetFirstChild<FontSize>().Val = "30";
                                }
                            }
                        }

                        else
                        {
                            FontSize fontSize1 = new FontSize() { Val = "30" };
                            rpr.Append(fontSize1);
                        }
                    }
                }
            }
            if (!flag2)
            {
                XmlElement errText = xmlDoc.CreateElement("Text");
                errText.InnerText = "参考文献标题字体错误，应为黑体";
                errRoot.AppendChild(errText);
            }
            if (!flag1)
            {
                XmlElement error = xmlDoc.CreateElement("Text");
                error.InnerText = "参考文献标题字号错误，应为小三号";
                errRoot.AppendChild(error);
            }
        }

        /* 检测参考文献内容的Paragraph */
        private void checkReferenceParqagraph(Paragraph para, Paragraph paranext, WordprocessingDocument doc, bool isCnRef, bool twopara, RefTypes refType,XmlDocument xmlDoc,XmlNode errRoot)
        {
            bool flag1 = true;
            bool flag2 = true;
            List<Paragraph> list = new List<Paragraph>();
            string paraText = "";
            if (paranext == null)
            {
                paraText = para.InnerText;
                list.Add(para);
            }
            else
            {
                paraText = para.InnerText + paranext.InnerText;
                list.Add(para);
                list.Add(paranext);
            }

            foreach (Paragraph singlep in list)
            {
                Run run = singlep.GetFirstChild<Run>();
                IEnumerable<Run> pRunList = singlep.Elements<Run>();
                ParagraphProperties pPr = null;
                if (run != null)
                {
                    pPr = singlep.GetFirstChild<ParagraphProperties>();
                }
                if (pRunList != null)
                {
                    //段前、段后间距和行距的检测
                    if (pPr != null)
                    {
                        if (pPr.GetFirstChild<SpacingBetweenLines>() != null)
                        {

                            //行距
                            if (pPr.GetFirstChild<SpacingBetweenLines>().Line != null)
                            {
                                if (pPr.GetFirstChild<SpacingBetweenLines>().Line.Value != "300")
                                {
                                    XmlElement errText = xmlDoc.CreateElement("Text");
                                    errText.InnerText = "此条参考文献行间距错误，应为多倍行距1.25：" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                                    errRoot.AppendChild(errText);
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

                    foreach (Run r in pRunList)
                    {
                        if (r != null)
                        {
                            RunProperties rpr = r.GetFirstChild<RunProperties>();
                            if (rpr != null)
                            {
                                if (rpr.GetFirstChild<RunFonts>() != null)
                                {
                                    if (rpr.GetFirstChild<RunFonts>().Ascii != null)
                                    {
                                        if (rpr.GetFirstChild<RunFonts>().Ascii != "宋体")
                                        {
                                            flag2 = false;
                                        }
                                    }

                                    rpr.GetFirstChild<RunFonts>().Ascii = "宋体";
                                    rpr.GetFirstChild<RunFonts>().HighAnsi = "宋体";
                                    rpr.GetFirstChild<RunFonts>().ComplexScript = "宋体";
                                    rpr.GetFirstChild<RunFonts>().EastAsia = "宋体";

                                }

                                else
                                {

                                    RunFonts runfont = new RunFonts() { Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "宋体", EastAsia = "宋体" };
                                    rpr.Append(runfont);


                                }

                                if (rpr.GetFirstChild<FontSize>() != null)
                                {
                                    if (rpr.GetFirstChild<FontSize>().Val != null)
                                    {
                                        if (rpr.GetFirstChild<FontSize>().Val != "21")
                                        {
                                            flag1 = false;
                                            rpr.GetFirstChild<FontSize>().Val = "21";
                                        }
                                    }
                                }

                                else
                                {
                                    FontSize fontSize1 = new FontSize() { Val = "21" };
                                    rpr.Append(fontSize1);
                                }
                            }
                        }
                    }
                }
            }
            if (!flag2)
            {
                XmlElement error = xmlDoc.CreateElement("Text");
                error.InnerText = "此条参考文献行字体错误，应为宋体：" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                errRoot.AppendChild(error);
            }
            if (!flag1)
            {
                XmlElement error = xmlDoc.CreateElement("Text");
                error.InnerText = "此条参考文献行字号错误，应为5号：" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                errRoot.AppendChild(error);
            }



            //参考文献类型检测，酌情添加对应类型的处理函数

            switch (refType)
            {

                case RefTypes.M: checkRefType_M(para, paranext,doc,isCnRef,xmlDoc,errRoot);break;//普通图书,
                case RefTypes.J: checkRefType_J(para, paranext, doc, isCnRef, xmlDoc, errRoot); break;//期刊
                case RefTypes.C: checkRefType_C(para, paranext, doc, isCnRef, xmlDoc, errRoot); break; //论文集、会议录
                case RefTypes.D: checkRefType_D(para, paranext, doc, isCnRef, xmlDoc, errRoot); break;//学位论文
                case RefTypes.P: checkRefType_P(para, paranext, doc, isCnRef, xmlDoc, errRoot); break;//专利文献
                case RefTypes.P_OL: break;//专利文献包含链接地址, 标志为"[P/OL]"
                case RefTypes.R: break;//科技报告
                case RefTypes.N: break;//报纸
                case RefTypes.S: break;//标准
                case RefTypes.G: break;//汇编
                case RefTypes.J_OL: checkRefType_online(para, paranext, doc, isCnRef, twopara, xmlDoc, errRoot); break;//电子文献一种,标志为"[J/OL]"
                case RefTypes.EB_OL: checkRefType_online(para, paranext, doc, isCnRef, twopara, xmlDoc, errRoot); break;//电子文献一种,电子公告，标志为"[EB/OL]"
                case RefTypes.C_OL: checkRefType_online(para, paranext, doc, isCnRef, twopara, xmlDoc, errRoot); break; //电子文献一种,标志为"[C/OL]"
                case RefTypes.M_OL: checkRefType_online(para, paranext, doc, isCnRef, twopara, xmlDoc, errRoot); break;//电子文献一种,标志为"[M/OL]",
                case RefTypes.None:
                    {
                        XmlElement errText = xmlDoc.CreateElement("Text");
                        errText.InnerText =  "此条参考文献缺少类型或写法有误" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                        errRoot.AppendChild(errText);
                        addComment(doc, para, "缺少参考文献类型或写法有误");
                        break;//不是参考文献的类型 
                    }
            }

            //return false;
        }
    

        /*
          * 普通图书类型检测方法
          * 标志为"M"
          */
        private void checkRefType_M(Paragraph para, Paragraph paranext, WordprocessingDocument doc, bool isCnRef,XmlDocument xmlDoc,XmlNode errRoot)
        {
            string paraText = "";
            //int t;
            if (paranext == null)
            {
                paraText = para.InnerText;
            }
            else
            {
                paraText = para.InnerText + paranext.InnerText;
            }

            if (isCnRef == true)
            {
                string[] TextArr = Regex.Split(paraText, @"\[\w*\]");
                string TextBefore = "";
                string TextAfter = "";
                if(para.ParagraphProperties.NumberingProperties != null)
                {
                    TextBefore=TextArr[0];
                    TextAfter=TextArr[1];
                }
                else
                {
                    TextBefore=TextArr[1];
                    TextAfter=TextArr[2];
                }

                //Console.WriteLine(TextAfter);
                

                int plength = paraText.Length;
                int Beforelength = TextBefore.Length;
                int Afterlength = TextAfter.Length;
                bool IsYear = true;

                int index = TextBefore.IndexOf('.');
                if (index == -1)
                {

                    XmlElement errText = xmlDoc.CreateElement("Text");
                    errText.InnerText =  "作者后应带标点“.”" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                    errRoot.AppendChild(errText);
                    addComment(doc, para, "作者后应带标点“.”");
                }
                else
                {
                    string[] autorname = Regex.Split(TextBefore.Substring(0, index), @",");
                    if (autorname.Length >= 3)
                    {
                        if (autorname[autorname.Length - 1].IndexOf('等') == -1 &&
                            autorname[autorname.Length - 1].IndexOf("et al") == -1)
                        {
                            XmlElement errText = xmlDoc.CreateElement("Text");
                            errText.InnerText = "超出三个作者应有等（中文）或et al(英文）" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                            errRoot.AppendChild(errText);
                            addComment(doc, para, "超出三个作者应有等（中文）或et al(英文）");
                        }
                    }

                    else if (autorname.Length <= 2)
                    {
                        if (autorname[autorname.Length - 1].IndexOf('等') != -1 ||
                            autorname[autorname.Length - 1].IndexOf("et al") != -1)
                        {
                            XmlElement errText = xmlDoc.CreateElement("Text");
                            errText.InnerText =  "不超过三个作者应全部列出" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                            errRoot.AppendChild(errText);
                            addComment(doc, para, "不超过三个作者应全部列出");
                        }
                    }
                }

             

                if (paraText[paraText.Length - 1] == '.')
                {
                    for (int i = plength - 2; i > plength - 6; i--)
                    {
                        if (paraText[i] < 48 || paraText[i] > 57)
                        {
                         
                            
                            IsYear = false;
                            break;
                        }
                    }
                }
                else if(paraText[paraText.Length-2]=='.')
                {
                    for (int i = plength - 3; i > plength - 7; i--)
                    {
                        if (paraText[i] < 48 || paraText[i] > 57)
                        {
                          
                            IsYear = false;
                            break;
                        }
                    }
                }
                else if (paraText[paraText.Length - 3] == '.')
                {
                    for (int i = plength - 4; i > plength - 8; i--)
                    {
                        if (paraText[i] < 48 || paraText[i] > 57)
                        {

                            IsYear = false;
                            break;
                        }
                    }
                }
                else
                { }
              

            


                if (!IsYear)
                {
                   
                    addComment(doc, para, "图书类参考文献结尾应以“年份.”结尾");
                }

                bool HasC = false;
                for (int i = 0; i < plength; i++)
                {
                    if (paraText[i] == ':' /*|| paraText[i] == '：'*/)
                    {
                        HasC = true;
                        break;
                    }
                }

                if (!HasC)
                {
                    XmlElement errText = xmlDoc.CreateElement("Text");
                    errText.InnerText = "图书类参考文献的文献类型与出版社之间应标注出版社所在城市，格式为“城市：出版社" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                    errRoot.AppendChild(errText);
                    addComment(doc, para, "图书类参考文献的文献类型与出版社之间应标注出版社所在城市，格式为“城市：出版社");
                }

                bool HasComma = false;
                for (int i = 0; i < plength; i++)
                {
                    if (paraText[i] == ',' /*|| paraText[i] == '，'*/)
                    {
                        HasComma = true;
                        break;
                    }
                }

                if (!HasComma)
                {
                    XmlElement errText = xmlDoc.CreateElement("Text");
                    errText.InnerText =  "图书类参考文献缺少标点，在出版社与年份之间应有“,”" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                    errRoot.AppendChild(errText);
                    addComment(doc, para, "图书类参考文献缺少标点，在出版社与年份之间应有“,”");
                }

                bool HasDot = false;
                for (int i = 0; i < Beforelength; i++)
                {
                    if (TextBefore[i] == '.' /*|| TextBefore[i] == '．'*/)
                    {
                        HasDot = true;
                        break;
                    }
                }
                if (!HasDot)
                {

                    XmlElement errText = xmlDoc.CreateElement("Text");
                    errText.InnerText =  "图书类参考文献缺少标点，在作者和专著名之间应有“.”" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                    errRoot.AppendChild(errText);
                    addComment(doc, para, "图书类参考文献缺少标点，在作者和专著名之间应有“.”");

                }

                HasDot = false;

                for (int i = 0; i < Afterlength; i++)
                {
                    if (TextAfter[i] == '.'&&i<3)
                    {
                        HasDot = true;
                        break;
                    }
                }

                if (!HasDot)
                {
                    XmlElement errText = xmlDoc.CreateElement("Text");
                    errText.InnerText ="图书类参考文献缺少标点，在类型标志和出版地之间应有“.”" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                    errRoot.AppendChild(errText);
                    addComment(doc, para, "图书类参考文献缺少标点，在类型标志和出版地之间应有“.”");
                }
            }

            if (paraText.Trim()[paraText.Trim().Length - 1] != '.')
            {
                XmlElement errText = xmlDoc.CreateElement("Text");
                errText.InnerText = "参考文献应以标点“.”结尾" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                errRoot.AppendChild(errText);
               
                string changed = "";
                if (paranext == null)
                {
                    if (paraText.Trim()[paraText.Trim().Length - 1] < 48 || paraText.Trim()[paraText.Trim().Length - 1] > 57)
                    {
                        changed = paraText.Trim().Substring(0, paraText.Trim().Length - 1) + ".";
                    }
                    else
                    {
                        changed = paraText.Trim() + ".";
                    }
                    if (para.Elements<Run>().Count() == 1)
                    {
                        para.GetFirstChild<Run>().GetFirstChild<Text>().Text = changed;
                    }
                    else
                    {
                        IEnumerable<Run> runs = para.Elements<Run>();
                        int num = 0;
                        foreach (Run rr in runs)
                        {

                            num++;
                            if (num != 1)
                            {
                                if (rr != null)
                                {
                                    if (rr.GetFirstChild<Text>() != null)
                                    {
                                        if (rr.GetFirstChild<Text>().Text != null)
                                        {
                                            rr.GetFirstChild<Text>().Text = null;
                                        }
                                    }
                                }
                            }
                        }

                        para.GetFirstChild<Run>().GetFirstChild<Text>().Text = changed;
                    }

                }

                else
                {
                    paraText = paranext.InnerText;
                    if (paraText != "")
                    {
                        if (paraText.Trim()[paraText.Trim().Length - 1] < 48 || paraText.Trim()[paraText.Trim().Length - 1] > 57)
                        {
                            changed = paraText.Trim().Substring(0, paraText.Trim().Length - 1) + ".";
                        }
                        else
                        {
                            changed = paraText.Trim() + ".";
                        }
                        if (para.Elements<Run>().Count() == 1)
                        {
                            para.GetFirstChild<Run>().GetFirstChild<Text>().Text = changed;
                        }
                        else
                        {
                            IEnumerable<Run> runs = para.Elements<Run>();
                            int num = 0;
                            foreach (Run rr in runs)
                            {

                                num++;
                                if (num != 1)
                                {
                                    if (rr != null)
                                    {
                                        if (rr.GetFirstChild<Text>() != null)
                                        {
                                            if (rr.GetFirstChild<Text>().Text != null)
                                            {
                                                rr.GetFirstChild<Text>().Text = null;
                                            }
                                        }
                                    }
                                }
                            }

                            para.GetFirstChild<Run>().GetFirstChild<Text>().Text = changed;
                        }
                    }

                }

            }
        }

      

        /*
         * 期刊类型检测检测方法
         * 标志为"J"
         */
        public void checkRefType_J(Paragraph para, Paragraph paranext,WordprocessingDocument doc,  bool isCnRef,XmlDocument xmlDoc,XmlNode errRoot)
        {
            string paraText = "";
            int t;
            if (paranext == null)
            {
                paraText = para.InnerText;
            }
            else
            {
                paraText = para.InnerText + paranext.InnerText;
            }
            string[] textArr = Regex.Split(paraText, "[[J]]", RegexOptions.IgnoreCase);/*(paraText, @"\[J\w*\]");*///用中括号分割参考文献条目
            string textBef = "";
            string textAft="";
            //if (para.ParagraphProperties.NumberingProperties != null)
            //{
                textBef = textArr[0];
                textAft = textArr[1];
            //}
           /* else
            {
                textBef = textArr[1];
                textAft = textArr[2];
            }*/
            int AftLength = textAft.Length;
            string txt = Tool.test();
            int index = textBef.IndexOf('.');
            if (index == -1)
            {
                XmlElement errText = xmlDoc.CreateElement("Text");
                errText.InnerText = "作者后应带标点“.”" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                errRoot.AppendChild(errText);
                addComment(doc, para, "作者后应带标点“.”");
            }
            else
            {
                string[] autorname = Regex.Split(textBef.Substring(0, index), @",");
                if (autorname.Length >= 3)
                {
                    if (autorname[autorname.Length - 1].IndexOf('等') == -1 &&
                        autorname[autorname.Length - 1].IndexOf("et al") == -1)
                    {
                        XmlElement errText = xmlDoc.CreateElement("Text");
                        errText.InnerText = "超出三个作者应有等（中文）或et al(英文）" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                        errRoot.AppendChild(errText);
                        addComment(doc, para, "超出三个作者应有等（中文）或et al(英文）");
                    }
                }

                else if (autorname.Length <= 2)
                {
                    if (autorname[autorname.Length - 1].IndexOf('等') != -1 ||
                        autorname[autorname.Length - 1].IndexOf("et al") != -1)
                    {
                        XmlElement errText = xmlDoc.CreateElement("Text");
                        errText.InnerText =  "不超过三个作者应全部列出" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                        errRoot.AppendChild(errText);
                        addComment(doc, para, "不超过三个作者应全部列出");
                    }
                }
            }
          
            for ( t = 0; t <3; t++)
            {
                if (textAft[t] == '.')
                    break;
            }

            if (t == 3)
            {
                XmlElement errText = xmlDoc.CreateElement("Text");
                errText.InnerText = "类型标号后应有“.”" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                errRoot.AppendChild(errText);
                addComment(doc, para, "类型标号后应有“.”");
            }
            int colonP = 0;
            for (int a = AftLength - 1; a > 0; a--)
            {
                if (textAft[a] == ':')
                {
                    colonP = a;
                    break;
                }
            }

            if (colonP == 0)
            {

                XmlElement errText = xmlDoc.CreateElement("Text");
                errText.InnerText = "期刊类参考文献应以“:页码范围.”结尾”：" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                errRoot.AppendChild(errText);
                addComment(doc, para, "期刊类参考文献应以“:页码范围.”结尾”：");
            }

            else
            {
                XmlElement errText = xmlDoc.CreateElement("Text");
                errText.InnerText =  "期刊类参考文献应以“-”连接页码：" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                errRoot.AppendChild(errText);
                string changed = "";

                for (int a = colonP; a < AftLength; a++)
                {
                    
                    if (textAft[a] == '~')
                    {
                        
                       
                        if(paranext!=null)
                        {
                            changed=paranext.InnerText.Substring(0,paranext.InnerText.IndexOf("~"))+"-"+paranext.InnerText.Substring(paranext.InnerText.IndexOf("~")+1);
                           
                            if (paranext.Elements<Run>().Count() == 1)
                                {
                                    paranext.GetFirstChild<Run>().GetFirstChild<Text>().Text = changed;
                                }
                                else
                                {
                                    IEnumerable<Run> runs = paranext.Elements<Run>();
                                    int num = 0;
                                    foreach (Run rr in runs)
                                    {

                                        num++;
                                        if (num != 1 )
                                        {

                                            rr.GetFirstChild<Text>().Text = null;
                                        }
                                    }

                                    paranext.GetFirstChild<Run>().GetFirstChild<Text>().Text = changed;
                                }
                            paraText = para.InnerText + paranext.InnerText;
                        }
                        else
                        {
                            changed=para.InnerText.Substring(0,para.InnerText.IndexOf("~"))+"-"+para.InnerText.Substring(para.InnerText.IndexOf("~")+1);
                         
                            if (para.Elements<Run>().Count() == 1)
                                {
                                    para.GetFirstChild<Run>().GetFirstChild<Text>().Text = changed;
                                }
                                else
                                {
                                    IEnumerable<Run> runs = para.Elements<Run>();
                                    int num = 0;
                                    foreach (Run rr in runs)
                                    {

                                        num++;
                                        if (num != 1)
                                        {
                                            if(rr.GetFirstChild<Text>()!=null)
                                            rr.GetFirstChild<Text>().Text = null;
                                        }
                                    }

                                    para.GetFirstChild<Run>().GetFirstChild<Text>().Text = changed;
                                }
                            paraText = para.InnerText;
                        }


                           

                        break;
                    }
                }


                bool Findissue = false;
                int right = 0;
                int left = 0;
                for (int a = colonP - 1; a > 0; a--)
                {
                    if (textAft[a] == ')' || textAft[a] == '）')
                    {
                        right = a;

                        for (int b = a; b > 0; b--)
                        {
                            if (textAft[b] == '(' || textAft[a] == '（')
                            {

                                left = b;
                                Findissue = true;
                                break;
                            }
                        }
                        break;
                    }
                }


                bool AllNum = true;
                if (Findissue)
                {

                    for (int a = left + 1; a < right; a++)
                    {
                        if (textAft[a] < 48 || textAft[a] > 57)
                        {
                            AllNum = false;
                            break;
                        }
                    }
                }

                if (Findissue && !AllNum)
                {
                     errText = xmlDoc.CreateElement("Text");
                    errText.InnerText =  "期刊类参考文献页码前的期号应全为数字,且其中不应有空格" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                    errRoot.AppendChild(errText);
                    
                    addComment(doc, para, "期刊类参考文献页码前的期号应全为数字,且其中不应有空格");
                }

                else if (!Findissue)
                {
                    errText = xmlDoc.CreateElement("Text");
                    errText.InnerText = "期刊类参考文献页码前应有期号，格式为“(number)”" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                    errRoot.AppendChild(errText);
                    addComment(doc, para, "期刊类参考文献页码前应有期号，格式为“(number)”");
                    
                    
                }

                else
                {

                    string issue = textAft.Substring(left + 1, right - left - 1);
                    if (Convert.ToInt32(issue) > 20)
                    {
                        errText = xmlDoc.CreateElement("Text");
                        errText.InnerText = "警告：期刊类参考文献期号一般不超过20，此条参考文献期号过大 或是未写卷号" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                        errRoot.AppendChild(errText);
                       
                        addComment(doc, para, "警告：期刊类参考文献期号一般不超过20，此条参考文献期号过大 或是未写卷号");
                        
                    }

                    int reelP = 0;
                    bool hasReel = true;
                    for (int a = left - 1; a > 0; a--)
                    {
                        if (textAft[a] == ',' /*|| textAft[a] == '，'*/)
                        {
                            reelP = a;
                            break;
                        }
                    }
                    bool realReel = true;

                    if (reelP == left - 1)
                    {
                        hasReel = false;
                    }

                    else
                    {

                        for (int a = reelP + 1; a < left; a++)
                        {
                            if (textAft[a] < 48 || textAft[a] > 57)
                            {
                                realReel = false;


                                errText = xmlDoc.CreateElement("Text");
                                errText.InnerText =  "期刊类参考文献卷号应全为数字,且其中不应有空格" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                                errRoot.AppendChild(errText);
                                    addComment(doc, para, "期刊类参考文献卷号应全为数字,且其中不应有空格");
                                    
                                    break;
                               
                            }
                        }
                    }

                    bool Findyear = true;
                    int yearP = 0;
                    for (int a = 0; a < reelP; a++)
                    {
                        if (textAft[a] >= 48 && textAft[a] <= 57)
                        {
                            yearP = a;
                            for (int b = a + 1; b < reelP; b++)
                            {
                                if (textAft[b] < 48 || textAft[b] > 57)
                                {
                                    errText = xmlDoc.CreateElement("Text");
                                    errText.InnerText = "期刊类参考文献卷号前的年份应全为数字，且其中不能有空格" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                                    errRoot.AppendChild(errText);
                                   
                                    addComment(doc, para, "期刊类参考文献卷号前的年份应全为数字，且其中不能有空格");
                                    Findyear = false;
                                    break;
                                }
                            }
                            break;
                        }
                    }

                    if (Findyear && hasReel && realReel && yearP != 0)
                    {

                        //Console.WriteLine(paraText);
                        string Writtenreel = textAft.Substring(reelP + 1, left - reelP - 1);
                        int writtenreel = Convert.ToInt32(Writtenreel);
                        string Writtenyear = textAft.Substring(yearP , 4);
                        int writtenyear = Convert.ToInt32(Writtenyear);
                        if (textAft[0] == '.')
                        {
                            string PublishHouse = "";
                            for (int i = 1; i < textAft.Length; i++)
                            {
                                if (textAft[i] == ',' || textAft[i] == '，')
                                {
                                    PublishHouse = textAft.Substring(1, i - 1);
                                    
                                    break;
                                }
                            }
                            int indexx = txt.IndexOf(PublishHouse.Trim());

                            if (indexx != -1)
                            {
                               // Console.WriteLine(PublishHouse.Trim());
                                int firstyearbeg = 0;
                                for (int a = indexx; a < txt.Length; a++)
                                {
                                    if (txt[a] >= 48 && txt[a] <= 57)
                                    {
                                        firstyearbeg = a;
                                        break;
                                    }
                                }

                                if (firstyearbeg != 0)
                                {
                                    string firstyear = txt.Substring(firstyearbeg, 4);
                                    int Firstyear = Convert.ToInt32(firstyear);

                                    if ((Firstyear + writtenreel) - writtenyear - 1 == 0)
                                    { }


                                    else if (((Firstyear + writtenreel - writtenyear - 1 < 10) && (Firstyear + writtenreel - writtenyear - 1 > 0)) ||
                                           ((Firstyear + writtenreel - writtenyear - 1 > -10) && (Firstyear + writtenreel - writtenyear - 1 < 0)))
                                    {
                                        errText = xmlDoc.CreateElement("Text");
                                        errText.InnerText =  "警告：卷号与出版社创刊年份不符" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                                        errRoot.AppendChild(errText);
                                        addComment(doc, para, "警告：卷号与出版社创刊年份不符");
                                    }

                                    else
                                    {
                                        errText = xmlDoc.CreateElement("Text");
                                        errText.InnerText =  "卷号与出版社标注的卷号相差过大" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                                        errRoot.AppendChild(errText);
                                        addComment(doc, para, "卷号与出版社标注的卷号相差过大");
                                    }
                                }
                            }
                        }
                    }
                }
            }

          
            if (paraText.Trim()[paraText.Trim().Length - 1] != '.')
            {
                XmlNode errText = xmlDoc.CreateElement("Text");
                errText.InnerText = "参考文献应以“.”结尾" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                errRoot.AppendChild(errText);

                string changed = "";
                if (paranext == null)
                {
                    if (paraText.Trim()[paraText.Trim().Length - 1] < 48 || paraText.Trim()[paraText.Trim().Length - 1] > 57)
                    {
                        changed = paraText.Trim().Substring(0, paraText.Trim().Length - 1) + ".";
                    }
                    else
                    {
                        changed = paraText.Trim() + ".";
                    }
                    if (para.Elements<Run>().Count() == 1)
                    {
                        para.GetFirstChild<Run>().GetFirstChild<Text>().Text = changed;
                    }
                    else
                    {
                        IEnumerable<Run> runs = para.Elements<Run>();
                        int num = 0;
                        foreach (Run rr in runs)
                        {

                            num++;
                            if (num != 1)
                            {
                                if (rr != null)
                                {
                                    if (rr.GetFirstChild<Text>() != null)
                                    {
                                        if (rr.GetFirstChild<Text>().Text != null)
                                        {
                                            rr.GetFirstChild<Text>().Text = null;
                                        }
                                    }
                                }
                            }
                        }

                        para.GetFirstChild<Run>().GetFirstChild<Text>().Text = changed;
                    }

                }

                else
                {
                    paraText = paranext.InnerText;
                    if (paraText != "")
                    {
                        if (paraText.Trim()[paraText.Trim().Length - 1] < 48 || paraText.Trim()[paraText.Trim().Length - 1] > 57)
                        {
                            changed = paraText.Trim().Substring(0, paraText.Trim().Length - 1) + ".";
                        }
                        else
                        {
                            changed = paraText.Trim() + ".";
                        }
                        if (para.Elements<Run>().Count() == 1)
                        {
                            para.GetFirstChild<Run>().GetFirstChild<Text>().Text = changed;
                        }
                        else
                        {
                            IEnumerable<Run> runs = para.Elements<Run>();
                            int num = 0;
                            foreach (Run rr in runs)
                            {

                                num++;
                                if (num != 1)
                                {
                                    if (rr != null)
                                    {
                                        if (rr.GetFirstChild<Text>() != null)
                                        {
                                            if (rr.GetFirstChild<Text>().Text != null)
                                            {
                                                rr.GetFirstChild<Text>().Text = null;
                                            }
                                        }
                                    }
                                }
                            }

                            para.GetFirstChild<Run>().GetFirstChild<Text>().Text = changed;
                        }

                    }
                }

            }

            //return false;
        }

        /*
         * 论文集、会议录类型检测方法
         * 标志为"C"
         */
        private void checkRefType_C(Paragraph para, Paragraph paranext, WordprocessingDocument doc,  bool isCnRef,XmlDocument xmlDoc,XmlNode errRoot)
        {
            string paraText = "";
            if (paranext == null)
            {
                paraText = para.InnerText;
            }
            else
            {
                paraText = para.InnerText + paranext.InnerText;
            }
            string[] textArr = Regex.Split(paraText, @"\[\w*\]");//用中括号分割参考文献条目
            string textBef = "";
            string textAft = "";
            if (para.ParagraphProperties.NumberingProperties != null)
            {
                textBef = textArr[0];
                textAft = textArr[1];
            }
            else
            {
                textBef = textArr[1];
                textAft = textArr[2];
            }

            int index = textBef.IndexOf('.');
            if (index == -1)
            {
                XmlNode errText = xmlDoc.CreateElement("Text");
                errText.InnerText = "作者后应带标点“.”" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                errRoot.AppendChild(errText);
                addComment(doc, para, "作者后应带标点“.”");
            }
            else
            {
                string[] autorname = Regex.Split(textBef.Substring(0, index), @",");
                if (autorname.Length >= 3)
                {
                    if (autorname[autorname.Length - 1].IndexOf('等') == -1 &&
                        autorname[autorname.Length - 1].IndexOf("et al") == -1)
                    {
                        XmlNode errText = xmlDoc.CreateElement("Text");
                        errText.InnerText = "超出三个作者应有等（中文）或et al(英文）" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                        errRoot.AppendChild(errText);
                        addComment(doc, para, "超出三个作者应有等（中文）或et al(英文）");
                    }
                }

                else if (autorname.Length <= 2)
                {
                    if (autorname[autorname.Length - 1].IndexOf('等') != -1 ||
                        autorname[autorname.Length - 1].IndexOf("et al") != -1)
                    {

                        XmlNode errText = xmlDoc.CreateElement("Text");
                        errText.InnerText = "不超过三个作者应全部列出" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                        errRoot.AppendChild(errText);
                        addComment(doc, para, "不超过三个作者应全部列出");  
                    }
                }
            }


            if (textArr.Length == 2||(textArr.Length==3&&para.ParagraphProperties.NumberingProperties==null))
            {

                int indexcomma = -1;

                if (textAft.IndexOf(':') == -1 && textAft.IndexOf('：') == -1)
                {

                    XmlNode errText = xmlDoc.CreateElement("Text");
                    errText.InnerText = "出版者或出版社前缺少标点符号“：”" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                    errRoot.AppendChild(errText);
                    addComment(doc, para, "出版者或出版社前缺少标点符号“：”");
                }
                Match match = Regex.Match(textAft, @"[1-2][0-9][0-9][0-9]");
                if (match.Success)
                {
                    int year = Convert.ToInt32(match.Value);
                    DateTime now = DateTime.Now;
                    if (year > now.Year)
                    {
                        XmlNode errText = xmlDoc.CreateElement("Text");
                        errText.InnerText = "年份超出当前年份，不合法" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                        errRoot.AppendChild(errText);
                        addComment(doc, para, "年份超出当前年份，不合法");
                    }

                    indexcomma = textAft.IndexOf(',') != -1 ? textAft.IndexOf(',') : textAft.IndexOf('，');
                    if (indexcomma != match.Index - 1 && indexcomma != match.Index - 2)
                    {
                        XmlNode errText = xmlDoc.CreateElement("Text");
                        errText.InnerText = "出版社或出版者与年份间缺少“，”" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                        errRoot.AppendChild(errText);
                        string changedd;

                        if (paranext != null)
                        {
                            Match match2 = Regex.Match(paranext.InnerText, @"[1-2][0-9][0-9][0-9]");
                            if (Regex.IsMatch(textAft[index - 1].ToString(), @"[\u4e00-\u9fbb]+$"))
                            {
                                changedd = paranext.InnerText.Substring(0, match2.Index) + "," + paranext.InnerText.Substring(match2.Index);

                            }
                            else
                            {
                                changedd = paranext.InnerText.Substring(0, match2.Index - 1) + "," + paranext.InnerText.Substring(match2.Index);
                            }

                            if (paranext.Elements<Run>().Count() == 1)
                            {
                                paranext.GetFirstChild<Run>().GetFirstChild<Text>().Text = changedd;
                            }
                            else
                            {
                                IEnumerable<Run> runs = paranext.Elements<Run>();
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

                                paranext.GetFirstChild<Run>().GetFirstChild<Text>().Text = changedd;
                            }
                            paraText = para.InnerText + paranext.InnerText;
                        }
                        else
                        {
                            Match match2 = Regex.Match(para.InnerText, @"[1-2][0-9][0-9][0-9]");
                            if (Regex.IsMatch(textAft[index - 1].ToString(), @"[\u4e00-\u9fbb]+$"))
                            {
                                changedd = para.InnerText.Substring(0, match2.Index) + "," + para.InnerText.Substring(match2.Index);

                            }
                            else
                            {
                                changedd = para.InnerText.Substring(0, match2.Index - 1) + "," + para.InnerText.Substring(match2.Index);
                            }

                            if (para.Elements<Run>().Count() == 1)
                            {
                                para.GetFirstChild<Run>().GetFirstChild<Text>().Text = changedd;
                            }
                            else
                            {
                                IEnumerable<Run> runs = para.Elements<Run>();
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

                                para.GetFirstChild<Run>().GetFirstChild<Text>().Text = changedd;
                            }
                            paraText = para.InnerText;
                        }
                    }
                }

                else
                {
                    XmlNode errText = xmlDoc.CreateElement("Text");
                    errText.InnerText = "年份超出当前年份，不合法" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                    errRoot.AppendChild(errText);
                    addComment(doc, para, "年份超出当前年份，不合法");
                }

            

            }

            else if (textArr.Length == 3 && para.ParagraphProperties.NumberingProperties != null)
            {
                string TextLast = textArr[2];
                int indexcomma = -1;

                if (textAft.IndexOf(':') == -1 && textAft.IndexOf('：') == -1)
                {
                    XmlNode errText = xmlDoc.CreateElement("Text");
                    errText.InnerText = "[出版者不详]前缺少标点符号“：”" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                    errRoot.AppendChild(errText);
                    addComment(doc, para, "[出版者不详]前缺少标点符号“：”");
                }
                Match match = Regex.Match(TextLast, @"[1-2][0-9][0-9][0-9]");
                if (match.Success)
                {
                    int year = Convert.ToInt32(match.Value);
                    DateTime now = DateTime.Now;
                    if (year > now.Year)
                    {
                        XmlNode errText = xmlDoc.CreateElement("Text");
                        errText.InnerText = "年份超出当前年份，不合法" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                        errRoot.AppendChild(errText);
                        addComment(doc, para, "年份超出当前年份，不合法");
                    }

                    indexcomma = TextLast.IndexOf(',') != -1 ? TextLast.IndexOf(',') : TextLast.IndexOf('，');
                    if (indexcomma != match.Index - 1 && indexcomma != match.Index - 2)
                    {
                        XmlNode errText = xmlDoc.CreateElement("Text");
                        errText.InnerText = "出版社或出版者与年份间缺少“，”" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                        errRoot.AppendChild(errText);
                       string changedd;

                        if (paranext != null)
                        {
                            Match match2 = Regex.Match(paranext.InnerText, @"[1-2][0-9][0-9][0-9]");
                            if (Regex.IsMatch(textAft[index - 1].ToString(), @"[\u4e00-\u9fbb]+$"))
                            {
                                changedd = paranext.InnerText.Substring(0, match2.Index) + "," + paranext.InnerText.Substring(match2.Index);

                            }
                            else
                            {
                                changedd = paranext.InnerText.Substring(0, match2.Index - 1) + "," + paranext.InnerText.Substring(match2.Index);
                            }

                            if (paranext.Elements<Run>().Count() == 1)
                            {
                                paranext.GetFirstChild<Run>().GetFirstChild<Text>().Text = changedd;
                            }
                            else
                            {
                                IEnumerable<Run> runs = paranext.Elements<Run>();
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

                                paranext.GetFirstChild<Run>().GetFirstChild<Text>().Text = changedd;
                            }
                            paraText = para.InnerText + paranext.InnerText;
                        }
                        else
                        {
                            Match match2 = Regex.Match(para.InnerText, @"[1-2][0-9][0-9][0-9]");
                            if (Regex.IsMatch(textAft[index - 1].ToString(), @"[\u4e00-\u9fbb]+$"))
                            {
                                changedd = para.InnerText.Substring(0, match2.Index) + "," + para.InnerText.Substring(match2.Index);

                            }
                            else
                            {
                                changedd = para.InnerText.Substring(0, match2.Index - 1) + "," + para.InnerText.Substring(match2.Index);
                            }

                            if (para.Elements<Run>().Count() == 1)
                            {
                                para.GetFirstChild<Run>().GetFirstChild<Text>().Text = changedd;
                            }
                            else
                            {
                                IEnumerable<Run> runs = para.Elements<Run>();
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

                                para.GetFirstChild<Run>().GetFirstChild<Text>().Text = changedd;
                            }
                            paraText = para.InnerText;
                        }
                    }
                    
                }
                else
                {

                    XmlNode errText = xmlDoc.CreateElement("Text");
                    errText.InnerText = "年份不合法" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                    errRoot.AppendChild(errText);
                    addComment(doc, para, "年份不合法");
                }


            }

            else if (textArr.Length == 4 && para.ParagraphProperties.NumberingProperties == null)
            {
                string TextLast = textArr[3];
                int indexcomma = -1;

                if (textAft.IndexOf(':') == -1 && textAft.IndexOf('：') == -1)
                {

                    XmlNode errText = xmlDoc.CreateElement("Text");
                    errText.InnerText = "[出版者不详]前缺少标点符号“：”" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                    errRoot.AppendChild(errText);
                    addComment(doc, para, "[出版者不详]前缺少标点符号“：”");
                }
                Match match = Regex.Match(TextLast, @"[1-2][0-9][0-9][0-9]");
                if (match.Success)
                {
                    int year = Convert.ToInt32(match.Value);
                    DateTime now = DateTime.Now;
                    
                    if (year > now.Year)
                    {
                        XmlNode errText = xmlDoc.CreateElement("Text");
                        errText.InnerText = "年份超出当前年份，不合法" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                        errRoot.AppendChild(errText);
                        addComment(doc, para, "年份超出当前年份，不合法");
                    }

                    indexcomma = TextLast.IndexOf(',') != -1 ? TextLast.IndexOf(',') : TextLast.IndexOf('，');
                    if (indexcomma != match.Index - 1 && indexcomma != match.Index - 2)
                    {
                        XmlNode errText = xmlDoc.CreateElement("Text");
                        errText.InnerText = "出版社或出版者与年份间缺少“，”" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                        errRoot.AppendChild(errText);
                        string changedd;

                        if (paranext != null)
                        {
                            Match match2 = Regex.Match(paranext.InnerText, @"[1-2][0-9][0-9][0-9]");
                            if (Regex.IsMatch(textAft[index - 1].ToString(), @"[\u4e00-\u9fbb]+$"))
                            {
                                changedd = paranext.InnerText.Substring(0, match2.Index) + "," + paranext.InnerText.Substring(match2.Index);

                            }
                            else
                            {
                                changedd = paranext.InnerText.Substring(0, match2.Index - 1) + "," + paranext.InnerText.Substring(match2.Index);
                            }

                            if (paranext.Elements<Run>().Count() == 1)
                            {
                                paranext.GetFirstChild<Run>().GetFirstChild<Text>().Text = changedd;
                            }
                            else
                            {
                                IEnumerable<Run> runs = paranext.Elements<Run>();
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

                                paranext.GetFirstChild<Run>().GetFirstChild<Text>().Text = changedd;
                            }
                            paraText = para.InnerText + paranext.InnerText;
                        }
                        else
                        {
                            Match match2 = Regex.Match(para.InnerText, @"[1-2][0-9][0-9][0-9]");
                            if (Regex.IsMatch(textAft[index - 1].ToString(), @"[\u4e00-\u9fbb]+$"))
                            {
                                changedd = para.InnerText.Substring(0, match2.Index) + "," + para.InnerText.Substring(match2.Index);

                            }
                            else
                            {
                                changedd = para.InnerText.Substring(0, match2.Index - 1) + "," + para.InnerText.Substring(match2.Index);
                            }

                            if (para.Elements<Run>().Count() == 1)
                            {
                                para.GetFirstChild<Run>().GetFirstChild<Text>().Text = changedd;
                            }
                            else
                            {
                                IEnumerable<Run> runs = para.Elements<Run>();
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

                                para.GetFirstChild<Run>().GetFirstChild<Text>().Text = changedd;
                            }
                            paraText = para.InnerText;
                        }
                    }

                }
                else
                {


                    XmlNode errText = xmlDoc.CreateElement("Text");
                    errText.InnerText = "年份不合法" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                    errRoot.AppendChild(errText);
                    addComment(doc, para, "年份不合法");
                }


            }

            else
            {
                
                addComment(doc, para, "多“[]”");
            }

            if (paraText.Trim()[paraText.Trim().Length - 1] != '.')
            {
                XmlNode errText = xmlDoc.CreateElement("Text");
                errText.InnerText = "参考文献应以“.”结尾" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                errRoot.AppendChild(errText);

                string changed = "";
                if (paranext == null)
                {
                    if (paraText.Trim()[paraText.Trim().Length - 1] < 48 || paraText.Trim()[paraText.Trim().Length - 1] > 57)
                    {
                        changed = paraText.Trim().Substring(0, paraText.Trim().Length - 1) + ".";
                    }
                    else
                    {
                        changed = paraText.Trim() + ".";
                    }
                    if (para.Elements<Run>().Count() == 1)
                    {
                        para.GetFirstChild<Run>().GetFirstChild<Text>().Text = changed;
                    }
                    else
                    {
                        IEnumerable<Run> runs = para.Elements<Run>();
                        int num = 0;
                        foreach (Run rr in runs)
                        {

                            num++;
                            if (num != 1)
                            {
                                if (rr != null)
                                {
                                    if (rr.GetFirstChild<Text>() != null)
                                    {
                                        if (rr.GetFirstChild<Text>().Text != null)
                                        {
                                            rr.GetFirstChild<Text>().Text = null;
                                        }
                                    }
                                }
                            }
                        }

                        para.GetFirstChild<Run>().GetFirstChild<Text>().Text = changed;
                    }

                }

                else
                {
                    paraText = paranext.InnerText;
                    if (paraText.Trim()[paraText.Trim().Length - 1] < 48 || paraText.Trim()[paraText.Trim().Length - 1] > 57)
                    {
                        changed = paraText.Trim().Substring(0, paraText.Trim().Length - 1) + ".";
                    }
                    else
                    {
                        changed = paraText.Trim() + ".";
                    }
                    if (para.Elements<Run>().Count() == 1)
                    {
                        para.GetFirstChild<Run>().GetFirstChild<Text>().Text = changed;
                    }
                    else
                    {
                        IEnumerable<Run> runs = para.Elements<Run>();
                        int num = 0;
                        foreach (Run rr in runs)
                        {

                            num++;
                            if (num != 1)
                            {
                                if (rr != null)
                                {
                                    if (rr.GetFirstChild<Text>() != null)
                                    {
                                        if (rr.GetFirstChild<Text>().Text != null)
                                        {
                                            rr.GetFirstChild<Text>().Text = null;
                                        }
                                    }
                                }
                            }
                        }

                        para.GetFirstChild<Run>().GetFirstChild<Text>().Text = changed;
                    }

                }
            }

            //return false;
        }

        /*
         * 学位论文类型检测方法
         * 标志为"D"
         */
        private void checkRefType_D(Paragraph para, Paragraph paranext, WordprocessingDocument doc, bool isCnRef, XmlDocument xmlDoc, XmlNode errRoot)
        {
            string paraText = "";
            if (paranext == null)
            {
                paraText = para.InnerText;
            }
            else
            {
                paraText = para.InnerText + paranext.InnerText;
            }
            string[] textArr = Regex.Split(paraText, @"\[\w*\]");//用中括号分割参考文献条目
            string textBef = "";
            string textAft = "";
            if (para.ParagraphProperties.NumberingProperties != null)
            {
                textBef = textArr[0];
                textAft = textArr[1];
            }
            else
            {
                textBef = textArr[1];
                textAft = textArr[2];
            }
            int index = textBef.IndexOf('.');
            if (index <= 0)
            {
                XmlNode errText = xmlDoc.CreateElement("Text");
                errText.InnerText = "作者后应带标点“.”" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                errRoot.AppendChild(errText);
                addComment(doc, para, "作者后应带标点“.”");
            }
            else
            {
                //超出三位作者加“等”
                string[] autornames = Regex.Split(textBef.Substring(0, index), @",");
                if (autornames.Length > 3)
                {
                    if (isCnRef)
                    {
                        if (textBef.Substring(0, index).IndexOf('等') == -1)
                        {

                            XmlNode errText = xmlDoc.CreateElement("Text");
                            errText.InnerText = "超出三位作者加“等”" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                            errRoot.AppendChild(errText);
                            addComment(doc, para, "超出三位作者加“等”" );
                        }
                    }
                    else
                    {
                        if (textBef.Substring(0, index).IndexOf("et al") == -1 && textBef.Substring(0, index).IndexOf("etc") == -1)
                        {

                            XmlNode errText = xmlDoc.CreateElement("Text");
                            errText.InnerText = "超出三位作者加“et al”" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                            errRoot.AppendChild(errText);
                            addComment(doc, para, "超出三位作者加et al”");
                        }
                    }
                }
                //作者不超过三个时应全部列出
                if (autornames.Length <= 2)
                {
                    if (textBef.IndexOf('等') != -1 || textBef.IndexOf("et al") != -1)
                    {

                        XmlNode errText = xmlDoc.CreateElement("Text");
                        errText.InnerText = "作者不超过三个时应全部列出" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                        errRoot.AppendChild(errText);
                        addComment(doc, para, "作者不超过三个时应全部列出");
                    }
                }
            }
            if (isCnRef)//中文
            {
                if ((textArr.Length == 3 && para.ParagraphProperties.NumberingProperties == null) || (textArr.Length == 2 && para.ParagraphProperties.NumberingProperties != null))
                {
                    int indexcolon = -1;
                    int indexcomma = -1;
                    //学院院系前缺少标点符号
                    if ((indexcolon =  textAft.IndexOf(':')) <= 2 && textArr[1].IndexOf('：') <= 2)
                    {

                        XmlNode errText = xmlDoc.CreateElement("Text");
                        errText.InnerText = "学院院系前缺少标点符号“：”" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                        errRoot.AppendChild(errText);
                        addComment(doc, para, "学院院系前缺少标点符号“：”");

                    }
                    //年份检测
                    Match match = Regex.Match(textAft, @"[1-2][0-9][0-9][0-9]");
                    if (match.Success)
                    {
                        int year = Convert.ToInt32(match.Value);
                        DateTime now = DateTime.Now;
                       
                        if (year > now.Year)
                        {

                            XmlNode errText = xmlDoc.CreateElement("Text");
                            errText.InnerText = "超出当前年份，不合法" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                            errRoot.AppendChild(errText);
                            addComment(doc, para, "超出当前年份，不合法");
                        }
                        //出版社后逗号检测
                        indexcomma = textAft.IndexOf(',') == -1 ? textAft.IndexOf('，') : textAft.IndexOf(',');
                    
                        if (indexcomma != match.Index - 1&&indexcomma != match.Index - 2)
                        {

                            XmlNode errText = xmlDoc.CreateElement("Text");
                            errText.InnerText = "出版社或出版者与年份间缺少“，”" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                            errRoot.AppendChild(errText);
                            string changedd;

                            if (paranext != null)
                            {
                                Match match2 = Regex.Match(paranext.InnerText, @"[1-2][0-9][0-9][0-9]");
                                if (Regex.IsMatch(textAft[index - 1].ToString(), @"[\u4e00-\u9fbb]+$"))
                                {
                                    changedd = paranext.InnerText.Substring(0, match2.Index) + "," + paranext.InnerText.Substring(match2.Index);

                                }
                                /*else if (textAft[index - 2] == ',' || textAft[index - 2] == '，')
                                { }*/
                                else
                                {
                                    changedd = paranext.InnerText.Substring(0, match2.Index - 1) + "," + paranext.InnerText.Substring(match2.Index);
                                }

                                if (paranext.Elements<Run>().Count() == 1)
                                {
                                    paranext.GetFirstChild<Run>().GetFirstChild<Text>().Text = changedd;
                                }
                                else
                                {
                                    IEnumerable<Run> runs = paranext.Elements<Run>();
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

                                    paranext.GetFirstChild<Run>().GetFirstChild<Text>().Text = changedd;
                                }
                                paraText = para.InnerText + paranext.InnerText;
                            }
                            else
                            {
                                Match match2 = Regex.Match(para.InnerText, @"[1-2][0-9][0-9][0-9]");
                                if (Regex.IsMatch(textAft[index - 1].ToString(), @"[\u4e00-\u9fbb]+$"))
                                {
                                    changedd = para.InnerText.Substring(0, match2.Index) + "," + para.InnerText.Substring(match2.Index);

                                }
                                else
                                {
                                    changedd = para.InnerText.Substring(0, match2.Index - 1) + "," + para.InnerText.Substring(match2.Index);
                                }

                                if (para.Elements<Run>().Count() == 1)
                                {
                                    para.GetFirstChild<Run>().GetFirstChild<Text>().Text = changedd;
                                }
                                else
                                {
                                    IEnumerable<Run> runs = para.Elements<Run>();
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

                                    para.GetFirstChild<Run>().GetFirstChild<Text>().Text = changedd;
                                }
                                paraText = para.InnerText;
                            }
                        }
                        //学院，系不能缺少
                        string school = textAft.Substring(indexcolon + 1);
                        if (school.IndexOf("学院") == -1 && school.IndexOf("系") == -1)
                        {
                            XmlNode errText = xmlDoc.CreateElement("Text");
                            errText.InnerText = "不能缺少院系" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                            errRoot.AppendChild(errText);
                            addComment(doc, para, "不能缺少院系");
                        }
                    }
                    else
                    {
                        XmlNode errText = xmlDoc.CreateElement("Text");
                        errText.InnerText = "年份不合法" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                        errRoot.AppendChild(errText);
                        
                        addComment(doc, para, "年份不合法");
                    }
                 
                }
                else
                {

                    addComment(doc, para, "多“[]”");
                }
            }
            else//英文
            {
                int indexcolon = -1;
                int indexcomma = -1;
                //出版地后冒号检测
                if ((indexcolon = textAft.IndexOf(':')) <= 2 && textAft.IndexOf('：') <= 2)
                {

                    XmlNode errText = xmlDoc.CreateElement("Text");
                    errText.InnerText = "缺少出版社（者）或者出版社前缺少标点符号“：”" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                    errRoot.AppendChild(errText);
                    addComment(doc, para, "缺少出版社（者）或者出版社前缺少标点符号“：”");
                }
                //年份检测
                Match match = Regex.Match(textAft, @"[1-2][0-9][0-9][0-9]");
                if (match.Success)
                {
                    int year = Convert.ToInt32(match.Value);
                    DateTime now = DateTime.Now;
                    if (year > now.Year)
                    {

                        XmlNode errText = xmlDoc.CreateElement("Text");
                        errText.InnerText = "超出当前年份，不合法" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                        errRoot.AppendChild(errText);
                        addComment(doc, para, "超出当前年份，不合法");

                    }
                    //出版社后逗号检测
                    indexcomma = textAft.IndexOf(',') == -1 ? textAft.IndexOf('，') : textAft.IndexOf(',');
                    if (indexcomma != match.Index - 1 && indexcomma != match.Index - 2)
                    {
                       
                        XmlNode errText = xmlDoc.CreateElement("Text");
                        errText.InnerText = "出版社或出版者与年份间缺少“，”" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                        errRoot.AppendChild(errText);
                        string changedd;

                        if (paranext != null)
                        {
                            Match match2 = Regex.Match(paranext.InnerText, @"[1-2][0-9][0-9][0-9]");
                            if ((textAft[index - 1] > 64 && textAft[index - 1] < 91) || (textAft[index - 1] > 96 && textAft[index - 1] < 123))
                            {
                                changedd = paranext.InnerText.Substring(0, match2.Index) + "," + paranext.InnerText.Substring(match2.Index);

                            }
                            else
                            {
                                changedd = paranext.InnerText.Substring(0, match2.Index - 1) + "," + paranext.InnerText.Substring(match2.Index);
                            }

                            if (paranext.Elements<Run>().Count() == 1)
                            {
                                paranext.GetFirstChild<Run>().GetFirstChild<Text>().Text = changedd;
                            }
                            else
                            {
                                IEnumerable<Run> runs = paranext.Elements<Run>();
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

                                paranext.GetFirstChild<Run>().GetFirstChild<Text>().Text = changedd;
                            }
                            paraText = para.InnerText + paranext.InnerText;
                        }
                        else
                        {
                            Match match2 = Regex.Match(para.InnerText, @"[1-2][0-9][0-9][0-9]");
                            if ((textAft[index - 1] > 64 && textAft[index - 1] < 91) || (textAft[index - 1] > 96 && textAft[index - 1] < 123))
                            {
                                changedd = para.InnerText.Substring(0, match2.Index) + "," + para.InnerText.Substring(match2.Index);

                            }
                            else
                            {
                                changedd = para.InnerText.Substring(0, match2.Index - 1) + "," + para.InnerText.Substring(match2.Index);
                            }

                            if (para.Elements<Run>().Count() == 1)
                            {
                                para.GetFirstChild<Run>().GetFirstChild<Text>().Text = changedd;
                            }
                            else
                            {
                                IEnumerable<Run> runs = para.Elements<Run>();
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

                                para.GetFirstChild<Run>().GetFirstChild<Text>().Text = changedd;
                            }
                            paraText = para.InnerText;
                        }
                    }
                  
                }
                else
                {
                    XmlNode errText = xmlDoc.CreateElement("Text");
                    errText.InnerText = "年份不合法" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                    errRoot.AppendChild(errText);
                    addComment(doc, para, "年份不合法");
                }
              
            }

            if (paraText.Trim()[paraText.Trim().Length - 1] != '.')
            {
                XmlNode errText = xmlDoc.CreateElement("Text");
                errText.InnerText = "参考文献应以“.”结尾" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                errRoot.AppendChild(errText);
                string changed = "";
                if (paranext == null)
                {
                    if (paraText.Trim()[paraText.Trim().Length - 1] < 48 || paraText.Trim()[paraText.Trim().Length - 1] > 57)
                    {
                        changed = paraText.Trim().Substring(0, paraText.Trim().Length - 1) + ".";
                    }
                    else
                    {
                        changed = paraText.Trim() + ".";
                    }
                    if (para.Elements<Run>().Count() == 1)
                    {
                        para.GetFirstChild<Run>().GetFirstChild<Text>().Text = changed;
                    }
                    else
                    {
                        IEnumerable<Run> runs = para.Elements<Run>();
                        int num = 0;
                        foreach (Run rr in runs)
                        {

                            num++;
                            if (num != 1)
                            {
                                if (rr != null)
                                {
                                    if (rr.GetFirstChild<Text>() != null)
                                    {
                                        if (rr.GetFirstChild<Text>().Text != null)
                                        {
                                            rr.GetFirstChild<Text>().Text = null;
                                        }
                                    }
                                }
                            }
                        }

                        para.GetFirstChild<Run>().GetFirstChild<Text>().Text = changed;
                    }

                }

                else
                {
                    paraText = paranext.InnerText;
                    if (paraText.Trim()[paraText.Trim().Length - 1] < 48 || paraText.Trim()[paraText.Trim().Length - 1] > 57)
                    {
                        changed = paraText.Trim().Substring(0, paraText.Trim().Length - 1) + ".";
                    }
                    else
                    {
                        changed = paraText.Trim() + ".";
                    }
                    if (para.Elements<Run>().Count() == 1)
                    {
                        para.GetFirstChild<Run>().GetFirstChild<Text>().Text = changed;
                    }
                    else
                    {
                        IEnumerable<Run> runs = para.Elements<Run>();
                        int num = 0;
                        foreach (Run rr in runs)
                        {

                            num++;
                            if (num != 1)
                            {
                                if (rr != null)
                                {
                                    if (rr.GetFirstChild<Text>() != null)
                                    {
                                        if (rr.GetFirstChild<Text>().Text != null)
                                        {
                                            rr.GetFirstChild<Text>().Text = null;
                                        }
                                    }
                                }
                            }
                        }

                        para.GetFirstChild<Run>().GetFirstChild<Text>().Text = changed;
                    }

                }
            }
            //return false;
        }

        /*
         * 专利文献类型检测方法
         * 标志为"P"
         */
        private void checkRefType_P(Paragraph para, Paragraph paranext, WordprocessingDocument doc, bool isCnRef, XmlDocument xmlDoc, XmlNode errRoot)
        {
            //return false;
        }
        private void checkRefType_online(Paragraph para, Paragraph paranext, WordprocessingDocument doc, bool isCnRef, bool twopara, XmlDocument xmlDoc, XmlNode errRoot)
        {

            string paraText = "";
            if (paranext == null)
            {
                paraText = para.InnerText;
            }
            else
            {
                paraText = para.InnerText + paranext.InnerText;
            }
            //string[] textArr = Regex.Split(paraText, @"\[\w*\]");//用中括号分割参考文献条目
           // Console.WriteLine(textArr[0]);
            //Console.WriteLine(textArr[1]);
            string[] textArr = Regex.Split(paraText, "[OL]", RegexOptions.IgnoreCase);
            string textBef = "";
            string textAft = "";
            textBef = textArr[0];
            textAft = textArr[1];
            /*if (para.ParagraphProperties.NumberingProperties != null)
            {

                string[] textArr2 = Regex.Split(paraText, @"\[.*]");
                textBef = textArr2[0];

                textAft = textArr2[1];
            }
            else
            {
                string[] textArr2 = Regex.Split(textArr[1], @"\[.*]");
                textBef = textArr2[1];
                textAft = textArr2[2];
            }*/
            int index = textBef.IndexOf('.');
            if (index <= 0)
            {
                XmlNode errText = xmlDoc.CreateElement("Text");
                errText.InnerText = "作者后应带标点“.”" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                errRoot.AppendChild(errText);
                addComment(doc, para, "作者后应带标点“.”");
            }
            else
            {
                //超出三位作者加“等”
                string[] autornames = Regex.Split(textBef.Substring(0, index), @",");
                if (autornames.Length > 3)
                {
                    if (isCnRef)
                    {
                        if (textBef.Substring(0, index).IndexOf('等') == -1)
                        {
                            XmlNode errText = xmlDoc.CreateElement("Text");
                            errText.InnerText = "超出三位作者加“等”" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                            errRoot.AppendChild(errText);
                            addComment(doc, para, "超出三位作者加“等”");
                        }
                    }
                    else
                    {
                        if (textArr[0].Substring(0, index).IndexOf("et al") == -1 && textArr[0].Substring(0, index).IndexOf("etc") == -1)
                        {
                            XmlNode errText = xmlDoc.CreateElement("Text");
                            errText.InnerText = "超出三位作者加“et al”" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                            errRoot.AppendChild(errText);
                            addComment(doc, para, "超出三位作者加“et al”");
                        }
                    }
                }
                //作者不超过三个时应全部列出
                if (autornames.Length <= 2)
                {
                    if (textArr[0].IndexOf('等') != -1)
                    {
                        XmlNode errText = xmlDoc.CreateElement("Text");
                        errText.InnerText = "作者不超过三个时应全部列出" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}";
                        errRoot.AppendChild(errText);
                        addComment(doc, para, "作者不超过三个时应全部列出");
                    }
                }
                if (!isCnRef)
                {
                    string change="";
                    bool flag = true;
                    //英文作者首字母大写
                    for (int i = 0; i < autornames.Length; i++)
                    {
                        //找到第一个字母位置
                        Match match = Regex.Match(autornames[i], @"\w");
                        if (match.Index != -1)
                        {
                            if (autornames[i][match.Index] <= 'A' || autornames[i][match.Index] >= 'Z')
                            {
                                flag = false;

                                if (i != autornames.Length - 1)
                                {
                                    change += (autornames[i][0].ToString().ToUpper() + autornames[i].Substring(1) + ",");
                                }
                                else
                                {
                                    change += (autornames[i][0].ToString().ToUpper() + autornames[i].Substring(1));
                                }


                            }
                            else
                            {
                                if (i != autornames.Length - 1)
                                {
                                    change += (autornames[i] + ",");
                                }

                                else
                                {
                                    change += autornames[i];
                                }
                            }

                            if (!flag)
                            {
                                int Authorsuml = 0;
                                for (int ii = 0; ii < autornames.Length; ii++)
                                {
                                    Authorsuml += autornames[ii].Length;
                                }
                                change += para.InnerText.Substring(Authorsuml + 2);
                                if (para.Elements<Run>().Count() == 1)
                                {
                                    para.GetFirstChild<Run>().GetFirstChild<Text>().Text = change;
                                }
                                else
                                {
                                    IEnumerable<Run> runs = para.Elements<Run>();
                                    int num = 0;
                                    foreach (Run rr in runs)
                                    {

                                        num++;
                                        if (num != 1)
                                        {
                                            if (rr != null)
                                            {
                                                if (rr.GetFirstChild<Text>() != null)
                                                {
                                                    if (rr.GetFirstChild<Text>().Text != null)
                                                    {
                                                        rr.GetFirstChild<Text>().Text = null;
                                                    }
                                                }
                                            }
                                        }
                                    }

                                    para.GetFirstChild<Run>().GetFirstChild<Text>().Text = change;
                                }
                            }
                        }
                        else
                        {
                            XmlElement xml = xmlDoc.CreateElement("Text");
                            xml.InnerText = "缺少作者”" + "{" + paraText.Substring(0, paraText.Length < 10 ? paraText.Length : 10) + "}" + autornames[i];
                            errRoot.AppendChild(xml);
                            addComment(doc, para, "缺少作者”");
                        }
                    }
                }
            }
            //return twopara;
        }
        
      
        private void isRefNumberingCorrect(Paragraph para, int count)
        {
            
            string changed="";
                if (para.InnerText.Trim().IndexOf('[' + Convert.ToString(count) + ']') ==-1)
                {
                    Match match1 = Regex.Match(para.InnerText, @"\[[0-9](/OL)?\]");
                    Match match2 = Regex.Match(para.InnerText, @"\[(1|2|3|4)[0-9](/OL)?\]");
                    if (match1.Success)
                    {

                        changed = "[" + Convert.ToString(count) + "]" + para.InnerText.Substring(match1.Index + match1.Length);
                    }
                    else if (match2.Success)
                    {
                        changed = "[" + Convert.ToString(count) + "]" + para.InnerText.Substring(match2.Index + match2.Length);
                    }
                    else
                    {
                        changed = "[" + Convert.ToString(count) + "]" + para.InnerText.Substring(para.InnerText.IndexOf("]")+1);
                    }

                    if (para.Elements<Run>().Count() == 1)
                    {
                        para.GetFirstChild<Run>().GetFirstChild<Text>().Text = changed;
                    }
                    else
                    {
                        IEnumerable<Run> runs = para.Elements<Run>();
                        int num = 0;
                        foreach (Run rr in runs)
                        {
                            num++;
                            if (num != 1)
                            {
                                if(rr.GetFirstChild<Text>()!=null)
                                rr.GetFirstChild<Text>().Text = null;
                            }
                        }

                        para.GetFirstChild<Run>().GetFirstChild<Text>().Text = changed;
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

            Match match = Regex.Match(paraText, @"\[[A-Z]*(/OL)?\]");
            if (match.Success)
            {
                string type = match.Groups[0].Value;
                string typenormal = null;
                typenormal = type.Substring(1, type.Length - 2);
                //Console.WriteLine(typenormal);
                switch (typenormal)
                {
                    case "M":return RefTypes.M; //普通图书,
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
        private void isNumberingCorrectinContents(int RefCount, WordprocessingDocument doc)
        {
            IEnumerable<Paragraph> paras = doc.MainDocumentPart.Document.Body.Elements<Paragraph>();
            List<int> list = new List<int>();
            int maxnumber = 0;
            int location = -1;
            
            bool flagover = false;
            Paragraph p=new Paragraph();
            foreach (Paragraph para in paras)
            {
                
                location++;
                string runText = null;
                IEnumerable<Run> runs = para.Elements<Run>();
                List<Run> Psrunlist = runs.ToList<Run>();
                int runn = -1;
                
                foreach (Run run in runs)
                {
                    runn++;
                    if (run.RunProperties != null)
                    {
                        if (run.RunProperties.VerticalTextAlignment != null)
                        {
                            runText = run.InnerText;
                            Match match = Regex.Match(runText, @"\[\d+\-*\d*\]");
                            if (match.Success)
                            {

                               
                                int index = match.Value.IndexOf('-');
                                if (index == -1)
                                {
                                    if (flagover)
                                    {
                                        maxnumber = Convert.ToInt16(runText.Substring(match.Index + 1, match.Length - 2)) - 1;
                                        flagover = false;
                                        
                                    }
                                    if (Convert.ToInt16(runText.Substring(match.Index + 1, match.Length - 2)) != maxnumber + 1)
                                    {
                                        maxnumber++;
                                        run.GetFirstChild<Text>().Text = "[" + maxnumber + "]";
                                    }
                                    else
                                    {
                                        maxnumber++;
                                    }
                                   
                                }
                                else
                                {
                                    //[m-n]
                                    //m
                                    if (flagover)
                                    {
                                        maxnumber = Convert.ToInt16(match.Value.Substring(1, index - 1)) - 1;
                                        flagover = false;
                                    }
                                    string number1 = match.Value.Substring(1, index - 1);
                                    if (Convert.ToInt16(number1) != maxnumber + 1)
                                    {
                                        maxnumber++;
                                        run.GetFirstChild<Text>().Text = "[" + maxnumber + runText.Substring(index);
                                    }

                                    //n
                                    maxnumber = Convert.ToInt16(match.Value.Substring(index + 1, match.Length - (index + 2)));
                                    if (maxnumber > RefCount)
                                    {
                                        flagover = true;
                                        //Console.WriteLine(match);
                                        addComment(doc, para, "此段落参考文献角标超过总参考文献数目");
                                        continue;
                                    }
                                }

                                runText = null;
                            }

                            else if (runText.IndexOf("[") != -1)
                            {
                                int num = runn + 1;
                                if (num < Psrunlist.Count)
                                {
                                    while (Psrunlist[num].RunProperties != null && Psrunlist[num].RunProperties.VerticalTextAlignment != null)
                                    {
                                        if (Psrunlist[num].RunProperties.VerticalTextAlignment.Val != null)
                                        {
                                            if (Psrunlist[num].RunProperties.VerticalTextAlignment.Val == VerticalPositionValues.Superscript)
                                            {
                                                runText += Psrunlist[num].InnerText;
                                                num++;
                                                if (num >= Psrunlist.Count)
                                                    break;
                                            }
                                        }
                                    }
                                }

                                Match match2 = Regex.Match(runText, @"\[\d+\-*\d*\]");
                                if (match2.Success)
                                {
                                    
                                    int index = match2.Value.IndexOf('-');
                                    if (index == -1)
                                    {

                                        if (flagover)
                                        {
                                            maxnumber = Convert.ToInt16(runText.Substring(match2.Index + 1, match2.Length - 2)) - 1;
                                            flagover = false;
                                          
                                        }
                                        if (Convert.ToInt16(runText.Substring(match2.Index + 1, match2.Length - 2)) != maxnumber + 1)
                                        {
                                            maxnumber++;
                                            num = runn + 1;
                                            if (num < Psrunlist.Count)
                                            {
                                                //if (Psrunlist[num].RunProperties != null)
                                                //{
                                                   // Console.WriteLine("000ppp");
                                                  //  if (Psrunlist[num].RunProperties.VerticalTextAlignment != null)
                                                   // {
                                                        while (Psrunlist[num].RunProperties != null &&Psrunlist[num].RunProperties.VerticalTextAlignment != null)
                                                        {
                                                            if (Psrunlist[num].RunProperties.VerticalTextAlignment.Val != null)
                                                            {
                                                                if (Psrunlist[num].RunProperties.VerticalTextAlignment.Val == VerticalPositionValues.Superscript)
                                                                {
                                                                    int loc = -1;
                                                                    foreach (Run r in runs)
                                                                    {
                                                                        loc++;
                                                                        if (loc == num)
                                                                        {
                                                                            r.GetFirstChild<Text>().Text = null;
                                                                            break;
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                            num++;
                                                            if (num >= Psrunlist.Count)
                                                                break;
                                                        }
                                                    //}
                                             //   }
                                            
                                            }

                                            run.GetFirstChild<Text>().Text = "[" + maxnumber + "]";
                                        }
                                        else
                                        {
                                            maxnumber++;
                                        }

                                    }
                                    else
                                    {
                                        //[m-n]
                                        //m
                                        if (flagover)
                                        {
                                            maxnumber = Convert.ToInt16(match2.Value.Substring(1, index - 1)) - 1;
                                            flagover = false;
                                        }
                                        string number1 = match2.Value.Substring(1, index - 1);
                                        if (Convert.ToInt16(number1) != maxnumber + 1)
                                        {
                                            maxnumber++;
                                            num = runn + 1;
                                            if (num < Psrunlist.Count)
                                            {
                                                while (Psrunlist[num].RunProperties != null && Psrunlist[num].RunProperties.VerticalTextAlignment != null)
                                                {
                                                    if (Psrunlist[num].RunProperties.VerticalTextAlignment.Val != null)
                                                    {
                                                        if (Psrunlist[num].RunProperties.VerticalTextAlignment.Val == VerticalPositionValues.Superscript)
                                                        {
                                                            int loc = -1;
                                                            foreach (Run r in runs)
                                                            {
                                                                loc++;
                                                                if (loc == num)
                                                                {
                                                                    r.GetFirstChild<Text>().Text = null;
                                                                    break;
                                                                }
                                                            }
                                                        }
                                                    }
                                                    num++;
                                                    if (num >= Psrunlist.Count)
                                                        break;
                                                }
                                            }

                                            run.GetFirstChild<Text>().Text = "[" + maxnumber + runText.Substring(index);
                                        }

                                        //n
                                        maxnumber = Convert.ToInt16(match2.Value.Substring(index + 1, match2.Length - (index + 2)));
                                        if (maxnumber > RefCount)
                                        {
                                            flagover = true;

                                           
                                            addComment(doc, para, "此段落参考文献角标超过总参考文献数目");
                                            continue;
                                        }
                                    }

                                    runText = null;
                                }
                                


                            }
                            else
                            {
                                continue;
                            }
                        }
                    }
                   
                }
               
                p=para;
                
            }
            if (maxnumber < RefCount - 1 )
                {
                    
                    addComment(doc, p, "正文中缺少参考文献角标，请补全");
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

       
    }
}
    

