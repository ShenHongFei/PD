using System;
using System.Collections.Generic;
using System.Xml;
using DocumentFormat.OpenXml.Packaging;
using Wp = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Wordprocessing;
using PaperFormatDetection.Tools;
using PaperFormatDetection.Frame;
using System.Text.RegularExpressions;

namespace PaperFormatDetection.Format
{
    public class FigureStyle : ModuleFormat
    {
        //构造函数
        public FigureStyle(List<Module> modList, PageLocator locator, int masterType) : base(modList, locator, masterType)
        {

        }

        public override void getStyle(WordprocessingDocument doc, String fileName)
        {
            string xmlFullPath = fileName + "\\FigureStyle.xml";
            CreateXmlFile(xmlFullPath);
            getFigureXml(doc, xmlFullPath);
        }

        private static void CreateXmlFile(string xmlPath)
        {
            XmlDocument xmlDocx = new XmlDocument();
            //创建类型声明节点  
            XmlNode node = xmlDocx.CreateXmlDeclaration("1.0", "utf-8", "");
            xmlDocx.AppendChild(node);
            //创建根节点  
            XmlNode root = xmlDocx.CreateElement("FigureStyle");
            XmlElement xe1 = xmlDocx.CreateElement("spErroInfo");
            xe1.SetAttribute("name", "特殊错误信息");
            XmlElement xe2 = xmlDocx.CreateElement("partName");
            xe2.SetAttribute("name", "提示名称");
            XmlElement xe3 = xmlDocx.CreateElement("Text");
            xe3.InnerText = "-----------------图-----------------";
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

        private static void getFigureXml(WordprocessingDocument docx, String xmlFullPath)
        {
            Body body = docx.MainDocumentPart.Document.Body;
            IEnumerable<Paragraph> paras = body.Elements<Paragraph>();
            XmlDocument xmlDocx = new XmlDocument();
            xmlDocx.Load(xmlFullPath);
            XmlNode root = xmlDocx.SelectSingleNode("FigureStyle/spErroInfo");
            int count = -1;
            string chapter = "";                            //章节
            List<Paragraph> pList = toList(paras);
            List<int> iList = Tool.getTitlePosition(docx);
            string last_chapter = null;

            int s = 0;
            int[] Count_spare;
            Count_spare = new int[1000];
            int t = 0;
            int[] Count_addspare;
            Count_addspare = new int[1000];

            #region  检测模块
            foreach (Paragraph p in paras)
            {
                string tu = null;
                string temp_name = null;                    //中文图名
                int Chinesename_position = 1;                    //中文名位置
                count++;
                Run r = p.GetFirstChild<Run>();
                if (r != null)
                {
                    bool flag1 = false; //EmbeddedObject 是否是图
                    Drawing d = r.GetFirstChild<Drawing>();
                    Picture pic = r.GetFirstChild<Picture>();
                    EmbeddedObject e = r.GetFirstChild<EmbeddedObject>();
                    if (r.GetFirstChild<DocumentFormat.OpenXml.AlternateContent>() != null && r.GetFirstChild<DocumentFormat.OpenXml.AlternateContent>().GetFirstChild<DocumentFormat.OpenXml.AlternateContentChoice>() != null && r.GetFirstChild<DocumentFormat.OpenXml.AlternateContent>().GetFirstChild<DocumentFormat.OpenXml.AlternateContentChoice>().GetFirstChild<Drawing>() != null)
                    {
                        Drawing d1 = r.GetFirstChild<DocumentFormat.OpenXml.AlternateContent>().GetFirstChild<DocumentFormat.OpenXml.AlternateContentChoice>().GetFirstChild<Drawing>();
                        if (d1 != null)
                        {
                            if (d1.GetFirstChild<Wp.Inline>() != null)
                            {
                                flag1 = true;
                            }
                            else
                            {
                                Tool.addComment(docx, p, "图应选择嵌入型布局");
                            }
                        }
                    }
                    /*
                    if (d != null && d.GetFirstChild<Wp.Inline>() != null)
                    {
                        Tool.addComment(docx, p, "图应选择嵌入型布局");
                        break;
                    }
                    if (pic!= null && pic.GetFirstChild<Wp.Inline>() != null)
                    {
                        Tool.addComment(docx, p, "图应选择嵌入型布局");
                        break;
                    }*/
                    if (e != null)
                    {
                        for (int i = 1; i < 3; i++)
                        {
                            if (count + i < pList.Count - 1)
                            {
                                Paragraph temp = pList[count + i];
                                string title1 = getFullText(temp).Trim();
                                if (title1.IndexOf('图') >= 0)
                                {
                                    flag1 = true;
                                }
                            }
                        }
                    }

                    if (d != null || pic != null || flag1 == true)
                    {
                        if (p != null)
                        {
                            int runnum= 0;
                            IEnumerable<Run> runs = p.Elements<Run>();
                            foreach (Run run in runs)
                            {
                                runnum++;
                            }
                            tu = getFullText(p).Trim();
                            if (tu != ""&&runnum!=1)
                            { 
                                XmlElement xmlx = xmlDocx.CreateElement("Text");
                                xmlx.InnerText = "图名在图片所在段中,导致此处无法修改：{||" + chapter + "}";
                                root.AppendChild(xmlx);
                                Tool.addComment(docx, p, "图名在图片所在段中,导致此处无法修改");
                                continue;
                            }
                        }
                        /**********************图的位置*****************************/
                        #region
                        Paragraph pict = pList[count];
                        ParagraphProperties pPr_position = null;
                        XmlDocument xmlDocx1 = new XmlDocument();
                        xmlDocx1.Load(xmlFullPath);
                        XmlNode root1 = xmlDocx1.SelectSingleNode("Figure/FigureName/FigurePosition");
                        Run r_position = pict.GetFirstChild<Run>();
                        if (r_position == null)
                            continue;
                        Drawing d_position = r_position.GetFirstChild<Drawing>();
                        Picture p_position = r_position.GetFirstChild<Picture>();
                        EmbeddedObject e_position = r_position.GetFirstChild<EmbeddedObject>();
                        if (d_position != null || p_position != null || e_position != null)
                        {
                            pPr_position = pict.GetFirstChild<ParagraphProperties>();
                            if (pPr_position != null)
                            {
                                if (pPr_position.GetFirstChild<Justification>() != null)
                                {
                                    if (pPr_position.GetFirstChild<Justification>().Val != null && pPr_position.GetFirstChild<Justification>().Val != "center")
                                    {
                                        Tool.change_center(pict);
                                        XmlElement xml = xmlDocx.CreateElement("Text");
                                        xml.InnerText = "图未居中：{" + temp_name + "||" + chapter + "}";
                                        root.AppendChild(xml);
                                    }
                                }
                                if (pPr_position.GetFirstChild<Justification>() == null)
                                {
                                    Tool.change_center(pict);
                                    XmlElement xml = xmlDocx.CreateElement("Text");
                                    xml.InnerText = "图未居中：{" + temp_name + "||" + chapter + "}";
                                    root.AppendChild(xml);
                                }

                            }
                            if (d_position != null)
                            {
                                if (d_position.GetFirstChild<Wp.Inline>() != null)
                                {
                                    Wp.Inline wp_inline = d_position.GetFirstChild<Wp.Inline>();
                                    if (wp_inline.GetFirstChild<Wp.Extent>() != null)
                                    {
                                        string a = wp_inline.GetFirstChild<Wp.Extent>().Cx;
                                        int p_size = int.Parse(a);
                                        if (a != null && p_size > 5727941)
                                        {
                                            Tool.addComment(docx, pList[count], "图的宽度不得超出页边距");
                                            XmlElement xmlx = xmlDocx.CreateElement("Text");
                                            xmlx.InnerText = "图的宽度不得超出页边距：{" + temp_name + "||" + chapter + "}";
                                            root.AppendChild(xmlx);
                                        }
                                    }
                                }
                            }
                        }
                        //图下图
                        Paragraph pict2 = pList[count + 1];
                        Run r_position2 = pict2.GetFirstChild<Run>();
                        if (r_position2 != null)
                        {
                            Drawing d_position2 = r_position2.GetFirstChild<Drawing>();
                            Picture p_position2 = r_position2.GetFirstChild<Picture>();
                            EmbeddedObject e_position2 = r_position2.GetFirstChild<EmbeddedObject>();
                            if (d_position2 != null || p_position2 != null||e_position2!=null)
                            {
                                continue;
                            }
                        }
                        #endregion
                        /**************************图名判断******************************/
                        #region
                        if (pList != null && count < pList.Count - 1)
                        {
                            Chinesename_position = 1;                    //中文名位置
                            Paragraph temp = pList[count + 1];              //中文图名所在段
                            List<int> listchapter = Tool.getTitlePosition(docx);     //章节目录
                            chapter = Chapter(listchapter, count, body);            //图所在章
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
                            if (temp != null)
                            {
                                temp_name = getFullText(temp).Trim();           //中文名（去除开头结尾空格
                                if (temp_name == "")
                                {
                                    Paragraph temp_find = pList[count + Chinesename_position];

                                    while (getFullText(temp_find) == "")
                                    {
                                        Count_spare[s] = count + Chinesename_position;
                                        s++;
                                        Chinesename_position++;
                                        temp_find = pList[count + Chinesename_position];
                                    }
                                    string temp_name2 = null;
                                    temp_name2 = getFullText(temp_find);
                                    if (temp_name2[0] == '图')
                                    {
                                        XmlElement xml = xmlDocx.CreateElement("Text");
                                        xml.InnerText = "图与图名之间不应有空行：{" + temp_name2 + "}";
                                        root.AppendChild(xml);
                                    }
                                }
                                temp = pList[count + Chinesename_position];
                                temp_name = getFullText(temp).Trim();
                                if (temp_name != "")
                                {
                                    Run tr = temp.GetFirstChild<Run>();   //图名的run
                                    if (tr != null)
                                    {
                                        Text trt = tr.GetFirstChild<Text>();
                                    }
                                    ParagraphProperties pPr = null;
                                    pPr = temp.GetFirstChild<ParagraphProperties>();
                                    string temp_name2 = null;
                                    temp_name2 = getFullText(temp);
                                    if (temp_name2[0] != '图' && temp_name2[0] != '注')
                                    {
                                        Tool.addComment(docx, p, "疑似缺失图名");
                                        XmlElement xml = xmlDocx.CreateElement("Text");
                                        xml.InnerText = "图的下一行应为图名，疑似缺失图名：{" + chapter + "之后的第" + num_incap + "个图}";
                                        root.AppendChild(xml);
                                        continue;
                                    }
                                    //图名居中
                                    if (pPr != null)
                                    {
                                        if (pPr.GetFirstChild<Justification>() != null)
                                        {
                                            if (pPr.GetFirstChild<Justification>().Val != null)
                                            {
                                                if (pPr.GetFirstChild<Justification>().Val != "center")
                                                {
                                                    XmlElement xml = xmlDocx.CreateElement("Text");
                                                    xml.InnerText = "中文图名未居中：{" + temp_name + "||" + chapter + "}";
                                                    root.AppendChild(xml);
                                                }

                                            }
                                        }
                                        else
                                        {
                                            XmlElement xml = xmlDocx.CreateElement("Text");
                                            xml.InnerText = "中文图名未居中：{" + temp_name + "||" + chapter + "}";
                                            root.AppendChild(xml);
                                        }
                                        Tool.change_center(temp);
                                    }
                                    //图名字体
                                    if (pPr != null)
                                    {

                                        ParagraphMarkRunProperties rPr = null;
                                        rPr = pPr.GetFirstChild<ParagraphMarkRunProperties>();

                                        if (rPr != null)
                                        {
                                            if (rPr.GetFirstChild<RunFonts>() != null)
                                            {
                                                if (rPr.GetFirstChild<RunFonts>().Ascii != null && rPr.GetFirstChild<RunFonts>().Ascii != "宋体")
                                                {

                                                    XmlElement xml = xmlDocx.CreateElement("Text");
                                                    xml.InnerText = "中文图名字体错误,应为宋体：{" + temp_name + "||" + chapter + "}";
                                                    root.AppendChild(xml);
                                                }
                                            }
                                            //图名字号
                                            bool CorrectFontSize = true;
                                            if (Tool.correctsize(temp, docx, "21") == false)
                                            {
                                                CorrectFontSize = false;
                                                XmlElement xml = xmlDocx.CreateElement("Text");
                                                xml.InnerText = "中文图名字号错误,应为五号字：{" + temp_name + "||" + chapter + "}";
                                                root.AppendChild(xml);
                                            }
                                        }
                                        Tool.change_fontsize(temp, "21");
                                        Tool.change_rfonts(temp, "宋体");
                                        Tool.remmovejiachu(temp);
                                    }
                                    //图名中间空两格
                                    string num = number(temp_name);

                                    int[] a1 = new int[3];
                                    if (num != null)
                                    {
                                        a1 = numberstyle(temp_name, num.Length);
                                        //序号前无空格
                                        if (a1[0] == 0)
                                        {
                                            XmlElement xe2 = xmlDocx.CreateElement("Text");
                                            xe2.InnerText = "中文图序号与“图”之间不应有空格：" + "{" + temp_name + "||" + chapter + "}";
                                            root.AppendChild(xe2);
                                            //去除中文表名序号前空格
                                            CNdeleteSpacingBeforeNumber(temp);
                                        }
                                        if (a1[1] == 0)
                                        {
                                            XmlElement xe2 = xmlDocx.CreateElement("Text");
                                            xe2.InnerText = "中文序号与图名之间应空两格：" + "{" + temp_name + "||" + chapter + "}";
                                            root.AppendChild(xe2);
                                            //序号后空格
                                            CNSpacingAfterNumebr(temp, num);

                                        }
                                        if (a1[2] == 0)
                                        {
                                            XmlElement xe2 = xmlDocx.CreateElement("Text");
                                            xe2.InnerText = "图中文序号不是M.N形式：" + "{" + temp_name + "||" + chapter + "}";
                                            root.AppendChild(xe2);
                                            //加批注
                                            if (temp != null)
                                            {
                                                Tool.addComment(docx, temp, "中文序号不是M.N形式");
                                            }
                                        }
                                    }
                                    else
                                    {
                                        XmlElement xe2 = xmlDocx.CreateElement("Text");
                                        xe2.InnerText = "图中文序号不是M.N形式：" + "{" + temp_name + "||" + chapter + "}";
                                        root.AppendChild(xe2);
                                        Tool.addComment(docx, temp, "中文序号不是M.N形式");
                                    }

                                }

                            }
                        }
                        #endregion
                        /*********************图上下空行***************************/
                        #region
                        if (temp_name != "")
                        {
                            XmlDocument xmlDocx2 = new XmlDocument();
                            xmlDocx2.Load(xmlFullPath);
                            XmlNode root2 = xmlDocx2.SelectSingleNode("Figure/Spaceline");
                            Paragraph p_text1 = null;
                            p_text1 = pList[count - 1];
                            Paragraph p_text2 = null;
                            if (count + Chinesename_position + 1 < pList.Count - 1)
                            {
                                p_text2 = pList[count + Chinesename_position + 1];
                            }
                            Run r_text1 = null;
                            Run r_text2 = null;
                            r_text1 = p_text1.GetFirstChild<Run>();
                            if (p_text2 != null)
                                r_text2 = p_text2.GetFirstChild<Run>();
                            string t1 = getFullText(p_text1).Trim();
                            string t2 = getFullText(p_text2).Trim();

                            Run temp_r = p.GetFirstChild<Run>();
                            if (t1 != "")
                            {
                                Count_addspare[t] = count - 1;
                                t++;
                                Tool.addComment(docx, p, "图上方应为空行");
                                XmlElement xml = xmlDocx.CreateElement("Text");
                                xml.InnerText = "如图不在该页的起始位置，则图上方应为空行：{" + temp_name + "||" + chapter + "}";
                                root.AppendChild(xml);
                            }
                            if (t2 != "")
                            {
                                Count_addspare[t] = count + Chinesename_position;
                                t++;
                                Tool.addComment(docx, p, "图名下一行应为空行");
                                XmlElement xml = xmlDocx.CreateElement("Text");
                                xml.InnerText = "如图名不是该页的最后一行，则图名下一行应为空行：{" + temp_name + "||" + chapter + "}";
                                root.AppendChild(xml);
                            }
                            Paragraph p_text4 = null;
                            p_text4 = pList[count - 2];
                            string t4 = getFullText(p_text4).Trim();
                            if (p_text4.GetFirstChild<Drawing>() != null && p_text4.GetFirstChild<Picture>() != null && p_text4.GetFirstChild<EmbeddedObject>() != null)
                                if (t4 == "" && t1 == "")
                                {
                                    if (pList[count - 1].GetFirstChild<Drawing>() != null && pList[count - 1].GetFirstChild<Picture>() != null)
                                    {
                                        Count_spare[s] = count - 1;
                                        s++;
                                    }

                                    XmlElement xml = xmlDocx.CreateElement("Text");
                                    xml.InnerText = "图上方应只有一个空行：{" + temp_name + "||" + chapter + "}";
                                    root.AppendChild(xml);
                                }
                            Paragraph p_text3 = null;
                            if (count + Chinesename_position + 2 < pList.Count - 1)
                                p_text3 = pList[count + Chinesename_position + 2];
                            string t3 = getFullText(p_text3).Trim();
                            if (p_text3.GetFirstChild<Drawing>() != null && p_text3.GetFirstChild<Picture>() != null && p_text3.GetFirstChild<EmbeddedObject>() != null)
                                if (t3 == "" && t2 == "")
                                {
                                    if (count + Chinesename_position + 1 < pList.Count - 1)
                                    {
                                        if (pList[count + Chinesename_position + 1] != null)
                                            if (pList[count + Chinesename_position + 2].GetFirstChild<Drawing>() != null && pList[count + Chinesename_position + 2].GetFirstChild<Picture>() != null)
                                            {
                                                Count_spare[s] = count + Chinesename_position + 1;
                                                s++;
                                            }
                                    }
                                    XmlElement xml = xmlDocx.CreateElement("Text");
                                    xml.InnerText = "图下方应只有一个空行：{" + temp_name + "||" + chapter + "}";
                                    root.AppendChild(xml);
                                }
                        }
                        #endregion
                        /*********************图上下空行***************************/
                    }
                }
            }
            #endregion
            //删除空行
            for (int i = 0; i < s; i++)
            {
                Paragraph sp = pList[Count_spare[i]];
                if (sp != null)//&& sp.Parent.Equals("DocumentFormat.OpenXml.Wordprocessing.Body"))
                {
                    try
                    {
                        body.RemoveChild<Paragraph>(sp);
                    }
                    catch (Exception)
                    {
                        XmlElement xml = xmlDocx.CreateElement("Text");
                        xml.InnerText = "检查是否用了浮于文字表面的图：{ ||" + chapter + "}";
                        root.AppendChild(xml);
                        Tool.addComment(docx, pList[Count_spare[i] + 2], "检查是否用了浮于文字表面的图");
                    }

                }

            }
            //添加空行
            for (int i = 0; i < t; i++)
            {
                //Console.WriteLine(Count_addspare[i]);
                Paragraph sp = pList[Count_addspare[i]];
                if (sp != null)
                    sp.InsertAfterSelf<Paragraph>(Tool.Generatespaceline());
            }
            xmlDocx.Save(xmlFullPath);
        }
        //判断各种出错情况a(0) 前无空格 a(1) hou liangge 
        static int[] numberstyle(string title, int numlen)
        {
            int l = -1;
            int i = -1;
            int[] a = new int[3] { 1, 1, 1 };
            foreach (char c in title)//寻找第一个数字字母位置
            {
                i++;
                if ((c <= '9' && c >= '0') || (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z'))
                {
                    l = i;
                    break;
                }
            }
            if (l == -1)//没写序号的情况
            {
                a[2] = 0;
                //加批注
                return a;
            }
            //序号前无空格  
            if (l - 1 >= 0)
            {
                if (title[l - 1] == ' ')
                {
                    a[0] = 0;
                }
            }
            //序号后两空格
            if (l + numlen + 1 < title.Length)
            {

                if (title[l + numlen] != ' ' || title[l + numlen + 1] != ' ' || title[l + numlen + 2] == ' ')
                {
                    a[1] = 0;
                }

            }
            //m.n格式
            if (l + 2 < title.Length && l >= 0)
            {
                if (title[l + 1] == '.')//m为一位数
                {
                    for (int j = 2; j < numlen; j++)
                    {
                        if (title[l + j] < '0' || title[l + j] > '9')
                        {
                            a[2] = 0;
                        }
                    }
                }
                else if (title[l + 2] == '.')//m为两位数
                {
                    for (int j = 3; j < numlen; j++)
                    {
                        if (title[l + j] <= '0' || title[l + j] >= '9')
                        {
                            a[2] = 0;
                        }
                    }
                }
                else if (title[l + 2] != '.' && (title[l + 1] != '.'))
                {
                    a[2] = 0;
                }
            }

            return a;
        }
        //去除中文图名序号前空格
        static void CNdeleteSpacingBeforeNumber(Paragraph p)
        {
            string s = p.InnerText;
            int index = s.IndexOf('图');
            int endIndex = index;
            if (index == -1)
            {
                return;
            }
            char c = s[index + 1];
            while (c == ' ')
            {
                if (endIndex == s.Length - 1)
                    break;
                endIndex++;
                c = s[endIndex];
            }
            s = s.Substring(0, index + 1) + s.Substring(endIndex);
            //替换此段落内容
            Tool.replaceText(p, s);
        }
        //获得num
        static string number(string title)
        {
            if (title != null)
            {
                string num = null;
                int l = -1;
                int i = -1;
                int j = 0;
                int len = title == null ? -1 : title.Length;
                //获得第一个数字位置用l记录
                foreach (char c in title)
                {
                    i++;
                    if ((c <= '9' && c >= '0') || (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z'))
                    {
                        l = i;
                        break;
                    }
                }
                for (j = 0; j < 5; j++)
                {
                    if (j + l < len && j + l >= 0)
                    {
                        if ((title[j + l] >= '0' && title[j + l] <= '9') || title[j + l] == '.' || (title[j + l] >= 'A' && title[j + l] <= 'Z') || (title[j + l] >= 'a' && title[j + l] <= 'z'))
                        {
                            num += title[j + l];
                        }
                        else
                        {
                            break;
                        }
                    }
                }
                return num;
            }
            else { return null; }
        }
        //保证中文图名序号后空格
        static void CNSpacingAfterNumebr(Paragraph p, string num)
        {
            string s = p.InnerText;
            if (s.IndexOf(num) == -1)
                return;
            int index = s.IndexOf(num) + num.Length;
            //匹配数字之后的空格
            Match m = Regex.Match(s.Substring(index), @"^\s+");
            if (m.Index == -1)
            {
                s = s.Substring(0, index) + "  " + s.Substring(index);
            }
            else
            {
                if (m.Length != 2)
                {
                    s = s.Substring(0, index) + "  " + s.Substring(index + m.Length);
                }
            }
            //替换此段落内容
            Tool.replaceText(p, s);
        }
        private static List<Paragraph> toList(IEnumerable<Paragraph> paras)
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
        private static String getFullText(Paragraph p)
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
        static string Chapter(List<int> titlePosition, int location, Body body)  //找图所在章节
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
        static int Item(List<int> titlePosition, int location, Body body)
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
