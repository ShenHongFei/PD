using System;
using System.Collections.Generic;
using System.Xml;
using DocumentFormat.OpenXml.Packaging;
using Wp = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Wordprocessing;
using PaperFormatDetection.Tools;
using PaperFormatDetection.Frame;

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
            string chapter = "";
            List<Paragraph> pList = toList(paras);
            List<int> iList = Tool.getTitlePosition(docx);
            string last_chapter = null;
            foreach (Paragraph p in paras)
            {
                string temp_name = null;
                string tempEnglish_name = null;
                count++;
                Run r = p.GetFirstChild<Run>();
                bool samep = false;
                if (r != null)
                {
                    Drawing d = r.GetFirstChild<Drawing>();
                    
                    Picture pic = r.GetFirstChild<Picture>();
                    //EmbeddedObject obj = r.GetFirstChild<EmbeddedObject>();
                   
                    if (d != null || pic != null)
                    {
                        if (pList != null && count < pList.Count - 1)
                        {
                            bool spaceline = false;
                            int Chinesename_position = 1;
                            int Englishname_position = 2;
                            Paragraph temp = pList[count + 1];
                            List<int> listchapter = Tool.getTitlePosition(docx);
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
                            if (temp != null)
                            {
                                temp_name = getFullText(temp).Trim();
                                if (temp_name != "")
                                {
                                    //Console.WriteLine(temp_name);
                                    Run tr = temp.GetFirstChild<Run>();
                                    if (tr != null)
                                    {
                                        Text trt = tr.GetFirstChild<Text>();

                                    }
                                    ParagraphProperties pPr = null;
                                    pPr = temp.GetFirstChild<ParagraphProperties>();
                                    if (chapter != "")
                                    {
                                        if (temp_name[0] != '图' && chapter[0] != '附')
                                        {
                                            if (temp_name[0] != 'F')
                                            {
                                                XmlElement xml = xmlDocx.CreateElement("Text");
                                                xml.InnerText = "图名错误，应为“图M.N  图的内容”：{" + temp_name + "||" + chapter + "}";
                                                root.AppendChild(xml);
                                            }
                                            else if (temp_name[0] == 'F')
                                            {
                                                samep = true;
                                            }
                                        }
                                    }
                                    //图名居中
                                    if (samep == false)
                                    {
                                        if (pPr != null)
                                        {

                                            if (pPr.GetFirstChild<Justification>() != null)
                                            {
                                                if (pPr.GetFirstChild<Justification>().Val != null)
                                                {
                                                    if (pPr.GetFirstChild<Justification>().Val != "center")
                                                    {
                                                        XmlElement xml = xmlDocx.CreateElement("Text");
                                                        xml.InnerText = "图名未居中：{" + temp_name + "||" + chapter + "}";
                                                        root.AppendChild(xml);
                                                    }

                                                }
                                            }


                                        }
                                        //图名字体字号
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
                                                        xml.InnerText = "图名字体错误,应为宋体：{" + temp_name + "||" + chapter + "}";
                                                        root.AppendChild(xml);
                                                    }
                                                }
                                                /*if (rPr.GetFirstChild<FontSize>() != null)
                                                {
                                                    if (rPr.GetFirstChild<FontSize>().Val != null && rPr.GetFirstChild<FontSize>().Val != "21")
                                                    {
                                                        XmlElement xml = xmlDocx.CreateElement("Text");
                                                        xml.InnerText = "图名字号错误,应为五号字：{" + temp_name + "||" + chapter + "}";
                                                        root.AppendChild(xml);
                                                    }
                                                }
                                                if (rPr.GetFirstChild<FontSize>() == null)
                                                {

                                                    XmlElement xml = xmlDocx.CreateElement("Text");
                                                    xml.InnerText = "图名字号错误,应为五号字：{" + temp_name + "||" + chapter + "}";
                                                    root.AppendChild(xml);

                                                }*/
                                                bool CorrectFontSize = true;
                                                if (Tool.correctsize(temp, docx, "21") == false)
                                                {
                                                    CorrectFontSize = false;
                                                    XmlElement xml = xmlDocx.CreateElement("Text");
                                                    xml.InnerText = "中文图名字号错误,应为五号字：{" + temp_name + "||" + chapter + "}";
                                                    root.AppendChild(xml);
                                                }
                                            }
                                        }
                                        //图名中间空两格
                                        int i = -1;
                                        int m = -1;
                                        int name_length = temp_name.Length;
                                        foreach (char c in temp_name)
                                        {
                                            i++;
                                            if (temp_name != null && (temp_name[0] > 57 || temp_name[0] < 48))
                                            {
                                                if (c <= 57 && c >= 48 && (temp_name[i - 1] == '.' || (temp_name[i - 1] <= 57 && temp_name[i - 1] >= 48)))
                                                {
                                                    if ((i + 1) < name_length)
                                                    {
                                                        if ((temp_name[i + 1] > 57 || temp_name[i + 1] < 48) && temp_name[i + 1] != '.')
                                                        {
                                                            m = i;
                                                            break;
                                                        }
                                                        if (temp_name[i + 1] <= 57 && temp_name[i + 1] >= 48)
                                                        {
                                                            m = i + 1;
                                                            break;
                                                        }
                                                        if ((i + 2) < name_length && temp_name[i + 1] == '.')
                                                        {
                                                            XmlElement xml = xmlDocx.CreateElement("Text");
                                                            xml.InnerText = "图名冗余，不符合论文要求，应为“图M.N  图的内容”：{" + temp_name + "||" + chapter + "}";
                                                            root.AppendChild(xml);
                                                        }
                                                    }

                                                }
                                            }
                                        }

                                        if ((m + 3) <= temp_name.Length && m != -1)
                                        {
                                            if (temp_name[m + 1] == ' ' && temp_name[m + 2] == ' ')//英文文只用一个m-1,中文还用 m-2
                                            {

                                            }
                                            else
                                            {
                                                XmlElement xml = xmlDocx.CreateElement("Text");
                                                xml.InnerText = "图名与序号中间应空两格：{" + temp_name + "||" + chapter + "}";
                                                root.AppendChild(xml);
                                            }
                                        }
                                        if (chapter != "")
                                        {
                                            if ((m == -1 || (m + 3) > temp_name.Length) && chapter[0] != '附')
                                            {
                                                XmlElement xml = xmlDocx.CreateElement("Text");
                                                xml.InnerText = "图名不完整，格式应为“图M.N  图的内容”：{" + temp_name + "||" + chapter + "}";
                                                root.AppendChild(xml);
                                            }
                                        }
                                    }

                                }
                                else if (temp_name == "")
                                {
                                    /*int count_title = 0;
                                    string title_name = null;
                                    int a = 0;
                                    foreach (int iL in iList)
                                    {
                                        if (iL < count)
                                        {
                                            a++;
                                            count_title = a;
                                        }

                                    }
                                    Paragraph ptitle = pList[count_title];
                                    if (ptitle != null)
                                    {
                                        Run rtitle = ptitle.GetFirstChild<Run>();
                                        if (rtitle != null)
                                        {

                                            title_name = getFullText(ptitle);
                                        }
                                    }

                                    XmlElement xml = xmlDocx.CreateElement("图名缺失");
                                    xml.InnerText = "图的下一行应为图名，疑似缺失图名：{" + title_name + "}";
                                    root.AppendChild(xml);*/
                                    /*Paragraph temp2 = pList[count + 2];
                                    Run r2 = temp2.GetFirstChild<Run>();
                                    string temp_name2 = null;
                                    temp_name2 = getFullText(temp2);
                                    Text trt2 = r2.GetFirstChild<Text>();*/
                                    Chinesename_position = 2;
                                    Paragraph temp_find = pList[count + Chinesename_position];

                                    for (Chinesename_position = 2; getFullText(temp_find) == ""; Chinesename_position++)
                                    {
                                        temp_find = pList[count + Chinesename_position];
                                    }
                                    string temp_name2 = null;
                                    temp_name2 = getFullText(temp_find);
                                    if (temp_name2[0] == '图')
                                    {
                                        XmlElement xml = xmlDocx.CreateElement("Text");
                                        xml.InnerText = "图与图名之间不应有空行：{" + temp_name2 + "}";
                                        root.AppendChild(xml);
                                        spaceline = true;
                                    }
                                    if (temp_name2[0] != '图')
                                    {
                                        XmlElement xml = xmlDocx.CreateElement("Text");
                                        xml.InnerText = "图的下一行应为图名，疑似缺失图名：{" + chapter + "之后的第" + num_incap + "个图}";
                                        root.AppendChild(xml);
                                    }
                                }

                            }
                            Paragraph temp_english = null;
                            if (samep == false)
                            {
                                temp_english = pList[count + 2];
                            }
                            else if (samep == true)
                            {
                                temp_english = pList[count + 1];
                            }
                            if (temp_english != null)
                            {
                                tempEnglish_name = getFullText(temp_english).Trim();
                                if (tempEnglish_name != "")
                                {
                                    bool head_correct = true;
                                    if (tempEnglish_name[0] != 'F' || tempEnglish_name[1] != 'i' || tempEnglish_name[2] != 'g' || tempEnglish_name[3] != '.')
                                    {
                                        XmlElement xml = xmlDocx.CreateElement("Text");
                                        xml.InnerText = "图英文名错误，应为“Fig. M.N  Name”：{" + tempEnglish_name + "||" + chapter + "}";
                                        root.AppendChild(xml);
                                        head_correct = false;
                                    }
                                    if (tempEnglish_name[4] != ' ' && head_correct == true)
                                    {
                                        XmlElement xml = xmlDocx.CreateElement("Text");
                                        xml.InnerText = "英文图名格式错误，“Fig.”与图片编号间应有一个空格：{" + tempEnglish_name + "||" + chapter + "}";
                                        root.AppendChild(xml);
                                    }
                                    ParagraphProperties pPr = null;
                                    pPr = temp_english.GetFirstChild<ParagraphProperties>();
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
                                                    xml.InnerText = "英文图名未居中：{" + tempEnglish_name + "||" + chapter + "}";
                                                    root.AppendChild(xml);
                                                }

                                            }
                                        }


                                    }
                                    //图名字体字号
                                    /*if (pPr != null)
                                    {

                                        ParagraphMarkRunProperties rPr = null;
                                        rPr = pPr.GetFirstChild<ParagraphMarkRunProperties>();

                                        if (rPr != null)
                                        {
                                            if (rPr.GetFirstChild<RunFonts>() != null)
                                            {
                                                if (rPr.GetFirstChild<RunFonts>().Ascii != null && rPr.GetFirstChild<RunFonts>().Ascii != "Times New Roman")
                                                {
                                                    XmlElement xml = xmlDocx.CreateElement("Text");
                                                    xml.InnerText = "英文图名字体错误,应为Times New Roman：{" + tempEnglish_name + "||" + chapter + "}";
                                                    root.AppendChild(xml);
                                                }
                                            }
                                            if (rPr.GetFirstChild<FontSize>() != null)
                                            {
                                                if (rPr.GetFirstChild<FontSize>().Val != null && rPr.GetFirstChild<FontSize>().Val != "21")
                                                {
                                                    if (rPr.GetFirstChild<FontSizeComplexScript>().Val != null && rPr.GetFirstChild<FontSizeComplexScript>().Val != "24")
                                                    {
                                                        string a = rPr.GetFirstChild<FontSizeComplexScript>().Val;
                                                        Console.WriteLine(a);
                                                        XmlElement xml = xmlDocx.CreateElement("Text");
                                                        xml.InnerText = "英文图名字号错误,应为五号字：{" + tempEnglish_name + "||" + chapter + "}";
                                                        root.AppendChild(xml);
                                                    }
                                                }
                                            }
                                            if (rPr.GetFirstChild<FontSize>() == null)
                                            {

                                                XmlElement xml = xmlDocx.CreateElement("Text");
                                                xml.InnerText = "英文图名字号错误,应为五号字：{" + tempEnglish_name + "||" + chapter + "}";
                                                root.AppendChild(xml);

                                            }
                                        }
                                    }*/
                                    if (Tool.correctsize(temp_english, docx, "21") == false)
                                    {
                                        XmlElement xml = xmlDocx.CreateElement("Text");
                                        xml.InnerText = "英文图名字号错误,应为五号字：{" + tempEnglish_name + "||" + chapter + "}";
                                        root.AppendChild(xml);
                                    }
                                    //图名中间空两格
                                    int i = -1;
                                    int m = -1;
                                    bool space_correct = true;
                                    int name_length = tempEnglish_name.Length;
                                    foreach (char c in tempEnglish_name)
                                    {
                                        i++;
                                        if (tempEnglish_name != null && (tempEnglish_name[0] > 57 || tempEnglish_name[0] < 48))
                                        {
                                            if (c <= 57 && c >= 48 && (tempEnglish_name[i - 1] == '.' || (tempEnglish_name[i - 1] <= 57 && tempEnglish_name[i - 1] >= 48)))
                                            {
                                                if ((i + 1) < name_length)
                                                {
                                                    if ((tempEnglish_name[i + 1] > 57 || tempEnglish_name[i + 1] < 48) && tempEnglish_name[i + 1] != '.')
                                                    {
                                                        m = i;
                                                        break;
                                                    }
                                                    if (tempEnglish_name[i + 1] <= 57 && tempEnglish_name[i + 1] >= 48)
                                                    {
                                                        m = i + 1;
                                                        break;
                                                    }
                                                    if ((i + 2) < name_length && tempEnglish_name[i + 1] == '.')
                                                    {
                                                        /*XmlElement xml = xmlDocx.CreateElement("Text");
                                                        xml.InnerText = "图名冗余，不符合论文要求，应为“图M.N  图的内容”：{" + tempEnglish_name + "||" + chapter + "}";
                                                        root.AppendChild(xml);*/
                                                        space_correct = false;
                                                    }
                                                }

                                            }
                                        }
                                    }

                                    if ((m + 3) <= tempEnglish_name.Length && m != -1)
                                    {
                                        if (tempEnglish_name[m + 1] == ' ' && tempEnglish_name[m + 2] == ' ')//英文文只用一个m-1,中文还用 m-2
                                        {

                                        }
                                        else
                                        {
                                            XmlElement xml = xmlDocx.CreateElement("Text");
                                            xml.InnerText = "图名与序号中间应空两格：{" + tempEnglish_name + "||" + chapter + "}";
                                            root.AppendChild(xml);
                                            space_correct = false;
                                        }
                                    }
                                    if (chapter != "")
                                    {
                                        if ((m == -1 || (m + 3) > tempEnglish_name.Length) && chapter[0] != '附')
                                        {
                                            XmlElement xml = xmlDocx.CreateElement("Text");
                                            xml.InnerText = "图名不完整，格式应为“Fig. M.N  Name”：{" + tempEnglish_name + "||" + chapter + "}";
                                            root.AppendChild(xml);
                                            space_correct = false;
                                        }
                                    }
                                    if (space_correct == true && (m + 3) <= tempEnglish_name.Length)
                                    {
                                        if (tempEnglish_name[m + 3] >= 97 && tempEnglish_name[m + 3] <= 122)
                                        {
                                            XmlElement xml = xmlDocx.CreateElement("Text");
                                            xml.InnerText = "英文图名的首字母应大写”：{" + tempEnglish_name + "||" + chapter + "}";
                                            root.AppendChild(xml);
                                        }
                                    }

                                }
                                else if (tempEnglish_name == "")
                                {
                                    Englishname_position = 3;
                                    Paragraph temp_find = pList[count + Englishname_position];

                                    for (Englishname_position = 3; getFullText(temp_find) == ""; Englishname_position++)
                                    {
                                        temp_find = pList[count + Englishname_position];
                                    }
                                    string temp_name2 = null;
                                    temp_name2 = getFullText(temp_find);
                                    if (temp_name2[0] == 'F'&& spaceline == false)
                                    {
                                        XmlElement xml = xmlDocx.CreateElement("Text");
                                        xml.InnerText = "英文图名与中文图名图名之间不应有空行：{" + temp_name2 + "}";
                                        root.AppendChild(xml);
                                    }
                                    if (temp_name2[0] != 'F')
                                    {
                                        XmlElement xml = xmlDocx.CreateElement("Text");
                                        xml.InnerText = "中文图名的下一行应为英文图名，疑似缺失图名：{" + chapter + "之后的第" + num_incap + "个图}";
                                        root.AppendChild(xml);
                                    }
                                }
                                if (Chinesename_position > Englishname_position)
                                {
                                    XmlElement xml = xmlDocx.CreateElement("Text");
                                    xml.InnerText = "英文图名应在中文图名下方：{" + tempEnglish_name + "||" + chapter + "}";
                                    root.AppendChild(xml);
                                }
                            }
                        }
                        /*********************图上下空行***************************/
                        if (temp_name != "" && samep == false)
                        {
                            XmlDocument xmlDocx2 = new XmlDocument();
                            xmlDocx2.Load(xmlFullPath);
                            XmlNode root2 = xmlDocx2.SelectSingleNode("Figure/Spaceline");
                            Paragraph p_text1 = null;
                            p_text1 = pList[count - 1];
                            Run r_text1 = null;
                            Run r_text2 = null;
                            r_text1 = p_text1.GetFirstChild<Run>();
                            /*if (r_text1 != null)
                            {
                                Console.WriteLine(count + " ");
                                Text t1 = r_text1.GetFirstChild<Text>();
                                if (t1 != null)
                                {
                                    XmlElement xml = xmlDocx.CreateElement("Text");
                                    xml.InnerText = "如图不在该页的起始位置，则图上方应为空行：{" + temp_name + "||" + chapter + "}";
                                    root.AppendChild(xml);
                                }
                            }*/
                            string t1 = getFullText(p_text1);
                            if (t1 != "")
                            {
                                XmlElement xml = xmlDocx.CreateElement("Text");
                                xml.InnerText = "如图不在该页的起始位置，则图上方应为空行：{" + temp_name + "||" + chapter + "}";
                                root.AppendChild(xml);
                            }
                            if (temp_name != "" && count < pList.Count - 1)
                            {
                                Paragraph p_text2 = null;
                                p_text2 = pList[count + 3];
                                
                                r_text2 = p_text2.GetFirstChild<Run>();
                                /*if (r_text2 != null)
                                {
                                    Text t2 = r_text1.GetFirstChild<Text>();
                                    if (t2 != null)
                                    {
                                        XmlElement xml = xmlDocx.CreateElement("Text");
                                        xml.InnerText = "如图名不是该页的最后一行，则图名下一行应为空行：{" + temp_name + "||" + chapter + "}";
                                        root.AppendChild(xml);
                                    }
                                }*/
                                string t2 = getFullText(p_text2);
                                if (t2 != "")
                                {
                                    XmlElement xml = xmlDocx.CreateElement("Text");
                                    xml.InnerText = "如图名不是该页的最后一行，则图名下一行应为空行：{" + temp_name + "||" + chapter + "}";
                                    root.AppendChild(xml);
                                }

                            }
                            Paragraph p_text4 = null;
                            p_text4 = pList[count - 2];
                            Run r_text4 = null;
                            r_text4 = p_text4.GetFirstChild<Run>();
                            if (r_text4 == null && r_text1 == null)
                            {
                                XmlElement xml = xmlDocx.CreateElement("Text");
                                xml.InnerText = "图上方应只有一个空行：{" + temp_name + "||" + chapter + "}";
                                root.AppendChild(xml);
                            }
                            Paragraph p_text3 = null;
                            p_text3 = pList[count + 4];
                            Run r_text3 = null;
                            r_text3 = p_text3.GetFirstChild<Run>();
                            if (r_text3 == null && r_text2 == null)
                            {
                                XmlElement xml = xmlDocx.CreateElement("Text");
                                xml.InnerText = "图下方应只有一个空行：{" + temp_name + "||" + chapter + "}";
                                root.AppendChild(xml);
                            }
                        }
                        /*********************图上下空行***************************/
                    }


                }
                /**********************图的位置*****************************/
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
                if (d_position != null || p_position != null)
                {
                    pPr_position = pict.GetFirstChild<ParagraphProperties>();
                    //rPr = r.GetFirstChild<RunProperties>();
                    //图居中
                    if (pPr_position != null)
                    {

                        if (pPr_position.GetFirstChild<Justification>() != null)
                        {
                            if (pPr_position.GetFirstChild<Justification>().Val != null && pPr_position.GetFirstChild<Justification>().Val != "center")
                            {
                                XmlElement xml = xmlDocx.CreateElement("Text");
                                xml.InnerText = "图未居中：{" + temp_name + "||" + chapter + "}";
                                root.AppendChild(xml);
                            }
                        }
                        if (pPr_position.GetFirstChild<Justification>() == null)
                        {
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
                                    XmlElement xmlx = xmlDocx.CreateElement("Text");
                                    xmlx.InnerText = "图的宽度不得超出页边距：{" + temp_name + "||" + chapter + "}";
                                    root.AppendChild(xmlx);
                                }
                            }
                        }
                    }
                }
            }
            xmlDocx.Save(xmlFullPath);
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
        static string Chapter(List<int> titlePosition, int location, Body body)
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
