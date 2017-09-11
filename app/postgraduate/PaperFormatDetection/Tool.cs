using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO;
using System.Text;

namespace PaperFormatDetection.Tools
{
    /**
     * 工具类
     * 所有方法都是静态方法，可以直接使用"类名.方法名"方式调用
     */
    public class Tool
    {
        /* 获取字符串的字符个数 */
        public static int GetHanNumFromString(string str)
        {
            int count = 0;
            Regex regex = new Regex(@"^[\u4E00-\u9FA5]{0,}$");
            for (int i = 0; i < str.Length; i++)
            {
                if (regex.IsMatch(str[i].ToString()))
                {
                    count++;
                }
            }
            return count;
        }
        /*常用期刊合集*/
        static byte[] byData = new byte[100];
        static char[] charData = new char[1000];
        static string str = null;
        public static int Read()
        {
            try
            {
                /*string str = System.Environment.CurrentDirectory;
                str += "\\常用期刊合集.txt";*/
                
                
                FileStream file = new FileStream(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "常用期刊合集.txt"), FileMode.Open);
                file.Seek(0, SeekOrigin.Begin);
                file.Read(byData, 0, 100); //byData传进来的字节数组,用以接受FileStream对象中的数据,第2个参数是字节数组中开始写入数据的位置,它通常是0,表示从数组的开端文件中向数组写数据,最后一个参数规定从文件读多少字符.
                Decoder d = Encoding.Default.GetDecoder();
                d.GetChars(byData, 0, byData.Length, charData, 0);
                //Console.WriteLine(charData);
                file.Close();
                return 1;
            }
            catch (IOException e)
            {
                Console.WriteLine(e.ToString());
                return 0;
            }

        }
        public static string test()
        {
            int a = Read();
            str = new string(charData);
            return str;
        }
        /* 获取段落完整文本 */
        public static String getFullText(Paragraph p)
        {
            var list = p.ChildElements;
            Run temp_r = new Run();
            SmartTagRun temp_sr = new SmartTagRun();
            InsertedRun temp_ir = new InsertedRun();
            Hyperlink temp_ih = new Hyperlink();
            String text = "";
            foreach (var t in list)
            {
                if (t.GetType() == temp_r.GetType())
                {
                    Text pText = t.GetFirstChild<Text>();
                    if (pText != null)
                    {
                        text += pText.Text;
                    }
                }
                else if (t.GetType() == temp_sr.GetType())
                {
                    IEnumerable<Run> rr = t.Elements<Run>();
                    if (rr != null)
                    {
                        foreach (Run tr in rr)
                        {
                            Text pText = tr.GetFirstChild<Text>();
                            if (pText != null)
                            {
                                text += pText.Text;
                            }
                        }
                    }
                }
                else if (t.GetType() == temp_ir.GetType())
                {
                    Run r = t.GetFirstChild<Run>();
                    if (r != null)
                    {
                        Text pText = r.GetFirstChild<Text>();
                        if (pText != null)
                        {
                            text += pText.Text;
                        }

                    }
                }
                else if (t.GetType() == temp_ih.GetType())
                {
                    IEnumerable<Run> rr = t.Elements<Run>();
                    if (rr != null)
                    {
                        foreach (Run tr in rr)
                        {
                            Text pText = tr.GetFirstChild<Text>();
                            if (pText != null)
                            {
                                text += pText.Text;
                            }
                        }
                    }
                }
            }
            return text;
        }

        public static String getPargraphStyleId(Paragraph p)
        {
            ParagraphProperties pPr = new ParagraphProperties();
            String id = "";
            if (p.GetFirstChild<ParagraphProperties>() != null)
            {
                pPr = p.GetFirstChild<ParagraphProperties>();
            }
            if (pPr.GetFirstChild<ParagraphStyleId>() != null)
            {
                id = pPr.GetFirstChild<ParagraphStyleId>().Val;
            }
            return id;
        }

        public static int getSpaceCount(String str)
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

        /* 获取文档标题所在位置 */
        public static List<int> getTitlePosition(WordprocessingDocument doc)
        {
            Body body = doc.MainDocumentPart.Document.Body;
            var list = body.ChildElements;
            Paragraph temp_p = new Paragraph();
            List<int> titlePosition = new List<int>();
            int count = 0;
            string text = "";
            List<String> content = new List<string>();
            content = getContent(doc);
            Regex titleOne = new Regex("[1-9]");
            foreach (var p in list)
            {
                count++;
                if (p.GetType() == temp_p.GetType())
                {
                    Paragraph pp = (Paragraph)p;
                    Run r = p.GetFirstChild<Run>();
                    if (r != null)
                    {
                        text = getFullText(pp);
                        //trim去掉字符串左右两段的空格
                        if (text.Trim().Length > 1)
                        {
                            if (isTitle(text, content))
                            {
                                if (getFullText(pp).Trim().Length >= 3)
                                {
                                    if (pp.GetFirstChild<BookmarkStart>() != null || getFullText(pp).Trim().Substring(0, 3).Split('.').Length > 1)
                                    {
                                        titlePosition.Add(count);
                                    }
                                }
                                continue;
                            }
                            if (p.GetFirstChild<BookmarkStart>() != null)
                            {
                                if (titleOne.Match(text.Trim().Substring(0, 1)).Success)
                                {
                                    titlePosition.Add(count);
                                }
                            }
                        }
                    }
                }
            }
            return titlePosition;
        }

        /* 获得章标题位置 */
        public static List<int> getchaptertitleposition(WordprocessingDocument doc)
        {
            Body body = doc.MainDocumentPart.Document.Body;
            var list = body.ChildElements;
            Paragraph temp_p = new Paragraph();
            List<int> titlePosition = new List<int>();
            int count = 0;
            string text = "";
            Tool tool = new Tool();
            List<String> content = new List<string>();
            content = getContent(doc);
            Regex titleOne = new Regex("[1-9]");
            foreach (var p in list)
            {
                count++;
                if (p.GetType() == temp_p.GetType())
                {
                    Paragraph pp = (Paragraph)p;
                    Run r = p.GetFirstChild<Run>();
                    if (r != null)
                    {
                        text = getFullText(pp);
                        //trim去掉字符串左右两端的空格
                        if (text.Trim().Length > 1)
                        {
                            if (p.GetFirstChild<BookmarkStart>() != null)
                            {
                                if (isTitle(text, content))
                                {
                                    if (getFullText(pp).Trim().Length >= 3)
                                    {
                                        if (pp.GetFirstChild<BookmarkStart>() != null || getFullText(pp).Trim().Substring(0, 3).Split('.').Length > 1)
                                        {
                                            if (text.Trim().IndexOf('.') != 3 && text.Trim().IndexOf('.') != 1)
                                            {
                                                titlePosition.Add(count);
                                            }
                                        }
                                    }
                                    continue;
                                }
                                if (titleOne.Match(text.Trim().Substring(0, 1)).Success)
                                {
                                    if (text.Trim().IndexOf('.') != 3 && text.Trim().IndexOf('.') != 1)
                                    {
                                        titlePosition.Add(count);
                                    }
                                }
                            }
                        }
                    }
                }
            }
            return titlePosition;
        }

        /* 获取目录内容 */
        public static List<String> getContent(WordprocessingDocument doc)
        {
            Body body = doc.MainDocumentPart.Document.Body;
            MainDocumentPart mainPart = doc.MainDocumentPart;
            IEnumerable<Paragraph> paras = body.Elements<Paragraph>();
            List<String> contens = new List<string>();
            bool isContents = false;
            Tool tool = new Tool();
            foreach (Paragraph p in paras)
            {
                Run r = p.GetFirstChild<Run>();
                String fullText = "";
                if (r != null)
                {
                    fullText = getFullText(p).Trim();
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

        public static String getHyperlinkFullText(Hyperlink p)
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
        public static bool isTitle(String str, List<String> content)
        {
            foreach (String s in content)
            {
                if (s.Contains(str))
                {
                    return true;
                }
            }
            return false;
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


        /* 判断是从某个章开始的第几个表格,图也可用 */
        public static int NumTblCha(List<int> chapter, List<int> table, int location)
        {
            int a = 0;
            int i;
            int index = -1;
            List<int> chaptertable = new List<int>();
            if (chapter.Count != 0)
            {
                for (i = 0; chapter[i] < location; i++)
                {
                    index = i;
                    if (i == chapter.Count - 1)
                        break;
                }
            }
            foreach (int tbl in table)
            {
                if (index != -1 && index <= chapter.Count - 2)
                {
                    if (tbl >= chapter[index] - 1 && tbl <= chapter[index + 1] - 1)
                    {
                        chaptertable.Add(tbl);
                    }
                }
                if (index != -1 && index == chapter.Count - 1)
                {
                    if (tbl >= chapter[index])
                    {
                        chaptertable.Add(tbl);
                    }
                }
            }
            for (int j = 0; j < chaptertable.LongCount(); j++)
            {
                if (chaptertable[j] == location)
                    a = j + 1;
            }
            if (a == 0)
            {
                return a;
            }
            return a;
        }


        /* 判断段落字体是否正确 */
        public static bool correctfonts(Paragraph p, WordprocessingDocument doc, string CNfonts, string ENfonts)
        {
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
                string rtext = Regex.Replace(r.InnerText, @"\s*", "");
                //调试用,记得删去
                /*if(rtext!= "资料编目模块测试结果")
                {
                    continue;
                }*/
                if (rtext.Length != 0)
                {
                    bool isChinese = true;
                    Match match = Regex.Match(rtext, @"[\u4e00-\u9fa5]");
                    isChinese = match.Success ? true : false;
                    //过滤数字，数字字体没有硬性要求
                    bool isNumber = false;
                    match = Regex.Match(rtext, @"[0-9]");
                    isNumber = match.Success ? true : false;
                    if (isNumber)
                    {
                        continue;
                    }
                    //英文字母
                    bool isEnglish = false;
                    match = Regex.Match(rtext, @"[a-z]");
                    if (match.Success)
                    {
                        isEnglish = true;
                    }
                    match = Regex.Match(rtext, "@[A-Z]");
                    if (match.Success)
                    {
                        isEnglish = true;
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
                    if (isChinese)
                    {
                        if (CNrfonts != null)
                        {
                            if (CNrfonts != CNfonts)
                            {
                                return false;
                            }
                            else
                            {
                                return true;
                            }
                        }
                        else if (CNrstylefonts != null)
                        {
                            if (CNrstylefonts != CNfonts)
                            {
                                return false;
                            }
                            else
                            {
                                return true;
                            }
                        }
                        else if (CNrBasefonts != null)
                        {
                            if (CNrBasefonts != CNfonts)
                            {
                                return false;
                            }
                            else
                            {
                                return true;
                            }
                        }
                        else if (CNpstylefonts != null)
                        {
                            if (CNpstylefonts != CNfonts)
                            {
                                return false;
                            }
                            else
                            {
                                return true;
                            }
                        }
                        else if (CNpbasestylefonts != null)
                        {
                            if (CNpbasestylefonts != CNfonts)
                            {
                                return false;
                            }
                            else
                            {
                                return true;
                            }
                        }
                        else if (CNNormalfonts != null)
                        {
                            if (CNNormalfonts != CNfonts)
                            {
                                return false;
                            }
                            else
                            {
                                return true;
                            }
                        }
                        else if (CNDefaultsfonts != null)
                        {
                            if (CNDefaultsfonts != CNfonts)
                            {
                                return false;
                            }
                            else
                            {
                                return true;
                            }
                        }

                    }
                    if (isEnglish)
                    {
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
                    }
                }
                return true;
            }
            return true;
        }

        //判断段落字体是否正确
        //ye.2016/6/10
        public static bool correctsize(Paragraph p, WordprocessingDocument doc, string size)
        {
            IEnumerable<Style> style = doc.MainDocumentPart.StyleDefinitionsPart.Styles.Elements<Style>();
            //段落style id

            //正文style
            Style Normalst = null;
            string Normalsize = null;
            foreach (Style s in style)
            {
                if (s.Type == StyleValues.Paragraph && s.Default == true)
                {
                    Normalst = s;
                    if (Normalst.StyleRunProperties != null)
                    {
                        if (Normalst.StyleRunProperties.FontSize != null)
                        {
                            if (Normalst.StyleRunProperties.FontSize.Val != null)
                            {
                                Normalsize = Normalst.StyleRunProperties.FontSize.Val;
                            }
                        }
                    }
                    break;
                }
            }
            //defaults
            string Defaultssize = null;
            if (doc.MainDocumentPart.StyleDefinitionsPart.Styles.DocDefaults != null)
            {
                if (doc.MainDocumentPart.StyleDefinitionsPart.Styles.DocDefaults.RunPropertiesDefault != null)
                {
                    if (doc.MainDocumentPart.StyleDefinitionsPart.Styles.DocDefaults.RunPropertiesDefault.RunPropertiesBaseStyle != null)
                    {
                        if (doc.MainDocumentPart.StyleDefinitionsPart.Styles.DocDefaults.RunPropertiesDefault.RunPropertiesBaseStyle.FontSize != null)
                        {
                            Defaultssize = doc.MainDocumentPart.StyleDefinitionsPart.Styles.DocDefaults.RunPropertiesDefault.RunPropertiesBaseStyle.FontSize.Val;
                        }
                    }
                }
            }
            //pstyleid
            string pstyleid = null;
            string pbasestyleid = null;
            //段落style
            Style pstyle = null;
            string pstylesize = null;
            Style pbasestyle = null;
            string pbasestylesize = null;
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
                                    if (pstyle.StyleRunProperties.FontSize != null)
                                    {
                                        if (pstyle.StyleRunProperties.FontSize.Val != null)
                                        {
                                            pstylesize = pstyle.StyleRunProperties.FontSize.Val;
                                        }
                                    }
                                }
                                break;
                            }
                        }
                    }
                }
            }
            if (pstyle != null)
            {
                if (pstyle.BasedOn != null)
                {
                    string Basestyleid = null;//Baseonstyleid
                    if (pstyle.BasedOn.Val != null)
                    {
                        Basestyleid = pstyle.BasedOn.Val;
                        foreach (Style s in style)
                        {
                            if (s.StyleId == Basestyleid)
                            {
                                pbasestyle = s;
                                if (pbasestyle.StyleRunProperties != null)
                                {
                                    if (pbasestyle.StyleRunProperties.FontSize != null)
                                    {
                                        if (pbasestyle.StyleRunProperties.FontSize.Val != null)
                                        {
                                            pbasestylesize = pbasestyle.StyleRunProperties.FontSize.Val;
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
                if (r.InnerText != null)
                {
                    //rstyleid
                    string rstyleid = null;
                    //rBaseonstyleid
                    string rBasestyleid = null;
                    //rfonts
                    string rsize = null;
                    //rBaseonfonts
                    string rBasesize = null;
                    //rstyle
                    Style rstyle = null;
                    string rstylesize = null;
                    //rbasestyle
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
                                            if (rstyle.StyleRunProperties.FontSize != null)
                                            {
                                                if (rstyle.StyleRunProperties.FontSize.Val != null)
                                                {
                                                    rstylesize = rstyle.StyleRunProperties.FontSize.Val;
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
                                            if (rBasestyle.StyleRunProperties.FontSize != null)
                                            {
                                                if (rBasestyle.StyleRunProperties.FontSize.Val != null)
                                                {
                                                    rBasesize = rBasestyle.StyleRunProperties.FontSize.Val;
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
                        if (rpr.FontSize != null)
                        {
                            if (rpr.FontSize.Val != null)
                            {
                                rsize = rpr.FontSize.Val;
                            }
                        }
                    }
                    if (rsize != null)
                    {
                        if (rsize != size)
                        {
                            return false;
                        }
                        else { return true; }
                    }
                    else if (rstylesize != null)
                    {
                        if (rstylesize != size)
                        {
                            return false;
                        }
                        else { return true; }
                    }
                    else if (rBasesize != null)
                    {
                        if (rBasesize != size)
                        {
                            return false;
                        }
                        else { return true; }
                    }
                    else if (pstylesize != null)
                    {
                        if (pstylesize != size)
                        {
                            return false;
                        }
                        else { return true; }
                    }
                    else if (pbasestylesize != null)
                    {
                        if (pbasestylesize != size)
                        {
                            return false;
                        }
                        else
                        {
                            return true;
                        }
                    }
                    else if (Normalsize != null)
                    {
                        if (Normalsize != size)
                        {
                            return false;
                        }
                        else
                        {
                            return true;
                        }
                    }
                    else if (Defaultssize != null)
                    {
                        if (Defaultssize != size)
                        {
                            return false;
                        }
                        else { return true; }
                    }

                }
                return true;
            }
            return true;
        }
        public static bool correctJustification(Paragraph p, WordprocessingDocument doc, string justification)
        {
            ParagraphProperties ppr = p.GetFirstChild<ParagraphProperties>();
            IEnumerable<Style> style = doc.MainDocumentPart.StyleDefinitionsPart.Styles.Elements<Style>();
            string pstyleid = null;
            Style pstyle = null;
            string pbasestyleid = null;
            //Style pbasestyle = null;
            if (ppr != null)
            {
                if (ppr.GetFirstChild<Justification>() != null)
                {
                    Justification tj = ppr.GetFirstChild<Justification>();
                    if (tj.Val != justification)
                    {
                        return false;
                    }
                    else
                    {
                        return true;
                    }
                }
                if (p.GetFirstChild<ParagraphProperties>().ParagraphStyleId != null)
                {
                    if (p.GetFirstChild<ParagraphProperties>().ParagraphStyleId.Val != null)
                    {
                        pstyleid = p.GetFirstChild<ParagraphProperties>().ParagraphStyleId.Val;
                        foreach (Style st in style)
                        {
                            pstyle = st;
                            if (st.StyleId == pstyleid)
                            {
                                if (st.StyleParagraphProperties != null)
                                {
                                    if (st.StyleParagraphProperties.Justification != null)
                                    {
                                        if (st.StyleParagraphProperties.Justification.Val.ToString() != justification)
                                        {
                                            return false;
                                        }
                                        else
                                        {
                                            return true;
                                        }
                                    }
                                }
                                break;
                            }
                        }
                        while (pstyle.BasedOn != null)
                        {
                            if (pstyle.BasedOn.Val != null)
                            {
                                pbasestyleid = pstyle.BasedOn.Val;
                                foreach (Style st in style)
                                {
                                    if (st.StyleId == pbasestyleid)
                                    {
                                        pstyle = st;
                                        if (st.StyleParagraphProperties != null)
                                        {
                                            if (st.StyleParagraphProperties.Justification != null)
                                            {
                                                if (st.StyleParagraphProperties.Justification.Val.ToString() != justification)
                                                { return false; }
                                                else { return true; }
                                            }
                                        }
                                        break;
                                    }
                                }
                            }
                        }
                    }
                }
            }
            return true;
        }
        public static bool correctSpacingBetweenLines_line(Paragraph p, WordprocessingDocument doc, string spacing_line)
        {
            IEnumerable<Style> style = doc.MainDocumentPart.StyleDefinitionsPart.Styles.Elements<Style>();

            if (p.ParagraphProperties != null)
            {
                if (p.ParagraphProperties.ParagraphPropertiesChange == null)
                {
                    if (p.ParagraphProperties.SpacingBetweenLines != null)
                    {
                        if (p.ParagraphProperties.SpacingBetweenLines.Line != null)
                        {
                            if (p.ParagraphProperties.SpacingBetweenLines.Line.Value != spacing_line)
                            {
                                return false;
                            }
                            else
                            {
                                return true;
                            }
                        }
                    }
                }
            }
            else
            {
                if (p.ParagraphProperties.ParagraphPropertiesChange.ParagraphPropertiesExtended != null)
                {
                    if (p.ParagraphProperties.ParagraphPropertiesChange.ParagraphPropertiesExtended.SpacingBetweenLines != null)
                    {
                        if (p.ParagraphProperties.ParagraphPropertiesChange.ParagraphPropertiesExtended.SpacingBetweenLines.Line != null)
                        {
                            if (p.ParagraphProperties.ParagraphPropertiesChange.ParagraphPropertiesExtended.SpacingBetweenLines.Line.Value != spacing_line)
                            {
                                return false;
                            }
                            else
                            {
                                return true;
                            }
                        }
                    }
                }
            }
            if (p.GetFirstChild<ParagraphProperties>().ParagraphStyleId != null)
            {
                if (p.GetFirstChild<ParagraphProperties>().ParagraphStyleId.Val != null)
                {
                    string pstyleid = p.GetFirstChild<ParagraphProperties>().ParagraphStyleId.Val;
                    foreach (Style s in style)
                    {
                        if (s.StyleId == pstyleid)
                        {
                            if (s.StyleParagraphProperties != null)
                            {
                                if (s.StyleParagraphProperties.SpacingBetweenLines != null)
                                {
                                    if (s.StyleParagraphProperties.SpacingBetweenLines.Line != null)
                                    {
                                        if (s.StyleParagraphProperties.SpacingBetweenLines.Line.Value != spacing_line)
                                        {
                                            return false;
                                        }
                                        else
                                        {
                                            return true;
                                        }
                                    }
                                }
                            }
                            if (s.BasedOn != null)
                            {
                                if (s.BasedOn.Val != null)
                                {
                                    string pbaseid = s.BasedOn.Val;
                                    foreach (Style s2 in style)
                                    {
                                        if (s2.StyleId == pbaseid)
                                        {
                                            if (s2.StyleParagraphProperties != null)
                                            {
                                                if (s2.StyleParagraphProperties.SpacingBetweenLines != null)
                                                {
                                                    if (s2.StyleParagraphProperties.SpacingBetweenLines.Line != null)
                                                    {
                                                        if (s2.StyleParagraphProperties.SpacingBetweenLines.Line.Value != spacing_line)
                                                        {
                                                            return false;
                                                        }
                                                        else
                                                        {
                                                            return true;
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            break;
                        }
                    }
                }
            }
            return true;
        }
        public static bool correctSpacingBetweenLines_Be(Paragraph p, WordprocessingDocument doc, string spacing_Before)
        {
            IEnumerable<Style> style = doc.MainDocumentPart.StyleDefinitionsPart.Styles.Elements<Style>();
            if (p.ParagraphProperties != null)
            {
                if (p.ParagraphProperties.ParagraphPropertiesChange == null)
                {
                    if (p.ParagraphProperties.SpacingBetweenLines != null)
                    {

                        if (p.ParagraphProperties.SpacingBetweenLines.Before != null)
                        {
                            if (Convert.ToInt32(p.ParagraphProperties.SpacingBetweenLines.Before.Value) < Convert.ToInt32(spacing_Before) - 10 ||
                            Convert.ToInt32(p.ParagraphProperties.SpacingBetweenLines.Before.Value) > Convert.ToInt32(spacing_Before) + 10)
                            {
                                return false;
                            }
                            else
                            {
                                return true;
                            }
                        }
                    }
                }
                else
                {
                    if (p.ParagraphProperties.ParagraphPropertiesChange.ParagraphPropertiesExtended != null)
                    {
                        if (p.ParagraphProperties.ParagraphPropertiesChange.ParagraphPropertiesExtended.SpacingBetweenLines != null)
                        {

                            if (p.ParagraphProperties.ParagraphPropertiesChange.ParagraphPropertiesExtended.SpacingBetweenLines.Before != null)
                            {
                                if (Convert.ToInt32(p.ParagraphProperties.ParagraphPropertiesChange.ParagraphPropertiesExtended.SpacingBetweenLines.Before.Value)
                                < Convert.ToInt32(spacing_Before) - 10 ||
                                Convert.ToInt32(p.ParagraphProperties.ParagraphPropertiesChange.ParagraphPropertiesExtended.SpacingBetweenLines.Before.Value)
                                > Convert.ToInt32(spacing_Before) + 10)
                                {
                                    return false;
                                }
                                else
                                {
                                    return true;
                                }
                            }
                        }
                    }
                }
                if (p.GetFirstChild<ParagraphProperties>().ParagraphStyleId != null)
                {
                    if (p.GetFirstChild<ParagraphProperties>().ParagraphStyleId.Val != null)
                    {
                        string pstyleid = p.GetFirstChild<ParagraphProperties>().ParagraphStyleId.Val;
                        foreach (Style s in style)
                        {
                            if (s.StyleId == pstyleid)
                            {
                                if (s.StyleParagraphProperties != null)
                                {
                                    if (s.StyleParagraphProperties.SpacingBetweenLines != null)
                                    {
                                        if (s.StyleParagraphProperties.SpacingBetweenLines.Before != null)
                                        {
                                            if (Convert.ToInt32(s.StyleParagraphProperties.SpacingBetweenLines.Before.Value)
                                            < Convert.ToInt32(spacing_Before) - 10 ||
                                            Convert.ToInt32(s.StyleParagraphProperties.SpacingBetweenLines.Before.Value)
                                            > Convert.ToInt32(spacing_Before) + 10)
                                            {
                                                return false;
                                            }
                                            else
                                            {
                                                return true;
                                            }
                                        }
                                    }
                                }
                                if (s.BasedOn != null)
                                {
                                    if (s.BasedOn.Val != null)
                                    {
                                        string pbaseid = s.BasedOn.Val;
                                        foreach (Style s2 in style)
                                        {
                                            if (s2.StyleId == pbaseid)
                                            {
                                                if (s2.StyleParagraphProperties != null)
                                                {
                                                    if (s2.StyleParagraphProperties.SpacingBetweenLines != null)
                                                    {
                                                        if (s2.StyleParagraphProperties.SpacingBetweenLines.Before != null)
                                                        {
                                                            if (Convert.ToInt32(s2.StyleParagraphProperties.SpacingBetweenLines.Before.Value)
                                                            < Convert.ToInt32(spacing_Before) - 10 ||
                                                            Convert.ToInt32(s2.StyleParagraphProperties.SpacingBetweenLines.Before.Value)
                                                            > Convert.ToInt32(spacing_Before) + 10)
                                                            {
                                                                return false;
                                                            }
                                                            else
                                                            {
                                                                return true;
                                                            }
                                                        }
                                                    }
                                                }
                                                break;
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
            return true;
        }
        public static bool correctSpacingBetweenLines_Af(Paragraph p, WordprocessingDocument doc, string spacing_After)
        {
            IEnumerable<Style> style = doc.MainDocumentPart.StyleDefinitionsPart.Styles.Elements<Style>();
            if (p.ParagraphProperties != null)
            {
                if (p.ParagraphProperties.ParagraphPropertiesChange == null)
                {
                    if (p.ParagraphProperties.SpacingBetweenLines != null)
                    {
                        if (p.ParagraphProperties.SpacingBetweenLines.After != null)
                        {
                            if (Convert.ToInt32(p.ParagraphProperties.SpacingBetweenLines.After.Value) <
                                Convert.ToInt32(spacing_After) - 10 ||
                                Convert.ToInt32(p.ParagraphProperties.SpacingBetweenLines.After.Value) >
                                Convert.ToInt32(spacing_After) + 10)
                            {
                                return false;
                            }
                            else
                            {
                                return true;
                            }
                        }
                    }
                }
                else
                {
                    if (p.ParagraphProperties.ParagraphPropertiesChange.ParagraphPropertiesExtended != null)
                    {
                        if (p.ParagraphProperties.ParagraphPropertiesChange.ParagraphPropertiesExtended.SpacingBetweenLines != null)
                        {
                            if (p.ParagraphProperties.ParagraphPropertiesChange.ParagraphPropertiesExtended.SpacingBetweenLines.After != null)
                            {
                                if (Convert.ToInt32(p.ParagraphProperties.ParagraphPropertiesChange.ParagraphPropertiesExtended.SpacingBetweenLines.After.Value) <
                                    Convert.ToInt32(spacing_After) - 10 ||
                                    Convert.ToInt32(p.ParagraphProperties.ParagraphPropertiesChange.ParagraphPropertiesExtended.SpacingBetweenLines.After.Value) >
                                    Convert.ToInt32(spacing_After) + 10
                                    )
                                {
                                    return false;
                                }
                                else
                                {
                                    return true;
                                }
                            }
                        }
                    }
                }
                if (p.GetFirstChild<ParagraphProperties>().ParagraphStyleId != null)
                {
                    if (p.GetFirstChild<ParagraphProperties>().ParagraphStyleId.Val != null)
                    {
                        string pstyleid = p.GetFirstChild<ParagraphProperties>().ParagraphStyleId.Val;
                        foreach (Style s in style)
                        {
                            if (s.StyleId == pstyleid)
                            {
                                if (s.StyleParagraphProperties != null)
                                {
                                    if (s.StyleParagraphProperties.SpacingBetweenLines != null)
                                    {

                                        if (s.StyleParagraphProperties.SpacingBetweenLines.After != null)
                                        {
                                            if (Convert.ToInt32(s.StyleParagraphProperties.SpacingBetweenLines.After.Value)
                                                < Convert.ToInt32(spacing_After) - 10 ||
                                                Convert.ToInt32(s.StyleParagraphProperties.SpacingBetweenLines.After.Value)
                                                > Convert.ToInt32(spacing_After) + 10
                                                    )
                                            {
                                                return false;
                                            }
                                            else
                                            {
                                                return true;
                                            }
                                        }
                                    }
                                }
                                if (s.BasedOn != null)
                                {
                                    if (s.BasedOn.Val != null)
                                    {
                                        string pbaseid = s.BasedOn.Val;
                                        foreach (Style s2 in style)
                                        {
                                            if (s2.StyleId == pbaseid)
                                            {
                                                if (s2.StyleParagraphProperties != null)
                                                {
                                                    if (s2.StyleParagraphProperties.SpacingBetweenLines != null)
                                                    {

                                                        if (s2.StyleParagraphProperties.SpacingBetweenLines.After != null)
                                                        {
                                                            if (Convert.ToInt32(s2.StyleParagraphProperties.SpacingBetweenLines.After.Value) <
                                                                Convert.ToInt32(spacing_After) - 10 ||
                                                                Convert.ToInt32(s2.StyleParagraphProperties.SpacingBetweenLines.After.Value) >
                                                                Convert.ToInt32(spacing_After) + 10
                                                                )
                                                            {
                                                                return false;
                                                            }
                                                            else
                                                            {
                                                                return true;
                                                            }
                                                        }
                                                    }
                                                }
                                                break;
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
            return true;
        }

        //获得表(或其他)所在章节号
        public static string Chapter(List<int> titlePosition, int location, Body body)
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
            DocumentFormat.OpenXml.Wordprocessing.Paragraph p = null;
            if (titlelocation >= 0)
            {
                if (titlePosition[titlelocation] - 1 >= 0)
                {
                    p = (DocumentFormat.OpenXml.Wordprocessing.Paragraph)body.ChildElements.GetItem(titlePosition[titlelocation] - 1);
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
