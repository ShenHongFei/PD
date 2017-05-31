using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using PaperFormatDetection.Tools;
using PaperFormatDetection.Frame;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;

namespace PaperFormatDetection.Format
{

    public class TableStyle : ModuleFormat
    {
        //构造函数
        public TableStyle(List<Module> modList, PageLocator locator, int masterType)
            : base(modList, locator, masterType)
        {

        }
        public override void getStyle(WordprocessingDocument doc, String fileName)
        {
            string xmlFullPath = fileName + "\\TableStyle.xml";
            CreateXmlFile(xmlFullPath);
            GetTextXml(doc, xmlFullPath);
        }


        private static void CreateXmlFile(string xmlPath)
        {
            XmlDocument xmlDoc = new XmlDocument();
            XmlNode node = xmlDoc.CreateXmlDeclaration("1.0", "utf-8", "");
            xmlDoc.AppendChild(node);
            XmlNode root = xmlDoc.CreateElement("TableStyle");
            XmlElement xe1 = xmlDoc.CreateElement("spErroInfo");
            xe1.SetAttribute("name", "特殊错误信息");
            XmlElement xe2 = xmlDoc.CreateElement("partName");
            xe2.SetAttribute("name", "提示名称");
            XmlElement xe3 = xmlDoc.CreateElement("Text");
            xe3.InnerText = "-----------------表-----------------";
            xe2.AppendChild(xe3);
            root.AppendChild(xe1);
            root.AppendChild(xe2);
            xmlDoc.AppendChild(root);
            try
            {
                xmlDoc.Save(xmlPath);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        private static void GetTextXml(WordprocessingDocument doc, string xmlFullPath)
        {
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(xmlFullPath);
            XmlNode root = xmlDoc.SelectSingleNode("TableStyle/spErroInfo");

            Body body = doc.MainDocumentPart.Document.Body;
            List<int> list = new List<int>();
            list = TableLocation(body); //获得表格位置用list保存
            IEnumerable<DocumentFormat.OpenXml.Wordprocessing.Table> tbl = body.Elements<DocumentFormat.OpenXml.Wordprocessing.Table>();
            int table = -1; //表数目
            int continued = 0;//统计续表个数
            foreach (DocumentFormat.OpenXml.Wordprocessing.Table tbls in tbl)
            {
                table++;
                int location = 0;
                if (table >= 0 && table < list.LongCount())
                {
                    location = list[table];
                }
                //获得章节号以及第几个表
                string chapter = "";
                List<int> listchapter = Tool.getTitlePosition(doc);
                int numbertable = Tool.NumTblCha(listchapter, list, location);//第几个表
                chapter = Chapter(listchapter, location, body);//获得章节号
                //表名位置
                int[] index = locationOfTitleAndBlankLine(doc, location);
                //中文表名
                string CNtitle = null;
                //英文表名
                string ENtitle = null;
                /*表前后位置*/
                #region 表前后位置
                //中文表名位置检测
                if (index[0] == -1)
                {
                    XmlElement xe1 = xmlDoc.CreateElement("Text");
                    xe1.InnerText = "某处中文表名不存在" + "{" + chapter + "之后的第" + numbertable + "个表" + "}";
                    root.AppendChild(xe1);
                    //加批注
                    if (body.ChildElements.GetItem(location - 1) != null)
                    {
                        if (body.ChildElements.GetItem(location - 1).GetType().ToString() == "DocumentFormat.OpenXml.Wordprocessing.Paragraph")
                        {
                            Paragraph commentPara = (Paragraph)body.ChildElements.GetItem(location - 1);
                            Tool.addComment(doc, commentPara, "缺少中文表名");
                        }
                    }
                }
                else
                {
                    CNtitle = ((Paragraph)body.ChildElements.GetItem(index[0])).InnerText.Trim();
                    if (index[0] != location - 1)
                    {
                        XmlElement xe1 = xmlDoc.CreateElement("Text");
                        xe1.InnerText = "中文表名位置不规范" + "{" + CNtitle + "||" + chapter + "}";
                        root.AppendChild(xe1);
                        //加批注
                        if (body.ChildElements.GetItem(location - 1) != null)
                        {
                            if (body.ChildElements.GetItem(location - 1).GetType().ToString() == "DocumentFormat.OpenXml.Wordprocessing.Paragraph")
                            {
                                Paragraph commentPara = (Paragraph)body.ChildElements.GetItem(location - 1);
                                Tool.addComment(doc, commentPara, "中文表名位置不规范或者多空行");
                            }
                        }
                    }
                }
                //英文表名位置检测
                if (index[1] != -1)
                {
                    ENtitle = ((Paragraph)body.ChildElements.GetItem(index[1])).InnerText.Trim();
                        //加批注
                        if (body.ChildElements.GetItem(location - 1) != null)
                        {
                            if (body.ChildElements.GetItem(location - 1).GetType().ToString() == "DocumentFormat.OpenXml.Wordprocessing.Paragraph")
                            {
                                Paragraph commentPara = (Paragraph)body.ChildElements.GetItem(location - 1);
                                Tool.addComment(doc, commentPara, "不应该有英文表名");
                            }
                        }
                }
                //表前空行
                if (index[2] == -1)
                {
                    if (ENtitle == null && CNtitle == null)
                    {
                        XmlElement xe1 = xmlDoc.CreateElement("Text");
                        xe1.InnerText = "某未命名表前忘记空行（若表不在页首）" + "{" + chapter + "之后的第" + numbertable + "个表" + "}";
                        root.AppendChild(xe1);
                    }
                    else
                    {
                        XmlElement xe1 = xmlDoc.CreateElement("Text");
                        if (CNtitle != null)
                        {
                            xe1.InnerText = "若表不在该页首，则表上方应有一行空行：" + "{ " + CNtitle + " || " + chapter + "}";
                        }
                        root.AppendChild(xe1);
                    }
                    //加批注
                    if (body.ChildElements.GetItem(location - 1) != null)
                    {
                        if (body.ChildElements.GetItem(location - 1).GetType().ToString() == "DocumentFormat.OpenXml.Wordprocessing.Paragraph")
                        {
                            Paragraph commentPara = (Paragraph)body.ChildElements.GetItem(location - 1);
                            Tool.addComment(doc, commentPara, "表不在该页首，则表上方应有一行空行");
                        }
                    }
                }
                else
                {
                    if (index[2] != location - 2)
                    {
                        XmlElement xe1 = xmlDoc.CreateElement("Text");
                        if (CNtitle != null)
                        {
                            xe1.InnerText = "表前空行位置不规范，应在中文标题前一行：" + "{ " + CNtitle + " || " + chapter + "}";
                        }
                        else
                        {
                            if (ENtitle != null)
                            {
                                xe1.InnerText = "表前空行位置不规范，应在中文标题前一行：" + "{ " + ENtitle + " || " + chapter + "}";
                            }
                            else
                            {
                                xe1.InnerText = "某未命名表前空行位置不规范，应在中文标题前一行：" + "{" + chapter + "之后的第" + numbertable + "个表" + "}";
                            }
                        }
                        root.AppendChild(xe1);
                        //加批注
                        if (body.ChildElements.GetItem(location - 1) != null)
                        {
                            if (body.ChildElements.GetItem(location - 1).GetType().ToString() == "DocumentFormat.OpenXml.Wordprocessing.Paragraph")
                            {
                                Paragraph commentPara = (Paragraph)body.ChildElements.GetItem(location - 1);
                                Tool.addComment(doc, commentPara, "表前空行位置不规范，应在中文标题前一行");
                            }
                        }
                    }
                }
                //表后空行
                if (index[3] == -1)
                {
                    if (ENtitle == null && CNtitle == null)
                    {
                        XmlElement xe1 = xmlDoc.CreateElement("Text");
                        xe1.InnerText = "某未命名表后忘记空行（若表不在页尾）" + "{" + chapter + "之后的第" + numbertable + "个表" + "}";
                        root.AppendChild(xe1);
                    }
                    else
                    {
                        XmlElement xe1 = xmlDoc.CreateElement("Text");
                        if (CNtitle != null)
                        {
                            xe1.InnerText = "若表不在该页首，则表下方应有一行空行：" + "{ " + CNtitle + " || " + chapter + "}";
                        }
                        else
                        {
                            xe1.InnerText = "若表不在该页首，则表下方应有一行空行：" + "{ " + ENtitle + " || " + chapter + "}";
                        }
                        root.AppendChild(xe1);
                    }
                    //加批注
                    if (body.ChildElements.GetItem(location) != null)
                    {
                        if (body.ChildElements.GetItem(location).GetType().ToString() == "DocumentFormat.OpenXml.Wordprocessing.Paragraph")
                        {
                            Paragraph commentPara = (Paragraph)body.ChildElements.GetItem(location + 1);
                            Tool.addComment(doc, commentPara, "表不在该页尾，则表上方应有一行空行");
                        }
                    }
                }
                else
                {
                    if (index[3] != location + 1)
                    {
                        XmlElement xe1 = xmlDoc.CreateElement("Text");
                        if (CNtitle != null)
                        {
                            xe1.InnerText = "表后空行位置不规范，应在表的后一行：" + "{ " + CNtitle + " || " + chapter + "}";
                        }
                        else
                        {
                            if (ENtitle != null)
                            {
                                xe1.InnerText = "表后空行位置不规范，应在表的后一行：" + "{ " + ENtitle + " || " + chapter + "}";
                            }
                            else
                            {
                                xe1.InnerText = "某未命名表后空行位置不规范，应在表的后一行：" + "{" + chapter + "之后的第" + numbertable + "个表" + "}";
                            }
                        }
                        root.AppendChild(xe1);
                        //加批注
                        if (body.ChildElements.GetItem(location) != null)
                        {
                            if (body.ChildElements.GetItem(location).GetType().ToString() == "DocumentFormat.OpenXml.Wordprocessing.Paragraph")
                            {
                                Paragraph commentPara = (Paragraph)body.ChildElements.GetItem(location + 1);
                                Tool.addComment(doc, commentPara, "表后空行位置不规范，应在表的后一行(若表不在页尾)");
                            }
                        }
                    }

                }
                #endregion

                //获取表名位置
                #region 表名位置
                string[] title = { ENtitle, CNtitle };

                //number序号m.n
                string num = number(title[1]);
                string Ennum = Ennumber(title[0]);

                List<int> listchapter2 = Tool.getchaptertitleposition(doc);//获得章标题位置
                int numbertablechapter = Tool.NumTblCha(listchapter2, list, location);//判断是从某个章开始的第几个表格
                //续表统计
                //新节开始，将continued置为0
                if (numbertablechapter == 1)
                {
                    continued = 0;
                }
                if (continuedtable(title) == true)
                {
                    continued++;
                    //   Console.WriteLine("continued"+continued);
                }
                SectionProperties sectpr = sectPr(location, body);
                string s = null;
                if (title[1] != null)
                {
                    s = "{" + title[1] + "||" + chapter + "}";
                }
                else if (title[0] != null)
                {
                    s = "{" + title[0] + "||" + chapter + "}";
                }
                else
                {
                    s = "{" + chapter + "之后的第" + numbertable + "个表" + "}";
                }
                #endregion
                //先判断是否有中文表名，若无
                #region 中文表名
                if (title[1] != null)
                {
                    //获取中文题目所在段落
                    Paragraph p = (Paragraph)body.ChildElements.GetItem(index[0]);
                    //5.21新加  *******表题目格式判断*********
                    string id = "";
                    id = TtitleStyleid(body, index[0]);
                    int[] c = TtitleStyle(id, doc, index[0]);

                    //字体
                    if (c[0] == 0)
                    {
                        XmlElement xe1 = xmlDoc.CreateElement("Text");
                        xe1.InnerText = "中文表名字体错误，应为中文宋体、英文Times New Roman：" + "{" + title[1] + "||" + chapter + "}";
                        root.AppendChild(xe1);
                    }
                    if (c[1] == 0)
                    {
                        XmlElement xe1 = xmlDoc.CreateElement("Text");
                        xe1.InnerText = "中文表名字号错误，应为五号：" + "{" + title[1] + "||" + chapter + "}";
                        root.AppendChild(xe1);
                    }
                    if (c[2] == 0)
                    {
                        XmlElement xe1 = xmlDoc.CreateElement("Text");
                        xe1.InnerText = "中文表名未居中：" + "{" + title[1] + "||" + chapter + "}";
                        root.AppendChild(xe1);
                    }


                    //字体、字号格式化
                    changeFontFontSize(p, "宋体", "Times New Roman", "21");

                    //**************5.24新加中文表名格式检测***********************
                    int[] e = numberstyle(title[1], num.Length);
                    //序号前无空格
                    if (e[0] == 0)
                    {
                        XmlElement xe2 = xmlDoc.CreateElement("Text");
                        xe2.InnerText = "中文表格序号与“表”之间不应有空格：" + "{" + title[1] + "||" + chapter + "}";
                        root.AppendChild(xe2);
                        //去除中文表名序号前空格
                        CNdeleteSpacingBeforeNumber(p);
                    }
                    if (e[1] == 0)
                    {
                        XmlElement xe2 = xmlDoc.CreateElement("Text");
                        xe2.InnerText = "中文序号与表名之间应空两格：" + "{" + title[1] + "||" + chapter + "}";
                        root.AppendChild(xe2);
                        //序号后空格
                        CNSpacingAfterNumebr(p, num);
                    }
                    if (e[2] == 0)
                    {
                        XmlElement xe2 = xmlDoc.CreateElement("Text");
                        xe2.InnerText = "中文序号不是M.N形式：" + "{" + title[1] + "||" + chapter + "}";
                        root.AppendChild(xe2);
                        //加批注
                        if (body.ChildElements.GetItem(index[0]) != null)
                        {
                            if (body.ChildElements.GetItem(index[0]).GetType().ToString() == "DocumentFormat.OpenXml.Wordprocessing.Paragraph")
                            {
                                Paragraph commentPara = (Paragraph)body.ChildElements.GetItem(index[0]);
                                Tool.addComment(doc, commentPara, "中文序号不是M.N形式");
                            }
                        }
                    }
                    if (!correctm(num, chapter))
                    {
                        XmlElement xe2 = xmlDoc.CreateElement("Text");
                        xe2.InnerText = "中文表名序号与章节号不一致：" + "{" + title[1] + "||" + chapter + "}";
                        root.AppendChild(xe2);
                        //序号与章节号修改为一致，同时加批注
                        //修改
                        changeM(p, num, chapter);
                        //加批注
                        if (body.ChildElements.GetItem(index[0]) != null)
                        {
                            if (body.ChildElements.GetItem(index[0]).GetType().ToString() == "DocumentFormat.OpenXml.Wordprocessing.Paragraph")
                            {
                                Paragraph commentPara = (Paragraph)body.ChildElements.GetItem(index[0]);
                                Tool.addComment(doc, commentPara, "中文表名序号与章节号不一致（已改）");
                            }
                        }
                    }
                    if (!correctn(num, numbertablechapter, continued))
                    {
                        XmlElement xe2 = xmlDoc.CreateElement("Text");
                        xe2.InnerText = "中文序号M.N的N未连续：" + "{" + title[1] + "||" + chapter + "}";
                        root.AppendChild(xe2);
                        //修改
                        changeN(p, numbertable, continued);
                    }

                }
                #endregion
                //看是否有英文表名
                #region 英文表名
                if (title[0] != null)
                {
                    //获取英文题目所在段落
                    Paragraph p = (Paragraph)body.ChildElements.GetItem(index[1]);
                    int[] e = EnNumberStyle(title[0], Ennum.Length);
                    if (e[0] == 0)
                    {
                        XmlElement xe2 = xmlDoc.CreateElement("Text");
                        xe2.InnerText = "英文表名开头应为“Tab.”：" + "{" + title[0] + "||" + chapter + "}";
                        root.AppendChild(xe2);
                        //添加上
                        addTab(p, Ennum);
                    }
                    if (e[1] == 0)
                    {
                        XmlElement xe2 = xmlDoc.CreateElement("Text");
                        xe2.InnerText = "“Tab.”后应有且仅有一个空格：" + "{" + title[0] + "||" + chapter + "}";
                        root.AppendChild(xe2);
                        //1空格
                        ENSapcingBeforeNumber(p, Ennum);
                    }
                    if (e[2] == 0)
                    {
                        XmlElement xe2 = xmlDoc.CreateElement("Text");
                        xe2.InnerText = "英文表名序号与表名之间应空两格：" + "{" + title[0] + "||" + chapter + "}";
                        root.AppendChild(xe2);
                        //2空格
                        ENSapcingAfterNumber(p, Ennum);
                    }
                    if (e[3] == 0)
                    {
                        XmlElement xe2 = xmlDoc.CreateElement("Text");
                        xe2.InnerText = "英文表名序号不是M.N形式：" + "{" + title[0] + "||" + chapter + "}";
                        root.AppendChild(xe2);
                        //加批注
                        if (body.ChildElements.GetItem(index[1]) != null)
                        {
                            if (body.ChildElements.GetItem(index[1]).GetType().ToString() == "DocumentFormat.OpenXml.Wordprocessing.Paragraph")
                            {
                                Paragraph commentPara = (Paragraph)body.ChildElements.GetItem(index[1]);
                                Tool.addComment(doc, commentPara, "英文表名序号不是M.N形式");
                            }
                        }
                    }
                    if (!correctm(Ennum, chapter))
                    {
                        XmlElement xe2 = xmlDoc.CreateElement("Text");
                        xe2.InnerText = "英文表名序号与章节号不一致：" + "{" + title[0] + "||" + chapter + "}";
                        root.AppendChild(xe2);
                        //修改且加批注
                        changeM(p, Ennum, chapter);
                    }
                    if (!correctn(Ennum, numbertablechapter, continued))
                    {
                        XmlElement xe2 = xmlDoc.CreateElement("Text");
                        xe2.InnerText = "英文表名序号M.N的N未连续：" + "{" + title[0] + "||" + chapter + "}";
                        root.AppendChild(xe2);
                        //修改且加批注
                        changeN(p, numbertable, continued);
                    }
                    string id = "";
                    id = TtitleStyleid(body, index[1]);
                    int[] c = TtitleStyle(id, doc, index[1]);
                    if (c[0] == 0)
                    {
                        XmlElement xe2 = xmlDoc.CreateElement("Text");
                        xe2.InnerText = "英文表名字体错误，应为Times New Roman：" + "{" + title[0] + "||" + chapter + "}";
                        root.AppendChild(xe2);
                    }
                    if (c[1] == 0)
                    {
                        XmlElement xe2 = xmlDoc.CreateElement("Text");
                        xe2.InnerText = "英文表名字号错误，应为5号：" + "{" + title[0] + "||" + chapter + "}";
                        root.AppendChild(xe2);
                    }
                    if (c[2] == 0)
                    {
                        XmlElement xe2 = xmlDoc.CreateElement("Text");
                        xe2.InnerText = "英文表名没有居中：" + "{" + title[0] + "||" + chapter + "}";
                        root.AppendChild(xe2);
                    }
                    //字体、字号格式化
                    changeFontFontSize(p, "宋体", "Times New Roman", "21");
                }
                #endregion
                //表内文本的检测
                #region 表内文本的检测
                int[] b = TableText(tbls, doc);
                //若字体错误
                if (b[0] == 0)
                {
                    XmlElement xe2 = xmlDoc.CreateElement("Text");
                    xe2.InnerText = "表格内字体错误，应为中文宋体，英文Times New Roman：" + s;
                    root.AppendChild(xe2);
                }
                //若字号错误
                if (b[1] == 0)
                {
                    XmlElement xe2 = xmlDoc.CreateElement("Text");
                    xe2.InnerText = "表格内字号不是五号：" + s;
                    root.AppendChild(xe2);
                }
                if (b[2] == 0)
                {
                    XmlElement xe2 = xmlDoc.CreateElement("Text");
                    xe2.InnerText = "表未居中：" + s;
                    root.AppendChild(xe2);
                }

                //格式化表内字体、字号
                IEnumerable<Paragraph> ps = tbls.Elements<Paragraph>();
                foreach (Paragraph p in ps)
                {
                    changeFontFontSize(p, "宋体", "Times New Roman", "21");
                    Tool.change_center(p);
                }
                #endregion
                //三线表
                if (!correctTable(tbls))
                {
                    XmlElement xe2 = xmlDoc.CreateElement("Text");
                    xe2.InnerText = "不是三线表：" + s;
                    root.AppendChild(xe2);
                }
                //格式化为三线表，不写在上一个里面是可能会检测失误，此函数一般情况下不会出错
                //格式化为三线表
                threeLineTable(tbls);

                //表过宽
                if (!width(tbls, sectpr))
                {
                    XmlElement xe2 = xmlDoc.CreateElement("Text");
                    xe2.InnerText = "表过宽：" + s;
                    root.AppendChild(xe2);
                }
            }
            xmlDoc.Save(xmlFullPath);
        }



        static void changeFontFontSize(Paragraph p, string CNfont, string Enfont, string size)
        {
            Tool.changeFont(p, CNfont, Enfont);
            Tool.changeFontSize(p, size);
        }
        static bool correctTable(DocumentFormat.OpenXml.Wordprocessing.Table t)//三线表 
        {
            int tcCount = 0;
            IEnumerable<TableRow> trList = t.Elements<TableRow>();
            int rowCount = trList.Count<TableRow>();
            TableProperties tpr = t.GetFirstChild<TableProperties>();
            TableBorders tb = tpr.GetFirstChild<TableBorders>();
            if (tpr != null)
            {

                if (tb != null)
                {
                    if (rowCount <= 2)
                    {
                        return true;
                    }
                    foreach (TableRow tr in trList)
                    {
                        tcCount++;
                        IEnumerable<TableCell> tcList = tr.Elements<TableCell>();
                        foreach (TableCell tc in tcList)
                        {
                            TableCellProperties tcp = tc.GetFirstChild<TableCellProperties>();
                            int bottom = 1;
                            if (tcp != null)
                            {
                                TableCellBorders tcb = tcp.GetFirstChild<TableCellBorders>();
                                if (tcb != null)
                                {
                                    if (tcb.GetFirstChild<LeftBorder>() != null)
                                    {
                                        if (tcb.GetFirstChild<LeftBorder>().Val != "nil")
                                        {
                                            return false;
                                        }
                                    }
                                    if (tcb.GetFirstChild<RightBorder>() != null)
                                    {
                                        if (tcb.GetFirstChild<RightBorder>().Val != "nil")
                                        {
                                            return false;
                                        }
                                    }
                                    //第一行
                                    if (tcCount == 1)
                                    {
                                        if (tcb.GetFirstChild<BottomBorder>() != null)
                                        {
                                            if (tcb.GetFirstChild<BottomBorder>().Val == "nil")
                                            {
                                                bottom = 0;
                                            }
                                        }
                                        else
                                        {
                                            if (tb.GetFirstChild<InsideHorizontalBorder>() != null)
                                            {
                                                if (tb.GetFirstChild<InsideHorizontalBorder>().Val == "none")
                                                {
                                                    return false;
                                                }
                                            }

                                        }
                                        if (tcb.GetFirstChild<TopBorder>() != null)
                                        {
                                            if (tcb.GetFirstChild<TopBorder>().Val == "nil")
                                            {
                                                return false;
                                            }
                                        }
                                        else
                                        {
                                            if (tb.GetFirstChild<TopBorder>() != null)
                                            {
                                                if (tb.GetFirstChild<TopBorder>().Val == "none")
                                                {
                                                    return false;
                                                }
                                            }
                                        }
                                    }
                                    //第二行的top
                                    if (tcCount == 2)
                                    {
                                        if (tcb.GetFirstChild<TopBorder>() != null)
                                        {
                                            if (tcb.GetFirstChild<TopBorder>().Val == "nil" && bottom == 0)
                                            {
                                                return false;
                                            }
                                        }
                                    }
                                    //除去第一行和最后一行的其他所有行
                                    if (tcCount != 1 && tcCount != rowCount)
                                    {
                                        if (tcb.GetFirstChild<BottomBorder>() != null)
                                        {
                                            if (tcb.GetFirstChild<BottomBorder>().Val == "single")
                                            {
                                                return false;
                                            }
                                        }
                                        else
                                        {
                                            if (tcCount != 2 && tb.GetFirstChild<InsideHorizontalBorder>() != null && tb.GetFirstChild<InsideHorizontalBorder>().Val == "single")
                                            {
                                                return false;
                                            }
                                        }
                                    }
                                    //最后一行并且不是第二行
                                    if (tcCount == rowCount && tcCount != 2)
                                    {
                                        if (tcb.GetFirstChild<TopBorder>() != null)
                                        {
                                            if (tcb.GetFirstChild<TopBorder>().Val == "single")
                                            {
                                                return false;
                                            }
                                        }
                                        else
                                        {
                                            if (tb.GetFirstChild<InsideHorizontalBorder>() != null && tb.GetFirstChild<InsideHorizontalBorder>().Val == "single")
                                            {
                                                return false;
                                            }
                                        }
                                        if (tcb.GetFirstChild<BottomBorder>() != null)
                                        {
                                            if (tcb.GetFirstChild<BottomBorder>().Val == "nil")
                                            {
                                                return false;
                                            }
                                        }
                                        else
                                        {
                                            if (tb.GetFirstChild<BottomBorder>() != null)
                                            {
                                                if (tb.GetFirstChild<BottomBorder>().Val == "none")
                                                {
                                                    return false;
                                                }
                                            }
                                        }
                                    }
                                }
                                //没有tcb的情况
                                else
                                {
                                    //第一行
                                    if (tcCount == 1)
                                    {
                                        if (tb.GetFirstChild<TopBorder>() != null)
                                        {
                                            if (tb.GetFirstChild<TopBorder>().Val == "none")
                                            {
                                                return false;
                                            }
                                        }
                                        if (tb.GetFirstChild<InsideHorizontalBorder>() != null)
                                        {
                                            if (tb.GetFirstChild<InsideHorizontalBorder>().Val == "none")
                                            {
                                                return false;
                                            }
                                        }
                                    }
                                    //中间行
                                    if (tcCount != 1 && tcCount != rowCount)
                                    {
                                        if (tcCount != 2 && tb.GetFirstChild<InsideHorizontalBorder>() != null && tb.GetFirstChild<InsideHorizontalBorder>().Val == "single")
                                        {
                                            return false;
                                        }
                                    }
                                    //最后一行
                                    if (tcCount == rowCount && tcCount - 1 != rowCount)
                                    {
                                        if (tb.GetFirstChild<InsideHorizontalBorder>() != null && tb.GetFirstChild<InsideHorizontalBorder>().Val == "single")
                                        {
                                            return false;
                                        }
                                        if (tb.GetFirstChild<BottomBorder>() != null)
                                        {
                                            if (tb.GetFirstChild<BottomBorder>().Val == "none")
                                            {
                                                return false;
                                            }
                                        }
                                    }

                                }
                            }
                        }

                    }

                }

            }
            return true;
        }
        //获得表格位置用list保存
        private static List<int> TableLocation(Body body)
        {
            List<int> list = new List<int>(30);
            int l = body.ChildElements.Count();
            for (int i = 0; i < l; i++)
            {
                if (body.ChildElements.GetItem(i).GetType().ToString() == "DocumentFormat.OpenXml.Wordprocessing.Table")
                {
                    list.Add(i);
                }
            }
            return list;
        }
        /*
        location[Chinese title,English title,blank line before table,blank line after table]
        */
        static int[] locationOfTitleAndBlankLine(WordprocessingDocument wordPro, int tablelocation)
        {
            int[] location = { -1, -1, -1, -1 };
            bool[] find = { false, false, false, false };
            Regex[] reg;
            reg = new Regex[9];
            reg[0] = new Regex(@"表[1-9][0-9]*\.[1-9][0-9]*  ");//中文表匹配  关键字段：表m.n空格空格
            reg[1] = new Regex(@"表[1-9][0-9]*\.[1-9][0-9]*");//中文表匹配  关键字段：表m.n
            reg[2] = new Regex(@"表\ *[1-9][0-9]*");//中文表匹配  关键字段：表m
            reg[3] = new Regex(@"Tab.\ *[1-9][0-9]*\.[1-9][0-9]*  ");//英文表匹配  关键字段Tab.空格m.n空格空格
            reg[4] = new Regex(@"Tab. [1-9][0-9]*\.[1-9][0-9]*");//英文表匹配  关键字段Tab.空格m.n
            reg[5] = new Regex(@"Tab.[1-9][0-9]*\.[1-9][0-9]*");//英文表匹配  关键字段Tab.m.n
            reg[6] = new Regex(@"Tab(. [1-9][0-9]*)*");//英文表匹配  关键字段Tab.空格m
            reg[7] = new Regex(@"Tab.[1-9][0-9]*");//英文表匹配  关键字段Tab.m
            reg[8] = new Regex(@"Tab[1-9][0-9]*");//英文表匹配  关键字段Tab.m
            Body body = wordPro.MainDocumentPart.Document.Body;
            //从table往前找
            for (int index = tablelocation - 1; index > tablelocation - 5 && index >= 0; index--)
            {
                if (body.ChildElements.GetItem(index).GetType().ToString() == "DocumentFormat.OpenXml.Wordprocessing.Paragraph")
                {
                    Paragraph p = (Paragraph)body.ChildElements.GetItem(index);
                    string s = p.InnerText.Trim();
                    if (s.Length == 0 && find[2] == false)
                    {
                        if (find[2] == false)//表前空行匹配
                        {
                            location[2] = index;
                            find[2] = true;
                        }
                    }
                    else if (s.Length > 0 && s.Length < 100)//长度过滤
                    {
                        //中文表名匹配
                        for (int i = 0; i <= 2; i++)
                        {
                            Match m = reg[i].Match(s);
                            if (m.Success)
                            {
                                if (find[0] == false && s.Length <= 40)
                                {
                                    location[0] = index;
                                    find[0] = true;
                                    break;
                                }
                            }
                        }
                        //英文表名匹配
                        for (int j = 3; j <= 8; j++)
                        {
                            Match m = reg[j].Match(s);
                            if (m.Success)
                            {
                                if (find[1] == false)
                                {
                                    location[1] = index;
                                    find[1] = true;
                                    break;
                                }
                            }
                        }
                    }
                }
            }
            //从table往后找
            for (int index = tablelocation + 1; index < tablelocation + 3 && index < body.Count(); index++)
            {
                if (body.ChildElements.GetItem(index).GetType().ToString() == "DocumentFormat.OpenXml.Wordprocessing.Paragraph")
                {
                    Paragraph p = (Paragraph)body.ChildElements.GetItem(index);
                    string s = p.InnerText.Trim();
                    if (s.Length == 0 && find[3] == false)
                    {
                        location[3] = index;
                        find[3] = true;
                        break;
                    }
                }
            }
            return location;
        }
        static bool allEnglish(string s)
        {
            return !Regex.IsMatch(s, @"[\u4e00-\u9fa5]");
        }
        //取得表名判断是否为空表 

        //中文
        //*******5.24新增 表格序号检测
        //检测项  1.序号前无空格  
        //       2.序号后两空格
        //       3.是否满足m.n格式
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
            }

            return a;
        }
        //英文
        //*******5.24新增 表格序号检测
        //检测项  1.Tab.正确否
        //        2.Tab.后有空格 
        //         3.序号后两空格
        //       4.是否满足m.n格式
        static int[] EnNumberStyle(string title, int numlen)
        {
            int l = -1;
            int i = -1;
            int[] a = new int[4] { 1, 1, 1, 1 };
            foreach (char c in title)//寻找第一个数字位置
            {
                i++;
                if (c <= '9' && c >= '0')
                {
                    l = i;
                    break;
                }
            }
            //没标号，找不到数字
            if (l == -1)
            {
                a[2] = 0;
                return a;
            }
            if (title.IndexOf("Tab.") < 0)
            {
                a[0] = 0;
                if (title.IndexOf("Tab") >= 0)
                {
                    if (title.IndexOf("Tab ") == -1)
                    {
                        a[1] = 0;
                    }
                }
            }
            else
            {
                if (title.IndexOf("Tab. ") < 0)
                {
                    a[1] = 0;//若没有空格报错
                }
                else
                {
                    //多空格报错
                    if (title.IndexOf("Tab.  ") >= 0)
                    {
                        a[1] = 0;
                    }
                }
            }
            //序号后两空格
            if (l + numlen + 1 < title.Length)
            {
                if (title[l + numlen] != ' ' || title[l + numlen + 1] != ' ' || title[l + numlen + 2] == ' ')
                {
                    a[2] = 0;
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
                            a[3] = 0;
                        }
                    }
                }
                else if (title[l + 2] == '.')//m为两位数
                {
                    for (int j = 3; j < numlen; j++)
                    {
                        if (title[l + j] <= '0' || title[l + j] >= '9')
                        {
                            a[3] = 0;
                        }
                    }
                }
            }
            return a;
        }
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
        static string Ennumber(string title)
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
                    if ((c <= '9' && c >= '0'))
                    {
                        l = i;
                        break;
                    }
                }
                for (j = 0; j < 5; j++)
                {
                    if (j + l < len && j + l >= 0)
                    {
                        if ((title[j + l] >= '0' && title[j + l] <= '9') || title[j + l] == '.')
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
        static bool continuedtable(string[] title)
        {
            if (title[0] != null && title[1] != null)
            {
                if (title[0].Length <= 4 || title[1].Length <= 2)
                {
                    return false;
                }
                else
                {
                    if (title[0].IndexOf("Cont") >= 0 || (title[1].IndexOf("续") >= 0))
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }
            }
            else
            {
                if (title[0] != null && title[1] == null)
                {
                    if (title[0].Length <= 4)
                    {
                        return false;
                    }
                    if (title[0].IndexOf("Cont") >= 0)
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }
                else if (title[0] == null && title[1] != null)
                {
                    if (title[1].Length <= 2)
                    {
                        return false;
                    }
                    if (title[1].IndexOf("续") >= 0)
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }
                return false;
            }
        }

        //字体正确返回1错误返回0
        //字号正确返回1错误返回0
        //Center正确返回1错误返回0
        static int[] TableText(DocumentFormat.OpenXml.Wordprocessing.Table table, WordprocessingDocument doc)
        {
            int[] a = new int[3] { 1, 1, 1 };
            IEnumerable<TableRow> tr = table.Elements<TableRow>();
            foreach (TableRow trs in tr)
            {
                IEnumerable<TableCell> tc = trs.Elements<TableCell>();
                foreach (TableCell tcs in tc)
                {
                    IEnumerable<DocumentFormat.OpenXml.Wordprocessing.Paragraph> paras = tcs.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>();
                    foreach (DocumentFormat.OpenXml.Wordprocessing.Paragraph p in paras)
                    {
                        if (Tool.correctfonts(p, doc, "宋体", "Times New Roman") == false)
                        {
                            a[0] = 0;
                        }
                        if (Tool.correctsize(p, doc, "21") == false)
                        {
                            a[1] = 0;
                        }
                    }
                }
            }
            //居中检测
            TableProperties tpr = table.GetFirstChild<TableProperties>();
            if (tpr != null)
            {
                if (tpr.GetFirstChild<TableJustification>() != null)
                {
                    TableJustification tj = tpr.GetFirstChild<TableJustification>();
                    if (tj.Val.ToString() != "center" && tj.Val.ToString() != null)
                    {
                        a[2] = 0;
                    }
                }
                else
                {
                    a[2] = 0;
                }
            }
            return a;
        }
        //获得title的Pstyle
        static string TtitleStyleid(Body body, int location)
        {
            string id = "";
            if (location > 0)
            {
                if (body.ChildElements.GetItem(location - 1) != null)
                {
                    if (body.ChildElements.GetItem(location - 1).GetFirstChild<ParagraphProperties>() != null)
                    {
                        if (body.ChildElements.GetItem(location - 1).GetFirstChild<ParagraphProperties>().GetFirstChild<ParagraphStyleId>() != null)
                        {
                            id = body.ChildElements.GetItem(location - 1).GetFirstChild<ParagraphProperties>().GetFirstChild<ParagraphStyleId>().Val;
                        }
                    }
                }
            }
            return id;
        }
        //中文标题样式字体字号和居中
        static int[] TtitleStyle(string id, WordprocessingDocument doc, int location)
        {
            int[] a = new int[3] { 1, 1, 1 };
            IEnumerable<DocumentFormat.OpenXml.Wordprocessing.Style> style = doc.MainDocumentPart.StyleDefinitionsPart.Styles.Elements<DocumentFormat.OpenXml.Wordprocessing.Style>();
            Body body = doc.MainDocumentPart.Document.Body;
            int i = -1;
            if (body.ChildElements.GetItem(location) != null)
            {
                DocumentFormat.OpenXml.Wordprocessing.Paragraph p = (DocumentFormat.OpenXml.Wordprocessing.Paragraph)body.ChildElements.GetItem(location);
                if (p != null)
                {
                    if (!Tool.correctfonts(p, doc, "宋体", "Times New Roman"))
                    {
                        a[0] = 0;
                    }
                    //return a;
                    if (Tool.correctsize(p, doc, "21") == false)
                    {
                        a[1] = 0;
                    }

                    if (id != null)
                    {
                        foreach (DocumentFormat.OpenXml.Wordprocessing.Style s in style)
                        {
                            i++;
                            if (s.StyleId == id)
                                break;
                        }
                    }
                    DocumentFormat.OpenXml.Wordprocessing.Style st = null;
                    if (i >= 0)
                    {
                        st = style.ElementAt<DocumentFormat.OpenXml.Wordprocessing.Style>(i);
                    }
                    DocumentFormat.OpenXml.Wordprocessing.Style st2 = null;
                    foreach (DocumentFormat.OpenXml.Wordprocessing.Style s in style)
                    {
                        if (s.StyleName != null)
                        {
                            if (s.StyleName.Val == "Normal")
                            {
                                st2 = s;
                                break;
                            }
                        }
                    }
                    if (body.ChildElements.GetItem(location) != null)
                    {
                        if (p != null)
                        {
                            IEnumerable<Run> run = p.Elements<Run>();
                            foreach (Run r in run)
                            {
                                RunProperties rPr = null;
                                if (r != null)
                                {
                                    rPr = r.GetFirstChild<RunProperties>();
                                }
                                if (rPr != null)
                                {
                                    //居中1
                                    ParagraphProperties ppr = p.GetFirstChild<ParagraphProperties>();
                                    if (ppr != null)
                                    {
                                        if (ppr.GetFirstChild<Justification>() != null)
                                        {
                                            Justification tj = ppr.GetFirstChild<Justification>();
                                            if (tj.Val != "center")
                                            {
                                                a[2] = 0;
                                            }
                                        }
                                        else if (st != null && id != null)
                                        {
                                            if (st.StyleParagraphProperties != null)
                                            {
                                                if (st.StyleParagraphProperties.Justification != null)
                                                {

                                                    if (st.StyleParagraphProperties.Justification.Val.ToString() != "center")
                                                    {
                                                        a[2] = 0;
                                                    }
                                                }
                                                else
                                                {
                                                    if (st2 != null)
                                                    {
                                                        if (st2.StyleParagraphProperties != null)
                                                        {
                                                            if (st2.StyleParagraphProperties.Justification != null)
                                                            {

                                                                if (st2.StyleParagraphProperties.Justification.Val.ToString() != "center")
                                                                {
                                                                    a[2] = 0;
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }

                                        }
                                        if (ppr.Indentation != null)
                                        {
                                            if (ppr.Indentation.FirstLine != null)
                                            {
                                                if (ppr.Indentation.FirstLine != "0")
                                                {
                                                    a[2] = 0;
                                                }
                                            }
                                        }
                                        else if (st != null)
                                        {
                                            if (st.StyleParagraphProperties != null && id != "")
                                            {
                                                if (st.StyleParagraphProperties.Indentation != null)
                                                {
                                                    if (st.StyleParagraphProperties.Indentation.FirstLine != "0")
                                                    {
                                                        a[2] = 0;
                                                    }
                                                }
                                                else
                                                {
                                                    if (st2 != null)
                                                    {
                                                        if (st2.StyleParagraphProperties != null && id != "")
                                                        {
                                                            if (st2.StyleParagraphProperties.Indentation != null)
                                                            {
                                                                if (st2.StyleParagraphProperties.Indentation.FirstLine != "0")
                                                                {
                                                                    a[2] = 0;
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
                        }
                    }
                }
            }
            return a;
        }


        //获得表所在章节号
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

        //编号中的m应与章号一致
        static bool correctm(string num, string chapter)
        {
            char m1 = '\0';
            char m2 = '\0';
            if (chapter != "")
            {
                m1 = chapter[0];
            }
            if (num != "")
            {
                m2 = num[0];
            }
            //带章节号的比对
            if (m1 >= '0' && m1 <= '9')
            {
                if (m1 == m2)
                {
                }
                else
                {
                    return false;
                }
            }
            else if (m1 == '附')
            {
                //附录X
                if (chapter.Length >= 3)
                {
                    if (chapter[2] != m2)
                    {
                        return false;
                    }
                }
            }
            return true;
        }
        //序号连续
        static bool correctn(string num, int numbertable, int continued)
        {
            int i = num.IndexOf('.');
            string n = "";
            if (i < 0)
            {
                return false;
            }
            else
            {
                if (i + 1 < num.Length)
                {
                    if (i + 2 < num.Length)
                    {
                        if (num[i + 1] >= '1' && num[i + 1] <= '9')
                        {
                            if (num[i + 2] >= '0' && num[i + 2] <= '9')
                            {
                                n = num.Substring(i + 1, 2);
                            }
                            else
                            {
                                n = num.Substring(i + 1, 1);
                            }
                        }
                        else
                        {
                            return false;
                        }
                    }
                    else
                    {

                        if (num[i + 1] >= '1' && num[i + 1] <= '9')
                        {
                            n = num.Substring(i + 1, 1);
                        }
                        else
                        {
                            return false;
                        }
                    }
                }
            }
            if (n != "")
            {
                if (n != (numbertable - continued).ToString())
                {
                    return false;
                }
            }
            return true;
        }

        static bool width(DocumentFormat.OpenXml.Wordprocessing.Table table, SectionProperties sectPr)
        {

            uint width = 0;
            uint pagewidth = 0;
            uint leftmargin = 0;
            uint rightmargin = 0;
            //获得表宽
            if (table.GetFirstChild<TableProperties>() != null)
            {
                if (table.GetFirstChild<TableProperties>().GetFirstChild<TableWidth>() != null)
                {
                    width = Convert.ToUInt32(table.GetFirstChild<TableProperties>().GetFirstChild<TableWidth>().Width.Value);
                }
            }
            if (width == 0)
            {
                if (table.GetFirstChild<TableGrid>() != null)
                {
                    IEnumerable<GridColumn> gridCols = table.GetFirstChild<TableGrid>().Elements<GridColumn>();
                    foreach (GridColumn gridCol in gridCols)
                    {
                        width += Convert.ToUInt32(gridCol.Width.Value);
                    }
                }
            }
            //获得左、右间距、页宽
            if (sectPr != null)
            {
                if (sectPr.GetFirstChild<PageMargin>() != null)
                {
                    if (sectPr.GetFirstChild<PageMargin>().Left != null)
                    {
                        leftmargin = sectPr.GetFirstChild<PageMargin>().Left.Value;
                    }
                    if (sectPr.GetFirstChild<PageMargin>().Right != null)
                    {
                        rightmargin = sectPr.GetFirstChild<PageMargin>().Right;
                    }
                }
                if (sectPr.GetFirstChild<PageSize>() != null)
                {
                    pagewidth = sectPr.GetFirstChild<PageSize>().Width.Value;
                }
            }
            //1.若是浮动型
            if (table.GetFirstChild<TableProperties>() != null)
            {
                if (table.GetFirstChild<TableProperties>().GetFirstChild<TablePositionProperties>() != null)
                {
                    TablePositionProperties tblpPr = table.GetFirstChild<TableProperties>().GetFirstChild<TablePositionProperties>();
                    if(tblpPr.HorizontalAnchor!= null)
                    {
                        string s = tblpPr.HorizontalAnchor.Value.ToString();
                        if (tblpPr.HorizontalAnchor.Value.ToString() == "Margin")
                        {
                            if (tblpPr.TablePositionX != null && tblpPr.TablePositionXAlignment == null)
                            {
                                if (tblpPr.TablePositionX.Value >= 0 && tblpPr.TablePositionX.Value + width + leftmargin < pagewidth - rightmargin)
                                {
                                    return true;
                                }
                                else
                                {
                                    return false;
                                }
                            }
                            if (tblpPr.TablePositionX == null && tblpPr.TablePositionXAlignment == null)
                            {
                                return true;
                            }
                            if (tblpPr.TablePositionXAlignment != null)
                            {
                                if (pagewidth - leftmargin - rightmargin >= width)
                                {
                                    return true;
                                }
                                else
                                {
                                    return false;
                                }
                            }
                        }
                        if (tblpPr.HorizontalAnchor.Value.ToString() == "Page")
                        {
                            if (tblpPr.TablePositionX != null && tblpPr.TablePositionXAlignment == null)
                            {
                                if (tblpPr.TablePositionX.Value >= leftmargin && tblpPr.TablePositionX.Value + width < pagewidth - rightmargin)
                                {
                                    return true;
                                }
                                else
                                {
                                    return false;
                                }
                            }
                            if (tblpPr.TablePositionX == null && tblpPr.TablePositionXAlignment == null)
                            {
                                return true;
                            }
                            if (tblpPr.TablePositionXAlignment != null)
                            {
                                if (pagewidth - leftmargin - rightmargin >= width)
                                {
                                    return true;
                                }
                                else
                                {
                                    return false;
                                }
                            }
                        }
                    }
                    
                }
                //若是无环绕型
                else if (table.GetFirstChild<TableProperties>().GetFirstChild<TableIndentation>() != null)
                {
                    int indentation = table.GetFirstChild<TableProperties>().GetFirstChild<TableIndentation>().Width.Value;
                    if (indentation < 0)
                    {
                        return false;
                    }
                    else
                    {
                        if (width - indentation + leftmargin > pagewidth - rightmargin)
                        {
                            return false;
                        }
                        else
                        {
                            return true;
                        }
                    }
                }
                else
                {
                    if (table.GetFirstChild<TableProperties>().GetFirstChild<TableJustification>() != null)
                    {
                        if (width > pagewidth - leftmargin - rightmargin)
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
        static SectionProperties sectPr(int location, Body body)
        {
            SectionProperties sectPr = new SectionProperties();
            int flag = 0;
            for (int i = location; i < body.ChildElements.Count(); i++)
            {
                if (body.ChildElements.GetItem(i).GetFirstChild<ParagraphProperties>() != null)
                {
                    if (body.ChildElements.GetItem(i).GetFirstChild<ParagraphProperties>().GetFirstChild<SectionProperties>() != null)
                    {
                        flag = 1;
                        sectPr = body.ChildElements.GetItem(i).GetFirstChild<ParagraphProperties>().GetFirstChild<SectionProperties>();
                        return sectPr;
                    }
                }
            }
            if (flag == 0)
            {
                if (body.ChildElements.GetItem(body.ChildElements.Count() - 1).GetType().ToString() == "DocumentFormat.OpenXml.Wordprocessing.SectionProperties")
                {
                    sectPr = (SectionProperties)body.ChildElements.GetItem(body.ChildElements.Count() - 1);
                }
            }
            return sectPr;
        }


        //三线表
        static void threeLineTable(Table t)
        {
            //获取行数
            IEnumerable<TableRow> trs = t.Elements<TableRow>();
            int len = trs.Count<TableRow>();
            if (t.GetFirstChild<TableProperties>() != null)
            {
                TableProperties tblPr = t.GetFirstChild<TableProperties>();
                if (tblPr.TableBorders != null)
                {
                    tblPr.RemoveChild<TableBorders>(tblPr.TableBorders);
                }
                TableBorders tableBorders1 = new TableBorders();
                TopBorder topBorder1 = new TopBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
                LeftBorder leftBorder1 = new LeftBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
                BottomBorder bottomBorder1 = new BottomBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
                RightBorder rightBorder1 = new RightBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
                InsideHorizontalBorder insideHorizontalBorder1 = new InsideHorizontalBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
                InsideVerticalBorder insideVerticalBorder1 = new InsideVerticalBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };

                tableBorders1.Append(topBorder1);
                tableBorders1.Append(leftBorder1);
                tableBorders1.Append(bottomBorder1);
                tableBorders1.Append(rightBorder1);
                tableBorders1.Append(insideHorizontalBorder1);
                tableBorders1.Append(insideVerticalBorder1);

                tblPr.AppendChild<TableBorders>(tableBorders1);
            }
            int i = 0;
            foreach (TableRow tr in trs)
            {
                i++;
                /**************格式化第一行开始**********************/
                if (i == 1)
                {
                    IEnumerable<TableCell> tcs = tr.Elements<TableCell>();
                    foreach (TableCell tc in tcs)
                    {
                        if (tc.TableCellProperties != null)
                        {
                            TableCellProperties tcPr = tc.TableCellProperties;
                            if (tcPr.TableCellBorders != null)
                            {
                                tcPr.RemoveChild<TableCellBorders>(tcPr.TableCellBorders);
                            }
                            TableCellBorders tableCellBorders1 = new TableCellBorders();
                            TopBorder topBorder1 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
                            BottomBorder bottomBorder1 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };

                            tableCellBorders1.Append(topBorder1);
                            tableCellBorders1.Append(bottomBorder1);
                            tcPr.Append(tableCellBorders1);
                        }
                    }
                }
                /**************格式化最后一行开始**********************/
                else if (i == len)
                {
                    IEnumerable<TableCell> tcs = tr.Elements<TableCell>();
                    foreach (TableCell tc in tcs)
                    {
                        if (tc.TableCellProperties != null)
                        {
                            TableCellProperties tcPr = tc.TableCellProperties;
                            if (tcPr.TableCellBorders != null)
                            {
                                tcPr.RemoveChild<TableCellBorders>(tcPr.TableCellBorders);
                            }
                            TableCellBorders tableCellBorders1 = new TableCellBorders();
                            if (len == 2)
                            {
                                TopBorder topBorder1 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
                                BottomBorder bottomBorder1 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
                                tableCellBorders1.Append(topBorder1);
                                tableCellBorders1.Append(bottomBorder1);
                                tcPr.Append(tableCellBorders1);
                            }
                            else
                            {
                                BottomBorder bottomBorder1 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
                                tableCellBorders1.Append(bottomBorder1);
                                tcPr.Append(tableCellBorders1);
                            }
                        }
                    }
                    if (len == 2)
                    {
                        break;
                    }
                }
                /**************格式化第2行开始（且总行数大于2）**********************/
                else if (i == 2)
                {
                    IEnumerable<TableCell> tcs = tr.Elements<TableCell>();
                    foreach (TableCell tc in tcs)
                    {
                        if (tc.TableCellProperties != null)
                        {
                            TableCellProperties tcPr = tc.TableCellProperties;
                            if (tcPr.TableCellBorders != null)
                            {
                                tcPr.RemoveChild<TableCellBorders>(tcPr.TableCellBorders);
                            }
                            TableCellBorders tableCellBorders1 = new TableCellBorders();
                            TopBorder topBorder1 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };

                            tableCellBorders1.Append(topBorder1);
                            tcPr.Append(tableCellBorders1);
                        }
                    }
                }
                /**************格式化中间行开始（且总行数大于2）**********************/
                else
                {
                    IEnumerable<TableCell> tcs = tr.Elements<TableCell>();
                    foreach (TableCell tc in tcs)
                    {
                        if (tc.TableCellProperties != null)
                        {
                            TableCellProperties tcPr = tc.TableCellProperties;
                            if (tcPr.TableCellBorders != null)
                            {
                                tcPr.RemoveChild<TableCellBorders>(tcPr.TableCellBorders);
                            }
                        }
                    }
                }
            }
        }

        //去除中文表名序号前空格
        static void CNdeleteSpacingBeforeNumber(Paragraph p)
        {
            string s = p.InnerText;
            int index = s.IndexOf('表');
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
        //保证中文表名序号后空格
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

        //修改M的值
        //chapter为章节标题，不是章节号
        static void changeM(Paragraph p, string num, string chapter)
        {
            string s = p.InnerText;
            if (s.IndexOf(num) == -1)
                return;
            int start = s.IndexOf(num);
            if (start == -1)
                return;

            Match m = Regex.Match(chapter, @"\d+");
            if (m.Index == -1)
                return;
            string correctM = m.Value;

            //.的位置
            m = Regex.Match(s, @"\d\.");
            int dian = m.Value.IndexOf('.') + m.Index;
            if (dian == -1)
                return;
            int Indexm = num.IndexOf(correctM);
            if (Indexm != 0)
            {
                s = s.Substring(0, start) + correctM + s.Substring(dian);
            }
            //替换此段落内容
            Tool.replaceText(p, s);
        }

        //修改N的值
        //chapter为章节标题，不是章节号
        static void changeN(Paragraph p, int numbertable, int continued)
        {
            string s = p.InnerText;
            string correctN = (numbertable - continued).ToString();

            //.的位置
            Match m = Regex.Match(s, @"\d\.");
            int dian = m.Value.IndexOf('.') + m.Index;
            if (dian == -1)
                return;
            m = Regex.Match(s, @".\d+");
            int nLen = m.Length - 1;
            if (s.Substring(dian).IndexOf(correctN) != 1)
            {
                s = s.Substring(0, dian + 1) + correctN + s.Substring(dian + nLen + 1);
            }
            //替换此段落内容
            Tool.replaceText(p, s);
        }
        static void addTab(Paragraph p, string num)
        {
            string s = p.InnerText;
            int index = s.IndexOf(num);
            if (index == -1)
            {
                return;
            }
            else
            {
                s = "Tab. " + s.Substring(index);
            }
            //替换此段落内容
            Tool.replaceText(p, s);
        }
        static void ENSapcingBeforeNumber(Paragraph p, string number)
        {
            string s = p.InnerText;
            int index1 = s.IndexOf(number);
            if (index1 == -1)
            {
                return;
            }
            //若没有空格
            int index = index1;
            if (s.Substring(0, index).IndexOf(' ') == -1)
            {
                s = s.Substring(0, index) + ' ' + s.Substring(index);
            }
            //若多个空格
            else if (s.Substring(0, index).IndexOf("  ") != -1)
            {
                char c = '1';
                if (index > 0)
                    c = s[--index];
                while (index >= 0 && c == ' ')
                {
                    index--;
                    c = s[index];
                }
                s = s.Substring(0, index) + ' ' + s.Substring(index1);
            }
            //替换此段落内容
            Tool.replaceText(p, s);
        }
        private static void ENSapcingAfterNumber(Paragraph p, string ennum)
        {
            string s = p.InnerText;
            int index = s.IndexOf(ennum);
            char c = '\0';
            if (index + ennum.Length + 1 < s.Length)
                c = s[index + ennum.Length];
            int kongge = index + ennum.Length;
            while (c == ' ')
            {
                if (kongge == s.Length)
                    break;
                c = s[++kongge];
            }
            s = s.Substring(0, index + ennum.Length) + "  " + s.Substring(kongge);
            //替换此段落内容
            Tool.replaceText(p, s);
        }
    }
}