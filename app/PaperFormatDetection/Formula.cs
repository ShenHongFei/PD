using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Xml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using PaperFormatDetection.Tools;
using PaperFormatDetection.Frame;
using System.Linq;
using System.IO;
using DocumentFormat.OpenXml;
namespace PaperFormatDetection.Format

{
    public class Formula : ModuleFormat
    {
        public Formula(List<Module> modList, PageLocator locator, int masterType)
            : base(modList, locator, masterType)
        {

        }

        public override void getStyle(WordprocessingDocument doc, String fileName)
        {
            string xmlFullPath = fileName + "\\Formula.xml";
            CreateXmlFile(xmlFullPath);
            getFormulaXml(doc,xmlFullPath);
            //Console.ReadLine();
        }

        public static void CreateXmlFile(string xmlPath)
        {
            XmlDocument xmlDocx = new XmlDocument();
            //创建类型声明节点  
            XmlNode node = xmlDocx.CreateXmlDeclaration("1.0", "utf-8", "");
            xmlDocx.AppendChild(node);
            //创建根节点  
            XmlNode root = xmlDocx.CreateElement("Formula");
            XmlElement xe1 = xmlDocx.CreateElement("spErroInfo");
            xe1.SetAttribute("name", "特殊错误信息");
            XmlElement xe2 = xmlDocx.CreateElement("partName");
            xe2.SetAttribute("name", "提示名称");
            XmlElement xe3 = xmlDocx.CreateElement("Text");
            xe3.InnerText = "-----------------公式-----------------";
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

        public static void getFormulaXml(WordprocessingDocument doc,string xmlFullPath)
        {
            Body body = doc.MainDocumentPart.Document.Body;
            IEnumerable<Paragraph> paras = body.Elements<Paragraph>();
            List<Paragraph> plist = toList(paras);
            var list = body.ChildElements;
            Paragraph temp = new Paragraph();
            int count = 0;
            int number = -1;
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(xmlFullPath);
            XmlNode sproot = xmlDoc.SelectSingleNode("Formula/spErroInfo");
            string CharpterNum = "";
            int CharpterFormulaN = 0;
            foreach (var t in list)
            {
                count++;
                if (t.GetType() == temp.GetType())
                {
                    Paragraph p = (Paragraph)t;
                    Run r = t.GetFirstChild<Run>();
                    if (r != null)
                    {
                        number++;
                        EmbeddedObject ob = r.GetFirstChild<EmbeddedObject>();
                        OfficeMath om = r.GetFirstChild<OfficeMath>();

                        if (ob != null || om != null)
                        {
                            if (plist != null && number < plist.Count - 1)
                            {
                                //Console.WriteLine("pppp");
                                
                                
                                List<int> titleposition = Tool.getTitlePosition(doc);
                                string chaper = Chapter(titleposition, count, body);
                                if (CharpterNum != chaper.Trim().Substring(0, 1))
                                {
                                    CharpterNum = chaper.Trim().Substring(0, 1);
                                    CharpterFormulaN = 0;
                                }
                                string content = getFullText(p);
                                string showcontent = content.Trim();
                                int firstnum = 0;
                                bool rightformula = false;
                                bool formula = false;
                                bool samep = false;
                                bool haspicname = false;
                                int i;
                                

                                for (i = 0; i < content.Length; i++)
                                {
                                    if (content[i] == '图')
                                        samep = true;
                                }

                                for (i = 0; i < content.Length; i++)
                                {
                                    if (content[i] > 48 && content[i] < 58 && firstnum == 0 && samep == false)
                                    {
                                        
                                        rightformula = true;
                                        formula = true;
                                        if (i != 0)
                                            firstnum = i;
                                    }
                                }
                                if (!rightformula)
                                {

                                    Paragraph nextl = plist[number + 1];
                                    string nl = getFullText(nextl);
                                   

                                    for (i = 0; i < nl.Length; i++)
                                    {
                                        if (nl[i] == '图' || nl[i] == 'F')
                                        {
                                            haspicname = true;

                                        }
                                    }
                                }




                                if (!samep && !haspicname)
                                {
                                    formula = true;
                                    CharpterFormulaN++;
                                }

                                if (formula && !rightformula)
                                {
                                    //Console.WriteLine("ooo");
                                    XmlElement xml = xmlDoc.CreateElement("Text");
                                    xml.InnerText = "公式缺少编号或编号格式有误{" + CharpterNum + "}";
                                    sproot.AppendChild(xml);
                                    Run run = new Run { };
                                    Text text = new Text { };
                                    text.Text ="       "+ "(" + CharpterNum + "." + CharpterFormulaN.ToString() + ")";
                                    run.Append(text);
                                    p.Append(run);
                                    addComment(doc, p, "请调整公式与公式序号之间的距离，使公式部分居中显示");
                                    
                                    
                                }

                                

                                ParagraphProperties ppr = p.GetFirstChild<ParagraphProperties>();
                                if (ppr != null && formula)
                                {
                                    
                                    if (ppr.GetFirstChild<Justification>() != null)
                                    {
                                        if (ppr.GetFirstChild<Justification>().Val != null)
                                        {
                                            if (ppr.GetFirstChild<Justification>().Val != "right")
                                            {
                                                XmlElement xml = xmlDoc.CreateElement("Text");
                                                xml.InnerText = "公式整行应右对齐";
                                                sproot.AppendChild(xml);
                                                if (showcontent != null)
                                                {
                                                 

                                                    ppr.GetFirstChild<Justification>().Val = JustificationValues.Right;
                                                   
                                                }
                                            }
                                        }
                                    }
                                    else
                                    {
                                        Justification justification = new Justification() { Val = JustificationValues.Right };
                                        ppr.Append(justification);
                                    }


                                    if (ppr.GetFirstChild<SpacingBetweenLines>() != null)
                                    {
                                        if (ppr.GetFirstChild<SpacingBetweenLines>().BeforeLines == null || ppr.GetFirstChild<SpacingBetweenLines>().Before == null ||
                                            ppr.GetFirstChild<SpacingBetweenLines>().AfterLines == null || ppr.GetFirstChild<SpacingBetweenLines>().After == null)
                                        {
                                            ppr.RemoveChild<SpacingBetweenLines>(ppr.GetFirstChild<SpacingBetweenLines>());
                                            SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines()
                                            {
                                                Line = "300",
                                                LineRule = LineSpacingRuleValues.Auto,
                                                BeforeLines = 50,
                                                AfterLines = 50
                                            };
                                            ppr.Append(spacingBetweenLines1);
                                        }

                                      
                                        else
                                        {
                                            ppr.GetFirstChild<SpacingBetweenLines>().BeforeLines = 50;
                                            ppr.GetFirstChild<SpacingBetweenLines>().AfterLines = 50;
                                            ppr.GetFirstChild<SpacingBetweenLines>().Before.Value = "157";
                                            ppr.GetFirstChild<SpacingBetweenLines>().After.Value = "157";
                                        }

                                    }
                                    else
                                    {
                                       
                                        SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines()
                                        {
                                            Before = "157",
                                            BeforeLines = 50,
                                            After = "157",
                                            AfterLines = 50
                                        };
                                        ppr.Append(spacingBetweenLines1);
                                    }
                                }
                            }
                        }
                    }
                }
            }
            xmlDoc.Save(xmlFullPath);
        }
                                

                                    



                          


                      






  
        public static List<Paragraph> toList(IEnumerable<Paragraph> paras)
        {
            List<Paragraph> list = new List<Paragraph>();
            foreach (Paragraph p in paras)
            {
                Paragraph t = p as Paragraph;
                if (t != null)
                {
                    Run r = t.GetFirstChild<Run>();
                    if(r!=null)
                    list.Add(t);
                }
            }
            return list;
        }
        public static String getFullText(Paragraph p)
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
        public static int Item(List<int> titlePosition, int location, Body body)
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

    }
}
