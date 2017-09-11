using Saxon.Api;
using System.Xml;
using System.IO;
using System.Collections.Generic;

namespace PaperFormatDetection.Frame
{
    /**
     * 对比规则类，继承自ModuleProcessor
     */
    public class RuleComparator : ModuleProcessor
    {   

        public RuleComparator(string paperPath, List<Module> modList) : base(paperPath, modList)
        {

        }

        public override void excute()
        {
            compare();
        }

        private void compare()
        {
            //遍历所有子节点 
            foreach (Module mod in modList)
            {
                if (mod.Detect == true && mod.XqueryUse == true)
                {
                    string className = mod.ClassName;//模块名称
                    string paperName = Path.GetFileNameWithoutExtension(paperPath);//待测论文文件夹名称
                    string compareName = className;//规则文件名称 
                    compareModule(paperName, compareName);
                }
            }
        }

        private void compareModule(string paperName,string compareName) {
            Processor pr = new Processor();
            XQueryCompiler xqc = pr.NewXQueryCompiler();

            string comparePath = "Papers\\" + paperName +"\\" + compareName + "_rule.xml";
            string curPath = Directory.GetCurrentDirectory();
            string configPath = curPath + "\\Config\\" + compareName + ".config";
            string errPath = "Papers\\" + paperName + "\\" + compareName + "_err.xml";

            XmlDocument doc = new XmlDocument();
            doc.Load(comparePath);
            XmlElement root = doc.DocumentElement;
            XmlNodeList ruleNodes = root.GetElementsByTagName("rule");

            XmlDocument errDoc = new XmlDocument();
            XmlNode node = errDoc.CreateXmlDeclaration("1.0", "utf-8", "");//创建类型声明节点
            errDoc.AppendChild(node);
            XmlNode err = errDoc.CreateElement("error");//创建根节点

            foreach (XmlNode xrule in ruleNodes)
            {
                string rule = xrule.InnerText;

                XQueryExecutable xqe = xqc.Compile(rule);
                XQueryEvaluator xqev = xqe.Load();
                if (!paperName.Contains("#"))
                {
                    XdmValue result = xqev.Evaluate(); //对比结果

                    foreach (XdmItem r in result)
                    {
                        string rstr = r.ToString();
                        string[] rarr = rstr.Split('/');
                        string errstr = "for $x in doc(\"" + configPath + "\")/conf/";
                        if (rarr.Length == 3)
                        {
                            errstr += "module[@tag=" + "\"" + rarr[0] + "\"]/";
                            errstr += "section[@tag=" + "\"" + rarr[1] + "\"]/";
                            errstr += "item[@tag=" + "\"" + rarr[2] + "\"]/error-msg/text() return $x";
                        }
                        else
                        {
                            errstr += "module[@tag=" + "\"" + rarr[0] + "\"]/";
                            errstr += "section[@tag=" + "\"" + rarr[1] + "\"]/";
                            errstr += "part[@tag=" + "\"" + rarr[2] + "\"]/";
                            errstr += "item[@tag=" + "\"" + rarr[3] + "\"]/error-msg/text() return $x";
                        }
                        XmlElement err_rule = errDoc.CreateElement("err-rule");
                        //string errstr = rstr;

                        err_rule.InnerText = errstr;
                        err.AppendChild(err_rule);
                    }
                    errDoc.AppendChild(err);
                    errDoc.Save(errPath);
                }
                else
                {
                    errDoc.AppendChild(err);
                    errDoc.Save(errPath);
                }
            }
        }
    }
}
