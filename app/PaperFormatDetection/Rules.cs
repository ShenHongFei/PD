using System;
using System.Xml;
using System.IO;
using System.Collections.Generic;

namespace PaperFormatDetection.Frame
{
    /**
     * 生成对比规则类，继承自ModuleProcessor
     */
    public class Rules : ModuleProcessor
    {   
        public Rules(string templatePath, string paperPath, List<Module> modList) : base(templatePath, paperPath, modList)
        {

        }

        public override void excute()
        {
            generateRules();
        }
        //生成对比文件
        private void generateRules()
        {

            //遍历所有子节点 
            foreach (Module  mod in modList)
            {
                if (mod.Detect == true && mod.XqueryUse == true)
                {
                    string configName = mod.Rule;//配置文件名称
                    string templateName = Path.GetFileNameWithoutExtension(templatePath);//论文模板文件夹名称
                    string paperName = Path.GetFileNameWithoutExtension(paperPath);//待测论文文件夹名称
                    string ruleName = mod.ClassName;
                    createRule(ruleName, templateName, paperName, configName);
                }
            }
        }

        //创建对比规则
        private void createRule(string ruleName, string templateName, string paperName, string configName)
        {
            try
            {
                string rulePath = "Papers\\" + paperName + "\\" + ruleName + "_rule.xml";

                string curPath = Directory.GetCurrentDirectory();
                string templatePath = curPath + "/Templates/" + templateName + "/" + ruleName + ".xml";
                string paperPath = curPath +"/Papers/" + paperName + "/" + ruleName + ".xml";

                XmlDocument ruleDoc = new XmlDocument();
                XmlNode node = ruleDoc.CreateXmlDeclaration("1.0", "utf-8", "");//创建类型声明节点
                ruleDoc.AppendChild(node);
                XmlNode root = ruleDoc.CreateElement("Compare");//创建根节点 

                XmlDocument configDoc = new XmlDocument();
                configDoc.Load("Config\\" + configName); //加载xml文件
                XmlNode xmod = configDoc.SelectSingleNode("conf").SelectSingleNode("module");
                XmlElement mod = (XmlElement)xmod; //转换类型
                string modTag = mod.GetAttribute("tag");

                XmlNodeList secList = xmod.ChildNodes;
                foreach (XmlNode xsec in secList)
                {
                    XmlElement sec = (XmlElement)xsec;
                    string secTag = sec.GetAttribute("tag");

                    string sonName = xsec.FirstChild.Name;
                    if (sonName == "part")
                    {
                        XmlNodeList partList = xsec.ChildNodes;
                        foreach (XmlNode xpart in partList)
                        {
                            XmlElement part = (XmlElement)xpart;
                            string partTag = part.GetAttribute("tag");

                            XmlNodeList itemList = xpart.ChildNodes;
                            foreach (XmlNode xitem in itemList)
                            {
                                XmlElement item = (XmlElement)xitem;
                                string itemTag = item.GetAttribute("tag");
                                
                                string rulestr = "for $Standard in doc(\"" + templatePath + "\")/"
                                    + modTag + "/" + secTag + "/" + partTag + "/" + itemTag + " ";
                                rulestr += "for $Under_test in doc(\"" + paperPath + "\")/"
                                    + modTag + "/" + secTag + "/" + partTag + "/" + itemTag + " "; ;
                                rulestr += "return if ( $Standard != $Under_test ) ";
                                rulestr += " then '" + modTag + "/" + secTag + "/" + partTag + "/" + itemTag + "'";
                                rulestr += " else ()";
                                XmlElement xe = ruleDoc.CreateElement("rule");
                                xe.InnerText = rulestr;
                                root.AppendChild(xe);

                            }
                        }
                    }
                    else
                    {
                        XmlNodeList itemList = xsec.ChildNodes;
                        foreach (XmlNode xitem in itemList)
                        {
                            XmlElement item = (XmlElement)xitem;
                            string itemTag = item.GetAttribute("tag");

                            string rulestr = "for $Standard in doc(\"" + templatePath + "\")/"
                              + modTag + "/" + secTag + "/" + itemTag + " ";
                            rulestr += "for $Under_test in doc(\"" + paperPath + "\")/"
                                + modTag + "/" + secTag  + "/" + itemTag + " ";
                            rulestr += "return if ( $Standard != $Under_test ) ";
                            rulestr += " then '" + modTag + "/" + secTag  + "/" + itemTag + "'";
                            rulestr += " else ()";
                            XmlElement xe = ruleDoc.CreateElement("rule");
                            xe.InnerText = rulestr;
                            root.AppendChild(xe);
                        }
                    }
                }
                ruleDoc.AppendChild(root);
                ruleDoc.Save(rulePath);
                configDoc.Save("Config\\" + configName);

            }
            catch (IOException e)
            {
                Console.WriteLine(e.ToString());
            }
        }
    }
}
