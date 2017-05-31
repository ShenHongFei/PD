using System.Text;

namespace PaperFormatDetection.Frame
{
    public class Module
    {
        private string className;
        private string cnName;
        private string rule;
        private bool detect;
        private bool xqueryUse;



        public Module(string className, string cnName, string rule, bool detect, bool xqueryUse)
        {
            this.className = className;
            this.cnName = cnName;
            this.rule = rule;
            this.detect = detect;
            this.xqueryUse = xqueryUse;
        }

        public string ClassName
        {
            get
            {
                return className;
            }
        }

        public string CnName
        {
            get
            {
                return cnName;
            }
        }

        public string Rule
        {
            get
            {
                return rule;
            }
        }

        public bool Detect
        {
            get
            {
                return detect;
            }
        }

        public bool XqueryUse
        {
            get
            {
                return xqueryUse;
            }
        }

        public override string ToString() {
            StringBuilder sb = new StringBuilder();
            sb.Append("{");
            sb.Append("ClassName: ");
            sb.Append(ClassName);
            sb.Append(", CnName: ");
            sb.Append(CnName);
            sb.Append(", Rule: ");
            sb.Append(Rule);
            sb.Append(", Detect: ");
            sb.Append(Detect);
            sb.Append(", XqueryUse: ");
            sb.Append(XqueryUse);
            sb.Append("}");
            return sb.ToString();
        }
    }
}
