
namespace zlMsgModel
{
    public class MsgTemplate
    {
        public long TemplateID { get; set; }
        public MsgProvider Provider { get; set; }
        public string TemplateCode { get; set; }
        public string SignCode { get; set; }
        public string SignNumber { get; set; }
        public string TemplatePara { get; set; }
        public string TemplateKind { get; set; }
        public string TemplateText { get; set; }

        public MsgTemplate(long templateID, MsgProvider provider, string templateCode, string signCode, string signNumber, string templatePara, string templateKind, string templateText)
        {
            TemplateID = templateID;
            Provider = provider;
            TemplateCode = templateCode;
            SignCode = signCode;
            SignNumber = signNumber;
            TemplatePara = templatePara;
            TemplateKind = templateKind;
            TemplateText = templateText;
        }
    }
}
