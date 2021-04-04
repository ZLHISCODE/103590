using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using zlMsgDAL;
using zlMsgModel;

namespace zlMsgBLL
{
    public class MsgTemplateBLL
    {
        private MsgTemplateDal templateDal = new MsgTemplateDal();

        public List<MsgTemplate> GetTemplates()
        {
            return templateDal.GetTemplateLists();
        }

        public long GetTemplateID()
        {
            return templateDal.GetTemplateID();
        }

        public bool TemplateUpdate(MsgTemplate template)
        {
            return templateDal.UpdateTemplate(template);
        }

        public bool TemplateInsert(MsgTemplate template)
        {
            return templateDal.AddTemplate(template);
        }

        public bool DeleteTempLateByid(List<long> TemplateIDs)
        {
            templateDal.DeleteTemplate(TemplateIDs);
            return true;
        }

    }
}
