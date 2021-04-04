using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DevExpress.XtraEditors.Controls;
using System.Runtime.InteropServices;

namespace ZLSOFT.HIS.PreTriage.ComLib
{
    [ComVisible(true), ClassInterface(ClassInterfaceType.None)]
    public class MessboxClass : Localizer
    {
        public override string GetLocalizedString(DevExpress.XtraEditors.Controls.StringId id)
        {
            switch (id)
            {
                case StringId.XtraMessageBoxCancelButtonText:
                    return "取消";
                case StringId.XtraMessageBoxOkButtonText:
                    return "确定";
                case StringId.XtraMessageBoxYesButtonText:
                    return "是";
                case StringId.XtraMessageBoxNoButtonText:
                    return "否";
                default:
                    return base.GetLocalizedString(id);
            }
        }
    }
}
