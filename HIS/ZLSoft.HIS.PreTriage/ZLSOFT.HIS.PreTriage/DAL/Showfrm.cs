using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;

namespace ZLSOFT.HIS.PreTriage.DAL
{
    public static class ShowForm
    {
        //选择器窗体显示

        public static bool ChooseData(DataTable dtDate, string strTxt, bool bolSearch, out string strOut)
        {
            //dtDate传入的数据,strTxt文本框内容,bolSearch是否查找，strOut用分号分隔的字符串

            try
            {
                if (dtDate.Columns["选择"] != null)
                {
                    dtDate.Columns.Remove(dtDate.Columns["选择"]);
                }
                frmcc frm = new frmcc(dtDate, strTxt, bolSearch);
                frm.ShowDialog();
                strOut = frm.gstrOut;

                if (dtDate.Columns["选择"] != null)
                {
                    dtDate.Columns.Remove(dtDate.Columns["选择"]);
                }
                return frm.gbolOK;
            }
            catch
            {
                strOut = "";
                return false;
            }
        }
        //选择器窗体显示

        public static bool ChooseRuleData(DataTable dtDate, string strTxt, bool bolSearch, out string strOut)
        {
            //dtDate传入的数据,strTxt文本框内容,bolSearch是否查找，strOut用分号分隔的字符串

            try
            {
                if (dtDate.Columns["选择"] != null) {
                    dtDate.Columns.Remove(dtDate.Columns["选择"]);
                }
                frmRules frm = new frmRules(dtDate, strTxt, bolSearch);
                frm.ShowDialog();
                strOut = frm.gstrOut;

                if (dtDate.Columns["选择"] != null)
                {
                    dtDate.Columns.Remove(dtDate.Columns["选择"]);
                }
                return frm.gbolOK;

        }
            catch
            {
                strOut = "";
                return false;
            }
}
    }
}
