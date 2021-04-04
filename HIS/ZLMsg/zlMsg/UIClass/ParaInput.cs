using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Text;
using System.Windows.Forms;

namespace zlShortMsg
{
    [ToolboxItem(true)]
    public partial class ParaInput : UserControl
    {
        public delegate void AddEvent(object sender, EventArgs e);
        public delegate void DeleteEvent(object sender, EventArgs e);

        public event AddEvent AddClick;
        public event DeleteEvent DeleteClick;

        public ParaInput()
        {
            InitializeComponent();
        }

        public ParaInput(bool blnFirst)
        {
            InitializeComponent();
            if (blnFirst)
            {
                lblDelete.Enabled = false;
                lblDelete.Visible = false;
            }
        }

        public string GetValue()
        {
            return txtPara.Text;
        }

        public void SetValue(string strValue)
        {
            txtPara.Text = strValue;
        }

        /// <summary>
        /// 功能:设置功能按钮是否可用
        /// </summary>
        public void SetBtnEnable(bool blnEnable,bool blnFirst= false)
        {
            lblAdd.Enabled = blnEnable;
            lblAdd.Visible = blnEnable;

            lblDelete.Enabled = blnEnable && !blnFirst;
            lblDelete.Visible = blnEnable && !blnFirst; ;
        }

        private void lblAdd_Click(object sender, EventArgs e)
        {
            AddClick?.Invoke(sender, e);
        }

        private void lblDelete_Click(object sender, EventArgs e)
        {
            DeleteClick?.Invoke(sender, e);
        }

        private void txtPara_Enter(object sender, EventArgs e)
        {
            BeginInvoke((Action)delegate
            {
                (sender as TextBox).SelectAll();
            });
        }
    }
}
