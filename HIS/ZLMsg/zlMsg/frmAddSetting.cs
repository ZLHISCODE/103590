using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using zlMsgBLL;
using zlMsgModel;

namespace zlShortMsg
{
    public partial class frmAddSetting : Form
    {
        private MsgProvider msgProvider = null;
        private MsgTemplate msgTemplate = null;

        private MsgProviderBll providerBll = new MsgProviderBll();
        private MsgTemplateBLL templateBLL = new MsgTemplateBLL();
        private List<ParaInput> paraInputList= new List<ParaInput>();   //自定义控件列表


        public frmAddSetting()
        {
            InitializeComponent();
        }


        /// <summary>
        /// 新增
        /// </summary>
        /// <returns></returns>
        public new MsgTemplate ShowDialog()
        {
            this.Text = "新增";
            base.ShowDialog();
            return msgTemplate;
        }


        /// <summary>
        /// 传入对象,作为修改
        /// </summary>
        /// <param name="template"></param>
        public MsgTemplate ShowDialog(MsgTemplate template)
        {
            this.Text = "修改";

            msgTemplate = template;
            msgProvider = template.Provider;

            cbxPlat.SelectedIndexChanged -= cbxPlat_SelectedIndexChanged;       //现将事件取消绑定,避免多次加载数据
            cbxPlat.Text = template.Provider.PrividerName;
            cbxPlat.SelectedIndexChanged += cbxPlat_SelectedIndexChanged;

            txtKey.Text = template.Provider.AppKey;
            txtSecret.Text = ZLSM4.Sm4DecryptEcb(template.Provider.AppSecret);
            txtCode.Text = template.TemplateCode;
            txtSign.Text = template.SignCode;
            txtText.Text = template.TemplateText;
            txtKind.Text = template.TemplateKind;

            string [] arrPara = template.TemplatePara.Split(new string[] { "||" }, StringSplitOptions.RemoveEmptyEntries);

            if (arrPara.Length== 0)
            {
                AddParaInput();
            }

            foreach (string s in arrPara)
            {
                AddParaInput();
                paraInputList[paraInputList.Count - 1].SetValue(s);
            }

            base.ShowDialog();
            return msgTemplate;
        }

        private void frmAddSetting_Load(object sender, EventArgs e)
        {
            if (msgTemplate == null)    //对象为空,说明是新增界面
            {
                AddParaInput();
            }

            //绑定全选事件
            foreach (var con in grpPlat.Controls)
            {
                if (con is TextBox t)
                {
                    t.Enter += EnterSelectAll;
                }
            }

            foreach (var con in grpTemp.Controls)
            {
                if (con is TextBox t)
                {
                    t.Enter += EnterSelectAll;
                }
            }

        }

        private void AddControl(object sender, EventArgs e)
        {
            AddParaInput();
        }

        private void DeleteControl(object sender, EventArgs e)
        {
            DeletePataInput();
        }

        /// <summary>
        /// 向窗体中添加自定义控件
        /// </summary>
        private void AddParaInput()
        {
            //首先将上一个控件的按钮可用性设为无
            if (paraInputList.Count > 0)
            {
                ParaInput paraInputLast = (ParaInput)paraInputList[paraInputList.Count - 1];
                paraInputLast.SetBtnEnable(false);
            }

            ParaInput paraInput = new ParaInput(paraInputList.Count == 0);
            paraInput.Size = new Size(264, 25);
            paraInput.Location = new Point(18, 180+ (paraInputList.Count+1 )*3 + paraInputList.Count * paraInput.Height );

            //事件bind
            paraInput.AddClick += AddControl;   
            paraInput.DeleteClick += DeleteControl;

            paraInputList.Add(paraInput);

            //添加控件后,调整容器的大小
            grpTemp.Controls.Add(paraInput);
            grpTemp.Height = grpTemp.Height  + 3+ paraInput.Height;
            this.Height = this.Height + 3 + paraInput.Height;
        }

        /// <summary>
        /// 删除窗体中的自定义控件
        /// </summary>
        private void DeletePataInput()
        {
            //点击删除的,肯定是最后一个按钮
            ParaInput paraInput = (ParaInput)paraInputList[paraInputList.Count - 1];
            grpTemp.Controls.Remove(paraInput);
            paraInputList.Remove(paraInput);

            //改变上一个控件的按钮可用性
            paraInput = (ParaInput)paraInputList[paraInputList.Count - 1];
            paraInput.SetBtnEnable(true, paraInputList.Count == 1);

            //删除控件后,调整容器大小
            grpTemp.Height = grpTemp.Height - 3 - paraInput.Height;
            this.Height = this.Height - 3 - paraInput.Height;
        }

        private void cbxPlat_SelectedIndexChanged(object sender, EventArgs e)
        {
            msgProvider = new MsgProviderBll().GetMsgProviderByName(cbxPlat.Text);

            if (msgProvider == null)
            {
                txtKey.Text = "";
                txtSecret.Text = "";
            }
            else
            {
                txtKey.Text = msgProvider.AppKey;
                txtSecret.Text = msgProvider.AppSecret;
            }
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            if (msgProvider == null )
            {
                msgProvider = new MsgProvider(providerBll.GetProviderCode(), cbxPlat.Text, txtKey.Text,ZLSM4.Sm4EncryptEcb(txtSecret.Text));
                providerBll.ProviderInsert(msgProvider);
            }
            else if (msgProvider.AppKey != txtKey.Text || msgProvider.AppSecret != ZLSM4.Sm4EncryptEcb(txtSecret.Text))
            {
                msgProvider.AppKey = txtKey.Text;
                msgProvider.AppSecret =ZLSM4.Sm4EncryptEcb(txtSecret.Text);
                providerBll.ProviderUpdate(msgProvider);
            }



            string para = "";
            foreach (ParaInput p  in  paraInputList)
            {
                if (para == "")
                    para = p.GetValue();
                else
                    para = para + "||" + p.GetValue();
            }

            if (msgTemplate == null)    //全局对象为空,说明是新增
            {
                msgTemplate = new MsgTemplate(templateBLL.GetTemplateID(), msgProvider, txtCode.Text, txtSign.Text,"",para, txtKind.Text, txtText.Text);
                if (templateBLL.TemplateInsert(msgTemplate))
                {
                    MessageBox.Show("添加成功！");
                    Close();
                }
            }
            else
            {
                msgTemplate = new MsgTemplate(msgTemplate.TemplateID, msgProvider, txtCode.Text, txtSign.Text, "",para, txtKind.Text, txtText.Text);
                if (templateBLL.TemplateUpdate(msgTemplate))
                {
                    MessageBox.Show("修改成功！");
                    Close();
                }
            }
        }

        private void EnterSelectAll(object sender, EventArgs e)
        {
            BeginInvoke((Action)delegate
            {
                (sender as TextBox).SelectAll();
            });
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
