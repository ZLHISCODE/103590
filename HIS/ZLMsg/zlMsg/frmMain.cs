using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Text;
using System.Windows.Forms;
using zlMsgBLL;

namespace zlShortMsg
{
    public partial class frmMain : Form
    {
        frmLog frmLog = null;
        frmFailure frmFailure = null;
        frmMsgSet frmMsgSet = null;

        public frmMain()
        {
            InitializeComponent();
            menuTop.Renderer = new MenuItemRenderer();
            lblLog.Click += pnlLog_Click;
            picLog.Click += pnlLog_Click;
            lblFailure.Click += pnlFailure_Click;
            picFailure.Click += pnlFailure_Click;
            lblSetting.Click += pnlSetting_Click;
            picSetting.Click += pnlSetting_Click;
        }

        private static GraphicsPath DrawRoundRect(int x, int y, int width, int height, int radius)
        {
            //四边圆角
            GraphicsPath gp = new GraphicsPath();
            gp.AddArc(x, y, radius, radius, 180, 90);
            gp.AddArc(width - radius, y, radius, radius, 270, 90);
            gp.AddArc(width - radius, height - radius, radius, radius, 0, 90);
            gp.AddArc(x, height - radius, radius, radius, 90, 90);
            gp.CloseAllFigures();
            return gp;
        }

        private void pnlNavi_Paint(object sender, PaintEventArgs e)
        {
            Graphics g = e.Graphics; g.SmoothingMode = SmoothingMode.AntiAlias;

            Rectangle rectangle = new Rectangle(0, 0, pnlNavi.Width, pnlNavi.Height);
            LinearGradientBrush b = new LinearGradientBrush(rectangle,ColorTranslator.FromHtml("#EBEBEB"), ColorTranslator.FromHtml("#ECECEC"),(float)10);
            GraphicsPath gp= DrawRoundRect(0, 0, pnlNavi.Width - 3, pnlNavi.Height-3 , 5);
            
            g.DrawPath(Pens.Black, gp);
            g.FillPath(b, gp);
        }

        private void pnlTip_Paint(object sender, PaintEventArgs e)
        {
            Graphics g = e.Graphics;g.SmoothingMode = SmoothingMode.AntiAlias;

            Rectangle rectangle = new Rectangle(0, 0, pnlTip.Width, pnlTip.Height);
            LinearGradientBrush b = new LinearGradientBrush(rectangle, ColorTranslator.FromHtml("#D3D3D3"), ColorTranslator.FromHtml("#D8D8D8"), (float)10);
            GraphicsPath gp = DrawRoundRect(0, 0, pnlTip.Width-3 , pnlTip.Height , 5);

            g.DrawPath(Pens.Black, gp);
            g.FillPath(b, gp);
        }
        private void pnlBoder_Paint(object sender, PaintEventArgs e)
        {
            Graphics g = e.Graphics; g.SmoothingMode = SmoothingMode.AntiAlias;

            Rectangle rectangle = new Rectangle(0, 0, pnlBoder.Width, pnlBoder.Height);
            LinearGradientBrush b = new LinearGradientBrush(rectangle, ColorTranslator.FromHtml("#D3D3D3"), ColorTranslator.FromHtml("#D8D8D8"), (float)10);
            GraphicsPath gp = DrawRoundRect(0, 0, pnlBoder.Width - 3, pnlBoder.Height - 3, 5);

            g.DrawPath(Pens.Black, gp);
            g.FillPath(b, gp);
        }


        private void frmMain_Load(object sender, EventArgs e)
        {
            staTime.Text = DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss");
            SetHot(1);

            frmLog = new frmLog();
            frmLog.TopLevel = false;
            frmLog.Dock = DockStyle.Fill;
            
            pnlContainer.Controls.Add(frmLog);

            frmLog.Show();
            frmLog.StartDcn();
        }

        private void Timer_Tick(object sender, EventArgs e)
        {
            staTime.Text = DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss");
        }

        private void pnlLog_Click(object sender, EventArgs e)
        {
            if (frmLog.Visible) return;

            frmLog.Visible = true ;
            if (frmFailure != null)  frmFailure.Visible =  false;
            if (frmMsgSet != null)  frmMsgSet.Visible = false;

            SetHot(1);
        }

        private void pnlFailure_Click(object sender, EventArgs e)
        {
            if (frmFailure == null)
            {
                frmFailure = new frmFailure();
                frmFailure.TopLevel = false;
                frmFailure.Dock = DockStyle.Fill;
                pnlContainer.Controls.Add(frmFailure);
                frmFailure.Show();
            }

            frmLog.Visible = false;
            frmFailure.Visible = true;
            if (frmMsgSet != null) frmMsgSet.Visible = false;

            SetHot(2);
        }

        private void pnlSetting_Click(object sender, EventArgs e)
        {
            if (frmMsgSet == null)
            {
                frmMsgSet = new frmMsgSet();
                frmMsgSet.TopLevel = false;
                frmMsgSet.Dock = DockStyle.Fill;
                pnlContainer.Controls.Add(frmMsgSet);
                frmMsgSet.Show();  
            }

            frmLog.Visible = false;
            frmMsgSet.Visible = true;
            if (frmFailure != null) frmFailure.Visible = false;

            SetHot(3);
        }

        private void 开启短信服务ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (开启短信服务ToolStripMenuItem.Text == "开启短信服务")
            {
                开启短信服务ToolStripMenuItem.Text = "关闭短信服务";
                frmLog.StartDcn();
                staState.Text = "短信服务状态:开启";
            }
            else
            {
                开启短信服务ToolStripMenuItem.Text = "开启短信服务";
                frmLog.EndDcn();
                staState.Text = "短信服务状态:关闭";
            }
        }

        private void notifyIcon_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (this.Visible)
            {
                this.Hide();
            }
            else
            {
                this.Visible = true;
                this.WindowState = FormWindowState.Normal;
                this.Activate();
            }
        }

        private void 最小化ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Hide();
        }

        private void 退出ToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            notifyIcon.Dispose();
            Environment.Exit(0);
        }

        private void frmMain_FormClosing(object sender, FormClosingEventArgs e)
        {
            notifyIcon.Dispose();
            Environment.Exit(0);
        }

        private void frmMain_Resize(object sender, EventArgs e)
        {
            this.Refresh();
        }

        /// <summary>
        /// 设置选中菜单热图效果
        /// </summary>
        /// <param name="intActice">选中菜单:1=日志 2=错误 3=设置</param>
        private void SetHot(int intActice)
        {
            FontFamily family= new FontFamily("宋体");
            Font fBold = new Font(family,(float)10.5,FontStyle.Bold);
            Font fNormal= new Font(family, (float)10.5);

            switch (intActice)
            {
                case 1:
                    //设置图片
                    picLog.Image = Properties.Resources.消息hot;
                    picFailure.Image = Properties.Resources.错误;
                    picSetting.Image = Properties.Resources.设置;

                    //字体
                    lblLog.Font = fBold;
                    lblFailure.Font = fNormal;
                    lblSetting.Font = fNormal;

                    break;
                case 2:
                    picLog.Image = Properties.Resources.消息;
                    picFailure.Image = Properties.Resources.错误hot;
                    picSetting.Image = Properties.Resources.设置;

                    lblLog.Font = fNormal;
                    lblFailure.Font = fBold;
                    lblSetting.Font = fNormal;

                    break;
                case 3:
                    picLog.Image = Properties.Resources.消息;
                    picFailure.Image = Properties.Resources.错误;
                    picSetting.Image = Properties.Resources.设置hot;

                    lblLog.Font = fNormal;
                    lblFailure.Font = fNormal;
                    lblSetting.Font = fBold;

                    break;
            }
        }

    }
}
