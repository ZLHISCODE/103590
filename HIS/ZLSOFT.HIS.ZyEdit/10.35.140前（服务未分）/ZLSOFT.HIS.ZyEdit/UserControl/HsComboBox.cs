using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Threading;

namespace ZLSOFT.HIS.ZyEdit
{
     partial class HsComboBox : ComboBox
    {
        //用于模拟键盘输入
        [DllImport("user32.dll")]
        private static extern void keybd_event(
            byte bVk, byte bScan, uint dwFlags, int dwExtraInfo);

        //显示控件
        //Control Control;
        //下拉面板
        Panel panel;

        //绘制面板用变量
        //光标前一位置
        Point pPoint;
        //光标当前位置
        Point cPoint;
        //鼠标是否已按下
        bool isMouseDown = false;

        //关闭下拉时光标是否在ComboBox上
        bool isCursorOnComboBox = false;
        
        [Browsable(false)]
        public Control Control { get; set; }
        [Browsable(false)]
        public CheckedListBox CheckedListBox { get; set; }
        [Browsable(false)]
        public TreeView TreeView { get; set; }

        public DataGridView DataGridView { get; set; }

        public DataGridView GridViewTmp { get; set; }

        public HsToolStripDropDown DropDown { get; set; }


        //控件类型
        public enum TypeC
        {
            //默认普通ComboBox
            ComboBox = 0,
            CheckedListBox = 1,
            TreeView = 2,
            DataGridView = 3,
            //普通Control
            Control = 99
        }

        /// <summary>
        /// 单位是否可见
        /// </summary>
        private TypeC _CtlType;
        [Description("设置下拉类型。")]
        [DefaultValue(0)]
        public TypeC CtlType
        {
            get
            {
                return _CtlType;
            }
            set
            {
                _CtlType = value;
                switch (_CtlType)
                {
                    case TypeC.ComboBox:
                        break;
                    case TypeC.CheckedListBox:
                        CheckedListBox CheckedListBox = new CheckedListBox();
                        this.SetDropDown(CheckedListBox);
                        break;
                    case TypeC.TreeView:
                        TreeView TreeView = new TreeView();
                        this.SetDropDown(TreeView);
                        break;
                    case TypeC.DataGridView:
                        DataGridView DataGridView = new DataGridView();
                        this.SetDropDown(DataGridView);
                        break;
                    default:
                        Control Control = new Panel();
                        this.SetDropDown(Control);
                        break;
                }
            }
        }


        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == Keys.Enter && this.DataGridView.Focused == true)
            {
                SendKeys.Send("{tab}");
                return true;
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }


        public HsComboBox()
        {
            InitializeComponent();
        }

        public HsComboBox(IContainer container)
        {
            container.Add(this);

            InitializeComponent();


            //设置下拉样式为DropDownList，不能手动输入
            this.DropDownStyle = ComboBoxStyle.DropDown;
            //绘制下拉面板
            this.DrawPanel();
        }

        /// <summary>
        /// 点击事件
        /// </summary>
        protected override void  OnMouseClick(MouseEventArgs e)
        {
            //如果为Null则为原生ComboBox
            if (this._CtlType != TypeC.ComboBox)
            {
                if (isCursorOnComboBox)
                {
                    isCursorOnComboBox = false;
                    //模拟Enter键，取消掉下拉状态
                    keybd_event(0xD, 0, 0, 0);
                    keybd_event(0xD, 0, 0x0002, 0);
                }
                else
                {
                    //创建下拉窗
                    ToolStripControlHost toolStripControlHost = new ToolStripControlHost(this.panel);
                    HsToolStripDropDown toolStripDropDown = new HsToolStripDropDown();
                    //设置边框
                    toolStripControlHost.Margin = Padding.Empty;
                    toolStripControlHost.Padding = Padding.Empty;
                    toolStripControlHost.AutoSize = false;
                    toolStripDropDown.Padding = Padding.Empty;
                    //添加
                    toolStripDropDown.Items.Add(toolStripControlHost);
                    toolStripDropDown.Show(this, 0, this.Height);
                    //设置宽度最小值
                    if (this.panel.Width < this.Width)
                    {
                        this.panel.Size = new System.Drawing.Size(this.Width, this.panel.Height);
                    }
                    //判断关闭时光标在ComboBox组件内
                    toolStripDropDown.Closed += delegate(object sender, ToolStripDropDownClosedEventArgs e1)
                    {
                        Rectangle rec = new Rectangle(0, 0, this.Width, this.Height);
                        this.isCursorOnComboBox = rec.Contains(this.PointToClient(Cursor.Position));
                    };
                    //设置焦点
                    toolStripDropDown.Focus();
                    DropDown = toolStripDropDown;
                }
            }
        }

        /// <summary>
        /// 绘制下拉面板
        /// </summary>
        public void DrawPanel()
        {
            this.panel = new Panel();
            this.panel.Size = new System.Drawing.Size(this.Width, 100);
            this.panel.Padding = new Padding(1, 1, 1, 13);
            this.panel.BackColor = Color.FromArgb(245, 245, 247);
            //绘制边线
            this.panel.Paint += delegate(object sender, PaintEventArgs e)
            {
                ControlPaint.DrawBorder(e.Graphics,
                               this.panel.ClientRectangle,
                               Color.DarkGray,
                               1,
                               ButtonBorderStyle.Solid,
                               Color.DarkGray,
                               1,
                               ButtonBorderStyle.Solid,
                               Color.DarkGray,
                               1,
                               ButtonBorderStyle.Solid,
                               Color.DarkGray,
                               1,
                               ButtonBorderStyle.Solid);
            };
            //使用Label实现右下角拖动按钮
            picDown.Parent = this.panel;
            picDown.Anchor = (AnchorStyles.Bottom | AnchorStyles.Right);
            picDown.Location = new Point(this.panel.Location.X + this.panel.Size.Width - picDown.Width,
                this.panel.Location.Y + this.panel.Size.Height - picDown.Height);
            //实现缩放功能
            picDown.MouseDown += delegate(object sender, MouseEventArgs e1)
            {
                this.pPoint = Cursor.Position;
                this.isMouseDown = true;
            };
            picDown.MouseLeave += delegate(object sender, EventArgs e1)
            {
                this.isMouseDown = false;
            };
            picDown.MouseMove += delegate(object sender, MouseEventArgs e1)
            {
                this.cPoint = Cursor.Position;
                if (e1.Button == MouseButtons.Left && isMouseDown )
                {
                    if  (Math.Max(this.panel.Height + cPoint.Y - pPoint.Y, 23) > 0 && Math.Max(this.panel.Width + cPoint.X - pPoint.X, this.Width) > 0){
                        this.panel.Height = Math.Max(this.panel.Height + cPoint.Y - pPoint.Y, 23);
                        this.panel.Width = Math.Max(this.panel.Width + cPoint.X - pPoint.X, this.Width);
                        pPoint = Cursor.Position;

                    }

                }
                else
                {
                    picDown.Cursor = Cursors.SizeNWSE;
                }
            };
        }

        /// <summary>
        /// 设置下拉内容-普通Control
        /// </summary>
        public void SetDropDown(Control control)
        {
            //需要将此前的Control注销掉，否则会一直显示在Panel中
            if(this.Control != null)
            {
                this.Control.Dispose();
            }
            this.Control = control;
            this.Control.Location = new Point(0, 0);
            this.Control.Dock = DockStyle.Fill;
            this.Control.Parent = this.panel;
            //将下拉高度设为1,实现隐藏效果
            this.DropDownHeight = 1;
        }

        /// <summary>
        /// 设置下拉内容-普通Control
        /// </summary>
        public void SetDropDown(DataGridView DataGridView)
        {
            //需要将此前的Control注销掉，否则会一直显示在Panel中
            //if (this.DataGridView != null)
            //{
            //    this.DataGridView.Dispose();
                
            //}
            this.DataGridView = DataGridView;
            this.DataGridView.Location = new Point(0, 0);
            this.DataGridView.Dock = DockStyle.Fill;
            this.DataGridView.Parent = this.panel;

            //将下拉高度设为1,实现隐藏效果
            this.DropDownHeight = 1;
        }


        /// <summary>
        /// 设置下拉内容-复选框列表
        /// </summary>
        public void SetDropDown(CheckedListBox checkedListBox)
        {
            //单击可选中
            checkedListBox.CheckOnClick = true;
            //边框格式
            checkedListBox.BorderStyle = BorderStyle.None;
            //去掉强制的高度修改,CheckedListBox的一个特殊高度判断，如果为true则会在底端多出一段空白
            checkedListBox.IntegralHeight = false;
            //选中事件添加监听（此时还未选中，刚Check还没有Checked）
            checkedListBox.ItemCheck += delegate(object sender, ItemCheckEventArgs e)
            {
                String text = "";
                for (int i = 0; i < ((CheckedListBox)(this.Control)).Items.Count; i++ )
                {
                    //使用异或特殊处理当前正在check的条目
                    if ((i == e.Index) != ((CheckedListBox)(this.Control)).GetItemChecked(i))
                    {
                        text += ((CheckedListBox)(this.Control)).Items[i].ToString() + ",";
                    }
                }
                text = text.Substring(0,Math.Max(text.Length-1,0));
                //显示所有内容
                ShowText(text);
            };
            this.CheckedListBox = checkedListBox;
            this.SetDropDown((Control)checkedListBox);
        }

        /// <summary>
        /// 设置下拉内容-树形列表
        /// </summary>
        public void SetDropDown(TreeView treeView)
        {
            //整行选择
            treeView.FullRowSelect = true;
            //边框格式
            treeView.BorderStyle = BorderStyle.None;
            //选中事件添加监听
            treeView.AfterSelect += delegate(object sender, TreeViewEventArgs e)
            {
                ShowText(((TreeView)(this.Control)).SelectedNode.Text);
            };
            this.TreeView = treeView;
            this.SetDropDown((Control)treeView);
        }

        /// <summary>
        /// 显示信息
        /// </summary>
        /// <param name="Text">信息内容</param>
        public void ShowText(String Text)
        {
            //当DropDownStyle = DropDownList时不能直接对Text赋值
            this.Items.Clear();
            this.Items.Add(Text);
            this.SelectedIndex = 0;
        }


        /// <summary>
        /// 直接打开下拉窗
        /// </summary>
        public void ShowDiag(int intWidth,int intHeight)
        {
            panel.SuspendLayout();
            DrawPanel();
            SetDropDown(this.GridViewTmp);

            //创建下拉窗
            ToolStripControlHost toolStripControlHost = new ToolStripControlHost(this.panel);
            HsToolStripDropDown toolStripDropDown = new HsToolStripDropDown();
            //设置边框
            toolStripControlHost.Margin = Padding.Empty;
            toolStripControlHost.Padding = Padding.Empty;
            toolStripControlHost.AutoSize = false;
            toolStripDropDown.Padding = Padding.Empty;
            //添加
            toolStripDropDown.Items.Add(toolStripControlHost);
            toolStripDropDown.Show(this, 0, this.Height);
            //设置宽度
            this.panel.Size = new System.Drawing.Size(this.Width, this.panel.Height);
            this.panel.Size = new System.Drawing.Size(intWidth, this.panel.Height);
            if (this.panel.Height< intHeight)
            {
                this.panel.Size = new System.Drawing.Size(this.panel.Width, intHeight);
            }
            //判断关闭时光标在ComboBox组件内
            toolStripDropDown.Closed += delegate (object sender, ToolStripDropDownClosedEventArgs e1)
            {
                Rectangle rec = new Rectangle(0, 0, this.Width, this.Height);
                this.isCursorOnComboBox = rec.Contains(this.PointToClient(Cursor.Position));
            };

            DropDown = toolStripDropDown;
            //设置焦点
            toolStripDropDown.Focus();
        }

        //关闭弹窗
        public void CloseDiag() {
            if (DropDown != null) {
            isCursorOnComboBox = false;
            DropDown.Close();
            }
        }





    }

    /// <summary>
    /// 重写ToolStripDropDown
    /// 使用双缓存减少闪烁
    /// </summary>
    public class HsToolStripDropDown : ToolStripDropDown
    {
        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams cp = base.CreateParams;
                cp.ExStyle |= 0x02000000;//双缓存
                return cp;
            }
        }
    }
    
}
