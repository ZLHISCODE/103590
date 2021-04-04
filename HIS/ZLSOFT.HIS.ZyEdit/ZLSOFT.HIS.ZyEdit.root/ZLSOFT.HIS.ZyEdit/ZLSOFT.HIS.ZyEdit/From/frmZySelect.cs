using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ZLSOFT.HIS.ZyEdit
{
    partial class frmZySelect : Form
    {
        string mstrName = string.Empty;//窗体名称
        DataTable mdtTmp = null;//显示数据源

        public DataRow dataRowOut = null;//返回dataRow
        public frmZySelect(string strName, DataTable dtTmp)
        {
            InitializeComponent();
            mstrName = strName;
            mdtTmp = dtTmp;
        }

        private void IntdgvList(DataGridView dgv)
        {
            ComLib.DoubleBuffered(dgv, true);
            //DataGridView控件基础设置
            // 表格上下左右自适应
            dgv.Anchor = (AnchorStyles.Top | AnchorStyles.Right | AnchorStyles.Bottom | AnchorStyles.Left);
            dgv.AllowUserToAddRows = false;
            dgv.AllowUserToDeleteRows = false;
            dgv.ReadOnly = true;
            // 列手工排序

            dgv.AllowUserToOrderColumns = false;
            // 列头系统样式，设置为false，自定义才生效

            dgv.EnableHeadersVisualStyles = false;
            // 列头高度大小模式
            dgv.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            // 列头高度大小
            dgv.ColumnHeadersHeight = 25;
            // 列头居中
            dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            dgv.ColumnHeadersDefaultCellStyle.Font = new Font("宋体", 9);

            dgv.RowsDefaultCellStyle.Font = new Font("宋体", 9);

            // 列头边框样式
            dgv.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single;
            // 列头背景色

            dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(235, 236, 239);
            // 列头前景色

            dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.Black;
            // 网格线颜色

            dgv.GridColor = Color.FromArgb(245, 245, 247);
            // 背景色

            dgv.BackgroundColor = ColorTranslator.FromHtml("#FFFFFF");
            // 行头边框样式
            dgv.RowHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single;
            // 行头背景色

            dgv.RowHeadersDefaultCellStyle.BackColor = ColorTranslator.FromHtml("#FFFFFF");
            // 行高（要在窗体初始化的地方InitializeComponent调用才生效）
            dgv.RowTemplate.Height = 25;
            // 单元格内容居中

            dgv.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            // 单元格背景色
            dgv.DefaultCellStyle.BackColor = ColorTranslator.FromHtml("#FFFFFF");
            // 隔行背景色

            dgv.AlternatingRowsDefaultCellStyle.BackColor = ColorTranslator.FromHtml("#FFFFFF");

            dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect;//设置为整行


            dgv.DefaultCellStyle.SelectionBackColor = ColorTranslator.FromHtml("#66CCFF");//设置背景色

            dgv.DefaultCellStyle.SelectionForeColor = Color.Black;//设置前景色


            dgv.MultiSelect = false;

            dgv.RowHeadersVisible = false;
            //选中第一行

            if (dgvSelect.Rows.Count > 0)
            {
                dgvSelect.Rows[0].Selected = true;
            }

            //设置隐藏列

            if (dgvSelect.Columns[""] != null)
            {
                dgvSelect.Columns[""].Visible = false;
            }
            if (dgvSelect.Columns["ID"] != null)
            {
                dgvSelect.Columns["ID"].Visible = false;
            }
            if (dgvSelect.Columns["是否保密"] != null)
            {
                dgvSelect.Columns["是否保密"].Visible = false;
            }

            if (mstrName == "方剂选择器")
            {
                if (dgvSelect.Columns["简码"] != null)
                {
                    dgvSelect.Columns["简码"].Visible = false;
                }
                if (dgvSelect.Columns["别名简码"] != null)
                {
                    dgvSelect.Columns["别名简码"].Visible = false;
                }
            }

            //列宽适应
            dgvSelect.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
        }


        private void frmZyList_Load(object sender, EventArgs e)
        {
            dgvSelect.DataSource = mdtTmp;
            IntdgvList(dgvSelect);
            dgvSelect.Refresh();
            this.Refresh();
            //计算列宽
            int intWidth = 150;
            for (int j = 1; j < dgvSelect.Columns.Count; j++)
            {
                if (dgvSelect.Columns[j].Visible)
                {
                    intWidth = intWidth + dgvSelect.Columns[j].Width;
                }
            }
            intWidth = intWidth > 800 ? 800 : intWidth;
            this.Width = intWidth;
            int xWidth = SystemInformation.PrimaryMonitorSize.Width;//获取显示器屏幕宽度

            int yHeight = SystemInformation.PrimaryMonitorSize.Height;//高度
            this.Location = new Point(xWidth / 2 - this.Width / 2, yHeight / 2 - this.Height / 2);

            this.Text = mstrName;
        }

        private void dgvSelect_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                SetSelectData(1);
                this.Close();
            }
            else if (e.KeyChar == (char)Keys.Escape)
            {
                this.Close();
                this.dataRowOut = null;
            }
        }

        private void dgvSelect_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            SetSelectData();
        }

        //选中数据进行设置
        private void SetSelectData(int intType = 0)
        {
            if (dgvSelect.Rows.Count > 0 && dgvSelect.CurrentCell != null)
            {
                if (dgvSelect.CurrentCell.RowIndex >= 0)
                {
                    int index = 0;
                    if (intType == 1)
                    {
                        index = ComLib.StrToInt(dgvSelect.Tag.ToString()); //获取选中行的行号
                    }
                    else
                    {
                        index = dgvSelect.CurrentRow.Index; //获取选中行的行号
                    }
                    index = index < 0 ? 0 : index;
                    dataRowOut = (dgvSelect.Rows[index].DataBoundItem as DataRowView).Row;
                    this.Close();
                }
            }
        }

        private void dgvSelect_KeyDown(object sender, KeyEventArgs e)
        {
            if (dgvSelect.CurrentRow != null)
            {
                dgvSelect.Tag = dgvSelect.CurrentRow.Index;
            }
        }
    }
}