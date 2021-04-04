using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace zlShortMsg
{
    public static class DgvDrawer
    {   
        /// <summary>
        /// 初始化表头
        /// </summary>
        /// <param name="gridView"></param>
        /// <param name="strHeader">表头信息 格式为:  列1:是否可见;列2:是否可见</param>
        public static void ChangeDgvStyle(ref DataGridView gridView,string strHeader = "")
        {
            int j = 0;
            gridView.AllowUserToAddRows = false;
            //初始化表头数据
            if (strHeader != "")
            {
                DataGridViewTextBoxColumn gridViewColumn;

                string[] arrHeader = strHeader.Split(';');
                string[] arrSplit;
                for(int i=0;i<arrHeader.Length;i++)
                {
                    arrSplit = arrHeader[i].Split(':');
                    gridViewColumn = new DataGridViewTextBoxColumn();
                    gridViewColumn.HeaderText = arrSplit[0];
                    gridViewColumn.Name = arrSplit[1];
                    if (arrSplit[2] == "1")
                    {
                        gridViewColumn.Visible = true;
                        j = i;  //记录可现实列的列号,用于设置最后一列填充整个控件
                    }
                    else
                    {
                        gridViewColumn.Visible = false;
                    }
                    gridViewColumn.ReadOnly = true;
                    gridView.Columns.Add(gridViewColumn);
                }

            }

            //Header样式设置
            gridView.EnableHeadersVisualStyles = false;

            gridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;
            gridView.Columns[j + 1].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            //表格颜色
            gridView.AlternatingRowsDefaultCellStyle.BackColor =SystemColors.InactiveBorder;
        }

    }
}
