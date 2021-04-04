using Microsoft.Win32;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using zlMsgBLL;

namespace zlShortMsg
{
    public partial class frmLogin : Form
    {
        public frmLogin()
        {
            InitializeComponent();
        }

        private void InitComponent()
        {
            lblTip.Parent = pctLeft;
            lblTip1.Parent = pctLeft;
            lblTip2.Parent = pctLeft;
            lblTip3.Parent = pctLeft;

            //设置文本框内容
            txtUser.Text = RegistryHelper.GetValue("Username") ?? "";
            cboList.Text = RegistryHelper.GetValue("Servername") ?? "";
            cboRole.SelectedIndex = (int)RegistryHelper.GetValue("Role").Val();

            if (RegistryHelper.GetValue("SavePwd") =="1")
            {
                chkPwd.Checked = true;
                string strPwd = RegistryHelper.GetValue("Userpwd") ?? "";
                if (!string.IsNullOrEmpty(strPwd))
                {
                    txtPwd.Text = ZLSM4.Sm4DecryptEcb(strPwd);
                }
                
            }

            //为文本框绑定Keypress事件
            txtUser.KeyPress += new KeyPressEventHandler((sender, args) =>
            {
                if (args.KeyChar == (char)Keys.Enter)
                {
                    SendKeys.Send("{TAB}");
                }
            });
            txtPwd.KeyPress += new KeyPressEventHandler(PressEnter);
            cboList.KeyPress += new KeyPressEventHandler(PressEnter);

            //绑定全选事件
            foreach (var con in this.Controls)
            {
                if (con is TextBox)
                {
                    TextBox t = (TextBox) con;
                    t.Enter += EnterSelectAll;
                }
            }
        }

        private void frmLogin_Load(object sender, EventArgs e)
        {
            
            InitComponent();
            LoadServer();
        }

        private void PressEnter(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                pctLogin_Click(null, null);
            }
        }

        private void picLogin_MouseEnter(object sender, EventArgs e)
        {
            pctLogin.Image = Properties.Resources.active;
        }

        private void picLogin_MouseLeave(object sender, EventArgs e)
        {
            pctLogin.Image = Properties.Resources.normal;
        }

        private void lblMin_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void lblMax_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        [DllImport("user32.dll")]
        public static extern bool ReleaseCapture();
        [DllImport("user32.dll")]
        public static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);
        private void pctForm_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                //这里一定要判断鼠标左键按下状态，否则会出现一个很奇葩的BUG，不信邪可以试一下~~
                ReleaseCapture();
                SendMessage(Handle, 0x00A1, 2, 0);
            }
        }


        /// <summary>
        /// 加载服务器列表
        /// </summary>
        private void LoadServer()
        {
            string dbPath = string.Empty;
            string strFile = string.Empty;

            dbPath = Environment.GetEnvironmentVariable("TNS_ADMIN");   //首先从环境变量中获取文件
            if (dbPath != "")
            {
                strFile = dbPath + "\\tnsnames.ora";
                if (!File.Exists(strFile))
                {
                    strFile = dbPath + "\\network\\ADMIN\\tnsnames.ora";
                    if (!File.Exists(strFile))
                    {
                        strFile = string.Empty;
                    }
                }
            }

            if (strFile == "")    //如果环境变量中没有找到文件,那么就从注册表中寻找
            {
                RegistryKey Rkey = Registry.LocalMachine.OpenSubKey("SOFTWARE\\Oracle");
                string[] arrTmp = Rkey.GetSubKeyNames();
                foreach (string strKey in arrTmp)
                {
                    if (strKey.Like("KEY_ORA*HOME*"))
                    {
                        strFile = Rkey.OpenSubKey(strKey).GetValue(@"ORACLE_HOME") + @"\network\ADMIN\tnsnames.ora";

                        if (File.Exists(strFile))
                        {
                            break;
                        }
                        else
                        {
                            strFile = string.Empty;
                        }
                    }
                }

                if (!string.IsNullOrEmpty(strFile))
                {
                    //如果注册表中可以找到文件,就复制到环境变量中
                    string strPath = strFile.Substring(0, strFile.LastIndexOf(@"\"));
                    Environment.SetEnvironmentVariable("TNS_ADMIN", strPath, EnvironmentVariableTarget.Machine);
                }
            }

            string[] output = null;
            if (strFile != "")
            {
                output = GetDatabases(strFile);
            }
            else
            {
                return;
            }
            //循环加载到下拉列表中去
            for (int i = 0; i < output.Length; i++)
            {
                if (output[i].ToString() != "")
                {
                    this.cboList.Items.Add(output[i].ToString());
                }
            }
        }

        /// <summary>
        /// 读取得数据库列表
        /// </summary>
        /// <param name="oraFile">oracle的ora文件地址</param>
        /// <returns>数据库列表字符数组</returns>
        private  string[] GetDatabases(string oraFile)
        {
            #region 读取TNS文件
            string output = "";
            string fileLine;
            System.Collections.Stack parens = new Stack();
            StreamReader sr;
            try
            {
                sr = new StreamReader(@"" + oraFile);
            }
            catch (System.IO.FileNotFoundException ex)
            {
                throw ex;
            }
            //  读取文件的第一行   
            fileLine = sr.ReadLine();
            #endregion
            try
            {
                #region 循环，读取每一行

                while (fileLine != null)
                {
                    if (fileLine.Trim().Length > 0)
                    {
                        if (fileLine.Trim().Substring(0, 1) != "#")
                        {
                            char lineChar;
                            for (int i = 0; i < fileLine.Length; i++)
                            {
                                lineChar = fileLine[i];

                                if (lineChar == '(')
                                {
                                    //如果第一个字符是 "(" 整行放入 堆栈。   
                                    parens.Push(lineChar);
                                }
                                else if (lineChar == ')')
                                {
                                    //  如果字符是")",一个一个移出  
                                    parens.Pop();
                                }
                                else
                                {
                                    if (parens.Count == 0)
                                    {
                                        output += lineChar;
                                    }
                                }
                            }
                        }
                    }
                    // 读取文件下一行 
                    fileLine = sr.ReadLine();
                }
                sr.Close();
                #endregion
                #region 处理=号
                string[] split = output.Split('=');

                //  以"="号为分隔符。截掉，放入split内
                for (int i = 0; i < split.Length; i++)
                {
                    split[i] = split[i].Trim();
                }
                Array.Sort(split);
                return split;
                #endregion
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        private bool Login(string strUser,string strPwd,string strServer,string strRole)
        {
            //创建连接
            try
            {
                OracleConnect.GetConn(strUser, strPwd, strServer, strRole);
            }
            catch (Exception e)
            {
                MessageBox.Show("登录发生错误: " + e.Message, "错误");
                return false;
            }
            return true;
        }

        private void pctLogin_Click(object sender, EventArgs e)
        {
            if (!Login(txtUser.Text, txtPwd.Text, cboList.Text, cboRole.Text)) return;

            RegistryHelper.SetValue("Username", txtUser.Text);
            RegistryHelper.SetValue("Servername", cboList.Text);
            RegistryHelper.SetValue("Role", cboRole.SelectedIndex.ToString());

            if (chkPwd.Checked)
            {
                RegistryHelper.SetValue("Userpwd", ZLSM4.Sm4EncryptEcb(txtPwd.Text));
                RegistryHelper.SetValue("SavePwd", "1");
            }
            else
            {
                RegistryHelper.SetValue("SavePwd", "0");
            }

            this.Hide();   new frmMain().Show(); 
        }

        private void EnterSelectAll(object sender, EventArgs e)
        {
            BeginInvoke((Action)delegate
            {
                (sender as TextBox).SelectAll();
            });
        }
    }
}
