VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmIdentifyNBYKT 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "宁波一卡通身份识别"
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8445
   Icon            =   "frmIdentifyNBYKT.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   8445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.OptionButton opt 
      Caption         =   "异地农保卡"
      Height          =   225
      Index           =   6
      Left            =   6900
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   210
      Width           =   1275
   End
   Begin VB.OptionButton opt 
      Caption         =   "异地医保卡"
      Height          =   225
      Index           =   5
      Left            =   5610
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   210
      Width           =   1215
   End
   Begin VB.OptionButton opt 
      Caption         =   "农保卡"
      Height          =   225
      Index           =   4
      Left            =   4710
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   210
      Width           =   885
   End
   Begin VB.OptionButton opt 
      Caption         =   "医保卡"
      Height          =   225
      Index           =   3
      Left            =   3750
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   210
      Width           =   885
   End
   Begin VB.CheckBox chk卡类型 
      Caption         =   "新卡"
      Height          =   225
      Left            =   1140
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   210
      Width           =   675
   End
   Begin VB.CommandButton cmdCard 
      Caption         =   "确定(&O)"
      Height          =   350
      Index           =   0
      Left            =   5730
      TabIndex        =   64
      Top             =   6600
      Width           =   1100
   End
   Begin VB.CommandButton cmdCard 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Index           =   1
      Left            =   6990
      TabIndex        =   65
      Top             =   6600
      Width           =   1100
   End
   Begin TabDlg.SSTab sstab 
      Height          =   4425
      Left            =   240
      TabIndex        =   10
      Top             =   2010
      Width           =   7995
      _ExtentX        =   14102
      _ExtentY        =   7805
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      Enabled         =   0   'False
      TabCaption(0)   =   "病人信息"
      TabPicture(0)   =   "frmIdentifyNBYKT.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl姓名"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbl性别"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lbl出生日期"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lbl证件类型"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lbl证件号"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lbl婚姻状况"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblEMAIL"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lbl说明"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lbl档案号"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lbl省"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lbl区"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lbl街道"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "lbl地址"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "lbl邮编"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "lbl电话"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "lbl工作单位"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "lbl单位地址"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "lbl单位邮编"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "lbl单位电话"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "lbl职业"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "lbl手机号"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "lbl家属姓名"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "lbl家属电话"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "lbl卡类型"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "lbl卡号"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "lbl保险号"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "txt姓名"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "cbo性别"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "msk出生日期"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "cbo证件类型"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "txt证件号"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "cbo婚姻状况"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "txtEMAIL"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "txt说明"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "txt档案号"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "Frame1"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "txt省"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "txt区"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "txt街道"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "txt地址"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "txt邮编"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "txt电话"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "txt工作单位"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "txt单位地址"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "txt单位邮编"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "txt单位电话"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "cbo职业"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "txt手机号"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "txt家属姓名"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "txt家属电话"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "txt卡号码"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "txt保险号"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "cbo卡类型"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).ControlCount=   53
      Begin VB.ComboBox cbo卡类型 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1020
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   1710
         Width           =   1515
      End
      Begin VB.TextBox txt保险号 
         Enabled         =   0   'False
         Height          =   300
         Left            =   6330
         MaxLength       =   50
         TabIndex        =   34
         Top             =   1710
         Width           =   1515
      End
      Begin VB.TextBox txt卡号码 
         Enabled         =   0   'False
         Height          =   300
         Left            =   3600
         MaxLength       =   50
         TabIndex        =   32
         Top             =   1710
         Width           =   1515
      End
      Begin VB.TextBox txt家属电话 
         Height          =   300
         Left            =   6330
         MaxLength       =   50
         TabIndex        =   63
         Top             =   3930
         Width           =   1485
      End
      Begin VB.TextBox txt家属姓名 
         Height          =   300
         Left            =   3600
         MaxLength       =   50
         TabIndex        =   61
         Top             =   3930
         Width           =   1485
      End
      Begin VB.TextBox txt手机号 
         Height          =   300
         Left            =   1020
         MaxLength       =   50
         TabIndex        =   59
         Top             =   3930
         Width           =   1485
      End
      Begin VB.ComboBox cbo职业 
         Height          =   300
         Left            =   6330
         TabIndex        =   57
         Text            =   "cbo职业"
         Top             =   3540
         Width           =   1515
      End
      Begin VB.TextBox txt单位电话 
         Height          =   300
         Left            =   3600
         MaxLength       =   50
         TabIndex        =   55
         Top             =   3540
         Width           =   1485
      End
      Begin VB.TextBox txt单位邮编 
         Height          =   300
         Left            =   1020
         MaxLength       =   50
         TabIndex        =   53
         Top             =   3540
         Width           =   1485
      End
      Begin VB.TextBox txt单位地址 
         Height          =   300
         Left            =   4770
         MaxLength       =   50
         TabIndex        =   51
         Top             =   3150
         Width           =   3045
      End
      Begin VB.TextBox txt工作单位 
         Height          =   300
         Left            =   1020
         MaxLength       =   50
         TabIndex        =   49
         Top             =   3150
         Width           =   2295
      End
      Begin VB.TextBox txt电话 
         Height          =   300
         Left            =   6330
         MaxLength       =   50
         TabIndex        =   47
         Top             =   2760
         Width           =   1485
      End
      Begin VB.TextBox txt邮编 
         Height          =   300
         Left            =   3990
         MaxLength       =   50
         TabIndex        =   45
         Top             =   2760
         Width           =   1095
      End
      Begin VB.TextBox txt地址 
         Height          =   300
         Left            =   1020
         MaxLength       =   50
         TabIndex        =   43
         Top             =   2760
         Width           =   2295
      End
      Begin VB.TextBox txt街道 
         Height          =   300
         Left            =   6330
         MaxLength       =   50
         TabIndex        =   41
         Top             =   2370
         Width           =   1485
      End
      Begin VB.TextBox txt区 
         Height          =   300
         Left            =   3600
         MaxLength       =   50
         TabIndex        =   39
         Top             =   2370
         Width           =   1485
      End
      Begin VB.TextBox txt省 
         Height          =   300
         Left            =   1020
         MaxLength       =   50
         TabIndex        =   37
         Top             =   2370
         Width           =   1485
      End
      Begin VB.Frame Frame1 
         Height          =   135
         Left            =   60
         TabIndex        =   35
         Top             =   2070
         Width           =   7875
      End
      Begin VB.TextBox txt档案号 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1020
         MaxLength       =   50
         TabIndex        =   12
         Top             =   540
         Width           =   1485
      End
      Begin VB.TextBox txt说明 
         Height          =   300
         Left            =   6330
         MaxLength       =   50
         TabIndex        =   28
         Top             =   1320
         Width           =   1515
      End
      Begin VB.TextBox txtEMAIL 
         Height          =   300
         Left            =   3600
         MaxLength       =   50
         TabIndex        =   26
         Top             =   1320
         Width           =   1515
      End
      Begin VB.ComboBox cbo婚姻状况 
         Height          =   300
         Left            =   1020
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   1320
         Width           =   1485
      End
      Begin VB.TextBox txt证件号 
         Height          =   300
         Left            =   6330
         MaxLength       =   50
         TabIndex        =   22
         Top             =   930
         Width           =   1515
      End
      Begin VB.ComboBox cbo证件类型 
         Height          =   300
         Left            =   3600
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   930
         Width           =   1515
      End
      Begin MSMask.MaskEdBox msk出生日期 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "yyyy-MM-dd"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   3
         EndProperty
         Height          =   300
         Left            =   1020
         TabIndex        =   18
         Top             =   930
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Format          =   "yyyy-MM-dd"
         Mask            =   "####-##-##"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox cbo性别 
         Height          =   300
         Left            =   6330
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   540
         Width           =   1515
      End
      Begin VB.TextBox txt姓名 
         Height          =   300
         Left            =   3600
         MaxLength       =   50
         TabIndex        =   14
         Top             =   540
         Width           =   1515
      End
      Begin VB.Label lbl保险号 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "保险号"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   5700
         TabIndex        =   33
         Top             =   1770
         Width           =   540
      End
      Begin VB.Label lbl卡号 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "卡号"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3150
         TabIndex        =   31
         Top             =   1770
         Width           =   360
      End
      Begin VB.Label lbl卡类型 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "卡类型"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   390
         TabIndex        =   29
         Top             =   1770
         Width           =   540
      End
      Begin VB.Label lbl家属电话 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "家属电话"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   5520
         TabIndex        =   62
         Top             =   3990
         Width           =   720
      End
      Begin VB.Label lbl家属姓名 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "家属姓名"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   2790
         TabIndex        =   60
         Top             =   3990
         Width           =   720
      End
      Begin VB.Label lbl手机号 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "手机号"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   390
         TabIndex        =   58
         Top             =   3990
         Width           =   540
      End
      Begin VB.Label lbl职业 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "职业"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   5880
         TabIndex        =   56
         Top             =   3600
         Width           =   360
      End
      Begin VB.Label lbl单位电话 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "单位电话"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   2790
         TabIndex        =   54
         Top             =   3600
         Width           =   720
      End
      Begin VB.Label lbl单位邮编 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "单位邮编"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   210
         TabIndex        =   52
         Top             =   3600
         Width           =   720
      End
      Begin VB.Label lbl单位地址 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "单位地址"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3960
         TabIndex        =   50
         Top             =   3210
         Width           =   720
      End
      Begin VB.Label lbl工作单位 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "工作单位"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   210
         TabIndex        =   48
         Top             =   3210
         Width           =   720
      End
      Begin VB.Label lbl电话 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "电话"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   5880
         TabIndex        =   46
         Top             =   2820
         Width           =   360
      End
      Begin VB.Label lbl邮编 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "邮编"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3540
         TabIndex        =   44
         Top             =   2820
         Width           =   360
      End
      Begin VB.Label lbl地址 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "地址"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   570
         TabIndex        =   42
         Top             =   2820
         Width           =   360
      End
      Begin VB.Label lbl街道 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "街道"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   5880
         TabIndex        =   40
         Top             =   2430
         Width           =   360
      End
      Begin VB.Label lbl区 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "区"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3330
         TabIndex        =   38
         Top             =   2430
         Width           =   180
      End
      Begin VB.Label lbl省 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "省/市"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   480
         TabIndex        =   36
         Top             =   2430
         Width           =   450
      End
      Begin VB.Label lbl档案号 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "档案号"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   390
         TabIndex        =   11
         Top             =   600
         Width           =   540
      End
      Begin VB.Label lbl说明 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "说明"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   5880
         TabIndex        =   27
         Top             =   1380
         Width           =   360
      End
      Begin VB.Label lblEMAIL 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "EMAIL"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3060
         TabIndex        =   25
         Top             =   1380
         Width           =   450
      End
      Begin VB.Label lbl婚姻状况 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "婚姻状况"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   210
         TabIndex        =   23
         Top             =   1380
         Width           =   720
      End
      Begin VB.Label lbl证件号 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "证件号"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   5700
         TabIndex        =   21
         Top             =   990
         Width           =   540
      End
      Begin VB.Label lbl证件类型 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "证件类型"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   2790
         TabIndex        =   19
         Top             =   990
         Width           =   720
      End
      Begin VB.Label lbl出生日期 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "出生日期"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   210
         TabIndex        =   17
         Top             =   990
         Width           =   720
      End
      Begin VB.Label lbl性别 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "性别"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   5880
         TabIndex        =   15
         Top             =   600
         Width           =   360
      End
      Begin VB.Label lbl姓名 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "姓名"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3150
         TabIndex        =   13
         Top             =   600
         Width           =   360
      End
   End
   Begin VB.OptionButton opt 
      Caption         =   "身份证"
      Height          =   225
      Index           =   2
      Left            =   2820
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   210
      Width           =   885
   End
   Begin VB.OptionButton opt 
      Caption         =   "明码"
      Height          =   225
      Index           =   1
      Left            =   2010
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   210
      Width           =   675
   End
   Begin VB.OptionButton opt 
      Caption         =   "就诊卡"
      Height          =   225
      Index           =   0
      Left            =   270
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   210
      Value           =   -1  'True
      Width           =   885
   End
   Begin VB.TextBox txt卡号 
      Height          =   300
      Left            =   960
      MaxLength       =   50
      TabIndex        =   9
      Top             =   570
      Width           =   4005
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshFact 
      Height          =   915
      Left            =   300
      TabIndex        =   67
      Top             =   960
      Width           =   7860
      _ExtentX        =   13864
      _ExtentY        =   1614
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   250
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      MouseIcon       =   "frmIdentifyNBYKT.frx":0028
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label lbl卡状态 
      AutoSize        =   -1  'True
      Caption         =   "遗失注销"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   5100
      TabIndex        =   66
      Top             =   600
      Width           =   840
   End
   Begin VB.Label lbl请刷卡 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "请刷卡"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   330
      TabIndex        =   8
      Top             =   630
      Width           =   540
   End
End
Attribute VB_Name = "frmIdentifyNBYKT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstr操作类型 As String
Private mstr明码 As String
Private mstrUser As String
Private mstrPwd As String
Private mbln消息转发 As Boolean
Private mstr完整地址 As String
Private mstr档案号 As String
Private mdomOutput As New MSXML2.DOMDocument
Dim intCount As Integer             '记录新卡,旧卡使用次数,以便下次设置缺省

Private Enum MSHCol
    姓名
    性别
    住址
    发卡医院
End Enum

Public Function ReadCard(ByVal str完整地址 As String, ByVal strUser As String, ByVal strPwd As String, ByVal bln消息转发 As Boolean) As String
    mstr档案号 = ""
    mstr操作类型 = ""
    mstrUser = strUser
    mstrPwd = strPwd
    mbln消息转发 = bln消息转发
    mstr完整地址 = str完整地址
    Me.Show 1
    ReadCard = mstr档案号
End Function

Private Sub chk卡类型_Click()
    mstr操作类型 = IIf(chk卡类型.value = 1, "通用就诊卡", "就诊卡")
End Sub

Private Sub cmdCard_Click(Index As Integer)
    Dim blnNew As Boolean           '是否建立新的病人档案
    Dim strSQL As String
    Dim str档案号 As String
    Dim lng病人ID As Long
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    '由于新建档案的病人无档案号,将病人ID做为档案号上传,以后再使用到该病人,更新档案号
    
    If Index = 0 Then
        If Me.txt姓名.Text = "" Then
            MsgBox "病人姓名不能为空！"
            Exit Sub
        End If
        If Val(lbl卡状态.Tag) <> 0 Then
            MsgBox "当前卡状态为：" & lbl卡状态.Caption & "，不允许使用！"
            Exit Sub
        End If
        
        '确定,保存或更新病人信息
        str档案号 = txt档案号.Text
        '检查是否存在该病人的信息
        '1、如果有此档案号，说明存在该病人
        '2、如果病人发卡记录中旧卡号存在，说明存在该病人
        strSQL = " Select * From 病人信息 Where IC卡号=[1]"
        'Call OpenRecordset(rsTemp, "检查是否存在该病人的信息", strSQL)
        Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "检查是否存在该病人的信息", str档案号)
        If rsTemp.RecordCount = 0 Then
            strSQL = " Select * From 病人信息 Where 病人ID=(Select 病人ID From 病人发卡记录 Where 新卡号=[1])"
            'Call OpenRecordset(rsTemp, "检查是否存在该病人的信息", strSQL)
            Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "检查是否存在该病人的信息", Me.txt卡号.Text)
            If rsTemp.RecordCount = 0 Then
                blnNew = True
            End If
        End If
        
        If blnNew Then
            '无保险帐户则认为没有病人信息
            If lng病人ID = 0 Then lng病人ID = gobjDatabase.GetNextNO(1)
            strSQL = "zl_病人信息_Insert(" & lng病人ID & ",NULL,NULL,'自费医疗'," & _
                "'" & txt姓名.Text & "','" & cbo性别.Text & "'," & DateDiff("yyyy", msk出生日期.Text, gobjDatabase.CurrentDate()) & "," & _
                "To_Date('" & Me.msk出生日期.Text & "','YYYY-MM-DD')," & _
                "NULL,'" & IIf(Me.cbo证件类型.ListIndex = 0, Me.txt证件号.Text, "") & "',NULL,'" & cbo职业.Text & "'," & _
                "NULL,NULL,NULL,NULL,'" & Me.txt地址.Text & "','" & Me.txt电话.Text & "','" & Me.txt邮编.Text & "'," & _
                "'" & Me.txt家属姓名.Text & "',NULL,NULL,'" & Me.txt家属电话.Text & "',NULL,'" & txt工作单位.Text & "','" & txt单位电话.Text & "','" & txt单位邮编.Text & "','" & txt单位地址.Text & "'," & _
                "NULL,NULL,NULL,NULL,To_Date('" & Format(gobjDatabase.CurrentDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),NULL,NULL,NULL,NULL,NULL,'" & IIf(Me.cbo证件类型.ListIndex = 0, "", Me.txt证件号.Text) & "')"
        Else
            lng病人ID = rsTemp!病人ID
            strSQL = "zl_病人信息_Update(" & _
                lng病人ID & "," & Nvl(rsTemp!门诊号, "NULL") & "," & Nvl(rsTemp!住院号, "NULL") & ",'" & Nvl(rsTemp!费别) & "'," & _
                "'" & Nvl(rsTemp!医疗付款方式) & "','" & txt姓名.Text & "','" & Me.cbo性别.Text & "'," & DateDiff("yyyy", msk出生日期.Text, gobjDatabase.CurrentDate()) & "," & _
                "To_Date('" & msk出生日期.Text & "','YYYY-MM-DD')," & _
                "'" & IIf(IsNull(rsTemp!出生地点), "", rsTemp!出生地点) & "','" & IIf(Me.cbo证件类型.ListIndex = 0, Me.txt证件号.Text, "") & "'," & _
                "'" & IIf(IsNull(rsTemp!身份), "", rsTemp!身份) & "','" & cbo职业.Text & "'," & _
                "'" & IIf(IsNull(rsTemp!民族), "", rsTemp!民族) & "','" & IIf(IsNull(rsTemp!国籍), "", rsTemp!国籍) & "'," & _
                "'" & IIf(IsNull(rsTemp!学历), "", rsTemp!学历) & "','" & IIf(IsNull(rsTemp!婚姻状况), "", rsTemp!婚姻状况) & "'," & _
                "'" & txt地址.Text & "','" & txt电话.Text & "','" & txt邮编.Text & "','" & txt家属姓名.Text & "'," & _
                "'" & IIf(IsNull(rsTemp!联系人关系), "", rsTemp!联系人关系) & "','" & IIf(IsNull(rsTemp!联系人地址), "", rsTemp!联系人地址) & "'," & _
                "'" & txt家属电话.Text & "'," & IIf(IsNull(rsTemp!合同单位ID), "NULL", rsTemp!合同单位ID) & "," & _
                "'" & txt工作单位.Text & "','" & txt单位电话.Text & "'," & _
                "'" & txt单位邮编.Text & "','" & IIf(IsNull(rsTemp!单位开户行), "", rsTemp!单位开户行) & "'," & _
                "'" & IIf(IsNull(rsTemp!单位帐号), "", rsTemp!单位帐号) & "','" & IIf(IsNull(rsTemp!担保人), "", rsTemp!担保人) & "'," & _
                "" & IIf(IsNull(rsTemp!担保额), "NULL", rsTemp!担保额) & "," & Nvl(rsTemp!险类, "NULL") & ")"
        End If
        gcnConnect.Execute strSQL, , adCmdStoredProc
        
        '更新IC卡号,就诊卡号
        If InStr(1, "1005,1006,1007", mshFact.Tag) <> 0 Then
            strSQL = "zl_病人信息_更新信息(" & lng病人ID & ",'就诊卡号','''" & txt卡号码.Text & "''')"        '这是医保返回的卡号
            gcnConnect.Execute strSQL, , adCmdStoredProc
        End If
        strSQL = "zl_病人信息_更新信息(" & lng病人ID & ",'IC卡号','''" & IIf(str档案号 = "", lng病人ID, str档案号) & "''')"
        gcnConnect.Execute strSQL, , adCmdStoredProc
        strSQL = "zl_病人信息_更新信息(" & lng病人ID & ",'一卡通建档时间','''" & Me.Tag & "''')"
        gcnConnect.Execute strSQL, , adCmdStoredProc
        strSQL = "zl_病人信息_更新信息(" & lng病人ID & ",'操作类型','''" & mstr操作类型 & "''')"
        gcnConnect.Execute strSQL, , adCmdStoredProc
'        strSQL = "zl_病人信息_更新信息(" & lng病人ID & ",'备注','''" & txt省.Text & "|" & txt区.Text & "|" & txt街道.Text & "|" & txt单位地址.Text & "|" & txt手机号.Text & "|" & txtEMAIL.Text & "''')"
'        gcnConnect.Execute strSQL, , adCmdStoredProc
        gcnConnect.Execute "zl_病人信息从表_Update(" & lng病人ID & ",'省','" & txt省.Text & "')", , adCmdStoredProc
        gcnConnect.Execute "zl_病人信息从表_Update(" & lng病人ID & ",'区','" & txt区.Text & "')", , adCmdStoredProc
        gcnConnect.Execute "zl_病人信息从表_Update(" & lng病人ID & ",'街道','" & txt街道.Text & "')", , adCmdStoredProc
        gcnConnect.Execute "zl_病人信息从表_Update(" & lng病人ID & ",'单位地址','" & txt单位地址.Text & "')", , adCmdStoredProc
        gcnConnect.Execute "zl_病人信息从表_Update(" & lng病人ID & ",'手机号','" & txt手机号.Text & "')", , adCmdStoredProc
        gcnConnect.Execute "zl_病人信息从表_Update(" & lng病人ID & ",'EMAIL','" & txtEMAIL.Text & "')", , adCmdStoredProc
        
        'todo:只要在我们系统是插入新病人,则产生一条病人发卡记录
        If rsTemp.RecordCount = 0 Then
            '如果是用的旧的就诊卡,txt卡号码肯定为空,此时应该将输入的旧卡号写入病人发卡记录中
            strSQL = "zl_病人发卡记录_换补卡(" & lng病人ID & ",' " & IIf(txt卡号码.Text = "", Me.txt卡号.Text, txt卡号码.Text) & "',NULL," & _
                "'" & mshFact.TextMatrix(mshFact.Row, 发卡医院) & "'," & Me.cbo卡类型.ItemData(Me.cbo卡类型.ListIndex) & ",'" & mstr明码 & "','" & cmdCard(0).Tag & "')"
            gcnConnect.Execute strSQL, , adCmdStoredProc
        End If
        
        mstr档案号 = IIf(str档案号 = "", lng病人ID, str档案号)
    End If
    
  '  MsgBox lng病人ID & "|" & mstr档案号
    Unload Me
    Exit Sub
errHand:
    MsgBox Err.Description
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call gobjCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    
    With Me.cbo婚姻状况
        .AddItem "已婚"
        .ItemData(.NewIndex) = 0
        .AddItem "未婚"
        .ItemData(.NewIndex) = 1
        .AddItem "丧偶"
        .ItemData(.NewIndex) = 2
        .AddItem "离婚"
        .ItemData(.NewIndex) = 3
        .AddItem "其他"
        .ItemData(.NewIndex) = 9
        .ListIndex = 0
    End With
    
    With Me.cbo性别
        .AddItem "男"
        .ItemData(.NewIndex) = 0
        .AddItem "女"
        .ItemData(.NewIndex) = 1
        .AddItem "其他"
        .ItemData(.NewIndex) = 9
        .ListIndex = 0
    End With
    
    With Me.cbo卡类型
        .AddItem "医保卡"
        .ItemData(.NewIndex) = 0
        .AddItem "农保卡"
        .ItemData(.NewIndex) = 1
        .AddItem "就诊卡"
        .ItemData(.NewIndex) = 2
        .AddItem "其他卡"
        .ItemData(.NewIndex) = 9
        .ListIndex = 0
    End With
    
    With Me.cbo证件类型
        .AddItem "身份证"
        .ItemData(.NewIndex) = 0
        .AddItem "其他"
        .ItemData(.NewIndex) = 9
        .ListIndex = 0
    End With
    
    strSQL = "Select 编码,名称 From 职业 Order by 编码"
    Call OpenRecordset(rsTemp, "提取职业", strSQL)
    With rsTemp
        Do While Not .EOF
            Me.cbo职业.AddItem !名称
            .MoveNext
        Loop
        Me.cbo职业.ListIndex = 0
    End With
    
    Call ClearCons
    
    '根据昨天或当天新卡旧卡使用次数设定缺省
    Dim lng新卡 As Long, lng旧卡 As Long
    '更新注册表中新卡,旧卡累计次数
    lng旧卡 = Val(GetSetting("ZLSOFT", "宁波一卡通", "旧卡", 0))
    lng新卡 = Val(GetSetting("ZLSOFT", "宁波一卡通", "新卡", 0))
    chk卡类型.value = IIf(lng新卡 >= lng旧卡, 1, 0)
    mstr操作类型 = IIf(lng新卡 >= lng旧卡, "通用就诊卡", "就诊卡")
    Call InitMsh
End Sub

Private Sub mshFact_EnterCell()
    Dim nodRowset As MSXML2.IXMLDOMElement, nodRow As MSXML2.IXMLDOMElement
    On Error GoTo errHand
    
    If mshFact.TextMatrix(mshFact.Row, 姓名) = "" Then
        sstab.Enabled = True
        txt卡号.Text = ""
        Exit Sub
    End If
    msk出生日期.Text = "2000-01-01"         '有些病人没有出生日期,设个缺省值先
    cmdCard(0).Tag = ""
    
    '根据选择的病人显示详细信息
    Select Case mshFact.Tag
    Case "1001"
        Me.lbl卡状态.Tag = 0
        Me.lbl卡状态.Caption = "正常"
        Set nodRowset = mdomOutput.childNodes(1).childNodes(0).childNodes(0).childNodes(0)
        For Each nodRow In nodRowset.childNodes
            If mshFact.TextMatrix(mshFact.Row, 姓名) = nodRow.selectSingleNode("name").Text And mshFact.TextMatrix(mshFact.Row, 住址) = nodRow.selectSingleNode("address").Text Then
                '姓名与住址相同的提出来
                'Me.txt卡号码.Text = nodRow.selectSingleNode("id").selectSingleNode("cardNumber").Text '旧卡号不更新到就诊卡字段中,必须重复数据出现
                Me.txt卡号码.Text = ""
                Me.txtEMAIL.Text = nodRow.selectSingleNode("email").Text
                Me.txt单位地址.Text = nodRow.selectSingleNode("companyAddress").Text
                Me.txt单位电话.Text = nodRow.selectSingleNode("companyPhone").Text
                Me.txt单位邮编.Text = nodRow.selectSingleNode("companyPostcode").Text
                Me.txt档案号.Text = ""  '老卡病人肯定没有档案号
                Me.txt地址.Text = nodRow.selectSingleNode("address").Text
                Me.txt电话.Text = nodRow.selectSingleNode("homePhone").Text
                Me.txt工作单位.Text = nodRow.selectSingleNode("company").Text
                Me.txt家属电话.Text = nodRow.selectSingleNode("folkPhoneNumber").Text
                Me.txt家属姓名.Text = nodRow.selectSingleNode("folkName").Text
                Me.txt街道.Text = nodRow.selectSingleNode("street").Text
                Me.txt区.Text = nodRow.selectSingleNode("district").Text
                Me.txt省.Text = nodRow.selectSingleNode("province").Text
                Me.txt手机号.Text = nodRow.selectSingleNode("mobile").Text
                Me.txt姓名.Text = nodRow.selectSingleNode("name").Text
                Me.txt邮编.Text = nodRow.selectSingleNode("homePostcode").Text
                Me.txt保险号.Text = ""
                Me.txt证件号.Text = ""
                If Val(nodRow.selectSingleNode("cftype").Text) = 0 Then
                    Me.txt证件号.Text = nodRow.selectSingleNode("cfnumber").Text
                End If
                'Me.Tag = nodRow.selectSingleNode("createTime").Text     '保存病人信息的建档时间,后面补卡时要用
                Me.cbo婚姻状况.ListIndex = Val(nodRow.selectSingleNode("wedlock").Text)
                Me.cbo性别.ListIndex = Val(nodRow.selectSingleNode("sex").Text)
                Me.cbo证件类型.ListIndex = IIf(Val(nodRow.selectSingleNode("cftype").Text) = 0, 0, 1)
                Me.cbo职业.Text = nodRow.selectSingleNode("metier").Text
                
                If nodRow.selectSingleNode("birthday").Text <> "" Then
                    Me.msk出生日期.Text = Mid(nodRow.selectSingleNode("birthday").Text, 1, 10)
                End If
            End If
        Next
    Case Else
        '只可能有一条记录
        Me.txt保险号.Text = ""
        Set nodRowset = mdomOutput.childNodes(1).childNodes(0).childNodes(0).childNodes(0)
        For Each nodRow In nodRowset.childNodes
            If nodRow.nodeName = "CardInfoRet" Then
                Me.txt卡号码.Text = nodRow.selectSingleNode("cardNumber").Text
                '更新卡状态(0：正常；1：暂停；2：注销；3：挂失；4：损坏注销；5：遗失注销)
                Me.lbl卡状态.Tag = Val(nodRow.selectSingleNode("cardStatus").Text)
                Select Case Val(Me.lbl卡状态.Tag)
                Case 0
                    Me.lbl卡状态.Caption = "正常"
                Case 1
                    Me.lbl卡状态.Caption = "暂停"
                Case 2
                    Me.lbl卡状态.Caption = "注销"
                Case 3
                    Me.lbl卡状态.Caption = "挂失"
                Case 4
                    Me.lbl卡状态.Caption = "损坏注销"
                Case 5
                    Me.lbl卡状态.Caption = "遗失注销"
                End Select
                cmdCard(0).Tag = nodRow.selectSingleNode("reportTime").Text '旧卡的建卡时间
            End If
            If nodRow.nodeName = "TPersonBasalInfo" Then
                Me.txt档案号.Text = nodRow.selectSingleNode("personid").Text
                Me.txt姓名.Text = nodRow.selectSingleNode("name").Text
                Me.txtEMAIL.Text = nodRow.selectSingleNode("email").Text
                Me.cbo性别.ListIndex = Val(nodRow.selectSingleNode("sex").Text)
                If nodRow.selectSingleNode("birthday").Text <> "" Then
                    Me.msk出生日期.Text = Mid(nodRow.selectSingleNode("birthday").Text, 1, 10)
                End If
                Me.cbo婚姻状况.ListIndex = Val(nodRow.selectSingleNode("wedlock").Text)
                Me.cbo证件类型.ListIndex = IIf(Val(nodRow.selectSingleNode("cftype").Text) = 0, 0, 1)
                Me.txt证件号.Text = ""
                If Val(nodRow.selectSingleNode("cftype").Text) = 0 Then
                    Me.txt证件号.Text = nodRow.selectSingleNode("cfnumber").Text
                End If
                Me.Tag = nodRow.selectSingleNode("createTime").Text     '保存病人信息的建档时间,后面补卡时要用
            End If
            If nodRow.nodeName = "TPersonExtendInfo" Then
                Me.txt单位地址.Text = nodRow.selectSingleNode("companyAddress").Text
                Me.txt单位电话.Text = nodRow.selectSingleNode("companyPhone").Text
                Me.txt单位邮编.Text = nodRow.selectSingleNode("companyPostcode").Text
                Me.txt地址.Text = nodRow.selectSingleNode("address").Text
                Me.txt电话.Text = nodRow.selectSingleNode("homePhone").Text
                Me.txt工作单位.Text = nodRow.selectSingleNode("company").Text
                Me.txt家属电话.Text = nodRow.selectSingleNode("folkPhoneNumber").Text
                Me.txt家属姓名.Text = nodRow.selectSingleNode("folkName").Text
                Me.txt街道.Text = nodRow.selectSingleNode("street").Text
                Me.txt区.Text = nodRow.selectSingleNode("district").Text
                Me.txt省.Text = nodRow.selectSingleNode("province").Text
                Me.txt手机号.Text = nodRow.selectSingleNode("mobile").Text
                Me.txt邮编.Text = nodRow.selectSingleNode("homePostcode").Text
                Me.cbo职业.Text = nodRow.selectSingleNode("metier").Text
            End If
        Next
    
    End Select
    
    If mshFact.Tag = "1004" Then Me.txt保险号.Text = Me.txt卡号.Text
    If Me.Tag <> "" Then Me.Tag = Format(Mid(Me.Tag, 1, 10), "YYYYMMdd") & Format(Mid(Me.Tag, 12, 8), "HHmmss")
    If cmdCard(0).Tag <> "" Then cmdCard(0).Tag = Format(Mid(cmdCard(0).Tag, 1, 10), "YYYYMMdd") & Format(Mid(cmdCard(0).Tag, 12, 8), "HHmmss")
    
    Exit Sub
errHand:
    MsgBox "显示指定病人信息时发生错误:" & Err.Description
    Resume
End Sub

Private Sub opt_Click(Index As Integer)
    chk卡类型.Enabled = (Index = 0)
    
    Select Case Index
    Case 0
        If chk卡类型.value = 1 Then
            mstr操作类型 = "通用就诊卡"
        Else
            mstr操作类型 = "就诊卡"
        End If
    Case 1
        mstr操作类型 = "明码"
    Case 2
        mstr操作类型 = "身份证"
    Case 3
        mstr操作类型 = "医保卡"
    Case 4
        mstr操作类型 = "农保卡"
    Case 5
        mstr操作类型 = "异地医保卡"
    Case 6
        mstr操作类型 = "异地农保卡"
    End Select
End Sub

Private Sub txt卡号_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    mstr明码 = ""
    Call ClearCons
    Call InitMsh
    sstab.Enabled = False
    Call ReadPatient
End Sub

Private Sub ClearCons()
    Dim objControl As Object
    
    Me.lbl卡状态.Tag = 0
    Me.lbl卡状态.Caption = ""
    For Each objControl In Me.Controls
        If UCase(objControl.Container.Name) = "SSTAB" Then
            Select Case Mid(UCase(objControl.Name), 1, 3)
            Case "TXT"
                objControl.Text = ""
            Case "CBO"
                objControl.ListIndex = 0
            End Select
        End If
    Next
    
    If Me.opt(4).value Or Me.opt(6).value Then
        Me.cbo卡类型.ListIndex = 1
    ElseIf Me.opt(3).value Or Me.opt(5).value Then
        Me.cbo卡类型.ListIndex = 0
    ElseIf Me.opt(0).value Then
        Me.cbo卡类型.ListIndex = 2
    Else
        Me.cbo卡类型.ListIndex = 3
    End If
    
End Sub

Private Function ReadPatient() As Boolean
    Dim strType As String
    Dim strPatient As String
    On Error GoTo errHand
    
    If opt(0).value Then
        If chk卡类型.value = 1 Then
            strType = "1005"
        Else
            strType = "1001"
        End If
    ElseIf opt(1).value Then
        strType = "1003"
    ElseIf opt(2).value Then
        strType = "1002"
    ElseIf opt(3).value Or opt(4).value Then
        strType = 1004
    ElseIf opt(5).value Then
        strType = 1007
    ElseIf opt(6).value Then
        strType = 1006
    End If
    
    '调用WebServices获取病人身份
    '1． 查询操作
    'a)  .参数说明
    'i.类型操作参数
    'SearchType:  查询操作类型
    'ii.查询关键字参数
    'Cardnumber:  卡号
    'Sfzj:  身份证号
    'Jzkmm:  就诊卡明码
    'Bxh:  保险号
    '
    'b)  .操作定义
    'i.  getPersonInfo(String SearchType,String [查询关键字参数])
    'ii.类型操作定义说明
    '1001: 通过卡号查询旧卡信息
    '1002: 通过身份证号码查询卡信息
    '1003: 通过就诊卡明码查询卡信息
    '1004: 通过保险号查询卡信息
    '1005: 通过卡号查询卡信息
    If Not 调用接口("getPersonInfo", strType, txt卡号.Text) Then
        sstab.Enabled = True
        txt卡号.Text = ""
        Exit Function
    End If
    '将病人信息填入表格中
    Me.mshFact.Tag = strType
    Call AnalysePatient
    
    '更新注册表中新卡,旧卡累计次数
    Dim lng新卡 As Long, lng旧卡 As Long, str日期 As String
    lng旧卡 = Val(GetSetting("ZLSOFT", "宁波一卡通", "旧卡", 0))
    lng新卡 = Val(GetSetting("ZLSOFT", "宁波一卡通", "新卡", 0))
    str日期 = Format(gobjDatabase.CurrentDate(), "yyyyMMdd")
    If str日期 <> GetSetting("ZLSOFT", "宁波一卡通", "日期", "") Then lng新卡 = 0: lng旧卡 = 0
    If opt(0).value Then
        If chk卡类型.value = 1 Then
            lng新卡 = lng新卡 + 1
        Else
            lng旧卡 = lng旧卡 + 1
        End If
        Call SaveSetting("ZLSOFT", "宁波一卡通", "旧卡", lng旧卡)
        Call SaveSetting("ZLSOFT", "宁波一卡通", "新卡", lng新卡)
        Call SaveSetting("ZLSOFT", "宁波一卡通", "日期", str日期)
    End If
    
    If mstr操作类型 = "明码" Then mstr明码 = Me.txt卡号.Text
    If opt(2).value Then Me.txt卡号.Text = ""
    mshFact.Row = 1: mshFact.Col = 0
    Call mshFact_EnterCell
    Exit Function
errHand:
    MsgBox Err.Description
End Function

Private Function AnalysePatient() As Boolean
    Dim intRow As Integer
    Dim nodRowset As MSXML2.IXMLDOMElement, nodRow As MSXML2.IXMLDOMElement
    On Error GoTo errHand
    '解析病人信息返回串
    
    intRow = 1
    Select Case mshFact.Tag
    Case "1001"
        Set nodRowset = mdomOutput.childNodes(1).childNodes(0).childNodes(0).childNodes(0)
        For Each nodRow In nodRowset.childNodes
            mshFact.TextMatrix(intRow, 姓名) = nodRow.selectSingleNode("name").Text
            mshFact.TextMatrix(intRow, 性别) = IIf(nodRow.selectSingleNode("sex").Text = "0", "男", "女")
            mshFact.TextMatrix(intRow, 住址) = nodRow.selectSingleNode("address").Text
            mshFact.TextMatrix(intRow, 发卡医院) = nodRow.selectSingleNode("id").selectSingleNode("hospitalid").Text
            intRow = intRow + 1
            mshFact.Rows = mshFact.Rows + 1
        Next
    Case Else       '1005
        Set nodRowset = mdomOutput.childNodes(1).childNodes(0).childNodes(0).childNodes(0)
        For Each nodRow In nodRowset.childNodes
            If nodRow.nodeName = "CardInfoRet" Then
                mshFact.TextMatrix(intRow, 发卡医院) = nodRow.selectSingleNode("hospitalNumber").Text
            End If
            If nodRow.nodeName = "TPersonBasalInfo" Then
                mshFact.TextMatrix(intRow, 姓名) = nodRow.selectSingleNode("name").Text
                mshFact.TextMatrix(intRow, 性别) = IIf(nodRow.selectSingleNode("sex").Text = "0", "男", "女")
            End If
            If nodRow.nodeName = "TPersonExtendInfo" Then
                mshFact.TextMatrix(intRow, 住址) = nodRow.selectSingleNode("address").Text
            End If
        Next
    End Select
    
    AnalysePatient = True
    Exit Function
errHand:
    MsgBox "装载病人信息时发生错误:" & Err.Description
End Function

Private Sub OpenRecordset(rsTemp As ADODB.Recordset, ByVal strCaption As String, Optional strSQL As String = "")
'功能：打开记录集
    If rsTemp.State = adStateOpen Then rsTemp.Close
    rsTemp.CursorLocation = adUseClient
    
    rsTemp.Open strSQL, gcnConnect, adOpenStatic, adLockReadOnly
    Set rsTemp.ActiveConnection = Nothing
End Sub

Private Function Nvl(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'功能：相当于Oracle的NVL，将Null值改成另外一个预设值
    Dim varReturn As Variant
    varReturn = IIf(IsNull(varValue), DefaultValue, varValue)
    Nvl = Replace(varReturn, "'", "")
End Function

Private Function GetElemnetValue(ByVal Name As String) As String
'功能：得到指定元素的值
    Dim xmlElement As MSXML2.IXMLDOMElement
    
    Set xmlElement = mdomOutput.documentElement.selectSingleNode(Name)
    If Not xmlElement Is Nothing Then
        '找到指定子元素
        GetElemnetValue = xmlElement.Text
'    Else
'        '取消
'        Debug.Assert False
    End If
End Function

Private Function GetAttributeValue(xmlElement As MSXML2.IXMLDOMElement, ByVal Name As String) As String
'功能：得到指定属性的值
    Dim varAttribute As Variant
    
    varAttribute = xmlElement.getAttribute(Name)
    If IsNull(varAttribute) = False Then
        GetAttributeValue = varAttribute
    End If
End Function

Private Function 调用接口(ByVal strFunction As String, ByVal strType As String, ByVal strKey As String) As Boolean
'    ----------------------------------------------------------------
    '功能描述   ：调用接口函数
    '编写人     ：朱玉宝
'    编写日期   ：2009-07-31
'    ----------------------------------------------------------------
    Dim str日期 As String, lng序列号 As Long, str错误信息 As String
    Dim strURL As String, strSoapRequest As String
    Dim objHttp As MSXML2.XMLHTTP
    On Error GoTo errHand
    
    Set objHttp = New MSXML2.XMLHTTP
    strURL = mstr完整地址 & "?op=" & strFunction
    
    strSoapRequest = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "utf-8" & Chr(34) & "?>" & _
                "<soapenv:Envelope xmlns:soapenv=" & Chr(34) & "http://schemas.xmlsoap.org/soap/envelope/" & Chr(34) & ">" & _
                "<soapenv:Header>" & _
                    "<ns:" & strFunction & " xmlns:ns=" & Chr(34) & "http://service.wondersgroup.com" & Chr(34) & ">" & _
                        "<ns:user>" & mstrUser & "</ns:user>" & _
                        "<ns:pwd>" & mstrPwd & "</ns:pwd>" & _
                    "</ns:" & strFunction & ">" & _
                "</soapenv:Header>" & _
                "<soapenv:Body>" & _
                    "<ns:" & strFunction & " xmlns:ns=" & Chr(34) & "http://service.wondersgroup.com" & Chr(34) & ">" & _
                        "<ns:SearchType>" & strType & "</ns:SearchType>"
    Select Case strType
    Case "1001"
        strSoapRequest = strSoapRequest & "<ns:CardNumber>" & strKey & "</ns:CardNumber>"
    Case "1002"
        strSoapRequest = strSoapRequest & "<ns:Sfzh>" & strKey & "</ns:Sfzh>"
    Case "1003"
        strSoapRequest = strSoapRequest & "<ns:Jzkmm>" & strKey & "</ns:Jzkmm>"
    Case "1004"
        strSoapRequest = strSoapRequest & "<ns:Bxh>" & strKey & "</ns:Bxh>"
    Case "1005"
        strSoapRequest = strSoapRequest & "<ns:CardNumber>" & strKey & "</ns:CardNumber>"
    Case "1006"
        strSoapRequest = strSoapRequest & "<ns:CardNumber>" & strKey & "</ns:CardNumber>"
    Case "1007"
        strSoapRequest = strSoapRequest & "<ns:CardNumber>" & strKey & "</ns:CardNumber>"
    End Select
                        
    strSoapRequest = strSoapRequest & _
                    "</ns:" & strFunction & ">" & _
                "</soapenv:Body>" & _
                "</soapenv:Envelope>"
    If mbln消息转发 = False Then
        objHttp.Open "post", strURL, False
        objHttp.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
        objHttp.setRequestHeader "Content-Length", Len(strSoapRequest)
        objHttp.setRequestHeader "SOAPAction", strURL
        
        '根据返回的状态信息来判断是否成功
        objHttp.send (strSoapRequest)
        If objHttp.status <> 200 Then
            MsgBox "返回信息：[" & objHttp.status & "]" & objHttp.responseText
            Exit Function
        End If
    Else
        '写入数据
        If Not SendRequest(str日期, lng序列号, strFunction, strURL, strSoapRequest) Then Exit Function
        
        '显示等待窗体
        If frmWait.SendRequest(str日期, lng序列号, str错误信息) = False Then
            If str错误信息 <> "" Then MsgBox "返回信息：" & str错误信息
            Exit Function
        End If
    End If
    
    '断点设置处
    Set mdomOutput = New MSXML2.DOMDocument
    If mbln消息转发 = False Then
        If mdomOutput.loadXML(objHttp.responseText) = False Then
            MsgBox "交易函数：" & strFunction & "，返回数据格式不正确！"
            Exit Function
        End If
    Else
        If mdomOutput.loadXML(str错误信息) = False Then
            MsgBox "交易函数：" & strFunction & "，返回数据格式不正确！"
            Exit Function
        End If
    End If
    
    调用接口 = True
    Exit Function
errHand:
    MsgBox Err.Description
End Function

Private Function SendRequest(str日期 As String, lng序列号 As Long, _
    ByVal strFuncName As String, ByVal strURL As String, ByVal strSoapRequest As String) As Boolean
    Dim blnTrans As Boolean
    Dim strRow As String
    Dim intRow As Integer, intCount As Integer
    On Error GoTo errHand
    '将待发送数据写入数据表
    
    str日期 = Format(gobjDatabase.CurrentDate, "yyyyMMdd")
    lng序列号 = gobjDatabase.GetNextId("消息转发")
    
    gcnConnect.BeginTrans
    blnTrans = True
    
    '插入主表
    gcnConnect.Execute "zl_消息主表_Insert('" & str日期 & "'," & lng序列号 & ",'" & strFuncName & "','" & strURL & "')", , adCmdStoredProc
    
    '插入待发送数据
    intCount = Len(strSoapRequest) \ 1000
    If Len(strSoapRequest) Mod 1000 <> 0 Then intCount = intCount + 1
    For intRow = 0 To intCount
        strRow = Mid(strSoapRequest, intRow * 1000 + 1, 1000)
        gcnConnect.Execute "zl_消息转发_Insert('" & str日期 & "'," & lng序列号 & "," & intRow + 1 & ",'" & strRow & "')", , adCmdStoredProc
    Next
    
    gcnConnect.CommitTrans
    blnTrans = False
    SendRequest = True
    Exit Function
errHand:
    If blnTrans Then gcnConnect.RollbackTrans
    MsgBox Err.Description
End Function

Private Sub InitMsh()
    With mshFact
        .Clear
        .Rows = 2: .Cols = 4
        .TextMatrix(0, 姓名) = "姓名"
        .TextMatrix(0, 性别) = "性别"
        .TextMatrix(0, 住址) = "住址"
        .TextMatrix(0, 发卡医院) = "发卡医院"
        .ColWidth(姓名) = 1200
        .ColWidth(性别) = 500
        .ColWidth(住址) = 3000
        .ColWidth(发卡医院) = 1000
    End With
End Sub
