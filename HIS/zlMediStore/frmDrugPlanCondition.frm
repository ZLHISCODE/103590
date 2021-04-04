VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDrugPlanCondition 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "条件设置"
   ClientHeight    =   8490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7635
   Icon            =   "frmDrugPlanCondition.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   7635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin TabDlg.SSTab sstConditon 
      Height          =   7680
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   7395
      _ExtentX        =   13044
      _ExtentY        =   13547
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "计划(&1)"
      TabPicture(0)   =   "frmDrugPlanCondition.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl库房"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Lbl剂型"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Lvw剂型"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "fra区间"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "fra计划类型"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "fra方式"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "fra计划方法"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Chk不产生计划数量"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cbo库房"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "chkBaseMedi"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "fra常备药"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "chkOnlyBaseMedi"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "frm毒理分类"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "chk不考虑现库存"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Chk仅提取低取下限的药品"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "chkClearZeroPlan"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "chk参考销量"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "fra辅助条件"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Chk剂型"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).ControlCount=   20
      TabCaption(1)   =   "药品分类(&2)"
      TabPicture(1)   =   "frmDrugPlanCondition.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label3"
      Tab(1).Control(1)=   "tvw用途"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "供应商(&3)"
      TabPicture(2)   =   "frmDrugPlanCondition.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "tvw供货单位"
      Tab(2).Control(1)=   "chk中标单位"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "来源药房(&4)"
      TabPicture(3)   =   "frmDrugPlanCondition.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lvw药房"
      Tab(3).Control(1)=   "Label2"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "来源库房(&5)"
      TabPicture(4)   =   "frmDrugPlanCondition.frx":007C
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label5"
      Tab(4).Control(1)=   "Label6"
      Tab(4).Control(2)=   "lvw库房"
      Tab(4).ControlCount=   3
      Begin VB.CheckBox chk中标单位 
         Caption         =   "无上次供应商以中标单位为准(&W)"
         Enabled         =   0   'False
         Height          =   225
         Left            =   -74880
         TabIndex        =   50
         Top             =   420
         Width           =   2985
      End
      Begin VB.CheckBox Chk剂型 
         Appearance      =   0  'Flat
         Caption         =   "全选"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   840
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   2925
         Width           =   675
      End
      Begin VB.Frame fra辅助条件 
         Caption         =   " 辅助条件 "
         Enabled         =   0   'False
         Height          =   1710
         Left            =   3720
         TabIndex        =   38
         Top             =   4980
         Width           =   3495
         Begin VB.TextBox txt上限天数 
            Height          =   300
            Left            =   1560
            TabIndex        =   40
            Top             =   270
            Width           =   795
         End
         Begin VB.TextBox txt下限天数 
            Height          =   300
            Left            =   1560
            TabIndex        =   39
            Top             =   660
            Width           =   795
         End
         Begin VB.Label lbl上限天数 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "上限天数(&X)"
            ForeColor       =   &H00000000&
            Height          =   180
            Left            =   390
            TabIndex        =   42
            Top             =   330
            Width           =   990
         End
         Begin VB.Label lbl下限天数 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "下限天数(&T)"
            ForeColor       =   &H00000000&
            Height          =   180
            Left            =   390
            TabIndex        =   41
            Top             =   720
            Width           =   990
         End
      End
      Begin VB.CheckBox chk参考销量 
         Appearance      =   0  'Flat
         Caption         =   "按销量产生计划数量"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3240
         TabIndex        =   37
         Top             =   1140
         Width           =   3120
      End
      Begin VB.CheckBox chkClearZeroPlan 
         Appearance      =   0  'Flat
         Caption         =   "不产生计划数量为0的药品记录"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   240
         TabIndex        =   36
         Top             =   900
         Width           =   2760
      End
      Begin VB.CheckBox Chk仅提取低取下限的药品 
         Appearance      =   0  'Flat
         Caption         =   "仅提取低于下限的药品"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   240
         TabIndex        =   35
         Top             =   660
         Width           =   2205
      End
      Begin VB.CheckBox chk不考虑现库存 
         Appearance      =   0  'Flat
         Caption         =   "不考虑现库存数量"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   240
         TabIndex        =   34
         Top             =   1140
         Width           =   1800
      End
      Begin VB.Frame frm毒理分类 
         Caption         =   " 毒理分类选择"
         Height          =   615
         Left            =   120
         TabIndex        =   29
         Top             =   2220
         Width           =   7095
         Begin VB.CheckBox chk毒理 
            Caption         =   "提取普通药"
            Height          =   180
            Index           =   0
            Left            =   240
            TabIndex        =   33
            Top             =   280
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin VB.CheckBox chk毒理 
            Caption         =   "提取毒性药"
            Height          =   180
            Index           =   1
            Left            =   1680
            TabIndex        =   32
            Top             =   280
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin VB.CheckBox chk毒理 
            Caption         =   "提取精神类药"
            Height          =   180
            Index           =   2
            Left            =   3240
            TabIndex        =   31
            Top             =   280
            Value           =   1  'Checked
            Width           =   1455
         End
         Begin VB.CheckBox chk毒理 
            Caption         =   "提取麻醉药"
            Height          =   180
            Index           =   3
            Left            =   5040
            TabIndex        =   30
            Top             =   280
            Value           =   1  'Checked
            Width           =   1215
         End
      End
      Begin VB.CheckBox chkOnlyBaseMedi 
         Appearance      =   0  'Flat
         Caption         =   "仅含基本药物"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4800
         TabIndex        =   28
         Top             =   900
         Width           =   1800
      End
      Begin VB.Frame fra常备药 
         Caption         =   " 常备药选择 "
         Height          =   645
         Left            =   120
         TabIndex        =   24
         Top             =   1500
         Width           =   7095
         Begin VB.OptionButton opt常备药 
            Caption         =   "仅提取非常备药"
            Height          =   180
            Index           =   1
            Left            =   4440
            TabIndex        =   27
            Top             =   300
            Width           =   1695
         End
         Begin VB.OptionButton opt常备药 
            Caption         =   "仅提取常备药"
            Height          =   180
            Index           =   0
            Left            =   2400
            TabIndex        =   26
            Top             =   300
            Width           =   1575
         End
         Begin VB.OptionButton opt常备药 
            Caption         =   "不考虑是否常备药"
            Height          =   180
            Index           =   2
            Left            =   240
            TabIndex        =   25
            Top             =   300
            Value           =   -1  'True
            Width           =   1815
         End
      End
      Begin VB.CheckBox chkBaseMedi 
         Appearance      =   0  'Flat
         Caption         =   "包含基本药物"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3240
         TabIndex        =   23
         Top             =   900
         Value           =   1  'Checked
         Width           =   1440
      End
      Begin VB.ComboBox cbo库房 
         Height          =   276
         Left            =   1050
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   330
         Width           =   3210
      End
      Begin VB.CheckBox Chk不产生计划数量 
         Appearance      =   0  'Flat
         Caption         =   "不产生计划数量"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3240
         TabIndex        =   14
         Top             =   660
         Width           =   1560
      End
      Begin VB.Frame fra计划方法 
         Caption         =   " 编制方法 "
         Height          =   1710
         Left            =   120
         TabIndex        =   9
         Top             =   4980
         Width           =   3435
         Begin VB.OptionButton opt方法 
            Caption         =   "自定义区间参照法(&5)"
            Height          =   195
            Index           =   4
            Left            =   735
            TabIndex        =   17
            Top             =   1440
            Width           =   2190
         End
         Begin VB.OptionButton opt方法 
            Caption         =   "药品日出库量参照法(&4)"
            Height          =   195
            Index           =   3
            Left            =   735
            TabIndex        =   13
            Top             =   1170
            Width           =   2190
         End
         Begin VB.OptionButton opt方法 
            Caption         =   "药品储备定额参照法(&3)"
            Height          =   195
            Index           =   2
            Left            =   735
            TabIndex        =   12
            Top             =   885
            Width           =   2190
         End
         Begin VB.OptionButton opt方法 
            Caption         =   "临近期间平均参照法(&2)"
            Height          =   195
            Index           =   1
            Left            =   720
            TabIndex        =   11
            Top             =   585
            Width           =   2190
         End
         Begin VB.OptionButton opt方法 
            Caption         =   "往年同期线性参照法(&1)"
            Height          =   195
            Index           =   0
            Left            =   735
            TabIndex        =   10
            Top             =   270
            Value           =   -1  'True
            Width           =   2190
         End
      End
      Begin VB.Frame fra方式 
         Caption         =   " 产生数量方式"
         Height          =   1710
         Left            =   3720
         TabIndex        =   43
         Top             =   4980
         Width           =   3495
         Begin VB.OptionButton opt上限 
            Caption         =   "按库存上限产生计划数量"
            Height          =   195
            Left            =   240
            TabIndex        =   45
            Top             =   360
            Value           =   -1  'True
            Width           =   2415
         End
         Begin VB.OptionButton opt下限 
            Caption         =   "按库存下限产生计划数量"
            Height          =   195
            Left            =   240
            TabIndex        =   44
            Top             =   720
            Width           =   2775
         End
      End
      Begin VB.Frame fra计划类型 
         Caption         =   " 计划类型 "
         Height          =   765
         Left            =   120
         TabIndex        =   4
         Top             =   6780
         Width           =   7095
         Begin VB.OptionButton opt计划 
            Caption         =   "月度计划(&A)"
            Height          =   210
            Index           =   0
            Left            =   1845
            TabIndex        =   8
            Top             =   405
            Value           =   -1  'True
            Width           =   1290
         End
         Begin VB.OptionButton opt计划 
            Caption         =   "季度计划(&B)"
            Height          =   210
            Index           =   1
            Left            =   3555
            TabIndex        =   7
            Top             =   405
            Width           =   1290
         End
         Begin VB.OptionButton opt计划 
            Caption         =   "年度计划(&C)"
            Height          =   210
            Index           =   2
            Left            =   5130
            TabIndex        =   6
            Top             =   405
            Width           =   1290
         End
         Begin VB.OptionButton opt计划 
            Caption         =   "周计划(&W)"
            Height          =   210
            Index           =   3
            Left            =   120
            TabIndex        =   5
            Top             =   405
            Width           =   1290
         End
      End
      Begin VB.Frame fra区间 
         Caption         =   " 自定义区间"
         Height          =   765
         Left            =   120
         TabIndex        =   18
         Top             =   6780
         Visible         =   0   'False
         Width           =   7095
         Begin MSComCtl2.DTPicker dtp开始时间 
            Height          =   315
            Left            =   1320
            TabIndex        =   19
            Top             =   360
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy年MM月dd日"
            Format          =   303759363
            CurrentDate     =   36263
         End
         Begin MSComCtl2.DTPicker dtp结束时间 
            Height          =   315
            Left            =   3225
            TabIndex        =   20
            Top             =   360
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy年MM月dd日"
            Format          =   303759363
            CurrentDate     =   36263
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "时间范围"
            Height          =   180
            Left            =   360
            TabIndex        =   22
            Top             =   420
            Width           =   720
         End
         Begin VB.Label lbl至 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "～"
            Height          =   180
            Index           =   0
            Left            =   2985
            TabIndex        =   21
            Top             =   420
            Width           =   180
         End
      End
      Begin MSComctlLib.ListView Lvw剂型 
         Height          =   1680
         Left            =   120
         TabIndex        =   47
         Top             =   3180
         Width           =   7065
         _ExtentX        =   12462
         _ExtentY        =   2963
         View            =   2
         Arrange         =   1
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "img16"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "名称"
            Object.Width           =   3528
         EndProperty
      End
      Begin MSComctlLib.TreeView tvw用途 
         Height          =   6840
         Left            =   -74880
         TabIndex        =   49
         Top             =   660
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   12065
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   476
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         ImageList       =   "img16"
         Appearance      =   1
      End
      Begin MSComctlLib.TreeView tvw供货单位 
         Height          =   6840
         Left            =   -74880
         TabIndex        =   51
         Top             =   660
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   12065
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   494
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         ImageList       =   "img16"
         Appearance      =   1
      End
      Begin MSComctlLib.ListView lvw库房 
         Height          =   5880
         Left            =   -74880
         TabIndex        =   56
         Top             =   900
         Width           =   7068
         _ExtentX        =   12462
         _ExtentY        =   10372
         View            =   2
         Arrange         =   1
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "img16"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "名称"
            Object.Width           =   3528
         EndProperty
      End
      Begin MSComctlLib.ListView lvw药房 
         Height          =   6840
         Left            =   -74880
         TabIndex        =   52
         Top             =   660
         Width           =   7065
         _ExtentX        =   12462
         _ExtentY        =   12065
         View            =   2
         Arrange         =   1
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "img16"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "名称"
            Object.Width           =   3528
         EndProperty
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "(若都不勾选，库房为[全院]时默认统计所有库房库存，否则为当前库房库存)"
         Height          =   180
         Left            =   -74760
         TabIndex        =   58
         Top             =   650
         Width           =   6144
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "勾选下列库房作为统计库存数量的库房"
         Height          =   180
         Left            =   -74760
         TabIndex        =   57
         Top             =   420
         Width           =   3060
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "不选择剂型默认提取所有剂型药品"
         Height          =   180
         Left            =   1680
         TabIndex        =   55
         Top             =   2940
         Width           =   2700
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "如果不选择分类，则不产生计划内容"
         Height          =   180
         Left            =   -74760
         TabIndex        =   54
         Top             =   420
         Width           =   2880
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "勾选下列药房作为统计药品销售的发药药房，都不勾选时默认为所有药房"
         Height          =   180
         Left            =   -74760
         TabIndex        =   53
         Top             =   420
         Width           =   5760
      End
      Begin VB.Label Lbl剂型 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "剂型(&G)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   120
         TabIndex        =   48
         Top             =   2940
         Width           =   630
      End
      Begin VB.Label lbl库房 
         AutoSize        =   -1  'True
         Caption         =   "库房(&K)"
         Height          =   180
         Left            =   240
         TabIndex        =   16
         Top             =   390
         Width           =   630
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   240
      TabIndex        =   2
      Top             =   8040
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6360
      TabIndex        =   1
      Top             =   8040
      Width           =   1100
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   1800
      Top             =   7920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPlanCondition.frx":0098
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPlanCondition.frx":0F72
            Key             =   "Folder1"
            Object.Tag             =   "Folder1"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPlanCondition.frx":13C4
            Key             =   "Card"
            Object.Tag             =   "Card"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPlanCondition.frx":1816
            Key             =   "Folder"
            Object.Tag             =   "Folder"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   5040
      TabIndex        =   0
      Top             =   8040
      Width           =   1100
   End
End
Attribute VB_Name = "frmDrugPlanCondition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnSelect As Boolean
Private mstr用途ID As String
Private mstr剂型 As String
Private mlng库房ID As Long
Private mint计划类型 As Integer
Private mint编制方法 As Integer             '0-往年同期线性参照法；1-临近期间平均参照法；2-药品储备定额参照法；3-药品日销售量参照法；4-自定义区间参照法
Private mbln下限 As Boolean
Private mint上限 As Integer
Private mint下限 As Integer
Private mbln中药库房 As Boolean                     '中药库房
Private mfrmMain As Form
Private mbln计划数量 As Boolean
Private mstr供货商ID As String
Private mbln中标单位 As Boolean
Private mstrBeginDate As String
Private mstrEndDate As String
Private mbln不考虑库存 As Boolean
Private mblnClearZeroPlan As Boolean
Private mblnBaseMedi As Boolean
Private mblnOnlyBaseMedi As Boolean
Private mintStock As Integer            '常备药选择：0-只提取常备药；1-只提取非常备药；2-不区别是否常备药；
Private mbln数量方式 As Boolean         ' false-上限方式 true-下限方式
Private mintPlanPoint As Integer        '全院计划不管站点 0-要管站点，1-不管站点
Private mstrToxicologyClass As String       '毒理分类
Private mbln按销量产生计划 As Boolean   '按销量产生计划数量
Private mstr来源药房 As String               '格式:药房id1,药房id2...
Private mstr来源库房 As String               '格式:药房id1,药房id2...
Private mstrAll来源药房 As String       '所有来源药房。格式:药房id1,药房id2...
Private mstrAll来源库房 As String       '所有来源药房。格式:药房id1,药房id2...

Private Enum zlDrugPlan
    P0_往年同期线性参照法 = 0
    P1_临近期间平均参照法 = 1
    P2_药品储备定额参照法 = 2
    P3_药品日销售量参照法 = 3
    P4_自定义区间参照法 = 4
End Enum
Public Function GetCondition(FrmMain As Form, ByRef str用途ID, ByRef str剂型 As String, _
    ByRef lng库房ID As Long, ByRef int计划类型 As Integer, ByRef int编制方法 As Integer, _
    ByRef bln下限 As Boolean, ByRef int上限 As Integer, ByRef int下限 As Integer, ByRef bln计划数量 As Boolean, _
    ByRef str供货商ID As String, ByRef bln中标单位 As Boolean, ByRef strBeginDate As String, ByRef strEndDate As String, _
    ByRef bln不考虑库存 As Boolean, ByRef blnClearZeroPlan As Boolean, ByRef blnBaseMedi As Boolean, ByRef intStock As Integer, _
    ByRef bln数量方式 As Boolean, ByRef blnOnlyBaseMedi As Boolean, ByRef strToxicologyClass As String, ByRef bln按销量产生计划 As Boolean, _
    ByRef str来源药房 As String, ByRef str来源库房 As String, Optional ByRef strAll来源药房 As String, Optional ByRef strAll来源库房 As String) As Boolean

    mstr用途ID = ""
    mstr剂型 = ""
    mlng库房ID = 0
    mint计划类型 = 0
    mint编制方法 = 0
    mblnSelect = False
    mblnClearZeroPlan = False
    mblnBaseMedi = False
    mintStock = 0
    
    Set mfrmMain = FrmMain
    Me.Show vbModal, FrmMain
    GetCondition = mblnSelect
    
    bln中标单位 = mbln中标单位
    str供货商ID = mstr供货商ID
    
    str用途ID = mstr用途ID
    str剂型 = mstr剂型
    lng库房ID = mlng库房ID
    int计划类型 = mint计划类型
    int编制方法 = mint编制方法 + 1
    bln下限 = mbln下限
    int上限 = mint上限
    int下限 = mint下限
    bln计划数量 = mbln计划数量
    strBeginDate = mstrBeginDate
    strEndDate = mstrEndDate
    bln不考虑库存 = mbln不考虑库存
    blnClearZeroPlan = mblnClearZeroPlan
    blnBaseMedi = mblnBaseMedi
    intStock = mintStock
    bln数量方式 = mbln数量方式
    blnOnlyBaseMedi = mblnOnlyBaseMedi
    strToxicologyClass = mstrToxicologyClass
    bln按销量产生计划 = mbln按销量产生计划
    str来源药房 = mstr来源药房
    str来源库房 = mstr来源库房
    strAll来源药房 = mstrAll来源药房
    strAll来源库房 = mstrAll来源库房
End Function


Private Sub cmdCancel_Click()
    mblnSelect = False
    Unload Me
End Sub


Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
End Sub

Private Sub cmdOk_Click()
    Dim intItem As Integer, intItems As Integer
    Dim Str期间 As String
    Dim intMonth As Integer
    Dim i As Integer
    Dim bln毒理 As Boolean
        
    If opt方法(3).Value Then
        '库存上限天数不能小于库存下限天数
        '库存上限天数与库存下限天数不能为零
        If Trim(txt上限天数.Text) = "" Then
            MsgBox "请输入库存上限天数！", vbInformation, gstrSysName
            txt上限天数.SetFocus
            Exit Sub
        End If
        If Trim(txt下限天数.Text) = "" Then
            MsgBox "请输入库存下限天数！", vbInformation, gstrSysName
            txt下限天数.SetFocus
            Exit Sub
        End If
        If Not IsNumeric(txt上限天数.Text) Then
            MsgBox "库存上限天数中含有非法字符！", vbInformation, gstrSysName
            txt上限天数.SetFocus
            Exit Sub
        End If
        If Not IsNumeric(txt下限天数.Text) Then
            MsgBox "库存下限天数中含有非法字符！", vbInformation, gstrSysName
            txt下限天数.SetFocus
            Exit Sub
        End If
        If Val(txt上限天数.Text) <= 0 Then
            MsgBox "库存上限天数不能小于零！", vbInformation, gstrSysName
            txt上限天数.SetFocus
            Exit Sub
        End If
        If Val(txt下限天数.Text) <= 0 Then
            MsgBox "库存下限天数不能小于零！", vbInformation, gstrSysName
            txt下限天数.SetFocus
            Exit Sub
        End If
        If Val(txt上限天数.Text) < Val(txt下限天数.Text) Then
            MsgBox "库存上限天数不能小于库存下限天数！", vbInformation, gstrSysName
            txt上限天数.SetFocus
            Exit Sub
        End If
        If Val(txt上限天数.Text) > 300 Then
            MsgBox "库存上限天数不能大于300天！", vbInformation, gstrSysName
            txt上限天数.SetFocus
            Exit Sub
        End If
        mint上限 = Val(txt上限天数.Text)
        mint下限 = Val(txt下限天数.Text)
    End If

    mstr用途ID = ""
    For intItem = 1 To tvw用途.Nodes.count
        If tvw用途.Nodes(intItem).Key = "Root" And tvw用途.Nodes(intItem).Checked Then
            mstr用途ID = "所有分类"
            Exit For
        End If
        
        If tvw用途.Nodes(intItem).Key <> "Root" And _
            tvw用途.Nodes(intItem).Key <> "_中成药" And _
            tvw用途.Nodes(intItem).Key <> "_中草药" And _
            tvw用途.Nodes(intItem).Key <> "_西成药" And _
            tvw用途.Nodes(intItem).Checked Then
            mstr用途ID = mstr用途ID & "," & Mid(tvw用途.Nodes(intItem).Key, 2)
        End If
    Next
    
    If mstr用途ID <> "" And mstr用途ID <> "所有分类" Then
        mstr用途ID = Mid(mstr用途ID, 2)
    End If
    
    mstr剂型 = ""
    
    intItems = Me.Lvw剂型.ListItems.count
    If intItems > 0 Then
        For intItem = 1 To intItems
            If Lvw剂型.ListItems(intItem).Checked Then
                mstr剂型 = mstr剂型 & "," & "'" & Lvw剂型.ListItems(intItem).Text & "'"
            End If
        Next
    End If
    If mstr剂型 <> "" Then mstr剂型 = Mid(mstr剂型, 2)
    
    mlng库房ID = cbo库房.ItemData(cbo库房.ListIndex)
    
    frmDrugPlanCard.LblTitle.Tag = cbo库房.Text

    For intItem = 0 To opt计划.count - 1
       If opt计划(intItem).Value Then
           frmDrugPlanCard.txt计划类型.Caption = Mid(opt计划(intItem).Caption, 1, InStr(1, opt计划(intItem).Caption, "(") - 1)
           mint计划类型 = intItem + 1
           Exit For
       End If
    Next

    For intItem = 0 To opt方法.count - 1
       If opt方法(intItem).Value Then
           frmDrugPlanCard.txt编制方法.Caption = Mid(opt方法(intItem).Caption, 1, InStr(1, opt方法(intItem).Caption, "(") - 1)
           mint编制方法 = intItem
           Exit For
       End If
    Next
    
    mstr供货商ID = ""
    For i = 1 To tvw供货单位.Nodes.count
        If tvw供货单位.Nodes(i).Key <> "Root" And _
            tvw供货单位.Nodes(i).Checked Then
            If tvw供货单位.Nodes(i).Tag = "1" Then
                mstr供货商ID = mstr供货商ID & "," & Mid(tvw供货单位.Nodes(i).Key, 2)
            End If
        End If
    Next
    If mstr供货商ID <> "" Then mstr供货商ID = Mid(mstr供货商ID, 2)
    mbln中标单位 = chk中标单位.Value = 1
    
    mbln下限 = (Chk仅提取低取下限的药品.Value = 1)
    mbln计划数量 = (Chk不产生计划数量.Value <> 1)
    mbln不考虑库存 = (chk不考虑现库存.Value = 1)
    mblnClearZeroPlan = (chkClearZeroPlan.Value = 1)
    mblnBaseMedi = (chkBaseMedi.Value = 1)
    mblnOnlyBaseMedi = (chkOnlyBaseMedi.Value = 1)
    mintStock = IIf(opt常备药(0).Value = True, 0, IIf(opt常备药(1).Value = True, 1, 2))
    mbln按销量产生计划 = (chk参考销量.Value = 1)
    
    If mint编制方法 = zlDrugPlan.P2_药品储备定额参照法 Or mint编制方法 = zlDrugPlan.P4_自定义区间参照法 Then
        mstrBeginDate = Format(dtp开始时间.Value, "yyyy-mm-dd")
        mstrEndDate = Format(dtp结束时间.Value, "yyyy-mm-dd")
    End If
    
    If opt上限.Value = True Then
        mbln数量方式 = False
    Else
        mbln数量方式 = True
    End If
    
    For i = 0 To chk毒理.count - 1
        If chk毒理(i).Value = 0 Then
            bln毒理 = True
            Exit For
        End If
    Next
    
    mstrToxicologyClass = ""
    If bln毒理 = True Then
        If chk毒理(0).Value = 1 Then
            mstrToxicologyClass = " t.毒理分类='普通药'"
        End If
        If chk毒理(1).Value = 1 Then
            If mstrToxicologyClass = "" Then
                mstrToxicologyClass = " t.毒理分类='毒性药'"
            Else
                mstrToxicologyClass = mstrToxicologyClass & " or t.毒理分类='毒性药'"
            End If
        End If
        If chk毒理(2).Value = 1 Then
            If mstrToxicologyClass = "" Then
                mstrToxicologyClass = "  t.毒理分类 ='精神I类' or t.毒理分类 ='精神II类' "
            Else
                mstrToxicologyClass = mstrToxicologyClass & " or t.毒理分类 ='精神I类' or t.毒理分类 ='精神II类'"
            End If
        End If
        If chk毒理(3).Value = 1 Then
            If mstrToxicologyClass = "" Then
                mstrToxicologyClass = " t.毒理分类 ='麻醉药'"
            Else
                mstrToxicologyClass = mstrToxicologyClass & " or t.毒理分类 ='麻醉药'"
            End If
        End If
        
        If mstrToxicologyClass <> "" Then
            mstrToxicologyClass = "(" & mstrToxicologyClass & ")"
        End If
    End If
    
    mstr来源药房 = ""
    intItems = Me.lvw药房.ListItems.count
    If intItems > 0 Then
        For intItem = 1 To intItems
            If lvw药房.ListItems(intItem).Checked Then
                mstr来源药房 = IIf(mstr来源药房 = "", "", mstr来源药房 & ",") & Mid(lvw药房.ListItems(intItem).Key, 2)
            End If
        Next
    End If
    
    mstr来源库房 = ""
    intItems = Me.lvw库房.ListItems.count
    If intItems > 0 Then
        For intItem = 1 To intItems
            If lvw库房.ListItems(intItem).Checked Then
                mstr来源库房 = IIf(mstr来源库房 = "", "", mstr来源库房 & ",") & Mid(lvw库房.ListItems(intItem).Key, 2)
            End If
        Next
    End If
    
    If mstr来源库房 <> "" Then
        If InStr(1, "," & mstr来源库房 & ",", "," & cbo库房.ItemData(cbo库房.ListIndex) & ",") = 0 Then
            mstr来源库房 = mstr来源库房 & "," & cbo库房.ItemData(cbo库房.ListIndex)
        End If
    End If
    
    '如果没有选择药品分类，提示是否继续
    If mstr用途ID = "" Then
        If MsgBox("未选择药品分类，将产生空的计划，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    End If
    
    mblnSelect = True
    Unload Me
End Sub

Private Sub cbo库房_Click()
    Dim blnEXIST As Boolean
    Dim rsTemp As New ADODB.Recordset
    '如果是全院计划，提取所有药品剂型
    On Error GoTo errHandle
    If Me.cbo库房.ItemData(Me.cbo库房.ListIndex) = 0 Then
        gstrSQL = "Select 编码,名称 From 药品剂型 Order by 编码"
        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "提取所有药品剂型")
        
'        opt常备药(0).Enabled = False
'        opt常备药(1).Enabled = False
    Else
        '提取该库房现有剂型，供用户选择
        mbln中药库房 = False
        gstrSQL = "Select 1 From 部门性质说明 " & _
                 " Where 工作性质 Like '中药%' And 部门ID=[1] "
        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "[检查部门性质]", Me.cbo库房.ItemData(cbo库房.ListIndex))
        
        If Not rsTemp.EOF Then mbln中药库房 = True
    
        gstrSQL = "Select Distinct J.编码,J.名称 " & _
                 " From 诊疗执行科室 A,药品特性 B,药品剂型 J " & _
                 " Where A.诊疗项目ID=B.药名ID And B.药品剂型=J.名称" & _
                 " And A.执行科室ID=[1]"
        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "[提取该库房现在剂型]", Me.cbo库房.ItemData(cbo库房.ListIndex))
    End If

    Lvw剂型.ListItems.Clear
    With rsTemp
        Do While Not .EOF
            If blnEXIST = False Then
                blnEXIST = (!名称 = "方剂")
            End If
            Lvw剂型.ListItems.Add , "K" & !编码, !名称, , 1
            .MoveNext
        Loop
        If mbln中药库房 And blnEXIST = False Then
            Lvw剂型.ListItems.Add , "KK1", "方剂", , 1
        End If
    End With
    If Chk剂型.Value <> 2 Then
        Chk剂型_Click
    Else
        Chk剂型.Value = 0
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey (vbKeyTab)
End Sub

Private Sub Form_Load()
    Dim rsTmp As New ADODB.Recordset
    Dim objNode As Node
    Dim i As Integer
    Dim blnSelectStock As String
    Dim strIco As String, strID As String
    Dim strTemp As String
    Dim objItem As ListItem
    
    On Error GoTo errH

    blnSelectStock = GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\药品计划管理", "库房", "0")
    mintPlanPoint = Val(zlDataBase.GetPara("全院计划不管站点", glngSys, 1330, 0))
    
    sstConditon.Tab = 0
    
    With mfrmMain.cboStock
        cbo库房.Clear
        For i = 0 To .ListCount - 1
            cbo库房.AddItem .List(i)
            cbo库房.ItemData(cbo库房.NewIndex) = .ItemData(i)
        Next
        cbo库房.ListIndex = .ListIndex
    End With

    If zlStr.IsHavePrivs(gstrprivs, "所有库房") Then
        If blnSelectStock = "0" Then
            cbo库房.Enabled = False
        Else
            cbo库房.Enabled = True
        End If
    Else
        cbo库房.Enabled = False
    End If
    
    '用途
    gstrSQL = "Select Level as 层,ID,上级ID,名称,DECODE(类型,1,'西成药',2,'中成药','中草药') As 材质 " & _
        " From 诊疗分类目录" & _
        " Where 类型 in (1,2,3)" & _
        " Start With 上级ID is NULL Connect by Prior ID=上级ID" & _
        " Order by Level"
    Set rsTmp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption)

    Set objNode = tvw用途.Nodes.Add(, , "Root", "所有用途", "Item")
    Set objNode = tvw用途.Nodes.Add("Root", 4, "_西成药", "西成药", "Item")
    Set objNode = tvw用途.Nodes.Add("Root", 4, "_中草药", "中草药", "Item")
    Set objNode = tvw用途.Nodes.Add("Root", 4, "_中成药", "中成药", "Item")

    Do While Not rsTmp.EOF
        If rsTmp!层 = 1 Then
            Set objNode = tvw用途.Nodes.Add("_" & rsTmp!材质, 4, "_" & rsTmp!Id, rsTmp!名称, "Item")
        Else
            Set objNode = tvw用途.Nodes.Add("_" & rsTmp!上级ID, 4, "_" & rsTmp!Id, rsTmp!名称, "Item")
        End If
        rsTmp.MoveNext
    Loop
    tvw用途.Nodes("Root").Selected = True
    tvw用途.Nodes("Root").Expanded = True
    '0-要管站点，1-不管站点
    mlng库房ID = cbo库房.ItemData(cbo库房.ListIndex)
    If mlng库房ID <> 0 Or (mlng库房ID = 0 And mintPlanPoint = 0 And (gstrNodeNo <> "-" Or gstrNodeNo <> "0")) Then
        strTemp = "(站点 = [1] Or 站点 is Null) And "
    End If
    gstrSQL = "" & _
        "   Select Level as 层,ID,上级ID,编码||'-'||名称 名称,末级 " & _
        "   From 供应商" & _
        "   where " & strTemp & "(substr(类型,1,1)=1 Or Nvl(末级,0)=0) And (To_Char(撤档时间,'yyyy-MM-dd')='3000-01-01' or 撤档时间 is null)" & _
        "   Start With 上级ID is NULL Connect by Prior ID=上级ID" & _
        "   Order by Level"
    Set rsTmp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "-供应商", gstrNodeNo)
    
    tvw供货单位.Nodes.Clear
    Set objNode = tvw供货单位.Nodes.Add(, , "Root", "所有药品供货商", "Folder")
    objNode.Sorted = True
    Do While Not rsTmp.EOF
        strIco = IIf(Val(NVL(rsTmp!末级)) = 1, "Card", "Folder")
        If rsTmp!层 = 1 Then
            Set objNode = tvw供货单位.Nodes.Add("Root", 4, "_" & rsTmp!Id, rsTmp!名称, strIco)
            strID = strID & rsTmp!Id & ";"
        Else
            If InStr(strID, rsTmp!Id & ";") = 0 Then
                Set objNode = tvw供货单位.Nodes.Add("Root", 4, "_" & rsTmp!Id, rsTmp!名称, strIco)
            Else
                Set objNode = tvw供货单位.Nodes.Add("_" & rsTmp!上级ID, 4, "_" & rsTmp!Id, rsTmp!名称, strIco)
            End If
        End If
        If strIco = "Card" Then
            objNode.Tag = "1"
        End If
        objNode.Sorted = True
        rsTmp.MoveNext
    Loop
    tvw供货单位.Nodes("Root").Selected = True
    tvw供货单位.Nodes("Root").Expanded = True
    
    Me.dtp结束时间 = Sys.Currentdate
    Me.dtp开始时间 = DateAdd("m", -1, Me.dtp结束时间)
    fra方式.Visible = False
    opt上限.Value = True    '默认是按上限
    
    mint编制方法 = Val(GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "药品计划管理", "条件设置编辑方式", 0))
    If mint编制方法 >= 0 And mint编制方法 <= 4 Then
        opt方法(mint编制方法).Value = True
    Else
        opt方法(0).Value = True
    End If
    
    '来源药房
    mstr来源药房 = GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "药品计划管理", "来源药房", "")
    gstrSQL = "Select Distinct a.id,a.编码,a.名称 " & _
        " From 部门表 a,部门性质说明 b " & _
        " Where a.id=b.部门id And b.工作性质  In ('中药房','西药房','成药房') And TO_CHAR(a.撤档时间, 'yyyy-MM-dd') = '3000-01-01'  Order By 名称 "
    Set rsTmp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "-药房")
    
    lvw药房.ListItems.Clear
    mstrAll来源药房 = ""
    With rsTmp
        Do While Not .EOF
            Set objItem = lvw药房.ListItems.Add(, "K" & !Id, "[" & !编码 & "]" & !名称, , 1)
                        
            If InStr(1, "," & mstr来源药房 & ",", "," & !Id & ",") > 0 Then
                objItem.Checked = True
            End If
            
            mstrAll来源药房 = IIf(mstrAll来源药房 = "", "", mstrAll来源药房 & ",") & !Id
            
            .MoveNext
        Loop
    End With
    
    '来源库房
    mstr来源库房 = GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "药品计划管理", "来源库房", "")
    gstrSQL = "Select Distinct a.id,a.编码,a.名称 " & _
        " From 部门表 a,部门性质说明 b " & _
        " Where a.id=b.部门id And b.工作性质  In ('中药房','西药房','成药房','西药库', '成药库', '中药库') And TO_CHAR(a.撤档时间, 'yyyy-MM-dd') = '3000-01-01'  Order By 名称 "
    Set rsTmp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "-药房")
    
    lvw库房.ListItems.Clear
    mstrAll来源库房 = ""
    With rsTmp
        Do While Not .EOF
            Set objItem = lvw库房.ListItems.Add(, "K" & !Id, "[" & !编码 & "]" & !名称, , 1)
                        
            If InStr(1, "," & mstr来源库房 & ",", "," & !Id & ",") > 0 Then
                objItem.Checked = True
            End If
            
            mstrAll来源库房 = IIf(mstrAll来源库房 = "", "", mstrAll来源库房 & ",") & !Id
            
            .MoveNext
        Loop
    End With
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\药品计划管理", "条件设置编辑方式", mint编制方法)
    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\药品计划管理", "来源药房", mstr来源药房)
    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\药品计划管理", "来源库房", mstr来源库房)
End Sub


Private Sub opt方法_Click(index As Integer)
    fra辅助条件.Visible = False
    fra计划类型.Visible = False
    fra区间.Visible = False
    chk参考销量.Enabled = True
    
    Select Case index
    Case zlDrugPlan.P0_往年同期线性参照法, zlDrugPlan.P1_临近期间平均参照法
        '0-往年同期线性参照法；1-临近期间平均参照法
        fra计划类型.Visible = True
        fra辅助条件.Visible = True
        fra辅助条件.Enabled = False
        fra辅助条件.ZOrder 0
    Case zlDrugPlan.P2_药品储备定额参照法
        '药品储备定额参照法
        '显示区间，产生数量方式
        chk参考销量.Value = 0
        chk参考销量.Enabled = False
        fra区间.Visible = True
        fra方式.Visible = True
        fra方式.ZOrder 0
    Case zlDrugPlan.P3_药品日销售量参照法
        '药品日销售量参照法
        fra计划类型.Visible = True
        fra辅助条件.Visible = True
        fra辅助条件.Enabled = True
        fra辅助条件.ZOrder 0
    Case zlDrugPlan.P4_自定义区间参照法
        '自定义区间参照法
        fra区间.Visible = True
        fra辅助条件.Visible = True
        fra辅助条件.Enabled = False
        fra辅助条件.ZOrder 0
    End Select
    
    mint编制方法 = index

End Sub

Private Sub opt计划_Click(index As Integer)
    opt方法(0).Enabled = True
    opt方法(1).Enabled = True
    opt方法(2).Enabled = True
    opt方法(3).Enabled = True

    Select Case index
    Case 1
        If opt方法(3).Value Then
            opt方法(3).Value = False
            opt方法(0).Value = True
        End If
        opt方法(3).Enabled = False
    Case 2
        If opt方法(0).Value Or opt方法(3).Value Then
            opt方法(0).Value = False
            opt方法(3).Value = False
            opt方法(1).Value = True
        End If
        opt方法(0).Enabled = False
        opt方法(3).Enabled = False
    Case 3
        If opt方法(0).Value Then
            opt方法(0).Value = False
            opt方法(1).Value = True
        End If
        opt方法(0).Enabled = False
    End Select
End Sub



Private Sub tvw供货单位_NodeCheck(ByVal Node As MSComctlLib.Node)
    Dim blnAllUnCheck As Boolean
    
    CheckNode Node, Node.Checked
    SetParentNode Node, Node.Checked, False
    
    blnAllUnCheck = True
    Do While Not Node Is Nothing
        If Node.Checked = True Then
            blnAllUnCheck = False
            Exit Do
        End If
        Set Node = Node.Next
    Loop
    
    If blnAllUnCheck Then
        chk中标单位.Value = 0
        chk中标单位.Enabled = False
    ElseIf chk中标单位.Enabled = False Then
        chk中标单位.Enabled = True
    End If
End Sub

Private Sub tvw用途_NodeCheck(ByVal Node As MSComctlLib.Node)
    CheckNode Node, Node.Checked
    SetParentNode Node, Node.Checked
End Sub

Private Sub SetParentNode1(ByVal Node As MSComctlLib.Node, blnCheck As Boolean)
    Dim intIdx As Integer

    If Not Node.Parent Is Nothing Then
        If blnCheck = True Then
            '看是否他的兄弟接点是否也全是TRUE，如是，则置其父节点也为TRUE，否则，不管
            intIdx = Node.FirstSibling.index
            Do While intIdx <> Node.LastSibling.index
                If tvw用途.Nodes(intIdx).Checked = False Then
                    Node.Parent.Checked = False
                    Exit Do
                End If
                intIdx = tvw用途.Nodes(intIdx).Next.index
            Loop
            If intIdx = Node.LastSibling.index Then
                If tvw用途.Nodes(intIdx).Checked = True Then
                    Node.Parent.Checked = True
                End If
            End If
        Else
            Node.Parent.Checked = False
        End If

        Set Node = Node.Parent
        If Not Node Is Nothing Then
            SetParentNode Node, blnCheck
        End If
    End If
End Sub


Private Sub SetParentNode(ByVal Node As MSComctlLib.Node, blnCheck As Boolean, Optional blnTvw用途 As Boolean = True)
    Dim intIdx As Integer
    
    If Not Node.Parent Is Nothing Then
        If blnCheck = True Then
            '看是否他的兄弟接点是否也全是TRUE，如是，则置其父节点也为TRUE，否则，不管
            intIdx = Node.FirstSibling.index
            Do While intIdx <> Node.LastSibling.index
                If blnTvw用途 = True Then
                    If tvw用途.Nodes(intIdx).Checked = False Then
                        Node.Parent.Checked = False
                        Exit Do
                    End If
                    intIdx = tvw用途.Nodes(intIdx).Next.index
                Else
                    If tvw供货单位.Nodes(intIdx).Checked = False Then
                        Node.Parent.Checked = False
                        Exit Do
                    End If
                    intIdx = tvw供货单位.Nodes(intIdx).Next.index
                End If
            Loop
            If intIdx = Node.LastSibling.index Then
                If blnTvw用途 = True Then
                       If tvw用途.Nodes(intIdx).Checked = True Then
                           Node.Parent.Checked = True
                       End If
                Else
                       If tvw供货单位.Nodes(intIdx).Checked = True Then
                           Node.Parent.Checked = True
                       End If
                End If
            End If
        Else
            Node.Parent.Checked = False
        End If
        
        Set Node = Node.Parent
        If Not Node Is Nothing Then
            SetParentNode Node, blnCheck, blnTvw用途
        End If
    End If
End Sub
Private Function CheckNode(ByVal Node As Object, blnCheck As Boolean)
    Dim intIdx As Integer

    If Node.Children > 0 Then
        Set Node = Node.Child
        Do While Not Node Is Nothing
            Node.Checked = blnCheck
            If Node.Children > 0 Then
                CheckNode Node, blnCheck
            End If
            Set Node = Node.Next
        Loop
    Else
        Node.Checked = blnCheck
    End If
End Function

Private Function CheckCount() As Integer
    Dim i As Integer
    For i = 1 To tvw用途.Nodes.count
        If tvw用途.Nodes(i).Checked Then CheckCount = CheckCount + 1
    Next
End Function


Private Sub Chk剂型_Click()
    If Chk剂型.Value = 2 Then Exit Sub
    Call SetSelect(Lvw剂型, Chk剂型.Value)
End Sub

Private Sub Lvw剂型_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Call ItemCheck(Lvw剂型, Item)
End Sub

Private Sub SetSelect(ByVal lvwObj As Object, Optional ByVal BlnSelect As Boolean = True)
    Dim intSelect As Integer
    With lvwObj
        For intSelect = 1 To .ListItems.count
            .ListItems(intSelect).Checked = BlnSelect
        Next
    End With
End Sub

Private Sub ItemCheck(ByVal lvwObj As Object, ByVal Item As MSComctlLib.ListItem)
    Dim lngCheck As Long, blnCheck As Boolean, intCount As Integer
    
    intCount = 0
    With lvwObj
        For lngCheck = 1 To .ListItems.count
            If .ListItems(lngCheck).Checked = True Then
                intCount = intCount + 1
            End If
        Next
        
        If intCount = lvwObj.ListItems.count Then
            Chk剂型.Value = 1
        ElseIf intCount > 0 Then
            Chk剂型.Value = 2
        Else
            Chk剂型.Value = 0
        End If
    End With
End Sub
