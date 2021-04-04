VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmParClinic 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "临床参数设置"
   ClientHeight    =   8790
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12630
   Icon            =   "frmParClinic.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8790
   ScaleWidth      =   12630
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picPar 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   7455
      Index           =   12
      Left            =   2340
      ScaleHeight     =   7425
      ScaleWidth      =   10185
      TabIndex        =   273
      Tag             =   "j"
      Top             =   0
      Width           =   10215
      Begin TabDlg.SSTab StabNurse 
         Height          =   7245
         Left            =   105
         TabIndex        =   274
         Top             =   90
         Width           =   9945
         _ExtentX        =   17542
         _ExtentY        =   12779
         _Version        =   393216
         TabHeight       =   520
         TabCaption(0)   =   "记录单"
         TabPicture(0)   =   "frmParClinic.frx":6852
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fra(1)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "体温单 "
         TabPicture(1)   =   "frmParClinic.frx":686E
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "lbl(22)"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "lbl(23)"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "lbl(24)"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "lbl(30)"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "lbl(31)"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).Control(5)=   "lbl(32)"
         Tab(1).Control(5).Enabled=   0   'False
         Tab(1).Control(6)=   "lblBreathe"
         Tab(1).Control(6).Enabled=   0   'False
         Tab(1).Control(7)=   "lbl(33)"
         Tab(1).Control(7).Enabled=   0   'False
         Tab(1).Control(8)=   "lbl(34)"
         Tab(1).Control(8).Enabled=   0   'False
         Tab(1).Control(9)=   "lbl(35)"
         Tab(1).Control(9).Enabled=   0   'False
         Tab(1).Control(10)=   "Line2"
         Tab(1).Control(10).Enabled=   0   'False
         Tab(1).Control(11)=   "lbl(36)"
         Tab(1).Control(11).Enabled=   0   'False
         Tab(1).Control(12)=   "lbl(37)"
         Tab(1).Control(12).Enabled=   0   'False
         Tab(1).Control(13)=   "lbl(38)"
         Tab(1).Control(13).Enabled=   0   'False
         Tab(1).Control(14)=   "lbl(39)"
         Tab(1).Control(14).Enabled=   0   'False
         Tab(1).Control(15)=   "lbl(27)"
         Tab(1).Control(15).Enabled=   0   'False
         Tab(1).Control(16)=   "Line1"
         Tab(1).Control(16).Enabled=   0   'False
         Tab(1).Control(17)=   "lbl(40)"
         Tab(1).Control(17).Enabled=   0   'False
         Tab(1).Control(18)=   "lbl(41)"
         Tab(1).Control(18).Enabled=   0   'False
         Tab(1).Control(19)=   "chk(127)"
         Tab(1).Control(19).Enabled=   0   'False
         Tab(1).Control(20)=   "picControl(1)"
         Tab(1).Control(20).Enabled=   0   'False
         Tab(1).Control(21)=   "ud(22)"
         Tab(1).Control(21).Enabled=   0   'False
         Tab(1).Control(22)=   "ud(23)"
         Tab(1).Control(22).Enabled=   0   'False
         Tab(1).Control(23)=   "ud(24)"
         Tab(1).Control(23).Enabled=   0   'False
         Tab(1).Control(24)=   "ud(26)"
         Tab(1).Control(24).Enabled=   0   'False
         Tab(1).Control(25)=   "ud(27)"
         Tab(1).Control(25).Enabled=   0   'False
         Tab(1).Control(26)=   "fra(15)"
         Tab(1).Control(26).Enabled=   0   'False
         Tab(1).Control(27)=   "txtUD(22)"
         Tab(1).Control(27).Enabled=   0   'False
         Tab(1).Control(28)=   "txtUD(23)"
         Tab(1).Control(28).Enabled=   0   'False
         Tab(1).Control(29)=   "txtUD(24)"
         Tab(1).Control(29).Enabled=   0   'False
         Tab(1).Control(30)=   "txtUD(26)"
         Tab(1).Control(30).Enabled=   0   'False
         Tab(1).Control(31)=   "txtUD(27)"
         Tab(1).Control(31).Enabled=   0   'False
         Tab(1).Control(32)=   "cbo(18)"
         Tab(1).Control(32).Enabled=   0   'False
         Tab(1).Control(33)=   "cbo(17)"
         Tab(1).Control(33).Enabled=   0   'False
         Tab(1).Control(34)=   "picLineColor(2)"
         Tab(1).Control(34).Enabled=   0   'False
         Tab(1).Control(35)=   "cbo(15)"
         Tab(1).Control(35).Enabled=   0   'False
         Tab(1).Control(36)=   "cbo(16)"
         Tab(1).Control(36).Enabled=   0   'False
         Tab(1).Control(37)=   "txt(28)"
         Tab(1).Control(37).Enabled=   0   'False
         Tab(1).Control(38)=   "chk(121)"
         Tab(1).Control(38).Enabled=   0   'False
         Tab(1).Control(39)=   "chk(122)"
         Tab(1).Control(39).Enabled=   0   'False
         Tab(1).Control(40)=   "chk(123)"
         Tab(1).Control(40).Enabled=   0   'False
         Tab(1).Control(41)=   "chk(124)"
         Tab(1).Control(41).Enabled=   0   'False
         Tab(1).Control(42)=   "chk(125)"
         Tab(1).Control(42).Enabled=   0   'False
         Tab(1).Control(43)=   "chk(126)"
         Tab(1).Control(43).Enabled=   0   'False
         Tab(1).Control(44)=   "chk(128)"
         Tab(1).Control(44).Enabled=   0   'False
         Tab(1).Control(45)=   "Frame1(4)"
         Tab(1).Control(45).Enabled=   0   'False
         Tab(1).Control(46)=   "picPoly"
         Tab(1).Control(46).Enabled=   0   'False
         Tab(1).Control(47)=   "PicInsert"
         Tab(1).Control(47).Enabled=   0   'False
         Tab(1).Control(48)=   "chk(130)"
         Tab(1).Control(48).Enabled=   0   'False
         Tab(1).Control(49)=   "chk(131)"
         Tab(1).Control(49).Enabled=   0   'False
         Tab(1).Control(50)=   "picLineColor(3)"
         Tab(1).Control(50).Enabled=   0   'False
         Tab(1).Control(51)=   "cbo(19)"
         Tab(1).Control(51).Enabled=   0   'False
         Tab(1).Control(52)=   "txt(25)"
         Tab(1).Control(52).Enabled=   0   'False
         Tab(1).Control(53)=   "chk(129)"
         Tab(1).Control(53).Enabled=   0   'False
         Tab(1).Control(54)=   "picLineColor(1)"
         Tab(1).Control(54).Enabled=   0   'False
         Tab(1).Control(55)=   "picOut"
         Tab(1).Control(55).Enabled=   0   'False
         Tab(1).Control(56)=   "picEnemaStool"
         Tab(1).Control(56).Enabled=   0   'False
         Tab(1).Control(57)=   "chk(132)"
         Tab(1).Control(57).Enabled=   0   'False
         Tab(1).Control(58)=   "picLineColor(4)"
         Tab(1).Control(58).Enabled=   0   'False
         Tab(1).Control(59)=   "chk(180)"
         Tab(1).Control(59).Enabled=   0   'False
         Tab(1).ControlCount=   60
         TabCaption(2)   =   "产程图"
         TabPicture(2)   =   "frmParClinic.frx":688A
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "lbl(11)"
         Tab(2).Control(1)=   "fra1(1)"
         Tab(2).Control(2)=   "fra1(0)"
         Tab(2).Control(3)=   "fra2"
         Tab(2).Control(4)=   "cbo(6)"
         Tab(2).Control(5)=   "chk(109)"
         Tab(2).Control(6)=   "chk(110)"
         Tab(2).Control(7)=   "chk(111)"
         Tab(2).Control(8)=   "fraPoliceLine"
         Tab(2).ControlCount=   9
         Begin VB.CheckBox chk 
            Caption         =   "曲线项目骑线显示"
            Height          =   180
            Index           =   180
            Left            =   -74685
            TabIndex        =   339
            Top             =   5805
            Width           =   2055
         End
         Begin VB.PictureBox picLineColor 
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            ForeColor       =   &H80000008&
            Height          =   200
            Index           =   4
            Left            =   -72045
            ScaleHeight     =   165
            ScaleWidth      =   270
            TabIndex        =   324
            TabStop         =   0   'False
            Top             =   2760
            Width           =   300
         End
         Begin VB.CheckBox chk 
            Caption         =   "体温单打印时,不输出下方的曲线说明信息"
            Height          =   195
            Index           =   132
            Left            =   -74685
            TabIndex        =   336
            Top             =   4950
            Width           =   3840
         End
         Begin VB.PictureBox picEnemaStool 
            BorderStyle     =   0  'None
            Height          =   190
            Left            =   -73140
            ScaleHeight     =   195
            ScaleWidth      =   2565
            TabIndex        =   354
            TabStop         =   0   'False
            Top             =   6870
            Width           =   2565
            Begin VB.OptionButton OptEnemaStool 
               Caption         =   "分子分母"
               Height          =   180
               Index           =   1
               Left            =   1230
               TabIndex        =   356
               Top             =   5
               Width           =   1050
            End
            Begin VB.OptionButton OptEnemaStool 
               Caption         =   "上下标"
               Height          =   180
               Index           =   0
               Left            =   15
               TabIndex        =   355
               Top             =   5
               Value           =   -1  'True
               Width           =   960
            End
         End
         Begin VB.PictureBox picOut 
            BorderStyle     =   0  'None
            Height          =   190
            Left            =   -73140
            ScaleHeight     =   195
            ScaleWidth      =   2565
            TabIndex        =   350
            TabStop         =   0   'False
            Top             =   6615
            Width           =   2565
            Begin VB.OptionButton OptOut 
               Caption         =   "医嘱内容"
               Height          =   180
               Index           =   1
               Left            =   1230
               TabIndex        =   352
               Top             =   5
               Width           =   1170
            End
            Begin VB.OptionButton OptOut 
               Caption         =   "出院方式"
               Height          =   180
               Index           =   0
               Left            =   0
               TabIndex        =   351
               Top             =   5
               Value           =   -1  'True
               Width           =   1170
            End
         End
         Begin VB.PictureBox picLineColor 
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            ForeColor       =   &H80000008&
            Height          =   200
            Index           =   1
            Left            =   -66525
            ScaleHeight     =   165
            ScaleWidth      =   270
            TabIndex        =   361
            TabStop         =   0   'False
            Top             =   1365
            Width           =   300
         End
         Begin VB.CheckBox chk 
            Caption         =   "手术后标注天数内再次手术时,停止前一次手术标注"
            Height          =   180
            Index           =   129
            Left            =   -74700
            TabIndex        =   330
            Top             =   3330
            Width           =   4500
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   180
            Index           =   25
            Left            =   -73740
            MaxLength       =   2
            TabIndex        =   326
            Text            =   "14"
            Top             =   3045
            Width           =   255
         End
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   19
            ItemData        =   "frmParClinic.frx":68A6
            Left            =   -71175
            List            =   "frmParClinic.frx":68A8
            Style           =   2  'Dropdown List
            TabIndex        =   329
            Top             =   2985
            Width           =   1725
         End
         Begin VB.PictureBox picLineColor 
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            ForeColor       =   &H80000008&
            Height          =   200
            Index           =   3
            Left            =   -72765
            ScaleHeight     =   165
            ScaleWidth      =   270
            TabIndex        =   327
            TabStop         =   0   'False
            Top             =   3045
            Width           =   300
         End
         Begin VB.CheckBox chk 
            Caption         =   "婴儿体温单首日天数从0开始"
            Height          =   180
            Index           =   131
            Left            =   -68325
            TabIndex        =   359
            Top             =   1110
            Width           =   2595
         End
         Begin VB.CheckBox chk 
            Caption         =   "体温单输出时打印医院名称"
            Height          =   180
            Index           =   130
            Left            =   -68325
            TabIndex        =   358
            Top             =   840
            Width           =   2550
         End
         Begin VB.PictureBox PicInsert 
            BorderStyle     =   0  'None
            Height          =   190
            Left            =   -73140
            ScaleHeight     =   195
            ScaleWidth      =   2565
            TabIndex        =   346
            TabStop         =   0   'False
            Top             =   6360
            Width           =   2565
            Begin VB.OptionButton OptInsert 
               Caption         =   "脉搏/心率"
               Height          =   180
               Index           =   0
               Left            =   0
               TabIndex        =   347
               Top             =   5
               Value           =   -1  'True
               Width           =   1170
            End
            Begin VB.OptionButton OptInsert 
               Caption         =   "心率/脉搏"
               Height          =   180
               Index           =   1
               Left            =   1230
               TabIndex        =   348
               Top             =   5
               Width           =   1170
            End
         End
         Begin VB.PictureBox picPoly 
            BorderStyle     =   0  'None
            Height          =   190
            Left            =   -73140
            ScaleHeight     =   195
            ScaleWidth      =   3360
            TabIndex        =   341
            TabStop         =   0   'False
            Top             =   6105
            Width           =   3360
            Begin VB.OptionButton optPloy 
               Caption         =   "斜线"
               Height          =   180
               Index           =   0
               Left            =   0
               TabIndex        =   342
               Top             =   5
               Value           =   -1  'True
               Width           =   720
            End
            Begin VB.OptionButton optPloy 
               Caption         =   "直线"
               Height          =   180
               Index           =   1
               Left            =   1230
               TabIndex        =   343
               Top             =   5
               Width           =   720
            End
            Begin VB.OptionButton optPloy 
               Caption         =   "不填充"
               Height          =   180
               Index           =   2
               Left            =   2400
               TabIndex        =   344
               Top             =   5
               Width           =   840
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "体温单文件的开始时间"
            Height          =   660
            Index           =   4
            Left            =   -68325
            TabIndex        =   387
            Top             =   6390
            Width           =   2970
            Begin VB.OptionButton optFileTime 
               Caption         =   "入科时间"
               Height          =   195
               Index           =   1
               Left            =   1680
               TabIndex        =   389
               Top             =   300
               Width           =   1125
            End
            Begin VB.OptionButton optFileTime 
               Caption         =   "入院时间"
               Height          =   195
               Index           =   0
               Left            =   300
               TabIndex        =   388
               Top             =   300
               Value           =   -1  'True
               Width           =   1125
            End
         End
         Begin VB.CheckBox chk 
            Caption         =   "体温单上显示病人的诊断信息"
            Height          =   180
            Index           =   128
            Left            =   -74685
            TabIndex        =   338
            Top             =   5520
            Width           =   2790
         End
         Begin VB.CheckBox chk 
            Caption         =   "体温单打印时,不输出心率列(仅在心率单独使用有效)"
            Height          =   180
            Index           =   126
            Left            =   -74685
            TabIndex        =   335
            Top             =   4680
            Width           =   4770
         End
         Begin VB.CheckBox chk 
            Caption         =   "汇总、波动项目显示当天数据（不勾显示昨天）"
            Height          =   180
            Index           =   125
            Left            =   -74685
            TabIndex        =   332
            Top             =   3870
            Width           =   4215
         End
         Begin VB.CheckBox chk 
            Caption         =   "全天汇总允许录入、显示汇总时间(h)"
            Height          =   180
            Index           =   124
            Left            =   -74685
            TabIndex        =   333
            Top             =   4140
            Width           =   3330
         End
         Begin VB.CheckBox chk 
            Caption         =   "病人术后不足标注天数出院,手术标记当页显示完全"
            Height          =   180
            Index           =   123
            Left            =   -74685
            TabIndex        =   331
            Top             =   3600
            Width           =   4500
         End
         Begin VB.CheckBox chk 
            Caption         =   "体温单只显示入科标识时，不自动转换为入院"
            Height          =   180
            Index           =   122
            Left            =   -74685
            TabIndex        =   337
            Top             =   5235
            Width           =   4020
         End
         Begin VB.CheckBox chk 
            Caption         =   "体温单日期栏每页首列日期格式固定为:年-月-日"
            Height          =   180
            Index           =   121
            Left            =   -74700
            TabIndex        =   334
            Top             =   4410
            Width           =   4365
         End
         Begin VB.TextBox txt 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   180
            Index           =   28
            Left            =   -72840
            MaxLength       =   2
            TabIndex        =   323
            Text            =   "v"
            Top             =   2745
            Width           =   255
         End
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   16
            ItemData        =   "frmParClinic.frx":68AA
            Left            =   -73215
            List            =   "frmParClinic.frx":68AC
            Style           =   2  'Dropdown List
            TabIndex        =   317
            Top             =   1800
            Width           =   1590
         End
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   15
            ItemData        =   "frmParClinic.frx":68AE
            Left            =   -73215
            List            =   "frmParClinic.frx":68B0
            Style           =   2  'Dropdown List
            TabIndex        =   313
            Top             =   1470
            Width           =   1590
         End
         Begin VB.PictureBox picLineColor 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            ForeColor       =   &H80000008&
            Height          =   200
            Index           =   2
            Left            =   -71070
            ScaleHeight     =   165
            ScaleWidth      =   270
            TabIndex        =   315
            TabStop         =   0   'False
            Top             =   1500
            Width           =   300
         End
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   17
            ItemData        =   "frmParClinic.frx":68B2
            Left            =   -71490
            List            =   "frmParClinic.frx":68B4
            Style           =   2  'Dropdown List
            TabIndex        =   321
            Top             =   2445
            Width           =   2430
         End
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   18
            ItemData        =   "frmParClinic.frx":68B6
            Left            =   -71490
            List            =   "frmParClinic.frx":68B8
            Style           =   2  'Dropdown List
            TabIndex        =   319
            Top             =   2115
            Width           =   2430
         End
         Begin VB.TextBox txtUD 
            Alignment       =   2  'Center
            Height          =   270
            IMEMode         =   3  'DISABLE
            Index           =   27
            Left            =   -73215
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   298
            Text            =   "0"
            Top             =   525
            Width           =   350
         End
         Begin VB.TextBox txtUD 
            Alignment       =   2  'Center
            Height          =   270
            IMEMode         =   3  'DISABLE
            Index           =   26
            Left            =   -72390
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   307
            Text            =   "0"
            Top             =   840
            Width           =   350
         End
         Begin VB.TextBox txtUD 
            Alignment       =   2  'Center
            Height          =   270
            Index           =   24
            Left            =   -70650
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   304
            Text            =   "6"
            Top             =   525
            Width           =   350
         End
         Begin VB.TextBox txtUD 
            Alignment       =   2  'Center
            Height          =   270
            Index           =   23
            Left            =   -72000
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   301
            Text            =   "18"
            Top             =   525
            Width           =   350
         End
         Begin VB.TextBox txtUD 
            Alignment       =   2  'Center
            Height          =   270
            IMEMode         =   3  'DISABLE
            Index           =   22
            Left            =   -72390
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   310
            Text            =   "0"
            Top             =   1155
            Width           =   350
         End
         Begin VB.Frame fra 
            Caption         =   "体温自动标志--颜色    "
            Height          =   4845
            Index           =   15
            Left            =   -68325
            TabIndex        =   360
            Top             =   1395
            Width           =   2970
            Begin VB.ComboBox cbo 
               Height          =   300
               Index           =   21
               ItemData        =   "frmParClinic.frx":68BA
               Left            =   1605
               List            =   "frmParClinic.frx":68BC
               Style           =   2  'Dropdown List
               TabIndex        =   386
               Top             =   4410
               Width           =   840
            End
            Begin VB.CheckBox chk 
               Caption         =   "按顺序在当天排列"
               Height          =   180
               Index           =   120
               Left            =   135
               TabIndex        =   382
               Top             =   3630
               Width           =   2505
            End
            Begin VB.ComboBox cbo 
               Height          =   300
               Index           =   33
               ItemData        =   "frmParClinic.frx":68BE
               Left            =   705
               List            =   "frmParClinic.frx":68C0
               Style           =   2  'Dropdown List
               TabIndex        =   377
               Top             =   2580
               Width           =   2100
            End
            Begin VB.ComboBox cbo 
               Height          =   300
               Index           =   32
               ItemData        =   "frmParClinic.frx":68C2
               Left            =   705
               List            =   "frmParClinic.frx":68C4
               Style           =   2  'Dropdown List
               TabIndex        =   375
               Top             =   2250
               Width           =   2100
            End
            Begin VB.ComboBox cbo 
               Height          =   300
               Index           =   22
               ItemData        =   "frmParClinic.frx":68C6
               Left            =   705
               List            =   "frmParClinic.frx":68C8
               Style           =   2  'Dropdown List
               TabIndex        =   363
               Top             =   270
               Width           =   2100
            End
            Begin VB.ComboBox cbo 
               Height          =   300
               Index           =   23
               ItemData        =   "frmParClinic.frx":68CA
               Left            =   705
               List            =   "frmParClinic.frx":68CC
               Style           =   2  'Dropdown List
               TabIndex        =   365
               Top             =   600
               Width           =   2100
            End
            Begin VB.ComboBox cbo 
               Height          =   300
               Index           =   24
               ItemData        =   "frmParClinic.frx":68CE
               Left            =   705
               List            =   "frmParClinic.frx":68D0
               Style           =   2  'Dropdown List
               TabIndex        =   367
               Top             =   930
               Width           =   2100
            End
            Begin VB.ComboBox cbo 
               Height          =   300
               Index           =   25
               ItemData        =   "frmParClinic.frx":68D2
               Left            =   705
               List            =   "frmParClinic.frx":68D4
               Style           =   2  'Dropdown List
               TabIndex        =   369
               Top             =   1260
               Width           =   2100
            End
            Begin VB.ComboBox cbo 
               Height          =   300
               Index           =   30
               ItemData        =   "frmParClinic.frx":68D6
               Left            =   705
               List            =   "frmParClinic.frx":68D8
               Style           =   2  'Dropdown List
               TabIndex        =   371
               Top             =   1590
               Width           =   2100
            End
            Begin VB.ComboBox cbo 
               Height          =   300
               Index           =   31
               ItemData        =   "frmParClinic.frx":68DA
               Left            =   705
               List            =   "frmParClinic.frx":68DC
               Style           =   2  'Dropdown List
               TabIndex        =   373
               Top             =   1920
               Width           =   2100
            End
            Begin VB.ComboBox cbo 
               Height          =   300
               Index           =   34
               ItemData        =   "frmParClinic.frx":68DE
               Left            =   705
               List            =   "frmParClinic.frx":68E0
               Style           =   2  'Dropdown List
               TabIndex        =   379
               Top             =   2910
               Width           =   2100
            End
            Begin VB.CheckBox chk 
               Caption         =   "超出40刻度缩小字体显示"
               Height          =   180
               Index           =   119
               Left            =   135
               TabIndex        =   384
               Top             =   4170
               Width           =   2580
            End
            Begin VB.ComboBox cbo 
               Height          =   300
               Index           =   35
               ItemData        =   "frmParClinic.frx":68E2
               Left            =   705
               List            =   "frmParClinic.frx":68E4
               Style           =   2  'Dropdown List
               TabIndex        =   381
               Top             =   3240
               Width           =   2100
            End
            Begin VB.CheckBox chk 
               Caption         =   "顶格输出(不勾为体温42℃)"
               Height          =   180
               Index           =   118
               Left            =   135
               TabIndex        =   383
               Top             =   3900
               Width           =   2655
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "说明与时间之间以          连接"
               Height          =   180
               Index           =   26
               Left            =   135
               TabIndex        =   385
               Top             =   4455
               Width           =   2700
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "出生"
               Height          =   180
               Index           =   21
               Left            =   315
               TabIndex        =   376
               Top             =   2640
               Width           =   360
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "分娩"
               Height          =   180
               Index           =   20
               Left            =   315
               TabIndex        =   374
               Top             =   2310
               Width           =   360
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "入院"
               Height          =   180
               Index           =   44
               Left            =   315
               TabIndex        =   362
               Top             =   315
               Width           =   360
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "入科"
               Height          =   180
               Index           =   45
               Left            =   315
               TabIndex        =   364
               Top             =   645
               Width           =   360
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "转科"
               Height          =   180
               Index           =   46
               Left            =   315
               TabIndex        =   366
               Top             =   990
               Width           =   360
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "换床"
               Height          =   180
               Index           =   48
               Left            =   315
               TabIndex        =   368
               Top             =   1320
               Width           =   360
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "手术"
               Height          =   180
               Index           =   49
               Left            =   315
               TabIndex        =   370
               Top             =   1650
               Width           =   360
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "出院"
               Height          =   180
               Index           =   50
               Left            =   315
               TabIndex        =   372
               Top             =   1980
               Width           =   360
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "回室"
               Height          =   180
               Index           =   19
               Left            =   315
               TabIndex        =   378
               Top             =   2970
               Width           =   360
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "转病区"
               Height          =   180
               Index           =   18
               Left            =   135
               TabIndex        =   380
               Top             =   3300
               Width           =   540
            End
         End
         Begin VB.Frame fra 
            BorderStyle     =   0  'None
            Height          =   4110
            Index           =   1
            Left            =   135
            TabIndex        =   421
            Top             =   510
            Width           =   5535
            Begin VB.PictureBox picControl 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   1845
               Index           =   0
               Left            =   3060
               ScaleHeight     =   1845
               ScaleWidth      =   2295
               TabIndex        =   292
               TabStop         =   0   'False
               Top             =   1215
               Visible         =   0   'False
               Width           =   2295
               Begin VB.PictureBox PicColorCollect 
                  Appearance      =   0  'Flat
                  AutoRedraw      =   -1  'True
                  AutoSize        =   -1  'True
                  BackColor       =   &H80000005&
                  BorderStyle     =   0  'None
                  ForeColor       =   &H80000008&
                  Height          =   1350
                  Index           =   0
                  Left            =   60
                  Picture         =   "frmParClinic.frx":68E6
                  ScaleHeight     =   1350
                  ScaleWidth      =   2160
                  TabIndex        =   293
                  TabStop         =   0   'False
                  Top             =   90
                  Width           =   2160
                  Begin VB.Shape shpBorder 
                     BorderColor     =   &H00C56A31&
                     FillColor       =   &H00FF8080&
                     Height          =   270
                     Index           =   0
                     Left            =   1890
                     Top             =   1080
                     Visible         =   0   'False
                     Width           =   270
                  End
                  Begin VB.Shape shpValue 
                     BorderColor     =   &H00C56A31&
                     FillColor       =   &H00FF8080&
                     Height          =   270
                     Index           =   0
                     Left            =   0
                     Top             =   0
                     Visible         =   0   'False
                     Width           =   270
                  End
               End
               Begin VB.PictureBox picColor 
                  Appearance      =   0  'Flat
                  AutoRedraw      =   -1  'True
                  BackColor       =   &H00FFFFFF&
                  ForeColor       =   &H80000008&
                  Height          =   200
                  Index           =   0
                  Left            =   90
                  ScaleHeight     =   165
                  ScaleWidth      =   165
                  TabIndex        =   294
                  TabStop         =   0   'False
                  Top             =   1575
                  Width           =   200
               End
               Begin VB.CommandButton cmdUnVisible 
                  Height          =   315
                  Index           =   0
                  Left            =   1815
                  Picture         =   "frmParClinic.frx":705C
                  Style           =   1  'Graphical
                  TabIndex        =   296
                  TabStop         =   0   'False
                  Top             =   1500
                  Width           =   450
               End
               Begin VB.Label lblColor 
                  Caption         =   "&HFFFFFF"
                  Height          =   195
                  Index           =   0
                  Left            =   405
                  TabIndex        =   295
                  Top             =   1575
                  UseMnemonic     =   0   'False
                  Width           =   1365
               End
            End
            Begin VB.ComboBox cbo 
               Height          =   300
               Index           =   3
               ItemData        =   "frmParClinic.frx":75E6
               Left            =   1860
               List            =   "frmParClinic.frx":75E8
               Style           =   2  'Dropdown List
               TabIndex        =   276
               Top             =   60
               Width           =   2565
            End
            Begin VB.CheckBox chk 
               Caption         =   "预览、打印时签名人显示签名图片"
               Height          =   180
               Index           =   117
               Left            =   360
               TabIndex        =   288
               Top             =   2865
               Width           =   3645
            End
            Begin VB.CheckBox chk 
               Caption         =   "护理文件页码按文件顺序编号"
               Height          =   180
               Index           =   116
               Left            =   360
               TabIndex        =   291
               Top             =   3750
               Width           =   3135
            End
            Begin VB.CheckBox chk 
               Caption         =   "住院病人同一时间需要记录多份护理文件"
               Height          =   180
               Index           =   115
               Left            =   360
               TabIndex        =   286
               Top             =   2280
               Width           =   3645
            End
            Begin VB.CheckBox chk 
               Caption         =   "只在当前页中显示跨页数据（不勾两页均显示）"
               Height          =   180
               Index           =   114
               Left            =   360
               TabIndex        =   290
               Top             =   3450
               Width           =   4215
            End
            Begin VB.ComboBox cbo 
               Height          =   300
               Index           =   14
               ItemData        =   "frmParClinic.frx":75EA
               Left            =   1860
               List            =   "frmParClinic.frx":75EC
               Style           =   2  'Dropdown List
               TabIndex        =   280
               Top             =   885
               Width           =   2565
            End
            Begin VB.TextBox txtUD 
               Alignment       =   2  'Center
               Height          =   270
               IMEMode         =   3  'DISABLE
               Index           =   21
               Left            =   1860
               Locked          =   -1  'True
               MaxLength       =   2
               TabIndex        =   284
               Text            =   "1"
               Top             =   1920
               Width           =   375
            End
            Begin VB.ComboBox cbo 
               Height          =   300
               Index           =   13
               ItemData        =   "frmParClinic.frx":75EE
               Left            =   1860
               List            =   "frmParClinic.frx":75F0
               Style           =   2  'Dropdown List
               TabIndex        =   278
               Top             =   480
               Width           =   2565
            End
            Begin VB.CheckBox chk 
               Caption         =   "预览、打印时同一页相同日期显示一次"
               Height          =   180
               Index           =   113
               Left            =   360
               TabIndex        =   289
               Top             =   3165
               Width           =   3540
            End
            Begin VB.PictureBox picLineColor 
               Appearance      =   0  'Flat
               BackColor       =   &H000000FF&
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   0
               Left            =   1860
               ScaleHeight     =   180
               ScaleWidth      =   2535
               TabIndex        =   282
               TabStop         =   0   'False
               Top             =   1575
               Width           =   2565
            End
            Begin MSComCtl2.UpDown ud 
               Height          =   270
               Index           =   21
               Left            =   2235
               TabIndex        =   285
               TabStop         =   0   'False
               Top             =   1920
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   476
               _Version        =   393216
               Value           =   1
               BuddyControl    =   "txtUD(21)"
               BuddyDispid     =   196669
               BuddyIndex      =   21
               OrigLeft        =   2160
               OrigTop         =   1830
               OrigRight       =   2415
               OrigBottom      =   2100
               Max             =   30
               SyncBuddy       =   -1  'True
               BuddyProperty   =   65547
               Enabled         =   -1  'True
            End
            Begin VB.CheckBox chk 
               Caption         =   "不同护理文件之间允许数据同步"
               Height          =   180
               Index           =   167
               Left            =   645
               TabIndex        =   287
               Top             =   2580
               Width           =   2895
            End
            Begin VB.CheckBox chk 
               Caption         =   "不画未记录汇总数据的下划线"
               Height          =   180
               Index           =   188
               Left            =   1890
               TabIndex        =   649
               Top             =   1260
               Width           =   2655
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "审签模式"
               Height          =   180
               Index           =   17
               Left            =   1080
               TabIndex        =   275
               Top             =   120
               Width           =   720
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "小结缺省标识"
               Height          =   180
               Index           =   16
               Left            =   720
               TabIndex        =   279
               Top             =   960
               Width           =   1080
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "允许录入超过当前        天的护理记录数据"
               Height          =   180
               Index           =   15
               Left            =   360
               TabIndex        =   283
               Top             =   1950
               Width           =   3600
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "护士、签名列显示模式"
               Height          =   180
               Index           =   14
               Left            =   15
               TabIndex        =   277
               Top             =   540
               Width           =   1800
            End
            Begin VB.Label lblLineColor 
               AutoSize        =   -1  'True
               Caption         =   "小结标识颜色"
               Height          =   180
               Index           =   0
               Left            =   720
               TabIndex        =   281
               Top             =   1590
               Width           =   1080
            End
         End
         Begin VB.Frame fraPoliceLine 
            Height          =   1230
            Left            =   -69285
            TabIndex        =   414
            Top             =   1350
            Width           =   3735
            Begin VB.ComboBox cbo 
               Height          =   300
               Index           =   5
               ItemData        =   "frmParClinic.frx":75F2
               Left            =   1590
               List            =   "frmParClinic.frx":75FC
               Style           =   2  'Dropdown List
               TabIndex        =   418
               Top             =   765
               Width           =   1905
            End
            Begin VB.ComboBox cbo 
               Height          =   300
               Index           =   4
               ItemData        =   "frmParClinic.frx":760C
               Left            =   1575
               List            =   "frmParClinic.frx":7616
               Style           =   2  'Dropdown List
               TabIndex        =   416
               Top             =   345
               Width           =   1905
            End
            Begin VB.CheckBox chk 
               Caption         =   "产程图上显示警戒、异常线"
               Height          =   180
               Index           =   112
               Left            =   0
               TabIndex        =   413
               Top             =   0
               Width           =   2550
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               Caption         =   "警戒线显示为"
               Height          =   180
               Index           =   13
               Left            =   465
               TabIndex        =   415
               Top             =   390
               Width           =   1080
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               Caption         =   "异常线显示为"
               Height          =   180
               Index           =   12
               Left            =   465
               TabIndex        =   417
               Top             =   825
               Width           =   1080
            End
         End
         Begin VB.CheckBox chk 
            Caption         =   "先露高低显示在左侧(不勾为右侧)"
            Height          =   180
            Index           =   111
            Left            =   -69285
            TabIndex        =   412
            Top             =   1065
            Width           =   3330
         End
         Begin VB.CheckBox chk 
            Caption         =   "产程图模式为交叉式(不勾为伴行式)"
            Height          =   180
            Index           =   110
            Left            =   -69285
            TabIndex        =   411
            Top             =   795
            Width           =   3330
         End
         Begin VB.CheckBox chk 
            Caption         =   "产程图上显示产程时间"
            Height          =   180
            Index           =   109
            Left            =   -69285
            TabIndex        =   410
            Top             =   525
            Width           =   2490
         End
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   6
            ItemData        =   "frmParClinic.frx":7626
            Left            =   -67185
            List            =   "frmParClinic.frx":7630
            Style           =   2  'Dropdown List
            TabIndex        =   420
            Top             =   2685
            Width           =   1650
         End
         Begin VB.Frame fra2 
            Caption         =   "生产措施标志"
            Height          =   1230
            Left            =   -74475
            TabIndex        =   405
            Top             =   3405
            Width           =   3735
            Begin VB.ComboBox cbo 
               Height          =   300
               Index           =   9
               ItemData        =   "frmParClinic.frx":764C
               Left            =   1110
               List            =   "frmParClinic.frx":7659
               Style           =   2  'Dropdown List
               TabIndex        =   407
               Top             =   330
               Width           =   2100
            End
            Begin VB.ComboBox cbo 
               Height          =   300
               Index           =   10
               ItemData        =   "frmParClinic.frx":7687
               Left            =   1110
               List            =   "frmParClinic.frx":7691
               Style           =   2  'Dropdown List
               TabIndex        =   409
               Top             =   765
               Width           =   2100
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "标志内容"
               Height          =   180
               Index           =   10
               Left            =   360
               TabIndex        =   406
               Top             =   390
               Width           =   720
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "标志位置"
               Height          =   180
               Index           =   9
               Left            =   360
               TabIndex        =   408
               Top             =   825
               Width           =   720
            End
         End
         Begin VB.Frame fra1 
            Caption         =   "生产曲线标志(顺产)"
            Height          =   1230
            Index           =   0
            Left            =   -74475
            TabIndex        =   395
            Top             =   525
            Width           =   3735
            Begin VB.ComboBox cbo 
               Height          =   300
               Index           =   12
               ItemData        =   "frmParClinic.frx":76AD
               Left            =   1110
               List            =   "frmParClinic.frx":76BA
               Style           =   2  'Dropdown List
               TabIndex        =   399
               Top             =   765
               Width           =   2100
            End
            Begin VB.ComboBox cbo 
               Height          =   300
               Index           =   11
               ItemData        =   "frmParClinic.frx":76E4
               Left            =   1110
               List            =   "frmParClinic.frx":76F1
               Style           =   2  'Dropdown List
               TabIndex        =   397
               Top             =   330
               Width           =   2100
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "先露下降"
               Height          =   180
               Index           =   8
               Left            =   360
               TabIndex        =   398
               Top             =   825
               Width           =   720
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "宫口扩大"
               Height          =   180
               Index           =   5
               Left            =   360
               TabIndex        =   396
               Top             =   390
               Width           =   720
            End
         End
         Begin VB.Frame fra1 
            Caption         =   "生产曲线标志(异常产)"
            Height          =   1230
            Index           =   1
            Left            =   -74475
            TabIndex        =   400
            Top             =   1965
            Width           =   3735
            Begin VB.ComboBox cbo 
               Height          =   300
               Index           =   7
               ItemData        =   "frmParClinic.frx":771B
               Left            =   1110
               List            =   "frmParClinic.frx":7728
               Style           =   2  'Dropdown List
               TabIndex        =   402
               Top             =   330
               Width           =   2100
            End
            Begin VB.ComboBox cbo 
               Height          =   300
               Index           =   8
               ItemData        =   "frmParClinic.frx":7752
               Left            =   1110
               List            =   "frmParClinic.frx":775F
               Style           =   2  'Dropdown List
               TabIndex        =   404
               Top             =   765
               Width           =   2100
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "宫口扩大"
               Height          =   180
               Index           =   6
               Left            =   360
               TabIndex        =   401
               Top             =   390
               Width           =   720
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "先露下降"
               Height          =   180
               Index           =   7
               Left            =   360
               TabIndex        =   403
               Top             =   825
               Width           =   720
            End
         End
         Begin MSComCtl2.UpDown ud 
            Height          =   270
            Index           =   27
            Left            =   -72870
            TabIndex        =   299
            TabStop         =   0   'False
            Top             =   525
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   476
            _Version        =   393216
            Value           =   2
            BuddyControl    =   "txtUD(27)"
            BuddyDispid     =   196669
            BuddyIndex      =   27
            OrigLeft        =   2190
            OrigTop         =   870
            OrigRight       =   2430
            OrigBottom      =   1170
            Max             =   4
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown ud 
            Height          =   270
            Index           =   26
            Left            =   -72045
            TabIndex        =   308
            TabStop         =   0   'False
            Top             =   840
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   476
            _Version        =   393216
            Value           =   2
            BuddyControl    =   "txtUD(26)"
            BuddyDispid     =   196669
            BuddyIndex      =   26
            OrigLeft        =   2190
            OrigTop         =   870
            OrigRight       =   2430
            OrigBottom      =   1170
            Max             =   30
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown ud 
            Height          =   270
            Index           =   24
            Left            =   -70320
            TabIndex        =   305
            TabStop         =   0   'False
            Top             =   525
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   476
            _Version        =   393216
            Value           =   8
            BuddyControl    =   "txtUD(24)"
            BuddyDispid     =   196669
            BuddyIndex      =   24
            OrigLeft        =   4580
            OrigTop         =   885
            OrigRight       =   4835
            OrigBottom      =   1170
            Max             =   23
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown ud 
            Height          =   270
            Index           =   23
            Left            =   -71670
            TabIndex        =   302
            TabStop         =   0   'False
            Top             =   525
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   476
            _Version        =   393216
            Value           =   18
            BuddyControl    =   "txtUD(23)"
            BuddyDispid     =   196669
            BuddyIndex      =   23
            OrigLeft        =   3230
            OrigTop         =   885
            OrigRight       =   3485
            OrigBottom      =   1170
            Max             =   23
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown ud 
            Height          =   270
            Index           =   22
            Left            =   -72045
            TabIndex        =   311
            TabStop         =   0   'False
            Top             =   1155
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   476
            _Version        =   393216
            BuddyControl    =   "txtUD(22)"
            BuddyDispid     =   196669
            BuddyIndex      =   22
            OrigLeft        =   2190
            OrigTop         =   870
            OrigRight       =   2430
            OrigBottom      =   1170
            Max             =   50
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin VB.PictureBox picControl 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1845
            Index           =   1
            Left            =   -70665
            ScaleHeight     =   1845
            ScaleWidth      =   2295
            TabIndex        =   390
            TabStop         =   0   'False
            Top             =   915
            Visible         =   0   'False
            Width           =   2295
            Begin VB.CommandButton cmdUnVisible 
               Height          =   315
               Index           =   1
               Left            =   1815
               Picture         =   "frmParClinic.frx":7789
               Style           =   1  'Graphical
               TabIndex        =   394
               TabStop         =   0   'False
               Top             =   1515
               Width           =   450
            End
            Begin VB.PictureBox PicColorCollect 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   1350
               Index           =   1
               Left            =   60
               Picture         =   "frmParClinic.frx":7D13
               ScaleHeight     =   1350
               ScaleWidth      =   2160
               TabIndex        =   391
               TabStop         =   0   'False
               Top             =   90
               Width           =   2160
               Begin VB.Shape shpValue 
                  BorderColor     =   &H00C56A31&
                  FillColor       =   &H00FF8080&
                  Height          =   270
                  Index           =   1
                  Left            =   0
                  Top             =   -15
                  Visible         =   0   'False
                  Width           =   270
               End
               Begin VB.Shape shpBorder 
                  BorderColor     =   &H00C56A31&
                  FillColor       =   &H00FF8080&
                  Height          =   270
                  Index           =   1
                  Left            =   1890
                  Top             =   1080
                  Visible         =   0   'False
                  Width           =   270
               End
            End
            Begin VB.PictureBox picColor 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H80000008&
               Height          =   200
               Index           =   1
               Left            =   90
               ScaleHeight     =   165
               ScaleWidth      =   165
               TabIndex        =   392
               TabStop         =   0   'False
               Top             =   1575
               Width           =   200
            End
            Begin VB.Label lblColor 
               Caption         =   "&HFFFFFF"
               Height          =   195
               Index           =   1
               Left            =   405
               TabIndex        =   393
               Top             =   1575
               UseMnemonic     =   0   'False
               Width           =   1365
            End
         End
         Begin VB.CheckBox chk 
            Caption         =   "体温单以单格显示(不勾为双格)"
            Height          =   180
            Index           =   127
            Left            =   -68325
            TabIndex        =   357
            Top             =   585
            Width           =   2895
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "灌肠大便显示格式："
            Height          =   180
            Index           =   41
            Left            =   -74700
            TabIndex        =   353
            Top             =   6840
            Width           =   1620
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "手术后标注    天,颜色"
            Height          =   180
            Index           =   40
            Left            =   -74700
            TabIndex        =   325
            Top             =   3045
            Width           =   1890
         End
         Begin VB.Line Line1 
            X1              =   -73770
            X2              =   -73485
            Y1              =   3255
            Y2              =   3255
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ",当天缺省格式"
            Height          =   180
            Index           =   27
            Left            =   -72390
            TabIndex        =   328
            Top             =   3045
            Width           =   1170
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "出院标志优先显示："
            Height          =   180
            Index           =   39
            Left            =   -74700
            TabIndex        =   349
            Top             =   6585
            Width           =   1620
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "脉搏短绌录入方式："
            Height          =   180
            Index           =   38
            Left            =   -74700
            TabIndex        =   345
            Top             =   6345
            Width           =   1620
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "脉搏短绌填充方式："
            Height          =   180
            Index           =   37
            Left            =   -74700
            TabIndex        =   340
            Top             =   6090
            Width           =   1620
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "体温复试合格显示符号    ,颜色"
            Height          =   180
            Index           =   36
            Left            =   -74700
            TabIndex        =   322
            Top             =   2760
            Width           =   2610
         End
         Begin VB.Line Line2 
            X1              =   -72885
            X2              =   -72600
            Y1              =   2955
            Y2              =   2955
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   ",颜色"
            Height          =   180
            Index           =   35
            Left            =   -71595
            TabIndex        =   314
            Top             =   1515
            Width           =   450
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "未记说明显示位置"
            Height          =   180
            Index           =   34
            Left            =   -74700
            TabIndex        =   312
            Top             =   1530
            Width           =   1440
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "体温不升显示方式"
            Height          =   180
            Index           =   33
            Left            =   -74700
            TabIndex        =   316
            Top             =   1860
            Width           =   1440
         End
         Begin VB.Label lblBreathe 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "呼吸项为表格时,呼吸机数据的显示方式"
            Height          =   180
            Left            =   -74700
            TabIndex        =   320
            Top             =   2490
            Width           =   3150
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "呼吸项为表格时,数据输出时的显示方式"
            Height          =   180
            Index           =   32
            Left            =   -74700
            TabIndex        =   318
            Top             =   2175
            Width           =   3150
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "体温开始记录时间"
            Height          =   180
            Index           =   31
            Left            =   -74700
            TabIndex        =   297
            Top             =   585
            Width           =   1440
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "体温表输出时,表格数据固定        行"
            Height          =   180
            Index           =   30
            Left            =   -74700
            TabIndex        =   306
            Top             =   885
            Width           =   3150
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "次日       点"
            Height          =   180
            Index           =   24
            Left            =   -71010
            TabIndex        =   303
            Top             =   585
            Width           =   1170
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "夜班从       点至"
            Height          =   180
            Index           =   23
            Left            =   -72540
            TabIndex        =   300
            Top             =   585
            Width           =   1530
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "体温表输出时,曲线固定添加        行"
            Height          =   180
            Index           =   22
            Left            =   -74700
            TabIndex        =   309
            Top             =   1185
            Width           =   3150
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "产程图0点与第一个曲线点"
            Height          =   180
            Index           =   11
            Left            =   -69285
            TabIndex        =   419
            Top             =   2745
            Width           =   2070
         End
      End
   End
   Begin VB.PictureBox picPar 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   7335
      Index           =   5
      Left            =   2400
      ScaleHeight     =   7305
      ScaleWidth      =   10065
      TabIndex        =   118
      Top             =   0
      Width           =   10095
      Begin VB.CheckBox chk 
         Caption         =   "科室药房对照按本机参数设置"
         Height          =   315
         Index           =   163
         Left            =   7335
         TabIndex        =   548
         Top             =   90
         Width           =   2700
      End
      Begin TabDlg.SSTab SST 
         Height          =   7095
         Left            =   120
         TabIndex        =   119
         Top             =   120
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   12515
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "门诊"
         TabPicture(0)   =   "frmParClinic.frx":8489
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lbl(28)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label8"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "lblDept(28)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "vsUseDept(28)"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "vsfDrugStore(28)"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "cbo(28)"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "cmdDel(28)"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "cmdAdd(28)"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).ControlCount=   8
         TabCaption(1)   =   "住院"
         TabPicture(1)   =   "frmParClinic.frx":84A5
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "cbo(29)"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "cmdDel(29)"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "cmdAdd(29)"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "vsUseDept(29)"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "vsfDrugStore(29)"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).Control(5)=   "lblDept(29)"
         Tab(1).Control(5).Enabled=   0   'False
         Tab(1).Control(6)=   "lblDept(0)"
         Tab(1).Control(6).Enabled=   0   'False
         Tab(1).Control(7)=   "lbl(29)"
         Tab(1).Control(7).Enabled=   0   'False
         Tab(1).ControlCount=   8
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   29
            Left            =   -74280
            Style           =   2  'Dropdown List
            TabIndex        =   130
            Top             =   452
            Width           =   1455
         End
         Begin VB.CommandButton cmdDel 
            Height          =   315
            Index           =   29
            Left            =   -72225
            Picture         =   "frmParClinic.frx":84C1
            Style           =   1  'Graphical
            TabIndex        =   129
            Top             =   445
            Width           =   450
         End
         Begin VB.CommandButton cmdAdd 
            Height          =   315
            Index           =   29
            Left            =   -72720
            Picture         =   "frmParClinic.frx":ED13
            Style           =   1  'Graphical
            TabIndex        =   128
            Top             =   445
            Width           =   450
         End
         Begin VB.CommandButton cmdAdd 
            Height          =   315
            Index           =   28
            Left            =   2280
            Picture         =   "frmParClinic.frx":15565
            Style           =   1  'Graphical
            TabIndex        =   123
            Top             =   443
            Width           =   450
         End
         Begin VB.CommandButton cmdDel 
            Height          =   315
            Index           =   28
            Left            =   2775
            Picture         =   "frmParClinic.frx":1BDB7
            Style           =   1  'Graphical
            TabIndex        =   122
            Top             =   443
            Width           =   450
         End
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   28
            Left            =   720
            Style           =   2  'Dropdown List
            TabIndex        =   121
            Top             =   450
            Width           =   1455
         End
         Begin VSFlex8Ctl.VSFlexGrid vsfDrugStore 
            Height          =   5475
            Index           =   28
            Left            =   4440
            TabIndex        =   124
            Top             =   1200
            Width           =   5055
            _cx             =   8916
            _cy             =   9657
            Appearance      =   1
            BorderStyle     =   1
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   14737632
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   4
            Cols            =   5
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmParClinic.frx":22609
            ScrollTrack     =   0   'False
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   -1  'True
            AutoSizeMode    =   0
            AutoSearch      =   0
            AutoSearchDelay =   2
            MultiTotals     =   -1  'True
            SubtotalPosition=   1
            OutlineBar      =   0
            OutlineCol      =   0
            Ellipsis        =   0
            ExplorerBar     =   0
            PicturesOver    =   0   'False
            FillStyle       =   0
            RightToLeft     =   0   'False
            PictureType     =   0
            TabBehavior     =   0
            OwnerDraw       =   0
            Editable        =   0
            ShowComboButton =   1
            WordWrap        =   0   'False
            TextStyle       =   0
            TextStyleFixed  =   0
            OleDragMode     =   0
            OleDropMode     =   0
            DataMode        =   0
            VirtualData     =   -1  'True
            DataMember      =   ""
            ComboSearch     =   3
            AutoSizeMouse   =   -1  'True
            FrozenRows      =   0
            FrozenCols      =   0
            AllowUserFreezing=   0
            BackColorFrozen =   0
            ForeColorFrozen =   0
            WallPaperAlignment=   9
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   24
         End
         Begin VSFlex8Ctl.VSFlexGrid vsUseDept 
            Height          =   5445
            Index           =   28
            Left            =   240
            TabIndex        =   125
            Top             =   1200
            Width           =   3885
            _cx             =   6853
            _cy             =   9604
            Appearance      =   1
            BorderStyle     =   1
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   16772055
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483638
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   7
            Cols            =   1
            FixedRows       =   0
            FixedCols       =   0
            RowHeightMin    =   245
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmParClinic.frx":226B6
            ScrollTrack     =   0   'False
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   -1  'True
            AutoSizeMode    =   1
            AutoSearch      =   0
            AutoSearchDelay =   2
            MultiTotals     =   -1  'True
            SubtotalPosition=   1
            OutlineBar      =   0
            OutlineCol      =   0
            Ellipsis        =   0
            ExplorerBar     =   0
            PicturesOver    =   0   'False
            FillStyle       =   0
            RightToLeft     =   0   'False
            PictureType     =   0
            TabBehavior     =   0
            OwnerDraw       =   0
            Editable        =   2
            ShowComboButton =   1
            WordWrap        =   -1  'True
            TextStyle       =   0
            TextStyleFixed  =   0
            OleDragMode     =   0
            OleDropMode     =   0
            DataMode        =   0
            VirtualData     =   -1  'True
            DataMember      =   ""
            ComboSearch     =   3
            AutoSizeMouse   =   -1  'True
            FrozenRows      =   0
            FrozenCols      =   0
            AllowUserFreezing=   0
            BackColorFrozen =   0
            ForeColorFrozen =   0
            WallPaperAlignment=   9
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   24
         End
         Begin VSFlex8Ctl.VSFlexGrid vsUseDept 
            Height          =   5445
            Index           =   29
            Left            =   -74760
            TabIndex        =   132
            Top             =   1200
            Width           =   3885
            _cx             =   6853
            _cy             =   9604
            Appearance      =   1
            BorderStyle     =   1
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483638
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   7
            Cols            =   1
            FixedRows       =   0
            FixedCols       =   0
            RowHeightMin    =   245
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmParClinic.frx":22862
            ScrollTrack     =   0   'False
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   -1  'True
            AutoSizeMode    =   1
            AutoSearch      =   0
            AutoSearchDelay =   2
            MultiTotals     =   -1  'True
            SubtotalPosition=   1
            OutlineBar      =   0
            OutlineCol      =   0
            Ellipsis        =   0
            ExplorerBar     =   0
            PicturesOver    =   0   'False
            FillStyle       =   0
            RightToLeft     =   0   'False
            PictureType     =   0
            TabBehavior     =   0
            OwnerDraw       =   0
            Editable        =   2
            ShowComboButton =   1
            WordWrap        =   -1  'True
            TextStyle       =   0
            TextStyleFixed  =   0
            OleDragMode     =   0
            OleDropMode     =   0
            DataMode        =   0
            VirtualData     =   -1  'True
            DataMember      =   ""
            ComboSearch     =   3
            AutoSizeMouse   =   -1  'True
            FrozenRows      =   0
            FrozenCols      =   0
            AllowUserFreezing=   0
            BackColorFrozen =   0
            ForeColorFrozen =   0
            WallPaperAlignment=   9
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   24
         End
         Begin VSFlex8Ctl.VSFlexGrid vsfDrugStore 
            Height          =   5415
            Index           =   29
            Left            =   -70560
            TabIndex        =   134
            Top             =   1200
            Width           =   5055
            _cx             =   8916
            _cy             =   9551
            Appearance      =   1
            BorderStyle     =   1
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   14737632
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   4
            Cols            =   4
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmParClinic.frx":22A0E
            ScrollTrack     =   0   'False
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   -1  'True
            AutoSizeMode    =   0
            AutoSearch      =   0
            AutoSearchDelay =   2
            MultiTotals     =   -1  'True
            SubtotalPosition=   1
            OutlineBar      =   0
            OutlineCol      =   0
            Ellipsis        =   0
            ExplorerBar     =   0
            PicturesOver    =   0   'False
            FillStyle       =   0
            RightToLeft     =   0   'False
            PictureType     =   0
            TabBehavior     =   0
            OwnerDraw       =   0
            Editable        =   0
            ShowComboButton =   1
            WordWrap        =   0   'False
            TextStyle       =   0
            TextStyleFixed  =   0
            OleDragMode     =   0
            OleDropMode     =   0
            DataMode        =   0
            VirtualData     =   -1  'True
            DataMember      =   ""
            ComboSearch     =   3
            AutoSizeMouse   =   -1  'True
            FrozenRows      =   0
            FrozenCols      =   0
            AllowUserFreezing=   0
            BackColorFrozen =   0
            ForeColorFrozen =   0
            WallPaperAlignment=   9
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   24
         End
         Begin VB.Label lblDept 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "发药药房"
            Height          =   180
            Index           =   29
            Left            =   -70560
            TabIndex        =   135
            Top             =   937
            Width           =   720
         End
         Begin VB.Label lblDept 
            Caption         =   "应用科室"
            Height          =   255
            Index           =   0
            Left            =   -74760
            TabIndex        =   133
            Top             =   900
            Width           =   855
         End
         Begin VB.Label lbl 
            Caption         =   "方案"
            Height          =   255
            Index           =   29
            Left            =   -74760
            TabIndex        =   131
            Top             =   490
            Width           =   375
         End
         Begin VB.Label lblDept 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "发药药房"
            Height          =   180
            Index           =   28
            Left            =   4440
            TabIndex        =   127
            Top             =   930
            Width           =   720
         End
         Begin VB.Label Label8 
            Caption         =   "应用科室"
            Height          =   255
            Left            =   240
            TabIndex        =   126
            Top             =   900
            Width           =   855
         End
         Begin VB.Label lbl 
            Caption         =   "方案"
            Height          =   255
            Index           =   28
            Left            =   240
            TabIndex        =   120
            Top             =   495
            Width           =   375
         End
      End
   End
   Begin VB.PictureBox picPar 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   8175
      Index           =   7
      Left            =   2400
      ScaleHeight     =   8145
      ScaleWidth      =   10065
      TabIndex        =   149
      Top             =   0
      Width           =   10095
      Begin VB.CheckBox chk 
         Caption         =   "病理诊断录入后可以不填写病理号"
         Height          =   255
         Index           =   187
         Left            =   240
         TabIndex        =   648
         Top             =   4800
         Width           =   4455
      End
      Begin VB.CheckBox chk 
         Caption         =   "诊断录入时附码自动提取"
         Height          =   255
         Index           =   170
         Left            =   240
         TabIndex        =   630
         Top             =   4560
         Width           =   4455
      End
      Begin VB.CheckBox chk 
         Caption         =   "按ICD-10录入时，病理诊断只允许录入M打头的肿瘤形态学编码"
         Height          =   255
         Index           =   164
         Left            =   240
         TabIndex        =   549
         Top             =   1560
         Width           =   6615
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   255
         Index           =   6
         Left            =   3360
         TabIndex        =   542
         Top             =   1920
         Width           =   4215
         Begin VB.OptionButton optICD附码 
            Caption         =   "必须填写"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   545
            Top             =   0
            Width           =   1095
         End
         Begin VB.OptionButton optICD附码 
            Caption         =   "提示是否填写"
            Height          =   255
            Index           =   2
            Left            =   1200
            TabIndex        =   544
            Top             =   0
            Width           =   1455
         End
         Begin VB.OptionButton optICD附码 
            Caption         =   "不检查"
            Height          =   255
            Index           =   0
            Left            =   2640
            TabIndex        =   543
            Top             =   0
            Value           =   -1  'True
            Width           =   855
         End
      End
      Begin VB.CheckBox chk 
         Caption         =   "路径变异原因从字典表中选取，不允许自由录入"
         Height          =   255
         Index           =   155
         Left            =   2520
         TabIndex        =   521
         Top             =   4250
         Width           =   4215
      End
      Begin VB.CheckBox chk 
         Caption         =   "诊断手术名称自由调整"
         Height          =   255
         Index           =   154
         Left            =   240
         TabIndex        =   520
         Top             =   4250
         Width           =   3375
      End
      Begin VB.CheckBox chk 
         Caption         =   "身份证加密显示"
         Height          =   255
         Index           =   136
         Left            =   240
         TabIndex        =   484
         Top             =   615
         Width           =   1575
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   255
         Index           =   3
         Left            =   3480
         TabIndex        =   168
         Top             =   1320
         Width           =   4215
         Begin VB.OptionButton opt病理诊断 
            Caption         =   "不检查"
            Height          =   255
            Index           =   0
            Left            =   2640
            TabIndex        =   171
            Top             =   0
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton opt病理诊断 
            Caption         =   "提示是否填写"
            Height          =   255
            Index           =   2
            Left            =   1200
            TabIndex        =   170
            Top             =   0
            Width           =   1455
         End
         Begin VB.OptionButton opt病理诊断 
            Caption         =   "必须填写"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   169
            Top             =   0
            Width           =   1095
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   255
         Index           =   1
         Left            =   3840
         TabIndex        =   164
         Top             =   960
         Width           =   3615
         Begin VB.OptionButton opt损伤中毒 
            Caption         =   "必须填写"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   167
            Top             =   0
            Width           =   1095
         End
         Begin VB.OptionButton opt损伤中毒 
            Caption         =   "提示是否填写"
            Height          =   255
            Index           =   2
            Left            =   1200
            TabIndex        =   166
            Top             =   0
            Width           =   1455
         End
         Begin VB.OptionButton opt损伤中毒 
            Caption         =   "不检查"
            Height          =   255
            Index           =   0
            Left            =   2640
            TabIndex        =   165
            Top             =   0
            Value           =   -1  'True
            Width           =   855
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   255
         Index           =   2
         Left            =   1440
         TabIndex        =   160
         Top             =   2200
         Width           =   4215
         Begin VB.OptionButton opt区域 
            Caption         =   "必须填写"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   163
            Top             =   0
            Width           =   1095
         End
         Begin VB.OptionButton opt区域 
            Caption         =   "提示是否填写"
            Height          =   255
            Index           =   2
            Left            =   1200
            TabIndex        =   162
            Top             =   0
            Width           =   1455
         End
         Begin VB.OptionButton opt区域 
            Caption         =   "不检查"
            Height          =   255
            Index           =   0
            Left            =   2640
            TabIndex        =   161
            Top             =   0
            Value           =   -1  'True
            Width           =   855
         End
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   0
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   159
         Top             =   240
         Width           =   2535
      End
      Begin VB.CheckBox chk 
         Caption         =   "使用手术结束时间"
         Height          =   255
         Index           =   69
         Left            =   2280
         TabIndex        =   158
         Top             =   615
         Width           =   2055
      End
      Begin VB.CheckBox chk 
         Caption         =   $"frmParClinic.frx":22A97
         Height          =   255
         Index           =   70
         Left            =   240
         TabIndex        =   157
         Top             =   2475
         Width           =   4095
      End
      Begin VB.CheckBox chk 
         Caption         =   "医生和护士分别填写病案首页"
         Height          =   255
         Index           =   71
         Left            =   240
         TabIndex        =   156
         Top             =   3480
         Width           =   4095
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "删除(&D)"
         Enabled         =   0   'False
         Height          =   350
         Left            =   8685
         TabIndex        =   152
         Top             =   5175
         Width           =   1100
      End
      Begin VB.CommandButton cmdModify 
         Caption         =   "修改(&M)"
         Enabled         =   0   'False
         Height          =   350
         Left            =   7485
         TabIndex        =   151
         Top             =   5175
         Width           =   1100
      End
      Begin VB.CommandButton cmdAddMed 
         Caption         =   "增加(&A)"
         Height          =   350
         Left            =   6285
         TabIndex        =   150
         Top             =   5175
         Width           =   1100
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfMecItem 
         Height          =   2415
         Left            =   120
         TabIndex        =   153
         Top             =   5640
         Width           =   9720
         _cx             =   17145
         _cy             =   4260
         Appearance      =   3
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16574424
         ForeColorSel    =   -2147483642
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   2
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   250
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Label lblICD附码 
         AutoSize        =   -1  'True
         Caption         =   "主要出院诊断为C00到D48时,ICD附码："
         Height          =   180
         Left            =   240
         TabIndex        =   546
         Top             =   1935
         Width           =   3240
      End
      Begin VB.Label lbl病理诊断 
         AutoSize        =   -1  'True
         Caption         =   "主要出院诊断为C00到D48时,病理诊断："
         Height          =   180
         Left            =   240
         TabIndex        =   177
         Top             =   1335
         Width           =   3330
      End
      Begin VB.Label lbl损伤中毒 
         AutoSize        =   -1  'True
         Caption         =   "主要出院诊断编码为S、T类时,损伤中毒诊断："
         Height          =   180
         Left            =   240
         TabIndex        =   176
         Top             =   975
         Width           =   3690
      End
      Begin VB.Label lbl区域 
         AutoSize        =   -1  'True
         Caption         =   "保存时区域项："
         Height          =   180
         Left            =   240
         TabIndex        =   175
         Top             =   2220
         Width           =   1260
      End
      Begin VB.Label lbl首页标准 
         Caption         =   "病案首页标准"
         Height          =   255
         Left            =   240
         TabIndex        =   174
         Top             =   285
         Width           =   1215
      End
      Begin VB.Label lbl中医 
         Caption         =   $"frmParClinic.frx":22AC3
         Height          =   585
         Left            =   480
         TabIndex        =   173
         Top             =   2790
         Width           =   6015
      End
      Begin VB.Label lblSeparEdit 
         Caption         =   "不良事件的项目、输液反应、引发药物、临床表现、住院期间身体约束、离院时透析(血透、腹透)尿素氮值等信息在启用该参数时只能由护士填写"
         Height          =   360
         Left            =   480
         TabIndex        =   172
         Top             =   3800
         Width           =   6015
      End
      Begin VB.Label Label3 
         Caption         =   "住院首页附加项目："
         Height          =   255
         Left            =   120
         TabIndex        =   154
         Top             =   5280
         Width           =   4095
      End
   End
   Begin VB.PictureBox picPar 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   7725
      Index           =   0
      Left            =   2400
      ScaleHeight     =   7695
      ScaleWidth      =   10065
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   10095
      Begin VB.CheckBox chk 
         Caption         =   "医嘱用药天数反算"
         Height          =   240
         Index           =   138
         Left            =   240
         TabIndex        =   486
         Top             =   2445
         Width           =   1875
      End
      Begin VB.CommandButton cmdAdvice 
         Caption         =   "医嘱内容定义(&F)"
         Height          =   405
         Left            =   240
         TabIndex        =   12
         Top             =   3000
         Width           =   1680
      End
      Begin VB.Frame fra入院诊断 
         Caption         =   "住院下达以下类别医嘱时检查是否填写入院诊断"
         Height          =   1365
         Left            =   240
         TabIndex        =   22
         Top             =   6165
         Width           =   4320
         Begin VB.CommandButton cmd住院检查入院诊断 
            Caption         =   "全清"
            Height          =   300
            Index           =   1
            Left            =   3120
            TabIndex        =   25
            Top             =   720
            Width           =   900
         End
         Begin VB.CommandButton cmd住院检查入院诊断 
            Caption         =   "全选"
            Height          =   300
            Index           =   0
            Left            =   3120
            TabIndex        =   24
            Top             =   360
            Width           =   900
         End
         Begin VB.ListBox lst 
            Columns         =   3
            ForeColor       =   &H80000012&
            Height          =   900
            IMEMode         =   3  'DISABLE
            Index           =   2
            Left            =   120
            Style           =   1  'Checkbox
            TabIndex        =   23
            Top             =   375
            Width           =   2940
         End
      End
      Begin VB.Frame fra抗菌目的 
         BorderStyle     =   0  'None
         Height          =   1935
         Index           =   2
         Left            =   240
         TabIndex        =   422
         Top             =   3600
         Width           =   4380
         Begin VB.Frame fra抗菌目的 
            Caption         =   "门诊"
            Height          =   680
            Index           =   0
            Left            =   0
            TabIndex        =   14
            Top             =   330
            Width           =   4305
            Begin VB.OptionButton opt抗菌目的门诊 
               Caption         =   "治疗"
               Height          =   180
               Index           =   2
               Left            =   3000
               TabIndex        =   17
               Top             =   300
               Value           =   -1  'True
               Width           =   680
            End
            Begin VB.OptionButton opt抗菌目的门诊 
               Caption         =   "预防"
               Height          =   180
               Index           =   1
               Left            =   1920
               TabIndex        =   16
               Top             =   300
               Width           =   680
            End
            Begin VB.OptionButton opt抗菌目的门诊 
               Caption         =   "下达时确定"
               Height          =   180
               Index           =   0
               Left            =   255
               TabIndex        =   15
               Top             =   300
               Width           =   1275
            End
         End
         Begin VB.Frame fra抗菌目的 
            Caption         =   "住院"
            Height          =   680
            Index           =   1
            Left            =   0
            TabIndex        =   18
            Top             =   1200
            Width           =   4305
            Begin VB.OptionButton opt抗菌目的住院 
               Caption         =   "下达时确定"
               Height          =   180
               Index           =   0
               Left            =   255
               TabIndex        =   19
               Top             =   300
               Width           =   1275
            End
            Begin VB.OptionButton opt抗菌目的住院 
               Caption         =   "预防"
               Height          =   180
               Index           =   1
               Left            =   1920
               TabIndex        =   20
               Top             =   300
               Width           =   680
            End
            Begin VB.OptionButton opt抗菌目的住院 
               Caption         =   "治疗"
               Height          =   180
               Index           =   2
               Left            =   3000
               TabIndex        =   21
               Top             =   300
               Value           =   -1  'True
               Width           =   680
            End
         End
         Begin VB.Label lbl抗菌目的 
            Caption         =   "抗菌药物缺省用药目的"
            Height          =   255
            Left            =   0
            TabIndex        =   13
            Top             =   60
            Width           =   1935
         End
      End
      Begin VB.Frame fra医嘱下达 
         Caption         =   "住院医嘱下达"
         Height          =   4170
         Left            =   4920
         TabIndex        =   423
         Top             =   3360
         Width           =   4770
         Begin VB.CheckBox chk 
            Caption         =   "皮试阳性用药"
            Height          =   195
            Index           =   178
            Left            =   2775
            TabIndex        =   625
            Top             =   3120
            Width           =   1935
         End
         Begin VB.CheckBox chk 
            Caption         =   "下达毒麻和第一类精神药品时必须登记代办人"
            Height          =   195
            Index           =   174
            Left            =   210
            TabIndex        =   621
            Top             =   1150
            Width           =   3960
         End
         Begin VB.CheckBox chk 
            Caption         =   "根据皮试结果限制医嘱发送"
            Height          =   225
            Index           =   166
            Left            =   210
            TabIndex        =   560
            Top             =   3120
            Width           =   2585
         End
         Begin VB.CheckBox chk 
            Caption         =   "住院医嘱下达时诊疗选择器显示药品库存"
            Height          =   240
            Index           =   51
            Left            =   210
            TabIndex        =   424
            Top             =   2550
            Width           =   4125
         End
         Begin VB.CheckBox chk 
            Caption         =   "长期医嘱缺省为次日生效"
            Height          =   240
            Index           =   24
            Left            =   210
            TabIndex        =   425
            Top             =   300
            Width           =   2280
         End
         Begin VB.CheckBox chk 
            Caption         =   "药品长期医嘱按规格下达"
            Height          =   240
            Index           =   4
            Left            =   210
            TabIndex        =   426
            Top             =   570
            Width           =   2280
         End
         Begin VB.CheckBox chk 
            Caption         =   "下达药品临嘱时可以指定用药天数"
            Height          =   195
            Index           =   15
            Left            =   210
            TabIndex        =   427
            Top             =   1440
            Width           =   3180
         End
         Begin VB.CheckBox chk 
            Caption         =   "临嘱的执行频率缺省为一次性"
            Height          =   195
            Index           =   13
            Left            =   210
            TabIndex        =   428
            Top             =   870
            Width           =   3180
         End
         Begin VB.CheckBox chk 
            Caption         =   "下达出院医嘱时检查出院诊断的填写"
            Height          =   195
            Index           =   16
            Left            =   210
            TabIndex        =   429
            Top             =   1725
            Width           =   3180
         End
         Begin VB.CheckBox chk 
            Caption         =   "下达临嘱时先输入单量"
            Height          =   195
            Index           =   14
            Left            =   2280
            TabIndex        =   430
            Top             =   2850
            Width           =   2205
         End
         Begin VB.CheckBox chk 
            Caption         =   "手术执行完成后才允许下达术后医嘱"
            Height          =   195
            Index           =   17
            Left            =   210
            TabIndex        =   431
            Top             =   2010
            Width           =   3180
         End
         Begin VB.CheckBox chk 
            Caption         =   "保存医嘱时自动排序"
            Height          =   255
            Index           =   20
            Left            =   210
            TabIndex        =   432
            Top             =   3780
            Width           =   1935
         End
         Begin VB.CommandButton cmdAdviceSortSet 
            Caption         =   "排序规则设置(&S)"
            Height          =   350
            Left            =   2280
            TabIndex        =   433
            Top             =   3720
            Width           =   1695
         End
         Begin VB.CheckBox chk 
            Caption         =   "允许给待入住病人下达医嘱"
            Height          =   195
            Index           =   18
            Left            =   210
            TabIndex        =   434
            Top             =   2280
            Value           =   1  'Checked
            Width           =   2895
         End
         Begin VB.CheckBox chk 
            Caption         =   "自动增加皮试医嘱"
            Height          =   225
            Index           =   19
            Left            =   210
            TabIndex        =   435
            Top             =   2850
            Width           =   1860
         End
         Begin VB.OptionButton opt未皮试限制医嘱 
            Caption         =   "所有药品"
            Height          =   255
            Index           =   0
            Left            =   945
            TabIndex        =   436
            Top             =   3405
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton opt未皮试限制医嘱 
            Caption         =   "配制中心接收的药品"
            Height          =   255
            Index           =   1
            Left            =   2025
            TabIndex        =   437
            Top             =   3405
            Width           =   1920
         End
         Begin VB.Label lblSTCheck 
            Caption         =   "限制"
            Height          =   255
            Left            =   480
            TabIndex        =   438
            Top             =   3420
            Width           =   375
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "门诊医嘱下达"
         Height          =   3030
         Left            =   4920
         TabIndex        =   26
         Top             =   240
         Width           =   4770
         Begin VB.CheckBox chk 
            Caption         =   "门诊西医科允许录入中医诊断"
            Height          =   180
            Index           =   186
            Left            =   240
            TabIndex        =   647
            Top             =   2715
            Width           =   3885
         End
         Begin VB.CheckBox chk 
            Caption         =   "皮试阳性用药"
            Height          =   195
            Index           =   177
            Left            =   2760
            TabIndex        =   624
            Top             =   1335
            Width           =   1935
         End
         Begin VB.CheckBox chk 
            Caption         =   "根据皮试结果限制医嘱发送"
            Height          =   195
            Index           =   165
            Left            =   240
            TabIndex        =   559
            Top             =   1335
            Width           =   2535
         End
         Begin VB.CheckBox chk 
            Caption         =   "门诊医嘱下达时诊疗选择器显示药品库存"
            Height          =   240
            Index           =   48
            Left            =   240
            TabIndex        =   439
            Top             =   2460
            Width           =   3660
         End
         Begin MSComCtl2.UpDown ud 
            Height          =   270
            Index           =   3
            Left            =   2370
            TabIndex        =   440
            TabStop         =   0   'False
            Top             =   1890
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   476
            _Version        =   393216
            Value           =   5
            BuddyControl    =   "txtUD(3)"
            BuddyDispid     =   196669
            BuddyIndex      =   3
            OrigLeft        =   2400
            OrigTop         =   1080
            OrigRight       =   2655
            OrigBottom      =   1350
            Max             =   999
            Min             =   1
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   0   'False
         End
         Begin VB.TextBox txtUD 
            Alignment       =   2  'Center
            BackColor       =   &H80000004&
            Enabled         =   0   'False
            Height          =   270
            IMEMode         =   3  'DISABLE
            Index           =   3
            Left            =   1970
            Locked          =   -1  'True
            MaxLength       =   3
            TabIndex        =   441
            Text            =   "5"
            Top             =   1890
            Width           =   375
         End
         Begin VB.TextBox txtUD 
            Alignment       =   2  'Center
            Height          =   270
            IMEMode         =   3  'DISABLE
            Index           =   11
            Left            =   1350
            MaxLength       =   4
            TabIndex        =   28
            Text            =   "10"
            Top             =   240
            Width           =   400
         End
         Begin VB.CheckBox chk 
            Caption         =   "门诊药嘱先作废后退药"
            Height          =   240
            Index           =   0
            Left            =   240
            TabIndex        =   442
            Top             =   2205
            Width           =   2460
         End
         Begin VB.CheckBox chk 
            Caption         =   "自动增加皮试医嘱"
            Height          =   195
            Index           =   2
            Left            =   240
            TabIndex        =   443
            Top             =   1080
            Width           =   1935
         End
         Begin VB.CheckBox chk 
            Caption         =   "下达药品医嘱时可以指定用药天数"
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   444
            Top             =   825
            Width           =   3360
         End
         Begin VB.CheckBox chk 
            Caption         =   "下达药品医嘱时必须录入药品单量"
            Height          =   195
            Index           =   3
            Left            =   240
            TabIndex        =   445
            Top             =   555
            Width           =   3360
         End
         Begin VB.CheckBox chk 
            Caption         =   "下达毒麻和第一类精神药品时必须登记代办人"
            Height          =   195
            Index           =   5
            Left            =   240
            TabIndex        =   446
            Top             =   1620
            Width           =   3960
         End
         Begin MSComCtl2.UpDown ud 
            Height          =   270
            Index           =   11
            Left            =   1760
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   476
            _Version        =   393216
            Value           =   10
            BuddyControl    =   "txtUD(11)"
            BuddyDispid     =   196669
            BuddyIndex      =   11
            OrigLeft        =   2100
            OrigTop         =   420
            OrigRight       =   2355
            OrigBottom      =   690
            Max             =   9999
            Min             =   1
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin VB.CheckBox chk 
            Caption         =   "一张处方最多允许        条药品医嘱"
            Height          =   240
            Index           =   52
            Left            =   240
            TabIndex        =   447
            Top             =   1905
            Width           =   3780
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "新开医嘱间隔        分钟则以当前时间作为开始时间"
            Height          =   180
            Index           =   25
            Left            =   240
            TabIndex        =   27
            Top             =   285
            Width           =   4320
         End
      End
      Begin VB.CheckBox chk 
         Caption         =   "允许处理超过挂号有效天数的病人"
         Height          =   195
         Index           =   82
         Left            =   240
         TabIndex        =   11
         Top             =   2160
         Width           =   3045
      End
      Begin VB.CheckBox chk 
         Caption         =   "过敏登记有效天数"
         Height          =   240
         Index           =   11
         Left            =   240
         TabIndex        =   4
         Top             =   690
         Width           =   1740
      End
      Begin VB.CheckBox chk 
         Caption         =   "一次申请多个检验项目"
         Height          =   240
         Index           =   34
         Left            =   240
         TabIndex        =   10
         Top             =   1860
         Width           =   2820
      End
      Begin VB.TextBox txtUD 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   270
         Index           =   7
         Left            =   2010
         MaxLength       =   3
         TabIndex        =   5
         Text            =   "1"
         Top             =   675
         Width           =   495
      End
      Begin VB.CheckBox chk 
         Caption         =   "下达医嘱时显示产地"
         Height          =   240
         Index           =   66
         Left            =   240
         TabIndex        =   9
         Top             =   1560
         Value           =   1  'Checked
         Width           =   1980
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   26
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox txtUD 
         Alignment       =   2  'Center
         Height          =   270
         IMEMode         =   3  'DISABLE
         Index           =   8
         Left            =   1710
         MaxLength       =   4
         TabIndex        =   2
         Text            =   "30"
         Top             =   240
         Width           =   465
      End
      Begin MSComCtl2.UpDown ud 
         Height          =   270
         Index           =   8
         Left            =   2205
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   240
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   476
         _Version        =   393216
         Value           =   30
         BuddyControl    =   "txtUD(8)"
         BuddyDispid     =   196669
         BuddyIndex      =   8
         OrigLeft        =   2100
         OrigTop         =   120
         OrigRight       =   2340
         OrigBottom      =   390
         Max             =   9999
         Min             =   10
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown ud 
         Height          =   270
         Index           =   7
         Left            =   2520
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   675
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   476
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtUD(7)"
         BuddyDispid     =   196669
         BuddyIndex      =   7
         OrigLeft        =   2400
         OrigTop         =   1380
         OrigRight       =   2655
         OrigBottom      =   1680
         Max             =   365
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   0   'False
      End
      Begin VB.Label lbl 
         Caption         =   "中药配方每行"
         Height          =   255
         Index           =   54
         Left            =   240
         TabIndex        =   7
         Top             =   1110
         Width           =   1095
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "补录医嘱识别间隔         分钟"
         Height          =   180
         Index           =   43
         Left            =   240
         TabIndex        =   1
         Top             =   300
         Width           =   2610
      End
   End
   Begin VB.PictureBox picPar 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   7455
      Index           =   3
      Left            =   2400
      ScaleHeight     =   7425
      ScaleWidth      =   10065
      TabIndex        =   112
      Top             =   0
      Width           =   10095
      Begin VB.Frame fraStopDept 
         Height          =   3215
         Left            =   5880
         TabIndex        =   561
         Top             =   4080
         Width           =   4095
         Begin VB.CheckBox chk 
            Caption         =   "停嘱时录入原因"
            Height          =   195
            Index           =   56
            Left            =   120
            TabIndex        =   562
            Top             =   0
            Width           =   1575
         End
         Begin VSFlex8Ctl.VSFlexGrid vsStopDept 
            Height          =   2640
            Left            =   120
            TabIndex        =   563
            Top             =   490
            Width           =   3900
            _cx             =   6879
            _cy             =   4657
            Appearance      =   2
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   16772055
            ForeColorSel    =   0
            BackColorBkg    =   -2147483636
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   9
            Cols            =   8
            FixedRows       =   0
            FixedCols       =   0
            RowHeightMin    =   280
            RowHeightMax    =   0
            ColWidthMin     =   500
            ColWidthMax     =   960
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmParClinic.frx":22B81
            ScrollTrack     =   0   'False
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   -1  'True
            AutoSizeMode    =   0
            AutoSearch      =   0
            AutoSearchDelay =   2
            MultiTotals     =   -1  'True
            SubtotalPosition=   1
            OutlineBar      =   0
            OutlineCol      =   0
            Ellipsis        =   0
            ExplorerBar     =   0
            PicturesOver    =   0   'False
            FillStyle       =   0
            RightToLeft     =   0   'False
            PictureType     =   0
            TabBehavior     =   0
            OwnerDraw       =   0
            Editable        =   2
            ShowComboButton =   1
            WordWrap        =   0   'False
            TextStyle       =   0
            TextStyleFixed  =   0
            OleDragMode     =   0
            OleDropMode     =   0
            DataMode        =   0
            VirtualData     =   -1  'True
            DataMember      =   ""
            ComboSearch     =   3
            AutoSizeMouse   =   -1  'True
            FrozenRows      =   0
            FrozenCols      =   0
            AllowUserFreezing=   0
            BackColorFrozen =   0
            ForeColorFrozen =   0
            WallPaperAlignment=   9
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   24
         End
         Begin VB.Label lblStopDept 
            AutoSize        =   -1  'True
            Caption         =   "设置可不填写停嘱原因的科室，例如：精神科。"
            Height          =   180
            Left            =   120
            TabIndex        =   564
            Top             =   240
            Width           =   3780
         End
      End
      Begin VB.Frame Frame14 
         Caption         =   "发送时按类别将本科执行的项目填为已执行"
         Height          =   3915
         Left            =   5880
         TabIndex        =   494
         Top             =   120
         Width           =   4065
         Begin VB.CommandButton cmdAdd 
            Height          =   315
            Index           =   37
            Left            =   2205
            Picture         =   "frmParClinic.frx":22C37
            Style           =   1  'Graphical
            TabIndex        =   640
            Top             =   1620
            Width           =   450
         End
         Begin VB.CommandButton cmdDel 
            Height          =   315
            Index           =   37
            Left            =   2700
            Picture         =   "frmParClinic.frx":29489
            Style           =   1  'Graphical
            TabIndex        =   639
            Top             =   1620
            Width           =   450
         End
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   37
            Left            =   645
            Style           =   2  'Dropdown List
            TabIndex        =   638
            Top             =   1620
            Width           =   1455
         End
         Begin VB.CommandButton cmd本科执行自动完成医嘱类别 
            Caption         =   "全清"
            Height          =   300
            Index           =   1
            Left            =   3195
            TabIndex        =   499
            Top             =   855
            Width           =   800
         End
         Begin VB.CommandButton cmd本科执行自动完成医嘱类别 
            Caption         =   "全选"
            Height          =   300
            Index           =   0
            Left            =   3195
            TabIndex        =   498
            Top             =   495
            Width           =   800
         End
         Begin VB.CheckBox chk 
            Caption         =   "临嘱"
            Height          =   255
            Index           =   40
            Left            =   975
            TabIndex        =   497
            Top             =   210
            Width           =   735
         End
         Begin VB.CheckBox chk 
            Caption         =   "长嘱"
            Height          =   255
            Index           =   39
            Left            =   135
            TabIndex        =   496
            Top             =   210
            Width           =   735
         End
         Begin VB.ListBox lst 
            Columns         =   3
            Height          =   900
            Index           =   8
            ItemData        =   "frmParClinic.frx":2FCDB
            Left            =   120
            List            =   "frmParClinic.frx":2FCE2
            Style           =   1  'Checkbox
            TabIndex        =   495
            Top             =   510
            Width           =   2985
         End
         Begin VSFlex8Ctl.VSFlexGrid vsUseDept 
            Height          =   1650
            Index           =   37
            Left            =   135
            TabIndex        =   641
            Top             =   2190
            Width           =   3795
            _cx             =   6694
            _cy             =   2910
            Appearance      =   1
            BorderStyle     =   1
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   16772055
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483638
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   7
            Cols            =   1
            FixedRows       =   0
            FixedCols       =   0
            RowHeightMin    =   245
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmParClinic.frx":2FCF0
            ScrollTrack     =   0   'False
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   -1  'True
            AutoSizeMode    =   1
            AutoSearch      =   0
            AutoSearchDelay =   2
            MultiTotals     =   -1  'True
            SubtotalPosition=   1
            OutlineBar      =   0
            OutlineCol      =   0
            Ellipsis        =   0
            ExplorerBar     =   0
            PicturesOver    =   0   'False
            FillStyle       =   0
            RightToLeft     =   0   'False
            PictureType     =   0
            TabBehavior     =   0
            OwnerDraw       =   0
            Editable        =   2
            ShowComboButton =   1
            WordWrap        =   -1  'True
            TextStyle       =   0
            TextStyleFixed  =   0
            OleDragMode     =   0
            OleDropMode     =   0
            DataMode        =   0
            VirtualData     =   -1  'True
            DataMember      =   ""
            ComboSearch     =   3
            AutoSizeMouse   =   -1  'True
            FrozenRows      =   0
            FrozenCols      =   0
            AllowUserFreezing=   0
            BackColorFrozen =   0
            ForeColorFrozen =   0
            WallPaperAlignment=   9
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   24
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "方案"
            Height          =   180
            Index           =   42
            Left            =   165
            TabIndex        =   643
            Top             =   1665
            Width           =   360
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "应用科室"
            Height          =   180
            Index           =   47
            Left            =   150
            TabIndex        =   642
            Top             =   1965
            Width           =   720
         End
      End
      Begin VB.Frame fra住院医嘱单打印 
         Caption         =   "住院医嘱单打印模式"
         Height          =   900
         Left            =   150
         TabIndex        =   77
         Top             =   6395
         Width           =   2700
         Begin VB.OptionButton opt住院医嘱单打印 
            Caption         =   "新开时打印"
            Height          =   180
            Index           =   1
            Left            =   150
            TabIndex        =   79
            Top             =   615
            Width           =   1440
         End
         Begin VB.OptionButton opt住院医嘱单打印 
            Caption         =   "校对后打印"
            Height          =   180
            Index           =   0
            Left            =   150
            TabIndex        =   78
            Top             =   270
            Value           =   -1  'True
            Width           =   1545
         End
      End
      Begin VB.Frame fra输血申请单打印 
         Caption         =   "输血申请单打印模式"
         Height          =   900
         Left            =   3015
         TabIndex        =   80
         Top             =   6395
         Width           =   2700
         Begin VB.OptionButton opt输血申请单打印 
            Caption         =   "发送时打印"
            Height          =   180
            Index           =   1
            Left            =   150
            TabIndex        =   81
            Top             =   270
            Value           =   -1  'True
            Width           =   1545
         End
         Begin VB.OptionButton opt输血申请单打印 
            Caption         =   "新开时打印"
            Height          =   180
            Index           =   0
            Left            =   150
            TabIndex        =   82
            Top             =   585
            Width           =   1440
         End
      End
      Begin VB.Frame fra后续处理 
         Caption         =   "住院医嘱处理与控制 "
         Height          =   2745
         Left            =   120
         TabIndex        =   67
         Top             =   3615
         Width           =   5625
         Begin VB.CheckBox chk 
            Caption         =   "会诊科室下达医嘱由会诊申请科室处理"
            Height          =   180
            Index           =   184
            Left            =   1980
            TabIndex        =   645
            Top             =   2415
            Width           =   3540
         End
         Begin VB.CheckBox chk 
            Caption         =   "叮嘱需要发送执行"
            Height          =   240
            Index           =   168
            Left            =   210
            TabIndex        =   565
            Top             =   2400
            Width           =   1845
         End
         Begin VB.CheckBox chk 
            Caption         =   "回退出院医嘱才能撤销预出院"
            Height          =   240
            Index           =   81
            Left            =   2520
            TabIndex        =   76
            Top             =   2070
            Width           =   2700
         End
         Begin VB.CheckBox chk 
            Caption         =   "下达出院医嘱才能出院"
            Height          =   240
            Index           =   50
            Left            =   210
            TabIndex        =   75
            Top             =   2070
            Width           =   2340
         End
         Begin VB.CheckBox chk 
            Caption         =   $"frmParClinic.frx":2FE9C
            Height          =   195
            Index           =   33
            Left            =   210
            TabIndex        =   74
            Top             =   1785
            Width           =   2460
         End
         Begin VB.CheckBox chk 
            Caption         =   "允许批量校对"
            Height          =   195
            Index           =   28
            Left            =   210
            TabIndex        =   70
            Top             =   907
            Width           =   1380
         End
         Begin VB.CheckBox chk 
            Caption         =   "允许批量暂停/启用"
            Height          =   195
            Index           =   29
            Left            =   2520
            TabIndex        =   71
            Top             =   907
            Width           =   1935
         End
         Begin VB.CheckBox chk 
            Caption         =   "校对,确认停止,重整医嘱后进行打印"
            Height          =   180
            Index           =   26
            Left            =   210
            TabIndex        =   68
            Top             =   360
            Width           =   3180
         End
         Begin VB.CheckBox chk 
            Caption         =   "允许对医技下达的医嘱进行后续处理"
            Height          =   195
            Index           =   31
            Left            =   210
            TabIndex        =   73
            Top             =   1515
            Width           =   3180
         End
         Begin VB.CheckBox chk 
            Caption         =   "校对和确认停止时使用电子签名"
            Height          =   195
            Index           =   27
            Left            =   210
            TabIndex        =   69
            Top             =   626
            Width           =   3165
         End
         Begin VB.CheckBox chk 
            Caption         =   "填写皮试结果时验证身份"
            Height          =   195
            Index           =   30
            Left            =   210
            TabIndex        =   72
            Top             =   1230
            Width           =   2445
         End
      End
      Begin VB.Frame fra住院发送 
         Caption         =   "住院发送选项"
         Height          =   3450
         Left            =   120
         TabIndex        =   59
         Top             =   120
         Width           =   5610
         Begin VB.CheckBox chk 
            Caption         =   "存在未发送医嘱时禁止处理转科医嘱"
            Height          =   180
            Index           =   185
            Left            =   240
            TabIndex        =   646
            Top             =   3195
            Width           =   4755
         End
         Begin VB.CheckBox chk 
            Caption         =   "发送完成后关闭医嘱窗体"
            Height          =   180
            Index           =   161
            Left            =   240
            TabIndex        =   533
            Top             =   2970
            Width           =   3285
         End
         Begin VB.CheckBox chk 
            Caption         =   "特殊药品分开发送"
            Height          =   180
            Index           =   160
            Left            =   240
            TabIndex        =   532
            Top             =   2720
            Width           =   1845
         End
         Begin VB.CheckBox chk 
            Caption         =   "检验医嘱发送时一组检验发送为一张单据"
            Height          =   180
            Index           =   141
            Left            =   240
            TabIndex        =   493
            Top             =   2420
            Width           =   4305
         End
         Begin VB.Frame Frame13 
            BorderStyle     =   0  'None
            Caption         =   "Frame7"
            Height          =   255
            Left            =   2085
            TabIndex        =   488
            Top             =   2140
            Width           =   2655
            Begin VB.OptionButton opt领药部门 
               Caption         =   "给药执行科室"
               Height          =   180
               Index           =   0
               Left            =   0
               TabIndex        =   489
               Top             =   0
               Width           =   1455
            End
            Begin VB.OptionButton opt领药部门 
               Caption         =   "病人病区"
               Height          =   180
               Index           =   1
               Left            =   1560
               TabIndex        =   490
               Top             =   0
               Value           =   -1  'True
               Width           =   1050
            End
         End
         Begin VB.CheckBox chk 
            Caption         =   "住院药嘱发送产生领药号"
            Height          =   240
            Index           =   64
            Left            =   240
            TabIndex        =   60
            Top             =   300
            Width           =   2280
         End
         Begin VB.CheckBox chk 
            Caption         =   "长嘱口服药发送结束时间"
            Height          =   195
            Index           =   25
            Left            =   240
            TabIndex        =   65
            Top             =   1820
            Width           =   2325
         End
         Begin VB.CheckBox chk 
            Caption         =   "有未校对或待发送医嘱禁止发送转科、出院、转院、死亡医嘱"
            Height          =   180
            Index           =   22
            Left            =   240
            TabIndex        =   63
            Top             =   1200
            Width           =   5220
         End
         Begin VB.CheckBox chk 
            Caption         =   "无须校对即可发送医嘱"
            Height          =   180
            Index           =   21
            Left            =   240
            TabIndex        =   61
            Top             =   600
            Width           =   2160
         End
         Begin VB.CheckBox chk 
            Caption         =   "发送时对医保病人检查项目是否审批"
            Height          =   195
            Index           =   38
            Left            =   240
            TabIndex        =   62
            Top             =   900
            Value           =   1  'Checked
            Width           =   3180
         End
         Begin VB.CheckBox chk 
            Caption         =   "药品长嘱的给药途径发送次数以结束时间为准计算"
            Height          =   180
            Index           =   23
            Left            =   240
            TabIndex        =   64
            Top             =   1500
            Width           =   4305
         End
         Begin MSComCtl2.DTPicker dtp口服结束时间 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "yyyy-M-d"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   0
            EndProperty
            Height          =   300
            Left            =   2565
            TabIndex        =   66
            Top             =   1767
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   159711234
            CurrentDate     =   0.999988425925926
         End
         Begin VB.Label lbl领药部门 
            Caption         =   "药品医嘱的领药部门为"
            Height          =   180
            Left            =   240
            TabIndex        =   491
            Top             =   2140
            Width           =   1845
         End
      End
   End
   Begin VB.PictureBox picPar 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   7875
      Index           =   1
      Left            =   2400
      ScaleHeight     =   7845
      ScaleWidth      =   10185
      TabIndex        =   100
      Top             =   0
      Visible         =   0   'False
      Width           =   10215
      Begin VB.Frame FrmBloodManager 
         Height          =   1725
         Left            =   120
         TabIndex        =   622
         Top             =   6075
         Width           =   3465
         Begin VB.CheckBox chk 
            Caption         =   "血液回室接收后才允许填写执行登记"
            Enabled         =   0   'False
            Height          =   240
            Index           =   183
            Left            =   120
            TabIndex        =   482
            Top             =   1080
            Value           =   1  'Checked
            Width           =   3210
         End
         Begin VB.CheckBox chk 
            Caption         =   "下达用血申请时确定发血信息"
            Enabled         =   0   'False
            Height          =   240
            Index           =   181
            Left            =   120
            TabIndex        =   481
            Top             =   810
            Width           =   2850
         End
         Begin VB.CheckBox chk 
            Caption         =   "启用血库管理系统"
            Height          =   240
            Index           =   135
            Left            =   120
            TabIndex        =   478
            Top             =   0
            Width           =   1785
         End
         Begin VB.ComboBox cbo 
            Enabled         =   0   'False
            Height          =   300
            Index           =   36
            ItemData        =   "frmParClinic.frx":2FEBA
            Left            =   2280
            List            =   "frmParClinic.frx":2FEBC
            Style           =   2  'Dropdown List
            TabIndex        =   483
            Top             =   1365
            Width           =   1065
         End
         Begin VB.CheckBox chk 
            Caption         =   "用血医嘱发送后输血科才能进行发血"
            Enabled         =   0   'False
            Height          =   255
            Index           =   176
            Left            =   120
            TabIndex        =   479
            Top             =   270
            Width           =   3180
         End
         Begin VB.CheckBox chk 
            Caption         =   "输血申请不显示血液库存信息"
            Enabled         =   0   'False
            Height          =   240
            Index           =   175
            Left            =   120
            TabIndex        =   480
            Top             =   540
            Width           =   2850
         End
         Begin VB.Label lblBloodManager 
            Caption         =   "输血采集默认检验诊疗类型"
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            TabIndex        =   623
            Top             =   1425
            Width           =   2235
         End
      End
      Begin VB.CheckBox chk 
         Caption         =   "住院"
         Height          =   225
         Index           =   159
         Left            =   8175
         TabIndex        =   530
         Top             =   5085
         Width           =   675
      End
      Begin VB.CheckBox chk 
         Caption         =   "门诊"
         Height          =   225
         Index           =   158
         Left            =   7230
         TabIndex        =   529
         Top             =   5085
         Width           =   675
      End
      Begin VB.CheckBox chk 
         Caption         =   "多科会诊由会诊代表科室书写一份意见"
         Height          =   240
         Index           =   146
         Left            =   3765
         TabIndex        =   509
         Top             =   7500
         Width           =   3420
      End
      Begin VB.Frame Frame1 
         Caption         =   "申请单启用环节"
         Height          =   900
         Index           =   5
         Left            =   3765
         TabIndex        =   503
         Top             =   4080
         Width           =   6225
         Begin VB.CheckBox chk 
            Caption         =   "手术"
            Height          =   225
            Index           =   157
            Left            =   3495
            TabIndex        =   528
            Top             =   570
            Width           =   675
         End
         Begin VB.CheckBox chk 
            Caption         =   "输血"
            Height          =   225
            Index           =   156
            Left            =   2595
            TabIndex        =   527
            Top             =   570
            Width           =   675
         End
         Begin VB.CheckBox chk 
            Caption         =   "检验"
            Height          =   225
            Index           =   152
            Left            =   1680
            TabIndex        =   526
            Top             =   570
            Width           =   675
         End
         Begin VB.CheckBox chk 
            Caption         =   "检查"
            Height          =   225
            Index           =   151
            Left            =   765
            TabIndex        =   522
            Top             =   570
            Width           =   675
         End
         Begin VB.CheckBox chk 
            Caption         =   "检查"
            Height          =   225
            Index           =   145
            Left            =   765
            TabIndex        =   508
            Top             =   285
            Width           =   675
         End
         Begin VB.CheckBox chk 
            Caption         =   "检验"
            Height          =   225
            Index           =   144
            Left            =   1680
            TabIndex        =   507
            Top             =   285
            Width           =   675
         End
         Begin VB.CheckBox chk 
            Caption         =   "输血"
            Height          =   225
            Index           =   143
            Left            =   2595
            TabIndex        =   506
            Top             =   285
            Width           =   675
         End
         Begin VB.CheckBox chk 
            Caption         =   "手术"
            Height          =   225
            Index           =   142
            Left            =   3495
            TabIndex        =   505
            Top             =   285
            Width           =   675
         End
         Begin VB.CheckBox chk 
            Caption         =   "会诊"
            Height          =   225
            Index           =   90
            Left            =   4380
            TabIndex        =   504
            Top             =   570
            Width           =   675
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "住院："
            Height          =   180
            Left            =   165
            TabIndex        =   525
            Top             =   570
            Width           =   540
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "门诊："
            Height          =   180
            Left            =   165
            TabIndex        =   524
            Top             =   285
            Width           =   540
         End
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   1
         ItemData        =   "frmParClinic.frx":2FEBE
         Left            =   1920
         List            =   "frmParClinic.frx":2FEC0
         Style           =   2  'Dropdown List
         TabIndex        =   147
         Top             =   5715
         Width           =   1065
      End
      Begin VB.TextBox txt 
         Height          =   600
         Index           =   1
         Left            =   3765
         MultiLine       =   -1  'True
         TabIndex        =   448
         Top             =   6675
         Width           =   6255
      End
      Begin VB.TextBox txt 
         Height          =   600
         Index           =   0
         Left            =   3765
         MultiLine       =   -1  'True
         TabIndex        =   449
         Top             =   5715
         Width           =   6255
      End
      Begin VB.CheckBox chk 
         Caption         =   $"frmParClinic.frx":2FEC2
         Height          =   255
         Index           =   79
         Left            =   120
         TabIndex        =   450
         Top             =   5415
         Width           =   3000
      End
      Begin VB.Frame fraCLKS 
         Height          =   3840
         Left            =   3760
         TabIndex        =   451
         Top             =   120
         Width           =   6250
         Begin VB.CheckBox chk 
            Caption         =   "医嘱超量时必须输入原因"
            Height          =   240
            Index           =   86
            Left            =   105
            TabIndex        =   452
            Top             =   0
            Width           =   2280
         End
         Begin VSFlex8Ctl.VSFlexGrid vsUnWriteDept 
            Height          =   3165
            Left            =   120
            TabIndex        =   453
            Top             =   525
            Width           =   6000
            _cx             =   10583
            _cy             =   5583
            Appearance      =   2
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   16772055
            ForeColorSel    =   0
            BackColorBkg    =   -2147483636
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   10
            Cols            =   8
            FixedRows       =   0
            FixedCols       =   0
            RowHeightMin    =   280
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmParClinic.frx":2FEE6
            ScrollTrack     =   0   'False
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   -1  'True
            AutoSizeMode    =   0
            AutoSearch      =   0
            AutoSearchDelay =   2
            MultiTotals     =   -1  'True
            SubtotalPosition=   1
            OutlineBar      =   0
            OutlineCol      =   0
            Ellipsis        =   0
            ExplorerBar     =   0
            PicturesOver    =   0   'False
            FillStyle       =   0
            RightToLeft     =   0   'False
            PictureType     =   0
            TabBehavior     =   0
            OwnerDraw       =   0
            Editable        =   2
            ShowComboButton =   1
            WordWrap        =   0   'False
            TextStyle       =   0
            TextStyleFixed  =   0
            OleDragMode     =   0
            OleDropMode     =   0
            DataMode        =   0
            VirtualData     =   -1  'True
            DataMember      =   ""
            ComboSearch     =   3
            AutoSizeMouse   =   -1  'True
            FrozenRows      =   0
            FrozenCols      =   0
            AllowUserFreezing=   0
            BackColorFrozen =   0
            ForeColorFrozen =   0
            WallPaperAlignment=   9
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   24
         End
         Begin VB.Label Label16 
            Caption         =   "设置可不录入超量原因的科室，例如：精神科。"
            Height          =   255
            Left            =   240
            TabIndex        =   454
            Top             =   300
            Width           =   4095
         End
      End
      Begin VB.Frame frmOPS 
         Height          =   1095
         Left            =   120
         TabIndex        =   455
         Top             =   3315
         Width           =   3465
         Begin VB.CheckBox chk 
            Caption         =   "主刀医师达到手术等级无需审核"
            Enabled         =   0   'False
            Height          =   240
            Index           =   83
            Left            =   120
            TabIndex        =   519
            Top             =   780
            Width           =   3105
         End
         Begin VB.CheckBox chk 
            Caption         =   "手术分级审核"
            Enabled         =   0   'False
            Height          =   240
            Index           =   140
            Left            =   120
            TabIndex        =   492
            Top             =   525
            Width           =   1425
         End
         Begin VB.CheckBox chk 
            Caption         =   "启用手术医师授权管理"
            Enabled         =   0   'False
            Height          =   240
            Index           =   49
            Left            =   120
            TabIndex        =   456
            Top             =   270
            Width           =   2220
         End
         Begin VB.CheckBox chk 
            Caption         =   "启用手术分级管理"
            Height          =   240
            Index           =   80
            Left            =   120
            TabIndex        =   457
            Top             =   0
            Width           =   1740
         End
      End
      Begin VB.Frame fraKSSStrict 
         Height          =   915
         Index           =   14
         Left            =   120
         TabIndex        =   458
         Top             =   2340
         Width           =   3465
         Begin VB.CheckBox chk 
            Caption         =   "按医疗小组进行抗菌药物审核"
            Enabled         =   0   'False
            Height          =   255
            Index           =   137
            Left            =   120
            TabIndex        =   485
            Top             =   585
            Width           =   2700
         End
         Begin VB.CheckBox chk 
            Caption         =   "启用抗菌药物分级管理"
            Height          =   255
            Index           =   75
            Left            =   120
            TabIndex        =   459
            Top             =   0
            Width           =   2175
         End
         Begin VB.CheckBox chk 
            Caption         =   "抗菌药物允许使用自备药"
            Enabled         =   0   'False
            Height          =   255
            Index           =   76
            Left            =   120
            TabIndex        =   460
            Top             =   300
            Width           =   2295
         End
      End
      Begin VB.Frame fraBlood 
         Height          =   885
         Left            =   120
         TabIndex        =   461
         Top             =   4470
         Width           =   3465
         Begin VB.CheckBox chk 
            Caption         =   "启用输血申请三级审核"
            Enabled         =   0   'False
            Height          =   255
            Index           =   53
            Left            =   120
            TabIndex        =   462
            Top             =   270
            Width           =   2175
         End
         Begin VB.CheckBox chk 
            Caption         =   "输血申请只能由中级及以上医师提出"
            Enabled         =   0   'False
            Height          =   200
            Index           =   85
            Left            =   120
            TabIndex        =   463
            Top             =   570
            Width           =   3230
         End
         Begin VB.CheckBox chk 
            Caption         =   "启用输血分级管理"
            Height          =   255
            Index           =   35
            Left            =   120
            TabIndex        =   464
            Top             =   0
            Width           =   1815
         End
      End
      Begin VB.Frame fraCheckDrug 
         Caption         =   "合理用药接口"
         Height          =   2160
         Left            =   120
         TabIndex        =   465
         Top             =   120
         Width           =   3465
         Begin VB.CommandButton cmdSet 
            Caption         =   "设置"
            Height          =   300
            Left            =   2700
            TabIndex        =   531
            Top             =   0
            Visible         =   0   'False
            Width           =   600
         End
         Begin VB.CheckBox chk 
            Caption         =   "禁忌药品要求填写原因"
            Height          =   240
            Index           =   139
            Left            =   120
            TabIndex        =   487
            Top             =   1200
            Width           =   2940
         End
         Begin VB.OptionButton optPASSVer 
            Caption         =   "美康4.0"
            Height          =   255
            Index           =   1
            Left            =   2040
            TabIndex        =   466
            Top             =   1740
            Width           =   975
         End
         Begin VB.OptionButton optPASSVer 
            Caption         =   "美康3.0"
            Height          =   255
            Index           =   0
            Left            =   1080
            TabIndex        =   467
            Top             =   1740
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.CheckBox chk 
            Caption         =   "允许下达院外执行的禁忌药品医嘱"
            Height          =   240
            Index           =   77
            Left            =   120
            TabIndex        =   469
            Top             =   666
            Width           =   3000
         End
         Begin VB.CheckBox chk 
            Caption         =   "禁止下达超极量药品医嘱"
            Height          =   240
            Index           =   63
            Left            =   120
            TabIndex        =   470
            Top             =   930
            Width           =   2940
         End
         Begin VB.CheckBox chk 
            Caption         =   "允许下达禁忌药品医嘱"
            Height          =   240
            Index           =   65
            Left            =   120
            TabIndex        =   471
            Top             =   400
            Width           =   2100
         End
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   20
            ItemData        =   "frmParClinic.frx":2FF9C
            Left            =   1230
            List            =   "frmParClinic.frx":2FF9E
            Style           =   2  'Dropdown List
            TabIndex        =   472
            Top             =   0
            Width           =   1410
         End
         Begin VB.ComboBox cbo 
            Enabled         =   0   'False
            Height          =   300
            Index           =   27
            ItemData        =   "frmParClinic.frx":2FFA0
            Left            =   1260
            List            =   "frmParClinic.frx":2FFA2
            Style           =   2  'Dropdown List
            TabIndex        =   473
            Top             =   1755
            Width           =   1770
         End
         Begin VB.CheckBox chk 
            Caption         =   "允许使用系统设置"
            Height          =   240
            Index           =   89
            Left            =   120
            TabIndex        =   468
            Top             =   1440
            Width           =   2940
         End
         Begin VB.CheckBox chk 
            Caption         =   "启用接口调用日志"
            Height          =   240
            Index           =   88
            Left            =   120
            TabIndex        =   103
            Top             =   1440
            Width           =   2940
         End
         Begin VB.Label lblPassVer 
            Caption         =   "当前版本："
            Height          =   255
            Left            =   120
            TabIndex        =   474
            Top             =   1740
            Width           =   2895
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "过敏输入来源"
            Height          =   180
            Index           =   0
            Left            =   105
            TabIndex        =   104
            Top             =   1815
            Width           =   1080
         End
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "启用申请单后必须使用申请单下达医嘱："
         Height          =   180
         Left            =   3780
         TabIndex        =   523
         Top             =   5115
         Width           =   3240
      End
      Begin VB.Label Label2 
         Caption         =   "住院医生站病人列表按          显示"
         Height          =   255
         Left            =   120
         TabIndex        =   146
         Top             =   5760
         Width           =   3375
      End
      Begin VB.Label lblBloodPrompt 
         Caption         =   "住院输血申请注意事项"
         Height          =   255
         Index           =   1
         Left            =   3765
         TabIndex        =   475
         Top             =   6405
         Width           =   2655
      End
      Begin VB.Label lblBloodPrompt 
         Caption         =   "门诊输血申请注意事项"
         Height          =   255
         Index           =   0
         Left            =   3765
         TabIndex        =   476
         Top             =   5475
         Width           =   2655
      End
   End
   Begin VB.PictureBox picPar 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   7335
      Index           =   9
      Left            =   2430
      ScaleHeight     =   7305
      ScaleWidth      =   10065
      TabIndex        =   196
      Top             =   45
      Width           =   10095
      Begin VB.Frame fraEprSign 
         Caption         =   "病历签名"
         Height          =   7005
         Left            =   5145
         TabIndex        =   534
         Top             =   120
         Width           =   4740
         Begin VB.TextBox txt 
            Enabled         =   0   'False
            Height          =   270
            Index           =   10
            Left            =   3915
            TabIndex        =   537
            Text            =   "50"
            Top             =   315
            Width           =   390
         End
         Begin VB.CheckBox chk 
            Caption         =   "签名使用图片时用原图"
            Height          =   195
            Index           =   98
            Left            =   75
            TabIndex        =   536
            Top             =   345
            Width           =   2235
         End
         Begin VB.CheckBox chk 
            Caption         =   "签名使用图片"
            Height          =   195
            Index           =   97
            Left            =   2460
            TabIndex        =   535
            Top             =   795
            Visible         =   0   'False
            Width           =   1395
         End
         Begin VSFlex8Ctl.VSFlexGrid vsfDepartSign 
            Height          =   5805
            Left            =   90
            TabIndex        =   538
            Top             =   1080
            Width           =   4575
            _cx             =   8070
            _cy             =   10239
            Appearance      =   0
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   16772055
            ForeColorSel    =   0
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   12698049
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   4
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   255
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmParClinic.frx":2FFA4
            ScrollTrack     =   0   'False
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   1
            MergeCompare    =   0
            AutoResize      =   -1  'True
            AutoSizeMode    =   1
            AutoSearch      =   0
            AutoSearchDelay =   2
            MultiTotals     =   -1  'True
            SubtotalPosition=   1
            OutlineBar      =   0
            OutlineCol      =   0
            Ellipsis        =   0
            ExplorerBar     =   0
            PicturesOver    =   0   'False
            FillStyle       =   0
            RightToLeft     =   0   'False
            PictureType     =   0
            TabBehavior     =   0
            OwnerDraw       =   0
            Editable        =   2
            ShowComboButton =   1
            WordWrap        =   -1  'True
            TextStyle       =   0
            TextStyleFixed  =   0
            OleDragMode     =   0
            OleDropMode     =   0
            DataMode        =   0
            VirtualData     =   -1  'True
            DataMember      =   ""
            ComboSearch     =   3
            AutoSizeMouse   =   -1  'True
            FrozenRows      =   0
            FrozenCols      =   0
            AllowUserFreezing=   0
            BackColorFrozen =   0
            ForeColorFrozen =   0
            WallPaperAlignment=   9
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   24
            Begin VB.CommandButton cmdDepartSelect 
               Caption         =   "…"
               Height          =   240
               Left            =   4260
               TabIndex        =   539
               Top             =   255
               Visible         =   0   'False
               Width           =   300
            End
         End
         Begin VB.Label lblEpr 
            AutoSize        =   -1  'True
            Caption         =   "使用图片签名科室设置"
            Height          =   180
            Index           =   3
            Left            =   90
            TabIndex        =   541
            Top             =   810
            Width           =   1800
         End
         Begin VB.Label lblEpr 
            AutoSize        =   -1  'True
            Caption         =   "指定签名图片高度     像素"
            Height          =   180
            Index           =   2
            Left            =   2430
            TabIndex        =   540
            Top             =   345
            Width           =   2250
         End
      End
      Begin VB.Frame fraEprWrite 
         Caption         =   "病历书写"
         Height          =   3150
         Left            =   255
         TabIndex        =   225
         Top             =   120
         Width           =   4650
         Begin VB.CheckBox chk 
            Caption         =   "填写诊断弹出的传染病报告卡需强制填写"
            Height          =   285
            Index           =   182
            Left            =   75
            TabIndex        =   644
            Top             =   2700
            Width           =   3660
         End
         Begin VB.CheckBox chk 
            Caption         =   "插入入/出院诊断时同步更新首页"
            Height          =   285
            Index           =   99
            Left            =   75
            TabIndex        =   232
            Top             =   2310
            Width           =   3420
         End
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   2
            Left            =   1410
            Style           =   2  'Dropdown List
            TabIndex        =   231
            Top             =   1853
            Width           =   2565
         End
         Begin VB.CheckBox chk 
            Caption         =   "将签名级别作为前缀加入(&P)"
            Height          =   225
            Index           =   95
            Left            =   75
            TabIndex        =   229
            Top             =   1124
            Width           =   2565
         End
         Begin VB.CheckBox chk 
            Caption         =   "显示手签位置(&H)"
            Height          =   240
            Index           =   96
            Left            =   75
            TabIndex        =   230
            Top             =   1506
            Width           =   1695
         End
         Begin VB.OptionButton optSign 
            Caption         =   "姓名"
            Height          =   180
            Index           =   0
            Left            =   1095
            TabIndex        =   227
            Top             =   330
            Value           =   -1  'True
            Width           =   1290
         End
         Begin VB.OptionButton optSign 
            Caption         =   "签名"
            Height          =   180
            Index           =   1
            Left            =   2535
            TabIndex        =   228
            Top             =   330
            Width           =   915
         End
         Begin VB.CheckBox chk 
            Caption         =   "签名移位(&S)"
            Height          =   285
            Index           =   94
            Left            =   75
            TabIndex        =   226
            Top             =   682
            Width           =   1305
         End
         Begin VB.Label lblEpr 
            AutoSize        =   -1  'True
            Caption         =   "签名时间(&T)"
            Height          =   180
            Index           =   1
            Left            =   75
            TabIndex        =   234
            Top             =   1913
            Width           =   990
         End
         Begin VB.Label lblEpr 
            Caption         =   "签名显示"
            Height          =   210
            Index           =   0
            Left            =   75
            TabIndex        =   233
            Top             =   315
            Width           =   870
         End
      End
      Begin VB.Frame fraEprIn 
         Caption         =   "住院病历"
         Height          =   3705
         Index           =   0
         Left            =   255
         TabIndex        =   210
         Top             =   3420
         Width           =   4650
         Begin VB.TextBox txt 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   180
            Index           =   12
            Left            =   2940
            TabIndex        =   221
            Text            =   "7"
            Top             =   2805
            Width           =   360
         End
         Begin VB.OptionButton optEprRead 
            Caption         =   "不连续预览，选中一份文件读一次。"
            Height          =   195
            Index           =   1
            Left            =   345
            TabIndex        =   220
            Top             =   2470
            Width           =   4260
         End
         Begin VB.OptionButton optEprRead 
            Caption         =   "连续预览，首次读取全部共享病历，后续只定位。"
            Height          =   195
            Index           =   0
            Left            =   345
            TabIndex        =   219
            Top             =   2150
            Value           =   -1  'True
            Width           =   4260
         End
         Begin VB.Frame fraEprsplit 
            Height          =   30
            Left            =   0
            TabIndex        =   218
            Top             =   1725
            Width           =   4635
         End
         Begin VB.CheckBox chk 
            Caption         =   "共享病历必须先书写被共享病历"
            Height          =   180
            Index           =   102
            Left            =   75
            TabIndex        =   217
            Top             =   1080
            Width           =   3720
         End
         Begin VB.CheckBox chk 
            Caption         =   "自动显示新增面板"
            Height          =   180
            Index           =   101
            Left            =   75
            TabIndex        =   216
            Top             =   705
            Width           =   3720
         End
         Begin VB.CheckBox chk 
            Caption         =   "(转科后要求书写)的共享病历另起一页打印"
            Height          =   330
            Index           =   100
            Left            =   75
            TabIndex        =   215
            Top             =   300
            Width           =   3720
         End
         Begin VB.TextBox txt 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   180
            Index           =   11
            Left            =   2595
            TabIndex        =   214
            Text            =   "13"
            Top             =   1440
            Width           =   360
         End
         Begin VB.OptionButton optEprRead 
            Caption         =   "连续预览，读取选中文件前后    天的共享病历。"
            Height          =   195
            Index           =   2
            Left            =   345
            TabIndex        =   222
            Top             =   2790
            Width           =   4260
         End
         Begin VB.Line lnEpr 
            Index           =   0
            X1              =   2595
            X2              =   2985
            Y1              =   1620
            Y2              =   1620
         End
         Begin VB.Label lblEpr 
            Caption         =   "列表产生滚动条后，共享病历超    行自动折叠"
            Height          =   240
            Index           =   4
            Left            =   75
            TabIndex        =   224
            Top             =   1440
            Width           =   3780
         End
         Begin VB.Line lnEpr 
            Index           =   1
            X1              =   2940
            X2              =   3330
            Y1              =   2985
            Y2              =   2985
         End
         Begin VB.Label lblEpr 
            Caption         =   "共享病历连读预览"
            Height          =   180
            Index           =   5
            Left            =   75
            TabIndex        =   223
            Top             =   1845
            Width           =   1650
         End
      End
   End
   Begin VB.PictureBox picPar 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   7815
      Index           =   8
      Left            =   2400
      ScaleHeight     =   7785
      ScaleWidth      =   10065
      TabIndex        =   155
      Top             =   30
      Width           =   10095
      Begin VB.CommandButton cmdPatiSign 
         Caption         =   "病人标记设置(&P)"
         Height          =   350
         Left            =   6555
         TabIndex        =   566
         Top             =   195
         Width           =   1680
      End
      Begin VB.Frame Frame11 
         Caption         =   "医技工作站"
         Height          =   2220
         Left            =   120
         TabIndex        =   209
         Top             =   5355
         Width           =   6240
         Begin VB.CheckBox chk 
            Caption         =   "血透室书写新版护理记录"
            Height          =   195
            Index           =   150
            Left            =   240
            TabIndex        =   518
            Top             =   1320
            Width           =   2520
         End
         Begin VB.Frame Frame15 
            BorderStyle     =   0  'None
            Caption         =   "Frame7"
            Height          =   255
            Left            =   1680
            TabIndex        =   513
            Top             =   1690
            Width           =   2895
            Begin VB.OptionButton opt病人过滤 
               Caption         =   "执行时间"
               Height          =   180
               Index           =   0
               Left            =   0
               TabIndex        =   515
               Top             =   0
               Value           =   -1  'True
               Width           =   1095
            End
            Begin VB.OptionButton opt病人过滤 
               Caption         =   "发送时间"
               Height          =   180
               Index           =   1
               Left            =   1200
               TabIndex        =   514
               Top             =   0
               Width           =   1095
            End
         End
         Begin VB.CheckBox chk 
            Caption         =   "填写皮试结果时验证身份"
            Height          =   195
            Index           =   92
            Left            =   240
            TabIndex        =   212
            Top             =   660
            Width           =   2280
         End
         Begin VB.CheckBox chk 
            Caption         =   "允许未收费病人完成执行"
            Height          =   195
            Index           =   91
            Left            =   240
            TabIndex        =   211
            Top             =   360
            Width           =   2280
         End
         Begin VB.CheckBox chk 
            Caption         =   "执行报到时收费或记账审核"
            Height          =   195
            Index           =   93
            Left            =   240
            TabIndex        =   213
            Top             =   975
            Width           =   2520
         End
         Begin VB.Label lbl病人过滤 
            Caption         =   "病人过滤条件按"
            Height          =   255
            Left            =   240
            TabIndex        =   516
            Top             =   1680
            Width           =   1335
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "门诊医生站"
         Height          =   1920
         Left            =   120
         TabIndex        =   178
         Top             =   120
         Width           =   6240
         Begin VB.TextBox txt 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   180
            Index           =   4
            Left            =   4725
            TabIndex        =   185
            Text            =   "0"
            Top             =   390
            Width           =   525
         End
         Begin VB.OptionButton opt接诊控制 
            Caption         =   "提示"
            Height          =   240
            Index           =   2
            Left            =   3060
            TabIndex        =   183
            Top             =   375
            Width           =   750
         End
         Begin VB.OptionButton opt接诊控制 
            Caption         =   "禁止"
            Height          =   240
            Index           =   1
            Left            =   2310
            TabIndex        =   182
            Top             =   375
            Width           =   855
         End
         Begin VB.OptionButton opt接诊控制 
            Caption         =   "不禁止"
            Height          =   240
            Index           =   0
            Left            =   1425
            TabIndex        =   181
            Top             =   375
            Width           =   870
         End
         Begin VB.CheckBox chk 
            Caption         =   "只接收已经分诊的病人"
            Height          =   195
            Index           =   78
            Left            =   240
            TabIndex        =   190
            Top             =   1611
            Width           =   2100
         End
         Begin VB.TextBox txt 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   180
            IMEMode         =   3  'DISABLE
            Index           =   5
            Left            =   885
            MaxLength       =   4
            TabIndex        =   187
            Text            =   "180"
            Top             =   690
            Width           =   465
         End
         Begin VB.Frame fraLine 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   15
            Left            =   885
            TabIndex        =   179
            Top             =   870
            Width           =   465
         End
         Begin VB.CheckBox chk 
            Caption         =   "医生呼叫人数限制含回诊病人"
            Height          =   180
            Index           =   72
            Left            =   240
            TabIndex        =   188
            Top             =   1020
            Value           =   1  'Checked
            Width           =   2685
         End
         Begin VB.CheckBox chk 
            Caption         =   "医生主动呼叫后才允许在队列中接诊"
            Height          =   195
            Index           =   73
            Left            =   240
            TabIndex        =   189
            Top             =   1308
            Value           =   1  'Checked
            Width           =   3360
         End
         Begin VB.Label lblReceptionTime 
            AutoSize        =   -1  'True
            Caption         =   "允许提前       分钟接诊"
            Height          =   180
            Left            =   3915
            TabIndex        =   184
            Top             =   405
            Width           =   2070
         End
         Begin VB.Line line 
            X1              =   4635
            X2              =   5340
            Y1              =   600
            Y2              =   600
         End
         Begin VB.Label lblReceptionMode 
            Caption         =   "病人接诊控制"
            Height          =   270
            Left            =   240
            TabIndex        =   180
            Top             =   380
            Width           =   1185
         End
         Begin VB.Label lblRefresh 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "每隔      秒自动刷新候诊/转诊病人清单"
            Height          =   180
            Left            =   480
            TabIndex        =   186
            Top             =   705
            Width           =   3330
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "住院护士站"
         Height          =   1635
         Left            =   120
         TabIndex        =   195
         Top             =   3465
         Width           =   6240
         Begin VB.CommandButton cmdLink 
            Caption         =   "验证"
            Height          =   300
            Left            =   4530
            TabIndex        =   208
            TabStop         =   0   'False
            Top             =   1185
            Width           =   570
         End
         Begin VB.TextBox txt 
            Height          =   300
            Index           =   22
            Left            =   3870
            MaxLength       =   4
            TabIndex        =   207
            Top             =   1200
            Width           =   600
         End
         Begin VB.TextBox txt 
            Height          =   300
            Index           =   21
            Left            =   2085
            MaxLength       =   15
            TabIndex        =   205
            Top             =   1185
            Width           =   1305
         End
         Begin VB.OptionButton optNewCard 
            Caption         =   "顺序号+床位号"
            Height          =   180
            Index           =   0
            Left            =   1860
            TabIndex        =   202
            Top             =   885
            Width           =   1515
         End
         Begin VB.OptionButton optNewCard 
            Caption         =   "顺序号+床位编制编号+床位号"
            Height          =   180
            Index           =   1
            Left            =   3435
            TabIndex        =   203
            Top             =   885
            Value           =   -1  'True
            Width           =   2715
         End
         Begin VB.CheckBox chk 
            Caption         =   "卡片余额将担保金额计算在内(新版专用)"
            Height          =   195
            Index           =   133
            Left            =   240
            TabIndex        =   200
            Top             =   555
            Width           =   3900
         End
         Begin VB.Frame Frame4 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   15
            Index           =   1
            Left            =   900
            TabIndex        =   199
            Top             =   420
            Width           =   300
         End
         Begin VB.TextBox txt 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   180
            IMEMode         =   3  'DISABLE
            Index           =   7
            Left            =   915
            MaxLength       =   3
            TabIndex        =   198
            Text            =   "1"
            Top             =   240
            Width           =   300
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "端口"
            Height          =   180
            Left            =   3465
            TabIndex        =   206
            Top             =   1230
            Width           =   360
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "整体护理服务器IP地址"
            Height          =   180
            Left            =   240
            TabIndex        =   204
            Top             =   1230
            Width           =   1800
         End
         Begin VB.Label Label12 
            Caption         =   "护士站床位排序按"
            Height          =   255
            Left            =   240
            TabIndex        =   201
            Top             =   870
            Width           =   1440
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "显示    天内的病案审查反馈数"
            Height          =   180
            Left            =   525
            TabIndex        =   197
            Top             =   255
            Width           =   2520
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "住院医生站"
         Height          =   1095
         Left            =   120
         TabIndex        =   191
         Top             =   2145
         Width           =   6240
         Begin VB.CheckBox chk 
            Caption         =   "对拥有全院病人权限的操作者，不显示没有床位的科室或病区"
            Height          =   255
            Index           =   68
            Left            =   240
            TabIndex        =   517
            Top             =   600
            Width           =   5640
         End
         Begin VB.TextBox txt 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   180
            IMEMode         =   3  'DISABLE
            Index           =   6
            Left            =   915
            MaxLength       =   3
            TabIndex        =   193
            Text            =   "1"
            Top             =   240
            Width           =   300
         End
         Begin VB.Frame Frame4 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   15
            Index           =   2
            Left            =   900
            TabIndex        =   194
            Top             =   420
            Width           =   300
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "显示    天内的病案审查反馈数"
            Height          =   180
            Left            =   525
            TabIndex        =   192
            Top             =   255
            Width           =   2520
         End
      End
   End
   Begin VB.PictureBox picPar 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   7695
      Index           =   6
      Left            =   2400
      ScaleHeight     =   7665
      ScaleWidth      =   10185
      TabIndex        =   136
      Top             =   0
      Width           =   10215
      Begin VB.Frame Frame6 
         Caption         =   "住院路径流程控制"
         Height          =   5895
         Left            =   120
         TabIndex        =   595
         Top             =   120
         Width           =   4335
         Begin VB.CheckBox chk 
            Caption         =   "门诊诊断不作为导入临床路径的诊断依据"
            Height          =   255
            Index           =   179
            Left            =   240
            TabIndex        =   634
            Top             =   5040
            Width           =   3975
         End
         Begin VB.CheckBox chk 
            Caption         =   "允许提前生成明天的路径项目"
            Height          =   180
            Index           =   47
            Left            =   240
            TabIndex        =   613
            Top             =   2880
            Width           =   3015
         End
         Begin VB.CheckBox chk 
            Caption         =   "允许前一天不评估就生成今天的路径项目"
            Height          =   180
            Index           =   46
            Left            =   240
            TabIndex        =   612
            Top             =   2520
            Width           =   3615
         End
         Begin VB.TextBox txt 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   180
            IMEMode         =   3  'DISABLE
            Index           =   2
            Left            =   3000
            MaxLength       =   3
            TabIndex        =   611
            Text            =   "30"
            Top             =   1080
            Width           =   300
         End
         Begin VB.CheckBox chk 
            Caption         =   "未评估时允许添加医嘱到昨天"
            Height          =   180
            Index           =   45
            Left            =   240
            TabIndex        =   610
            Top             =   2160
            Width           =   2775
         End
         Begin VB.Frame Frame1 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   15
            Index           =   0
            Left            =   3120
            TabIndex        =   609
            Top             =   1620
            Width           =   300
         End
         Begin VB.TextBox txt 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   180
            IMEMode         =   3  'DISABLE
            Index           =   3
            Left            =   3120
            MaxLength       =   1
            TabIndex        =   608
            Text            =   "1"
            Top             =   1440
            Width           =   300
         End
         Begin VB.CheckBox chk 
            Caption         =   "医技科室下达医嘱不在路径表上记录"
            Height          =   180
            Index           =   44
            Left            =   240
            TabIndex        =   607
            Top             =   1800
            Width           =   3375
         End
         Begin VB.Frame Frame2 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   15
            Left            =   2925
            TabIndex        =   606
            Top             =   1275
            Width           =   420
         End
         Begin VB.Frame fraPathExe 
            Height          =   615
            Left            =   240
            TabIndex        =   602
            Top             =   360
            Width           =   3135
            Begin VB.CheckBox chk 
               Caption         =   "护士"
               Height          =   255
               Index           =   43
               Left            =   1320
               TabIndex        =   605
               Top             =   240
               Width           =   735
            End
            Begin VB.CheckBox chk 
               Caption         =   "医生"
               Height          =   255
               Index           =   42
               Left            =   480
               TabIndex        =   604
               Top             =   240
               Width           =   735
            End
            Begin VB.CheckBox chk 
               Caption         =   "启用路径执行环节"
               Height          =   180
               Index           =   41
               Left            =   240
               TabIndex        =   603
               Top             =   0
               Width           =   1815
            End
         End
         Begin VB.CheckBox chk 
            Caption         =   "匹配时期效不同算路径外项目"
            Height          =   180
            Index           =   84
            Left            =   240
            TabIndex        =   601
            Top             =   3240
            Width           =   3015
         End
         Begin VB.CommandButton cmdPathSortSet 
            Caption         =   "路径项目生成顺序设置(&S)"
            Height          =   350
            Left            =   480
            TabIndex        =   600
            Top             =   5400
            Width           =   2415
         End
         Begin VB.CheckBox chk 
            Caption         =   "出院后不允许取消完成路径"
            Height          =   180
            Index           =   153
            Left            =   240
            TabIndex        =   599
            Top             =   4320
            Width           =   3015
         End
         Begin VB.CheckBox chk 
            Caption         =   "药品医嘱不匹配为路径外项目"
            Height          =   180
            Index           =   57
            Left            =   240
            TabIndex        =   598
            Top             =   3600
            Width           =   3015
         End
         Begin VB.CheckBox chk 
            Caption         =   "启用药剂科、医务科双审核模式"
            Height          =   255
            Index           =   58
            Left            =   240
            TabIndex        =   597
            Top             =   4680
            Width           =   3255
         End
         Begin VB.CheckBox chk 
            Caption         =   "药品医嘱相同分类不算路径外项目"
            Height          =   180
            Index           =   134
            Left            =   240
            TabIndex        =   596
            Top             =   3960
            Width           =   3015
         End
         Begin VB.Label lblSendAdvice 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "生成路径时，医嘱允许超过当前时间    天"
            Height          =   180
            Left            =   240
            TabIndex        =   615
            Top             =   1440
            Width           =   3420
         End
         Begin VB.Label lbl中药味数 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "中药配方允许修改的中药味数上限    %"
            Height          =   180
            Left            =   240
            TabIndex        =   614
            Top             =   1080
            Width           =   3150
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "门诊路径流程控制"
         Height          =   5535
         Left            =   5040
         TabIndex        =   587
         Top             =   120
         Width           =   4335
         Begin VB.Frame Frame17 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   15
            Left            =   2925
            TabIndex        =   592
            Top             =   960
            Width           =   420
         End
         Begin VB.TextBox txt 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   180
            IMEMode         =   3  'DISABLE
            Index           =   23
            Left            =   3120
            MaxLength       =   1
            TabIndex        =   591
            Text            =   "1"
            Top             =   1080
            Width           =   300
         End
         Begin VB.Frame Frame1 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   15
            Index           =   7
            Left            =   3120
            TabIndex        =   590
            Top             =   1260
            Width           =   300
         End
         Begin VB.TextBox txt 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   180
            IMEMode         =   3  'DISABLE
            Index           =   24
            Left            =   3000
            MaxLength       =   3
            TabIndex        =   589
            Text            =   "30"
            Top             =   720
            Width           =   300
         End
         Begin VB.CheckBox chk 
            Caption         =   "启用路径执行环节"
            Height          =   180
            Index           =   169
            Left            =   240
            TabIndex        =   588
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "中药配方允许修改的中药味数上限    %"
            Height          =   180
            Left            =   240
            TabIndex        =   594
            Top             =   720
            Width           =   3150
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "生成路径时，医嘱允许超过当前时间    天"
            Height          =   180
            Left            =   240
            TabIndex        =   593
            Top             =   1080
            Width           =   3420
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "住院路径打印控制"
         Height          =   1455
         Left            =   120
         TabIndex        =   577
         Top             =   6120
         Width           =   4335
         Begin VB.Frame Frame12 
            BorderStyle     =   0  'None
            Caption         =   "Frame12"
            Height          =   255
            Left            =   1920
            TabIndex        =   627
            Top             =   1080
            Width           =   2295
            Begin VB.OptionButton optPrintWay 
               Caption         =   "报表方式"
               Height          =   180
               Index           =   1
               Left            =   1200
               TabIndex        =   629
               Top             =   0
               Width           =   1095
            End
            Begin VB.OptionButton optPrintWay 
               Caption         =   "表格方式"
               Height          =   180
               Index           =   0
               Left            =   0
               TabIndex        =   628
               Top             =   0
               Value           =   -1  'True
               Width           =   1095
            End
         End
         Begin VB.Frame fraRule 
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   1800
            TabIndex        =   582
            Top             =   360
            Width           =   2415
            Begin VB.OptionButton opt路径打印规则 
               Caption         =   "按阶段打印"
               Height          =   180
               Index           =   1
               Left            =   1080
               TabIndex        =   584
               Top             =   0
               Width           =   1215
            End
            Begin VB.OptionButton opt路径打印规则 
               Caption         =   "按天打印"
               Height          =   180
               Index           =   0
               Left            =   0
               TabIndex        =   583
               Top             =   0
               Value           =   -1  'True
               Width           =   1095
            End
         End
         Begin VB.OptionButton opt每页路径打印天数 
            Caption         =   "3天"
            Height          =   180
            Index           =   3
            Left            =   2880
            TabIndex        =   581
            Top             =   720
            Width           =   615
         End
         Begin VB.OptionButton opt每页路径打印天数 
            Caption         =   "2天"
            Height          =   180
            Index           =   0
            Left            =   0
            TabIndex        =   580
            Top             =   -9000
            Width           =   615
         End
         Begin VB.OptionButton opt每页路径打印天数 
            Caption         =   "2天"
            Height          =   180
            Index           =   1
            Left            =   3480
            TabIndex        =   579
            Top             =   -9000
            Width           =   615
         End
         Begin VB.OptionButton opt每页路径打印天数 
            Caption         =   "2天"
            Height          =   180
            Index           =   2
            Left            =   2280
            TabIndex        =   578
            Top             =   720
            Width           =   615
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "路径表单打印方式："
            Height          =   180
            Left            =   240
            TabIndex        =   626
            Top             =   1080
            Width           =   1620
         End
         Begin VB.Label lblPrtRule 
            Caption         =   "路径表单打印规则："
            Height          =   255
            Left            =   240
            TabIndex        =   586
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label lblPrintDays 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "路径表单每页打印的天数"
            Height          =   180
            Left            =   240
            TabIndex        =   585
            Top             =   720
            Width           =   1980
         End
      End
      Begin VB.Frame Frame16 
         Caption         =   "门诊路径打印控制"
         Height          =   1455
         Left            =   5040
         TabIndex        =   567
         Top             =   5760
         Width           =   4335
         Begin VB.OptionButton opt门诊路径打印天数 
            Caption         =   "2天"
            Height          =   180
            Index           =   1
            Left            =   -9000
            TabIndex        =   617
            Top             =   0
            Width           =   615
         End
         Begin VB.OptionButton opt门诊路径打印天数 
            Caption         =   "2天"
            Height          =   180
            Index           =   0
            Left            =   -9000
            TabIndex        =   616
            Top             =   0
            Width           =   615
         End
         Begin VB.OptionButton opt每页路径打印天数 
            Caption         =   "2天"
            Height          =   180
            Index           =   4
            Left            =   3480
            TabIndex        =   568
            Top             =   -9000
            Width           =   615
         End
         Begin VB.Frame Frame18 
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   1800
            TabIndex        =   572
            Top             =   360
            Width           =   2415
            Begin VB.OptionButton opt门诊路径打印规则 
               Caption         =   "按阶段打印"
               Height          =   180
               Index           =   1
               Left            =   1080
               TabIndex        =   574
               Top             =   0
               Width           =   1215
            End
            Begin VB.OptionButton opt门诊路径打印规则 
               Caption         =   "按天打印"
               Height          =   180
               Index           =   0
               Left            =   0
               TabIndex        =   573
               Top             =   0
               Value           =   -1  'True
               Width           =   1095
            End
         End
         Begin VB.OptionButton opt门诊路径打印天数 
            Caption         =   "3天"
            Height          =   180
            Index           =   3
            Left            =   2880
            TabIndex        =   571
            Top             =   720
            Width           =   615
         End
         Begin VB.OptionButton opt门诊路径打印天数 
            Caption         =   "2天"
            Height          =   180
            Index           =   2
            Left            =   2280
            TabIndex        =   570
            Top             =   720
            Width           =   615
         End
         Begin VB.OptionButton opt每页路径打印天数 
            Caption         =   "2天"
            Height          =   180
            Index           =   5
            Left            =   3600
            TabIndex        =   569
            Top             =   -9000
            Width           =   615
         End
         Begin VB.Label Label20 
            Caption         =   "路径表单打印规则："
            Height          =   255
            Left            =   240
            TabIndex        =   576
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "路径表单每页打印的天数"
            Height          =   180
            Left            =   240
            TabIndex        =   575
            Top             =   720
            Width           =   1980
         End
      End
   End
   Begin VB.PictureBox picPar 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   7455
      Index           =   4
      Left            =   2400
      ScaleHeight     =   7425
      ScaleWidth      =   10065
      TabIndex        =   113
      Top             =   15
      Width           =   10095
      Begin VB.Frame fraAdvicePrint 
         Caption         =   "医嘱单打印"
         Height          =   2235
         Left            =   4320
         TabIndex        =   141
         Top             =   240
         Width           =   5625
         Begin VB.TextBox txt 
            Height          =   270
            Index           =   26
            Left            =   4050
            MaxLength       =   100
            TabIndex        =   637
            Top             =   1860
            Width           =   555
         End
         Begin VB.TextBox txt 
            Height          =   270
            Index           =   8
            Left            =   2685
            MaxLength       =   100
            TabIndex        =   636
            Top             =   1860
            Width           =   555
         End
         Begin VB.CheckBox chk 
            Caption         =   "术后"
            Height          =   255
            Index           =   172
            Left            =   2655
            TabIndex        =   631
            Top             =   1215
            Width           =   660
         End
         Begin VB.CheckBox chk 
            Caption         =   "重整"
            Height          =   255
            Index           =   173
            Left            =   3420
            TabIndex        =   620
            Top             =   915
            Width           =   660
         End
         Begin VB.CheckBox chk 
            Caption         =   "术后"
            Height          =   255
            Index           =   171
            Left            =   2655
            TabIndex        =   619
            Top             =   915
            Width           =   660
         End
         Begin VB.PictureBox picOptFra 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   200
            Index           =   1
            Left            =   2595
            ScaleHeight     =   195
            ScaleWidth      =   2895
            TabIndex        =   555
            Top             =   300
            Width           =   2900
            Begin VB.OptionButton optPrintDruUse 
               Caption         =   "一并给药打印"
               Height          =   180
               Index           =   2
               Left            =   1545
               TabIndex        =   558
               Top             =   0
               Width           =   1400
            End
            Begin VB.OptionButton optPrintDruUse 
               Caption         =   "打印"
               Height          =   180
               Index           =   1
               Left            =   870
               TabIndex        =   557
               Top             =   0
               Width           =   720
            End
            Begin VB.OptionButton optPrintDruUse 
               Caption         =   "不打印"
               Height          =   180
               Index           =   0
               Left            =   0
               TabIndex        =   556
               Top             =   0
               Value           =   -1  'True
               Width           =   855
            End
         End
         Begin VB.PictureBox picOptFra 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   200
            Index           =   0
            Left            =   2715
            ScaleHeight     =   195
            ScaleWidth      =   2700
            TabIndex        =   551
            Top             =   615
            Width           =   2700
            Begin VB.OptionButton opt转科死亡出院医嘱 
               Caption         =   "以上两者"
               ForeColor       =   &H00000000&
               Height          =   180
               Index           =   2
               Left            =   1725
               TabIndex        =   552
               Top             =   0
               Width           =   1150
            End
            Begin VB.OptionButton opt转科死亡出院医嘱 
               Caption         =   "临嘱单"
               ForeColor       =   &H00000000&
               Height          =   180
               Index           =   1
               Left            =   855
               TabIndex        =   553
               Top             =   0
               Width           =   900
            End
            Begin VB.OptionButton opt转科死亡出院医嘱 
               Caption         =   "长嘱单"
               ForeColor       =   &H00000000&
               Height          =   180
               Index           =   0
               Left            =   0
               TabIndex        =   554
               Top             =   0
               Value           =   -1  'True
               Width           =   900
            End
         End
         Begin VB.CheckBox chk 
            Caption         =   "转科"
            Height          =   285
            Index           =   162
            Left            =   1920
            TabIndex        =   547
            Top             =   1215
            Width           =   660
         End
         Begin VB.CheckBox chk 
            Caption         =   "转科"
            Height          =   255
            Index           =   59
            Left            =   1920
            TabIndex        =   143
            Top             =   915
            Width           =   660
         End
         Begin VB.CheckBox chk 
            Caption         =   "长嘱单"
            Height          =   285
            Index           =   60
            Left            =   3420
            TabIndex        =   142
            Top             =   1515
            Width           =   850
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "中药医嘱单行显示字数  长嘱单         临嘱单"
            Height          =   180
            Left            =   150
            TabIndex        =   635
            Top             =   1875
            Width           =   3870
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "转科换页后在首行打印""重开医嘱""字样"
            Height          =   180
            Left            =   135
            TabIndex        =   633
            Top             =   1560
            Width           =   3060
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "临嘱单另起一页打印"
            Height          =   180
            Left            =   135
            TabIndex        =   632
            Top             =   1245
            Width           =   1620
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "长嘱单另起一页打印"
            Height          =   180
            Left            =   135
            TabIndex        =   618
            Top             =   930
            Width           =   1620
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "西药、成药用法单独打印一行"
            Height          =   180
            Left            =   135
            TabIndex        =   550
            Top             =   285
            Width           =   2340
         End
         Begin VB.Label Label1 
            Caption         =   "转科、出院、死亡医嘱打印位置"
            Height          =   255
            Left            =   135
            TabIndex        =   144
            Top             =   600
            Width           =   2535
         End
      End
      Begin VB.Frame fra留存 
         Caption         =   "需要药品留存登记的给药途径"
         Height          =   4770
         Left            =   4320
         TabIndex        =   137
         Top             =   2520
         Width           =   5625
         Begin VB.ListBox lst 
            Columns         =   3
            ForeColor       =   &H80000012&
            Height          =   4050
            IMEMode         =   3  'DISABLE
            Index           =   7
            ItemData        =   "frmParClinic.frx":3002E
            Left            =   120
            List            =   "frmParClinic.frx":30035
            Style           =   1  'Checkbox
            TabIndex        =   140
            Top             =   240
            Width           =   5355
         End
         Begin VB.CommandButton cmd药品留存给药途径 
            Caption         =   "全选"
            Height          =   300
            Index           =   0
            Left            =   3600
            TabIndex        =   139
            Top             =   4360
            Width           =   900
         End
         Begin VB.CommandButton cmd药品留存给药途径 
            Caption         =   "全清"
            Height          =   300
            Index           =   1
            Left            =   4560
            TabIndex        =   138
            Top             =   4360
            Width           =   900
         End
      End
      Begin VB.Frame fra执行 
         Caption         =   " 医嘱执行 "
         Height          =   2235
         Left            =   240
         TabIndex        =   83
         Top             =   240
         Width           =   3975
         Begin VB.CheckBox chk 
            Caption         =   "皮试"
            Height          =   200
            Index           =   61
            Left            =   825
            TabIndex        =   502
            Top             =   280
            Width           =   690
         End
         Begin VB.CheckBox chk 
            Caption         =   "操作员只限以本人身份登记"
            Height          =   210
            Index           =   32
            Left            =   120
            TabIndex        =   500
            Top             =   1560
            Width           =   2520
         End
         Begin VB.CheckBox chk 
            Caption         =   "执行单打印时逐个病人换页打印"
            Height          =   195
            Index           =   55
            Left            =   120
            TabIndex        =   148
            Top             =   1275
            Width           =   3540
         End
         Begin VB.CheckBox chk 
            Caption         =   "指定医嘱在其他科室执行"
            Height          =   240
            Index           =   62
            Left            =   120
            TabIndex        =   88
            Top             =   945
            Value           =   1  'Checked
            Width           =   2280
         End
         Begin VB.CheckBox chk 
            Caption         =   "输血"
            Height          =   200
            Index           =   74
            Left            =   120
            TabIndex        =   84
            Top             =   280
            Width           =   690
         End
         Begin VB.TextBox txtUD 
            Alignment       =   2  'Center
            BackColor       =   &H80000004&
            Enabled         =   0   'False
            Height          =   270
            IMEMode         =   3  'DISABLE
            Index           =   0
            Left            =   1150
            MaxLength       =   3
            TabIndex        =   86
            Text            =   "999"
            Top             =   555
            Width           =   510
         End
         Begin VB.CheckBox chk 
            Caption         =   "允许取消"
            Height          =   200
            Index           =   87
            Left            =   120
            TabIndex        =   85
            Top             =   585
            Width           =   1035
         End
         Begin MSComCtl2.UpDown ud 
            Height          =   270
            Index           =   0
            Left            =   1680
            TabIndex        =   87
            TabStop         =   0   'False
            Top             =   555
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   476
            _Version        =   393216
            Value           =   1
            BuddyControl    =   "txtUD(0)"
            BuddyDispid     =   196669
            BuddyIndex      =   0
            OrigLeft        =   2400
            OrigTop         =   1380
            OrigRight       =   2655
            OrigBottom      =   1680
            Max             =   999
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   0   'False
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "医嘱执行后需要核对"
            Height          =   180
            Left            =   1665
            TabIndex        =   501
            Top             =   285
            Width           =   1620
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "天内的执行操作"
            Height          =   180
            Index           =   1
            Left            =   1995
            TabIndex        =   117
            Top             =   600
            Width           =   1260
         End
      End
      Begin VB.Frame fra超期收回 
         Caption         =   " 超期收回 "
         Height          =   4770
         Left            =   240
         TabIndex        =   89
         Top             =   2520
         Width           =   3975
         Begin VB.CheckBox chk 
            Caption         =   "收回时医嘱列表只显示当前病区的医嘱"
            Height          =   255
            Index           =   67
            Left            =   210
            TabIndex        =   145
            Top             =   1200
            Width           =   3615
         End
         Begin VB.OptionButton opt超期费用收回 
            Caption         =   "销帐申请"
            Height          =   255
            Index           =   0
            Left            =   2640
            TabIndex        =   92
            Top             =   300
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton opt超期费用收回 
            Caption         =   "负数记帐"
            Height          =   255
            Index           =   1
            Left            =   1440
            TabIndex        =   91
            Top             =   300
            Width           =   1095
         End
         Begin VB.CheckBox chk 
            Caption         =   "超期收回时自动审核本科执行的销帐申请"
            Height          =   195
            Index           =   36
            Left            =   210
            TabIndex        =   93
            Top             =   600
            Width           =   3680
         End
         Begin VB.CheckBox chk 
            Caption         =   "确认停止后自动执行超期收回"
            Height          =   195
            Index           =   37
            Left            =   210
            TabIndex        =   94
            Top             =   900
            Width           =   3180
         End
         Begin VB.ListBox lst 
            Columns         =   3
            ForeColor       =   &H80000012&
            Height          =   2790
            IMEMode         =   3  'DISABLE
            Index           =   6
            ItemData        =   "frmParClinic.frx":3003E
            Left            =   210
            List            =   "frmParClinic.frx":30045
            Style           =   1  'Checkbox
            TabIndex        =   96
            Top             =   1860
            Width           =   3525
         End
         Begin VB.Label lblRoll 
            Caption         =   "费用收回模式"
            Height          =   255
            Left            =   240
            TabIndex        =   90
            Top             =   300
            Width           =   1095
         End
         Begin VB.Label lblSend 
            Caption         =   "以下发药方式的西药一但发药就不收回"
            Height          =   255
            Left            =   240
            TabIndex        =   95
            Top             =   1590
            Width           =   3255
         End
      End
   End
   Begin VB.PictureBox picPar 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   7455
      Index           =   2
      Left            =   2400
      ScaleHeight     =   7425
      ScaleWidth      =   10065
      TabIndex        =   105
      Top             =   0
      Width           =   10095
      Begin VB.Frame fra诊断 
         Caption         =   "门诊发送以下类别时检查诊断填写"
         Height          =   1530
         Left            =   5160
         TabIndex        =   55
         Top             =   3600
         Width           =   4695
         Begin VB.CommandButton cmd门诊发送检查诊断 
            Caption         =   "全清"
            Height          =   300
            Index           =   1
            Left            =   3600
            TabIndex        =   58
            Top             =   720
            Width           =   900
         End
         Begin VB.CommandButton cmd门诊发送检查诊断 
            Caption         =   "全选"
            Height          =   300
            Index           =   0
            Left            =   3600
            TabIndex        =   57
            Top             =   360
            Width           =   900
         End
         Begin VB.ListBox lst 
            Columns         =   3
            ForeColor       =   &H80000012&
            Height          =   900
            IMEMode         =   3  'DISABLE
            Index           =   1
            Left            =   240
            Style           =   1  'Checkbox
            TabIndex        =   56
            Top             =   360
            Width           =   3300
         End
      End
      Begin VB.Frame fra门诊发送 
         Caption         =   "门诊发送选项"
         Height          =   1455
         Left            =   5160
         TabIndex        =   47
         Top             =   240
         Width           =   4695
         Begin VB.CheckBox chk 
            Caption         =   "门诊医嘱发送后启用诊间支付"
            Height          =   180
            Index           =   12
            Left            =   120
            TabIndex        =   49
            Top             =   720
            Width           =   3015
         End
         Begin VB.CheckBox chk 
            Caption         =   "发送时将本科执行的填为已执行"
            Height          =   195
            Index           =   10
            Left            =   120
            TabIndex        =   48
            Top             =   360
            Width           =   2820
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "门诊发送单据类别 "
         Height          =   1080
         Index           =   0
         Left            =   5160
         TabIndex        =   50
         Top             =   1920
         Width           =   4695
         Begin VB.OptionButton opt发送单据类型 
            Caption         =   "收费单据"
            Height          =   180
            Index           =   0
            Left            =   240
            TabIndex        =   51
            Top             =   330
            Width           =   1020
         End
         Begin VB.OptionButton opt发送单据类型 
            Caption         =   "记帐单据"
            Height          =   180
            Index           =   1
            Left            =   1395
            TabIndex        =   52
            Top             =   330
            Width           =   1020
         End
         Begin VB.OptionButton opt发送单据类型 
            Caption         =   "发送时再确定"
            Height          =   180
            Index           =   2
            Left            =   2565
            TabIndex        =   53
            Top             =   330
            Value           =   -1  'True
            Width           =   1380
         End
         Begin VB.CheckBox chk 
            Caption         =   "只有合约单位病人的医嘱才可以发送为记帐单"
            Height          =   195
            Index           =   6
            Left            =   255
            TabIndex        =   54
            Top             =   630
            Width           =   3960
         End
      End
      Begin VB.Frame fraSendNO 
         Caption         =   "门诊发送单据产生规则"
         Height          =   3435
         Left            =   240
         TabIndex        =   37
         Top             =   3600
         Width           =   4695
         Begin VB.CheckBox chk 
            Caption         =   "检验医嘱发送时一组检验发送为一张单据"
            Height          =   180
            Index           =   147
            Left            =   465
            TabIndex        =   510
            Top             =   2760
            Width           =   4080
         End
         Begin VB.CommandButton cmd门诊发送一张单据类别 
            Caption         =   "全清"
            Height          =   300
            Index           =   1
            Left            =   3600
            TabIndex        =   45
            Top             =   2160
            Width           =   900
         End
         Begin VB.CommandButton cmd门诊发送一张单据类别 
            Caption         =   "全选"
            Height          =   300
            Index           =   0
            Left            =   3600
            TabIndex        =   44
            Top             =   1800
            Width           =   900
         End
         Begin VB.ListBox lst 
            Columns         =   3
            Enabled         =   0   'False
            ForeColor       =   &H80000012&
            Height          =   900
            IMEMode         =   3  'DISABLE
            Index           =   0
            Left            =   480
            Style           =   1  'Checkbox
            TabIndex        =   43
            Top             =   1815
            Width           =   2940
         End
         Begin VB.OptionButton opt发送单据规则 
            Caption         =   "以下同一类别医嘱相同执行科室只产生一张单据"
            Height          =   180
            Index           =   0
            Left            =   240
            TabIndex        =   42
            Top             =   1515
            Width           =   4140
         End
         Begin VB.OptionButton opt发送单据规则 
            Caption         =   "每次发送医嘱只产生一张单据"
            Height          =   180
            Index           =   1
            Left            =   240
            TabIndex        =   40
            Top             =   900
            Width           =   3060
         End
         Begin VB.CheckBox chk 
            Caption         =   "一并给药的即使处方笺不同也发送为一张单据"
            Height          =   255
            Index           =   9
            Left            =   465
            TabIndex        =   46
            Top             =   3045
            Value           =   1  'Checked
            Width           =   3975
         End
         Begin VB.OptionButton opt发送单据规则 
            Caption         =   "所有类别医嘱在相同执行科室只产生一张单据"
            Height          =   180
            Index           =   2
            Left            =   240
            TabIndex        =   41
            Top             =   1200
            Width           =   4140
         End
         Begin VB.CheckBox chk 
            Caption         =   "不同诊断的医嘱分别产生单据"
            Height          =   180
            Index           =   7
            Left            =   240
            TabIndex        =   38
            Top             =   315
            Width           =   2760
         End
         Begin VB.CheckBox chk 
            Caption         =   "开始时间不是同一天的分别产生单据"
            Height          =   180
            Index           =   8
            Left            =   240
            TabIndex        =   39
            Top             =   600
            Width           =   3480
         End
      End
      Begin VB.Frame fra发送 
         Caption         =   "发送为记帐划价单的诊疗类别"
         Height          =   2775
         Left            =   240
         TabIndex        =   30
         Top             =   240
         Width           =   4695
         Begin VB.CheckBox chk 
            Caption         =   "检验医嘱发送时生成样本条码"
            Height          =   200
            Index           =   54
            Left            =   120
            TabIndex        =   36
            Top             =   2160
            Width           =   2640
         End
         Begin VB.CommandButton cmd发送划价类别 
            Caption         =   "全选"
            Height          =   300
            Index           =   0
            Left            =   3600
            TabIndex        =   34
            Top             =   600
            Width           =   900
         End
         Begin VB.CommandButton cmd发送划价类别 
            Caption         =   "全清"
            Height          =   300
            Index           =   1
            Left            =   3600
            TabIndex        =   35
            Top             =   960
            Width           =   900
         End
         Begin TabDlg.SSTab SendPriceType 
            Height          =   1665
            Left            =   120
            TabIndex        =   31
            Top             =   255
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   2937
            _Version        =   393216
            Style           =   1
            Tabs            =   2
            Tab             =   1
            TabsPerRow      =   2
            TabHeight       =   520
            TabMaxWidth     =   882
            TabCaption(0)   =   "门诊"
            TabPicture(0)   =   "frmParClinic.frx":3004E
            Tab(0).ControlEnabled=   0   'False
            Tab(0).Control(0)=   "lst(4)"
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "住院"
            TabPicture(1)   =   "frmParClinic.frx":3006A
            Tab(1).ControlEnabled=   -1  'True
            Tab(1).Control(0)=   "lst(3)"
            Tab(1).Control(0).Enabled=   0   'False
            Tab(1).ControlCount=   1
            Begin VB.ListBox lst 
               Columns         =   3
               Height          =   1110
               IMEMode         =   3  'DISABLE
               Index           =   3
               ItemData        =   "frmParClinic.frx":30086
               Left            =   75
               List            =   "frmParClinic.frx":30088
               Style           =   1  'Checkbox
               TabIndex        =   33
               Top             =   360
               Width           =   3180
            End
            Begin VB.ListBox lst 
               Columns         =   3
               Height          =   1110
               IMEMode         =   3  'DISABLE
               Index           =   4
               ItemData        =   "frmParClinic.frx":3008A
               Left            =   -74925
               List            =   "frmParClinic.frx":3008C
               Style           =   1  'Checkbox
               TabIndex        =   32
               Top             =   360
               Width           =   3180
            End
         End
      End
   End
   Begin VB.PictureBox picPar 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   7335
      Index           =   10
      Left            =   2565
      ScaleHeight     =   7305
      ScaleWidth      =   10065
      TabIndex        =   235
      Top             =   30
      Width           =   10095
      Begin VB.Frame fraEprControl 
         Caption         =   "电子病案"
         Height          =   7080
         Left            =   120
         TabIndex        =   236
         Top             =   150
         Width           =   9855
         Begin VB.CheckBox chk 
            Caption         =   "病案必须先编目后评分"
            Height          =   285
            Index           =   149
            Left            =   7635
            TabIndex        =   512
            Top             =   2093
            Width           =   2100
         End
         Begin VB.CheckBox chk 
            Caption         =   "评分自动写入病案主页"
            Height          =   285
            Index           =   148
            Left            =   5265
            TabIndex        =   511
            Top             =   2093
            Width           =   2130
         End
         Begin VB.TextBox txt 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   180
            Index           =   15
            Left            =   1815
            TabIndex        =   246
            Text            =   "7"
            Top             =   1800
            Width           =   300
         End
         Begin VB.TextBox txt 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   180
            Index           =   16
            Left            =   1815
            TabIndex        =   245
            Text            =   "10"
            Top             =   2160
            Width           =   300
         End
         Begin VB.TextBox txt 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   180
            Index           =   14
            Left            =   540
            TabIndex        =   244
            Text            =   "10"
            Top             =   660
            Width           =   435
         End
         Begin VB.TextBox txt 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   180
            Index           =   13
            Left            =   2325
            TabIndex        =   243
            Text            =   "7"
            Top             =   315
            Width           =   510
         End
         Begin VB.CheckBox chk 
            Caption         =   "必须录入借阅申请理由"
            Height          =   285
            Index           =   105
            Left            =   5265
            TabIndex        =   242
            Top             =   1748
            Width           =   2100
         End
         Begin VB.CheckBox chk 
            Caption         =   "允许自由录入借阅原因"
            Height          =   285
            Index           =   106
            Left            =   7635
            TabIndex        =   241
            Top             =   1748
            Width           =   2130
         End
         Begin VB.CommandButton cmdEprUp 
            Caption         =   "上移(&U)"
            Height          =   350
            Left            =   8565
            TabIndex        =   240
            Top             =   675
            Width           =   1200
         End
         Begin VB.CommandButton cmdEprDown 
            Caption         =   "下移(&D)"
            Height          =   350
            Left            =   8565
            TabIndex        =   239
            Top             =   1140
            Width           =   1200
         End
         Begin VB.CheckBox chk 
            Caption         =   "病案必须审查才能归档"
            Height          =   285
            Index           =   103
            Left            =   165
            TabIndex        =   238
            Top             =   1025
            Width           =   2505
         End
         Begin VB.CheckBox chk 
            Caption         =   "允许自由录入审查意见"
            Height          =   285
            Index           =   104
            Left            =   165
            TabIndex        =   237
            Top             =   1395
            Width           =   2160
         End
         Begin VSFlex8Ctl.VSFlexGrid vsfEpr 
            Height          =   1395
            Index           =   0
            Left            =   5280
            TabIndex        =   247
            Top             =   210
            Width           =   3180
            _cx             =   5609
            _cy             =   2461
            Appearance      =   0
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   16772055
            ForeColorSel    =   0
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   12698049
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   2
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   1
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   255
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   ""
            ScrollTrack     =   0   'False
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   1
            MergeCompare    =   0
            AutoResize      =   -1  'True
            AutoSizeMode    =   1
            AutoSearch      =   0
            AutoSearchDelay =   2
            MultiTotals     =   -1  'True
            SubtotalPosition=   1
            OutlineBar      =   0
            OutlineCol      =   0
            Ellipsis        =   0
            ExplorerBar     =   0
            PicturesOver    =   0   'False
            FillStyle       =   0
            RightToLeft     =   0   'False
            PictureType     =   0
            TabBehavior     =   0
            OwnerDraw       =   0
            Editable        =   0
            ShowComboButton =   1
            WordWrap        =   -1  'True
            TextStyle       =   0
            TextStyleFixed  =   0
            OleDragMode     =   0
            OleDropMode     =   0
            DataMode        =   0
            VirtualData     =   -1  'True
            DataMember      =   ""
            ComboSearch     =   3
            AutoSizeMouse   =   -1  'True
            FrozenRows      =   0
            FrozenCols      =   0
            AllowUserFreezing=   0
            BackColorFrozen =   0
            ForeColorFrozen =   0
            WallPaperAlignment=   9
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   24
         End
         Begin VSFlex8Ctl.VSFlexGrid vsfEpr 
            Height          =   4155
            Index           =   2
            Left            =   5265
            TabIndex        =   248
            Top             =   2820
            Width           =   4500
            _cx             =   7937
            _cy             =   7329
            Appearance      =   0
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   16772055
            ForeColorSel    =   0
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   12698049
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   2
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   2
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   255
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   ""
            ScrollTrack     =   0   'False
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   1
            MergeCompare    =   0
            AutoResize      =   -1  'True
            AutoSizeMode    =   1
            AutoSearch      =   0
            AutoSearchDelay =   2
            MultiTotals     =   -1  'True
            SubtotalPosition=   1
            OutlineBar      =   0
            OutlineCol      =   0
            Ellipsis        =   0
            ExplorerBar     =   0
            PicturesOver    =   0   'False
            FillStyle       =   0
            RightToLeft     =   0   'False
            PictureType     =   0
            TabBehavior     =   0
            OwnerDraw       =   0
            Editable        =   0
            ShowComboButton =   1
            WordWrap        =   -1  'True
            TextStyle       =   0
            TextStyleFixed  =   0
            OleDragMode     =   0
            OleDropMode     =   0
            DataMode        =   0
            VirtualData     =   -1  'True
            DataMember      =   ""
            ComboSearch     =   3
            AutoSizeMouse   =   -1  'True
            FrozenRows      =   0
            FrozenCols      =   0
            AllowUserFreezing=   0
            BackColorFrozen =   0
            ForeColorFrozen =   0
            WallPaperAlignment=   9
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   24
            Begin VB.CommandButton cmdEprSelect 
               Caption         =   "…"
               Height          =   240
               Index           =   2
               Left            =   4140
               TabIndex        =   258
               Top             =   270
               Visible         =   0   'False
               Width           =   300
            End
         End
         Begin VSFlex8Ctl.VSFlexGrid vsfEpr 
            Height          =   4155
            Index           =   1
            Left            =   165
            TabIndex        =   249
            Top             =   2820
            Width           =   4500
            _cx             =   7937
            _cy             =   7329
            Appearance      =   0
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   16772055
            ForeColorSel    =   0
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   12698049
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   2
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   2
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   255
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   ""
            ScrollTrack     =   0   'False
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   1
            MergeCompare    =   0
            AutoResize      =   -1  'True
            AutoSizeMode    =   1
            AutoSearch      =   0
            AutoSearchDelay =   2
            MultiTotals     =   -1  'True
            SubtotalPosition=   1
            OutlineBar      =   0
            OutlineCol      =   0
            Ellipsis        =   0
            ExplorerBar     =   0
            PicturesOver    =   0   'False
            FillStyle       =   0
            RightToLeft     =   0   'False
            PictureType     =   0
            TabBehavior     =   0
            OwnerDraw       =   0
            Editable        =   0
            ShowComboButton =   1
            WordWrap        =   -1  'True
            TextStyle       =   0
            TextStyleFixed  =   0
            OleDragMode     =   0
            OleDropMode     =   0
            DataMode        =   0
            VirtualData     =   -1  'True
            DataMember      =   ""
            ComboSearch     =   3
            AutoSizeMouse   =   -1  'True
            FrozenRows      =   0
            FrozenCols      =   0
            AllowUserFreezing=   0
            BackColorFrozen =   0
            ForeColorFrozen =   0
            WallPaperAlignment=   9
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   24
            Begin VB.CommandButton cmdEprSelect 
               Caption         =   "…"
               Height          =   240
               Index           =   1
               Left            =   4125
               TabIndex        =   257
               Top             =   255
               Visible         =   0   'False
               Width           =   300
            End
         End
         Begin VB.Label lblEpr 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "电子病案评分科室范围"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   10
            Left            =   5265
            TabIndex        =   256
            Tag             =   "0"
            Top             =   2520
            Width           =   1800
         End
         Begin VB.Label lblEpr 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "电子病案审查科室范围 "
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   9
            Left            =   165
            TabIndex        =   255
            Top             =   2520
            Width           =   1890
         End
         Begin VB.Line lnEpr 
            Index           =   4
            X1              =   1755
            X2              =   2185
            Y1              =   1995
            Y2              =   1995
         End
         Begin VB.Line lnEpr 
            Index           =   5
            X1              =   1755
            X2              =   2185
            Y1              =   2340
            Y2              =   2340
         End
         Begin VB.Label lblEpr 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "借阅时的期限缺省为     天"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   11
            Left            =   165
            TabIndex        =   254
            Top             =   1800
            Width           =   2250
         End
         Begin VB.Label lblEpr 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "借阅时的最长期限为     天"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   12
            Left            =   165
            TabIndex        =   253
            Top             =   2145
            Width           =   2250
         End
         Begin VB.Label lblEpr 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "档案排序顺序"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   8
            Left            =   8565
            TabIndex        =   252
            Top             =   210
            Width           =   1080
         End
         Begin VB.Line lnEpr 
            Index           =   2
            X1              =   2340
            X2              =   2875
            Y1              =   510
            Y2              =   510
         End
         Begin VB.Line lnEpr 
            Index           =   3
            X1              =   540
            X2              =   990
            Y1              =   840
            Y2              =   840
         End
         Begin VB.Label lblEpr 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "每隔     分钟自动刷新等待复查问题"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   7
            Left            =   165
            TabIndex        =   251
            Top             =   670
            Width           =   2970
         End
         Begin VB.Label lblEpr 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "处理反馈问题的缺省期限为      天"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   6
            Left            =   165
            TabIndex        =   250
            Top             =   315
            Width           =   2880
         End
      End
   End
   Begin VB.PictureBox picPar 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   5535
      Index           =   11
      Left            =   5640
      ScaleHeight     =   5505
      ScaleWidth      =   6345
      TabIndex        =   259
      TabStop         =   0   'False
      Top             =   600
      Width           =   6375
      Begin VB.CheckBox chk 
         Caption         =   "允许未收费的医嘱执行完成"
         Height          =   255
         Index           =   108
         Left            =   360
         TabIndex        =   270
         Top             =   2040
         Width           =   3015
      End
      Begin VB.CheckBox chk 
         Caption         =   "填写皮试结果时验证身份"
         Height          =   255
         Index           =   107
         Left            =   360
         TabIndex        =   269
         Top             =   1680
         Width           =   3015
      End
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   270
         Index           =   20
         Left            =   1155
         MaxLength       =   2
         TabIndex        =   268
         Top             =   1275
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   270
         Index           =   19
         Left            =   1155
         MaxLength       =   2
         TabIndex        =   266
         Top             =   915
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   270
         Index           =   18
         Left            =   3840
         MaxLength       =   2
         TabIndex        =   264
         Top             =   555
         Width           =   495
      End
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   270
         Index           =   17
         Left            =   1155
         MaxLength       =   3
         TabIndex        =   263
         Top             =   555
         Width           =   495
      End
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   270
         Index           =   9
         Left            =   600
         MaxLength       =   6
         TabIndex        =   261
         Top             =   195
         Width           =   615
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfWaittingMixDept 
         Height          =   2415
         Left            =   360
         TabIndex        =   272
         Top             =   2760
         Width           =   5415
         _cx             =   9551
         _cy             =   4260
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmParClinic.frx":3008E
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   2
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "待配液流程科室："
         Height          =   180
         Index           =   4
         Left            =   360
         TabIndex        =   271
         Top             =   2520
         Width           =   1440
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "皮试提前       分钟提醒"
         Height          =   180
         Index           =   3
         Left            =   360
         TabIndex        =   267
         Top             =   1320
         Visible         =   0   'False
         Width           =   2070
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "输液提前       分钟提醒"
         Height          =   180
         Index           =   2
         Left            =   360
         TabIndex        =   265
         Top             =   960
         Visible         =   0   'False
         Width           =   2070
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "默认滴速      （滴/分钟）， 默认滴系数"
         Height          =   180
         Index           =   1
         Left            =   360
         TabIndex        =   262
         Top             =   600
         Width           =   3420
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "每        秒自动刷新病人清单"
         Height          =   180
         Index           =   0
         Left            =   360
         TabIndex        =   260
         Top             =   240
         Width           =   2520
      End
   End
   Begin VB.PictureBox picFunc 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      FillColor       =   &H8000000A&
      ForeColor       =   &H80000008&
      Height          =   8205
      Left            =   0
      ScaleHeight     =   8205
      ScaleWidth      =   2415
      TabIndex        =   106
      Top             =   0
      Width           =   2415
      Begin VB.PictureBox picTPL 
         BorderStyle     =   0  'None
         Height          =   6135
         Left            =   0
         ScaleHeight     =   6135
         ScaleWidth      =   2250
         TabIndex        =   108
         Top             =   0
         Width           =   2250
         Begin XtremeSuiteControls.TaskPanel tplFunc 
            Height          =   5250
            Left            =   0
            TabIndex        =   109
            Top             =   720
            Width           =   2205
            _Version        =   589884
            _ExtentX        =   3889
            _ExtentY        =   9260
            _StockProps     =   64
            Behaviour       =   1
            ItemLayout      =   2
            HotTrackStyle   =   3
         End
         Begin XtremeCommandBars.ImageManager imgFunc 
            Left            =   1800
            Top             =   360
            _Version        =   589884
            _ExtentX        =   635
            _ExtentY        =   635
            _StockProps     =   0
            Icons           =   "frmParClinic.frx":30114
         End
         Begin XtremeSuiteControls.ShortcutCaption sccFunc 
            Height          =   300
            Left            =   0
            TabIndex        =   110
            Top             =   0
            Width           =   2200
            _Version        =   589884
            _ExtentX        =   3881
            _ExtentY        =   529
            _StockProps     =   6
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SubItemCaption  =   -1  'True
            Alignment       =   1
         End
      End
      Begin VB.PictureBox picVbar 
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         FillColor       =   &H8000000A&
         Height          =   5820
         Left            =   2280
         MousePointer    =   9  'Size W E
         ScaleHeight     =   5820
         ScaleWidth      =   45
         TabIndex        =   107
         Top             =   120
         Width           =   45
      End
      Begin XtremeSuiteControls.ShortcutBar scbFunc 
         Height          =   6765
         Left            =   0
         TabIndex        =   111
         Top             =   0
         Width           =   2400
         _Version        =   589884
         _ExtentX        =   4233
         _ExtentY        =   11933
         _StockProps     =   64
      End
      Begin XtremeCommandBars.ImageManager imgType 
         Left            =   0
         Top             =   0
         _Version        =   589884
         _ExtentX        =   635
         _ExtentY        =   635
         _StockProps     =   0
         Icons           =   "frmParClinic.frx":36010
      End
   End
   Begin VB.PictureBox PicBottom 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   590
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   12630
      TabIndex        =   477
      Top             =   8205
      Width           =   12630
      Begin VB.TextBox txtLocate 
         Height          =   300
         Index           =   1
         Left            =   4700
         TabIndex        =   115
         Top             =   120
         Width           =   1200
      End
      Begin VB.TextBox txtLocate 
         Height          =   300
         Index           =   0
         Left            =   2400
         TabIndex        =   102
         Top             =   120
         Width           =   1200
      End
      Begin VB.CommandButton cmdHelp 
         Caption         =   "帮助(&H)"
         CausesValidation=   0   'False
         Height          =   350
         Left            =   60
         TabIndex        =   99
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   11400
         TabIndex        =   98
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   10245
         TabIndex        =   97
         Top             =   120
         Width           =   1100
      End
      Begin VB.Label lblPrompt 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00C00000&
         Height          =   225
         Left            =   6000
         TabIndex        =   116
         Top             =   165
         Width           =   4095
      End
      Begin VB.Label lblLocate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "科室查找(&F)"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   3600
         TabIndex        =   114
         Top             =   165
         Width           =   1095
      End
      Begin VB.Label lblLocate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "参数查找(&S)"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   1200
         TabIndex        =   101
         Top             =   168
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmParClinic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mrsPar As ADODB.Recordset '参数与控件对应记录集（同一个参数可能对应一组多个控件）
Private marrFunc(2) As String
Private mlngPreFind As Long
Private mblnOk As Boolean
Private mobjPass As Object     '合理用药监测接口
Private mblnUseBlood As Boolean    '是否安装了了血库

Private Enum constTxtLocate
    txt_Par = 0
    txt_Dept = 1
End Enum

Private Enum constChk
    chk_门诊药嘱输天数 = 1
    chk_门诊下达加皮试 = 2
    chk_门诊药嘱输单量 = 3
    chk_门诊登记代办人 = 5
    chk_合约单位发送记帐 = 6
    chk_不同诊断分单据 = 7
    chk_不同天的分单据 = 8
    chk_一并给给一张单据 = 9
    
    chk_门诊发送本科自动执行 = 10
    chk_门诊医嘱发送后启用诊间支付 = 12
    chk_住院临嘱缺省一次性 = 13
    chk_住院药嘱输单量 = 14
    chk_住院药嘱输天数 = 15
    chk_下出院医嘱检查出院诊断 = 16
    chk_手术执行后下术后医嘱 = 17
    chk_待入住病人医嘱下达 = 18
    chk_住院下达加皮试 = 19
    chk_住院下达自动排序 = 20
    
    
    chk_住院无须校对发送 = 21
    chk_特殊医嘱发送检查未生效医嘱 = 22
    chk_药嘱发送限制结束时间 = 23
    chk_长嘱口服药发送结束时间 = 25
    chk_校对确诊停止重整后打印 = 26
    chk_校对确诊停止电子签名 = 27
    chk_批量校对 = 28
    chk_批量暂停 = 29
    chk_登记皮试验证身份 = 30
    chk_允许处理医技医嘱 = 31
    chk_实习医生停嘱需审核 = 33
    chk_超期收回自动审核本科 = 36
    chk_确认停止后自动收回 = 37
    chk_住院发送检查医保审批 = 38
    chk_住院本科自动执行长嘱 = 39
    chk_住院本科自动执行临嘱 = 40
    chk_叮嘱需要发送执行 = 168
    
    chk_过敏登记有效天数 = 11
    chk_门诊处方条数限制 = 52
    
    chk_药品按规格下医嘱 = 4
    chk_长期医嘱次日生效 = 24
    chk_下达医嘱时显示产地 = 66
    chk_一次申请多个检验项目 = 34
    chk_回退出院医嘱才允许撤销出院 = 81
    chk_下达出院医嘱才允许出院 = 50
    
    
    chk_检验医嘱发送生成条形码 = 54
    chk_住院药嘱发送产生领药号 = 64
    chk_未作废临嘱禁止退药 = 0
        
    chk_允许处理超过挂号有效天数的病人 = 82
    chk_临床工作站必须使用zlPlugIn部件 = 79
    chk_医嘱超量时必须输入原因 = 86
    chk_医嘱用药天数反算 = 138
    
    chk_输血医嘱执行后需要核对 = 74
    chk_皮试医嘱执行后需要核对 = 61
    
    chk_医嘱执行有效天数 = 87
    chk_指定医嘱在其他科室执行 = 62
                    
    chk_启用血库管理系统 = 135
    chk_用血医嘱发送后才能发血 = 176
    chk_输血申请不显示血液库存 = 175
    chk_下达用血申请时确定发血信息 = 181
    chk_血液接收后才允许执行登记 = 183
    
    chk_输血分级管理 = 35
    chk_输血申请三级审核 = 53
    chk_输血申请限制中级及以上医师 = 85
        
    
    chk_抗菌药物分级管理 = 75
    chk_抗菌药物使用自备药 = 76
    chk_按医疗小组进行抗菌药物审核 = 137
    
    chk_手术分级管理 = 80
    chk_手术授权管理 = 49
    chk_手术分级审核 = 140
    chk_主刀医师达到手术等级无需审核 = 83
            
    chk_禁忌药嘱 = 65
    chk_禁止下达超极量药品医嘱 = 63
    chk_允许院外执行禁忌药品 = 77
    chk_接口调用日志_大通 = 88   '大通接口日志调用 65522
    chk_使用系统设置_美康 = 89   '美康接口系统设置功能控制参数 65198
    chk_禁忌药品要求填写原因 = 139
    '临床路径应用
    chk_启用路径执行环节 = 41
    chk_路径执行环节医生启用 = 42
    chk_路径执行环节护士启用 = 43
    chk_路径病人医技医嘱显示 = 44
    chk_未评估时允许添加医嘱到昨天 = 45
    chk_允许前一天不评估就生成今天路径项目 = 46
    chk_允许提前生成明天的路径项目 = 47
    chk_匹配时期效不同算路径外项目 = 84
    chk_药品医嘱不匹配为路径外项目 = 57
    chk_出院后不允许取消完成路径 = 153
    chk_启用药剂科和医务科双审核模式 = 58
    chk_药品医嘱相同分类不算路径外医嘱 = 134
    
    chk_门诊启用路径执行环节 = 169

    chk_门诊医嘱下达时诊疗选择器显示药品库存 = 48
    chk_住院医嘱下达时诊疗选择器显示药品库存 = 51
    chk_执行单打印时逐个病人换页打印 = 55
    chk_长嘱单转科换页 = 59
    chk_转科换页后在首行打印重开医嘱 = 60
    chk_收回时医嘱列表只显示当前病区的医嘱 = 67
    chk_使用手术结束时间 = 69
    chk_中医科室不使用项目 = 70
    chk_医生和护士分别填写病案首页 = 71
    chk_医生呼叫人数限制含回诊病人 = 72
    chk_医生主动呼叫后才允许在队列中接诊 = 73
    chk_只接收已经分诊的病人 = 78
    chk_允许未收费病人完成执行 = 91
    chk_填写皮试结果时验证身份 = 92
    chk_执行报到时收费或记账审核 = 93
    chk_血透室书写新版护理记录 = 150
    
    chk_签名移位 = 94
    chk_签名级别前辍 = 95
    chk_显示手签位置 = 96
    chk_签名使用图片 = 97
    chk_签名使用原图 = 98
    chk_病历诊断同步首页 = 99
    chk_转科后书写的病历另起一页打印 = 100
    chk_住院病历自动显示新增面板 = 101
    chk_要求先书写被共享病历 = 102
    chk_病案必须审查才能归档 = 103
    chk_允许自由录入审查意见 = 104
    chk_必须录入借阅申请理由 = 105
    chk_允许自由录入借阅原因 = 106
    chk_评分自动写入病案主页 = 148
    chk_病案必须先编目后评分 = 149
    
    chk_门诊输液皮试验证身份 = 107
    chk_门诊输液未收费允许完成 = 108
    
    '新版住院护士工作站
    chk_卡片余额含担保金额 = 133
    
    chk_不显示无床位的病区科室 = 68
    chk_门诊诊断不作为导入临床路径的诊断依据 = 179
    
    '产程图
    chk_产程图显示产程时间 = 109
    chk_产程图模式 = 110
    chk_先露高低显示位置 = 111
    chk_产程图上显示警戒线 = 112
    
    '记录单
    chk_日期显示方式 = 113
    chk_只在当前页显示跨页数据 = 114
    chk_对应多份护理文件 = 115
    chk_文件页码规则 = 116
    chk_签名人显示方式 = 117
    chk_允许数据同步 = 167
    
    '体温单
    chk_自动标志顶格输出 = 118
    chk_自动标志超出40刻度缩小显示 = 119
    chk_自动标志按顺序当天排列 = 120
    chk_首列日期格式 = 121
    chk_入科标志不自动转为入院 = 122
    chk_不足术后天数出院 = 123
    chk_全天汇总显示小时 = 124
    chk_汇总项目显示当天数据 = 125
    chk_不输出心率列 = 126
    chk_绘图刻度单格显示 = 127
    chk_显示诊断信息 = 128
    chk_再次手术停止前次标注 = 129
    chk_打印医院名称 = 130
    chk_婴儿住院天数从0开始计算 = 131
    chk_输出曲线说明信息 = 132
    chk_曲线项目骑线显示 = 180
    chk_小结缺省标识 = 188
    
    
    '住院首页
    chk_身份证密文 = 136
    
    chk_住院检验医嘱发送时一组检验发送为一张单据 = 141
    chk_本人执行登记 = 32
    
    chk_申请单启用环节会诊 = 90
    chk_申请单启用环节门诊手术 = 142
    chk_申请单启用环节门诊输血 = 143
    chk_申请单启用环节门诊检验 = 144
    chk_申请单启用环节门诊检查 = 145
    
    chk_申请单启用环节住院手术 = 157
    chk_申请单启用环节住院输血 = 156
    chk_申请单启用环节住院检验 = 152
    chk_申请单启用环节住院检查 = 151
    
    chk_多科会诊意见书写要求 = 146
    chk_门诊检验医嘱发送时一组检验发送为一张单据 = 147
    chk_诊断手术名称自由调整 = 154
    chk_诊断录入时附码自动提取 = 170
    chk_路径变异原因从字典表中选取 = 155
    chk_转入转出医疗机构不允许自由录入 = 134
    chk_启用申请单后必须使用申请单下达医嘱门诊 = 158
    chk_启用申请单后必须使用申请单下达医嘱住院 = 159
    chk_特殊药品分开发送 = 160
    chk_发送完成后关闭医嘱窗体 = 161
    chk_停嘱时录入原因 = 56
    chk_临嘱单转科换页 = 162
    chk_科室药房对照按本机参数设置 = 163
    chk_病理诊断只允许录入M打头的肿瘤形态学编码 = 164
    
    chk_门诊医嘱发送皮试限制 = 165
    chk_住院医嘱发送皮试限制 = 166
    
    chk_长嘱单术后换页 = 171
    chk_长嘱单重整换页 = 173
    chk_住院登记代办人 = 174
    chk_皮试阳性用药门诊 = 177
    chk_皮试阳性用药住院 = 178
    
    chk_临嘱单术后换页 = 172
    
    chk_传染病报告卡强制填写 = 182
    chk_会诊科室下达医嘱由会诊申请科室处理 = 184
    chk_存在未发送医嘱时禁止处理转科医嘱 = 185
    chk_门诊西医科允许录入中医诊断 = 186
    chk_不填写病理号 = 187
End Enum

Private Enum constCbo
    cbo_住院首页标准 = 0
    cbo_住院医生站列表显示 = 1
    cbo_病历签名显示时间 = 2
    
    '记录单部分
    cbo_审签模式 = 3
    cbo_签名列显示模式 = 13
    cbo_小结缺省标识 = 14
    '产程图部分
    cbo_警戒线显示 = 4
    cbo_异常线显示 = 5
    cbo_零点与首次点连接 = 6
    cbo_宫口扩大异常产 = 7
    cbo_先露下降异常产 = 8
    cbo_生产标志内容 = 9
    cbo_生产标志位置 = 10
    cbo_宫口扩大顺产 = 11
    cbo_先露下降顺产 = 12
    '体温单
    cbo_未记说明显示位置 = 15
    cbo_体温不升显示方式 = 16
    cbo_呼吸机符号显示位置 = 17
    cbo_呼吸表格数据显示位置 = 18
    cbo_手术当天缺省格式 = 19
    cbo_标志说名与时间连接符号 = 21
    cbo_入院自动标志 = 22
    cbo_入科自动标志 = 23
    cbo_转科自动标志 = 24
    cbo_换床自动标志 = 25
    cbo_手术自动标志 = 30
    cbo_出院自动标志 = 31
    cbo_分娩自动标志 = 32
    cbo_出生自动标志 = 33
    cbo_回室自动标志 = 34
    cbo_转病区自动标志 = 35
    
    cbo_合理用药接口 = 20
    cbo_中药配方 = 26
    cmd_过敏输入来源 = 27
    cbo_门诊药房科室对照方案 = 28
    cbo_住院药房科室对照方案 = 29
    cbo_输血采集默认诊疗类型 = 36
    cbo_住院本科执行自动完成方案 = 37
End Enum

Private Enum constUpDown
    ud_医嘱执行有效天数 = 0
    ud_门诊处方条数限制 = 3
    ud_过敏登记有效天数 = 7
    
    ud_补录医嘱识别间隔 = 8
    ud_门诊新开医嘱间隔 = 11
    ud_记录单超期录入天数 = 21
    ud_体温曲线固定添加行数 = 22
    ud_体温夜班开始时点 = 23
    ud_体温夜班结束时点 = 24
    ud_体温表格固定添加行数 = 26
    ud_体温开始时间 = 27
    
End Enum

Private Enum constTxtUd
    txtud_体温夜班开始时点 = 23
    txtud_体温夜班结束时点 = 24
End Enum

Private Enum constTxt
    txt_门诊输血申请注意事项 = 0
    txt_住院输血申请注意事项 = 1
    txt_中药配方允许修改的中药味数上限 = 2
    txt_生成时医嘱允许超过当前时间天数 = 3
    txt_允许提前接诊分钟 = 4
    txt_门诊医生站病人列表自动刷新秒数 = 5
    txt_住院医生病案反馈数 = 6
    txt_住院护士病案反馈数 = 7

    
    txt_签名使用图片高度 = 10
    txt_共享病历N行自动折叠 = 11
    txt_共享病历连续预览N天 = 12
    txt_电子病案审查缺省期限 = 13
    txt_电子病案审查待复查刷新间隔 = 14
    txt_电子病案借阅缺省期限 = 15
    txt_电子病案借阅最大期限 = 16
    
    txt_门诊输液自动刷新病人 = 9
    txt_门诊输液滴速 = 17
    txt_门诊输液滴系数 = 18
    txt_门诊输液输液提醒 = 19
    txt_门诊输液皮试提醒 = 20
    txt_整体护理IP地址 = 21
    txt_整体护理IP端口 = 22
    txt_手术标注天数 = 25
    txt_体温复试合格符号 = 28
    txt_门诊中药配方允许修改的中药味数上限 = 24
    txt_门诊生成时医嘱允许超过当前时间天数 = 23
    txt_长嘱单中药医嘱单行显示字数 = 8
    txt_临嘱单中药医嘱单行显示字数 = 26
End Enum

Private Enum constListBox
    lst_门诊发送一张单据类别 = 0
    lst_门诊发送检查诊断 = 1
    lst_住院检查入院诊断 = 2
    lst_住院发送划价类别 = 3
    lst_门诊发送划价类别 = 4    '发送为划价单的诊疗类别
    lst_不收回的发药类型 = 6
    lst_药品留存给药途径 = 7
    lst_本科执行自动完成医嘱类别 = 8
End Enum

'合理用药接口
Private Enum lblEnum
    lbl_过敏输入来源 = 0
End Enum

Private mrs门诊药房对照 As ADODB.Recordset
Private mrs住院药房对照 As ADODB.Recordset
Private mrs住院执行对照 As ADODB.Recordset

Private mcol科室 As Collection '不用填写超量说明的科室
Private mcolStop科室 As Collection  '不用停嘱录入原因的科室
Private mrsAdvice As New ADODB.Recordset '记录医嘱内容定义

Private Sub cmdAdd_Click(Index As Integer)
    Dim lngIndex As Long, i As Long
    
    If cbo(Index).ListCount > 0 Then
        lngIndex = cbo(Index).ItemData(cbo(Index).ListCount - 1)
    Else
        lngIndex = 0
    End If
    cbo(Index).AddItem "方案" & lngIndex + 1
    cbo(Index).ItemData(cbo(Index).NewIndex) = lngIndex + 1
    
    If Index = cbo_门诊药房科室对照方案 Then
        mrs门诊药房对照.AddNew
        mrs门诊药房对照!方案 = lngIndex + 1
    ElseIf Index = cbo_住院药房科室对照方案 Then
        mrs住院药房对照.AddNew
        mrs住院药房对照!方案 = lngIndex + 1
    ElseIf Index = cbo_住院本科执行自动完成方案 Then
        mrs住院执行对照.AddNew
        mrs住院执行对照!方案 = lngIndex + 1
    End If
    cbo(Index).ListIndex = cbo(Index).ListCount - 1
        
    If Index = cbo_住院本科执行自动完成方案 Then
        For i = 0 To lst(lst_本科执行自动完成医嘱类别).ListCount - 1
            lst(lst_本科执行自动完成医嘱类别).Selected(i) = False
        Next
        Frame14.Tag = "已修改"
    Else
        With vsfDrugStore(Index)
            For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("可用")) = 0
                .TextMatrix(i, .ColIndex("缺省")) = ""
                If Index = cbo_门诊药房科室对照方案 Then
                    .TextMatrix(i, .ColIndex("发药窗口")) = "自动分配"
                End If
            Next
        End With
        SST.Tag = "已修改"
    End If
End Sub

Private Sub cmdAddMed_Click()
    frmInMedSetup.ShowMe "", "", "", "新增", Me
    cmdModify.Enabled = vsfMecItem.Rows > 1:  cmdDelete.Enabled = vsfMecItem.Rows > 1
End Sub

Private Sub cmdDel_Click(Index As Integer)
    On Error Resume Next
    If cbo(Index).ListCount = 0 Then Exit Sub
    If MsgBox("是否删除此方案？", vbQuestion + vbYesNo + vbDefaultButton1, Me.Caption) = vbYes Then
        If Index = cbo_门诊药房科室对照方案 Then
            mrs门诊药房对照.Delete
        ElseIf Index = cbo_住院药房科室对照方案 Then
            mrs住院药房对照.Delete
        ElseIf Index = cbo_住院本科执行自动完成方案 Then
            mrs住院执行对照.Delete
        End If
        cbo(Index).RemoveItem cbo(Index).ListIndex
        If cbo(Index).ListIndex = -1 Then
            If cbo(Index).ListCount > 0 Then
                cbo(Index).ListIndex = 0
            Else
                vsUseDept(Index).Rows = 0: vsUseDept(Index).Rows = 1
                vsUseDept(Index).Enabled = False
                If Index = cbo_住院本科执行自动完成方案 Then
                    lst(lst_本科执行自动完成医嘱类别).Enabled = False
                Else
                    vsfDrugStore(Index).Enabled = False
                End If
            End If
        Else
            Call cbo_Click(Index)
        End If
        
        If Index = cbo_住院本科执行自动完成方案 Then
            Frame14.Tag = "已修改"
        Else
            SST.Tag = "已修改"
        End If
    End If
End Sub

Private Sub cmdDelete_Click()
    Dim strSQL As String
    
    If vsfMecItem.Row > 0 Then
        If CheckMecItem = False Then vsfMecItem.SetFocus: Exit Sub
        If MsgBox("确认要删除[" & vsfMecItem.TextMatrix(vsfMecItem.Row, 1) & "]吗?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        strSQL = "zl_病案项目_edit(null,null,null,'" & vsfMecItem.TextMatrix(vsfMecItem.Row, 0) & "',2)"
        On Error GoTo ErrHandle
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        vsfMecItem.RemoveItem vsfMecItem.Row
        If vsfMecItem.Rows = 1 Then
            cmdModify.Enabled = False
            cmdDelete.Enabled = False
        End If
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdDepartSelect_Click()
    Dim lngCol As Long, lngRow As Long, X As Long, Y As Long, rsTmp As ADODB.Recordset, blnCancel As Boolean
    Dim i As Long, blnChange As Boolean, strNewValue As String
    
    lngCol = vsfDepartSign.MouseCol: lngRow = vsfDepartSign.MouseRow
    If lngRow < 1 Then Exit Sub
    vsfDepartSign.Row = vsfDepartSign.Rows - 1 '选中修改行
    gstrSQL = "Select b.Id, b.编码, b.名称,b.简码　" & vbNewLine & _
              "from (Select 参数id, 部门id, 参数值 From Zldeptparas Where 参数id = (Select ID From zlParameters Where 参数名 = '签名使用图片')) A,部门表 B" & vbNewLine & _
              "Where b.Id = a.部门id(+) And a.部门id Is Null" & vbNewLine & _
              "And (b.撤档时间 Is Null Or b.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD'))" & vbNewLine & _
              "Order By b.编码"

    Set rsTmp = FS.ShowSQLSelectEx(Me, cmdDepartSelect, gstrSQL, 0, "", False, "部门", "请选择部门", False, False, True, blnCancel, False, False, False, "")
    
    If blnCancel = False Then
        If rsTmp Is Nothing Then Exit Sub
        If rsTmp.EOF Then Exit Sub
        For i = 1 To vsfDepartSign.Rows - 1
            If vsfDepartSign.TextMatrix(i, vsfDepartSign.ColIndex("ID")) = rsTmp!ID Then Exit Sub
        Next
        vsfDepartSign.TextMatrix(lngRow, vsfDepartSign.ColIndex("ID")) = rsTmp!ID & ""
        vsfDepartSign.TextMatrix(lngRow, vsfDepartSign.ColIndex("编码")) = rsTmp!编码 & ""
        vsfDepartSign.TextMatrix(lngRow, vsfDepartSign.ColIndex("科室")) = rsTmp!名称 & ""
        vsfDepartSign.TextMatrix(lngRow, vsfDepartSign.ColIndex("启用")) = "1"
        blnChange = True
    End If
        
    If blnChange Then
        If vsfDepartSign.TextMatrix(vsfDepartSign.Rows - 1, vsfDepartSign.ColIndex("ID")) <> "" Then
            vsfDepartSign.Rows = vsfDepartSign.Rows + 1
        End If

        With vsfDepartSign
            For i = 1 To .Rows - 1
                If Val(.Cell(flexcpChecked, i, .ColIndex("启用"))) <> Decode(Val(.RowData(i)), 1, 1, 2) Then
                    vsfDepartSign.ForeColor = vbRed
                    Exit Sub
                End If
            Next
            vsfDepartSign.ForeColor = vbBlack
        End With
        cmdDepartSelect.Visible = False
        vsfDepartSign.Row = vsfDepartSign.Rows - 1
        Call vsfDepartSign.ShowCell(vsfDepartSign.Rows - 1, 0)
    End If
End Sub

Private Sub cmdEprDown_Click()
    If vsfEpr(0).Row < vsfEpr(0).Rows - 1 Then
        Dim intCount As Integer, strTmp As String
        '变更行存储值
        strTmp = CStr(vsfEpr(0).RowData(vsfEpr(0).Row))
        vsfEpr(0).RowData(vsfEpr(0).Row) = vsfEpr(0).RowData(vsfEpr(0).Row + 1)
        vsfEpr(0).RowData(vsfEpr(0).Row + 1) = Val(strTmp)
        '变更行显示值
        strTmp = vsfEpr(0).TextMatrix(vsfEpr(0).Row, 0)
        vsfEpr(0).TextMatrix(vsfEpr(0).Row, 0) = vsfEpr(0).TextMatrix(vsfEpr(0).Row + 1, 0)
        vsfEpr(0).TextMatrix(vsfEpr(0).Row + 1, 0) = strTmp
        '变更当前选中行
        vsfEpr(0).Row = vsfEpr(0).Row + 1
        '变更值
        strTmp = ""
        With vsfEpr(0)
            For intCount = 1 To .Rows - 1
                If Val(.RowData(intCount)) > 0 Then
                    strTmp = strTmp & ";" & Val(.RowData(intCount))
                End If
            Next
        End With
        If strTmp <> "" Then strTmp = Mid(strTmp, 2)
        Call SetParChange(vsfEpr, 0, mrsPar, True, strTmp)
    End If
End Sub

Private Sub cmdEprSelect_Click(Index As Integer)
Dim lngCol As Long, lngRow As Long, X As Long, Y As Long, rsTmp As New ADODB.Recordset, blnCancel As Boolean
Dim vPoint As POINTAPI, l As Long, blnChange As Boolean, strNewValue As String
    vPoint = zlControl.GetCoordPos(vsfEpr(Index).hwnd, cmdEprSelect(Index).Left, cmdEprSelect(Index).Top)
    lngCol = vsfEpr(Index).MouseCol: lngRow = vsfEpr(Index).MouseRow
    If lngRow < 1 Then Exit Sub
    If lngCol = 1 Then '姓名
        gstrSQL = "Select Distinct a.编号,a.Id, a.姓名, a.简码, c.名称 As 科室" & vbNewLine & _
                "From 人员表 A, 人员性质说明 B, 部门表 C, 部门人员 D" & vbNewLine & _
                "Where a.Id = b.人员id And c.Id = d.部门id And d.人员id = a.Id And d.缺省 = 1 And" & vbNewLine & _
                "      (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null) And b.人员性质 In ('医生')" & vbNewLine & _
                "Order By a.编号"
        Set rsTmp = zlDatabase.ShowSelect(Me, gstrSQL, 0, "", False, "姓名", "请选择一个审查人员", False, False, True, vPoint.X, vPoint.Y, , blnCancel)
        If blnCancel = False Then
            If rsTmp.EOF Then Exit Sub
            blnChange = True
            vsfEpr(Index).TextMatrix(lngRow, 0) = rsTmp!ID
            vsfEpr(Index).TextMatrix(lngRow, 1) = rsTmp!姓名
        End If
    ElseIf lngCol = 3 Then '科室
        gstrSQL = "Select a.Id,a.编码, a.名称, a.简码" & vbNewLine & _
                    "From 部门表 A, 部门性质说明 B" & vbNewLine & _
                    "Where a.Id = b.部门id And b.工作性质 In ('临床') And b.服务对象 In (2, 3) And" & vbNewLine & _
                    "      (To_Char(a.撤档时间, 'yyyy-MM-dd') = '3000-01-01' Or a.撤档时间 Is Null)" & vbNewLine & _
                    "Order By a.编码"
        Set rsTmp = zlDatabase.ShowSQLMultiSelect(Me, gstrSQL, 0, "", False, "", "请选择一个或多个病人科室", False, False, True, vPoint.X, vPoint.Y, 225, blnCancel, True, True)
        If blnCancel = False Then
            If rsTmp.EOF Then Exit Sub
            blnChange = True
            Do Until rsTmp.EOF
                If rsTmp.AbsolutePosition = 1 Then
                    vsfEpr(Index).TextMatrix(lngRow, 2) = rsTmp!ID
                    vsfEpr(Index).TextMatrix(lngRow, 3) = rsTmp!名称
                Else
                    vsfEpr(Index).TextMatrix(lngRow, 2) = vsfEpr(Index).TextMatrix(lngRow, 2) & "," & rsTmp!ID
                    vsfEpr(Index).TextMatrix(lngRow, 3) = vsfEpr(Index).TextMatrix(lngRow, 3) & vbCrLf & rsTmp!名称
                End If
                rsTmp.MoveNext
            Loop
            vsfEpr(Index).AutoSize 3, 3
        End If
    End If
    
    If blnChange Then
        If vsfEpr(Index).TextMatrix(vsfEpr(Index).Rows - 1, 0) <> "" And vsfEpr(Index).TextMatrix(vsfEpr(Index).Rows - 1, 2) <> "" Then
            vsfEpr(Index).Rows = vsfEpr(Index).Rows + 1
        End If
        
        For l = 1 To vsfEpr(Index).Rows - 1
            If vsfEpr(Index).TextMatrix(l, 0) <> "" And vsfEpr(Index).TextMatrix(l, 2) <> "" Then
                strNewValue = strNewValue & ";" & vsfEpr(Index).TextMatrix(l, 0) & "," & vsfEpr(Index).TextMatrix(l, 2)
             End If
        Next
        If strNewValue <> "" Then
            strNewValue = Mid(strNewValue, 2)
        End If
        
        Call SetParChange(vsfEpr, Index, mrsPar, True, strNewValue)
    End If
End Sub


Private Sub cmdEprUp_Click()
    If vsfEpr(0).Row > 1 Then
        Dim intCount As Integer, strTmp As String
        '变更行存储值
        strTmp = CStr(vsfEpr(0).RowData(vsfEpr(0).Row))
        vsfEpr(0).RowData(vsfEpr(0).Row) = vsfEpr(0).RowData(vsfEpr(0).Row - 1)
        vsfEpr(0).RowData(vsfEpr(0).Row - 1) = Val(strTmp)
        '变更行显示值
        strTmp = vsfEpr(0).TextMatrix(vsfEpr(0).Row, 0)
        vsfEpr(0).TextMatrix(vsfEpr(0).Row, 0) = vsfEpr(0).TextMatrix(vsfEpr(0).Row - 1, 0)
        vsfEpr(0).TextMatrix(vsfEpr(0).Row - 1, 0) = strTmp
        '变更当前选中行
        vsfEpr(0).Row = vsfEpr(0).Row - 1
        '变更值
        strTmp = ""
        With vsfEpr(0)
            For intCount = 1 To .Rows - 1
                If Val(.RowData(intCount)) > 0 Then
                    strTmp = strTmp & ";" & Val(.RowData(intCount))
                End If
            Next
        End With
        If strTmp <> "" Then strTmp = Mid(strTmp, 2)
        Call SetParChange(vsfEpr, 0, mrsPar, True, strTmp)
    End If
End Sub

Private Sub cmdHelp_Click()
     ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdLink_Click()
    Call CheckNurseIntegrateIP
End Sub

Private Function CheckNurseIntegrateIP(Optional ByVal blnComdValiade As Boolean = True) As Boolean
    Dim strIP As String, strErrMsg As String
    If txt(txt_整体护理IP地址).Text = "" Then
        If blnComdValiade = True Then
            MsgBox "请输入整体护理IP地址!", vbInformation, gstrSysName
            Call zlControl.ControlSetFocus(txt(txt_整体护理IP地址))
            Exit Function
        End If
    End If
    '判断IP地址是否正确
    If IsIPAddress(txt(txt_整体护理IP地址).Text) = False Then
        If blnComdValiade = True Then
            MsgBox "IP地址格式不正确，请重新输入！", vbInformation, gstrSysName
            Call zlControl.ControlSetFocus(txt(txt_整体护理IP地址))
        End If
        Exit Function
    End If
    '检查是否是正确移动服务器地址：连接服务器进行测试
    strIP = txt(txt_整体护理IP地址).Text & IIF(txt(txt_整体护理IP端口).Text = "", "", ":" & txt(txt_整体护理IP端口).Text)
    If InitNurseIntegrate(True) = True Then
        If gobjNurseIntegrate.IPAdreesCheck(strIP, strErrMsg) = False Then
            If blnComdValiade = True Then
                MsgBox strErrMsg, vbInformation, gstrSysName
            End If
            Exit Function
        Else
            If blnComdValiade = True Then MsgBox "IP地址设置成功！", vbInformation, gstrSysName
            cmdLink.Tag = "OK"
        End If
        CheckNurseIntegrateIP = True
    End If
End Function

Private Sub cmdModify_Click()
    If vsfMecItem.Row > 0 Then
        If CheckMecItem = False Then vsfMecItem.SetFocus: Exit Sub
        frmInMedSetup.ShowMe vsfMecItem.TextMatrix(vsfMecItem.Row, 0), vsfMecItem.TextMatrix(vsfMecItem.Row, 1), vsfMecItem.TextMatrix(vsfMecItem.Row, 2), "修改", Me
        vsfMecItem.SetFocus
    End If
End Sub

Private Function CheckMecItem() As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "select count(信息名) as 数量 from 病案主页从表 where 信息名='" & vsfMecItem.TextMatrix(vsfMecItem.Row, 1) & "'"
    
    Err = 0: On Error GoTo ErrHandle
    Call zlDatabase.OpenRecordset(rsTemp, strSQL, Me.Caption)
    If rsTemp!数量 > 0 Then
        MsgBox "该项目已经使用,不能进行修改或删除!"
        CheckMecItem = False
        Exit Function
    End If
    CheckMecItem = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub cmdPathSortSet_Click()
    frmParAdviceSort.mbytFun = 0
    frmParAdviceSort.Show vbModal, Me
End Sub


Private Sub cmdPatiSign_Click()
    frmPatiStamp.ShowMe Me
End Sub

Private Sub cmdSet_Click()
    Call mobjPass.Setup(Me, gcnOracle, glngSys)
End Sub

Private Sub cmdUnVisible_Click(Index As Integer)
    Dim lngIndex As Long
    
    picControl(Index).Visible = False
    If picControl(Index).Tag <> "" Then
        lngIndex = Val(picControl(Index).Tag)
        If picLineColor(lngIndex).Enabled And picLineColor(lngIndex).Visible Then picLineColor(lngIndex).SetFocus
    End If
    picControl(Index).Tag = ""
End Sub

Private Sub cmd药品留存给药途径_Click(Index As Integer)
    Call SetLstSelected(lst(lst_药品留存给药途径), Index = 0)
End Sub

Private Sub cmd本科执行自动完成医嘱类别_Click(Index As Integer)
    If lst(lst_本科执行自动完成医嘱类别).Enabled = False Then Exit Sub
    Call SetLstSelected(lst(lst_本科执行自动完成医嘱类别), Index = 0)
End Sub


Private Sub Form_Activate()
    If Me.Tag = "初始成功" Then
        Call scbFunc_SelectedChanged(scbFunc.Selected)
        Me.Tag = ""
    End If
End Sub

Private Sub Form_Load()
    Dim strCategory As String
    
    mblnOk = False
    strCategory = "参数设置,基础项目"
    
    '图标编号,TaskPanelItem的ID(同时也是参数容器Picture控件数组号),TaskPanelItem的标题;......
    marrFunc(0) = "100,0,医嘱下达选项;101,1,业务流程控制;102,2,门诊医嘱发送;103,3,住院医嘱操作;104,4,医嘱其他处理;" & _
                  "106,6,临床路径控制;107,8,临床工作站;108,7,住院首页;105,9,病历书写显示;107,10,电子病案管理;" & _
                  "103,11,门诊输液管理;112,12,新版护理文件"
    marrFunc(1) = "105,5,科室药房设置"

    '1.初始化快捷面板的一级分类列表,缺省选中第一个
    Call InitSCBItem(scbFunc, strCategory, picTPL.hwnd)
    Call scbFunc.Icons.AddIcons(imgType.Icons)
      
    '2.初始化任务面板的二级分类列表,缺省选中第一个
    Call InitTPLItem(sccFunc, tplFunc, scbFunc.Selected.Caption, marrFunc(0))
    Call tplFunc.Icons.AddIcons(imgFunc.Icons)
    
    
    Call InitData
    Call ShowErrParasMsg(Me, mrsPar)
    Me.Tag = "初始成功"
End Sub


Private Sub OptEnemaStool_Click(Index As Integer)
    If Me.Visible Then Call SetParChange(OptEnemaStool, Index, mrsPar)
End Sub

Private Sub OptEnemaStool_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub OptEnemaStool_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(OptEnemaStool, Index, mrsPar)
End Sub

Private Sub optEprRead_Click(Index As Integer)
    If Index = 2 Then
        txt(txt_共享病历连续预览N天).Enabled = True
        Call SetParChange(txt, txt_共享病历连续预览N天, mrsPar, True, Val(txt(txt_共享病历连续预览N天)))
        Call SetParChange(optEprRead, Index, mrsPar, True, Val(txt(txt_共享病历连续预览N天)))
    Else
        txt(txt_共享病历连续预览N天).Enabled = False
        Call SetParChange(optEprRead, Index, mrsPar, True, IIF(Index = 0, "-1", "0"))
    End If
End Sub

Private Sub optEprRead_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optEprRead, Index, mrsPar)
End Sub

Private Sub optFileTime_Click(Index As Integer)
    If Me.Visible Then Call SetParChange(optFileTime, Index, mrsPar)
End Sub

Private Sub optFileTime_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optFileTime_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optFileTime, Index, mrsPar)
End Sub

Private Sub optICD附码_Click(Index As Integer)
    If Me.Visible Then Call SetParChange(optICD附码, Index, mrsPar)
End Sub

Private Sub optICD附码_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optICD附码_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optICD附码, Index, mrsPar)
End Sub

Private Sub OptInsert_Click(Index As Integer)
    If Me.Visible Then Call SetParChange(OptInsert, Index, mrsPar)
End Sub

Private Sub OptInsert_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub OptInsert_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(OptInsert, Index, mrsPar)
End Sub

Private Sub optNewCard_Click(Index As Integer)
    If Me.Visible Then Call SetParChange(optNewCard, Index, mrsPar)
End Sub

Private Sub optNewCard_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optNewCard_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optNewCard, Index, mrsPar)
End Sub

Private Sub OptOut_Click(Index As Integer)
    If Me.Visible Then Call SetParChange(OptOut, Index, mrsPar)
End Sub

Private Sub OptOut_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub OptOut_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(OptOut, Index, mrsPar)
End Sub

Private Sub optPloy_Click(Index As Integer)
    If Me.Visible Then Call SetParChange(optPloy, Index, mrsPar)
End Sub

Private Sub optPloy_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optPloy_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optPloy, Index, mrsPar)
End Sub

Private Sub optSign_Click(Index As Integer)
    Call SetParChange(optSign, Index, mrsPar)
End Sub

Private Sub optSign_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optSign, Index, mrsPar)
End Sub

Private Sub opt病人过滤_Click(Index As Integer)
    If Me.Visible Then Call SetParChange(opt病人过滤, Index, mrsPar)
End Sub

Private Sub opt病人过滤_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(opt病人过滤, Index, mrsPar)
End Sub

Private Sub opt路径打印规则_Click(Index As Integer)
    If Me.Visible Then Call SetParChange(opt路径打印规则, Index, mrsPar)
End Sub

Private Sub opt路径打印规则_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub opt路径打印规则_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(opt路径打印规则, Index, mrsPar)
End Sub

Private Sub optPrintWay_Click(Index As Integer)
    If Me.Visible Then Call SetParChange(optPrintWay, Index, mrsPar)
End Sub

Private Sub optPrintWay_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optPrintWay_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optPrintWay, Index, mrsPar)
End Sub

Private Sub opt每页路径打印天数_Click(Index As Integer)
    If Me.Visible Then Call SetParChange(opt每页路径打印天数, Index, mrsPar)
End Sub

Private Sub opt每页路径打印天数_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub opt每页路径打印天数_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(opt每页路径打印天数, Index, mrsPar)
End Sub

Private Sub opt门诊路径打印规则_Click(Index As Integer)
    If Me.Visible Then Call SetParChange(opt门诊路径打印规则, Index, mrsPar)
End Sub

Private Sub opt门诊路径打印规则_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub opt门诊路径打印规则_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(opt门诊路径打印规则, Index, mrsPar)
End Sub

Private Sub opt门诊路径打印天数_Click(Index As Integer)
    If Me.Visible Then Call SetParChange(opt门诊路径打印天数, Index, mrsPar)
End Sub

Private Sub opt门诊路径打印天数_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub opt门诊路径打印天数_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(opt门诊路径打印天数, Index, mrsPar)
End Sub

Private Sub opt区域_Click(Index As Integer)
    If Me.Visible Then Call SetParChange(opt区域, Index, mrsPar)
End Sub

Private Sub opt区域_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub opt区域_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(opt区域, Index, mrsPar)
End Sub

Private Sub opt损伤中毒_Click(Index As Integer)
    If Me.Visible Then Call SetParChange(opt损伤中毒, Index, mrsPar)
End Sub

Private Sub opt损伤中毒_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub opt损伤中毒_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(opt损伤中毒, Index, mrsPar)
End Sub


Private Sub opt病理诊断_Click(Index As Integer)
    If Me.Visible Then Call SetParChange(opt病理诊断, Index, mrsPar)
End Sub

Private Sub opt病理诊断_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub opt病理诊断_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(opt病理诊断, Index, mrsPar)
End Sub

Private Sub opt转科死亡出院医嘱_Click(Index As Integer)
    If Me.Visible Then Call SetParChange(opt转科死亡出院医嘱, Index, mrsPar)
End Sub

Private Sub opt转科死亡出院医嘱_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(opt转科死亡出院医嘱, Index, mrsPar)
End Sub

Private Sub optPrintDruUse_Click(Index As Integer)
    If Me.Visible Then Call SetParChange(optPrintDruUse, Index, mrsPar)
End Sub

Private Sub optPrintDruUse_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optPrintDruUse, Index, mrsPar)
End Sub

Private Sub opt领药部门_Click(Index As Integer)
    If Me.Visible Then Call SetParChange(opt领药部门, Index, mrsPar)
End Sub

Private Sub opt领药部门_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(opt领药部门, Index, mrsPar)
End Sub

Private Sub opt接诊控制_Click(Index As Integer)
    If Me.Visible Then Call SetParChange(opt接诊控制, Index, mrsPar, True, IIF(opt接诊控制(0).value, "0|0", IIF(opt接诊控制(1).value, "1|" & NVL(txt(txt_允许提前接诊分钟).Text, "0"), "2|" & NVL(txt(txt_允许提前接诊分钟).Text, "0"))))
    txt(txt_允许提前接诊分钟).Enabled = Index > 0
End Sub

Private Sub opt接诊控制_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub opt接诊控制_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(opt接诊控制, Index, mrsPar)
End Sub

Private Sub PicColorCollect_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lRow As Long, lCol As Long, lX As Long, lY As Long
    
    If X > 0 And X < PicColorCollect(Index).ScaleWidth And Y > 0 And Y < PicColorCollect(Index).ScaleHeight Then
        SetCapture PicColorCollect(Index).hwnd
        shpBorder(Index).Visible = True
    Else
        ReleaseCapture
        shpBorder(Index).Visible = False
    End If

    
    lRow = Y \ (18 * Screen.TwipsPerPixelY)
    lCol = X \ (18 * Screen.TwipsPerPixelX)
    lX = ((lCol) * 18 + 4) * Screen.TwipsPerPixelX
    lY = ((lRow) * 18 + 4) * Screen.TwipsPerPixelY
    shpBorder(Index).Move lCol * 18 * Screen.TwipsPerPixelX, lRow * 18 * Screen.TwipsPerPixelY, 270, 270
    
    If PicColorCollect(Index).Point(lX, lY) = -1 Then Exit Sub
    picColor(Index).BackColor = PicColorCollect(Index).Point(lX, lY)
    Select Case CStr(Hex(picColor(Index).BackColor))
    Case "0"
        lblColor(Index) = "黑色"
    Case "3399"
        lblColor(Index) = "褐色"
    Case "3333"
        lblColor(Index) = "橄榄色"
    Case "3300"
        lblColor(Index) = "深绿"
    Case "663300"
        lblColor(Index) = "深青"
    Case "800000"
        lblColor(Index) = "深蓝"
    Case "993333"
        lblColor(Index) = "靛蓝"
    Case "333333"
        lblColor(Index) = "灰色-80%"
    Case "80"
        lblColor(Index) = "深红"
    Case "66FF"
        lblColor(Index) = "橙色"
    Case "8080"
        lblColor(Index) = "深黄"
    Case "8000"
        lblColor(Index) = "绿色"
    Case "808000"
        lblColor(Index) = "青色"
    Case "FF0000"
        lblColor(Index) = "蓝色"
    Case "996666"
        lblColor(Index) = "蓝-灰"
    Case "808080"
        lblColor(Index) = "灰色-50%"
    Case "FF"
        lblColor(Index) = "红色"
    Case "99FF"
        lblColor(Index) = "浅橙色"
    Case "CC99"
        lblColor(Index) = "酸橙色"
    Case "669933"
        lblColor(Index) = "海绿"
    Case "CCCC33"
        lblColor(Index) = "水绿色"
    Case "FF6633"
        lblColor(Index) = "浅蓝"
    Case "800080"
        lblColor(Index) = "紫罗兰"
    Case "999999"
        lblColor(Index) = "灰色-40%"
    Case "FF00FF"
        lblColor(Index) = "粉红"
    Case "CCFF"
        lblColor(Index) = "金色"
    Case "FFFF"
        lblColor(Index) = "黄色"
    Case "FF00"
        lblColor(Index) = "鲜绿"
    Case "FFFF00"
        lblColor(Index) = "青绿"
    Case "FFCC00"
        lblColor(Index) = "天蓝"
    Case "663399"
        lblColor(Index) = "梅红"
    Case "C0C0C0"
        lblColor(Index) = "灰色-25%"
    Case "CC99FF"
        lblColor(Index) = "玫瑰红"
    Case "99CCFF"
        lblColor(Index) = "茶色"
    Case "99FFFF"
        lblColor(Index) = "浅黄"
    Case "CCFFCC"
        lblColor(Index) = "浅绿"
    Case "FFFFCC"
        lblColor(Index) = "浅青绿"
    Case "FFCC99"
        lblColor(Index) = "淡蓝"
    Case "FF99CC"
        lblColor(Index) = "淡紫"
    Case "FFFFFF"
        lblColor(Index) = "白色"
    Case Else
        lblColor(Index) = "&H" & CStr(Hex(picColor(Index).BackColor))
    End Select
End Sub

Private Sub PicColorCollect_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strValue As String
    Dim lngIndex As Long
    
    If Button <> vbLeftButton Then Exit Sub
    '按指定颜色作图
    If picControl(Index).Tag = "" Then
        picControl(Index).Visible = False
        Exit Sub
    End If
    lngIndex = Val(picControl(Index).Tag)
    picLineColor(lngIndex).BackColor = picColor(Index).BackColor
    If Me.Visible Then
        Call SetParChange(picLineColor, lngIndex, mrsPar, True, Val(picLineColor(lngIndex).BackColor))
    End If
    picControl(Index).Visible = False
     picControl(Index).Tag = ""
    If picLineColor(lngIndex).Enabled And picLineColor(lngIndex).Visible Then picLineColor(lngIndex).SetFocus
End Sub

Private Sub picLineColor_Click(Index As Integer)
    Dim lngIndex As Long
    
    lngIndex = Index
    If lngIndex = 1 Or lngIndex = 2 Or lngIndex = 3 Or lngIndex = 4 Then lngIndex = 1 '确定颜色选择控件索引
    
    picControl(lngIndex).Top = picLineColor(Index).Top + picLineColor(Index).Height
    picControl(lngIndex).Left = picLineColor(Index).Left
    If picControl(lngIndex).Left + picControl(lngIndex).Width > StabNurse.Width - 60 Then
        picControl(lngIndex).Left = StabNurse.Width - 60 - picControl(lngIndex).Width
    End If
    If picControl(lngIndex).Left < 60 Then picControl(lngIndex).Left = 60
    picControl(lngIndex).Visible = True
    picControl(lngIndex).ZOrder 0
    picControl(lngIndex).Tag = Index
    Call SetCOLOR(Val(picLineColor(Index).BackColor), lngIndex)
End Sub

Private Sub picLineColor_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub picLineColor_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(picLineColor, Index, mrsPar)
End Sub

Private Sub picPar_Resize(Index As Integer)
    On Error Resume Next
    
    Select Case Index
    Case 11     '门诊输液管理
        With vsfWaittingMixDept
            .Width = picPar(Index).ScaleWidth - 360 * 2
            .Height = picPar(Index).ScaleHeight - .Top - 360
        End With
    End Select
End Sub

Private Sub StabNurse_Click(PreviousTab As Integer)
    If Me.Visible Then
        If StabNurse.Tab = 1 Then
            lblPrompt.Caption = "说明:此界面只针对标准通用体温部件,地区部件请在模块内部的护理选项中设置"
        Else
            lblPrompt.Caption = ""
        End If
    End If
End Sub

Private Sub tplFunc_ItemClick(ByVal Item As XtremeSuiteControls.ITaskPanelGroupItem)
    Dim i As Long
    
    For i = 0 To picPar.UBound
        picPar(i).Visible = (i = Item.ID)
    Next
    
    lblLocate(txt_Dept).Visible = (Item.ID = GetFuncID("业务流程控制", marrFunc) Or Item.ID = GetFuncID("住院医嘱操作", marrFunc))
    txtLocate(txt_Dept).Visible = lblLocate(txt_Dept).Visible
    If txtLocate(txt_Dept).Visible Then
        lblPrompt.Left = txtLocate(txt_Dept).Left + txtLocate(txt_Dept).Width + 60
        If Item.ID = GetFuncID("业务流程控制", marrFunc) Then
            lblLocate(txt_Dept).Caption = "科室查找(&F)"
        Else
            lblLocate(txt_Dept).Caption = "病区查找(&F)"
        End If
    Else
        lblPrompt.Left = txtLocate(txt_Par).Left + txtLocate(txt_Par).Width + 60
    End If
    lblPrompt.Width = cmdOK.Left - lblPrompt.Left - 120
    If Item.ID = GetFuncID("新版护理文件", marrFunc) Then
        Call StabNurse_Click(StabNurse.Tab)
    Else
        lblPrompt.Caption = ""
    End If
    mlngPreFind = 1
    
    tplFunc.Tag = Item.ID   '用于获取当前选中的TaskPanelItem
End Sub

Private Sub Form_Resize()
    Dim i As Long
    
    If Me.WindowState = vbMinimized Then Exit Sub
    
    If picVbar.Left < 1500 Then picVbar.Left = 1500
    If picVbar.Left > Me.ScaleWidth - 3000 Then picVbar.Left = Me.ScaleWidth - 3000
    picVbar.Top = 0
    
    picFunc.Width = picVbar.Left + picVbar.Width
    
    For i = 0 To picPar.UBound
        picPar(i).Top = Me.ScaleTop
        picPar(i).Left = picFunc.Left + picFunc.ScaleWidth
        picPar(i).Width = Me.ScaleWidth - picPar(i).Left
        picPar(i).Height = Me.ScaleHeight - PicBottom.ScaleHeight
    Next
End Sub


Private Sub scbFunc_ExpandButtonDown(CancelMenu As Boolean)
    CancelMenu = True
End Sub

Private Sub picBottom_Resize()
    cmdCancel.Left = PicBottom.ScaleWidth - cmdCancel.Width - 120
    cmdOK.Left = cmdCancel.Left - cmdOK.Width - 120
End Sub


Private Sub picFunc_Resize()
    scbFunc.Top = picFunc.ScaleTop
    scbFunc.Left = picFunc.ScaleLeft + 45
    scbFunc.Width = picFunc.ScaleWidth - picVbar.Width - 45
    scbFunc.Height = picFunc.ScaleHeight
    
    picVbar.Height = picFunc.ScaleHeight
End Sub

Private Sub picTPL_Resize()
    sccFunc.Left = picTPL.ScaleLeft
    sccFunc.Width = picTPL.ScaleWidth
    
    tplFunc.Left = picTPL.ScaleLeft
    tplFunc.Top = sccFunc.Top + sccFunc.Height
    tplFunc.Height = picTPL.ScaleHeight - sccFunc.Height
    tplFunc.Width = picTPL.ScaleWidth
End Sub


Private Sub picVbar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        picVbar.Left = IIF(picVbar.Left + X < 2000, 2000, picVbar.Left + X)
        Call Form_Resize
    End If
End Sub

Private Sub scbFunc_SelectedChanged(ByVal Item As XtremeSuiteControls.IShortcutBarItem)
    If Me.Visible Then
        Call InitTPLItem(sccFunc, tplFunc, Item.Caption, marrFunc(Item.ID - 1)) 'ID是从1开始的（因为同时为图标序号）,数组是从0开始
        Call tplFunc_ItemClick(tplFunc.Groups(1).Items(1))
    End If
End Sub


Public Sub LocateFuncItem(ByVal lngFunc As Long)
'功能：根据ID选中一级和二级分类
    Dim i As Long, j As Long, lngId As Long
    Dim arrTmp As Variant
    Dim n As Long
    
    For i = 0 To UBound(marrFunc)
        arrTmp = Split(marrFunc(i), ";")
        For j = 0 To UBound(arrTmp)
            lngId = Split(arrTmp(j), ",")(1)
            If lngFunc = lngId Then
                tplFunc.Tag = lngId
                Set scbFunc.Selected = scbFunc(i)
                
                For n = 1 To tplFunc.Groups(1).Items.Count
                    tplFunc.Groups(1).Items(n).Selected = tplFunc.Groups(1).Items(n).ID = lngId
                Next
            End If
        Next
    Next
End Sub



Private Sub Form_Unload(Cancel As Integer)
    If Not mblnOk Then
        mrsPar.Filter = "(修改状态=1 ANd ErrType =Null) OR  (修改状态=1 And ErrType=" & PET_值超限 & ")"
        If mrsPar.RecordCount > 0 Or cmdAdvice.Tag = "已修改" Or SST.Tag = "已修改" Or Frame14.Tag = "已修改" Then
            If MsgBox("你已修改部分参数，如果你就这样退出的话，所有的修改都不会生效。" & vbCrLf & "是否确认退出？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
                Cancel = 1: Exit Sub
            End If
        End If
    End If
    Set mrsAdvice = Nothing
    Set mcol科室 = Nothing
    Set mcolStop科室 = Nothing
    Set mrsPar = Nothing
    Set mrs门诊药房对照 = Nothing
    Set mrs住院药房对照 = Nothing
    Set mrs住院执行对照 = Nothing
    Set mobjPass = Nothing
    Set gobjNurseIntegrate = Nothing
End Sub

Private Sub InitData()
'功能：初始化界面控件,读取并加载数据
    
    '1.初始化变量
    mlngPreFind = 1
    mblnUseBlood = isPiotBlood
    Set mcol科室 = New Collection
    Set mcolStop科室 = New Collection
    
    Call InitSystemPara
    
    
    '2.初始化界面控件
    Call InitEnv
    
    
    '3.加载系统参数
    Call LoadPar
    
    '4.加载数据后的控件可用性控制
    chk_Click chk_住院下达自动排序
    chk_Click chk_允许前一天不评估就生成今天路径项目
    chk_Click chk_启用路径执行环节
    chk_Click chk_长嘱单转科换页
    chk_Click chk_住院本科自动执行长嘱
    chk_Click chk_科室药房对照按本机参数设置
End Sub
Private Sub LoadPar()
'功能：读取并加载参数到界面控件
    Dim strValue As String, strTmp As String, arrTmp As Variant, varTmp As Variant
    Dim i As Long, n As Long, intIndex As Integer, lngValue As Long
    Dim rsTmp As ADODB.Recordset, rs As ADODB.Recordset, rsData As ADODB.Recordset
    Dim arrObj As Variant  '数组对象：模块1,参数号1(参数名1),控件对象1,模块2,参数号2(参数名2),控件对象2,......,模块参数请使用参数名
    
    Set rsTmp = GetPar(mrsPar, p门诊医嘱下达 & "," & p住院医嘱下达 & "," & p住院医嘱发送 & "," & p临床路径应用 & _
        "," & p住院医生站 & "," & p门诊医生站 & "," & p住院护士站 & "," & p新版住院护士站 & "," & p医技工作站 & "," & p病历内部工具 & _
        "," & p住院病历管理 & "," & p电子病案审查 & "," & p电子病案借阅 & "," & p电子病案评分 & "," & p门诊输液管理 & _
        "," & p护理记录管理 & "," & p临床路径管理 & "," & p门诊路径应用)
    
     '1.设置CheckBox类参数
    strTmp = "0:27:" & chk_住院药嘱发送产生领药号 & _
            ",0:34:" & chk_指定医嘱在其他科室执行 & _
            ",0:43:" & chk_下达出院医嘱才允许出院 & _
            ",0:51:" & chk_本人执行登记 & _
            ",0:56:" & chk_门诊处方条数限制 & _
            ",0:68:" & chk_未作废临嘱禁止退药 & _
            ",0:69:" & chk_药品按规格下医嘱 & _
            ",0:70:" & chk_过敏登记有效天数 & _
            ",0:71:" & chk_长期医嘱次日生效 & _
            ",0:84:" & chk_一次申请多个检验项目 & _
            ",0:143:" & chk_检验医嘱发送生成条形码 & _
            ",0:161:" & chk_禁忌药嘱 & _
            ",0:162:" & chk_下达医嘱时显示产地 & _
            ",0:182:" & chk_禁止下达超极量药品医嘱 & _
            ",0:187:" & chk_抗菌药物分级管理 & _
            ",0:188:" & chk_抗菌药物使用自备药 & _
            ",0:189:" & chk_允许院外执行禁忌药品 & _
            ",0:192:" & chk_回退出院医嘱才允许撤销出院 & _
            ",0:271:" & chk_停嘱时录入原因 & _
            ",0:274:" & chk_科室药房对照按本机参数设置 & _
            ",0:288:" & chk_叮嘱需要发送执行 & _
            ",0:300:" & chk_传染病报告卡强制填写 & _
            ",0:302:" & chk_会诊科室下达医嘱由会诊申请科室处理

    strTmp = strTmp & _
            ",0:208:" & chk_临床工作站必须使用zlPlugIn部件 & _
            ",0:209:" & chk_手术分级管理 & _
            ",0:210:" & chk_允许处理超过挂号有效天数的病人 & _
            ",0:216:" & chk_输血分级管理 & _
            ",0:217:" & chk_手术授权管理 & _
            ",0:218:" & chk_输血申请三级审核 & _
            ",0:219:" & chk_输血申请限制中级及以上医师 & _
            ",0:225:" & chk_接口调用日志_大通 & _
            ",0:226:" & chk_使用系统设置_美康 & _
            ",0:230:" & chk_医嘱超量时必须输入原因 & _
            ",0:236:" & chk_启用血库管理系统 & _
            ",0:237:" & chk_多科会诊意见书写要求 & _
            ",0:240:" & chk_医嘱用药天数反算 & _
            ",0:247:" & chk_身份证密文 & _
            ",0:248:" & chk_按医疗小组进行抗菌药物审核 & _
            ",0:249:" & chk_禁忌药品要求填写原因 & _
            ",0:250:" & chk_手术分级审核 & _
            ",0:254:" & chk_主刀医师达到手术等级无需审核 & _
            ",0:257:" & chk_诊断手术名称自由调整 & _
            ",0:259:" & chk_路径变异原因从字典表中选取 & _
            ",0:262:" & chk_特殊药品分开发送 & _
            ",0:272:" & chk_用血医嘱发送后才能发血 & _
            ",0:286:" & chk_输血申请不显示血液库存 & _
            ",0:289:" & chk_诊断录入时附码自动提取
    strTmp = strTmp & _
            ",0:293:" & chk_下达用血申请时确定发血信息 & _
            ",0:301:" & chk_血液接收后才允许执行登记 & _
            ",0:307:" & chk_不填写病理号

    strTmp = strTmp & _
            "," & p门诊医嘱下达 & ":医嘱执行天数:" & chk_门诊药嘱输天数 & _
            "," & p门诊医嘱下达 & ":自动增加皮试医嘱:" & chk_门诊下达加皮试 & _
            "," & p门诊医嘱下达 & ":必须录入药品单量:" & chk_门诊药嘱输单量 & _
            "," & p门诊医嘱下达 & ":单位记帐:" & chk_合约单位发送记帐 & _
            "," & p门诊医嘱下达 & ":门诊本科自动执行:" & chk_门诊发送本科自动执行 & _
            "," & p门诊医嘱下达 & ":要求登记代办人:" & chk_门诊登记代办人 & _
            "," & p门诊医嘱下达 & ":一并给药发送为一张:" & chk_一并给给一张单据 & _
            "," & p门诊医嘱下达 & ":不同诊断的医嘱分别产生单据:" & chk_不同诊断分单据 & _
            "," & p门诊医嘱下达 & ":门诊医嘱发送后启用诊间支付:" & chk_门诊医嘱发送后启用诊间支付 & _
            "," & p门诊医嘱下达 & ":开始时间不是同一天的分别产生单据:" & chk_不同天的分单据 & _
            "," & p门诊医嘱下达 & ":显示药品库存:" & chk_门诊医嘱下达时诊疗选择器显示药品库存 & _
            "," & p门诊医嘱下达 & ":检验医嘱单独产生单据:" & chk_门诊检验医嘱发送时一组检验发送为一张单据 & _
            "," & p门诊医嘱下达 & ":医嘱发送皮试限制:" & chk_门诊医嘱发送皮试限制 & _
            "," & p门诊医嘱下达 & ":皮试阳性用药:" & chk_皮试阳性用药门诊 & _
            "," & p门诊医嘱下达 & ":门诊西医科允许录入中医诊断:" & chk_门诊西医科允许录入中医诊断
    
    strTmp = strTmp & _
            "," & p住院医嘱下达 & ":临嘱缺省一次性:" & chk_住院临嘱缺省一次性 & _
            "," & p住院医嘱下达 & ":临嘱先输入单量:" & chk_住院药嘱输单量 & _
            "," & p住院医嘱下达 & ":医嘱执行天数:" & chk_住院药嘱输天数 & _
            "," & p住院医嘱下达 & ":要求输入出院诊断:" & chk_下出院医嘱检查出院诊断 & _
            "," & p住院医嘱下达 & ":手术完成后下达术后医嘱:" & chk_手术执行后下术后医嘱 & _
            "," & p住院医嘱下达 & ":医嘱自动排序:" & chk_住院下达自动排序 & _
            "," & p住院医嘱下达 & ":实习医生停止医嘱需要审核:" & chk_实习医生停嘱需审核 & _
            "," & p住院医嘱下达 & ":允许给待入住病人下达医嘱:" & chk_待入住病人医嘱下达 & _
            "," & p住院医嘱下达 & ":自动增加皮试医嘱:" & chk_住院下达加皮试 & _
            "," & p住院医嘱下达 & ":显示药品库存:" & chk_住院医嘱下达时诊疗选择器显示药品库存 & _
            "," & p住院医嘱下达 & ":发送完成后关闭医嘱窗体:" & chk_发送完成后关闭医嘱窗体 & _
            "," & p住院医嘱下达 & ":要求登记代办人:" & chk_住院登记代办人 & _
            "," & p住院医嘱下达 & ":医嘱发送皮试限制:" & chk_住院医嘱发送皮试限制 & _
            "," & p住院医嘱下达 & ":皮试阳性用药:" & chk_皮试阳性用药住院
            
   strTmp = strTmp & _
            "," & p住院医嘱发送 & ":自动进入医嘱打印:" & chk_校对确诊停止重整后打印 & _
            "," & p住院医嘱发送 & ":批量医嘱校对:" & chk_批量校对 & _
            "," & p住院医嘱发送 & ":批量医嘱启停:" & chk_批量暂停 & _
            "," & p住院医嘱发送 & ":皮试验证身份:" & chk_登记皮试验证身份 & _
            "," & p住院医嘱发送 & ":医技医嘱后续处理:" & chk_允许处理医技医嘱 & _
            "," & p住院医嘱发送 & ":超期收回费用本科自动审核:" & chk_超期收回自动审核本科 & _
            "," & p住院医嘱发送 & ":校对医嘱电子签名:" & chk_校对确诊停止电子签名 & _
            "," & p住院医嘱发送 & ":药嘱发送限制结束时间:" & chk_药嘱发送限制结束时间 & _
            "," & p住院医嘱发送 & ":检查医保审批:" & chk_住院发送检查医保审批 & _
            "," & p住院医嘱发送 & ":发送前自动校对:" & chk_住院无须校对发送 & _
            "," & p住院医嘱发送 & ":停止后自动超期收回:" & chk_确认停止后自动收回 & _
            "," & p住院医嘱发送 & ":特殊医嘱发送前检查未生效医嘱:" & chk_特殊医嘱发送检查未生效医嘱 & _
            "," & p住院医嘱发送 & ":常用报表逐个打印:" & chk_执行单打印时逐个病人换页打印 & _
            "," & p住院医嘱发送 & ":长嘱单转科换页:" & chk_长嘱单转科换页 & _
            "," & p住院医嘱发送 & ":长嘱单术后换页:" & chk_长嘱单术后换页 & _
            "," & p住院医嘱发送 & ":长嘱单重整换页:" & chk_长嘱单重整换页 & _
            "," & p住院医嘱发送 & ":转科换页后在首行打印重开医嘱:" & chk_转科换页后在首行打印重开医嘱 & _
            "," & p住院医嘱发送 & ":只显示当前病区的医嘱:" & chk_收回时医嘱列表只显示当前病区的医嘱 & _
            "," & p住院医嘱发送 & ":检验医嘱单独产生单据:" & chk_住院检验医嘱发送时一组检验发送为一张单据 & _
            "," & p住院医嘱发送 & ":临嘱单转科换页:" & chk_临嘱单转科换页 & _
            "," & p住院医嘱发送 & ":临嘱单术后换页:" & chk_临嘱单术后换页 & _
            "," & p住院医嘱发送 & ":存在未发送医嘱时禁止处理转科医嘱:" & chk_存在未发送医嘱时禁止处理转科医嘱
    
    strTmp = strTmp & _
            "," & p临床路径应用 & ":是否启用路径执行环节:" & chk_启用路径执行环节 & _
            "," & p临床路径应用 & ":医技医嘱在路径表外:" & chk_路径病人医技医嘱显示 & _
            "," & p临床路径应用 & ":未评估时允许添加医嘱到昨天:" & chk_未评估时允许添加医嘱到昨天 & _
            "," & p临床路径应用 & ":允许前一天不评估就生成今天的路径项目:" & chk_允许前一天不评估就生成今天路径项目 & _
            "," & p临床路径应用 & ":允许提前生成明天的路径项目:" & chk_允许提前生成明天的路径项目 & _
            "," & p临床路径应用 & ":匹配时期效不同算路径外项目:" & chk_匹配时期效不同算路径外项目 & _
            "," & p临床路径应用 & ":出院后不允许取消完成路径:" & chk_出院后不允许取消完成路径 & _
            "," & p临床路径应用 & ":药品医嘱不匹配为路径外项目:" & chk_药品医嘱不匹配为路径外项目 & _
            "," & p临床路径应用 & ":药品医嘱相同分类不算路径外医嘱:" & chk_药品医嘱相同分类不算路径外医嘱
            
    strTmp = strTmp & "," & p临床路径管理 & ":双审核模式:" & chk_启用药剂科和医务科双审核模式
     
    strTmp = strTmp & _
            "," & p门诊路径应用 & ":是否启用路径执行环节:" & chk_门诊启用路径执行环节
    strTmp = strTmp & _
            "," & p住院医生站 & ":使用手术结束时间:" & chk_使用手术结束时间 & _
            "," & p住院医生站 & ":中医科室不使用西医病案首页项目:" & chk_中医科室不使用项目 & _
            "," & p住院医生站 & ":医生和护士分别填写病案首页:" & chk_医生和护士分别填写病案首页 & _
            "," & p住院医生站 & ":不显示无床位的病区科室:" & chk_不显示无床位的病区科室 & _
            "," & p住院医生站 & ":病理诊断只允许录入肿瘤形态学编码:" & chk_病理诊断只允许录入M打头的肿瘤形态学编码 & _
            "," & p住院医生站 & ":门诊诊断不作为导入临床路径的诊断依据:" & chk_门诊诊断不作为导入临床路径的诊断依据
            
    strTmp = strTmp & _
            "," & p门诊医生站 & ":就诊人数含回诊:" & chk_医生呼叫人数限制含回诊病人 & _
            "," & p门诊医生站 & ":医生主动呼叫后才允许接诊:" & chk_医生主动呼叫后才允许在队列中接诊 & _
            "," & p门诊医生站 & ":只接收已经分诊的病人:" & chk_只接收已经分诊的病人
            
    strTmp = strTmp & _
            "," & p医技工作站 & ":未收费完成:" & chk_允许未收费病人完成执行 & _
            "," & p医技工作站 & ":皮试验证身份:" & chk_填写皮试结果时验证身份 & _
            "," & p医技工作站 & ":执行报到时收费或记账审核:" & chk_执行报到时收费或记账审核 & _
            "," & p医技工作站 & ":血透室书写新版护理记录:" & chk_血透室书写新版护理记录
            
    strTmp = strTmp & _
            "," & p病历内部工具 & ":签名自动位移:" & chk_签名移位 & _
            "," & p病历内部工具 & ":显示手签位置:" & chk_显示手签位置 & _
            "," & p病历内部工具 & ":将签名级别作为前缀加入:" & chk_签名级别前辍 & _
            "," & p病历内部工具 & ":SyncPage:" & chk_病历诊断同步首页 & _
            "," & p病历内部工具 & ":签名使用原图:" & chk_签名使用原图 & _
            "," & p住院病历管理 & ":转科后要求书写的共享病历另起一页打印:" & chk_转科后书写的病历另起一页打印 & _
            "," & p住院病历管理 & ":自动显示新增面板:" & chk_住院病历自动显示新增面板 & _
            "," & p住院病历管理 & ":共享病历必须先书写被共享病历:" & chk_要求先书写被共享病历 & _
            "," & p电子病案审查 & ":接收才能归档:" & chk_病案必须审查才能归档 & _
            "," & p电子病案审查 & ":允许自由录入审查意见:" & chk_允许自由录入审查意见 & _
            "," & p电子病案借阅 & ":必须录入借阅原因:" & chk_必须录入借阅申请理由 & _
            "," & p电子病案借阅 & ":允许自由录入借阅原因:" & chk_允许自由录入借阅原因 & _
            "," & 0 & ":90:" & chk_评分自动写入病案主页 & _
            "," & 0 & ":91:" & chk_病案必须先编目后评分

    strTmp = strTmp & _
            "," & p门诊输液管理 & ":皮试验证身份:" & chk_门诊输液皮试验证身份 & _
            "," & p门诊输液管理 & ":未收费完成:" & chk_门诊输液未收费允许完成
    
    strTmp = strTmp & _
        "," & p新版住院护士站 & ":卡片余额含担保金额:" & chk_卡片余额含担保金额
    Call SetParToControl(strTmp, mrsPar, chk)
    '产程图
    strTmp = p护理记录管理 & ":产程图显示产程时间:" & chk_产程图显示产程时间 & _
            "," & p护理记录管理 & ":产程图模式:" & chk_产程图模式 & _
            "," & p护理记录管理 & ":先露高低显示位置:" & chk_先露高低显示位置 & _
            "," & p护理记录管理 & ":产程图显示警戒线:" & chk_产程图上显示警戒线
    '记录单
    strTmp = strTmp & _
            "," & p护理记录管理 & ":记录单日期显示方式:" & chk_日期显示方式 & _
            "," & p护理记录管理 & ":跨页数据只显示在第一页:" & chk_只在当前页显示跨页数据 & _
            "," & p护理记录管理 & ":对应多份护理文件:" & chk_对应多份护理文件 & _
            "," & p护理记录管理 & ":护理文件页码规则:" & chk_文件页码规则 & _
            "," & p护理记录管理 & ":记录单签名人显示方式:" & chk_签名人显示方式 & _
            "," & p护理记录管理 & ":允许数据同步:" & chk_允许数据同步
    '体温单
    strTmp = strTmp & _
            "," & p护理记录管理 & ":体温标志输出位置:" & chk_自动标志顶格输出 & _
            "," & p护理记录管理 & ":体温标志超出40刻度缩小字体显示:" & chk_自动标志超出40刻度缩小显示 & _
            "," & p护理记录管理 & ":再次手术停止前次标注:" & chk_再次手术停止前次标注 & _
            "," & p护理记录管理 & ":体温标志按顺序当天排列:" & chk_自动标志按顺序当天排列 & _
            "," & p护理记录管理 & ":首列日期格式:" & chk_首列日期格式 & _
            "," & p护理记录管理 & ":入科标识不自动转换为入院:" & chk_入科标志不自动转为入院 & _
            "," & p护理记录管理 & ":病人术后不足14天出院标记显示:" & chk_不足术后天数出院 & _
            "," & p护理记录管理 & ":全天汇总显示录入时间:" & chk_全天汇总显示小时 & _
            "," & p护理记录管理 & ":汇总波动显示当天数据:" & chk_汇总项目显示当天数据 & _
            "," & p护理记录管理 & ":体温单不打印心率列:" & chk_不输出心率列 & _
            "," & p护理记录管理 & ":体温单显示格式:" & chk_绘图刻度单格显示 & _
            "," & p护理记录管理 & ":体温单显示诊断:" & chk_显示诊断信息 & _
            "," & p护理记录管理 & ":婴儿体温单首日天数显示0:" & chk_婴儿住院天数从0开始计算 & _
            "," & p护理记录管理 & ":打印医院名称:" & chk_打印医院名称 & _
            "," & p护理记录管理 & ":体温单不打印曲线说明:" & chk_输出曲线说明信息 & _
            "," & p护理记录管理 & ":曲线项目骑线显示:" & chk_曲线项目骑线显示
    Call SetParToControl(strTmp, mrsPar, chk)
    
    '2.设置ComboBox类参数
    strTmp = "0:30:" & cbo_合理用药接口 & _
            ",0:224:" & cmd_过敏输入来源 & _
            "," & p住院医生站 & ":部门显示方式:" & cbo_住院医生站列表显示 & _
            "," & p住院医生站 & ":病案首页标准:" & cbo_住院首页标准 & _
            "," & p病历内部工具 & ":签名时间:" & cbo_病历签名显示时间
    '记录单
    strTmp = strTmp & _
            "," & p护理记录管理 & ":记录单审签模式:" & cbo_审签模式 & _
            "," & p护理记录管理 & ":护士、签名列显示模式:" & cbo_签名列显示模式
    '体温单
    strTmp = strTmp & _
            "," & p护理记录管理 & ":未记说明显示位置:" & cbo_未记说明显示位置 & _
            "," & p护理记录管理 & ":体温不升显示方式:" & cbo_体温不升显示方式 & _
            "," & p护理记录管理 & ":呼吸表格呼吸机输出方式:" & cbo_呼吸机符号显示位置 & _
            "," & p护理记录管理 & ":呼吸表格输出:" & cbo_呼吸表格数据显示位置 & _
            "," & p护理记录管理 & ":手术当天缺省格式:" & cbo_手术当天缺省格式 & _
            "," & p护理记录管理 & ":体温标志分隔符:" & cbo_标志说名与时间连接符号
    Call SetParToControl(strTmp, mrsPar, cbo)
            
    '3.设置UpDown类参数
    strTmp = "0:5:" & ud_补录医嘱识别间隔 & _
            ",0:220:" & ud_医嘱执行有效天数 & _
            ",0:223:" & ud_门诊新开医嘱间隔
    
    '记录单
    strTmp = strTmp & "," & p护理记录管理 & ":超期录入护理数据天数:" & ud_记录单超期录入天数
    '体温单
    strTmp = strTmp & _
            "," & p护理记录管理 & ":体温曲线固定添加行数:" & ud_体温曲线固定添加行数 & _
            "," & p护理记录管理 & ":体温开始时间:" & ud_体温开始时间 & _
            "," & p护理记录管理 & ":体温表格行数:" & ud_体温表格固定添加行数
    Call SetParToControl(strTmp, mrsPar, ud)     'mrsPar存储的控件名是txtUD
    
    '4.设置TextBox类参数
    strTmp = p门诊医嘱下达 & ":输血申请注意事项:" & txt_门诊输血申请注意事项 & _
            "," & p住院医嘱下达 & ":输血申请注意事项:" & txt_住院输血申请注意事项 & _
            "," & p临床路径应用 & ":中药配方允许修改的中药味数上限:" & txt_中药配方允许修改的中药味数上限 & _
            "," & p临床路径应用 & ":路径医嘱生成超前天数:" & txt_生成时医嘱允许超过当前时间天数 & _
            "," & p门诊路径应用 & ":中药配方允许修改的中药味数上限:" & txt_门诊中药配方允许修改的中药味数上限 & _
            "," & p门诊路径应用 & ":路径医嘱生成超前天数:" & txt_门诊生成时医嘱允许超过当前时间天数 & _
            "," & p门诊医生站 & ":候诊刷新间隔:" & txt_门诊医生站病人列表自动刷新秒数 & _
            "," & p住院医生站 & ":病案审查反馈天数:" & txt_住院医生病案反馈数 & _
            "," & p住院护士站 & ":病案审查反馈天数:" & txt_住院护士病案反馈数 & _
            "," & p住院医嘱发送 & ":长嘱单中药医嘱单行显示字数:" & txt_长嘱单中药医嘱单行显示字数 & _
            "," & p住院医嘱发送 & ":临嘱单中药医嘱单行显示字数:" & txt_临嘱单中药医嘱单行显示字数


    strTmp = strTmp & _
            "," & p病历内部工具 & ":签名图片高度:" & txt_签名使用图片高度 & _
            "," & p住院病历管理 & ":共享病历折叠起始行数:" & txt_共享病历N行自动折叠 & _
            "," & p电子病案审查 & ":反馈处理期限:" & txt_电子病案审查缺省期限 & _
            "," & p电子病案审查 & ":未复查刷新频率:" & txt_电子病案审查待复查刷新间隔 & _
            "," & p电子病案借阅 & ":病案借阅期限:" & txt_电子病案借阅缺省期限 & _
            "," & p电子病案借阅 & ":借阅最长期限:" & txt_电子病案借阅最大期限
    
    strTmp = strTmp & _
            "," & p门诊输液管理 & ":医技刷新间隔:" & txt_门诊输液自动刷新病人 & _
            "," & p门诊输液管理 & ":皮试提醒提前时间:" & txt_门诊输液皮试提醒 & _
            "," & p门诊输液管理 & ":默认滴速:" & txt_门诊输液滴速 & _
            "," & p门诊输液管理 & ":默认滴系数:" & txt_门诊输液滴系数 & _
            "," & p门诊输液管理 & ":输液提醒提前时间:" & txt_门诊输液输液提醒
    
    '体温单
    strTmp = strTmp & _
        "," & p护理记录管理 & ":手术后标注天数:" & txt_手术标注天数 & _
        "," & p护理记录管理 & ":体温复试合格符号:" & txt_体温复试合格符号
    Call SetParToControl(strTmp, mrsPar, txt)
    
    '5.设置ListBox类参数
    strTmp = "0:80:" & lst_住院发送划价类别 & _
            ",0:86:" & lst_门诊发送划价类别 & _
            "," & p门诊医嘱下达 & ":要求输入门诊诊断:" & lst_门诊发送检查诊断 & _
            "," & p门诊医嘱下达 & ":产生为同一单据的医嘱类别:" & lst_门诊发送一张单据类别 & _
            "," & p住院医嘱下达 & ":要求输入入院诊断:" & lst_住院检查入院诊断
    Call SetParToControl(strTmp, mrsPar, lst)
        
    strTmp = p住院医嘱发送 & ":发药后不收回:" & lst_不收回的发药类型
    Call SetParToControl(strTmp, mrsPar, lst, 2)
    
    strTmp = p住院医嘱发送 & ":留存登记给药途径限制:" & lst_药品留存给药途径
    Call SetParToControl(strTmp, mrsPar, lst, 3)
    
    '6.设置OptionButton类参数
    arrObj = Array(p门诊医嘱下达, "发送单据类型", opt发送单据类型, _
                    p门诊医嘱下达, "发送单据号规则", opt发送单据规则, _
                    p门诊医嘱下达, "抗菌药物缺省用药目的", opt抗菌目的门诊, _
                    p住院医嘱下达, "医嘱单打印模式", opt住院医嘱单打印, _
                    p住院医嘱下达, "抗菌药物缺省用药目的", opt抗菌目的住院, _
                    p住院医嘱下达, "根据皮试结果限制医嘱发送类型", opt未皮试限制医嘱, _
                    p住院医嘱发送, "超期收回产生负数费用", opt超期费用收回, _
                    p住院医嘱发送, "输血申请单打印模式", opt输血申请单打印, _
                    p住院医嘱发送, "转科和出院打印", opt转科死亡出院医嘱, _
                    p住院医嘱发送, "住院领药部门", opt领药部门, _
                    p临床路径应用, "路径表单打印规则", opt路径打印规则, _
                    p临床路径应用, "路径表单每页打印的天数", opt每页路径打印天数, _
                    p门诊路径应用, "路径表单打印规则", opt门诊路径打印规则, _
                    p门诊路径应用, "路径表单每页打印的天数", opt门诊路径打印天数, _
                    p住院医生站, "损伤中毒检查", opt损伤中毒, _
                    p住院医生站, "病理诊断检查", opt病理诊断, _
                    p住院医生站, "ICD附码检查", optICD附码, _
                    p住院医生站, "区域检查", opt区域, _
                    p医技工作站, "病人过滤方式", opt病人过滤, _
                    p病历内部工具, "SignShow", optSign, _
                    p新版住院护士站, "床位卡片排序方式", optNewCard, _
                    p住院医嘱发送, "药品用法单独打印一行", optPrintDruUse)
    Call SetParToControl("", mrsPar, arrObj)
    
    arrObj = Array(p临床路径应用, "路径表单打印方式", optPrintWay)
    Call SetParToControl("", mrsPar, arrObj)
    
    arrObj = Array(p门诊医生站, "病人接诊控制", opt接诊控制)
    Call SetParToControl("", mrsPar, arrObj, 1)
    '体温单
    arrObj = Array(p护理记录管理, "脉搏短绌填充方式", optPloy, _
                    p护理记录管理, "脉搏短绌以(心率/脉搏)方式录入", OptInsert, _
                    p护理记录管理, "出院标志优先显示", OptOut, _
                    p护理记录管理, "体温单文件开始时间", optFileTime, _
                    p护理记录管理, "灌肠后大便显示格式", OptEnemaStool)
    Call SetParToControl("", mrsPar, arrObj)
    
    '7.其他系统参数
    rsTmp.Filter = "模块=0"
    Do Until rsTmp.EOF
        strValue = "" & rsTmp!参数值
        Select Case rsTmp!参数号
        Case 56
            ud(ud_门诊处方条数限制).value = IIF(Val(strValue) = 0, 5, Val(strValue))
            
            Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")   '已有CheckBox控件，所以需要再产生一条记录
            Call SetParRelation(txtUD, ud_门诊处方条数限制, mrsPar)
                        
        Case 70
            ud(ud_过敏登记有效天数).value = IIF(Val(strValue) = 0, 1, Val(strValue))
            
            Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "") '已有CheckBox控件，所以需要再产生一条记录
            Call SetParRelation(txtUD, ud_过敏登记有效天数, mrsPar)
        Case 186
            chk(chk_输血医嘱执行后需要核对).value = Mid(strValue, 1, 1)
            chk(chk_皮试医嘱执行后需要核对).value = Mid(strValue, 2, 1)
            
            Call SetParRelation(chk, chk_输血医嘱执行后需要核对, mrsPar, rsTmp!参数号)
            Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
            Call SetParRelation(chk, chk_皮试医嘱执行后需要核对, mrsPar)
        Case 188
            chk(chk_抗菌药物使用自备药).Enabled = chk(chk_抗菌药物分级管理).value = 1
        Case 248
            chk(chk_按医疗小组进行抗菌药物审核).Enabled = chk(chk_抗菌药物分级管理).value = 1
        Case 213
            cbo(cbo_中药配方).ListIndex = IIF(Val(strValue) = 4, 1, 0)
            Call SetParRelation(cbo, cbo_中药配方, mrsPar, rsTmp!参数号)
        Case 220    '允许取消n天内登记的医嘱执行记录
            chk(chk_医嘱执行有效天数).value = IIF(Val(strValue) = 999, 0, 1)
            
            Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "") '已有txtUD控件，所以需要再产生一条记录
            Call SetParRelation(chk, chk_医嘱执行有效天数, mrsPar)
            
        Case 228   '美康合理用药接口版本
            strTmp = NVL(strValue, "3.0")
            If strTmp = "3.0" Then
                optPASSVer(0).value = True
            Else
                optPASSVer(1).value = True
            End If
            If Not mobjPass Is Nothing Then
                cmdSet.Visible = mobjPass.SetEnabled(cbo(cbo_合理用药接口).ListIndex, strTmp)
            End If
            Call SetParRelation(optPASSVer, 0, mrsPar, rsTmp!参数号)
            Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
            Call SetParRelation(optPASSVer, 1, mrsPar)
        Case 230
            Call SetVsfEditable(vsUnWriteDept, Val(strValue) = 1)
        Case 271
            Call SetVsfEditable(vsStopDept, Val(strValue) = 1)
        Case 273
            If mblnUseBlood Then
                Call zlControl.cbo.Locate(cbo(cbo_输血采集默认诊疗类型), strValue)
                Call SetParRelation(cbo, cbo_输血采集默认诊疗类型, mrsPar, rsTmp!参数号)
            End If
        Case 233
            Call Load科室(vsUnWriteDept, strValue)
            Call SetParRelation(vsUnWriteDept, 0, mrsPar, rsTmp!参数号)
        Case 285
            Call Load科室(vsStopDept, strValue)
            Call SetParRelation(vsStopDept, 0, mrsPar, rsTmp!参数号)
        Case 238
            Call Set申请单启用环节(strValue)
            Call SetParRelation(chk, chk_申请单启用环节门诊检查, mrsPar, rsTmp!参数号)
            Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
            Call SetParRelation(chk, chk_申请单启用环节门诊检验, mrsPar)
            Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
            Call SetParRelation(chk, chk_申请单启用环节门诊输血, mrsPar)
            Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
            Call SetParRelation(chk, chk_申请单启用环节门诊手术, mrsPar)
            Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
            Call SetParRelation(chk, chk_申请单启用环节住院检查, mrsPar)
            Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
            Call SetParRelation(chk, chk_申请单启用环节住院检验, mrsPar)
            Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
            Call SetParRelation(chk, chk_申请单启用环节住院输血, mrsPar)
            Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
            Call SetParRelation(chk, chk_申请单启用环节住院手术, mrsPar)
            Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
            Call SetParRelation(chk, chk_申请单启用环节会诊, mrsPar)
        Case 260
            chk(chk_启用申请单后必须使用申请单下达医嘱门诊).value = Mid(strValue, 1, 1)
            chk(chk_启用申请单后必须使用申请单下达医嘱住院).value = Mid(strValue, 2, 1)
            Call SetParRelation(chk, chk_启用申请单后必须使用申请单下达医嘱门诊, mrsPar, rsTmp!参数号)
            Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
            Call SetParRelation(chk, chk_启用申请单后必须使用申请单下达医嘱住院, mrsPar)
        End Select
        
        rsTmp.MoveNext
    Loop
    
    '8.其他模块参数设置
    rsTmp.Filter = "模块=" & p门诊医嘱下达
    Do Until rsTmp.EOF
        strValue = "" & rsTmp!参数值
        Select Case rsTmp!参数名
            Case "药房科室对照方案"
                '药房科室对照方案
                For i = 0 To UBound(Split(strValue, ";"))
                    cbo(cbo_门诊药房科室对照方案).AddItem "方案" & i + 1
                    cbo(cbo_门诊药房科室对照方案).ItemData(cbo(cbo_门诊药房科室对照方案).NewIndex) = i + 1
                    mrs门诊药房对照.AddNew
                    mrs门诊药房对照!科室IDs = Split(strValue, ";")(i)
                    mrs门诊药房对照!方案 = i + 1
                    mrs门诊药房对照!缺省西药房 = zlDatabase.GetPara("门诊缺省西药房", glngSys, p门诊医嘱下达, , , , , Val(Split(mrs门诊药房对照!科室IDs, ",")(0)))
                    mrs门诊药房对照!缺省成药房 = zlDatabase.GetPara("门诊缺省成药房", glngSys, p门诊医嘱下达, , , , , Val(Split(mrs门诊药房对照!科室IDs, ",")(0)))
                    mrs门诊药房对照!缺省中药房 = zlDatabase.GetPara("门诊缺省中药房", glngSys, p门诊医嘱下达, , , , , Val(Split(mrs门诊药房对照!科室IDs, ",")(0)))
                    mrs门诊药房对照!可用西药房 = zlDatabase.GetPara("门诊可用西药房", glngSys, p门诊医嘱下达, , , , , Val(Split(mrs门诊药房对照!科室IDs, ",")(0)))
                    mrs门诊药房对照!可用成药房 = zlDatabase.GetPara("门诊可用成药房", glngSys, p门诊医嘱下达, , , , , Val(Split(mrs门诊药房对照!科室IDs, ",")(0)))
                    mrs门诊药房对照!可用中药房 = zlDatabase.GetPara("门诊可用中药房", glngSys, p门诊医嘱下达, , , , , Val(Split(mrs门诊药房对照!科室IDs, ",")(0)))
                    mrs门诊药房对照!缺省发料部门 = zlDatabase.GetPara("门诊缺省发料部门", glngSys, p门诊医嘱下达, , , , , Val(Split(mrs门诊药房对照!科室IDs, ",")(0)))
                    mrs门诊药房对照!可用发料部门 = zlDatabase.GetPara("门诊可用发料部门", glngSys, p门诊医嘱下达, , , , , Val(Split(mrs门诊药房对照!科室IDs, ",")(0)))
                    mrs门诊药房对照.Update
                Next
                If cbo(cbo_门诊药房科室对照方案).ListCount > 0 Then
                    cbo(cbo_门诊药房科室对照方案).ListIndex = 0
                End If
                Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
                Call SetParRelation(vsUseDept, cbo_门诊药房科室对照方案, mrsPar)
            Case "门诊缺省西药房"
                '门诊缺省西药房
                Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
                Call SetParRelation(vsfDrugStore, cbo_门诊药房科室对照方案, mrsPar, , , "门诊缺省西药房")
            Case "门诊缺省成药房"
                '门诊缺省成药房
                Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
                Call SetParRelation(vsfDrugStore, cbo_门诊药房科室对照方案, mrsPar, , , "门诊缺省成药房")
            Case "门诊缺省中药房"
                '门诊缺省中药房
                Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
                Call SetParRelation(vsfDrugStore, cbo_门诊药房科室对照方案, mrsPar, , , "门诊缺省中药房")
            Case "门诊缺省发料部门"
                '门诊缺省发料部门
                Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
                Call SetParRelation(vsfDrugStore, cbo_门诊药房科室对照方案, mrsPar, , , "门诊缺省发料部门")
            Case "门诊可用西药房"
                '门诊可用西药房
                Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
                Call SetParRelation(vsfDrugStore, cbo_门诊药房科室对照方案, mrsPar, , , "门诊可用西药房")
            Case "门诊可用成药房"
                '门诊可用成药房
                Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
                Call SetParRelation(vsfDrugStore, cbo_门诊药房科室对照方案, mrsPar, , , "门诊可用成药房")
            Case "门诊可用中药房"
                '门诊可用中药房
                Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
                Call SetParRelation(vsfDrugStore, cbo_门诊药房科室对照方案, mrsPar, , , "门诊可用中药房")
            Case "门诊可用发料部门"
                '门诊可用发料部门
                Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
                Call SetParRelation(vsfDrugStore, cbo_门诊药房科室对照方案, mrsPar, , , "门诊可用发料部门")
        End Select
        
        rsTmp.MoveNext
    Loop
    
    rsTmp.Filter = "模块=" & p住院医嘱下达
    Do Until rsTmp.EOF
        strValue = "" & rsTmp!参数值
        Select Case rsTmp!参数名
            Case "药房科室对照方案"
                '药房科室对照方案
                For i = 0 To UBound(Split(strValue, ";"))
                    cbo(cbo_住院药房科室对照方案).AddItem "方案" & i + 1
                    cbo(cbo_住院药房科室对照方案).ItemData(cbo(cbo_住院药房科室对照方案).NewIndex) = i + 1
                    mrs住院药房对照.AddNew
                    mrs住院药房对照!科室IDs = Split(strValue, ";")(i)
                    mrs住院药房对照!方案 = i + 1
                    mrs住院药房对照!缺省西药房 = zlDatabase.GetPara("住院缺省西药房", glngSys, p住院医嘱下达, , , , , Val(Split(mrs住院药房对照!科室IDs, ",")(0)))
                    mrs住院药房对照!缺省成药房 = zlDatabase.GetPara("住院缺省成药房", glngSys, p住院医嘱下达, , , , , Val(Split(mrs住院药房对照!科室IDs, ",")(0)))
                    mrs住院药房对照!缺省中药房 = zlDatabase.GetPara("住院缺省中药房", glngSys, p住院医嘱下达, , , , , Val(Split(mrs住院药房对照!科室IDs, ",")(0)))
                    mrs住院药房对照!可用西药房 = zlDatabase.GetPara("住院可用西药房", glngSys, p住院医嘱下达, , , , , Val(Split(mrs住院药房对照!科室IDs, ",")(0)))
                    mrs住院药房对照!可用成药房 = zlDatabase.GetPara("住院可用成药房", glngSys, p住院医嘱下达, , , , , Val(Split(mrs住院药房对照!科室IDs, ",")(0)))
                    mrs住院药房对照!可用中药房 = zlDatabase.GetPara("住院可用中药房", glngSys, p住院医嘱下达, , , , , Val(Split(mrs住院药房对照!科室IDs, ",")(0)))
                    mrs住院药房对照!缺省发料部门 = zlDatabase.GetPara("住院缺省发料部门", glngSys, p住院医嘱下达, , , , , Val(Split(mrs住院药房对照!科室IDs, ",")(0)))
                    mrs住院药房对照!可用发料部门 = zlDatabase.GetPara("住院可用发料部门", glngSys, p住院医嘱下达, , , , , Val(Split(mrs住院药房对照!科室IDs, ",")(0)))
                    mrs住院药房对照.Update
                Next
                If cbo(cbo_住院药房科室对照方案).ListCount > 0 Then
                    cbo(cbo_住院药房科室对照方案).ListIndex = 0
                End If
                Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
                Call SetParRelation(vsUseDept, cbo_住院药房科室对照方案, mrsPar)
            Case "住院缺省西药房"
                '住院缺省西药房
                Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
                Call SetParRelation(vsfDrugStore, cbo_住院药房科室对照方案, mrsPar, , , "住院缺省西药房")
            Case "住院缺省成药房"
                '住院缺省成药房
                Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
                Call SetParRelation(vsfDrugStore, cbo_住院药房科室对照方案, mrsPar, , , "住院缺省成药房")
            Case "住院缺省中药房"
                '住院缺省中药房
                Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
                Call SetParRelation(vsfDrugStore, cbo_住院药房科室对照方案, mrsPar, , , "住院缺省中药房")
            Case "住院缺省发料部门"
                '住院缺省发料部门
                Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
                Call SetParRelation(vsfDrugStore, cbo_住院药房科室对照方案, mrsPar, , , "住院缺省发料部门")
            Case "住院可用西药房"
                '住院可用西药房
                Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
                Call SetParRelation(vsfDrugStore, cbo_住院药房科室对照方案, mrsPar, , , "住院可用西药房")
            Case "住院可用成药房"
                '住院可用成药房
                Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
                Call SetParRelation(vsfDrugStore, cbo_住院药房科室对照方案, mrsPar, , , "住院可用成药房")
            Case "住院可用中药房"
                '住院可用中药房
                Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
                Call SetParRelation(vsfDrugStore, cbo_住院药房科室对照方案, mrsPar, , , "住院可用中药房")
            Case "住院可用发料部门"
                '住院可用发料部门
                Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
                Call SetParRelation(vsfDrugStore, cbo_住院药房科室对照方案, mrsPar, , , "住院可用发料部门")
        End Select
        
        rsTmp.MoveNext
    Loop
    
    
    rsTmp.Filter = "模块=" & p住院医嘱发送
    Do Until rsTmp.EOF
        strValue = "" & rsTmp!参数值
        Select Case rsTmp!参数名
        Case "本科执行自动完成"
            chk(chk_住院本科自动执行长嘱).value = Mid(strValue, 1, 1)
            chk(chk_住院本科自动执行临嘱).value = Mid(strValue, 2, 1)
            
            Call SetParRelation(chk, chk_住院本科自动执行长嘱, mrsPar, rsTmp!参数名, p住院医嘱发送)
            Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
            Call SetParRelation(chk, chk_住院本科自动执行临嘱, mrsPar)
        
        Case "长嘱口服药发送结束时间"
            If InStr(strValue, "|") = 0 Then
                chk(chk_长嘱口服药发送结束时间).value = 0
            Else
                chk(chk_长嘱口服药发送结束时间).value = Val(Split(strValue, "|")(0))
                If chk(chk_长嘱口服药发送结束时间).value = 1 Then
                    dtp口服结束时间.value = Format(Split(strValue, "|")(1), "HH:MM:SS")
                    dtp口服结束时间.Enabled = True
                End If
            End If
            
            Call SetParRelation(chk, chk_长嘱口服药发送结束时间, mrsPar, rsTmp!参数名, p住院医嘱发送)
            Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
            Call SetParRelation(dtp口服结束时间, 0, mrsPar)
        
        Case "住院本科执行自动完成方案"
                '住院本科执行自动完成方案
                For i = 0 To UBound(Split(strValue, ";"))
                    cbo(cbo_住院本科执行自动完成方案).AddItem "方案" & i + 1
                    cbo(cbo_住院本科执行自动完成方案).ItemData(cbo(cbo_住院本科执行自动完成方案).NewIndex) = i + 1
                    mrs住院执行对照.AddNew
                    mrs住院执行对照!科室IDs = Split(strValue, ";")(i)
                    mrs住院执行对照!方案 = i + 1
                    mrs住院执行对照!医嘱类别 = zlDatabase.GetPara("本科执行自动完成医嘱类别", glngSys, p住院医嘱发送, , , , , Val(Split(mrs住院执行对照!科室IDs, ",")(0)))
                    mrs住院执行对照.Update
                Next
                If cbo(cbo_住院本科执行自动完成方案).ListCount > 0 Then
                    cbo(cbo_住院本科执行自动完成方案).ListIndex = 0
                End If
                Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
                Call SetParRelation(vsUseDept, cbo_住院本科执行自动完成方案, mrsPar)
            Case "本科执行自动完成医嘱类别"
                '本科执行自动完成医嘱类别
                Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
                Call SetParRelation(lst, lst_本科执行自动完成医嘱类别, mrsPar, , , "本科执行自动完成医嘱类别")
                
        End Select
        
        rsTmp.MoveNext
    Loop
    
    rsTmp.Filter = "模块=" & p临床路径应用
    Do Until rsTmp.EOF
        strValue = "" & rsTmp!参数值
        Select Case rsTmp!参数名
        Case "路径执行环节启用场合"
            chk(chk_路径执行环节医生启用).value = Mid(strValue, 1, 1)
            chk(chk_路径执行环节护士启用).value = Mid(strValue, 2, 1)
            
            Call SetParRelation(chk, chk_路径执行环节医生启用, mrsPar, rsTmp!参数名, p临床路径应用)
            Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
            Call SetParRelation(chk, chk_路径执行环节护士启用, mrsPar)
        End Select
        
        rsTmp.MoveNext
    Loop
    
    rsTmp.Filter = "模块=" & p门诊医生站
    Do Until rsTmp.EOF
        strValue = "" & rsTmp!参数值
        Select Case rsTmp!参数名
        Case "病人接诊控制"
            txt(txt_允许提前接诊分钟) = Mid(strValue, 3)
        End Select
        
        rsTmp.MoveNext
    Loop
    
    rsTmp.Filter = "模块=" & p住院病历管理
    Do Until rsTmp.EOF
        strValue = "" & rsTmp!参数值
        Select Case rsTmp!参数名
            Case "共享病历连读预览"
                If strValue = -1 Then
                    optEprRead(0).value = True
                ElseIf strValue = 0 Then
                    optEprRead(1).value = True
                Else
                    optEprRead(2).value = True
                    txt(txt_共享病历连续预览N天).Text = strValue
                End If
                Call SetParRelation(optEprRead, 0, mrsPar, rsTmp!参数名, p住院病历管理)
                Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
                Call SetParRelation(optEprRead, 1, mrsPar)
                Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
                Call SetParRelation(optEprRead, 2, mrsPar)
                Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
                Call SetParRelation(txt, txt_共享病历连续预览N天, mrsPar)
        End Select
        rsTmp.MoveNext
    Loop
    
    '使用图片签名科室
    Call setDepartSign
    
    rsTmp.Filter = "模块=" & p电子病案审查
    Do Until rsTmp.EOF
        strValue = "" & rsTmp!参数值
        Select Case rsTmp!参数名
            Case "档案排序顺序"
                If strValue = "" Then strValue = "5;1;6;2;3;4;8;7;9"
                arrTmp = Split(strValue, ";")
                With vsfEpr(0)
                    '1-住院医嘱;2-住院病历;3-护理病历;4-护理记录;5-首页记录;6-医嘱报告;7-疾病证明;8-知情文件
                    For i = 0 To UBound(arrTmp)
                        Select Case arrTmp(i)
                        Case "1"
                            .TextMatrix(i + 1, 0) = "住院医嘱"
                            .RowData(i + 1) = 1
                        Case "2"
                            .TextMatrix(i + 1, 0) = "住院病历"
                            .RowData(i + 1) = 2
                        Case "3"
                            .TextMatrix(i + 1, 0) = "护理病历"
                            .RowData(i + 1) = 3
                        Case "4"
                            .TextMatrix(i + 1, 0) = "护理记录"
                            .RowData(i + 1) = 4
                        Case "5"
                            .TextMatrix(i + 1, 0) = "首页记录"
                            .RowData(i + 1) = 5
                        Case "6"
                            .TextMatrix(i + 1, 0) = "医嘱报告"
                            .RowData(i + 1) = 6
                        Case "7"
                            .TextMatrix(i + 1, 0) = "疾病证明"
                            .RowData(i + 1) = 7
                        Case "8"
                            .TextMatrix(i + 1, 0) = "知情文件"
                            .RowData(i + 1) = 8
                        Case "9"
                            .TextMatrix(i + 1, 0) = "临床路径"
                            .RowData(i + 1) = 9
                        End Select
                    Next
                End With
                Call SetParRelation(vsfEpr, 0, mrsPar, rsTmp!参数名, p电子病案审查)
            Case "审查科室范围"
                '表格赋值
                gstrSQL = "Select ID,编号,姓名 From 人员表"
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
                
                gstrSQL = "Select a.ID,a.编码,a.名称 From 部门表 a,部门性质说明 b Where a.ID=b.部门id And b.工作性质='临床' and ( TO_CHAR (a.撤档时间, 'yyyy-MM-dd') = '3000-01-01' or a.撤档时间 is null)"
                Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
                
                With vsfEpr(1)
                    .Rows = 2
                    arrTmp = Split(strValue, ";")
                    For i = 0 To UBound(arrTmp)
                        varTmp = Split(arrTmp(i), ",") '科室ID
                        rs.Filter = ""
                        rs.Filter = "ID=" & Val(varTmp(0))
                        If rs.RecordCount > 0 Then '
                            If Val(.TextMatrix(.Rows - 1, 0)) > 0 Then .Rows = .Rows + 1
                            .TextMatrix(.Rows - 1, 1) = rs!姓名
                            .TextMatrix(.Rows - 1, 0) = rs!ID
                            
                            For n = 1 To UBound(varTmp)
                                rsData.Filter = ""
                                rsData.Filter = "ID=" & Val(varTmp(n))
                                If rsData.RecordCount > 0 Then
                                    If .TextMatrix(.Rows - 1, 3) = "" Then
                                        .TextMatrix(.Rows - 1, 3) = rsData!名称
                                        .TextMatrix(.Rows - 1, 2) = rsData!ID
                                    Else
                                        .TextMatrix(.Rows - 1, 3) = .TextMatrix(.Rows - 1, 3) & vbCrLf & rsData!名称
                                        .TextMatrix(.Rows - 1, 2) = .TextMatrix(.Rows - 1, 2) & "," & rsData!ID
                                    End If
                                End If
                            Next
                        End If
                    Next
                    .Rows = .Rows + 1
                    .AutoSize 3, 3
                End With
                
                Call SetParRelation(vsfEpr, 1, mrsPar, rsTmp!参数名, p电子病案审查)
        End Select
        rsTmp.MoveNext
    Loop
    
    rsTmp.Filter = "模块=" & p电子病案评分
    Do Until rsTmp.EOF
        strValue = "" & rsTmp!参数值
        Select Case rsTmp!参数名
            Case "评分科室范围"
                '表格赋值
                gstrSQL = "Select ID,编号,姓名 From 人员表"
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
                
                gstrSQL = "Select a.ID,a.编码,a.名称 From 部门表 a,部门性质说明 b Where a.ID=b.部门id And b.工作性质='临床' and ( TO_CHAR (a.撤档时间, 'yyyy-MM-dd') = '3000-01-01' or a.撤档时间 is null)"
                Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
                
                With vsfEpr(2)
                    .Rows = 2
                    arrTmp = Split(strValue, ";")
                    For i = 0 To UBound(arrTmp)
                        varTmp = Split(arrTmp(i), ",") '科室ID
                        rs.Filter = ""
                        rs.Filter = "ID=" & Val(varTmp(0))
                        If rs.RecordCount > 0 Then '
                            If Val(.TextMatrix(.Rows - 1, 0)) > 0 Then .Rows = .Rows + 1
                            .TextMatrix(.Rows - 1, 1) = rs!姓名
                            .TextMatrix(.Rows - 1, 0) = rs!ID
                            
                            For n = 1 To UBound(varTmp)
                                rsData.Filter = ""
                                rsData.Filter = "ID=" & Val(varTmp(n))
                                If rsData.RecordCount > 0 Then
                                    If .TextMatrix(.Rows - 1, 3) = "" Then
                                        .TextMatrix(.Rows - 1, 3) = rsData!名称
                                        .TextMatrix(.Rows - 1, 2) = rsData!ID
                                    Else
                                        .TextMatrix(.Rows - 1, 3) = .TextMatrix(.Rows - 1, 3) & vbCrLf & rsData!名称
                                        .TextMatrix(.Rows - 1, 2) = .TextMatrix(.Rows - 1, 2) & "," & rsData!ID
                                    End If
                                End If
                            Next
                        End If
                    Next
                    .Rows = .Rows + 1
                    .AutoSize 3, 3
                End With
            
                Call SetParRelation(vsfEpr, 2, mrsPar, rsTmp!参数名, p电子病案评分)
        End Select
        rsTmp.MoveNext
    Loop
    
    '门诊输液管理
    Call SetWaittingMixDept
    rsTmp.Filter = "模块=" & p门诊输液管理
    Do Until rsTmp.EOF
        strValue = "" & rsTmp!参数值
        Select Case rsTmp!参数名
        Case "待配液科室列表"
            Call SetParRelation(vsfWaittingMixDept, 0, mrsPar, rsTmp!参数名, p门诊输液管理)
        End Select
        rsTmp.MoveNext
    Loop
    
    '护理记录项目管理
    chk(chk_允许数据同步).Enabled = chk(chk_对应多份护理文件).value = 1           '96044,陈刘,允许同步为对应多份护理文件的子参数
    rsTmp.Filter = "模块=" & p护理记录管理
    Do Until rsTmp.EOF
        strValue = "" & rsTmp!参数值
        Select Case rsTmp!参数名
        Case "体温单标记" '体温单标记
            arrTmp = Split(strValue, ";")
            If UBound(arrTmp) >= 0 Then cbo(cbo_入院自动标志).ListIndex = Val(arrTmp(0))
            If UBound(arrTmp) >= 1 Then cbo(cbo_入科自动标志).ListIndex = Val(arrTmp(1))
            If UBound(arrTmp) >= 2 Then cbo(cbo_转科自动标志).ListIndex = Val(arrTmp(2))
            If UBound(arrTmp) >= 3 Then cbo(cbo_换床自动标志).ListIndex = Val(arrTmp(3))
            If UBound(arrTmp) >= 4 Then cbo(cbo_手术自动标志).ListIndex = Val(arrTmp(4))
            If UBound(arrTmp) >= 5 Then cbo(cbo_出院自动标志).ListIndex = Val(arrTmp(5))
            If UBound(arrTmp) >= 6 Then cbo(cbo_分娩自动标志).ListIndex = Val(arrTmp(6))
            If UBound(arrTmp) >= 7 Then cbo(cbo_出生自动标志).ListIndex = Val(arrTmp(7))
            If UBound(arrTmp) >= 8 Then cbo(cbo_回室自动标志).ListIndex = Val(arrTmp(8))
            If UBound(arrTmp) >= 9 Then cbo(cbo_转病区自动标志).ListIndex = Val(arrTmp(9))
            Call SetParRelation(cbo, cbo_入院自动标志, mrsPar, rsTmp!参数名, p护理记录管理)
            Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
            Call SetParRelation(cbo, cbo_入科自动标志, mrsPar)
            Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
            Call SetParRelation(cbo, cbo_转科自动标志, mrsPar)
            Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
            Call SetParRelation(cbo, cbo_换床自动标志, mrsPar)
            Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
            Call SetParRelation(cbo, cbo_手术自动标志, mrsPar)
            Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
            Call SetParRelation(cbo, cbo_出院自动标志, mrsPar)
            Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
            Call SetParRelation(cbo, cbo_分娩自动标志, mrsPar)
            Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
            Call SetParRelation(cbo, cbo_出生自动标志, mrsPar)
            Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
            Call SetParRelation(cbo, cbo_回室自动标志, mrsPar)
            Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
            Call SetParRelation(cbo, cbo_转病区自动标志, mrsPar)
        Case "体温时间夜班标志" '体温时间夜班标志
            arrTmp = Split(strValue, ";")
            If UBound(arrTmp) >= 1 Then
                ud(ud_体温夜班开始时点).value = Abs(Val(arrTmp(0)))
                ud(ud_体温夜班结束时点).value = Abs(Val(arrTmp(1)))
            Else
                ud(ud_体温夜班开始时点).value = Abs(Val(strValue))
            End If
            Call SetParRelation(txtUD, txtud_体温夜班开始时点, mrsPar, rsTmp!参数名, p护理记录管理)
            Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
            Call SetParRelation(txtUD, txtud_体温夜班结束时点, mrsPar)
        Case "产程生产曲线标志" '曲线标志(顺产)
            arrTmp = Split(strValue, ";")
            For i = 0 To 1
                intIndex = IIF(i = 0, cbo_宫口扩大顺产, cbo_先露下降顺产)
                If UBound(arrTmp) >= i Then
                    lngValue = Val(arrTmp(i))
                    If lngValue < 0 Or lngValue > cbo(intIndex).ListCount - 1 Then lngValue = 0
                    cbo(intIndex).ListIndex = lngValue
                Else
                    cbo(intIndex).ListIndex = 0
                End If
            Next
            
            Call SetParRelation(cbo, cbo_宫口扩大顺产, mrsPar, rsTmp!参数名, p护理记录管理)
            Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
            Call SetParRelation(cbo, cbo_先露下降顺产, mrsPar)
        Case "产程生产措施标志" '措施标志
            arrTmp = Split(strValue, ";")
            For i = 0 To 1
                intIndex = IIF(i = 0, cbo_生产标志内容, cbo_生产标志位置)
                If UBound(arrTmp) >= i Then
                    lngValue = Val(arrTmp(i))
                    If lngValue < 0 Or lngValue > cbo(intIndex).ListCount - 1 Then lngValue = 0
                    cbo(intIndex).ListIndex = lngValue
                Else
                    cbo(intIndex).ListIndex = 0
                End If
            Next
            
            Call SetParRelation(cbo, cbo_生产标志内容, mrsPar, rsTmp!参数名, p护理记录管理)
            Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
            Call SetParRelation(cbo, cbo_生产标志位置, mrsPar)
        Case "产程警戒异常线标志" '警戒线和异常线标志
            arrTmp = Split(strValue, ";")
            For i = 0 To 1
                intIndex = IIF(i = 0, cbo_警戒线显示, cbo_异常线显示)
                If UBound(arrTmp) >= i Then
                    lngValue = Val(arrTmp(i))
                    If lngValue < 0 Or lngValue > cbo(intIndex).ListCount - 1 Then lngValue = 0
                    cbo(intIndex).ListIndex = lngValue
                Else
                    cbo(intIndex).ListIndex = 0
                End If
            Next
            
            Call SetParRelation(cbo, cbo_警戒线显示, mrsPar, rsTmp!参数名, p护理记录管理)
            Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
            Call SetParRelation(cbo, cbo_异常线显示, mrsPar)
        Case "产程生产曲线标志(异)" '曲线标志(异常产)
            arrTmp = Split(strValue, ";")
            For i = 0 To 1
                intIndex = IIF(i = 0, cbo_宫口扩大异常产, cbo_先露下降异常产)
                If UBound(arrTmp) >= i Then
                    lngValue = Val(arrTmp(i))
                    If lngValue < 0 Or lngValue > cbo(intIndex).ListCount - 1 Then lngValue = 0
                    cbo(intIndex).ListIndex = lngValue
                Else
                    cbo(intIndex).ListIndex = 0
                End If
            Next
            
            Call SetParRelation(cbo, cbo_宫口扩大异常产, mrsPar, rsTmp!参数名, p护理记录管理)
            Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
            Call SetParRelation(cbo, cbo_先露下降异常产, mrsPar)
        Case "产程曲线点与0点连线" '0点与首次点连接
            lngValue = Val(strValue)
            If lngValue < 0 Or lngValue > cbo(cbo_零点与首次点连接).ListCount - 1 Then lngValue = 0
            cbo(cbo_零点与首次点连接).ListIndex = lngValue
            Call SetParRelation(cbo, cbo_零点与首次点连接, mrsPar, rsTmp!参数名, p护理记录管理)
        Case "小结标识颜色" '小结标识颜色
            lngValue = Val(strValue)
            picLineColor(0).BackColor = lngValue
            Call SetParRelation(picLineColor, 0, mrsPar, rsTmp!参数名, p护理记录管理)
        Case "体温单标记显示颜色" '体温标志颜色
            lngValue = Val(strValue)
             picLineColor(1).BackColor = lngValue
            Call SetParRelation(picLineColor, 1, mrsPar, rsTmp!参数名, p护理记录管理)
        Case "未记说明显示颜色" '未记说明颜色
            lngValue = Val(strValue)
             picLineColor(2).BackColor = lngValue
            Call SetParRelation(picLineColor, 2, mrsPar, rsTmp!参数名, p护理记录管理)
        Case "手术天数显示颜色" '手术天数颜色
            lngValue = Val(strValue)
            picLineColor(3).BackColor = lngValue
            Call SetParRelation(picLineColor, 3, mrsPar, rsTmp!参数名, p护理记录管理)
        Case "体温复试合格颜色" '体温复试颜色
            lngValue = Val(strValue)
            picLineColor(4).BackColor = lngValue
            Call SetParRelation(picLineColor, 4, mrsPar, rsTmp!参数名, p护理记录管理)
        Case "小结缺省格式"
            arrTmp = Split(strValue, ";")
            If UBound(arrTmp) >= 0 Then cbo(cbo_小结缺省标识).ListIndex = Val(arrTmp(0))
            If UBound(arrTmp) >= 1 Then chk(chk_小结缺省标识).value = Val(arrTmp(1))
            Call SetParRelation(cbo, cbo_小结缺省标识, mrsPar, rsTmp!参数名, p护理记录管理)
            Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
            Call SetParRelation(chk, chk_小结缺省标识, mrsPar)
            Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
        End Select
        rsTmp.MoveNext
    Loop
    '新版住院护士站
    txt(txt_整体护理IP地址) = ""
    txt(txt_整体护理IP端口) = ""
    rsTmp.Filter = "模块=" & p新版住院护士站
    Do Until rsTmp.EOF
        strValue = "" & rsTmp!参数值
        Select Case rsTmp!参数名
            Case "整体护理IP地址"
                If InStr(1, strValue, ":") <> 0 Then
                    txt(txt_整体护理IP地址) = Mid(strValue, 1, InStr(1, strValue, ":") - 1)
                    txt(txt_整体护理IP端口) = Mid(strValue, InStr(1, strValue, ":") + 1)
                Else
                    txt(txt_整体护理IP地址) = strValue
                End If
                Call SetParRelation(txt, txt_整体护理IP地址, mrsPar, rsTmp!参数名, p新版住院护士站)
                Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
                Call SetParRelation(txt, txt_整体护理IP端口, mrsPar)
        End Select
        rsTmp.MoveNext
    Loop
End Sub

Private Sub InitEnv()
'功能：初始化界面控件，加载基础数据
    Dim rsTmp As ADODB.Recordset
    Dim i As Integer, strTmp As String
    Dim blnTmp As Boolean
    
    On Error GoTo ErrHandle

    vsUnWriteDept.ComboList = "..."
    vsUnWriteDept.RowHeightMin = 280
    vsStopDept.ComboList = "..."
    vsStopDept.RowHeightMin = 280
    vsUseDept(cbo_门诊药房科室对照方案).ColWidth(0) = 3000
    vsUseDept(cbo_住院药房科室对照方案).ColWidth(0) = 3000
    vsUseDept(cbo_住院本科执行自动完成方案).ColWidth(0) = 3000
    vsUseDept(cbo_门诊药房科室对照方案).ColAlignment(0) = flexAlignLeftCenter
    vsUseDept(cbo_住院药房科室对照方案).ColAlignment(0) = flexAlignLeftCenter
    vsUseDept(cbo_住院本科执行自动完成方案).ColAlignment(0) = flexAlignLeftCenter
    
    cbo(cbo_中药配方).AddItem "0-三味中药"
    cbo(cbo_中药配方).AddItem "1-四味中药"
    cbo(cbo_中药配方).ListIndex = 0
    
    cbo(cbo_合理用药接口).AddItem "0-未使用"
    cbo(cbo_合理用药接口).AddItem "1-四川美康"
    cbo(cbo_合理用药接口).AddItem "2-上海大通"
    cbo(cbo_合理用药接口).AddItem "3-北京太元通"
    cbo(cbo_合理用药接口).AddItem "4-广州保进"
    cbo(cbo_合理用药接口).AddItem "5-杭州逸曜"
    cbo(cbo_合理用药接口).AddItem "6-中联信息"
    
    cbo(cbo_合理用药接口).ListIndex = 0
    
    cbo(cmd_过敏输入来源).AddItem "0-可选择输入来源"
    cbo(cmd_过敏输入来源).AddItem "1-按药品目录输入"
    cbo(cmd_过敏输入来源).AddItem "2-按过敏源输入"
    cbo(cmd_过敏输入来源).ListIndex = 0
    
    cbo(cbo_住院医生站列表显示).AddItem "科室"
    cbo(cbo_住院医生站列表显示).AddItem "病区"
    cbo(cbo_住院医生站列表显示).ListIndex = 0
    
    cbo(cbo_住院首页标准).AddItem "0-卫生部标准"
    cbo(cbo_住院首页标准).AddItem "1-四川省标准"
    cbo(cbo_住院首页标准).AddItem "2-云南省标准"
    cbo(cbo_住院首页标准).AddItem "3-湖南省标准"
    cbo(cbo_住院首页标准).ListIndex = 0
    
    '读取医嘱内容定义
    gstrSQL = "Select 诊疗类别,医嘱内容 From 医嘱内容定义 Order by 诊疗类别"
    Call zlDatabase.OpenRecordset(mrsAdvice, gstrSQL, Me.Caption)
    
    
    '读取医嘱发送为划价类别
    gstrSQL = "Select 编码,名称 From 诊疗项目类别 Where 编码 Not IN('5','6','7','8','9')" & _
        " Union All Select '5','药品' From Dual Order by 编码"
    Set rsTmp = New ADODB.Recordset
    Call zlDatabase.OpenRecordset(rsTmp, gstrSQL, Me.Caption)
  
    Do While Not rsTmp.EOF
        lst(lst_门诊发送划价类别).AddItem rsTmp!编码 & "-" & rsTmp!名称
        lst(lst_门诊发送划价类别).ItemData(lst(lst_门诊发送划价类别).NewIndex) = Asc(rsTmp!编码)
        
        lst(lst_住院发送划价类别).AddItem rsTmp!编码 & "-" & rsTmp!名称
        lst(lst_住院发送划价类别).ItemData(lst(lst_住院发送划价类别).NewIndex) = Asc(rsTmp!编码)
        
        rsTmp.MoveNext
    Loop

    If rsTmp.RecordCount > 0 Then rsTmp.Filter = "编码<>'4'"
    Do While Not rsTmp.EOF
        lst(lst_门诊发送检查诊断).AddItem rsTmp!编码 & "-" & rsTmp!名称
        lst(lst_门诊发送检查诊断).ItemData(lst(lst_门诊发送检查诊断).NewIndex) = Asc(rsTmp!编码)
        
        
        lst(lst_住院检查入院诊断).AddItem rsTmp!编码 & "-" & rsTmp!名称
        lst(lst_住院检查入院诊断).ItemData(lst(lst_住院检查入院诊断).NewIndex) = Asc(rsTmp!编码)
        rsTmp.MoveNext
    Loop
    lst(lst_门诊发送检查诊断).ListIndex = 0
    lst(lst_住院检查入院诊断).ListIndex = 0
    
    With lst(lst_门诊发送一张单据类别)
        If rsTmp.RecordCount > 0 Then rsTmp.Filter = "编码<>'4' And 编码<>'5'"
        Do While Not rsTmp.EOF
            .AddItem rsTmp!编码 & "-" & rsTmp!名称
            .ItemData(.NewIndex) = Asc(rsTmp!编码)
            rsTmp.MoveNext
        Loop
        .ListIndex = 0
    End With
    
    With lst(lst_本科执行自动完成医嘱类别)
        .Clear
        .AddItem "输液"
        .ItemData(.NewIndex) = 0
        .AddItem "注射"
        .ItemData(.NewIndex) = 1
        .AddItem "口服"
        .ItemData(.NewIndex) = 2
        .AddItem "采集"
        .ItemData(.NewIndex) = 3
        .AddItem "过敏试验"
        .ItemData(.NewIndex) = 4
        .AddItem "普通治疗"
        .ItemData(.NewIndex) = 5
        .AddItem "特殊治疗"
        .ItemData(.NewIndex) = 6
        .AddItem "其它给药途径"
        .ItemData(.NewIndex) = 7
        .AddItem "其它医嘱"
        .ItemData(.NewIndex) = 8
        .ListIndex = 0
    End With
     
    '不收回的发药类型
    gstrSQL = "Select 编码, 名称 From 发药类型 Order by 编码"
    Set rsTmp = New ADODB.Recordset
    Call zlDatabase.OpenRecordset(rsTmp, gstrSQL, Me.Caption)
    
    With lst(lst_不收回的发药类型)
                .Clear
        Do While Not rsTmp.EOF
            .AddItem rsTmp!名称
            rsTmp.MoveNext
        Loop
        If .ListIndex = -1 And .ListCount > 0 Then .ListIndex = 0
    End With
    
    Call InitRs药房对照(mrs门诊药房对照)
    Call InitRs药房对照(mrs住院药房对照)
    Set mrs住院执行对照 = New ADODB.Recordset
    mrs住院执行对照.Fields.Append "方案", adVarChar, 1000
    mrs住院执行对照.Fields.Append "科室IDs", adVarChar, 40000
    mrs住院执行对照.Fields.Append "医嘱类别", adVarChar, 40000
    mrs住院执行对照.CursorLocation = adUseClient
    mrs住院执行对照.LockType = adLockOptimistic
    mrs住院执行对照.CursorType = adOpenStatic
    mrs住院执行对照.Open
    
    '发药药房
    InitVsDrugStore cbo_门诊药房科室对照方案
    InitVsDrugStore cbo_住院药房科室对照方案
    
    '药品流程登记给药途径
    gstrSQL = "Select ID,编码,名称 From 诊疗项目目录" & _
        " Where 类别='E' And 操作类型='2' And 服务对象 IN(2,3) And (站点='" & gstrNodeNo & "' Or 站点 is Null) Order by 编码"
    Call zlDatabase.OpenRecordset(rsTmp, gstrSQL, Me.Caption)
    With lst(lst_药品留存给药途径)
        .Clear
        Do While Not rsTmp.EOF
            .AddItem rsTmp!名称
            .ItemData(.NewIndex) = rsTmp!ID
            rsTmp.MoveNext
        Loop
        If .ListIndex = -1 And .ListCount > 0 Then .ListIndex = 0
    End With
    
    With vsfMecItem
        .Rows = 1
        .Cols = 3
        .TextMatrix(0, 0) = "编码"
        .TextMatrix(0, 1) = "名称"
        .TextMatrix(0, 2) = "内容"
        .ColWidth(0) = 1000
        .ColWidth(1) = 2000
        .ColWidth(2) = 2000
        .Cell(flexcpAlignment, 0, 0, 0, 2) = 4
    End With
    gstrSQL = "select 编码,名称,内容 from 病案项目 order by 编码"
    Call zlDatabase.OpenRecordset(rsTmp, gstrSQL, Me.Caption)
    cmdModify.Enabled = rsTmp.RecordCount > 0: cmdDelete.Enabled = rsTmp.RecordCount > 0
    While Not rsTmp.EOF
        With vsfMecItem
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = rsTmp!编码 & ""
            .TextMatrix(.Rows - 1, 1) = rsTmp!名称 & ""
            .TextMatrix(.Rows - 1, 2) = rsTmp!内容 & ""
        End With
        rsTmp.MoveNext
    Wend
    
    With cbo(2)
        .Clear
        .AddItem "不显示"
        .AddItem Format(Now(), "yyyy-MM-dd hh:mm")
        .AddItem Format(Now(), "yyyy年MM月dd日 hh:mm")
    End With
    
    With vsfEpr(0)
        .Clear
        .Rows = 10
        .RowHidden(0) = True
    End With
    
    With vsfEpr(1)
        .Clear
        .Cols = 4
        .Rows = 1
        .TextMatrix(0, 0) = "姓名ID"
        .TextMatrix(0, 1) = "姓名"
        .TextMatrix(0, 2) = "科室ID"
        .TextMatrix(0, 3) = "科室"
        .ColWidth(0) = 0
        .ColWidth(1) = 1600
        .ColWidth(2) = 0
        .ColWidth(3) = 2600
        .ColAlignment(0) = flexAlignLeftCenter
        .ColAlignment(1) = flexAlignLeftCenter
        .Editable = flexEDKbd
    End With
    
    With vsfEpr(2)
        .Clear
        .Cols = 4
        .Rows = 1
        .TextMatrix(0, 0) = "姓名ID"
        .TextMatrix(0, 1) = "姓名"
        .TextMatrix(0, 2) = "科室ID"
        .TextMatrix(0, 3) = "科室"
        .ColWidth(0) = 0
        .ColWidth(1) = 1600
        .ColWidth(2) = 0
        .ColWidth(3) = 2600
        .ColAlignment(0) = flexAlignLeftCenter
        .ColAlignment(1) = flexAlignLeftCenter
        .Editable = flexEDKbd
    End With
    
    '判断是否能够启用新版血库系统
    If mblnUseBlood Then
        '读取诊疗检验类型
        gstrSQL = "Select 编码,名称,缺省标志 From 诊疗检验类型 order by 编码"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "诊疗检验类型")
        cbo(cbo_输血采集默认诊疗类型).Clear
        i = 0
        Do While Not rsTmp.EOF
            cbo(cbo_输血采集默认诊疗类型).AddItem rsTmp!编码 & "-" & rsTmp!名称
            If Val("" & rsTmp!缺省标志) = 1 Then
                cbo(cbo_输血采集默认诊疗类型).ListIndex = cbo(cbo_输血采集默认诊疗类型).NewIndex
            End If
            If "" & rsTmp!名称 = "血常规" Then i = cbo(cbo_输血采集默认诊疗类型).NewIndex
            rsTmp.MoveNext
        Loop
        If cbo(cbo_输血采集默认诊疗类型).ListIndex = -1 And cbo(cbo_输血采集默认诊疗类型).ListCount > 0 Then
            cbo(cbo_输血采集默认诊疗类型).ListIndex = i
        End If
        cbo(cbo_输血采集默认诊疗类型).Tag = cbo(cbo_输血采集默认诊疗类型).ListIndex
        
        FrmBloodManager.Visible = True
    Else
        FrmBloodManager.Visible = False
        chk(chk_启用血库管理系统).Left = chk(chk_多科会诊意见书写要求).Left
    End If
    
    Call InitNurseItem
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub InitVsDrugStore(ByVal intIndex As Long)
    Dim strSQL As String, rsTmp As Recordset
    Dim j As Long, i As Long
    Dim lngRow As Long, arrTmp As Variant
    
    '药房与发料部门
    strSQL = _
        "Select Distinct A.ID,A.编码,A.名称,B.工作性质 " & _
        " From 部门表 A,部门性质说明 B " & _
        " Where (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
        " AND B.部门ID=A.ID And B.服务对象 IN(" & IIF(intIndex = cbo_门诊药房科室对照方案, "1,3", "2,3") & ") and B.工作性质 in('中药房','西药房','成药房','发料部门')" & _
        " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
        " Order by 工作性质,编码"
    Set rsTmp = New ADODB.Recordset
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    
    With vsfDrugStore(intIndex)
        .Rows = .FixedRows
        .Editable = flexEDKbdMouse
        .MergeCol(.ColIndex("类别")) = True
        .MergeCells = flexMergeFixedOnly
        
        If intIndex = cbo_门诊药房科室对照方案 Then
            '隐藏 发药窗口  列，1252 模块这3个参数没有用到 '成药房窗口','西药房窗口','中药房窗口'  故先隐藏了
            .ColHidden(.ColIndex("发药窗口")) = True
        End If
        
        If Not rsTmp.EOF Then
            .Rows = .FixedRows + rsTmp.RecordCount
            lngRow = .FixedRows
            arrTmp = Split("西药房,成药房,中药房,发料部门", ",")
            For i = 0 To UBound(arrTmp)
                rsTmp.Filter = "工作性质='" & arrTmp(i) & "'"
                Do While Not rsTmp.EOF
                    .TextMatrix(lngRow, .ColIndex("类别")) = arrTmp(i)
                    .TextMatrix(lngRow, .ColIndex("药房")) = rsTmp!名称
                    .RowData(lngRow) = Val(rsTmp!ID)
                    
                    lngRow = lngRow + 1
                    rsTmp.MoveNext
                Loop
                If lngRow < .Rows - 1 Then  '划分隔线
                    .Select lngRow, .FixedCols, lngRow, .Cols - 1
                    .CellBorder vbBlue, 0, 1, 0, 0, 0, 0
                End If
            Next
        End If
    End With
End Sub

Private Sub cmdOK_Click()
    Dim lngTmp As Long
    
    If ValidateData() = False Then Exit Sub
    
    Call Save医嘱内容
    Call Save药房科室对照
    Call SaveDepartSign
    Call Save住院执行对照
    lngTmp = -1
    mrsPar.Filter = "参数名='科室药房对照按本机参数设置' and 修改状态=1"
    If Not mrsPar.EOF Then
        lngTmp = mrsPar!参数新值
    End If
    
    If SavePar(mrsPar, Me) = False Then Exit Sub
    
    Call Set药房科室对照参数性质(lngTmp)
    
    mblnOk = True
    Unload Me
End Sub

Private Function ValidateData() As Boolean
'功能：验证数据的有效性
    mrsPar.Filter = "模块=" & p新版住院护士站 & " and 参数名='整体护理IP地址' and 修改状态=1"
    If cmdLink.Tag <> "OK" And Not mrsPar.EOF Then
        If CheckNurseIntegrateIP(False) = False Then
            If MsgBox("临床工作站页面中的整体护理服务器IP地址设置不正确，如果医院使用了移动整体护理，将导致新版护士工作站无法使用整体护理功能。" & vbCrLf & "请问您是否要继续？", vbYesNo + vbDefaultButton2 + vbQuestion, gstrSysName) = vbNo Then
                Exit Function
            End If
        End If
    End If
    ValidateData = True
End Function

Private Sub cmdCancel_Click()
    mblnOk = False
    Unload Me
End Sub

Private Sub lst_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(lst, Index, mrsPar)
End Sub

Private Sub chk_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(chk, Index, mrsPar)
End Sub

Private Sub optPASSVer_Click(Index As Integer)
    Dim strValue As String
    
    If Me.Visible Then
        strValue = IIF(optPASSVer(0).value, "3.0", "4.0")
        Call SetParChange(optPASSVer, Index, mrsPar, True, strValue)
        cmdSet.Visible = mobjPass.SetEnabled(cbo(cbo_合理用药接口).ListIndex, strValue)
    End If
End Sub

Private Sub optPASSVer_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optPASSVer_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optPASSVer, Index, mrsPar)
End Sub

Private Sub opt抗菌目的门诊_Click(Index As Integer)
    If Me.Visible Then Call SetParChange(opt抗菌目的门诊, Index, mrsPar)
 
End Sub

Private Sub opt抗菌目的门诊_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub opt抗菌目的门诊_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(opt抗菌目的门诊, Index, mrsPar)
End Sub

Private Sub opt抗菌目的住院_Click(Index As Integer)
    If Me.Visible Then Call SetParChange(opt抗菌目的住院, Index, mrsPar)
End Sub

Private Sub opt抗菌目的住院_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub opt抗菌目的住院_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(opt抗菌目的住院, Index, mrsPar)
End Sub

Private Sub txt_Change(Index As Integer)
    Dim strValue As String
    Dim strIP As String, strIPProt As String
    If Me.Visible Then
        Select Case Index
            Case txt_允许提前接诊分钟
                Call SetParChange(opt接诊控制, 1, mrsPar, True, IIF(opt接诊控制(0).value, "0|0", IIF(opt接诊控制(1).value, "1|" & NVL(txt(txt_允许提前接诊分钟).Text, "0"), "2|" & NVL(txt(txt_允许提前接诊分钟).Text, "0"))))
            Case txt_共享病历连续预览N天
                Call SetParChange(optEprRead, 2, mrsPar, True, Val(txt(txt_共享病历连续预览N天)))
                Call SetParChange(txt, Index, mrsPar, True, Val(txt(txt_共享病历连续预览N天)))
            Case txt_整体护理IP地址, txt_整体护理IP端口
                If txt(txt_整体护理IP地址).Text = "" Then
                    strValue = ""
                ElseIf txt(txt_整体护理IP端口).Text = "" Then
                    strValue = txt(txt_整体护理IP地址).Text
                Else
                    strValue = txt(txt_整体护理IP地址).Text & ":" & txt(txt_整体护理IP端口).Text
                End If
                Call SetParChange(txt, txt_整体护理IP地址, mrsPar, True, strValue) '更新参数值
                mrsPar.Filter = "模块=" & p新版住院护士站 & " and 参数名='整体护理IP地址' "
                If Not mrsPar.EOF Then
                    If InStr(1, mrsPar!参数值 & "", ":") <> 0 Then
                        strIP = Mid(mrsPar!参数值, 1, InStr(1, mrsPar!参数值, ":") - 1)
                        strIPProt = Mid(mrsPar!参数值, InStr(1, mrsPar!参数值, ":") + 1)
                    Else
                        strIP = mrsPar!参数值 & ""
                        strIPProt = ""
                    End If
                    txt(txt_整体护理IP地址).ForeColor = IIF(txt(txt_整体护理IP地址).Text <> strIP, &HC0&, &H0&)
                    txt(txt_整体护理IP端口).ForeColor = IIF(txt(txt_整体护理IP端口).Text <> strIPProt, &HC0&, &H0&)
                End If
                cmdLink.Tag = ""
            Case Else
                Call SetParChange(txt, Index, mrsPar)
        End Select
    End If
End Sub

Private Sub txt_GotFocus(Index As Integer)
    Select Case Index
    Case txt_门诊输液自动刷新病人, txt_门诊输液滴速, txt_门诊输液滴系数, txt_门诊输液输液提醒, txt_门诊输液皮试提醒, _
            txt_手术标注天数, txt_体温复试合格符号, txt_整体护理IP地址, txt_整体护理IP端口
        Call zlControl.TxtSelAll(txt(Index))
    End Select
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    ElseIf KeyAscii = Asc(gstrParSplit1) Or KeyAscii = Asc(gstrParSplit2) Then
        KeyAscii = 0
    End If
    
    Select Case Index
        Case txt_签名使用图片高度, txt_共享病历N行自动折叠, txt_共享病历连续预览N天, txt_电子病案审查缺省期限, _
            txt_电子病案审查待复查刷新间隔, txt_电子病案借阅缺省期限, txt_电子病案借阅最大期限, _
            txt_门诊输液自动刷新病人, txt_门诊输液滴速, txt_门诊输液滴系数, txt_门诊输液输液提醒, txt_门诊输液皮试提醒, txt_手术标注天数, txt_整体护理IP端口
            If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then
                KeyAscii = 0
            End If
        Case txt_整体护理IP地址
            If InStr("0123456789." & Chr(8), Chr(KeyAscii)) = 0 Then
                KeyAscii = 0
            End If
    End Select
End Sub

Private Sub txt_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index = txt_允许提前接诊分钟 Then
        Call SetParTip(opt接诊控制, 1, mrsPar)
    Else
        Call SetParTip(txt, Index, mrsPar)
    End If
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    Select Case Index
    Case txt_门诊输液自动刷新病人
        If Val(txt(Index).Text) <> 0 And Val(txt(Index).Text) < 30 Then
            MsgBox "门诊输液管理的“自动刷新病人清单”录入不正确，范围（0或大于等于30）！", vbInformation, gstrSysName
            txt(Index).Text = 30
        End If
    Case txt_门诊输液滴速
        If Val(txt(Index).Text) < 10 Or Val(txt(Index).Text) > 100 Then
            MsgBox "门诊输液管理的“默认滴速”录入不正确，范围（10-100）！", vbInformation, gstrSysName
            txt(Index).Text = 40
        End If
    Case txt_门诊输液滴系数
        If InStr(",10,15,20,", "," & Trim(txt(Index).Text) & ",") <= 0 Then
            MsgBox "门诊输液管理的“默认滴系数”录入不正确，范围（10、15、20）！", vbInformation, gstrSysName
            txt(Index).Text = 20
        End If
    Case txt_门诊输液输液提醒
        If Val(txt(Index).Text) < 0 Or Val(txt(Index).Text) > 60 Then
            MsgBox "门诊输液管理的“输液提醒”录入不正确，范围（0-60）！", vbInformation, gstrSysName
            txt(Index).Text = 3
        End If
    Case txt_门诊输液皮试提醒
        If Val(txt(Index).Text) < 0 Or Val(txt(Index).Text) > 60 Then
            MsgBox "门诊输液管理的“皮试提醒”录入不正确，范围（0-60）！", vbInformation, gstrSysName
            txt(Index).Text = 0
        End If
    Case txt_整体护理IP地址
        If txt(Index).Text = "" Then Exit Sub
        If IsIPAddress(txt(Index).Text) = False Then
            MsgBox "整体护理服务器IP地址录入不正确，请检查！", vbInformation, gstrSysName
            Cancel = True
        End If
    End Select
End Sub

Private Sub txtUD_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(txtUD, Index, mrsPar)
End Sub

Private Sub cbo_GotFocus(Index As Integer)
    Call SetParTip(cbo, Index, mrsPar)
End Sub



Private Function SetDeptInput(vsTmp As VSFlexGrid, ByVal lngRow As Long, ByVal lngCol As Long, ByVal rsTmp As ADODB.Recordset) As Boolean
    '先检查下表格中是否存在
    Dim strTmp As String
    With vsTmp
        On Error Resume Next
        If vsTmp.Name = "vsStopDept" Then
            strTmp = mcolStop科室("_" & rsTmp!ID)
        Else
            strTmp = mcol科室("_" & rsTmp!ID)
        End If
        If Err.Number = 0 Then
            MsgBox "该科室已经存在，请重新输入。", vbInformation, gstrSysName
            .TextMatrix(lngRow, lngCol) = CStr(.Cell(flexcpData, lngRow, lngCol))
            Exit Function
        Else
            Err.Clear
        End If
        On Error GoTo 0
        
        If .TextMatrix(lngRow, lngCol + 4) <> "" Then
            If vsTmp.Name = "vsStopDept" Then
                Call mcolStop科室.Remove("_" & Val(.TextMatrix(.Row, .Col + 4)))
            Else
                Call mcol科室.Remove("_" & Val(.TextMatrix(.Row, .Col + 4)))
            End If
        End If
        
        .TextMatrix(lngRow, lngCol) = rsTmp!名称 & ""
        .Cell(flexcpData, lngRow, lngCol) = rsTmp!名称 & ""
        .TextMatrix(lngRow, lngCol + 4) = rsTmp!ID & ""
        If vsTmp.Name = "vsStopDept" Then
            Call mcolStop科室.Add(rsTmp!ID & "", "_" & rsTmp!ID)
        Else
            Call mcol科室.Add(rsTmp!ID & "", "_" & rsTmp!ID)
        End If
        SetDeptInput = True
    End With
End Function



Private Sub txtLocate_Change(Index As Integer)
    If Index = txt_Dept Then
        mlngPreFind = 1
    ElseIf Index = txt_Par Then
        txtLocate(Index).Tag = ""
    End If
End Sub

Private Sub txtLocate_GotFocus(Index As Integer)
    txtLocate(Index).SelStart = 0
    txtLocate(Index).SelLength = Len(txtLocate(Index).Text)
End Sub

Private Sub txtLocate_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Dim strFind As String
        
        If Trim(txtLocate(Index).Text) = "" Then Exit Sub
        strFind = UCase(Trim(txtLocate(Index).Text))
        
        Select Case Index
        Case txt_Par
            Call LocatePar(txtLocate(Index), Me)
        Case txt_Dept
            If vsUnWriteDept.Visible Then
                Call LocateDept(strFind, vsUnWriteDept)
            End If
        End Select
    End If
End Sub

Private Sub LocateDept(ByVal strFind As String, ByRef objTmp As Object)
'功能：检查不写超量的科室
    Dim i As Long, j As Long
    Dim lngRows As Long, lngStart As Long
    Dim strCode As String, strName As String
    
    If TypeName(objTmp) = "ListBox" Then 'lst_自动校对病区
        With objTmp
            lngRows = .ListCount - 1
            
            lngStart = IIF(mlngPreFind = 1, 0, mlngPreFind)
            For i = lngStart To .ListCount - 1
                strCode = Split(.List(i), "-")(0)
                strName = Split(.List(i), "-")(1)
                If strCode Like strFind & "*" Or strName Like IIF(gstrLike <> "", "*", "") & strFind & "*" Then
                    .ListIndex = i
                    .SetFocus
                    Exit For
                End If
            Next
        End With
        If i < lngRows Then
            mlngPreFind = i + 1
        Else
            If mlngPreFind = 1 Then
                MsgBox "没有找到匹配的，请检查输入的内容。", vbInformation, Me.Caption
                txtLocate(txt_Dept).SetFocus
            Else
                MsgBox "全部找完了，后面没有了。", vbInformation, Me.Caption
                mlngPreFind = 1
            End If
        End If
    Else
        '考虑到此功能的使用频率低，暂时不支持连续查找
        With objTmp
            For i = 0 To .Rows - 1
                For j = 0 To .Cols - 1
                    If .ColHidden(j) = False Then
                        If .TextMatrix(i, j) Like IIF(gstrLike <> "", "*", "") & strFind & "*" Then
                            .Row = i: .Col = j
                            .ShowCell i, j
                            Exit Sub
                        End If
                    End If
                Next
            Next
            
            MsgBox "没有找到匹配的科室，请检查输入的内容。", vbInformation, Me.Caption
            txtLocate(txt_Dept).SetFocus
        End With
    End If
End Sub

Private Sub vsfDepartSign_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim l As Long
    Dim strTmp As String
    
    If vsfDepartSign.TextMatrix(Row, vsfDepartSign.ColIndex("ID")) = "" Then Exit Sub

    With vsfDepartSign
        For l = 1 To .Rows - 1
            If Val(.Cell(flexcpChecked, l, .ColIndex("启用"))) <> Decode(Val(.RowData(l)), 1, 1, 2) Then
                vsfDepartSign.ForeColor = vbRed
                Exit Sub
            End If
        Next
        vsfDepartSign.ForeColor = vbBlack
    End With
End Sub

Private Sub vsfDepartSign_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsfDepartSign
        If Row = .Rows - 1 Then
            If .TextMatrix(Row, .ColIndex("ID")) = "" And .ColKey(Col) = "启用" Then Cancel = 1
            vsfDepartSign.TextMatrix(Row, Col) = ""
        Else
            If .ColKey(Col) <> "启用" Then Cancel = 1
        End If
    End With
End Sub

Private Sub vsfDepartSign_BeforeScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long, Cancel As Boolean)
    cmdDepartSelect.Visible = False
End Sub

Private Sub vsfDepartSign_DblClick()
    If vsfDepartSign.TextMatrix(vsfDepartSign.RowSel, vsfDepartSign.ColIndex("ID")) = "" Then Exit Sub
    vsfDepartSign.TextMatrix(vsfDepartSign.RowSel, vsfDepartSign.ColIndex("启用")) = IIF(vsfDepartSign.TextMatrix(vsfDepartSign.RowSel, vsfDepartSign.ColIndex("启用")) = "1", "0", "1")
    Call vsfDepartSign_AfterEdit(vsfDepartSign.RowSel, vsfDepartSign.ColSel)
End Sub

Private Sub vsfDepartSign_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then Exit Sub
    
    Dim lngCol As Long, lngRow As Long, X As Long, Y As Long, rsTmp As ADODB.Recordset, blnCancel As Boolean
    Dim vPoint As POINTAPI, i As Long, blnChange As Boolean, strNewValue As String
    
    vPoint = zlControl.GetCoordPos(vsfDepartSign.hwnd, cmdDepartSelect.Left, cmdDepartSelect.Top)
    lngCol = vsfDepartSign.Col: lngRow = vsfDepartSign.Row
    
    gstrSQL = "Select b.Id, b.编码, b.名称,b.简码　" & vbNewLine & _
                    "from (Select 参数id, 部门id, 参数值 From Zldeptparas Where 参数id = (Select ID From zlParameters Where 参数名 = '签名使用图片')) A,部门表 B" & vbNewLine & _
                    "Where b.Id = a.部门id(+) And a.部门id Is Null And (Instr(b.简码, Upper([1])) > 0 Or Instr(b.编码, Upper([1])) > 0)" & vbNewLine & _
                    "And (b.撤档时间 Is Null Or b.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD'))" & vbNewLine & _
                    "Order By b.编码"
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, gstrSQL, 0, "", False, "部门", "请选择部门", False, False, True, vPoint.X, vPoint.Y, 225, blnCancel, True, True, vsfDepartSign.EditText)
        
    If blnCancel Then
        vsfDepartSign.TextMatrix(lngRow, lngCol) = ""
    Else
        If rsTmp Is Nothing Then
            vsfDepartSign.TextMatrix(lngRow, lngCol) = ""
            Exit Sub
        End If
        For i = 1 To vsfDepartSign.Rows - 1
            If vsfDepartSign.TextMatrix(i, vsfDepartSign.ColIndex("ID")) = rsTmp!ID Then
                vsfDepartSign.TextMatrix(lngRow, lngCol) = ""
                Exit Sub
            End If
        Next
        vsfDepartSign.TextMatrix(lngRow, vsfDepartSign.ColIndex("ID")) = rsTmp!ID & ""
        vsfDepartSign.TextMatrix(lngRow, vsfDepartSign.ColIndex("编码")) = rsTmp!编码 & ""
        vsfDepartSign.TextMatrix(lngRow, vsfDepartSign.ColIndex("科室")) = rsTmp!名称 & ""
        vsfDepartSign.TextMatrix(lngRow, vsfDepartSign.ColIndex("启用")) = "1"
        blnChange = True
    End If

    If blnChange Then
        If vsfDepartSign.TextMatrix(vsfDepartSign.Rows - 1, vsfDepartSign.ColIndex("ID")) <> "" Then
            vsfDepartSign.Rows = vsfDepartSign.Rows + 1
        End If

        With vsfDepartSign
            For i = 1 To .Rows - 1
                If Val(.Cell(flexcpChecked, i, .ColIndex("启用"))) <> Decode(Val(.RowData(i)), 1, 1, 2) Then
                    vsfDepartSign.ForeColor = vbRed
                    Exit Sub
                End If
            Next
            vsfDepartSign.ForeColor = vbBlack
        End With
        vsfDepartSign.Row = vsfDepartSign.Rows - 1
        Call vsfDepartSign.ShowCell(vsfDepartSign.Rows - 1, 0)
    End If
End Sub

Private Sub vsfDepartSign_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngRow As Long, lngCol As Long, lngRowHeight As Long, lngTop As Long, lngLeft As Long, i As Integer, blnScroll As Boolean, lngRowsHeight As Long

    lngRow = vsfDepartSign.MouseRow: lngCol = vsfDepartSign.MouseCol
    If lngRow = vsfDepartSign.Rows - 1 And lngCol >= 0 Then
        For i = 1 To vsfDepartSign.Rows - 1
            lngRowsHeight = lngRowsHeight + vsfDepartSign.RowHeight(i) + IIF(vsfDepartSign.RowHeight(i) = vsfDepartSign.RowHeight(0), 30, 0)
        Next
        If lngRowsHeight > vsfDepartSign.ClientHeight - vsfDepartSign.RowHeight(0) - 30 Then blnScroll = True '认为产生滚动，CMD的位置相应向左移动240
            
        For i = vsfDepartSign.TopRow To lngRow - 1
            lngTop = lngTop + vsfDepartSign.RowHeight(i) + IIF(vsfDepartSign.RowHeight(i) = vsfDepartSign.RowHeight(0), 30, 0)
        Next
        lngTop = lngTop + vsfDepartSign.RowHeight(0) + 30
        'lngLeft = IIF(lngCol = 1, 1600, vsfDepartSign.Width - 10 - IIF(blnScroll, 240, 0)) - cmdDepartSelect.Width - 30
        lngLeft = vsfDepartSign.Width - 10 - IIF(blnScroll, 240, 0) - cmdDepartSelect.Width - 30
        lngRowHeight = vsfDepartSign.RowHeight(lngRow)
        Call cmdDepartSelect.Move(lngLeft, lngTop, cmdDepartSelect.Width, lngRowHeight)
        cmdDepartSelect.Visible = True
    Else
        cmdDepartSelect.Visible = False
    End If
        
    Call SetParTip(vsfDepartSign, 0, mrsPar)
End Sub

Private Sub vsfDrugStore_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strObjTag As String
    On Error Resume Next
    strObjTag = GetDrugTag(Index, vsfDrugStore(Index).MouseRow, vsfDrugStore(Index).MouseCol)
    Call SetParTip(vsfDrugStore, Index, mrsPar, strObjTag)
End Sub

Private Function GetDrugTag(ByVal lngIndex As Long, ByVal lngRow As Long, ByVal lngCol As Long) As String
    With vsfDrugStore(lngIndex)
        If lngCol = .ColIndex("可用") Then
            GetDrugTag = IIF(lngIndex = cbo_门诊药房科室对照方案, "门诊", "住院") & "可用" & .TextMatrix(lngRow, .ColIndex("类别"))
        ElseIf lngCol = .ColIndex("缺省") Then
            GetDrugTag = IIF(lngIndex = cbo_门诊药房科室对照方案, "门诊", "住院") & "缺省" & .TextMatrix(lngRow, .ColIndex("类别"))
        ElseIf lngCol = .ColIndex("发药窗口") Then
            GetDrugTag = .TextMatrix(lngRow, .ColIndex("类别")) & "窗口"
        End If
    End With
End Function

Private Sub vsfDrugStore_StartEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = vsfDrugStore(Index).ColIndex("可用") Then
        Call Set可用药房(Index, Row, True)
    ElseIf Col = vsfDrugStore(Index).ColIndex("可用") Then
        Call Set缺省药房(Index)
    End If
    If Col <> vsfDrugStore(Index).ColIndex("发药窗口") Then Cancel = True
End Sub

Private Sub vsfDrugStore_BeforeEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsfDrugStore(Index)
        Select Case Col
        Case .ColIndex("可用")
            Cancel = Val(.Cell(flexcpData, Row, Col)) <> 0
        Case .ColIndex("缺省")
            Cancel = Val(.Cell(flexcpData, Row, Col)) <> 0
        Case .ColIndex("发药窗口")
            Cancel = Val(.Cell(flexcpData, Row, Col)) <> 0
        Case Else
            Cancel = True
            Exit Sub
        End Select
    End With
End Sub

Private Sub vsfDrugStore_DblClick(Index As Integer)
    With vsfDrugStore(Index)
        If .MouseCol = .ColIndex("缺省") Then
            Call Set缺省药房(Index)
        ElseIf .MouseCol = .ColIndex("药房") Then
            Call Set可用药房(Index, .Row, True)
        ElseIf .MouseCol = .ColIndex("可用") And .MouseRow = .FixedRows - 1 Then
            Dim i As Long
            For i = .FixedRows To .Rows - 1
                Call Set可用药房(Index, i)
            Next
        End If
    End With
End Sub

Private Sub vsfDrugStore_EnterCell(Index As Integer)

    With vsfDrugStore(Index)
        If .Row > 0 Then
            .FocusRect = flexFocusLight
        End If
    End With
End Sub

Private Sub vsfDrugStore_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        If vsfDrugStore(Index).Col = vsfDrugStore(Index).ColIndex("缺省") Then
            Call Set缺省药房(Index)
        End If
    End If
End Sub

Private Sub SetRecord药房(ByVal lngIndex As Long, ByVal lngRow As Long, ByVal lngCol As Long)
    Dim strTmp As String, i As Long
    Dim strValue As String
    
    strTmp = Replace(Replace(GetDrugTag(lngIndex, lngRow, lngCol), "门诊", ""), "住院", "")
    
    With vsfDrugStore(lngIndex)
        If strTmp Like "可用*" Then
            For i = .FixedRows To .Rows - 1
                If .TextMatrix(i, .ColIndex("类别")) = .TextMatrix(lngRow, .ColIndex("类别")) Then
                    If .TextMatrix(i, .ColIndex("可用")) <> 0 Then
                        strValue = strValue & "," & .RowData(i)
                    End If
                End If
            Next
            strValue = Mid(strValue, 2)
        ElseIf strTmp Like "缺省*" Then
            strValue = IIF(.TextMatrix(lngRow, .ColIndex("缺省")) = "√", .RowData(lngRow), "")
        ElseIf strTmp Like "*窗口" Then
            For i = .FixedRows To .Rows - 1
                If .TextMatrix(i, .ColIndex("类别")) = .TextMatrix(lngRow, .ColIndex("类别")) Then
                    If .TextMatrix(i, .ColIndex("发药窗口")) <> "自动分配" Then
                        strValue = strValue & "," & .RowData(i) & ":" & .TextMatrix(i, .ColIndex("发药窗口"))
                    End If
                End If
            Next
            strValue = Mid(strValue, 2)
        End If
    End With
    IIF(lngIndex = cbo_门诊药房科室对照方案, mrs门诊药房对照, mrs住院药房对照).Fields(strTmp).value = strValue
    SST.Tag = "已修改"
End Sub

Private Sub Set缺省药房(ByVal Index As Integer)
'功能：设置当前行的缺省药房，同时处理相同类型的其他行的缺省药房
    Dim i As Long
    
    With vsfDrugStore(Index)
        If Val("" & .Cell(flexcpData, .Row, .ColIndex("缺省"))) = 0 Then  '该参数允许修改的情况下
            If .TextMatrix(.Row, .ColIndex("缺省")) = "√" Then
                .TextMatrix(.Row, .ColIndex("缺省")) = ""
            Else
                '当没有有权限修改可用时且可用为0（false)时不允许设置缺省
                If Not (Val(.TextMatrix(.Row, .ColIndex("可用"))) = 0 And Val("" & .Cell(flexcpData, .Row, .ColIndex("可用"))) = 1) Then
                    '同类别的其他行取消缺省
                    For i = .FixedRows To .Rows - 1
                        If .TextMatrix(.Row, .ColIndex("类别")) = .TextMatrix(i, .ColIndex("类别")) Then
                            If .TextMatrix(i, .ColIndex("缺省")) = "√" Then .TextMatrix(i, .ColIndex("缺省")) = ""
                        End If
                    Next
                    .TextMatrix(.Row, .ColIndex("可用")) = -1    '自动设置为可用
                    Call SetRecord药房(Index, .Row, .ColIndex("可用"))
                    .TextMatrix(.Row, .ColIndex("缺省")) = "√"
                Else
                    MsgBox "设置当前药房为缺省时，会同时将当前药房设置为可用，" & vbNewLine & "你没有修改可用药房的权限。", vbInformation, gstrSysName
                End If
            End If
            Call SetRecord药房(Index, .Row, .ColIndex("缺省"))
        Else
            MsgBox "你没有修改缺省药房的权限。", vbInformation, gstrSysName
        End If
    End With
End Sub

Private Sub Set可用药房(ByVal Index As Integer, ByVal lngRow As Long, Optional ByVal blnAsk As Boolean = False)
'功能：设置当前行的可用药房，同时处理当前行的缺省药房

    With vsfDrugStore(Index)
        If Val("" & .Cell(flexcpData, lngRow, .ColIndex("可用"))) = 0 Then   '该参数允许修改的情况下
            If Val(.TextMatrix(lngRow, .ColIndex("可用"))) = -1 Then
                '当前科室勾选可用
                If Not (Val("" & .Cell(flexcpData, lngRow, .ColIndex("缺省"))) = 1 And .TextMatrix(lngRow, .ColIndex("缺省")) = "√") Then
                    .TextMatrix(lngRow, .ColIndex("可用")) = 0
                    .TextMatrix(lngRow, .ColIndex("缺省")) = ""
                    Call SetRecord药房(Index, lngRow, .ColIndex("缺省"))
                    Call SetRecord药房(Index, lngRow, .ColIndex("可用"))
                Else
                    If blnAsk Then
                        MsgBox "取消当前药房可用时，会同时取消当前药房缺省，" & vbNewLine & "你没有修改缺省药房的权限。", vbInformation, gstrSysName
                    End If
                End If
            Else
                .TextMatrix(lngRow, .ColIndex("可用")) = -1    '自动设置为可用
                Call SetRecord药房(Index, lngRow, .ColIndex("可用"))
            End If
        Else
            If blnAsk Then
                MsgBox "你没有修改可用药房的权限。", vbInformation, gstrSysName
            End If
        End If
    End With
End Sub


Private Sub vsfDrugStore_ValidateEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Index = cbo_门诊药房科室对照方案 Then
        If Col = vsfDrugStore(Index).ColIndex("发药窗口") Then
            Call SetRecord药房(Index, Row, Col)
        End If
    End If
End Sub

Private Sub vsfEpr_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim l As Long, strNewValue As String
    If Not (Index = 1 Or Index = 2) Then Exit Sub '0档案排序不需要 1 2是审查、评分指定科室及人员范围
    
    If KeyCode = vbKeyDelete Then
        For l = vsfEpr(Index).Row To vsfEpr(Index).Rows - 1
            If l <> vsfEpr(Index).Rows - 1 Then
                vsfEpr(Index).TextMatrix(l, 0) = vsfEpr(Index).TextMatrix(l + 1, 0)
                vsfEpr(Index).TextMatrix(l, 1) = vsfEpr(Index).TextMatrix(l + 1, 1)
                vsfEpr(Index).TextMatrix(l, 2) = vsfEpr(Index).TextMatrix(l + 1, 2)
                vsfEpr(Index).TextMatrix(l, 3) = vsfEpr(Index).TextMatrix(l + 1, 3)
            Else
                vsfEpr(Index).TextMatrix(l, 0) = ""
                vsfEpr(Index).TextMatrix(l, 1) = ""
                vsfEpr(Index).TextMatrix(l, 2) = ""
                vsfEpr(Index).TextMatrix(l, 3) = ""
            End If
        Next
        vsfEpr(Index).Rows = vsfEpr(Index).Rows - 1
        If vsfEpr(Index).TextMatrix(vsfEpr(Index).Rows - 1, 0) <> "" And vsfEpr(Index).TextMatrix(vsfEpr(Index).Rows - 1, 2) <> "" Then
            vsfEpr(Index).Rows = vsfEpr(Index).Rows + 1
        End If
        vsfEpr(Index).RowHeight(vsfEpr(Index).Rows - 1) = 240
        vsfEpr(Index).AutoSize 3, 3
        vsfEpr(Index).Row = vsfEpr(Index).Rows - 1
        
        For l = 1 To vsfEpr(Index).Rows - 1
            If vsfEpr(Index).TextMatrix(l, 0) <> "" And vsfEpr(Index).TextMatrix(l, 2) <> "" Then
                strNewValue = strNewValue & ";" & vsfEpr(Index).TextMatrix(l, 0) & "," & vsfEpr(Index).TextMatrix(l, 2)
             End If
        Next
        If strNewValue <> "" Then
            strNewValue = Mid(strNewValue, 2)
        End If
        
        Call SetParChange(vsfEpr, Index, mrsPar, True, strNewValue)
    End If
End Sub

Private Sub vsfEpr_KeyPressEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
If Not (Index = 1 Or Index = 2) Then Exit Sub '0档案排序不需要 1 2是审查、评分指定科室及人员范围

If KeyAscii = vbKeyReturn Then
Dim lngCol As Long, lngRow As Long, X As Long, Y As Long, rsTmp As New ADODB.Recordset, blnCancel As Boolean
Dim vPoint As POINTAPI, l As Long, blnChange As Boolean, strNewValue As String, strOldValue As String
    vPoint = zlControl.GetCoordPos(vsfEpr(Index).hwnd, cmdEprSelect(Index).Left, cmdEprSelect(Index).Top)
    lngCol = vsfEpr(Index).Col: lngRow = vsfEpr(Index).Row
    strOldValue = vsfEpr(Index).TextMatrix(lngRow, lngCol)
    If lngCol = 1 Then '姓名
        gstrSQL = "Select Distinct a.编号,a.Id, a.姓名, a.简码, c.名称 As 科室" & vbNewLine & _
                "From 人员表 A, 人员性质说明 B, 部门表 C, 部门人员 D" & vbNewLine & _
                "Where a.Id = b.人员id And c.Id = d.部门id And d.人员id = a.Id And d.缺省 = 1 And" & vbNewLine & _
                "      (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null) And b.人员性质 In ('医生') And" & vbNewLine & _
                "      (Instr(a.姓名,[1])>0 or Instr(a.简码,Upper([1]))>0 or Instr(a.编号,[1])>0 ) " & vbNewLine & _
                "Order By a.编号"
        Set rsTmp = zlDatabase.ShowSQLSelect(Me, gstrSQL, 0, "", False, "姓名", "请选择一个审查人员", False, False, True, vPoint.X, vPoint.Y, 225, blnCancel, True, True, vsfEpr(Index).EditText)
        If blnCancel Then
            vsfEpr(Index).TextMatrix(lngRow, 1) = strOldValue
        Else
            If rsTmp.State = adStateClosed Then
                vsfEpr(Index).TextMatrix(lngRow, 1) = strOldValue: vsfEpr(Index).EditText = strOldValue
            ElseIf rsTmp.EOF Then
                vsfEpr(Index).TextMatrix(lngRow, 1) = strOldValue: vsfEpr(Index).EditText = strOldValue
            Else
                blnChange = True
                vsfEpr(Index).TextMatrix(lngRow, 0) = rsTmp!ID
                vsfEpr(Index).TextMatrix(lngRow, 1) = rsTmp!姓名
            End If
        End If
    ElseIf lngCol = 3 Then '科室
        gstrSQL = "Select a.Id,a.编码, a.名称, a.简码" & vbNewLine & _
                    "From 部门表 A, 部门性质说明 B" & vbNewLine & _
                    "Where a.Id = b.部门id And b.工作性质 In ('临床') And b.服务对象 In (2, 3) And" & vbNewLine & _
                    "      (To_Char(a.撤档时间, 'yyyy-MM-dd') = '3000-01-01' Or a.撤档时间 Is Null) And" & vbNewLine & _
                    "      (Instr(a.名称, [1])>0 or Instr(a.简码,Upper([1]))>0 or Instr(a.编码,[1])>0  ) " & vbNewLine & _
                    "Order By a.编码"
        Set rsTmp = zlDatabase.ShowSQLMultiSelect(Me, gstrSQL, 0, "", False, "", "请选择一个或多个病人科室", False, False, True, vPoint.X, vPoint.Y, 225, blnCancel, True, True, vsfEpr(Index).EditText)
        If blnCancel Then
            vsfEpr(Index).TextMatrix(lngRow, 3) = strOldValue
        Else
            If rsTmp.State = adStateClosed Then
                vsfEpr(Index).TextMatrix(lngRow, 3) = strOldValue: vsfEpr(Index).EditText = strOldValue
            ElseIf rsTmp.EOF Then
                vsfEpr(Index).TextMatrix(lngRow, 3) = strOldValue: vsfEpr(Index).EditText = strOldValue
            Else
                blnChange = True
                Do Until rsTmp.EOF
                    If rsTmp.AbsolutePosition = 1 Then
                        vsfEpr(Index).TextMatrix(lngRow, 2) = rsTmp!ID
                        vsfEpr(Index).TextMatrix(lngRow, 3) = rsTmp!名称
                    Else
                        vsfEpr(Index).TextMatrix(lngRow, 2) = vsfEpr(Index).TextMatrix(lngRow, 2) & "," & rsTmp!ID
                        vsfEpr(Index).TextMatrix(lngRow, 3) = vsfEpr(Index).TextMatrix(lngRow, 3) & vbCrLf & rsTmp!名称
                    End If
                    rsTmp.MoveNext
                Loop
            End If
            vsfEpr(Index).AutoSize 3, 3
        End If
    End If
    
    If blnChange Then
        If vsfEpr(Index).TextMatrix(vsfEpr(Index).Rows - 1, 0) <> "" And vsfEpr(Index).TextMatrix(vsfEpr(Index).Rows - 1, 2) <> "" Then
            vsfEpr(Index).Rows = vsfEpr(Index).Rows + 1
        End If
        
        For l = 1 To vsfEpr(Index).Rows - 1
            If vsfEpr(Index).TextMatrix(l, 0) <> "" And vsfEpr(Index).TextMatrix(l, 2) <> "" Then
                strNewValue = strNewValue & ";" & vsfEpr(Index).TextMatrix(l, 0) & "," & vsfEpr(Index).TextMatrix(l, 2)
             End If
        Next
        If strNewValue <> "" Then
            strNewValue = Mid(strNewValue, 2)
        End If
        
        Call SetParChange(vsfEpr, Index, mrsPar, True, strNewValue)
    End If
End If
End Sub

Private Sub vsfEpr_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim lngRow As Long, lngCol As Long, lngRowHeight As Long, lngTop As Long, lngLeft As Long, i As Integer, blnScroll As Boolean, lngRowsHeight As Long
    If Index = 1 Or Index = 2 Then
        lngRow = vsfEpr(Index).MouseRow: lngCol = vsfEpr(Index).MouseCol
        If lngRow >= 1 And lngCol >= 0 Then
            For i = 1 To vsfEpr(Index).Rows - 1
                lngRowsHeight = lngRowsHeight + vsfEpr(Index).RowHeight(i) + IIF(vsfEpr(Index).RowHeight(i) = vsfEpr(Index).RowHeight(0), 30, 0)
            Next
            If lngRowsHeight > vsfEpr(Index).ClientHeight - vsfEpr(Index).RowHeight(0) - 30 Then blnScroll = True '认为产生滚动，CMD的位置相应向左移动240
            
            For i = vsfEpr(Index).TopRow To lngRow - 1
                lngTop = lngTop + vsfEpr(Index).RowHeight(i) + IIF(vsfEpr(Index).RowHeight(i) = vsfEpr(Index).RowHeight(0), 30, 0)
            Next
            lngTop = lngTop + vsfEpr(Index).RowHeight(0) + 30
            lngLeft = IIF(lngCol = 1, 1600, vsfEpr(Index).Width - 10 - IIF(blnScroll, 240, 0)) - cmdEprSelect(Index).Width - 30
            lngRowHeight = vsfEpr(Index).RowHeight(lngRow)
            Call cmdEprSelect(Index).Move(lngLeft, lngTop, cmdEprSelect(Index).Width, lngRowHeight)
            cmdEprSelect(Index).Visible = True
        End If
    End If
    
    Call SetParTip(vsfEpr, Index, mrsPar)
End Sub

Private Sub vsfMecItem_DblClick()
    Call cmdModify_Click
End Sub

Private Sub vsfMecItem_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        If vsfMecItem.Row > 0 Then
            Call cmdDelete_Click
        End If
    End If
End Sub

Private Sub vsfWaittingMixDept_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim l As Long
    Dim strTmp As String
    
    With vsfWaittingMixDept
        For l = 1 To .Rows - 1
            If Val(.TextMatrix(l, .ColIndex("启用"))) <> 0 Then
                strTmp = strTmp & "," & Trim(.TextMatrix(l, .ColIndex("ID")))
            End If
        Next
        If strTmp <> "" Then strTmp = Mid(strTmp, 2)
        
        Call SetParChange(vsfWaittingMixDept, 0, mrsPar, True, strTmp)
    End With
End Sub

Private Sub vsfWaittingMixDept_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = vsfWaittingMixDept.ColKey(Col) <> "启用"
End Sub

Private Sub vsfWaittingMixDept_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(vsfWaittingMixDept, 0, mrsPar)
End Sub

Private Sub vsUnWriteDept_CellChanged(ByVal Row As Long, ByVal Col As Long)
    Dim strValue As String
    
    If Me.Visible Then
        strValue = Get科室(vsUnWriteDept)
        Call SetParChange(vsUnWriteDept, 0, mrsPar, True, strValue)
    End If
End Sub

Private Sub vsStopDept_CellChanged(ByVal Row As Long, ByVal Col As Long)
    Dim strValue As String
    
    If Me.Visible Then
        strValue = Get科室(vsStopDept)
        Call SetParChange(vsStopDept, 0, mrsPar, True, strValue)
    End If
End Sub

Private Sub vsUnWriteDept_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rsTmp As New ADODB.Recordset
    
    With vsUnWriteDept
        If KeyCode > 127 Then
            '解决直接输入汉字的问题
            Call vsUnWriteDept_KeyPress(KeyCode)
        ElseIf KeyCode = vbKeyDelete Then
            Call mcol科室.Remove("_" & Val(.TextMatrix(.Row, .Col + 4)))
            .TextMatrix(.Row, .Col) = ""
            .Cell(flexcpData, .Row, .Col) = ""
            .TextMatrix(.Row, .Col + 4) = ""
            
            Call vsUnWriteDept_AfterEdit(.Row, .Col)
        End If
        If KeyCode = vbKeyReturn Then
            If .Editable = flexEDNone Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
            Call EnterNextCell(vsUnWriteDept)
        End If
    End With
End Sub

Private Sub vsStopDept_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rsTmp As New ADODB.Recordset
    
    With vsStopDept
        If KeyCode > 127 Then
            '解决直接输入汉字的问题
            Call vsStopDept_KeyPress(KeyCode)
        ElseIf KeyCode = vbKeyDelete Then
            Call mcolStop科室.Remove("_" & Val(.TextMatrix(.Row, .Col + 4)))
            .TextMatrix(.Row, .Col) = ""
            .Cell(flexcpData, .Row, .Col) = ""
            .TextMatrix(.Row, .Col + 4) = ""
            
            Call vsStopDept_AfterEdit(.Row, .Col)
        End If
        If KeyCode = vbKeyReturn Then
            If .Editable = flexEDNone Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
            Call EnterNextCell(vsStopDept)
        End If
    End With
End Sub


Private Sub vsUnWriteDept_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If vsUnWriteDept.TextMatrix(Row, Col + 4) = "" Then vsUnWriteDept.TextMatrix(Row, Col) = ""
End Sub

Private Sub vsStopDept_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If vsStopDept.TextMatrix(Row, Col + 4) = "" Then vsStopDept.TextMatrix(Row, Col) = ""
End Sub


Private Sub vsUnWriteDept_KeyPress(KeyAscii As Integer)
    If vsUnWriteDept.Editable = flexEDNone Then Exit Sub

    With vsUnWriteDept
        If KeyAscii = 13 Then
            KeyAscii = 0
        Else
            If KeyAscii = Asc("*") Then
                KeyAscii = 0
                Call vsUnWriteDept_CellButtonClick(.Row, .Col)
            Else
                If KeyAscii = vbKeyBack Then Exit Sub
                .ComboList = "" '使按钮状态进入输入状态
            End If
        End If
    End With
End Sub

Private Sub vsStopDept_KeyPress(KeyAscii As Integer)
    If vsStopDept.Editable = flexEDNone Then Exit Sub

    With vsStopDept
        If KeyAscii = 13 Then
            KeyAscii = 0
        Else
            If KeyAscii = Asc("*") Then
                KeyAscii = 0
                Call vsStopDept_CellButtonClick(.Row, .Col)
            Else
                If KeyAscii = vbKeyBack Then Exit Sub
                .ComboList = "" '使按钮状态进入输入状态
            End If
        End If
    End With
End Sub

Private Sub Load科室(vsTmp As VSFlexGrid, ByVal strIn As String)
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim i As Integer
    Dim lngRow As Long
    Dim lngCol As Long
    
    vsTmp.Clear
    If strIn = "" Then Exit Sub
    
    strIn = Replace(strIn, "|", ",")
    strSQL = "select id,名称 from 部门表 where id in (Select Column_Value From Table(f_Num2list([1]))) Order by 编码"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strIn)
    If rsTmp.EOF Then Exit Sub
    
    With vsTmp
        lngRow = (rsTmp.RecordCount + 3) \ 4
        If lngRow > 5 Then .Rows = lngRow
        
        For i = 1 To rsTmp.RecordCount
            If vsTmp.Name = "vsStopDept" Then
                Call mcolStop科室.Add(rsTmp!ID & "", "_" & rsTmp!ID)
            Else
                Call mcol科室.Add(rsTmp!ID & "", "_" & rsTmp!ID)
            End If
            lngRow = (i - 1) \ 4
            lngCol = (i - 1) Mod 4
            
            .TextMatrix(lngRow, lngCol) = rsTmp!名称
            .Cell(flexcpData, lngRow, lngCol) = rsTmp!名称 & ""
            .TextMatrix(lngRow, lngCol + 4) = rsTmp!ID
            rsTmp.MoveNext
        Next
    End With
    
    Exit Sub
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub




Private Function Get科室(vsTmp As VSFlexGrid) As String
    Dim i As Integer
    Dim j As Integer
    Dim strIds As String
    
    With vsTmp
        For i = 0 To .Rows - 1
            For j = 0 To 3
                If .TextMatrix(i, j) <> "" Then
                    strIds = strIds & "|" & Val(.TextMatrix(i, j + 4))
                End If
            Next
        Next
    End With
    
    Get科室 = Mid(strIds, 2)
End Function



Private Sub vsUnWriteDept_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    With vsUnWriteDept
    If .Editable = flexEDNone Then
        .FocusRect = flexFocusLight
        .ComboList = ""
    Else
        .FocusRect = flexFocusSolid
        .ComboList = "..."
    End If
    End With
End Sub

Private Sub vsStopDept_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    With vsStopDept
    If .Editable = flexEDNone Then
        .FocusRect = flexFocusLight
        .ComboList = ""
    Else
        .FocusRect = flexFocusSolid
        .ComboList = "..."
    End If
    End With
End Sub

Private Sub vsUnWriteDept_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsUnWriteDept
    If .Editable = flexEDNone Then
        .FocusRect = flexFocusLight
        .ComboList = ""
    Else
        .FocusRect = flexFocusSolid
        .ComboList = "..."
    End If
    End With
End Sub

Private Sub vsStopDept_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsStopDept
    If .Editable = flexEDNone Then
        .FocusRect = flexFocusLight
        .ComboList = ""
    Else
        .FocusRect = flexFocusSolid
        .ComboList = "..."
    End If
    End With
End Sub


Private Sub vsUnWriteDept_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim vPoint As POINTAPI
    Dim blnCancel As Boolean, blnInput As Boolean
    
    On Error GoTo errH
    strSQL = "select Distinct a.id,a.编码,a.名称,a.简码 from 部门表 a,部门性质说明 b where a.id=b.部门id" & _
        " and b.工作性质='临床' And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null) Order by A.简码"
    With vsUnWriteDept
        vPoint = zlControl.GetCoordPos(.hwnd, .CellLeft, .CellTop)
        Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 0, "临床科室", , , , , , True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, , True)
        If Not rsTmp Is Nothing Then
            blnInput = SetDeptInput(vsUnWriteDept, Row, Col, rsTmp)
            If blnInput Then
                Call EnterNextCell(vsUnWriteDept)
                Call vsUnWriteDept_AfterEdit(Row, Col)
            Else
                Call vsUnWriteDept_AfterRowColChange(-1, -1, Row, Col)
            End If
        Else
            If Not blnCancel Then
                MsgBox "没有可用的科室数据，请先到部门管理中设置。", vbInformation, gstrSysName
            End If
        End If
    End With
    Exit Sub
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vsStopDept_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim vPoint As POINTAPI
    Dim blnCancel As Boolean, blnInput As Boolean
    
    On Error GoTo errH
    strSQL = "select Distinct a.id,a.编码,a.名称,a.简码 from 部门表 a,部门性质说明 b where a.id=b.部门id" & _
        " and b.工作性质='临床' And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null) Order by A.简码"
    With vsStopDept
        vPoint = zlControl.GetCoordPos(.hwnd, .CellLeft, .CellTop)
        Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 0, "临床科室", , , , , , True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, , True)
        If Not rsTmp Is Nothing Then
            blnInput = SetDeptInput(vsStopDept, Row, Col, rsTmp)
            If blnInput Then
                Call EnterNextCell(vsStopDept)
                Call vsStopDept_AfterEdit(Row, Col)
            Else
                Call vsStopDept_AfterRowColChange(-1, -1, Row, Col)
            End If
        Else
            If Not blnCancel Then
                MsgBox "没有可用的科室数据，请先到部门管理中设置。", vbInformation, gstrSysName
            End If
        End If
    End With
    Exit Sub
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vsUnWriteDept_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strInput As String
    Dim vPoint As POINTAPI, blnCancel As Boolean, blnInput As Boolean
    
    With vsUnWriteDept
        If .EditText = CStr(.Cell(flexcpData, Row, Col)) Then
            Call EnterNextCell(vsUnWriteDept)
            Exit Sub
        End If
        strInput = UCase(.EditText)
        strSQL = "select Distinct a.id,a.编码,a.名称,a.简码 from 部门表 a,部门性质说明 b where a.id=b.部门id" & _
            " and b.工作性质='临床' And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null)" & _
            " And (Upper(A.编码) Like [1] Or Upper(A.名称) Like [2] Or Upper(A.简码) Like [2])" & _
            " Order by A.简码"
        With vsUnWriteDept
            vPoint = zlControl.GetCoordPos(.hwnd, .CellLeft, .CellTop)
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "临床科室", False, "", "", False, False, True, _
                vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, strInput & "%", gstrLike & strInput & "%")
            If Not rsTmp Is Nothing Then
                If SetDeptInput(vsUnWriteDept, Row, Col, rsTmp) Then
                    .EditText = .TextMatrix(Row, Col)
                    Call EnterNextCell(vsUnWriteDept)
                    Exit Sub
                End If
            Else
                If Not blnCancel Then
                    MsgBox "没有找到匹配的科室。", vbInformation, gstrSysName
                End If
                Cancel = True
            End If
        End With
        Call vsUnWriteDept_AfterRowColChange(Row, Col, Row, Col) '重新显示按钮
    End With
End Sub

Private Sub vsStopDept_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strInput As String
    Dim vPoint As POINTAPI, blnCancel As Boolean, blnInput As Boolean
    
    With vsStopDept
        If .EditText = CStr(.Cell(flexcpData, Row, Col)) Then
            Call EnterNextCell(vsStopDept)
            Exit Sub
        End If
        strInput = UCase(.EditText)
        strSQL = "select Distinct a.id,a.编码,a.名称,a.简码 from 部门表 a,部门性质说明 b where a.id=b.部门id" & _
            " and b.工作性质='临床' And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null)" & _
            " And (Upper(A.编码) Like [1] Or Upper(A.名称) Like [2] Or Upper(A.简码) Like [2])" & _
            " Order by A.简码"
        With vsStopDept
            vPoint = zlControl.GetCoordPos(.hwnd, .CellLeft, .CellTop)
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "临床科室", False, "", "", False, False, True, _
                vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, strInput & "%", gstrLike & strInput & "%")
            If Not rsTmp Is Nothing Then
                If SetDeptInput(vsStopDept, Row, Col, rsTmp) Then
                    .EditText = .TextMatrix(Row, Col)
                    Call EnterNextCell(vsStopDept)
                    Exit Sub
                End If
            Else
                If Not blnCancel Then
                    MsgBox "没有找到匹配的科室。", vbInformation, gstrSysName
                End If
                Cancel = True
            End If
        End With
        Call vsStopDept_AfterRowColChange(Row, Col, Row, Col) '重新显示按钮
    End With
End Sub


Private Sub SaveDepartSign()
'保存部门使用图片签名参数
    Dim i As Long
    
    On Error GoTo ErrHandle
    With vsfDepartSign
        For i = 1 To .Rows - 1
            If Val(.Cell(flexcpChecked, i, .ColIndex("启用"))) <> Decode(Val(.RowData(i)), 1, flexChecked, flexUnchecked) Then
                Call zlDatabase.SetPara("签名使用图片", IIF(.TextMatrix(i, .ColIndex("启用")) = "0", "0", "1"), glngSys, p病历内部工具, , Val(.TextMatrix(i, .ColIndex("ID"))))
            End If
        Next
    End With
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Save医嘱内容()
'保存医嘱内容定义
    Dim blnTrans As Boolean

    On Error GoTo ErrHandle
    If cmdAdvice.Tag = "已修改" Then
        
        gcnOracle.BeginTrans: blnTrans = True
        gstrSQL = "zl_医嘱内容定义_Delete"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        
        mrsAdvice.Filter = 0
        Do While Not mrsAdvice.EOF
            If Not IsNull(mrsAdvice!医嘱内容) Then
                gstrSQL = "zl_医嘱内容定义_Insert('" & mrsAdvice!诊疗类别 & "','" & Replace(mrsAdvice!医嘱内容, "'", "''") & "')"
                Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            End If
            mrsAdvice.MoveNext
        Loop
        gcnOracle.CommitTrans: blnTrans = False
        cmdAdvice.Tag = ""
    End If
    
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
    If blnTrans Then gcnOracle.RollbackTrans
    Call SaveErrLog
End Sub

Private Sub Save药房科室对照()
'保存药房科室对照
    Dim i As Long, strOutFilter As String, strInFilter
    Dim str方案 As String
    
    If SST.Tag <> "已修改" Then
        mrsPar.Filter = "参数名='药房科室对照方案' and 修改状态=1 "
        If Not mrsPar.EOF Then SST.Tag = "已修改"
    End If
    
    If SST.Tag = "已修改" Then
        Call zlDatabase.DelDeptPara("门诊缺省西药房", glngSys, p门诊医嘱下达)
        Call zlDatabase.DelDeptPara("门诊缺省成药房", glngSys, p门诊医嘱下达)
        Call zlDatabase.DelDeptPara("门诊缺省中药房", glngSys, p门诊医嘱下达)
        Call zlDatabase.DelDeptPara("门诊可用西药房", glngSys, p门诊医嘱下达)
        Call zlDatabase.DelDeptPara("门诊可用成药房", glngSys, p门诊医嘱下达)
        Call zlDatabase.DelDeptPara("门诊可用中药房", glngSys, p门诊医嘱下达)
        Call zlDatabase.DelDeptPara("门诊缺省发料部门", glngSys, p门诊医嘱下达)
        Call zlDatabase.DelDeptPara("门诊可用发料部门", glngSys, p门诊医嘱下达)
        strOutFilter = mrs门诊药房对照.Filter: strInFilter = mrs住院药房对照.Filter
        mrs门诊药房对照.Filter = 0: mrs住院药房对照.Filter = 0
        If mrs门诊药房对照.RecordCount > 0 Then mrs门诊药房对照.MoveFirst
        If mrs住院药房对照.RecordCount > 0 Then mrs住院药房对照.MoveFirst
        Do While Not mrs门诊药房对照.EOF
            If mrs门诊药房对照!科室IDs & "" <> "" Then
                str方案 = str方案 & ";" & mrs门诊药房对照!科室IDs
                For i = 0 To UBound(Split(mrs门诊药房对照!科室IDs, ","))
                    Call zlDatabase.SetPara("门诊缺省西药房", mrs门诊药房对照!缺省西药房 & "", glngSys, p门诊医嘱下达, , Split(mrs门诊药房对照!科室IDs, ",")(i))
                    Call zlDatabase.SetPara("门诊缺省成药房", mrs门诊药房对照!缺省成药房 & "", glngSys, p门诊医嘱下达, , Split(mrs门诊药房对照!科室IDs, ",")(i))
                    Call zlDatabase.SetPara("门诊缺省中药房", mrs门诊药房对照!缺省中药房 & "", glngSys, p门诊医嘱下达, , Split(mrs门诊药房对照!科室IDs, ",")(i))
                    Call zlDatabase.SetPara("门诊可用西药房", mrs门诊药房对照!可用西药房 & "", glngSys, p门诊医嘱下达, , Split(mrs门诊药房对照!科室IDs, ",")(i))
                    Call zlDatabase.SetPara("门诊可用成药房", mrs门诊药房对照!可用成药房 & "", glngSys, p门诊医嘱下达, , Split(mrs门诊药房对照!科室IDs, ",")(i))
                    Call zlDatabase.SetPara("门诊可用中药房", mrs门诊药房对照!可用中药房 & "", glngSys, p门诊医嘱下达, , Split(mrs门诊药房对照!科室IDs, ",")(i))
                    Call zlDatabase.SetPara("门诊缺省发料部门", mrs门诊药房对照!缺省发料部门 & "", glngSys, p门诊医嘱下达, , Split(mrs门诊药房对照!科室IDs, ",")(i))
                    Call zlDatabase.SetPara("门诊可用发料部门", mrs门诊药房对照!可用发料部门 & "", glngSys, p门诊医嘱下达, , Split(mrs门诊药房对照!科室IDs, ",")(i))
                Next
            End If
            mrs门诊药房对照.MoveNext
        Loop
        Call zlDatabase.SetPara("药房科室对照方案", Mid(str方案, 2), glngSys, p门诊医嘱下达)
        str方案 = ""
        Call zlDatabase.DelDeptPara("住院缺省西药房", glngSys, p住院医嘱下达)
        Call zlDatabase.DelDeptPara("住院缺省成药房", glngSys, p住院医嘱下达)
        Call zlDatabase.DelDeptPara("住院缺省中药房", glngSys, p住院医嘱下达)
        Call zlDatabase.DelDeptPara("住院可用西药房", glngSys, p住院医嘱下达)
        Call zlDatabase.DelDeptPara("住院可用成药房", glngSys, p住院医嘱下达)
        Call zlDatabase.DelDeptPara("住院可用中药房", glngSys, p住院医嘱下达)
        Call zlDatabase.DelDeptPara("住院缺省发料部门", glngSys, p住院医嘱下达)
        Call zlDatabase.DelDeptPara("住院可用发料部门", glngSys, p住院医嘱下达)
        Do While Not mrs住院药房对照.EOF
            If mrs住院药房对照!科室IDs & "" <> "" Then
                str方案 = str方案 & ";" & mrs住院药房对照!科室IDs
                For i = 0 To UBound(Split(mrs住院药房对照!科室IDs, ","))
                    Call zlDatabase.SetPara("住院缺省西药房", mrs住院药房对照!缺省西药房 & "", glngSys, p住院医嘱下达, , Split(mrs住院药房对照!科室IDs, ",")(i))
                    Call zlDatabase.SetPara("住院缺省成药房", mrs住院药房对照!缺省成药房 & "", glngSys, p住院医嘱下达, , Split(mrs住院药房对照!科室IDs, ",")(i))
                    Call zlDatabase.SetPara("住院缺省中药房", mrs住院药房对照!缺省中药房 & "", glngSys, p住院医嘱下达, , Split(mrs住院药房对照!科室IDs, ",")(i))
                    Call zlDatabase.SetPara("住院可用西药房", mrs住院药房对照!可用西药房 & "", glngSys, p住院医嘱下达, , Split(mrs住院药房对照!科室IDs, ",")(i))
                    Call zlDatabase.SetPara("住院可用成药房", mrs住院药房对照!可用成药房 & "", glngSys, p住院医嘱下达, , Split(mrs住院药房对照!科室IDs, ",")(i))
                    Call zlDatabase.SetPara("住院可用中药房", mrs住院药房对照!可用中药房 & "", glngSys, p住院医嘱下达, , Split(mrs住院药房对照!科室IDs, ",")(i))
                    Call zlDatabase.SetPara("住院缺省发料部门", mrs住院药房对照!缺省发料部门 & "", glngSys, p住院医嘱下达, , Split(mrs住院药房对照!科室IDs, ",")(i))
                    Call zlDatabase.SetPara("住院可用发料部门", mrs住院药房对照!可用发料部门 & "", glngSys, p住院医嘱下达, , Split(mrs住院药房对照!科室IDs, ",")(i))
                Next
            End If
            mrs住院药房对照.MoveNext
        Loop
        Call zlDatabase.SetPara("药房科室对照方案", Mid(str方案, 2), glngSys, p住院医嘱下达)
        
        mrs门诊药房对照.Filter = IIF(strOutFilter = "0", 0, strOutFilter): mrs住院药房对照.Filter = IIF(strInFilter = "0", 0, strInFilter)
        SST.Tag = ""
    End If
End Sub

Private Sub Set药房科室对照参数性质(ByVal lngPar As Long)
'功能：改变参数性质
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strPar As String
    Dim blnTrans As Boolean
    Dim i As Long
    
    On Error GoTo errH
    If lngPar = 0 Then
        strPar = ",0,0,0,'" & gstrUserName & "','修改关联参数(科室药房对照按本机参数设置)影响',1)"
    ElseIf lngPar = 1 Then
        strPar = ",0,1,0,'" & gstrUserName & "','修改关联参数(科室药房对照按本机参数设置)影响',0)"
    End If
    If strPar <> "" Then
        strSQL = "Select n.id From zlParameters N Where n.模块 In (1252, 1253)" & vbNewLine & _
                "And  n.参数名 In ('门诊缺省西药房', '门诊缺省成药房', '门诊缺省中药房', '门诊缺省发料部门', '门诊可用西药房', '门诊可用成药房', '门诊可用中药房', '住院缺省西药房', '住院缺省成药房'," & vbNewLine & _
                "                '住院缺省中药房', '住院缺省发料部门', '住院可用西药房', '住院可用成药房', '住院可用中药房','门诊可用发料部门','住院可用发料部门')"

        Set rsTmp = New ADODB.Recordset
        Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
        gcnOracle.BeginTrans: blnTrans = True
        For i = 1 To rsTmp.RecordCount
            zlDatabase.ExecuteProcedure "Zl_Parameters_Change(" & rsTmp!ID & strPar, Me.Caption
            rsTmp.MoveNext
        Next
        gcnOracle.CommitTrans: blnTrans = False
    End If
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans: blnTrans = False
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub lst_ItemCheck(Index As Integer, Item As Integer)
    If Me.Visible Then
        If Index = lst_本科执行自动完成医嘱类别 Then
            Call SetValue住院执行
            Frame14.Tag = "已修改"
        Else
            Call SetParChange(lst, Index, mrsPar)
        End If
    End If
End Sub

Private Sub txtUD_Validate(Index As Integer, Cancel As Boolean)
    If Val(txtUD(Index).Text) > ud(Index).Max Or Val(txtUD(Index).Text) < ud(Index).Min Then
        txtUD(Index).Text = ud(Index).value
    End If
End Sub

Private Sub txtUD_Change(Index As Integer)
    Dim strValue As String
    
    If Me.Visible Then
        Select Case Index
        Case txtud_体温夜班开始时点, txtud_体温夜班结束时点
                    strValue = txtUD(txtud_体温夜班开始时点).Text & ";" & txtUD(txtud_体温夜班结束时点).Text
                    If Index = txtud_体温夜班开始时点 Then
                        Call SetParChange(txtUD, txtud_体温夜班结束时点, mrsPar, True, strValue)
                    Else
                        Call SetParChange(txtUD, txtud_体温夜班开始时点, mrsPar, True, strValue)
                    End If
                    Call SetParChange(txtUD, Index, mrsPar, True, strValue)
        Case Else
                Call SetParChange(txtUD, Index, mrsPar)
        End Select
    End If
End Sub

Private Sub txtUD_GotFocus(Index As Integer)
    Call zlControl.TxtSelAll(txtUD(Index))
End Sub

Private Sub txtUD_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub cmdAdvice_Click()
    If frmAdviceDefine.ShowMe(Me, mrsAdvice) Then
        '标记为已变化,需要保存
        cmdAdvice.Tag = "已修改"
    End If
End Sub

Private Sub cbo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub chk_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub


Private Sub lst_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub


Private Sub cbo_Click(Index As Integer)
    Dim blnValue As Boolean, strValue As String
    Dim arrIndex As Variant, i As Integer
    
    Select Case Index
    Case cbo_合理用药接口
        '美康时可见
        lblPassVer.Visible = InStr(",1,2,", "," & cbo(Index).ListIndex & ",") > 0
        optPASSVer(0).Visible = InStr(",1,2,", "," & cbo(Index).ListIndex & ",") > 0
        optPASSVer(1).Visible = InStr(",1,2,", "," & cbo(Index).ListIndex & ",") > 0
        
        If cbo(Index).ListIndex = 0 Then    '未启用接口
            chk(chk_禁忌药嘱).Enabled = False
            chk(chk_禁忌药嘱).value = 0
            chk(chk_禁止下达超极量药品医嘱).Enabled = False
            chk(chk_禁止下达超极量药品医嘱).value = 0
            chk(chk_允许院外执行禁忌药品).Enabled = False
            chk(chk_允许院外执行禁忌药品).value = 0
            chk(chk_禁忌药品要求填写原因).Enabled = False
            chk(chk_禁忌药品要求填写原因).value = 0
        
            chk(chk_接口调用日志_大通).Visible = False  '大通时可见
            chk(chk_使用系统设置_美康).Visible = False  '美康时可见
      
            '太元通时可见
            cbo(cmd_过敏输入来源).Visible = False
            lblInfo(lbl_过敏输入来源).Visible = False
            cmdSet.Visible = False
        Else
            chk(chk_禁忌药嘱).Enabled = True
            chk(chk_允许院外执行禁忌药品).Enabled = chk(chk_禁忌药嘱).value = 0
            chk(chk_禁忌药品要求填写原因).Enabled = chk(chk_禁忌药嘱).value = 1 And InStr(",1,2,3,6,", "," & cbo(cbo_合理用药接口).ListIndex & ",") > 0
            If Not chk(chk_禁忌药品要求填写原因).Enabled Then chk(chk_禁忌药品要求填写原因).value = 0
            
            If cbo(Index).ListIndex = 1 Then  '美康
                chk(chk_使用系统设置_美康).Visible = True
                chk(chk_使用系统设置_美康).Enabled = True
                optPASSVer(0).Caption = "美康3.0"
                optPASSVer(1).Caption = "美康4.0"
            Else
                chk(chk_使用系统设置_美康).Visible = False
                chk(chk_使用系统设置_美康).Enabled = False
            End If

            If cbo(Index).ListIndex = 2 Then  '大通
                chk(chk_禁止下达超极量药品医嘱).Enabled = True
                chk(chk_接口调用日志_大通).Visible = True
                optPASSVer(0).Caption = "CS版"
                optPASSVer(1).Caption = "BS版"
            Else
                chk(chk_禁止下达超极量药品医嘱).Enabled = False
                chk(chk_禁止下达超极量药品医嘱).value = 0
                chk(chk_接口调用日志_大通).Visible = False
            End If
            If cbo(Index).ListIndex = 3 Then    '太元通
                cbo(cmd_过敏输入来源).ListIndex = 0
                cbo(cmd_过敏输入来源).Visible = True
                lblInfo(lbl_过敏输入来源).Visible = True
                cbo(cmd_过敏输入来源).Enabled = True
                lblInfo(lbl_过敏输入来源).Enabled = True
            Else
                cbo(cmd_过敏输入来源).Visible = False
                lblInfo(lbl_过敏输入来源).Visible = False
            End If
            If mobjPass Is Nothing Then
                On Error Resume Next
                Set mobjPass = CreateObject("zlPassInterface.clsPass")
                Err.Clear: On Error GoTo 0
            End If
            If Not mobjPass Is Nothing Then
                If optPASSVer(1).Visible Then strValue = IIF(optPASSVer(1).value, "4.0", "3.0")
                cmdSet.Visible = mobjPass.SetEnabled(cbo(Index).ListIndex, strValue)
            End If
        End If
    Case cbo_中药配方
        blnValue = True
        strValue = IIF(cbo(cbo_中药配方).ListIndex = 1, 4, 3)
    Case cbo_门诊药房科室对照方案, cbo_住院药房科室对照方案
        Call Set科室药房对照(Index)
    Case cbo_住院本科执行自动完成方案
        Call Set自动完成方案(Index)
    Case cbo_宫口扩大顺产, cbo_先露下降顺产
        blnValue = True
        strValue = cbo(cbo_宫口扩大顺产).ListIndex & ";" & cbo(cbo_先露下降顺产).ListIndex
        If Index = cbo_宫口扩大顺产 Then
            If Me.Visible Then Call SetParChange(cbo, cbo_先露下降顺产, mrsPar, blnValue, strValue)
        Else
            If Me.Visible Then Call SetParChange(cbo, cbo_宫口扩大顺产, mrsPar, blnValue, strValue)
        End If
    Case cbo_宫口扩大异常产, cbo_先露下降异常产
        blnValue = True
        strValue = cbo(cbo_宫口扩大异常产).ListIndex & ";" & cbo(cbo_先露下降异常产).ListIndex
        If Index = cbo_宫口扩大异常产 Then
            If Me.Visible Then Call SetParChange(cbo, cbo_先露下降异常产, mrsPar, blnValue, strValue)
        Else
            If Me.Visible Then Call SetParChange(cbo, cbo_宫口扩大异常产, mrsPar, blnValue, strValue)
        End If
    Case cbo_生产标志内容, cbo_生产标志位置
        blnValue = True
        strValue = cbo(cbo_生产标志内容).ListIndex & ";" & cbo(cbo_生产标志位置).ListIndex
        If Index = cbo_生产标志内容 Then
            If Me.Visible Then Call SetParChange(cbo, cbo_生产标志位置, mrsPar, blnValue, strValue)
        Else
            If Me.Visible Then Call SetParChange(cbo, cbo_生产标志内容, mrsPar, blnValue, strValue)
        End If
    Case cbo_警戒线显示, cbo_异常线显示
        blnValue = True
        strValue = cbo(cbo_警戒线显示).ListIndex & ";" & cbo(cbo_异常线显示).ListIndex
        If Index = cbo_警戒线显示 Then
            If Me.Visible Then Call SetParChange(cbo, cbo_异常线显示, mrsPar, blnValue, strValue)
        Else
            If Me.Visible Then Call SetParChange(cbo, cbo_警戒线显示, mrsPar, blnValue, strValue)
        End If
    Case cbo_入院自动标志, cbo_入科自动标志, cbo_转科自动标志, cbo_换床自动标志, cbo_手术自动标志, cbo_出院自动标志, _
        cbo_分娩自动标志, cbo_出生自动标志, cbo_回室自动标志, cbo_转病区自动标志
        blnValue = True
        strValue = cbo(cbo_入院自动标志).ListIndex & ";" & cbo(cbo_入科自动标志).ListIndex & ";" & cbo(cbo_转科自动标志).ListIndex & ";" & _
            cbo(cbo_换床自动标志).ListIndex & ";" & cbo(cbo_手术自动标志).ListIndex & ";" & cbo(cbo_出院自动标志).ListIndex & ";" & _
            cbo(cbo_分娩自动标志).ListIndex & ";" & cbo(cbo_出生自动标志).ListIndex & ";" & cbo(cbo_回室自动标志).ListIndex & ";" & _
            cbo(cbo_转病区自动标志).ListIndex
        arrIndex = Array(cbo_入院自动标志, cbo_入科自动标志, cbo_转科自动标志, cbo_换床自动标志, cbo_手术自动标志, cbo_出院自动标志, _
            cbo_分娩自动标志, cbo_出生自动标志, cbo_回室自动标志, cbo_转病区自动标志)
        For i = 0 To UBound(arrIndex)
            If Index <> CInt(arrIndex(i)) Then
                If Me.Visible Then Call SetParChange(cbo, CInt(arrIndex(i)), mrsPar, blnValue, strValue)
            End If
        Next i
    Case cbo_输血采集默认诊疗类型
        blnValue = True
        If cbo(cbo_输血采集默认诊疗类型).ListIndex >= 0 Then
            strValue = zlCommFun.GetNeedName(cbo(cbo_输血采集默认诊疗类型).List(cbo(cbo_输血采集默认诊疗类型).ListIndex), "-")
        End If
    Case cbo_小结缺省标识
        blnValue = True
        strValue = cbo(cbo_小结缺省标识).ListIndex & ";" & chk(chk_小结缺省标识).value
        If Me.Visible Then Call SetParChange(cbo, cbo_小结缺省标识, mrsPar, blnValue, strValue)
    End Select
    
    If Me.Visible Then
        Call SetParChange(cbo, Index, mrsPar, blnValue, strValue)
    End If
End Sub

Private Sub Set科室药房对照(ByVal lngIndex As Long)
    Dim strSQL As String, rsTmp As Recordset
    Dim strDeptIDs As String
    Dim i As Long, j As Long
    Dim strDefault As String, strDSIDs As String, str窗口 As String, strPar As String
    
    If lngIndex = cbo_门诊药房科室对照方案 Then
        mrs门诊药房对照.Filter = "方案=" & cbo(lngIndex).ItemData(cbo(lngIndex).ListIndex)
        If mrs门诊药房对照.RecordCount > 0 Then strDeptIDs = mrs门诊药房对照!科室IDs & ""
    ElseIf lngIndex = cbo_住院药房科室对照方案 Then
        mrs住院药房对照.Filter = "方案=" & cbo(lngIndex).ItemData(cbo(lngIndex).ListIndex)
        If mrs住院药房对照.RecordCount > 0 Then strDeptIDs = mrs住院药房对照!科室IDs & ""
    End If
    strSQL = "select ID,名称 From 部门表 Where ID in(Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)))"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strDeptIDs)
    vsUseDept(lngIndex).Enabled = True
    vsfDrugStore(lngIndex).Enabled = True
 
    vsUseDept(lngIndex).Rows = 0
    Do While Not rsTmp.EOF
        vsUseDept(lngIndex).Rows = vsUseDept(lngIndex).Rows + 1
        vsUseDept(lngIndex).TextMatrix(vsUseDept(lngIndex).Rows - 1, 0) = rsTmp!名称 & ""
        vsUseDept(lngIndex).Cell(flexcpData, vsUseDept(lngIndex).Rows - 1, 0) = rsTmp!ID & ""

        rsTmp.MoveNext
    Loop
    If vsUseDept(lngIndex).Rows = 0 Then
        vsUseDept(lngIndex).Rows = 1
    Else
        '科室药房对照
        With vsfDrugStore(lngIndex)
            For i = 1 To .Rows - 1
                strDefault = IIF(lngIndex = cbo_门诊药房科室对照方案, mrs门诊药房对照, mrs住院药房对照).Fields("缺省" & .TextMatrix(i, .ColIndex("类别"))).value & ""
                strDSIDs = "," & IIF(lngIndex = cbo_门诊药房科室对照方案, mrs门诊药房对照, mrs住院药房对照).Fields("可用" & .TextMatrix(i, .ColIndex("类别"))).value & ","
                
                If Val(.RowData(i)) = Val(strDefault) Then
                    .TextMatrix(i, .ColIndex("缺省")) = "√"
                    .TextMatrix(i, .ColIndex("可用")) = -1   'true
                Else
                    .TextMatrix(i, .ColIndex("缺省")) = ""
                    .TextMatrix(i, .ColIndex("可用")) = IIF(InStr(strDSIDs, "," & Val(.RowData(i)) & ",") > 0, -1, 0)
                End If
                strPar = IIF(lngIndex = cbo_门诊药房科室对照方案, mrs门诊药房对照, mrs住院药房对照).Fields("缺省发料部门").value & ""
                 
            Next
        End With
    End If
End Sub

Private Sub chk_Click(Index As Integer)
    Dim blnValue As Boolean, strValue As String
    Dim strIndexs As String '对于参数是分成按位存取的要特殊处理。
    Dim varTmp As Variant
    Dim i As Long
    
    Select Case Index
        Case chk_过敏登记有效天数
            txtUD(ud_过敏登记有效天数).Enabled = chk(Index).value = 1
            txtUD(ud_过敏登记有效天数).BackColor = IIF(chk(Index).value = 1, RGB(255, 255, 255), Me.BackColor)
            ud(ud_过敏登记有效天数).Enabled = txtUD(ud_过敏登记有效天数).Enabled
            strValue = IIF(chk(Index).value = 1, ud(ud_过敏登记有效天数).value, "0")
            blnValue = True
        Case chk_门诊处方条数限制
            txtUD(ud_门诊处方条数限制).Enabled = chk(Index).value = 1
            txtUD(ud_门诊处方条数限制).BackColor = IIF(chk(Index).value = 1, RGB(255, 255, 255), Me.BackColor)
            ud(ud_门诊处方条数限制).Enabled = txtUD(ud_门诊处方条数限制).Enabled
            strValue = IIF(chk(Index).value = 1, ud(ud_门诊处方条数限制).value, "0")
            blnValue = True
        Case chk_抗菌药物分级管理
            chk(chk_抗菌药物使用自备药).Enabled = chk(Index).value = 1
            chk(chk_按医疗小组进行抗菌药物审核).Enabled = chk(Index).value = 1
        Case chk_禁忌药嘱
            chk(chk_允许院外执行禁忌药品).value = 0
            chk(chk_禁忌药品要求填写原因).value = 0
            chk(chk_允许院外执行禁忌药品).Enabled = chk(Index).value = 0
            chk(chk_禁忌药品要求填写原因).Enabled = chk(Index).value = 1 And InStr(",1,2,3,6,", "," & cbo(cbo_合理用药接口).ListIndex & ",") > 0
        Case chk_手术分级管理
            If chk(Index).value = 0 Then
                chk(chk_手术授权管理).value = 0
                chk(chk_手术分级审核).value = 0
                chk(chk_主刀医师达到手术等级无需审核).value = 0
            End If
            chk(chk_手术授权管理).Enabled = chk(Index).value = 1
            chk(chk_手术分级审核).Enabled = chk(Index).value = 1
            chk(chk_主刀医师达到手术等级无需审核).Enabled = chk(Index).value = 1
        Case chk_输血分级管理
            If chk(Index).value = 0 Then
                chk(chk_输血申请三级审核).value = 0
                chk(chk_输血申请限制中级及以上医师).value = 0
            End If
            chk(chk_输血申请三级审核).Enabled = chk(Index).value = 1
            chk(chk_输血申请限制中级及以上医师).Enabled = chk(Index).value = 1
        Case chk_启用血库管理系统
            If mblnUseBlood Then
                If chk(Index).value = 0 Then
                    chk(chk_用血医嘱发送后才能发血).value = 0
                    chk(chk_输血申请不显示血液库存).value = 0
                    chk(chk_下达用血申请时确定发血信息).value = 0
                    chk(chk_血液接收后才允许执行登记).value = 0
                    If cbo(cbo_输血采集默认诊疗类型).ListCount > 0 Then
                        cbo(cbo_输血采集默认诊疗类型).ListIndex = 0
                    End If
                End If
                chk(chk_用血医嘱发送后才能发血).Enabled = chk(Index).value = 1
                chk(chk_输血申请不显示血液库存).Enabled = chk(Index).value = 1
                chk(chk_下达用血申请时确定发血信息).Enabled = chk(Index).value = 1
                chk(chk_血液接收后才允许执行登记).Enabled = chk(Index).value = 1
                lblBloodManager.Enabled = chk(Index).value = 1
                cbo(cbo_输血采集默认诊疗类型).Enabled = chk(Index).value = 1
            End If
        Case chk_医嘱执行有效天数
            txtUD(ud_医嘱执行有效天数).Enabled = chk(Index).value = 1
            txtUD(ud_医嘱执行有效天数).BackColor = IIF(chk(Index).value = 1, RGB(255, 255, 255), Me.BackColor)
            ud(ud_医嘱执行有效天数).Enabled = txtUD(ud_医嘱执行有效天数).Enabled
            strValue = IIF(chk(Index).value = 1, ud(ud_医嘱执行有效天数).value, "999")
            blnValue = True
        Case chk_医嘱超量时必须输入原因
            Call SetVsfEditable(vsUnWriteDept, chk(Index).value = 1)
        Case chk_停嘱时录入原因
            Call SetVsfEditable(vsStopDept, chk(Index).value = 1)
        Case chk_长嘱口服药发送结束时间
            dtp口服结束时间.Enabled = chk(Index).value = 1
            blnValue = True
            strValue = IIF(chk(Index).value = 1, "1|" & dtp口服结束时间.value, "0")
        Case chk_输血医嘱执行后需要核对, chk_皮试医嘱执行后需要核对
            blnValue = True
            strValue = chk(chk_输血医嘱执行后需要核对).value & chk(chk_皮试医嘱执行后需要核对).value
            strIndexs = chk_输血医嘱执行后需要核对 & "|" & chk_皮试医嘱执行后需要核对
        Case chk_住院本科自动执行长嘱, chk_住院本科自动执行临嘱
            blnValue = True
            strValue = chk(chk_住院本科自动执行长嘱).value & chk(chk_住院本科自动执行临嘱).value
            lst(lst_本科执行自动完成医嘱类别).Enabled = Val(strValue) <> 0
        Case chk_住院下达自动排序
            cmdAdviceSortSet.Enabled = chk(Index).value = 1
        Case chk_住院下达加皮试
            opt未皮试限制医嘱(0).Enabled = chk(Index).value = 1
            opt未皮试限制医嘱(1).Enabled = opt未皮试限制医嘱(0).Enabled
        Case chk_路径执行环节医生启用, chk_路径执行环节护士启用
            blnValue = True
            strValue = chk(chk_路径执行环节医生启用).value & chk(chk_路径执行环节护士启用).value
        Case chk_启用路径执行环节
            If chk(Index).value = 0 Then
                chk(chk_路径执行环节医生启用).value = 0
                chk(chk_路径执行环节医生启用).Enabled = False
                chk(chk_路径执行环节护士启用).value = 0
                chk(chk_路径执行环节护士启用).Enabled = False
            Else
                chk(chk_路径执行环节医生启用).Enabled = True
                chk(chk_路径执行环节护士启用).Enabled = True
            End If
        Case chk_允许前一天不评估就生成今天路径项目
            If chk(Index).value = 0 Then
                chk(chk_允许提前生成明天的路径项目).value = 0
                chk(chk_允许提前生成明天的路径项目).Enabled = False
            Else
                chk(chk_允许提前生成明天的路径项目).Enabled = True
            End If
        Case chk_长嘱单转科换页
            chk(chk_转科换页后在首行打印重开医嘱).Enabled = chk(Index).value = 1
        Case chk_签名使用图片
            'chk(chk_签名使用原图).Enabled = (chk(chk_签名使用图片).value = 1)
            txt(txt_签名使用图片高度).Enabled = (chk(chk_签名使用图片).value = 1)
        Case chk_签名使用原图
            txt(txt_签名使用图片高度).Enabled = Not (chk(chk_签名使用原图).value = 1)
        Case chk_产程图上显示警戒线
            cbo(cbo_警戒线显示).Enabled = chk(Index).value = 1
            cbo(cbo_异常线显示).Enabled = chk(Index).value = 1
        Case chk_申请单启用环节门诊检查, chk_申请单启用环节住院检查, chk_申请单启用环节门诊检验, chk_申请单启用环节住院检验, chk_申请单启用环节门诊输血, _
                chk_申请单启用环节住院输血, chk_申请单启用环节门诊手术, chk_申请单启用环节住院手术, chk_申请单启用环节会诊
            blnValue = True
            strValue = Get申请单启用环节
            strIndexs = chk_申请单启用环节门诊检查 & "|" & chk_申请单启用环节住院检查 & "|" & chk_申请单启用环节门诊检验 & "|" & chk_申请单启用环节住院检验 & "|" & chk_申请单启用环节门诊输血 & "|" & _
                    chk_申请单启用环节住院输血 & "|" & chk_申请单启用环节门诊手术 & "|" & chk_申请单启用环节住院手术 & "|" & chk_申请单启用环节会诊
        Case chk_启用申请单后必须使用申请单下达医嘱门诊, chk_启用申请单后必须使用申请单下达医嘱住院
            blnValue = True
            strValue = chk(chk_启用申请单后必须使用申请单下达医嘱门诊).value & chk(chk_启用申请单后必须使用申请单下达医嘱住院).value
            strIndexs = chk_启用申请单后必须使用申请单下达医嘱门诊 & "|" & chk_启用申请单后必须使用申请单下达医嘱住院
        Case chk_科室药房对照按本机参数设置
            SST.Enabled = chk(Index).value = 0
        Case chk_对应多份护理文件
            chk(chk_允许数据同步).Enabled = chk(chk_对应多份护理文件).value = 1
        Case chk_药品医嘱相同分类不算路径外医嘱
            If chk(chk_药品医嘱相同分类不算路径外医嘱).value = 1 Then chk(chk_药品医嘱不匹配为路径外项目).value = 0
        Case chk_药品医嘱不匹配为路径外项目
            If chk(chk_药品医嘱不匹配为路径外项目).value = 1 Then chk(chk_药品医嘱相同分类不算路径外医嘱).value = 0
        Case chk_小结缺省标识
            blnValue = True
            strValue = cbo(cbo_小结缺省标识).ListIndex & ";" & chk(chk_小结缺省标识).value
    End Select
    
    If Me.Visible Then
        If strIndexs = "" Then
            Call SetParChange(chk, Index, mrsPar, blnValue, strValue)
        Else
            varTmp = Split(strIndexs, "|")
            For i = 0 To UBound(varTmp)
                Call SetParChange(chk, Val(varTmp(i)), mrsPar, blnValue, strValue)
            Next
        End If
    End If
End Sub

Private Sub cmd发送划价类别_Click(Index As Integer)
    Dim j As Long
    
    If SendPriceType.Tab = 0 Then
        j = lst_门诊发送划价类别
    Else
        j = lst_住院发送划价类别
    End If
    Call SetLstSelected(lst(j), Index = 0)
End Sub


Private Sub cmd门诊发送一张单据类别_Click(Index As Integer)
    If lst(lst_门诊发送一张单据类别).Enabled = False Then Exit Sub
    Call SetLstSelected(lst(lst_门诊发送一张单据类别), Index = 0)
End Sub

Private Sub cmd门诊发送检查诊断_Click(Index As Integer)
    Call SetLstSelected(lst(lst_门诊发送检查诊断), Index = 0)
End Sub

Private Sub cmd住院检查入院诊断_Click(Index As Integer)
    Call SetLstSelected(lst(lst_住院检查入院诊断), Index = 0)
End Sub


Private Sub opt超期费用收回_Click(Index As Integer)
     '负数记帐时，不使用自动审核申请单
     
    chk(chk_超期收回自动审核本科).Enabled = (Index = 1)
    If Index = 0 Then chk(chk_超期收回自动审核本科).value = 0
    
    If Me.Visible Then Call SetParChange(opt超期费用收回, Index, mrsPar)
End Sub

Private Sub opt发送单据规则_Click(Index As Integer)
    lst(lst_门诊发送一张单据类别).Enabled = opt发送单据规则(0).value
    
    chk(chk_一并给给一张单据).Enabled = opt发送单据规则(0).value
    chk(chk_门诊检验医嘱发送时一组检验发送为一张单据).Enabled = opt发送单据规则(0).value
    
    If Me.Visible Then Call SetParChange(opt发送单据规则, Index, mrsPar)
End Sub

Private Sub opt发送单据类型_Click(Index As Integer)
    If Me.Visible Then Call SetParChange(opt发送单据类型, Index, mrsPar)
End Sub


Private Sub opt输血申请单打印_Click(Index As Integer)
    If Me.Visible Then Call SetParChange(opt输血申请单打印, Index, mrsPar)
End Sub

Private Sub opt未皮试限制医嘱_Click(Index As Integer)
    If Me.Visible Then Call SetParChange(opt未皮试限制医嘱, Index, mrsPar)
End Sub

Private Sub opt住院医嘱单打印_Click(Index As Integer)
    If Me.Visible Then Call SetParChange(opt住院医嘱单打印, Index, mrsPar)
End Sub

Private Sub dtp口服结束时间_Change()
    If Me.Visible Then
        Call SetParChange(dtp口服结束时间, 0, mrsPar, True, "1|" & dtp口服结束时间.value)
    End If
End Sub


Private Sub vsUnWriteDept_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(vsUnWriteDept, 0, mrsPar)
End Sub

Private Sub vsStopDept_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(vsStopDept, 0, mrsPar)
End Sub

Private Sub opt超期费用收回_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(opt超期费用收回, Index, mrsPar)
End Sub

Private Sub opt发送单据规则_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(opt发送单据规则, Index, mrsPar)
End Sub

Private Sub opt发送单据类型_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(opt发送单据类型, Index, mrsPar)
End Sub


Private Sub opt输血申请单打印_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(opt输血申请单打印, Index, mrsPar)
End Sub

Private Sub opt未皮试限制医嘱_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(opt未皮试限制医嘱, Index, mrsPar)
End Sub


Private Sub opt住院医嘱单打印_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(opt住院医嘱单打印, Index, mrsPar)
End Sub

Private Sub cmdAdviceSortSet_Click()
    frmParAdviceSort.mbytFun = 1
    frmParAdviceSort.Show vbModal, Me
End Sub

Private Sub vsUseDept_AfterRowColChange(Index As Integer, ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If vsUseDept(Index).Editable = flexEDNone Then
        vsUseDept(Index).FocusRect = flexFocusLight
        vsUseDept(Index).ComboList = ""
    Else
        vsUseDept(Index).FocusRect = flexFocusSolid
        vsUseDept(Index).ComboList = "..."
    End If
End Sub


Private Sub vsUseDept_CellButtonClick(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    Dim vPoint As POINTAPI
    Dim strSQL As String, blnCancel As Boolean
    Dim rsTmp As Recordset, i As Long, strIds As String
    Dim strTmp As String, blnCheck As Boolean
    Dim strWhere As String
    
    If Col = 0 Then
        With vsUseDept(Index)
            For i = 0 To .Rows - 1
                If .Cell(flexcpData, i, 0) <> "" Then
                    strIds = strIds & "," & .Cell(flexcpData, i, 0)
                End If
            Next
            strIds = Mid(strIds, 2)
            
            If Index = cbo_门诊药房科室对照方案 Then
                strWhere = " and 工作性质 in ('临床','检查','检验','手术','治疗','营养') And B.服务对象 in (1,3)"
            ElseIf Index = cbo_住院药房科室对照方案 Then
                strWhere = " and 工作性质 in ('临床','检查','检验','手术','治疗','营养') And B.服务对象 in (2,3)"
            ElseIf Index = cbo_住院本科执行自动完成方案 Then
                strWhere = " And (  b.工作性质 = '临床' And ((b.服务对象 In (2, 3)) Or (b.服务对象 = 1 And Exists (Select 1 From 床位状况记录 C Where b.部门id = c.科室id)))" & vbNewLine & _
                            "    Or b.工作性质 = '护理' And b.服务对象 In (1, 2, 3))"
            End If
                        
            strSQL = "select distinct ID,Decode(instr([1],',' || ID || ','),0,0,1) AS CHECKID,编码,名称" & _
                " from 部门表 A,部门性质说明 B" & _
                " where A.ID=B.部门ID " & strWhere & _
                "       and (A.撤档时间 is null or A.撤档时间=to_date('3000-01-01','YYYY-MM-DD'))" & _
                " order by 编码"
            
            vPoint = zlControl.GetCoordPos(.hwnd, .CellLeft + 15, .CellTop)
            Set rsTmp = zlDatabase.ShowSQLMultiSelect(Me, strSQL, 0, "科室", False, "", "", False, _
                    False, True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, "," & strIds & ",")
                    
            If Not blnCancel Then
                If Not rsTmp Is Nothing Then
                    .Rows = 0
                    Do While Not rsTmp.EOF
                        If Index = cbo_门诊药房科室对照方案 Then
                            blnCheck = Check科室药房对照(mrs门诊药房对照, Val(mrs门诊药房对照!方案 & ""), Val(rsTmp!ID & ""), rsTmp!名称 & "")
                        ElseIf Index = cbo_住院药房科室对照方案 Then
                            blnCheck = Check科室药房对照(mrs住院药房对照, Val(mrs住院药房对照!方案 & ""), Val(rsTmp!ID & ""), rsTmp!名称 & "")
                        ElseIf Index = cbo_住院本科执行自动完成方案 Then
                            blnCheck = Check科室药房对照(mrs住院执行对照, Val(mrs住院执行对照!方案 & ""), Val(rsTmp!ID & ""), rsTmp!名称 & "")
                        End If
                        If blnCheck Then
                            .Rows = .Rows + 1
                            .TextMatrix(.Rows - 1, Col) = rsTmp!名称 & ""
                            .Cell(flexcpData, .Rows - 1, Col) = rsTmp!ID & ""
                            strTmp = strTmp & "," & .Cell(flexcpData, .Rows - 1, Col)
                        End If
                        rsTmp.MoveNext
                    Loop
                    If Index = cbo_门诊药房科室对照方案 Then
                        mrs门诊药房对照!科室IDs = Mid(strTmp, 2)
                        mrs门诊药房对照.Update
                        Call SetParChange(vsUseDept, Index, mrsPar, True, Get科室药房对照(mrs门诊药房对照))
                    ElseIf Index = cbo_住院药房科室对照方案 Then
                        mrs住院药房对照!科室IDs = Mid(strTmp, 2)
                        mrs住院药房对照.Update
                        Call SetParChange(vsUseDept, Index, mrsPar, True, Get科室药房对照(mrs住院药房对照))
                    ElseIf Index = cbo_住院本科执行自动完成方案 Then
                        mrs住院执行对照!科室IDs = Mid(strTmp, 2)
                        mrs住院执行对照.Update
                        Call SetParChange(vsUseDept, Index, mrsPar, True, Get科室药房对照(mrs住院执行对照))
                    End If
                Else
                    MsgBox "当前没有选择的科室。", vbInformation, Me.Caption
                    Exit Sub
                End If
            End If
        End With
    End If
End Sub

Private Function Check科室药房对照(ByVal rsTmp As Recordset, ByVal lng当前方案 As Long, ByVal lng科室ID As Long, ByVal str科室名称 As String) As Boolean
    Dim strValue As String, strFilter As String
    Dim blnYes As Boolean
    
    If rsTmp.RecordCount = 0 Then Exit Function
    strFilter = rsTmp.Filter
    rsTmp.Filter = "方案<>" & lng当前方案
    If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
    Do While Not rsTmp.EOF
        If InStr("," & rsTmp!科室IDs & ",", "," & lng科室ID & ",") > 0 Then
            If blnYes = False Then
                If MsgBox("""" & str科室名称 & """在其他方案中已经存在，是否要将" & """" & str科室名称 & """修改为本方案？", vbQuestion + vbYesNo + vbDefaultButton1, Me.Caption) = vbYes Then
                    blnYes = True
                Else
                    rsTmp.Filter = strFilter
                    Check科室药房对照 = False
                    Exit Function
                End If
            End If
            If blnYes Then
                rsTmp!科室IDs = Replace("," & rsTmp!科室IDs & ",", "," & lng科室ID & ",", ",")
                rsTmp!科室IDs = Mid(rsTmp!科室IDs, 2)
                If rsTmp!科室IDs <> "" Then rsTmp!科室IDs = Mid(rsTmp!科室IDs, 1, Len(rsTmp!科室IDs) - 1)
                rsTmp.Update
            End If
        End If
        rsTmp.MoveNext
    Loop
    rsTmp.Filter = strFilter
    Check科室药房对照 = True
End Function

Private Sub vsUseDept_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim i As Long
    Dim strTmp As String
    
    If KeyCode > 127 Then
        '解决直接输入汉字的问题
        Call vsUseDept_KeyPress(Index, KeyCode)
    ElseIf KeyCode = vbKeyDelete Then
        With vsUseDept(Index)
            .TextMatrix(.Row, 0) = ""
            .Cell(flexcpData, .Row, 0) = ""
            For i = 0 To .Rows - 1
                If .Cell(flexcpData, i, 0) <> "" Then
                    strTmp = strTmp & "," & .Cell(flexcpData, i, 0)
                End If
            Next
        End With
        If Index = cbo_门诊药房科室对照方案 Then
            mrs门诊药房对照!科室IDs = Mid(strTmp, 2)
            mrs门诊药房对照.Update
            Call SetParChange(vsUseDept, Index, mrsPar, True, Get科室药房对照(mrs门诊药房对照))
        ElseIf Index = cbo_住院药房科室对照方案 Then
            mrs住院药房对照!科室IDs = Mid(strTmp, 2)
            mrs住院药房对照.Update
            Call SetParChange(vsUseDept, Index, mrsPar, True, Get科室药房对照(mrs住院药房对照))
        ElseIf Index = cbo_住院本科执行自动完成方案 Then
            mrs住院执行对照!科室IDs = Mid(strTmp, 2)
            mrs住院执行对照.Update
            Call SetParChange(vsUseDept, Index, mrsPar, True, Get科室药房对照(mrs住院执行对照))
        End If
    ElseIf KeyCode = vbKeyReturn Then
        KeyCode = 0
        If vsUseDept(Index).Editable = flexEDNone Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
        Call EnterNextCell(vsUseDept(Index))
    End If
End Sub

Private Sub vsUseDept_KeyPress(Index As Integer, KeyAscii As Integer)
    vsUseDept(Index).ComboList = "" '使按钮状态进入输入状态
End Sub

Private Sub vsUseDept_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(vsUseDept, Index, mrsPar)
End Sub

Private Sub vsUseDept_ValidateEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim vPoint As POINTAPI
    Dim strSQL As String, blnCancel As Boolean
    Dim rsTmp As Recordset, i As Long, strFind As String, strFindSql As String
    Dim strTmp As String, blnCheck As Boolean
    Dim strWhere As String
    
    If Col = 0 Then
        With vsUseDept(Index)
            strFind = .EditText
            If strFind = "" Then
                vsUseDept_KeyDown Index, vbKeyDelete, 0
                Exit Sub
            End If
            If strFind = .TextMatrix(.Row, Col) Then
                Exit Sub
            End If
            If IsNumeric(strFind) Then
                strFindSql = " And (A.简码 like [2] OR A.名称 Like [2] or A.编码=[1]) "
            Else
                strFindSql = " And (A.简码 like [2] OR A.名称 Like [2] )"
            End If
            
            If Index = cbo_门诊药房科室对照方案 Then
                strWhere = " and 工作性质 in ('临床','检查','检验','手术','治疗','营养') And B.服务对象 in (1,3)"
            ElseIf Index = cbo_住院药房科室对照方案 Then
                strWhere = " and 工作性质 in ('临床','检查','检验','手术','治疗','营养') And B.服务对象 in (2,3)"
            ElseIf Index = cbo_住院本科执行自动完成方案 Then
                strWhere = " And (  b.工作性质 = '临床' And ((b.服务对象 In (2, 3)) Or (b.服务对象 = 1 And Exists (Select 1 From 床位状况记录 C Where b.部门id = c.科室id)))" & vbNewLine & _
                            "    Or b.工作性质 = '护理' And b.服务对象 In (1, 2, 3))"
            End If
            
            strSQL = "select distinct ID,编码,名称" & _
                " from 部门表 A,部门性质说明 B" & _
                " where A.ID=B.部门ID " & strWhere & _
                "       and (A.撤档时间 is null or A.撤档时间=to_date('3000-01-01','YYYY-MM-DD'))" & strFindSql & _
                " order by 编码"
            
            vPoint = zlControl.GetCoordPos(.hwnd, .CellLeft + 15, .CellTop)
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "科室", False, "", "", False, _
                    False, True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, strFind, "%" & UCase(strFind) & "%")
                    
            If Not blnCancel Then
                If Not rsTmp Is Nothing Then
                    For i = 0 To .Rows - 1
                        If .Cell(flexcpData, i, Col) = rsTmp!ID & "" And i <> Row Then
                            MsgBox "当前列表中已经存在""" & rsTmp!名称 & """了。", vbInformation, Me.Caption
                            Cancel = True
                            Exit Sub
                        End If
                    Next
                    If Index = cbo_门诊药房科室对照方案 Then
                        blnCheck = Check科室药房对照(mrs门诊药房对照, Val(mrs门诊药房对照!方案 & ""), Val(rsTmp!ID & ""), rsTmp!名称 & "")
                    ElseIf Index = cbo_住院药房科室对照方案 Then
                        blnCheck = Check科室药房对照(mrs住院药房对照, Val(mrs住院药房对照!方案 & ""), Val(rsTmp!ID & ""), rsTmp!名称 & "")
                    ElseIf Index = cbo_住院本科执行自动完成方案 Then
                        blnCheck = Check科室药房对照(mrs住院执行对照, Val(mrs住院执行对照!方案 & ""), Val(rsTmp!ID & ""), rsTmp!名称 & "")
                    End If
                    If blnCheck Then
                        .TextMatrix(.Row, Col) = rsTmp!名称 & ""
                        .Cell(flexcpData, .Row, Col) = rsTmp!ID & ""
                        .EditText = rsTmp!名称 & ""
                        For i = 0 To .Rows - 1
                            If .Cell(flexcpData, i, Col) <> "" Then
                                strTmp = strTmp & "," & .Cell(flexcpData, i, Col)
                            End If
                        Next
                        If Index = cbo_门诊药房科室对照方案 Then
                            mrs门诊药房对照!科室IDs = Mid(strTmp, 2)
                            mrs门诊药房对照.Update
                            Call SetParChange(vsUseDept, Index, mrsPar, True, Get科室药房对照(mrs门诊药房对照))
                        ElseIf Index = cbo_住院药房科室对照方案 Then
                            mrs住院药房对照!科室IDs = Mid(strTmp, 2)
                            mrs住院药房对照.Update
                            Call SetParChange(vsUseDept, Index, mrsPar, True, Get科室药房对照(mrs住院药房对照))
                        ElseIf Index = cbo_住院本科执行自动完成方案 Then
                            mrs住院执行对照!科室IDs = Mid(strTmp, 2)
                            mrs住院执行对照.Update
                            Call SetParChange(vsUseDept, Index, mrsPar, True, Get科室药房对照(mrs住院执行对照))
                        End If
                    Else
                        Cancel = True
                    End If
                Else
                    MsgBox "没有找到匹配的科室。", vbInformation, Me.Caption
                    Cancel = True
                End If
            End If
        End With
    End If
End Sub

Private Function Get科室药房对照(ByRef rsTmp As Recordset) As String
    Dim strValue As String, strFilter As String
    
    If rsTmp.RecordCount = 0 Then Exit Function
    strFilter = rsTmp.Filter
    rsTmp.Filter = 0
    If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
    Do While Not rsTmp.EOF
        If rsTmp!科室IDs <> "" Then
            strValue = strValue & ";" & rsTmp!科室IDs
        End If
        rsTmp.MoveNext
    Loop
    rsTmp.Filter = strFilter
    Get科室药房对照 = Mid(strValue, 2)
End Function

Private Sub setDepartSign()
'功能：加载数据至vsfDepartSign控件
    Dim strSQL As String, strTmp As String
    Dim rsTemp As ADODB.Recordset
    Dim i As Long
    
    vsfDepartSign.Rows = 1
    
    mrsPar.Filter = "模块=" & p病历内部工具 & " And 参数名='签名使用图片'"
    If mrsPar.RecordCount <= 0 Then Exit Sub

    strTmp = zlCommFun.NVL(mrsPar!参数值)
    
    On Error GoTo ErrHandle
    strSQL = "Select b.Id As ID, b.编码 As 编码, b.名称 As 科室, a.参数值 As 启用 " & _
                  "From Zldeptparas A, 部门表 B " & _
                  "Where a.部门id = b.Id And b.上级id Is Not Null And " & _
                  "a.参数id in (Select max(ID) From zlParameters Where 系统 = 100 And 模块 = 1070 And 参数名 = '签名使用图片') " & _
                  "order by ID "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取使用图片签名科室", strTmp, zl9ComLib.gstrNodeNo)
    Do While rsTemp.EOF = False
        With vsfDepartSign
            .Rows = .Rows + 1
            i = .Rows - 1
            .TextMatrix(i, .ColIndex("ID")) = rsTemp!ID
            .TextMatrix(i, .ColIndex("编码")) = rsTemp!编码 & ""
            .TextMatrix(i, .ColIndex("科室")) = rsTemp!科室 & ""
            .TextMatrix(i, .ColIndex("启用")) = rsTemp!启用 & ""
                        .RowData(i) = NVL(rsTemp!启用, 0)
        End With
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    vsfDepartSign.Rows = vsfDepartSign.Rows + 1
    Exit Sub
    
ErrHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub SetWaittingMixDept()
'功能：加载数据至vsfWaittingMixDept控件

    Dim strSQL As String, strTmp As String
    Dim rsTemp As ADODB.Recordset
    Dim l As Long
    
    vsfWaittingMixDept.Rows = 1
    
    mrsPar.Filter = "模块=" & p门诊输液管理 & " And 参数名='待配液科室列表'"
    If mrsPar.RecordCount <= 0 Then Exit Sub
    
    strTmp = zlCommFun.NVL(mrsPar!参数值)
    
    On Error GoTo ErrHandle
    strSQL = "Select Distinct a.Id, a.编码, a.名称, Decode(Nvl(c.Column_Value, 0), 0, 0, -1) 启用 " & vbNewLine & _
             "From 部门表 A, 部门性质说明 B, Table(f_Num2list([1], ',')) C " & vbNewLine & _
             "Where b.部门id = a.Id And a.Id = c.Column_Value(+) And (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null) " & vbNewLine & _
             "    And b.服务对象 In (1, 3) And b.工作性质 In ('治疗', '临床') And (a.站点 = [2] Or a.站点 Is Null) " & vbNewLine & _
             "Order By a.编码"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取门诊输液科室", strTmp, zl9ComLib.gstrNodeNo)
    Do While rsTemp.EOF = False
        With vsfWaittingMixDept
            .Rows = .Rows + 1
            l = .Rows - 1
            .TextMatrix(l, .ColIndex("ID")) = rsTemp!ID
            .TextMatrix(l, .ColIndex("编码")) = rsTemp!编码
            .TextMatrix(l, .ColIndex("科室")) = rsTemp!名称
            .TextMatrix(l, .ColIndex("启用")) = rsTemp!启用
        End With
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    
    Exit Sub
    
ErrHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub InitNurseItem()
'护理文件基础控件加载
    
    '产程图
    cbo(cbo_宫口扩大顺产).Clear
    cbo(cbo_宫口扩大顺产).AddItem "0-不显示"
    cbo(cbo_宫口扩大顺产).AddItem "1-显示虚线箭头"
    cbo(cbo_宫口扩大顺产).AddItem "2-显示实线箭头"
    cbo(cbo_宫口扩大顺产).ListIndex = 0
    
    cbo(cbo_先露下降顺产).Clear
    cbo(cbo_先露下降顺产).AddItem "0-不显示"
    cbo(cbo_先露下降顺产).AddItem "1-显示虚线箭头"
    cbo(cbo_先露下降顺产).AddItem "2-显示实线箭头"
    cbo(cbo_先露下降顺产).ListIndex = 0
    
    '73309:刘鹏飞,2014-06-24
    cbo(cbo_宫口扩大异常产).Clear
    cbo(cbo_宫口扩大异常产).AddItem "0-不显示"
    cbo(cbo_宫口扩大异常产).AddItem "1-显示虚线箭头"
    cbo(cbo_宫口扩大异常产).AddItem "2-显示实线箭头"
    cbo(cbo_宫口扩大异常产).ListIndex = 0
    
    cbo(cbo_先露下降异常产).Clear
    cbo(cbo_先露下降异常产).AddItem "0-不显示"
    cbo(cbo_先露下降异常产).AddItem "1-显示虚线箭头"
    cbo(cbo_先露下降异常产).AddItem "2-显示实线箭头"
    cbo(cbo_先露下降异常产).AddItem "3-显示直角虚线"
    cbo(cbo_先露下降异常产).ListIndex = 0
    
    cbo(cbo_生产标志内容).Clear
    cbo(cbo_生产标志内容).AddItem "0-不显示"
    cbo(cbo_生产标志内容).AddItem "1-显示生产"
    cbo(cbo_生产标志内容).AddItem "2-显示处理内容"
    cbo(cbo_生产标志内容).ListIndex = 0
    
    cbo(cbo_生产标志位置).Clear
    cbo(cbo_生产标志位置).AddItem "0-宫口扩大"
    cbo(cbo_生产标志位置).AddItem "1-先露下降"
    cbo(cbo_生产标志位置).ListIndex = 0
    
    cbo(cbo_警戒线显示).Clear
    cbo(cbo_警戒线显示).AddItem "0-虚线"
    cbo(cbo_警戒线显示).AddItem "1-实线"
    cbo(cbo_警戒线显示).ListIndex = 0
    
    cbo(cbo_异常线显示).Clear
    cbo(cbo_异常线显示).AddItem "0-虚线"
    cbo(cbo_异常线显示).AddItem "1-实线"
    cbo(cbo_异常线显示).ListIndex = 0
    
    '73309:刘鹏飞,2014-06-24
    cbo(cbo_零点与首次点连接).Clear
    cbo(cbo_零点与首次点连接).AddItem "0-不连线"
    cbo(cbo_零点与首次点连接).AddItem "1-以虚线连接"
    cbo(cbo_零点与首次点连接).AddItem "2-以实线连接"
    cbo(cbo_零点与首次点连接).ListIndex = 0
    
    '记录单
    '43588,刘鹏飞,2012-09-13,添加记录单审签模式
    cbo(cbo_审签模式).Clear
    cbo(cbo_审签模式).AddItem "0-聘任职务+审签权限"
    cbo(cbo_审签模式).AddItem "1-审签权限"
    cbo(cbo_审签模式).ListIndex = 0
    
    cbo(cbo_小结缺省标识).Clear
    cbo(cbo_小结缺省标识).AddItem "0-不处理"
    cbo(cbo_小结缺省标识).AddItem "1-上下画横线标识"
    cbo(cbo_小结缺省标识).AddItem "2-汇总值下方画双横线标识"
    cbo(cbo_小结缺省标识).AddItem "3-上方画横线标识"
    '72664:刘鹏飞,2014-07-18,添加小结标识
    cbo(cbo_小结缺省标识).AddItem "4-汇总值下方画单横线标识"
    cbo(cbo_小结缺省标识).ListIndex = 0
    
    '58414,刘鹏飞,2013-04-10,添加护士、签名列显示模式
    cbo(cbo_签名列显示模式).Clear
    cbo(cbo_签名列显示模式).AddItem "0-所有行显示"
    cbo(cbo_签名列显示模式).AddItem "1-首行显示"
    cbo(cbo_签名列显示模式).AddItem "2-首尾行显示"
    cbo(cbo_签名列显示模式).AddItem "3-尾行显示"
    cbo(cbo_签名列显示模式).ListIndex = 0
    
    '体温单
    cbo(cbo_入院自动标志).Clear
    cbo(cbo_入院自动标志).AddItem "0-不显示"
    cbo(cbo_入院自动标志).AddItem "1-显示说明"
    cbo(cbo_入院自动标志).AddItem "2-显示说明和时间"
    cbo(cbo_入院自动标志).ListIndex = 0
    
    cbo(cbo_入科自动标志).Clear
    cbo(cbo_入科自动标志).AddItem "0-不显示"
    cbo(cbo_入科自动标志).AddItem "1-显示说明"
    cbo(cbo_入科自动标志).AddItem "2-显示说明和时间"
    cbo(cbo_入科自动标志).ListIndex = 0
    
    cbo(cbo_转科自动标志).Clear
    cbo(cbo_转科自动标志).AddItem "0-不显示"
    cbo(cbo_转科自动标志).AddItem "1-显示说明"
    cbo(cbo_转科自动标志).AddItem "2-显示说明和时间"
    cbo(cbo_转科自动标志).AddItem "3-显示说明和科室"
    cbo(cbo_转科自动标志).AddItem "4-显示说明,科室,时间"
    cbo(cbo_转科自动标志).ListIndex = 0
    
    cbo(cbo_换床自动标志).Clear
    cbo(cbo_换床自动标志).AddItem "0-不显示"
    cbo(cbo_换床自动标志).AddItem "1-显示说明"
    cbo(cbo_换床自动标志).AddItem "2-显示说明和时间"
    cbo(cbo_换床自动标志).ListIndex = 0
    
    cbo(cbo_手术自动标志).Clear
    cbo(cbo_手术自动标志).AddItem "0-不显示"
    cbo(cbo_手术自动标志).AddItem "1-显示说明"
    cbo(cbo_手术自动标志).AddItem "2-显示说明和时间"
    cbo(cbo_手术自动标志).ListIndex = 0
    
    cbo(cbo_出院自动标志).Clear
    cbo(cbo_出院自动标志).AddItem "0-不显示"
    cbo(cbo_出院自动标志).AddItem "1-显示说明"
    cbo(cbo_出院自动标志).AddItem "2-显示说明和时间"
    cbo(cbo_出院自动标志).ListIndex = 0
    
    cbo(cbo_分娩自动标志).Clear
    cbo(cbo_分娩自动标志).AddItem "0-不显示"
    cbo(cbo_分娩自动标志).AddItem "1-显示说明"
    cbo(cbo_分娩自动标志).AddItem "2-显示说明和时间"
    cbo(cbo_分娩自动标志).ListIndex = 0
    
    cbo(cbo_出生自动标志).Clear
    cbo(cbo_出生自动标志).AddItem "0-不显示"
    cbo(cbo_出生自动标志).AddItem "1-显示说明"
    cbo(cbo_出生自动标志).AddItem "2-显示说明和时间"
    cbo(cbo_出生自动标志).ListIndex = 0
    
    cbo(cbo_回室自动标志).Clear
    cbo(cbo_回室自动标志).AddItem "0-不显示"
    cbo(cbo_回室自动标志).AddItem "1-显示说明"
    cbo(cbo_回室自动标志).AddItem "2-显示说明和时间"
    cbo(cbo_回室自动标志).ListIndex = 0
    
    '73235
    cbo(cbo_转病区自动标志).Clear
    cbo(cbo_转病区自动标志).AddItem "0-不显示"
    cbo(cbo_转病区自动标志).AddItem "1-显示说明"
    cbo(cbo_转病区自动标志).AddItem "2-显示说明和时间"
    cbo(cbo_转病区自动标志).AddItem "3-显示说明和病区"
    cbo(cbo_转病区自动标志).AddItem "4-显示说明,病区,时间"
    cbo(cbo_转病区自动标志).ListIndex = 0
    
    cbo(cbo_标志说名与时间连接符号).Clear
    cbo(cbo_标志说名与时间连接符号).AddItem "——"
    cbo(cbo_标志说名与时间连接符号).AddItem "于"
    cbo(cbo_标志说名与时间连接符号).AddItem "空格"
    cbo(cbo_标志说名与时间连接符号).AddItem ""
    cbo(cbo_标志说名与时间连接符号).ListIndex = 0
    
    '51338,刘鹏飞,2012-07-06
    cbo(cbo_手术当天缺省格式).Clear
    cbo(cbo_手术当天缺省格式).AddItem "0-不显示"
    cbo(cbo_手术当天缺省格式).AddItem "1-显示0"
    cbo(cbo_手术当天缺省格式).AddItem "2-显示手术次数"
    cbo(cbo_手术当天缺省格式).AddItem "3-显示汉字格式"
    cbo(cbo_手术当天缺省格式).ListIndex = 0
    
    '51512,刘鹏飞,2012-07-11
    cbo(cbo_未记说明显示位置).Clear
    cbo(cbo_未记说明显示位置).AddItem "0-显示在上面"
    cbo(cbo_未记说明显示位置).AddItem "1-显示在下面"
    cbo(cbo_未记说明显示位置).AddItem "2-不显示"
    cbo(cbo_未记说明显示位置).ListIndex = 0
    
    cbo(cbo_体温不升显示方式).Clear
    cbo(cbo_体温不升显示方式).AddItem "0-箭头"
    cbo(cbo_体温不升显示方式).AddItem "1-不升"
    cbo(cbo_体温不升显示方式).AddItem "2-符号+箭头"
    cbo(cbo_体温不升显示方式).AddItem "3-符号+不升"
    cbo(cbo_体温不升显示方式).ListIndex = 0
    
    '73316:刘鹏飞,2014-06-26
    cbo(cbo_呼吸机符号显示位置).Clear
    cbo(cbo_呼吸机符号显示位置).AddItem "0-表格栏填写呼吸机符号R(缺省方式)"
    cbo(cbo_呼吸机符号显示位置).AddItem "1-表格栏填写呼吸频次,相应时间呼吸栏横线上方纵向输出”呼吸机”,用↑标识开始↓标识终止"
    cbo(cbo_呼吸机符号显示位置).AddItem "2-表格栏填写A+呼吸值"
    zlControl.CboSetWidth cbo(cbo_呼吸机符号显示位置).hwnd, 7500
    '72663:刘鹏飞,2014-08-08
    cbo(cbo_呼吸表格数据显示位置).Clear
    cbo(cbo_呼吸表格数据显示位置).AddItem "0-由上到下(无数据继承)"
    cbo(cbo_呼吸表格数据显示位置).AddItem "1-由上到下(有数据继承)"
    cbo(cbo_呼吸表格数据显示位置).AddItem "2-由下到上(无数据继承)"
    cbo(cbo_呼吸表格数据显示位置).AddItem "3-由下到上(有数据继承)"
End Sub

Private Sub SetCOLOR(vData As OLE_COLOR, ByVal Index As Integer)
    Dim lRow As Long, lCol As Long
    shpValue(Index).Visible = True
    Select Case CStr(Hex(vData))
    Case "0"
        lblColor(Index) = "黑色"
        lRow = 0
        lCol = 0
    Case "3399"
        lblColor(Index) = "褐色"
        lRow = 0
        lCol = 1
    Case "3333"
        lblColor(Index) = "橄榄色"
        lRow = 0
        lCol = 2
    Case "3300"
        lblColor(Index) = "深绿"
        lRow = 0
        lCol = 3
    Case "663300"
        lblColor(Index) = "深青"
        lRow = 0
        lCol = 4
    Case "800000"
        lblColor(Index) = "深蓝"
        lRow = 0
        lCol = 5
    Case "993333"
        lblColor(Index) = "靛蓝"
        lRow = 0
        lCol = 6
    Case "333333"
        lblColor(Index) = "灰色-80%"
        lRow = 0
        lCol = 7
    Case "80"
        lblColor(Index) = "深红"
        lRow = 1
        lCol = 0
    Case "66FF"
        lblColor(Index) = "橙色"
        lRow = 1
        lCol = 1
    Case "8080"
        lblColor(Index) = "深黄"
        lRow = 1
        lCol = 2
    Case "8000"
        lblColor(Index) = "绿色"
        lRow = 1
        lCol = 3
    Case "808000"
        lblColor(Index) = "青色"
        lRow = 1
        lCol = 4
    Case "FF0000"
        lblColor(Index) = "蓝色"
        lRow = 1
        lCol = 5
    Case "996666"
        lblColor(Index) = "蓝-灰"
        lRow = 1
        lCol = 6
    Case "808080"
        lblColor(Index) = "灰色-50%"
        lRow = 1
        lCol = 7
    Case "FF"
        lblColor(Index) = "红色"
        lRow = 2
        lCol = 0
    Case "99FF"
        lblColor(Index) = "浅橙色"
        lRow = 2
        lCol = 1
    Case "CC99"
        lblColor(Index) = "酸橙色"
        lRow = 2
        lCol = 2
    Case "669933"
        lblColor(Index) = "海绿"
        lRow = 2
        lCol = 3
    Case "CCCC33"
        lblColor(Index) = "水绿色"
        lRow = 2
        lCol = 4
    Case "FF6633"
        lblColor(Index) = "浅蓝"
        lRow = 2
        lCol = 5
    Case "800080"
        lblColor(Index) = "紫罗兰"
        lRow = 2
        lCol = 6
    Case "999999"
        lblColor(Index) = "灰色-40%"
        lRow = 2
        lCol = 7
    Case "FF00FF"
        lblColor(Index) = "粉红"
        lRow = 3
        lCol = 0
    Case "CCFF"
        lblColor(Index) = "金色"
        lRow = 3
        lCol = 1
    Case "FFFF"
        lblColor(Index) = "黄色"
        lRow = 3
        lCol = 2
    Case "FF00"
        lblColor(Index) = "鲜绿"
        lRow = 3
        lCol = 3
    Case "FFFF00"
        lblColor(Index) = "青绿"
        lRow = 3
        lCol = 4
    Case "FFCC00"
        lblColor(Index) = "天蓝"
        lRow = 3
        lCol = 5
    Case "663399"
        lblColor(Index) = "梅红"
        lRow = 3
        lCol = 6
    Case "C0C0C0"
        lblColor(Index) = "灰色-25%"
        lRow = 3
        lCol = 7
    Case "CC99FF"
        lblColor(Index) = "玫瑰红"
        lRow = 4
        lCol = 0
    Case "99CCFF"
        lblColor(Index) = "茶色"
        lRow = 4
        lCol = 1
    Case "99FFFF"
        lblColor(Index) = "浅黄"
        lRow = 4
        lCol = 2
    Case "CCFFCC"
        lblColor(Index) = "浅绿"
        lRow = 4
        lCol = 3
    Case "FFFFCC"
        lblColor(Index) = "浅青绿"
        lRow = 4
        lCol = 4
    Case "FFCC99"
        lblColor(Index) = "淡蓝"
        lRow = 4
        lCol = 5
    Case "FF99CC"
        lblColor(Index) = "淡紫"
        lRow = 4
        lCol = 6
    Case "FFFFFF"
        lblColor(Index) = "白色"
        lRow = 4
        lCol = 7
    Case Else
        lblColor(Index) = "&H" & CStr(Hex(picColor(Index).BackColor))
    End Select
    shpBorder(Index).Visible = False
    shpValue(Index).Move lCol * 18 * Screen.TwipsPerPixelX, lRow * 18 * Screen.TwipsPerPixelY, 270, 270
    shpValue(Index).Visible = True
    If vData = -9999997 Or vData = -1 Then
    
    Else
        picColor(Index).BackColor = vData
    End If
End Sub

Private Sub Set申请单启用环节(ByVal strPar As String)
'功能：设置界面控件 参数名：申请单启用环节
    Dim strTmp As String
    Dim varTmp As Variant
    On Error GoTo errH
    varTmp = Split(strPar, "|")
    strTmp = varTmp(0)
    chk(chk_申请单启用环节门诊检查).value = Mid(strTmp, 1, 1)
    chk(chk_申请单启用环节住院检查).value = Mid(strTmp, 2, 1)
    strTmp = varTmp(1)
    chk(chk_申请单启用环节门诊检验).value = Mid(strTmp, 1, 1)
    chk(chk_申请单启用环节住院检验).value = Mid(strTmp, 2, 1)
    strTmp = varTmp(2)
    chk(chk_申请单启用环节门诊输血).value = Mid(strTmp, 1, 1)
    chk(chk_申请单启用环节住院输血).value = Mid(strTmp, 2, 1)
    strTmp = varTmp(3)
    chk(chk_申请单启用环节门诊手术).value = Mid(strTmp, 1, 1)
    chk(chk_申请单启用环节住院手术).value = Mid(strTmp, 2, 1)
    strTmp = varTmp(4)
    chk(chk_申请单启用环节会诊).value = Mid(strTmp, 1, 1)
    Exit Sub
errH:
    MsgBox "系统参数：申请单启用环节，参数值格式不对！", vbInformation, gstrSysName
    Err.Clear
End Sub

Private Function Get申请单启用环节() As String
'功能：从界面获取参数值
    Dim strTmp As String
    strTmp = chk(chk_申请单启用环节门诊检查).value & chk(chk_申请单启用环节住院检查).value & "|" & _
    chk(chk_申请单启用环节门诊检验).value & chk(chk_申请单启用环节住院检验).value & "|" & _
    chk(chk_申请单启用环节门诊输血).value & chk(chk_申请单启用环节住院输血).value & "|" & _
    chk(chk_申请单启用环节门诊手术).value & chk(chk_申请单启用环节住院手术).value & "|" & _
    chk(chk_申请单启用环节会诊).value
    Get申请单启用环节 = strTmp
End Function

Private Sub InitRs药房对照(ByRef rsIn As ADODB.Recordset)
'功能：初始化科室药房对应关系记录集
    Set rsIn = New ADODB.Recordset
    rsIn.Fields.Append "方案", adVarChar, 1000
    rsIn.Fields.Append "科室IDs", adVarChar, 400000
    rsIn.Fields.Append "可用发料部门", adVarChar, 40000
    rsIn.Fields.Append "可用西药房", adVarChar, 40000
    rsIn.Fields.Append "可用成药房", adVarChar, 40000
    rsIn.Fields.Append "可用中药房", adVarChar, 40000
    rsIn.Fields.Append "缺省发料部门", adVarChar, 40000
    rsIn.Fields.Append "缺省西药房", adVarChar, 40000
    rsIn.Fields.Append "缺省成药房", adVarChar, 40000
    rsIn.Fields.Append "缺省中药房", adVarChar, 40000
'    rsIn.Fields.Append "发料部门窗口", adVarChar, 40000
'    rsIn.Fields.Append "西药房窗口", adVarChar, 40000
'    rsIn.Fields.Append "成药房窗口", adVarChar, 40000
'    rsIn.Fields.Append "中药房窗口", adVarChar, 40000
    rsIn.CursorLocation = adUseClient
    rsIn.LockType = adLockOptimistic
    rsIn.CursorType = adOpenStatic
    rsIn.Open
End Sub

Private Function isPiotBlood() As Boolean
'是否启用血库
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    On Error GoTo ErrHand
    strSQL = "Select 1 From zlSystems Where 编号=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, 2200)
    isPiotBlood = (Not rsTmp.EOF)
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Save住院执行对照()
'功能：保存 住院本科执行自动完成方案
    Dim i As Long, strFilter As String
    Dim str方案 As String
    
    If Frame14.Tag <> "已修改" Then
        mrsPar.Filter = "参数名='住院本科执行自动完成方案' and 修改状态=1 "
        If Not mrsPar.EOF Then Frame14.Tag = "已修改"
    End If
    
    If Frame14.Tag = "已修改" Then
        str方案 = ""
        Call zlDatabase.DelDeptPara("本科执行自动完成医嘱类别", glngSys, p住院医嘱发送)
        strFilter = mrs住院执行对照.Filter
        mrs住院执行对照.Filter = 0
        If mrs住院执行对照.RecordCount > 0 Then mrs住院执行对照.MoveFirst
        Do While Not mrs住院执行对照.EOF
            If mrs住院执行对照!科室IDs & "" <> "" Then
                str方案 = str方案 & ";" & mrs住院执行对照!科室IDs
                For i = 0 To UBound(Split(mrs住院执行对照!科室IDs, ","))
                    Call zlDatabase.SetPara("本科执行自动完成医嘱类别", mrs住院执行对照!医嘱类别 & "", glngSys, p住院医嘱发送, , Split(mrs住院执行对照!科室IDs, ",")(i))
                Next
            End If
            mrs住院执行对照.MoveNext
        Loop
        Call zlDatabase.SetPara("住院本科执行自动完成方案", Mid(str方案, 2), glngSys, p住院医嘱发送)
        mrs住院执行对照.Filter = IIF(strFilter = "0", 0, strFilter)
        Frame14.Tag = ""
    End If
End Sub

Private Sub SetValue住院执行()
'功能：住院本科执行自动完成方案 某个方案的值
    Dim strTmp As String
    Dim i As Long
    For i = 0 To lst(lst_本科执行自动完成医嘱类别).ListCount - 1
        If lst(lst_本科执行自动完成医嘱类别).Selected(i) Then
            strTmp = strTmp & "," & i
        End If
    Next
    mrs住院执行对照!医嘱类别 = Mid(strTmp, 2)
End Sub

Private Sub Set自动完成方案(ByVal lngIndex As Long)
    Dim str类别 As String
    Dim i As Long
    Dim strDeptIDs As String
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    mrs住院执行对照.Filter = "方案=" & cbo(lngIndex).ItemData(cbo(lngIndex).ListIndex)
    If mrs住院执行对照.RecordCount > 0 Then
        str类别 = mrs住院执行对照!医嘱类别 & ""
        strDeptIDs = mrs住院执行对照!科室IDs & ""
    End If
    
    
    strSQL = "select ID,名称 From 部门表 Where ID in(Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)))"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strDeptIDs)
    vsUseDept(lngIndex).Enabled = True
    vsUseDept(lngIndex).Rows = 0
    Do While Not rsTmp.EOF
        vsUseDept(lngIndex).Rows = vsUseDept(lngIndex).Rows + 1
        vsUseDept(lngIndex).TextMatrix(vsUseDept(lngIndex).Rows - 1, 0) = rsTmp!名称 & ""
        vsUseDept(lngIndex).Cell(flexcpData, vsUseDept(lngIndex).Rows - 1, 0) = rsTmp!ID & ""
        rsTmp.MoveNext
    Loop
    
    If vsUseDept(lngIndex).Rows = 0 Then
        vsUseDept(lngIndex).Rows = 1
    End If
    
    vsUseDept(lngIndex).Enabled = True
    lst(lst_本科执行自动完成医嘱类别).Enabled = True
    If str类别 <> "" Then
        If str类别 = "*" Then str类别 = "012345678"
        For i = 0 To 8
            lst(lst_本科执行自动完成医嘱类别).Selected(i) = InStr(str类别, i) > 0
        Next
    End If
End Sub
