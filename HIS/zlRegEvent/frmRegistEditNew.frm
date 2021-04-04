VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Object = "*\A..\zlIDKind\zlIDKind.vbp"
Object = "*\A..\ZlPatiAddress\ZlPatiAddress.vbp"
Begin VB.Form frmRegistEditNew 
   AutoRedraw      =   -1  'True
   Caption         =   "门诊挂号处理"
   ClientHeight    =   10920
   ClientLeft      =   60
   ClientTop       =   1350
   ClientWidth     =   15105
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRegistEditNew.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   10920
   ScaleWidth      =   15105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picInfoFrame 
      AutoRedraw      =   -1  'True
      Height          =   10455
      Left            =   5880
      ScaleHeight     =   10395
      ScaleWidth      =   7875
      TabIndex        =   47
      Top             =   0
      Width           =   7935
      Begin VB.PictureBox picInfo 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         CausesValidation=   0   'False
         Height          =   9825
         Left            =   375
         ScaleHeight     =   9825
         ScaleWidth      =   7470
         TabIndex        =   48
         Top             =   330
         Width           =   7470
         Begin VB.PictureBox picDetailFee 
            BorderStyle     =   0  'None
            Height          =   2865
            Left            =   90
            ScaleHeight     =   2865
            ScaleWidth      =   7470
            TabIndex        =   93
            Top             =   4545
            Width           =   7470
            Begin VB.CheckBox chkExtra 
               Caption         =   "退附加费"
               Height          =   240
               Left            =   1320
               TabIndex        =   102
               Top             =   2160
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.TextBox txt发生时间 
               Height          =   360
               Left            =   4725
               Locked          =   -1  'True
               TabIndex        =   101
               TabStop         =   0   'False
               Top             =   2100
               Width           =   2550
            End
            Begin VB.ComboBox cbo备注 
               Height          =   330
               Left            =   525
               TabIndex        =   100
               Top             =   2520
               Width           =   6765
            End
            Begin VB.ComboBox cbo预约方式 
               Height          =   330
               Left            =   2475
               Style           =   2  'Dropdown List
               TabIndex        =   99
               Top             =   2115
               Visible         =   0   'False
               Width           =   1080
            End
            Begin VB.CheckBox chk病历费 
               Caption         =   "购买病历"
               Height          =   240
               Left            =   0
               TabIndex        =   98
               Top             =   2160
               Width           =   1275
            End
            Begin VB.ComboBox cbo费别 
               Height          =   330
               IMEMode         =   3  'DISABLE
               Left            =   4725
               Style           =   2  'Dropdown List
               TabIndex        =   97
               Top             =   0
               Width           =   2550
            End
            Begin VB.ComboBox cbo付款方式 
               Height          =   330
               Left            =   915
               Style           =   2  'Dropdown List
               TabIndex        =   96
               Top             =   435
               Width           =   2550
            End
            Begin VB.TextBox txt门诊号 
               Enabled         =   0   'False
               Height          =   360
               IMEMode         =   3  'DISABLE
               Left            =   4725
               TabIndex        =   95
               ToolTipText     =   "按空格产生新的门诊号"
               Top             =   435
               Width           =   2550
            End
            Begin VB.ComboBox cbo医疗类别 
               Height          =   330
               IMEMode         =   3  'DISABLE
               Left            =   915
               Style           =   2  'Dropdown List
               TabIndex        =   94
               Top             =   0
               Width           =   2550
            End
            Begin VSFlex8Ctl.VSFlexGrid vsfMoney 
               Height          =   1155
               Left            =   0
               TabIndex        =   103
               Top             =   885
               Width           =   7260
               _cx             =   12806
               _cy             =   2037
               Appearance      =   0
               BorderStyle     =   1
               Enabled         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   10.5
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
               GridColor       =   -2147483636
               GridColorFixed  =   -2147483632
               TreeColor       =   -2147483632
               FloodColor      =   192
               SheetBorder     =   -2147483633
               FocusRect       =   3
               HighLight       =   0
               AllowSelection  =   0   'False
               AllowBigSelection=   0   'False
               AllowUserResizing=   1
               SelectionMode   =   0
               GridLines       =   1
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   3
               Cols            =   3
               FixedRows       =   1
               FixedCols       =   0
               RowHeightMin    =   320
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"frmRegistEditNew.frx":014A
               ScrollTrack     =   0   'False
               ScrollBars      =   3
               ScrollTips      =   0   'False
               MergeCells      =   1
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
            Begin VB.Label lbl摘要 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "备注"
               Height          =   210
               Left            =   0
               TabIndex        =   110
               Top             =   2595
               Width           =   420
            End
            Begin VB.Label lbl发生时间 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "发生时间"
               Height          =   210
               Left            =   3840
               TabIndex        =   109
               Top             =   2175
               Width           =   840
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "费别"
               Height          =   210
               Left            =   4260
               TabIndex        =   108
               Top             =   60
               Width           =   420
            End
            Begin VB.Label lbl门诊号 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "门诊号"
               Height          =   210
               Left            =   4050
               TabIndex        =   107
               Top             =   495
               Width           =   630
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "付款方式"
               Height          =   210
               Left            =   45
               TabIndex        =   106
               Top             =   495
               Width           =   840
            End
            Begin VB.Label lbl医疗类别 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "医疗类别"
               Height          =   210
               Left            =   45
               TabIndex        =   105
               Top             =   60
               Width           =   840
            End
            Begin VB.Label lbl预约方式 
               AutoSize        =   -1  'True
               Caption         =   "预约方式"
               Height          =   210
               Left            =   1590
               TabIndex        =   104
               Top             =   2175
               Visible         =   0   'False
               Width           =   840
            End
         End
         Begin VB.PictureBox picTotal 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1260
            Left            =   45
            ScaleHeight     =   1260
            ScaleWidth      =   7260
            TabIndex        =   84
            TabStop         =   0   'False
            Top             =   7470
            Visible         =   0   'False
            Width           =   7260
            Begin VB.Label lblTotal 
               BackStyle       =   0  'Transparent
               Caption         =   "合计"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   24
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1005
               Left            =   90
               TabIndex        =   86
               Top             =   158
               Width           =   615
            End
            Begin VB.Label lbl合计 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "0.00"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   36
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   840
               Left            =   5655
               TabIndex        =   85
               Top             =   240
               Width           =   1410
            End
         End
         Begin VB.TextBox txtIDCard 
            Height          =   360
            Left            =   1200
            MaxLength       =   18
            TabIndex        =   11
            Tag             =   "身份证号"
            Top             =   2520
            Width           =   2550
         End
         Begin VB.PictureBox picBal 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   2265
            Left            =   4035
            ScaleHeight     =   2265
            ScaleWidth      =   3375
            TabIndex        =   63
            Top             =   7485
            Width           =   3375
            Begin VB.CommandButton cmdOK 
               Caption         =   "确定(&O)"
               Height          =   390
               Left            =   690
               TabIndex        =   23
               ToolTipText     =   "热键:F2"
               Top             =   1830
               Width           =   1200
            End
            Begin VB.CommandButton cmdCancel 
               Cancel          =   -1  'True
               Caption         =   "取消(&C)"
               Height          =   390
               Left            =   2100
               TabIndex        =   24
               Top             =   1845
               Width           =   1200
            End
            Begin VB.TextBox txt找补 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               IMEMode         =   3  'DISABLE
               Left            =   540
               Locked          =   -1  'True
               TabIndex        =   22
               TabStop         =   0   'False
               Text            =   "0.00"
               Top             =   1365
               Width           =   2760
            End
            Begin VB.TextBox txt缴款 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               IMEMode         =   3  'DISABLE
               Left            =   1995
               MaxLength       =   10
               TabIndex        =   21
               Text            =   "0.00"
               Top             =   915
               Width           =   1305
            End
            Begin VB.TextBox txt本次应缴 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00108000&
               Height          =   405
               Left            =   540
               Locked          =   -1  'True
               MaxLength       =   12
               TabIndex        =   19
               TabStop         =   0   'False
               Text            =   "0.00"
               ToolTipText     =   "本次应缴合计=累计实缴金额-累计个人帐户支付-累计冲预交额"
               Top             =   465
               Width           =   2760
            End
            Begin VB.TextBox txt合计 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00108000&
               Height          =   405
               Left            =   540
               Locked          =   -1  'True
               MaxLength       =   12
               TabIndex        =   18
               TabStop         =   0   'False
               Text            =   "0.00"
               ToolTipText     =   "本次应缴合计=累计实缴金额-累计个人帐户支付-累计冲预交额"
               Top             =   0
               Width           =   2760
            End
            Begin VB.ComboBox cbo结算方式 
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   14.25
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               IMEMode         =   3  'DISABLE
               Left            =   555
               Style           =   2  'Dropdown List
               TabIndex        =   20
               Top             =   915
               Width           =   1440
            End
            Begin VB.Label lbl找补 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "找补"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   0
               TabIndex        =   67
               Top             =   1440
               Width           =   510
            End
            Begin VB.Label lbl缴款 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "缴款"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   0
               TabIndex        =   66
               Top             =   990
               Width           =   510
            End
            Begin VB.Label lbl应缴 
               AutoSize        =   -1  'True
               Caption         =   "应缴"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   240
               Left            =   0
               TabIndex        =   65
               Top             =   540
               Width           =   510
            End
            Begin VB.Label lblSum 
               AutoSize        =   -1  'True
               Caption         =   "合计"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   240
               Left            =   0
               TabIndex        =   64
               Top             =   75
               Width           =   510
            End
         End
         Begin VB.TextBox txt家庭电话 
            Height          =   360
            Left            =   4815
            MaxLength       =   20
            TabIndex        =   13
            Top             =   2520
            Width           =   2550
         End
         Begin VB.TextBox txt年龄 
            Height          =   360
            IMEMode         =   2  'OFF
            Left            =   5730
            TabIndex        =   9
            Top             =   2055
            Width           =   930
         End
         Begin VB.ComboBox cbo性别 
            Height          =   330
            Left            =   600
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   2070
            Width           =   1185
         End
         Begin VB.CommandButton cmdMore 
            Height          =   330
            Left            =   5940
            Picture         =   "frmRegistEditNew.frx":01B4
            Style           =   1  'Graphical
            TabIndex        =   62
            TabStop         =   0   'False
            ToolTipText     =   "更多内容(Ctrl+M)"
            Top             =   1635
            Width           =   350
         End
         Begin VB.CommandButton cmdLookup 
            Height          =   330
            Left            =   5220
            Picture         =   "frmRegistEditNew.frx":073E
            Style           =   1  'Graphical
            TabIndex        =   61
            TabStop         =   0   'False
            ToolTipText     =   "查找病人(Ctrl+F)"
            Top             =   1635
            Width           =   350
         End
         Begin VB.ComboBox cbo年龄单位 
            Height          =   330
            Left            =   6675
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   2070
            Width           =   705
         End
         Begin VB.TextBox txtPatient 
            Height          =   360
            Left            =   1200
            TabIndex        =   5
            ToolTipText     =   "热键:F11"
            Top             =   1620
            Width           =   3960
         End
         Begin VB.CommandButton cmdCard 
            Height          =   330
            Left            =   5565
            Picture         =   "frmRegistEditNew.frx":0888
            Style           =   1  'Graphical
            TabIndex        =   60
            TabStop         =   0   'False
            ToolTipText     =   "绑定就诊卡:F10"
            Top             =   1635
            Width           =   350
         End
         Begin VB.CommandButton cmdComminuty 
            Height          =   330
            Left            =   6285
            Picture         =   "frmRegistEditNew.frx":0E12
            Style           =   1  'Graphical
            TabIndex        =   59
            TabStop         =   0   'False
            ToolTipText     =   "社区病人身份验证"
            Top             =   1635
            Width           =   350
         End
         Begin VB.CommandButton cmdYb 
            Caption         =   "医保"
            Height          =   330
            Left            =   6645
            TabIndex        =   58
            Top             =   1635
            Width           =   705
         End
         Begin VB.Frame Frame2 
            Height          =   30
            Left            =   0
            TabIndex        =   57
            Top             =   1515
            Width           =   20000
         End
         Begin VB.TextBox txt号别 
            BackColor       =   &H00EBFFFF&
            Height          =   360
            IMEMode         =   3  'DISABLE
            Left            =   600
            TabIndex        =   1
            ToolTipText     =   "F9定位及询问挂号科室，输入""+""仅购买病历,输入"".""键回退,输入""-""号表示显示所有号别"
            Top             =   630
            Width           =   2355
         End
         Begin VB.TextBox txt科室 
            Enabled         =   0   'False
            Height          =   360
            IMEMode         =   3  'DISABLE
            Left            =   600
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   1065
            Width           =   2355
         End
         Begin VB.ComboBox cbo医生 
            ForeColor       =   &H00C00000&
            Height          =   330
            IMEMode         =   2  'OFF
            Left            =   4815
            TabIndex        =   4
            ToolTipText     =   "当所选费别医生为空且本地参数要求输医生时才允许输入"
            Top             =   1080
            Width           =   2550
         End
         Begin VB.TextBox txtSN 
            Enabled         =   0   'False
            Height          =   360
            IMEMode         =   3  'DISABLE
            Left            =   4815
            MaxLength       =   5
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   630
            Width           =   1725
         End
         Begin VB.Frame Frame1 
            Height          =   30
            Left            =   0
            TabIndex        =   56
            Top             =   555
            Width           =   20000
         End
         Begin VB.TextBox txtFact 
            ForeColor       =   &H00C00000&
            Height          =   360
            IMEMode         =   3  'DISABLE
            Left            =   600
            TabIndex        =   55
            TabStop         =   0   'False
            Top             =   120
            Width           =   1590
         End
         Begin VB.CheckBox chkCancel 
            Caption         =   "退"
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   6585
            Style           =   1  'Graphical
            TabIndex        =   54
            ToolTipText     =   "热键:F8"
            Top             =   135
            Width           =   360
         End
         Begin VB.CheckBox chkPrint 
            Caption         =   "重"
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   6975
            Style           =   1  'Graphical
            TabIndex        =   53
            ToolTipText     =   "热键:F7"
            Top             =   135
            Width           =   360
         End
         Begin VB.ComboBox cboNO 
            ForeColor       =   &H00C00000&
            Height          =   330
            IMEMode         =   3  'DISABLE
            Left            =   4815
            TabIndex        =   52
            ToolTipText     =   "热键:F12"
            Top             =   135
            Width           =   1725
         End
         Begin VB.CheckBox chkBooking 
            Caption         =   "预"
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   6975
            Style           =   1  'Graphical
            TabIndex        =   51
            ToolTipText     =   "预挂将来的号,热键:Ctrl+F12"
            Top             =   645
            Width           =   360
         End
         Begin VB.CommandButton cmdPatiPic 
            Height          =   330
            Left            =   6585
            Picture         =   "frmRegistEditNew.frx":139C
            Style           =   1  'Graphical
            TabIndex        =   50
            ToolTipText     =   "显示病人照片,热键:Ctrl+W"
            Top             =   645
            Width           =   360
         End
         Begin VB.TextBox txtPatientPrint 
            Height          =   360
            Left            =   1200
            TabIndex        =   49
            ToolTipText     =   "热键:F11"
            Top             =   1620
            Visible         =   0   'False
            Width           =   1440
         End
         Begin zlIDKind.IDKindNew IDKind 
            Height          =   360
            Left            =   600
            TabIndex        =   68
            Top             =   1620
            Width           =   630
            _ExtentX        =   1111
            _ExtentY        =   635
            Appearance      =   2
            IDKindStr       =   $"frmRegistEditNew.frx":7BEE
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontSize        =   10.5
            FontName        =   "宋体"
            IDKind          =   -1
            ShowPropertySet =   -1  'True
            DefaultCardType =   "0"
            BackColor       =   -2147483633
         End
         Begin MSMask.MaskEdBox txt出生时间 
            Height          =   360
            Left            =   4500
            TabIndex        =   8
            Top             =   2055
            Width           =   645
            _ExtentX        =   1138
            _ExtentY        =   635
            _Version        =   393216
            MaxLength       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "hh:mm"
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txt出生日期 
            Height          =   360
            Left            =   3240
            TabIndex        =   7
            Top             =   2055
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   635
            _Version        =   393216
            AutoTab         =   -1  'True
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "YYYY-MM-DD"
            Mask            =   "####-##-##"
            PromptChar      =   "_"
         End
         Begin VSFlex8Ctl.VSFlexGrid vsfPay 
            Height          =   2220
            Left            =   90
            TabIndex        =   69
            Top             =   7485
            Width           =   3735
            _cx             =   6588
            _cy             =   3916
            Appearance      =   0
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   10.5
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
            GridColor       =   -2147483636
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483633
            FocusRect       =   3
            HighLight       =   0
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   3
            Cols            =   8
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmRegistEditNew.frx":7C9B
            ScrollTrack     =   0   'False
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   1
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
         Begin zlIDKind.IDKindNew IDKind证件 
            Height          =   375
            Left            =   600
            TabIndex        =   91
            TabStop         =   0   'False
            Top             =   2520
            Width           =   630
            _ExtentX        =   1111
            _ExtentY        =   661
            Appearance      =   2
            IDKindStr       =   "身|二代身份证|0|0|0|0|0|0"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontSize        =   12
            FontName        =   "宋体"
            IDKind          =   -1
            NotAutoAppendKind=   -1  'True
            BackColor       =   -2147483633
         End
         Begin VB.TextBox txt证件 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1200
            MaxLength       =   18
            TabIndex        =   12
            Tag             =   "证件"
            Top             =   2520
            Width           =   2550
         End
         Begin ZlPatiAddress.PatiAddress padd户口地址 
            Height          =   750
            Left            =   1005
            TabIndex        =   17
            Tag             =   "户口地址"
            Top             =   3750
            Visible         =   0   'False
            Width           =   6360
            _ExtentX        =   11218
            _ExtentY        =   1323
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxLength       =   100
            LineFeed        =   -1  'True
         End
         Begin ZlPatiAddress.PatiAddress padd家庭地址 
            Height          =   750
            Left            =   1005
            TabIndex        =   16
            Tag             =   "现住址"
            Top             =   2955
            Visible         =   0   'False
            Width           =   6360
            _ExtentX        =   11218
            _ExtentY        =   1323
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxLength       =   100
            LineFeed        =   -1  'True
         End
         Begin VB.ComboBox cbo户口地址 
            Height          =   330
            Left            =   1005
            TabIndex        =   15
            Top             =   3375
            Width           =   6360
         End
         Begin VB.ComboBox cbo家庭地址 
            Height          =   330
            Left            =   1005
            TabIndex        =   14
            Top             =   2955
            Width           =   6360
         End
         Begin VB.Label lblIDCard 
            AutoSize        =   -1  'True
            Caption         =   "证件"
            Height          =   210
            Left            =   135
            TabIndex        =   90
            ToolTipText     =   "证件号码"
            Top             =   2595
            Width           =   420
         End
         Begin VB.Label lbl户口地址 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "户口地址"
            Height          =   210
            Left            =   135
            TabIndex        =   83
            Top             =   3810
            Width           =   840
         End
         Begin VB.Label lbl家庭电话 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "家庭电话"
            Height          =   210
            Left            =   3930
            TabIndex        =   82
            Top             =   2595
            Width           =   840
         End
         Begin VB.Label lbl家庭地址 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "现住地址"
            Height          =   210
            Left            =   135
            TabIndex        =   81
            Top             =   3015
            Width           =   840
         End
         Begin VB.Label lbl出生日期 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "出生日期"
            Height          =   210
            Left            =   2340
            TabIndex        =   80
            Top             =   2130
            Width           =   840
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "年龄"
            Height          =   210
            Left            =   5280
            TabIndex        =   79
            Top             =   2130
            Width           =   420
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "病人"
            Height          =   210
            Left            =   135
            TabIndex        =   78
            Top             =   1695
            Width           =   420
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "性别"
            Height          =   210
            Left            =   135
            TabIndex        =   77
            Top             =   2130
            Width           =   420
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "号别"
            Height          =   210
            Left            =   135
            TabIndex        =   76
            Top             =   705
            Width           =   420
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "科室"
            Height          =   210
            Left            =   135
            TabIndex        =   75
            Top             =   1140
            Width           =   420
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "医生"
            Height          =   210
            Left            =   4350
            TabIndex        =   74
            Top             =   1140
            Width           =   420
         End
         Begin VB.Label lblSN 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "序号"
            Height          =   210
            Left            =   4350
            TabIndex        =   73
            Top             =   705
            Width           =   420
         End
         Begin VB.Label lblNO 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "单据号"
            ForeColor       =   &H00C00000&
            Height          =   210
            Left            =   4140
            TabIndex        =   72
            Top             =   195
            Width           =   630
         End
         Begin VB.Label lblFact 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "票号"
            ForeColor       =   &H00C00000&
            Height          =   210
            Left            =   135
            TabIndex        =   71
            Top             =   195
            Width           =   420
         End
         Begin VB.Label lblPrompt 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "卡号"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C000C0&
            Height          =   315
            Left            =   90
            TabIndex        =   70
            Top             =   8730
            Width           =   660
         End
      End
      Begin ZlPatiAddress.PatiAddress paddVerify 
         Height          =   330
         Left            =   -30
         TabIndex        =   89
         Tag             =   "户口地址"
         Top             =   960
         Visible         =   0   'False
         Width           =   6360
         _ExtentX        =   11218
         _ExtentY        =   582
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   100
      End
   End
   Begin VB.PictureBox picSerialInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   90
      Picture         =   "frmRegistEditNew.frx":7DC2
      ScaleHeight     =   1665
      ScaleWidth      =   1020
      TabIndex        =   92
      Top             =   7305
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.PictureBox picTop 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F5F5F5&
      BorderStyle     =   0  'None
      Height          =   405
      Left            =   9720
      ScaleHeight     =   405
      ScaleWidth      =   4530
      TabIndex        =   40
      Top             =   45
      Width           =   4530
      Begin VB.Label lblCancel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "退"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   390
         Left            =   975
         TabIndex        =   44
         Top             =   0
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Label lbl急 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "急"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   390
         Left            =   0
         TabIndex        =   43
         Top             =   0
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.Label lblFree 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "免"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   390
         Left            =   495
         TabIndex        =   42
         Top             =   0
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Label lbl险类 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "重庆市医保"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   345
         Left            =   0
         TabIndex        =   41
         Top             =   0
         Width           =   1605
      End
   End
   Begin VB.Timer timPlan 
      Interval        =   60000
      Left            =   4305
      Top             =   735
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   26
      Top             =   10560
      Width           =   15105
      _ExtentX        =   26644
      _ExtentY        =   635
      SimpleText      =   $"frmRegistEditNew.frx":90A4
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   882
            MinWidth        =   882
            Picture         =   "frmRegistEditNew.frx":90EB
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   16457
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1111
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "大写"
            TextSave        =   "大写"
            Key             =   "STACAPS"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   1376
            MinWidth        =   18
            Picture         =   "frmRegistEditNew.frx":997F
            Object.ToolTipText     =   "序号说明"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picPlan 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8700
      Left            =   615
      ScaleHeight     =   8700
      ScaleWidth      =   5805
      TabIndex        =   25
      Top             =   930
      Width           =   5805
      Begin VB.PictureBox picSplit 
         BorderStyle     =   0  'None
         Height          =   100
         Left            =   15
         MousePointer    =   7  'Size N S
         ScaleHeight     =   105
         ScaleWidth      =   3855
         TabIndex        =   45
         Top             =   5820
         Width           =   3855
      End
      Begin VB.PictureBox picTime 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   0
         ScaleHeight     =   465
         ScaleWidth      =   4470
         TabIndex        =   87
         Top             =   5925
         Width           =   4500
         Begin MSComCtl2.DTPicker dtpAppointmentTime 
            Height          =   360
            Left            =   3555
            TabIndex        =   31
            Top             =   75
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   635
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "HH:MM"
            Format          =   94371843
            UpDown          =   -1  'True
            CurrentDate     =   42340.4166666667
         End
         Begin VB.Label lblRegTotal 
            Caption         =   "剩余可挂合计:"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   0
            Left            =   60
            TabIndex        =   113
            ToolTipText     =   "所有限号号别的剩余可挂数量合计"
            Top             =   60
            Width           =   1875
         End
         Begin VB.Label lblRegTotal 
            AutoSize        =   -1  'True
            Caption         =   "333"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   285
            Index           =   1
            Left            =   1950
            TabIndex        =   112
            Top             =   60
            Width           =   450
         End
         Begin VB.Label lbl预约时间 
            AutoSize        =   -1  'True
            Caption         =   "时间"
            Height          =   210
            Left            =   3105
            TabIndex        =   88
            Top             =   150
            Width           =   420
         End
      End
      Begin VB.CheckBox chkShowAll 
         BackColor       =   &H00A0A0A0&
         Caption         =   "所有号别"
         ForeColor       =   &H00000005&
         Height          =   240
         Left            =   2565
         TabIndex        =   38
         TabStop         =   0   'False
         ToolTipText     =   "热键：F6(指允许的科室范围内所有号别)"
         Top             =   457
         Visible         =   0   'False
         Width           =   1380
      End
      Begin VB.PictureBox picBookingDate 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   0
         ScaleHeight     =   360
         ScaleWidth      =   5670
         TabIndex        =   27
         Top             =   0
         Width           =   5670
         Begin VB.ComboBox cboTime 
            Height          =   330
            Left            =   2955
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Top             =   0
            Width           =   840
         End
         Begin MSComCtl2.DTPicker dtpAppointmentDate 
            Height          =   360
            Left            =   960
            TabIndex        =   28
            Top             =   -15
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   635
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarTitleBackColor=   -2147483636
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   94371843
            CurrentDate     =   42340
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "预约日期"
            Height          =   210
            Left            =   90
            TabIndex        =   33
            Top             =   60
            Width           =   840
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "时段"
            Height          =   210
            Left            =   2520
            TabIndex        =   32
            Top             =   60
            Width           =   420
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfPlan 
         Height          =   5070
         Left            =   0
         TabIndex        =   30
         Top             =   750
         Width           =   3360
         _cx             =   5927
         _cy             =   8943
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
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
         BackColorAlternate=   16185078
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483633
         FocusRect       =   0
         HighLight       =   0
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   9
         GridLineWidth   =   1
         Rows            =   6
         Cols            =   9
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   322
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmRegistEditNew.frx":9EB8
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   1
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
         Begin VB.PictureBox picImgPlan 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   30
            ScaleHeight     =   225
            ScaleWidth      =   210
            TabIndex        =   111
            Top             =   60
            Width           =   210
            Begin VB.Image imgColPlan 
               Height          =   195
               Left            =   0
               Picture         =   "frmRegistEditNew.frx":9F81
               ToolTipText     =   "选择需要显示的列(Ctrl+E)"
               Top             =   0
               Width           =   195
            End
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfList 
         Height          =   2175
         Left            =   0
         TabIndex        =   39
         Top             =   6510
         Visible         =   0   'False
         Width           =   5925
         _cx             =   10451
         _cy             =   3836
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
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
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483633
         FocusRect       =   3
         HighLight       =   0
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   5
         Cols            =   3
         FixedRows       =   0
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmRegistEditNew.frx":A4CF
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
      Begin XtremeSuiteControls.ShortcutCaption sc安排 
         Height          =   315
         Left            =   15
         TabIndex        =   46
         Top             =   420
         Width           =   4605
         _Version        =   589884
         _ExtentX        =   8123
         _ExtentY        =   556
         _StockProps     =   6
         Caption         =   "挂号安排表"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H8000000C&
      Height          =   30
      Left            =   -60
      TabIndex        =   0
      Top             =   465
      Width           =   40000
   End
   Begin VB.PictureBox picPatiPicBack 
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      Height          =   2060
      Left            =   4020
      ScaleHeight     =   2055
      ScaleWidth      =   1755
      TabIndex        =   34
      Top             =   555
      Width           =   1760
      Begin VB.PictureBox picPatiPic 
         BorderStyle     =   0  'None
         Height          =   1800
         Left            =   30
         ScaleHeight     =   1800
         ScaleWidth      =   1695
         TabIndex        =   35
         Top             =   230
         Width           =   1700
         Begin VB.Image imgPatiPic 
            Height          =   1800
            Left            =   0
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1700
         End
         Begin VB.Label lblShow 
            BackStyle       =   0  'Transparent
            Caption         =   "无照片"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000A&
            Height          =   465
            Left            =   300
            TabIndex        =   36
            Top             =   750
            Width           =   1125
         End
      End
      Begin VB.Label lblClosePic 
         Alignment       =   2  'Center
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "×"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1530
         TabIndex        =   37
         Top             =   30
         Width           =   195
      End
   End
   Begin MSCommLib.MSComm com 
      Left            =   3135
      Top             =   630
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   420
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmRegistEditNew.frx":A5DB
      Left            =   1095
      Top             =   570
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
   Begin VB.Image imgDel 
      Height          =   240
      Left            =   2670
      Picture         =   "frmRegistEditNew.frx":A5EF
      Top             =   705
      Width           =   240
   End
End
Attribute VB_Name = "frmRegistEditNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
'公共入出口参数
Public mstrPrivs As String
Public mlngModul As Long
Public mbytMode As Integer '0-挂号,1-预约,2-接收,3-取消预约 ,4-退号 预约有两种模式:0-挂号,此时预约要收费,1-预约,不收费
Public mbytInState As Byte '0-执行,1-查阅
Public mintCancel As Integer '0-退号,1-退病历费,2-退附加费
Public int记录状态 As Integer '2-查阅冲销预约单据,3-查阅被冲销的原始单据 注：取消预约时 mbytinstate=1
Public mblnViewCancel As Boolean '是否查看退号单据
Public mstrNoIn As String '要接收或查阅的单据号
Public mblnCharge As Boolean '是否收费在调用
Public mstr划价NO As String '退号同时要删除的划价单
Public mblnICCard As Boolean 'IC卡发卡
'门诊医生站使用的变量
Public mblnStation As Boolean '是否医生工作站在调用挂号
Public mstrRoom As String '医生工作站的接诊诊室
Public mstrRegNo As String '医生站挂号成功时的挂号单号
Public mblnNoneCut As Boolean '是否不允许使用打折费别("挂号费别打折"权限)
Public mblnStationPrice As Boolean '医生站挂号时是否允许生成划价单收挂号费
Public mblnViewOriginal As Boolean
  
'消息模块使用的变量
Public mobjMsgModule As clsMipModule

'发卡相关变量，用于区分缺省读卡类型和缺省发卡类型
Private mCurSendCard As Ty_CardProperty   '卡费和挂号费一起收时有效，先发卡再输入姓名会引起发卡类型变量，需要用模块变量记录

'票据相关变量
Private mlngShareUseID As Long '共享领用批次ID
Private mstrUseType As String '使用类别
Private mintInvoiceFormat As Integer  '打印的发票格式,发票格式序号
Private mblnStartFactUseType As Boolean   '是否启用了使用类别
Private mintInvoicePrint As Integer  '0-不打印;1-自动打印;2-提示打印

'状态控制参数
Private mblnOneCard As Boolean      '是否启用了一卡通接口,此模式下，票号严格管理，票号范围外的发卡或绑定卡不收费
Private mrsOneCard As ADODB.Recordset
Private mlng磁卡领用ID As Long '当前使用的就诊卡领用批次
Private mblnOnVilidate As Boolean
Private mlng默认卡类别ID As Long
Private mstrCardPrivs As String, mstrForceNote As String
Private mblnPre连续 As Boolean
Private mblnUnitReg As Boolean  '在预约时是否检查挂号合作单位开放号码
Private mblnOk  As Boolean, mbln连续挂号 As Boolean
Private mblnStateChange As Boolean '用于在进行挂号序号状态处理的时候,不触发vsflex的事件
Private mstrPre号别 As String '上一个有效号别
Private mlngPreRow As Long  '上一个有效表列
Private mdbl预交余额 As Double, mblnCenter As Boolean
Private mcbrToolBar As CommandBar, mbln退号原因 As Boolean
Private mdbl个帐余额 As Double, mstr原摘要 As String
Private mdbl原金额 As Double
Private mblnCancel As Boolean
Private mblnActivate As Boolean
Private mblnReadBooking As Boolean
Private mblnAppointmentChange As Boolean
Private mblnUserCancel As Boolean

Private mblnCard As Boolean '当前是否就诊卡刷卡
Private mblnNewCard As Boolean '发新卡
Private mblnUnload As Boolean, mblnChange As Boolean
Private mblnSendCard As Boolean
Private mblnBuyHisBook As Boolean
Private mblnUnChange As Boolean
Private mblnManualInput As Boolean
Private mintSysAppLimit As Integer
Private mblnFirst As Boolean
Private mblnAlwaysSend As Boolean '非严格控制时始终发卡
Private mblnCheckNOValidity As Boolean
Private mstr门诊号 As String
Private mdatLast As Date, msngTime As Single, mlngRow As Long
Private mblnChangeByCode As Boolean
Private mcur病历 As Currency
Private mblnNoClearPrompt As Boolean
Public mblnNOMoved As Boolean
Public mintNOLength As Integer  '门诊号长度
Private mDatLastRefresh As Date '号表上次刷新时间
Private mblnReSetIDKind As Boolean '刷门诊号方式时,连续挂号后,恢复身份类别为门诊号方式
Private mblnIDCardKind  As Boolean '预约挂号时,输入身份证号后,新病人在保存后是否自动恢复到身份证号别中
Private mblnAddCardItem As Boolean '卡费和挂号费一起收取
Private mblnBoundPati As Boolean '绑定卡,不收取病人卡费
Private mblnNotClick As Boolean '是否点击了IDKind
Private mblnNotChange As Boolean '用于控制是否代码触发了txtsn的validate事件
Private mblnFinishReg As Boolean
Private mbln基本信息调整 As Boolean '是否允许调整病人基本信息
Private mblnStructAdress As Boolean  '病人地址结构化录入
Private mblnShowTown As Boolean      '乡镇地址结构化录入

'记录挂号相关费用信息
Private mrsItems As ADODB.Recordset '记录挂号项目(包括从属项目)
Private mrsInComes As ADODB.Recordset '记录收入项目(包含费用信息)
Private mrsDoctor As ADODB.Recordset '当允许输入医生时(gbln医生),客户端缓存医生信息
Private mrs家庭地址 As ADODB.Recordset  '缓存家庭地址,初始时读取地区表
Private mrsSNState As ADODB.Recordset   '当前号别的序号状态
Private mrs时间段 As ADODB.Recordset    ' 挂号安排时间段
Private mrsUnitReg As ADODB.Recordset  '合作单位控制
Private mrsBill As ADODB.Recordset     '预约接收时保存预约单据信息
Private mrsBillAdvance As ADODB.Recordset '退号时,单据对应的预交记录信息

Private mdblReg     As Double           '挂号费用
Private mlng挂号科室ID As Long
Private mstr医生姓名 As String
Private mlng医生ID As Long
Private mbln建病案 As Boolean
Private mstr退费项目IDs As String
Private mbln附加费 As Boolean, mbln主费用 As Boolean
Private mstr附加费 As String, mstr附加项目ID As String
Private mrs费别 As ADODB.Recordset '费别列表
Private mstr连续挂号_挂号NO As String, mstr连续挂号_就诊卡NO As String
Private mblnUnChkClick As Boolean  '不触发checkbox的Click事件
Private mrsALL时间段 As ADODB.Recordset '问题:45509
Private mstrCurKey As String '当前星期几

'本地模块变量
Private mobjCommunity As Object     '社区接口部件
Private mint社区 As Integer
Private mstr社区号 As String

Private mrsPlan As ADODB.Recordset '包含挂号安排信息
 
Private mrsInfo As ADODB.Recordset '包含挂号病人身份信息
Private mbln病历费 As Boolean '是否可以收取病历工本费
Private mbln包含病历费 As Boolean '退号的单据中是否包含病历费
Private mlng领用ID As Long
Private mblnLEDKey As Boolean
Private mstrSort As String '号别排序字段
Private mintIDKind As Integer '上次使用的身份类别控件
Private mbln加号   As Boolean '是否是加号这种情况

Private mstrPrePati As String '上次挂号的病人,或本次已输入或验证过身份的病人
Private mstrPreNO As String '上次号别
Private mcur合计 As Currency '当前累计到的合计金额
Private mcur应缴 As Currency '当前累计到的应缴金额
Private mint挂号数 As Integer     '连续挂号时，同一病人已挂号多张挂号数
Private mstrPrepayPrivs As String '预交权限
Private mlng记录ID As Long '预约的出诊记录ID，接收时获取
Private mlng锁号记录ID As Long '锁号的记录ID，提前接收或延后接收时与原记录ID不一致
Private mobjRegist As clsRegist

'医保相关变量
Private mintInsure As Integer
Private mlngOutModeMC As Long '本地医保设置的外挂式医保险类
Private mblnOlnyBJYB   As Boolean '仅仅是北京医保:见问题:问题:26982
Private mblnNotQuery As Boolean  '未找到插件中的数据,再保存挂号时,回填数据
Private mblnBrushPlugin As Boolean '当前是否从插件读取的病人信息
Private mstrYBPati As String '医保病人身份验证信息
Private mcur个帐余额 As Currency '个人帐户余额
Private mcur个帐透支 As Currency '个人帐户允许透支金额
Private mstr个人帐户 As String  '挂号是否允许使用个人帐户
Private mlng结帐ID As Long '医保退号时的结帐ID
'刘兴洪 问题:26962 日期:2009-12-25 11:25:27
Private Type Ty_ModulePara
    bln挂号生成队列         As Boolean '排队叫号生成队列:实质上是读取的是分诊管理的参数
    int同科限约数           As Integer  '同科室限约
    int同科限挂数           As Integer
    bln同科限挂急诊         As Boolean
    int病人预约科室数       As Integer
    int病人挂号科室数       As Integer
    lng预约有效时间         As Long
    int预约失效次数         As Integer
    bln预约接收确定挂号费   As Boolean
    bln允许住院病人挂号     As Boolean '31724
    bln预约不产生门诊号     As Boolean
    bln点击列头排序         As Boolean '是否允许点击列头排序
    bln随机序号选择         As Boolean ' 启用了序号的情况下 是否允许 操作员随机选择序号
    bln失约用于挂号         As Boolean '分时段时  失约用于挂号
    lngN天取消预约          As Long    '预约N天内不能取消预约
    bln退号审核             As Boolean '在N天内取消预约 是否需要通过审核
    lng预约限制时间         As Long    '限制预约与现在时间的最小间隔 __分钟
    lng预约缺省天数         As Long    '预约时缺省间隔天数
    bln挂号必须刷卡         As Boolean '38603
    byt家庭地址联想         As Byte  '挂号家庭地址输入方式 是否联想
    bln监护人录入           As Boolean '是否控制监护人录入
    lngN岁以下录入监护人    As Long '监护人录入控制年龄
    bln严格按时段挂号       As Boolean  '严格按时段挂号
    blnReuseCancelNO        As Boolean '已退序号允许挂号
    int专家号挂号限制       As Integer
    int专家号预约限制       As Integer
    bln禁止输入年龄         As Boolean
    byt缴款方式             As Byte
    byt接收模式             As Byte
End Type
Private Enum SortType
    by号别 '根据号别进行排序
    by科室 '根据 科室-->项目--已挂数 来进行排序
    by科室and已挂数
End Enum
Private mSortType As SortType '点击排序方式
Private mTy_Para As Ty_ModulePara
Private mstr当前星期 As String
Private mstrPre费别 As String, mstrCard结算方式 As String
Private mstr年龄 As String '原年龄
Private mstr性别 As String '原性别
Private mstr姓名 As String '原姓名
Private mstr年龄单位 As String
Private mstr出生日期 As String

'界面的一个处理流程类型
Private Enum CustomTime
    t_普通
    t_时段
End Enum
Private Enum ViewMode
     V_普通号
     v_专家号
     v_专家号分时段
     V_普通号分时段
End Enum
Private mViewMode    As ViewMode  '
Private mcustomTime  As CustomTime
Private Const SNCOLS = 10
Private Const SnArgCols = 7
Private WithEvents mobjIDCard As zlIDCard.clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1
Private WithEvents mobjICCard As clsICCard
Attribute mobjICCard.VB_VarHelpID = -1
Private WithEvents mobjfrmPatiInfo As frmPatiInfo
Attribute mobjfrmPatiInfo.VB_VarHelpID = -1
'-----------------------------------------------------------------------------------
'结算卡相关
Private Type Ty_PayMoney
    lng医疗卡类别ID As Long
    bln消费卡 As Boolean
    str结算方式 As String
    str名称 As String
    str刷卡卡号 As String
    str刷卡密码 As String
    dbl帐户余额 As Double
    objCard As Card
    Have挂号费 As Boolean
    Have卡费 As Boolean
End Type

Private mCurCardPay As Ty_PayMoney '本次卡支付
Private mstrPassWord As String
Private mcolCardPayMode As Collection
Private mobjPayCard As Card

'挂号相关状态数据类型' 2012-10-29 lgf
'暂时只用于序号控制,分时段 的状态保存
Private Type Ty_RegPlanState
    '状态记录
    str号别                 As String '选中的号别
    lngLastNO               As Long '最后的一个序号
    strLastNO_Time          As String '最后一个时段开始时间
    strLastNo_EndTime       As String '最有一个时段结束时间
    lngLastNO_X             As Long '最后一个序号所在的位置
    lngLastNO_Y             As Long '最后一个序号所在的位置
    bln序号控制             As Boolean '序号控制
    lng限号数               As Long '限号数
    lng限约数               As Long '限约数
    '状态控制变量
    '以下变量,主要用于,分时段,因为分时段的号,才有序号和时段同时存在的情况
    blnAdditionalNumber     As Boolean '是否已经追加序号 '追加序号的特点(挂出去的序号,序号大于设置的最大序号,或者时间大于或者等于,最后一个时段的结束时间)
    lngSelX                 As Long '选中的行
    lngSelY                 As Long '选中的列
    lngSelNO                As Long '选中的序号
    strSelTime              As String  '选中的序号对应时段的开始时间
End Type

Private mtyRegPlanState As Ty_RegPlanState '挂号状态类型
Private mbln发卡 As Boolean '标识当前操纵是否是发卡,True - 发卡 False - 绑定卡  问题号:56599
Private mobjHealthCard As Object '制卡接口对象
Private mblnRegReceiveByNo As Boolean '判断是否是通过在挂号窗口输入单据号进行预约接收操作 问题号:57423
'-----------------------------------------------------------------------------------
Private mobjDelCards As Cards '当前退号类别

Private Type TYPE_MedicarePAR
    医保接口打印票据 As Boolean
    使用个人帐户   As Boolean  'support挂号使用个人帐户
    连续挂号  As Boolean    'support连续挂号
    不收病历费 As Boolean   'support挂号不收取病历费
    挂号检查项目 As Boolean
End Type
Private MCPAR As TYPE_MedicarePAR
'-----------------------------------------------------------------------------------
Private Enum EM_REGISTFEE_MODE  '68991挂号费用收取方式
        EM_RG_现收 = 0
        EM_RG_划价 = 1
        EM_RG_记帐 = 2
End Enum
Private Enum EM_PATI_CHARGE_MODE    '病人收费模式
    EM_先结算后诊疗 = 0
    EM_先诊疗后结算 = 1
End Enum
Private mRegistFeeMode As EM_REGISTFEE_MODE '挂号费用收取方式
Private mPatiChargeMode As EM_PATI_CHARGE_MODE    '病人收费模式
Private mstr病人家属IDs As String '病人使用家属预交，79868
Private mblnNotEMPIQuery As Boolean '防止连续的调用接口
Private mlngEMPI病人ID As Long '接口中的病人ID
Private mstrPrePriceGrade As String
Private mblnGetBirth As Boolean '判断是否允许通过年龄计算生日

Private Sub initInsurePara(ByVal lng病人ID As Long)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化医保参数
    '入参:lng病人ID-病人ID
    '编制:刘兴洪
    '日期:2013-11-19 15:43:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    MCPAR.医保接口打印票据 = gclsInsure.GetCapability(support医保接口打印票据, lng病人ID, mintInsure)
    MCPAR.使用个人帐户 = gclsInsure.GetCapability(support挂号使用个人帐户, lng病人ID, mintInsure)
    MCPAR.连续挂号 = gclsInsure.GetCapability(support连续挂号, lng病人ID, mintInsure)
    MCPAR.不收病历费 = gclsInsure.GetCapability(support挂号不收取病历费, lng病人ID, mintInsure)
    MCPAR.挂号检查项目 = gclsInsure.GetCapability(support挂号检查项目, lng病人ID, mintInsure)
End Sub

Private Function SelectMemo(ByVal strInput As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:选择常用摘要
    '入参:strInput-输入串;为空时,表示全部
    '返回:选择成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-04 16:06:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnCancel As Boolean, strSQL As String, strWhere As String
    Dim rsInfo As ADODB.Recordset
    Dim vRect As RECT, strKey As String
    strKey = GetMatchingSting(strInput, False)
    If strInput <> "" Then
        If zlCommFun.IsCharChinese(cbo备注.Text) Then
             strWhere = " And  名称 like [1] "
        ElseIf zlCommFun.IsNumOrChar(cbo备注.Text) Then
             strWhere = " And (简码 like upper([1]) or 编码 like upper([1]))"
        End If
    End If
    
    strSQL = "" & _
     "   Select RowNum AS ID,编码,名称,简码  " & _
     "   From 常用挂号摘要 " & _
     "   Where 1=1 " & strWhere & _
     "   Order by 缺省标志"
     vRect = zlControl.GetControlRect(cbo备注.Hwnd)
     On Error GoTo Hd
     Set rsInfo = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "常用挂号摘要", False, _
                    "", "", False, False, True, vRect.Left, vRect.Top, cbo备注.Height, blnCancel, True, False, strKey)
     If blnCancel Then Exit Function
     If rsInfo Is Nothing Then
        If strInput = "" Then
            MsgBox "没有设置常用挂号摘要,请在字典管理中设置", vbOKOnly + vbInformation, gstrSysName
        End If
        zlCommFun.PressKey vbKeyTab: Exit Function
     End If
     zlControl.CboSetText Me.cbo备注, Nvl(rsInfo!名称)
     cbo备注.Tag = Nvl(rsInfo!名称)
     zlCommFun.PressKey vbKeyTab
     SelectMemo = True
     Exit Function
Hd:
    If ErrCenter() = 1 Then Resume
    SaveErrLog
End Function

Private Sub cboTime_Click()
    If mblnNotClick = True Then Exit Sub
    mblnUnChange = True
    txt号别.Text = ""
    mblnUnChange = False
    Call ShowPlans(, True)
End Sub

Private Sub cbo备注_Change()
    cbo备注.Tag = ""
End Sub

Private Sub cbo备注_Click()
    If mblnNotChange Then Exit Sub
    If chkCancel.Value = 1 Or mbytMode = 4 Then
        Call cbo备注_KeyDown(13, 0)
    End If
End Sub

Private Sub cbo备注_KeyDown(KeyCode As Integer, Shift As Integer)
    If chkCancel.Value = 1 Or mbytMode = 4 Then
        If KeyCode <> vbKeyReturn Then Exit Sub
        If Trim(cbo备注.Text) = "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
        If SetDelMemo(Trim(cbo备注.Text)) = True Then
            zlCommFun.PressKey vbKeyTab: Exit Sub
        End If
    Else
        If KeyCode <> vbKeyReturn Then Exit Sub
        If cbo备注.Tag <> "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
        If Trim(cbo备注.Text) = "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
        If SelectMemo(Trim(cbo备注.Text)) = False Then
            zlCommFun.PressKey vbKeyTab: Exit Sub
        End If
    End If
End Sub

Private Function SetDelMemo(ByVal strInput As String) As Boolean
    Dim rsMemo As ADODB.Recordset, strSQL As String
    On Error GoTo errH
    If mbln退号原因 = False Then SetDelMemo = True: Exit Function
    cbo备注.Clear
    If strInput = "" Then
        strSQL = "Select 名称,缺省标志 From 常用退号原因 Order By 缺省标志 Desc,编码"
        Set rsMemo = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        If rsMemo.RecordCount <> 0 Then
            Do While Not rsMemo.EOF
                cbo备注.AddItem rsMemo!名称
                If Val(Nvl(rsMemo!缺省标志)) = 1 Then
                    mblnNotChange = True
                    cbo备注.ListIndex = cbo备注.NewIndex: cbo备注.Tag = cbo备注.Text
                    mblnNotChange = False
                End If
                rsMemo.MoveNext
            Loop
        End If
    Else
        strSQL = "Select 名称,缺省标志,简码,编码 From 常用退号原因 Order By 缺省标志 Desc,编码"
        Set rsMemo = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        If rsMemo.RecordCount <> 0 Then
            Do While Not rsMemo.EOF
                cbo备注.AddItem rsMemo!名称

                If Nvl(rsMemo!简码) Like UCase(strInput) & "*" Or Nvl(rsMemo!编码) Like UCase(strInput) & "*" Or Nvl(rsMemo!名称) Like strInput & "*" Then
                    mblnNotChange = True
                    cbo备注.ListIndex = cbo备注.NewIndex
                    mblnNotChange = False
                    cbo备注.Tag = cbo备注.Text
                End If
                rsMemo.MoveNext
            Loop
            If cbo备注.Text = "" Then
                MsgBox "没有找到对应的退号原因,请重新输入", vbInformation, gstrSysName
                SetDelMemo = False
                Exit Function
            End If
        End If
    End If
    SetDelMemo = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    SaveErrLog
End Function

Private Sub cbo付款方式_Click()
    Dim strPriceGrade As String
    
    If mbytInState = 1 Then Exit Sub
    
    If gintPriceGradeStartType < 2 Then Exit Sub
    Call gobjPublicExpense.zlGetPriceGrade(gstrNodeNo, 0, 0, zlStr.NeedName(cbo付款方式.Text), , , strPriceGrade)
    mobjfrmPatiInfo.mstrPriceGrade = strPriceGrade
    If mstrPrePriceGrade = strPriceGrade Then Exit Sub
    mstrPrePriceGrade = strPriceGrade
    
    '31182:包含预约接收
    If (mbytMode = 2 Or mbytMode = 0 And mbytInState = 0 And mstrNoIn <> "") Then
        '预约接收
        If mTy_Para.bln预约接收确定挂号费 = False Then
            If Not mrsInfo Is Nothing Then
                Exit Sub
            End If
        End If
    End If
    
    If txt号别.Text <> "" Then
        mblnBuyHisBook = True
        Call ShowRegistFromInput
        mblnBuyHisBook = False
    End If
End Sub

Private Sub cbo年龄单位_LostFocus()
    Dim strBirth As String
    If cbo年龄单位.Locked Then Exit Sub
    '更正出生日期
    With mobjfrmPatiInfo
        '69026,冉俊明,2014-8-8,检查输入年龄
        If Trim(txt年龄.Text) <> "" Then
            If .mobjPubPatient Is Nothing Then Exit Sub
            If .mobjPubPatient.CheckPatiAge(Trim(txt年龄.Text) & cbo年龄单位.Text) = False Then
                If txt年龄.Visible And txt年龄.Enabled And Not txt年龄.Locked Then
                    txt年龄.SetFocus: Exit Sub
                End If
            End If
        End If
    
        .txt年龄.Text = txt年龄.Text
        .txt年龄.Tag = txt年龄.Text
        If .cbo年龄单位.ListCount = 0 Then CopyCboTofrmPatiInfo
        .cbo年龄单位.ListIndex = cbo年龄单位.ListIndex
        .cbo年龄单位.Visible = cbo年龄单位.Visible
        
        If cbo年龄单位.Tag <> cbo年龄单位.Text Then
            .mblnChange = False
            If mblnGetBirth Then
                If mobjfrmPatiInfo.mobjPubPatient.ReCalcBirthDay(Trim(txt年龄.Text) & cbo年龄单位.Text, strBirth) Then
                    .txt出生日期.Text = Format(strBirth, "yyyy-mm-dd")
                    .txt出生时间.Text = Format(strBirth, "hh:mm")
                End If
            End If
            .mblnChange = True
            Call ReLoadCardFee(, True)
        Else
            Exit Sub
        End If
        '89130:李南春,2015/10/13,更新出生日期
        mblnChange = False
        txt出生日期.Text = .txt出生日期.Text
        txt出生时间.Text = .txt出生时间.Text
        mblnChange = True
        cbo年龄单位.Tag = cbo年龄单位.Text
        Call ShowRegistFromInput
    End With
End Sub

Private Sub cbo性别_LostFocus()
    Call ReLoadCardFee(, True)
End Sub


Private Sub cbo预约方式_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cbo性别_Click()
    If mblnNotChange Then Exit Sub
    If cbo性别.Enabled = False Then Exit Sub
    If cbo性别.Tag <> cbo性别.Text Then
        Call ShowRegistFromInput
    End If
    cbo性别.Tag = cbo性别.Text
End Sub


Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case conMenu_File_Exit
            Unload Me
        Case 2605 '预留
            Call HoldRegNo
        Case 2604 '取消预留
            Call HoldRegNo
        Case conMenu_File_Print
            Call PrintSetup
        Case 3816 '缴预交
            Call AddDeposit
        Case conMenu_View_Refresh
            Call RefreshFace
        Case 816 '出诊表
            frmRegistList.ShowMe Me, Val(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("号源ID")))
        Case 4006 '历史挂号
            Call SelectHistoryRegist
    End Select
End Sub

Private Sub PrintSetup()
    Dim strTmp As String
    
    If gblnPrintCase Then
        strTmp = zlCommFun.ShowMsgbox("打印设置", "请选择对哪一种打印内容进行设置", "!挂号票据(&1),挂号凭条(&2),病历标签(&3)", Me, vbInformation)
        If strTmp = "挂号票据" Then
            Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1111", Me)
        End If
        If strTmp = "挂号凭条" Then
            Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1111_3", Me)
        End If
        If strTmp = "病历标签" Then
            Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1111_2", Me)
        End If
    Else
        Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1111", Me)
    End If
End Sub

Private Sub RefreshFace()
    mstrPreNO = vsfPlan.TextMatrix(vsfPlan.Row, GetCol("号别"))
    Call ShowPlans
    If gbln医生 And Not mblnStation Then Call GetAll医生
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If mbytInState = 0 And (mbytMode = 0 Or mbytMode = 1) Then
        Select Case Control.ID
        Case 816
            If vsfPlan.Row > vsfPlan.Rows - 1 Or vsfPlan.Col > vsfPlan.Cols - 1 Or vsfPlan.ColIndex("号源ID") = -1 Then
                Control.Enabled = False
            Else
                Control.Enabled = vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("号源ID")) <> ""
            End If
        Case 4006
            If mrsInfo Is Nothing Then
                Control.Enabled = False
            ElseIf mrsInfo.State <> 1 Then
                Control.Enabled = False
            ElseIf mrsInfo.RecordCount = 0 Then
                Control.Enabled = False
            ElseIf IsNull(mrsInfo!病人ID) Then
                Control.Enabled = False
            Else
                Control.Enabled = True
            End If
        End Select
    Else
        Select Case Control.ID
        Case conMenu_File_Exit
            Control.Visible = True
        Case Else
            Control.Visible = False
        End Select
    End If
End Sub

Private Sub chkBooking_Click()
    Dim blnBooking As Boolean, Curdate As Date
    
    Call SetCHKState(chkBooking)
    
    blnBooking = chkBooking.Value = 1
    picBookingDate.Visible = blnBooking
    If blnBooking Then
        lbl预约方式.Visible = True
        cbo预约方式.Visible = True
        picBookingDate.Visible = True
    Else
        lbl预约方式.Visible = False
        cbo预约方式.Visible = False
        picBookingDate.Visible = False
    End If
    lbl摘要.Visible = True: cbo备注.Visible = True
    Call SetPlanGrid
    Call SetPicTimeObjectVisible
    
    If chkBooking.Tag = "保存" Then Exit Sub
    
    mblnUnChange = True     '避免txt号别.Text = "" 时调用ShowPlans
    Call ClearBill(, False)
    mblnUnChange = False
    Curdate = zlDatabase.Currentdate
    If blnBooking And Curdate > dtpAppointmentDate.Value Then  '保留之前的预约时间
        If Curdate < gdatRegistTime Then
            dtpAppointmentDate.Value = Format(gdatRegistTime + IIf(gint预约天数 >= 7, 7, mTy_Para.lng预约缺省天数), "yyyy-MM-dd " & gstr上班时间)
            dtpAppointmentDate.MinDate = Format(gdatRegistTime, "yyyy-MM-dd 00:00")
        Else
            dtpAppointmentDate.Value = Format(Curdate + IIf(gint预约天数 >= 7, 7, mTy_Para.lng预约缺省天数), "yyyy-MM-dd " & gstr上班时间)
            dtpAppointmentDate.MinDate = Format(Curdate, "yyyy-MM-dd 00:00")  '27781:当前加一小时
        End If
    End If
    Call ShowPlans
    Call Form_Resize
    Call picPlan_Resize
    If txt号别.Visible And txt号别.Enabled Then txt号别.SetFocus
End Sub

Private Function GetPatiIDByComminuty(ByVal int社区 As Integer, ByVal strComminuty As String) As Long
    Dim rsTmp As ADODB.Recordset, strSQL As String
        
    strSQL = "Select 病人ID From 病人社区信息 Where 社区 = [1] And 社区号 = [2]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, int社区, strComminuty)
    If rsTmp.RecordCount > 0 Then GetPatiIDByComminuty = rsTmp!病人ID
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
 

Private Sub cmdComminuty_Click()
    Dim lng病人ID As Long
    Dim colInfo As Collection, strTmp As String
    
    If mrsInfo Is Nothing Then
        lng病人ID = 0
    Else
        lng病人ID = mrsInfo!病人ID
    End If
    If Not mobjCommunity Is Nothing Then
        If mobjCommunity.Identify(glngSys, mlngModul, mint社区, mstr社区号, colInfo, lng病人ID) Then
            strTmp = GetColItem(colInfo, "姓名")
            If lng病人ID = 0 Then
                lng病人ID = GetPatiIDByComminuty(mint社区, mstr社区号)
                If lng病人ID = 0 Then
                    txtPatient.Text = strTmp
                Else
                    txtPatient.Text = "-" & lng病人ID
                    Call txtPatient_Validate(False)
                End If
            Else
                If strTmp <> Trim(txtPatient.Text) Then
                    MsgBox "社区验证接口返回的病人姓名与当前病人姓名不符,请检查是否是同一病人!", vbInformation
                    Exit Sub
                End If
            End If
            strTmp = GetColItem(colInfo, "性别")
            If strTmp <> "" Then cbo性别.ListIndex = cbo.FindIndex(cbo性别, strTmp, True)
                
            strTmp = GetColItem(colInfo, "家庭地址")
            If strTmp <> "" Then cbo家庭地址.Text = strTmp
            '89242:李南春,2015/12/7,读取病人地址信息
            Call zlReadAddrInfo(padd家庭地址, lng病人ID, 0, 3, cbo家庭地址.Text)
                                       
            '详细病人信息设置
            
            Call CopyCboTofrmPatiInfo
            Call CopyInfoTofrmPatiInfo
            With mobjfrmPatiInfo
                strTmp = GetColItem(colInfo, "年龄")
                If strTmp <> "" Then Call LoadOldData(strTmp, .txt年龄, .cbo年龄单位)
                
                strTmp = GetColItem(colInfo, "出生日期")
                If IsDate(strTmp) Then
                    .mblnChange = False
                    .txt出生日期.Text = Format(strTmp, "YYYY-MM-DD")
                    .mblnChange = True
                    If CDate(.txt出生日期.Text) - CDate(strTmp) <> 0 Then .txt出生时间.Text = Format(strTmp, "HH:MM")
                    
                    .txt年龄.Text = ReCalcOld(CDate(.txt出生日期.Text), .cbo年龄单位, lng病人ID) '根据出生日期重算年龄
                    .txt年龄.Tag = .txt年龄.Text
                Else
                    .mblnChange = False
                    .txt出生日期.Text = ReCalcBirth(.txt年龄.Text, .cbo年龄单位.Text)
                    .mblnChange = True
                    .txt出生时间.Text = "__:__"
                End If
                
                txt年龄.Text = .txt年龄.Text
                txt年龄.Tag = txt年龄.Text
                cbo年龄单位.ListIndex = .cbo年龄单位.ListIndex
                Call txt年龄_Validate(False)
                
                strTmp = GetColItem(colInfo, "年龄")
                If strTmp <> "" Then .cbo国籍.ListIndex = cbo.FindIndex(.cbo国籍, strTmp, True)
                strTmp = GetColItem(colInfo, "民族")
                If strTmp <> "" Then .cbo民族.ListIndex = cbo.FindIndex(.cbo民族, strTmp, True)
                strTmp = GetColItem(colInfo, "婚姻状况")
                If strTmp <> "" Then .cbo婚姻.ListIndex = cbo.FindIndex(.cbo婚姻, strTmp, True)
                strTmp = GetColItem(colInfo, "职业")
                If strTmp <> "" Then .cbo职业.ListIndex = cbo.FindIndex(.cbo职业, strTmp)
                strTmp = GetColItem(colInfo, "身份证号")
                If strTmp <> "" Then .txt身份证号.Text = strTmp: .txt身份证号.Tag = .txt身份证号.Text
                
                strTmp = GetColItem(colInfo, "工作单位")
                If strTmp <> "" Then .txt单位名称.Text = strTmp
                strTmp = GetColItem(colInfo, "单位电话")
                If strTmp <> "" Then .txt单位电话.Text = strTmp
                strTmp = GetColItem(colInfo, "单位邮编")
                If strTmp <> "" Then .txt单位邮编.Text = strTmp
                
                strTmp = GetColItem(colInfo, "家庭电话")
                If strTmp <> "" Then .txt家庭电话.Text = strTmp
                strTmp = GetColItem(colInfo, "家庭地址邮编")
                If strTmp <> "" Then .txt家庭邮编.Text = strTmp
                strTmp = GetColItem(colInfo, "区域")
                If strTmp <> "" Then .txt区域.Text = strTmp: .txt区域.Tag = .txt区域.Text
            End With
        End If
    End If
End Sub

Private Function GetColItem(colInfo As Collection, strItem As String) As String
    If colInfo Is Nothing Then Exit Function
    
    Err.Clear: On Error Resume Next
    GetColItem = colInfo("_" & strItem)
    Err.Clear: On Error GoTo 0
End Function

Private Function CancelBespeakRegist() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:取消预约挂号
    '返回:取消成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-08 17:47:43
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    '取消预约
    If mstrNoIn = "" Then Exit Function
    If zlCommFun.ActualLen(Me.cbo备注.Text) > 50 Then
        MsgBox "备注最多只能输入25个汉字或50个字符,请检查!", vbInformation + vbOKOnly, gstrSysName
        If cbo备注.Enabled And cbo备注.Visible Then cbo备注.SetFocus
        Exit Function
    End If
    
    On Error GoTo errHandle
    
    strSQL = "zl_病人挂号记录_出诊_DELETE('" & mstrNoIn & "','" & UserInfo.编号 & "','" & UserInfo.姓名 & "','" & Me.cbo备注.Text & "')"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    CancelBespeakRegist = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub cmdOK_Click()
    If mbytMode = 3 And mbytInState = 1 Then
        '取消预约
        If CancelBespeakRegist = False Then Exit Sub
        mblnOk = True
        gblnOk = True: Unload Me
        Exit Sub
    End If
    Call SaveData
    If Trim(txtSN.Text) <> "" Then Call mobjRegist.zlCancelRegNo(mlng锁号记录ID)
End Sub

Private Sub cmdPatiPic_Click()
    '74430,冉俊明,2014-7-8,挂号界面显示病人照片的浮动窗体
    Call ShowPatiPic
End Sub

Private Sub cmdRemark_Click()
    If SelectMemo("") = False Then
        zlCommFun.PressKey vbKeyTab: Exit Sub
    End If
End Sub
Private Sub cmdYb_Click()
     '医保身份证验证
     Call zlInusreIdentify
End Sub
Private Sub cmd结束挂号_Click()
    Call SaveData(True)
End Sub
 
Private Sub dtpAppointmentTime_Change()
    Dim str日期 As String, i As Integer, lngRow As Long
    If Not dtpAppointmentTime.Visible Then Exit Sub
    If Not dtpAppointmentTime.Enabled Then Exit Sub
    If dtpAppointmentDate.Visible Then
        str日期 = Format(dtpAppointmentDate.Value, "yyyy-MM-dd")
    Else
        str日期 = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    End If
    If str日期 = "" Then str日期 = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    
    txt发生时间.Text = str日期 & " " & Format(dtpAppointmentTime.Value, "hh:mm:00")
    lngRow = 0
    If CDate(txt发生时间.Text) > CDate(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("结束时间"))) Then
        '超出时间的安排，重新寻找定位
        For i = 1 To vsfPlan.Rows - 1
            With vsfPlan
                If .TextMatrix(.Row, .ColIndex("号别")) = .TextMatrix(i, .ColIndex("号别")) And _
                    CDate(vsfPlan.TextMatrix(i, vsfPlan.ColIndex("结束时间"))) >= CDate(txt发生时间.Text) Then
                    lngRow = i
                    Exit For
                End If
            End With
        Next i
    ElseIf CDate(txt发生时间.Text) < CDate(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("挂号时间"))) Then
        '超出时间的安排，重新寻找定位
        For i = 1 To vsfPlan.Rows - 1
            With vsfPlan
                If .TextMatrix(.Row, .ColIndex("号别")) = .TextMatrix(i, .ColIndex("号别")) And _
                    CDate(vsfPlan.TextMatrix(i, vsfPlan.ColIndex("挂号时间"))) <= CDate(txt发生时间.Text) Then
                    lngRow = i
                    Exit For
                End If
            End With
        Next i
    End If
    If lngRow <> 0 Then
        mblnAppointmentChange = True
        vsfPlan.Select lngRow, 1
        mblnAppointmentChange = False
    End If
End Sub
 

Private Sub dtpAppointmentTime_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
         DoEvents
       If txtPatient.Enabled Then
         txtPatient.SetFocus
       Else
           zlCommFun.PressKey vbKeyTab
       End If
    End If
End Sub

 

Private Sub IDKind_Click(objCard As zlIDKind.Card)
    Dim lng卡类别ID As Long, strOutCardNO As String, strExpand
    Dim strOutPatiInforXML As String
    If txtPatient.Locked Then Exit Sub
    
    mbln发卡 = True '问题号:56599
    If objCard.名称 Like "IC卡*" And objCard.系统 Then
        '系统IC卡
        If Not mobjICCard Is Nothing Then
           txtPatient.Text = mobjICCard.Read_Card()
           If txtPatient.Text <> "" Then
                mblnUnChange = True
                Call txtPatient_Validate(False)
                mblnUnChange = False
                Call SetOneCardBalance
           End If
        End If
        Exit Sub
    End If
    
    lng卡类别ID = objCard.接口序号
    
    If lng卡类别ID <= 0 Then Exit Sub
    '    zlReadCard(frmMain As Object, _
    '    ByVal lngModule As Long, _
    '    ByVal lngCardTypeID As Long, _
    '    ByVal blnOlnyCardNO As Boolean, _
    '    ByVal strExpand As String, _
    '    ByRef strOutCardNO As String, _
    '    ByRef strOutPatiInforXML As String) As Boolean
    '    '---------------------------------------------------------------------------------------------------------------------------------------------
    '    '功能:读卡接口
    '    '入参:frmMain-调用的父窗口
    '    '       lngModule-调用的模块号
    '    '       strExpand-扩展参数,暂无用
    '    '       blnOlnyCardNO-仅仅读取卡号
    '    '出参:strOutCardNO-返回的卡号
    '    '       strOutPatiInforXML-(病人信息返回.XML串)
    '    '返回:函数返回    True:调用成功,False:调用失败\
'    If gobjSquare.objSquareCard.zlReadCard(Me, mlngModul, lng卡类别ID, True, strExpand, strOutCardNO, strOutPatiInforXML) = False Then Exit Sub
'    txtPatient.Text = strOutCardNO
    
'    If txtPatient.Text <> "" Then
'        mblnUnChange = True
'        Call txtPatient_Validate(False)
'        mblnUnChange = False
'    End If
    
End Sub

Private Sub IDKind_ItemClick(index As Integer, objCard As zlIDKind.Card)
    If mbytInState > 0 Then Exit Sub
    '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
    txtPatient.IMEMode = 0
    '需要清除信息,避免刷卡后,再切换,造成密文显示失去意义
    If txtPatient.Text <> "" And Not mblnNotClick Then txtPatient.Text = ""
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    
    zlControl.TxtSelAll txtPatient
    '83089:李南春,2015/3/17,重置缺省的发卡类别
    If IDKind.GetCurCard.名称 Like "姓名*" Then
        Call InitSendCardPreperty(mlngModul)
    End If
End Sub

Private Sub IDKind_KeyDown(KeyCode As Integer, Shift As Integer)
    If txtPatient.Locked Or txtPatient.Enabled = False Or txtPatient.Text <> "" Then Exit Sub
    If IDKind.GetCurCard Is Nothing Then Exit Sub
    '快键操作IDKind
    IDKind.ActiveFastKey
End Sub

Private Sub IDKind_ReadCard(ByVal objCard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
    Dim lngPreIDKind As Long, intIndex As Integer
    Dim dtDate As Date
    Dim blnNew As Boolean
    Dim blnCard As Boolean    '是否就诊卡

    If txtPatient.Locked Or txtPatient.Text <> "" Then Exit Sub    'Or Not Me.ActiveControl Is txtPatient
    '状态变量赋值
    mblnNotClick = True
    intIndex = IDKind.GetKindIndex(objCard.名称)
    lngPreIDKind = IDKind.IDKind
    If intIndex > 0 Then IDKind.IDKind = intIndex
    txtPatient.Text = objPatiInfor.卡号
    Call txtPatient_Validate(False)
    
    If mrsInfo Is Nothing Then
        blnNew = True
    ElseIf mrsInfo.State <> 1 Then
        blnNew = True
    End If
    '当成新病人
    If (txtPatient.Text = "" Or blnNew) And objPatiInfor.姓名 <> "" Then
        txtPatient.Text = objPatiInfor.姓名
        intIndex = IDKind.GetKindIndex("姓名")
        If intIndex > 0 Then IDKind.IDKind = IDKind.GetKindIndex("姓名")
        Call txtPatient_KeyPress(vbKeyReturn)
        If txtPatient.Text <> "" Then
            Call zlControl.CboLocate(cbo性别, objPatiInfor.性别)
            If IsDate(objPatiInfor.出生日期) = False Then
                txt年龄.Text = ReCalcOld(CDate(objPatiInfor.出生日期), cbo年龄单位)
            End If
        End If
    End If
    IDKind.IDKind = lngPreIDKind
    mblnNotClick = False
End Sub

Private Sub MovePatiPic()
    '----------------------------------------------------------------------------------------------------------------
    '功能：移动病人相框
    '编制：冉俊明
    '日期：2014-7-8
    '----------------------------------------------------------------------------------------------------------------
    ReleaseCapture
    SendMessage picPatiPicBack.Hwnd, WM_SYSCOMMAND, SC_MOVE + HTCAPTION, 0
    
    If picPatiPicBack.Left < 0 Then picPatiPicBack.Left = 0
    If picPatiPicBack.Top < 0 Then picPatiPicBack.Top = 0
    If picPatiPicBack.Left + picPatiPicBack.Width > Me.ScaleWidth Then
        picPatiPicBack.Left = Me.ScaleWidth - picPatiPicBack.Width
    End If
    If picPatiPicBack.Top + picPatiPicBack.Height > Me.ScaleHeight Then
        picPatiPicBack.Top = Me.ScaleHeight - picPatiPicBack.Height
    End If
End Sub

Private Sub IDKind证件_ItemClick(index As Integer, objCard As zlIDKind.Card)
    Dim blnVisible As Boolean, lngRow As Long, lngCol As Long
    If IDKind证件.IDKind = IDKind证件.GetKindIndex("身份证号") Then blnVisible = True
    If blnVisible And txtPatient = "" Then txtIDCard.Tag = "": txtIDCard.Text = ""
    txtIDCard.Visible = blnVisible:  txt证件.Visible = Not blnVisible
    If txtIDCard.Visible And txtIDCard.Enabled Then txtIDCard.SetFocus
    If txt证件.Visible And txt证件.Enabled Then txt证件.SetFocus
    txt证件.Text = "": txt证件.Tag = ""
    If blnVisible Then Exit Sub
    '105357:李南春，2017/2/6，界面初始化时会触发ItemClick
    If mobjfrmPatiInfo Is Nothing Then Exit Sub
    With mobjfrmPatiInfo.vsCertificate
        For lngRow = 1 To .Rows - 1
            For lngCol = 0 To .Cols - 1 Step 2
                If .TextMatrix(lngRow, lngCol) = IDKind证件.GetCurCard.名称 Then
                    txt证件.Tag = .TextMatrix(lngRow, lngCol + 1)
                    txt证件.Text = txt证件.Tag
                    Exit For
                End If
            Next
        Next
    End With
End Sub

Private Sub imgPatiPic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MovePatiPic
End Sub

Private Sub lblClosePic_Click()
    picPatiPicBack.Visible = False
End Sub

Private Sub lblShow_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MovePatiPic
End Sub

'72168,冉俊明,2014/4/22,挂号时通过挂号科室确定可选费别
Private Sub mobjfrmPatiInfo_ReturnVisitClick()
    Dim i As Long
    
    Call Init费别(mobjfrmPatiInfo.chk复诊.Value = 0, True)
    With mobjfrmPatiInfo
        .cbo费别.Clear
        For i = 0 To cbo费别.ListCount - 1
            .cbo费别.AddItem cbo费别.List(i)
            .cbo费别.ItemData(i) = cbo费别.ItemData(i)
        Next
        .cbo费别.ListIndex = cbo费别.ListIndex
    End With
End Sub

Private Sub mobjfrmPatiInfo_PatiMerged(病人ID As Long)
        '合并后的病人
        Call GetPatient(IDKind.GetCurCard, "-" & 病人ID, False)
End Sub

Private Sub mobjfrmPatiInfo_付款方式Click(index As Long)
    cbo付款方式.ListIndex = index
End Sub

Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, _
                            ByVal strNation As String, ByVal datBirthDay As Date, ByVal strAddress As String)
    Dim lngPreIDKind As Long
    Dim blnNewCard   As Boolean
    Dim blnAddCardItem  As Boolean
    
    If txt号别.Text <> "" And Not txtPatient.Locked And txtPatient.Text = "" And Me.ActiveControl Is txtPatient Then
        mblnNotClick = True
        lngPreIDKind = IDKind.IDKind
        IDKind.IDKind = IDKind.GetKindIndex("身份证号")
        txtPatient.Text = strID
        mblnNotEMPIQuery = True
        mblnUserCancel = False
        Call txtPatient_Validate(False)
        mblnNotEMPIQuery = False
        '107049:李南春,2017/4/14,如果his有记录，将his信息传给接口
        If Not mrsInfo Is Nothing Then Call zlQueryEMPIPatiInfo
        
        If txtPatient.Text = "" And mblnUserCancel = True Then mblnNotClick = False: Exit Sub
        
        If txtPatient.Text = "" Then   '新病人
            IDKind.IDKind = IDKind.GetKindIndex("姓名")
            txtPatient.Text = strName
            '107049:李南春,2017/4/14,为了将身份证上的信息传给接口
            mblnNotEMPIQuery = True
            Call txtPatient_Validate(False)
            If txtPatient.Text <> "" Then
                txtIDCard.Text = strID
                txtIDCard.Tag = strID
                With mobjfrmPatiInfo
                    .txt身份证号.Text = strID
                    Call zlControl.CboLocate(.cbo性别, strSex)
                    Call zlControl.CboLocate(.cbo民族, strNation)
                    .txt出生日期.Text = Format(datBirthDay, "yyyy-MM-dd")
                    .txt出生时间.Text = "00:00"
                    txt出生日期.Text = Format(datBirthDay, "yyyy-MM-dd")
                    txt出生时间.Text = "00:00"
                    .cbo家庭地址.Text = IIf(Trim(cbo家庭地址.Text) = "", strAddress, cbo家庭地址.Text)
                    .txtRegLocation.Text = strAddress
                    cbo户口地址.Text = .txtRegLocation.Text
                    cbo性别.ListIndex = .cbo性别.ListIndex
                    txt年龄.Text = .txt年龄.Text
                    txt年龄.Tag = .txt年龄.Text '38564
                    
                    cbo年龄单位.ListIndex = .cbo年龄单位.ListIndex
                    Call txt年龄_Validate(False)
                    cbo家庭地址.Text = .cbo家庭地址.Text
                    '89242:李南春,2015/12/7,读取病人地址信息
                    padd家庭地址.Value = cbo家庭地址.Text
                    padd户口地址.Value = cbo户口地址.Text
                    .padd家庭地址.Value = cbo家庭地址.Text
                    .padd户口地址.Value = cbo户口地址.Text
                    .cbo年龄单位.Tag = .cbo年龄单位.Text
                    cbo年龄单位.Tag = cbo年龄单位.Text
                End With
            End If
            mblnNotEMPIQuery = False
            Call zlQueryEMPIPatiInfo
        End If
        IDKind.IDKind = lngPreIDKind
        mblnNotClick = False
        mobjfrmPatiInfo.mblnNewPatient = False
        '75717,冉俊明,2014-7-22,挂号预约时读取新病人身份证照片
        If mobjfrmPatiInfo.imgPatient.Picture = 0 Then
            Call LoadIDImage
        End If
        If cbo户口地址.Text = "" Then
            mobjfrmPatiInfo.txtRegLocation.Text = strAddress
            cbo户口地址.Text = strAddress
            padd户口地址.Value = cbo户口地址.Text
            mobjfrmPatiInfo.padd户口地址.Value = cbo户口地址.Text
        Else
            If mblnStructAdress Then
                If padd户口地址.CheckDefrentValue(padd户口地址.Value, strAddress) = False Then
                    If MsgBox("身份证上的地址" & strAddress & "与原有病人的户口地址" & padd户口地址.Value & "不一致,是否将病人的户口地址更新为身份证上的地址?", vbQuestion + vbYesNo, gstrSysName) = vbYes Then
                        mobjfrmPatiInfo.txtRegLocation.Text = strAddress
                        cbo户口地址.Text = strAddress
                        padd户口地址.Value = cbo户口地址.Text
                        mobjfrmPatiInfo.padd户口地址.Value = cbo户口地址.Text
                    End If
                End If
            Else
                If cbo户口地址.Text <> strAddress Then
                    If MsgBox("身份证上的地址" & strAddress & "与原有病人的户口地址" & cbo户口地址.Text & "不一致,是否将病人的户口地址更新为身份证上的地址?", vbQuestion + vbYesNo, gstrSysName) = vbYes Then
                        mobjfrmPatiInfo.txtRegLocation.Text = strAddress
                        cbo户口地址.Text = strAddress
                        padd户口地址.Value = cbo户口地址.Text
                        mobjfrmPatiInfo.padd户口地址.Value = cbo户口地址.Text
                    End If
                End If
            End If
        End If
        '没有家庭地址的,更新家庭地址
        If cbo家庭地址.Text = "" Then
            mobjfrmPatiInfo.cbo家庭地址.Text = strAddress
            cbo家庭地址.Text = strAddress
            padd家庭地址.Value = cbo家庭地址.Text
            mobjfrmPatiInfo.padd家庭地址.Value = cbo家庭地址.Text
        End If
    End If
End Sub


Private Sub mobjICCard_ShowICCardInfo(ByVal strNO As String)
    Dim lngPreIDKind As Long
    
    If txt号别.Text <> "" And Not txtPatient.Locked And txtPatient.Text = "" And Me.ActiveControl Is txtPatient Then
        mblnNotClick = True
        lngPreIDKind = IDKind.IDKind
        IDKind.IDKind = IDKind.GetKindIndex("IC卡号")
        txtPatient.Text = strNO
        If txtPatient.Text <> "" Then
            mblnUnChange = True
            Call txtPatient_Validate(False)
            mblnUnChange = False
            Call SetOneCardBalance
        Else
            Call mobjICCard.SetEnabled(False) '如果不符合发卡条件，禁用继续自动读取
        End If
        IDKind.IDKind = lngPreIDKind
        mblnNotClick = False
        If mobjICCard Is Nothing Then Call NewCardObject
        If txt号别.Text <> "" And Not txtPatient.Locked And Me.ActiveControl Is txtPatient Then mobjICCard.SetEnabled (txtPatient.Text = "")
    End If
End Sub

Private Sub cbo费别_Click()
    Dim str费别 As String
    
    If mbytInState = 1 Or Not Visible Then Exit Sub
    '31182:包含预约接收
    If (mbytMode = 2 Or mbytMode = 0 And mbytInState = 0 And mstrNoIn <> "") And InStr(1, mstrPrivs, ";允许修改费别;") <= 0 Then
        '预约接收
        If mTy_Para.bln预约接收确定挂号费 = False Then
            If Not mrsInfo Is Nothing Then
                Exit Sub
            End If
        End If
    End If
   ' If Not (mbytMode = 2 Or mbytMode = 0 And mbytInState = 0 And mstrNoIn <> "") And mTy_Para.bln预约接收确定挂号费 = False And Not (mrsInfo Is Nothing And mbytMode = 2) Then Exit Sub
    
    str费别 = NeedName(cbo费别)
    If mstrPre费别 = str费别 Then Exit Sub
    mstrPre费别 = str费别
    
    If txt号别.Text <> "" Then
        mblnBuyHisBook = True
        Call ShowRegistFromInput
        mblnBuyHisBook = False
    End If
End Sub



Private Sub cbo年龄单位_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtPatientPrint.Visible Then
            If cmdOK.Enabled And cmdOK.Visible Then cmdOK.SetFocus
        Else
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    End If
End Sub

Private Sub cbo医疗类别_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If cbo医疗类别.Locked Then Exit Sub
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        lngIdx = zlControl.CboMatchIndex(cbo医疗类别.Hwnd, KeyAscii)
        If lngIdx = -1 And cbo医疗类别.ListCount > 0 Then lngIdx = 0
        cbo医疗类别.ListIndex = lngIdx
    End If
End Sub

Private Sub cbo医生_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub cbo医生_Validate(Cancel As Boolean)
    Dim i As Integer
    Dim strDoctor As String
    Dim blnFinded As Boolean
    
    If cbo医生.ListCount = 0 Then cbo医生.Text = "": Exit Sub
    
    strDoctor = cbo医生.Text
    
    If mrsDoctor.State = 1 Then
        If mrsDoctor.RecordCount = 0 Then cbo医生.Text = "": Exit Sub
        mrsDoctor.MoveFirst
        For i = 1 To mrsDoctor.RecordCount
            If strDoctor = mrsDoctor!编号 Or strDoctor = mrsDoctor!姓名 Or UCase(strDoctor) = mrsDoctor!简码 Or strDoctor = mrsDoctor!简码 & "-" & mrsDoctor!姓名 Then
                strDoctor = mrsDoctor!ID
                blnFinded = True
                Exit For
            End If
            mrsDoctor.MoveNext
        Next
        If Not blnFinded Then Call zlCommFun.PressKey(vbKeyF4)
    End If
        
    If blnFinded Then
        If zlControl.CboLocate(cbo医生, strDoctor, True) Then
            mstr医生姓名 = Mid(cbo医生.Text, InStr(1, cbo医生.Text, "-") + 1)
            mlng医生ID = cbo医生.ItemData(cbo医生.ListIndex)
        Else
            Call zlControl.TxtSelAll(cbo医生)
            Cancel = True
        End If
    Else
        Call zlControl.TxtSelAll(cbo医生)
        Cancel = mrsDoctor.State = 1
    End If
End Sub

Private Sub chkShowAll_Click()
    If mblnUnChkClick = True Or mblnReadBooking Then Exit Sub
    Call ShowPlans
End Sub

Private Sub chk病历费_GotFocus()
    chk病历费.ForeColor = vbBlue
End Sub

Private Sub chk病历费_LostFocus()
    chk病历费.ForeColor = &H80000012
End Sub

Private Sub SetCHKState(chkThis As CheckBox)
    If chkThis Is chkPrint Then
        chkBooking.Enabled = chkPrint.Value = 0
        chkCancel.Enabled = chkPrint.Value = 0
        cmdComminuty.Enabled = chkPrint.Value = 0
    ElseIf chkThis Is chkBooking Then
        chkPrint.Enabled = chkBooking.Value = 0
        chkCancel.Enabled = chkBooking.Value = 0
    ElseIf chkThis Is chkCancel Then
        chkPrint.Enabled = chkCancel.Value = 0
        chkBooking.Enabled = chkCancel.Value = 0
        cmdComminuty.Enabled = chkCancel.Value = 0
        cmdYb.Enabled = chkCancel.Value = 0
    End If
End Sub

Private Sub SetCodeEnable(ByVal blnEnable As Boolean)
    txt号别.Enabled = blnEnable
    txt科室.Enabled = blnEnable
    txtSN.Enabled = blnEnable
    cbo医生.Enabled = blnEnable
End Sub

Private Sub SetPatiEnable(ByVal blnEnable As Boolean)
    IDKind.Enabled = blnEnable
    txtPatient.Enabled = blnEnable
    cmdLookup.Enabled = blnEnable
    cmdCard.Enabled = blnEnable
    cmdMore.Enabled = blnEnable
    cmdComminuty.Enabled = blnEnable
    cmdYb.Enabled = blnEnable
    cbo性别.Enabled = blnEnable
    txt年龄.Enabled = blnEnable
    cbo年龄单位.Enabled = blnEnable
    txt出生日期.Enabled = blnEnable
    txt出生时间.Enabled = blnEnable
    txtIDCard.Enabled = blnEnable
    txt家庭电话.Enabled = blnEnable
    cbo家庭地址.Enabled = blnEnable
    cbo户口地址.Enabled = blnEnable
    padd户口地址.Enabled = blnEnable
    padd家庭地址.Enabled = blnEnable
    cbo医疗类别.Enabled = blnEnable
    cbo费别.Enabled = blnEnable
    cbo付款方式.Enabled = blnEnable
    txt门诊号.Enabled = blnEnable
    IDKind证件.Enabled = blnEnable
End Sub

Private Sub chkCancel_Click()
    cboNO.Text = ""
    
    SetCodeEnable chkCancel.Value = 0
    SetPatiEnable chkCancel.Value = 0
    vsfPlan.Enabled = chkCancel.Value = 0
    
    Call RemoveShowItem
    Call ClearBill
    
    mcur合计 = 0: mcur应缴 = 0: txt合计.Text = "0.00": txt本次应缴.Text = "0.00": mint挂号数 = 0
    txt缴款.Text = "0.00": txt缴款.Enabled = chkCancel.Value = 0
    txt找补.Text = "0.00": txt找补.Enabled = chkCancel.Value = 0
        
    Call SetCHKState(chkCancel)
    
    If chkCancel.Value = 0 Then
        chkCancel.ForeColor = 0
        lbl急.Visible = False
        txtFact.Locked = False
        txt号别.Locked = False
        
        txtPatient.Locked = False
        txt年龄.Locked = False
        cbo家庭地址.Locked = False
        cbo户口地址.Locked = False
        padd家庭地址.ControlLock = False
        padd户口地址.ControlLock = False
        txt门诊号.Locked = False
        
        cbo性别.Locked = False
        cbo付款方式.Locked = False
        cbo费别.Locked = False
        txtIDCard.Locked = False
        cbo结算方式.Locked = False
        
        chk病历费.Enabled = mbln病历费
        chk病历费.Caption = "购买病历"
        chkExtra.Visible = False
        lbl预约方式.Visible = mbytMode <> 0
        cbo预约方式.Visible = mbytMode <> 0
        '刷新票据号
        If mbytMode <> 1 And gbytInvoice <> 0 Then Call RefreshFact
        If mbytMode <> 1 Then Load支付方式
    Else
        chkCancel.ForeColor = vbRed
        
        lbl急.Visible = False
                
        txtFact.Locked = Not (InStr(1, mstrPrivs, ";修改票据号;") > 0) And gblnBill挂号  ' True:刘兴洪:20000,增加修改票据号权限
        txt号别.Locked = True
        
        txtPatient.Locked = True
        txt年龄.Locked = True
        cbo家庭地址.Locked = True
        cbo户口地址.Locked = True
        padd家庭地址.ControlLock = True
        padd户口地址.ControlLock = True
        txt门诊号.Locked = True
        txtIDCard.Locked = True
        cbo性别.Locked = True
        cbo付款方式.Locked = True
        cbo费别.Locked = True
        cbo结算方式.Visible = False
        
        chk病历费.Enabled = False
        chk病历费.Caption = "退病历费"
        cboNO.Text = "": txtFact.Text = ""
        If cboNO.Enabled And cboNO.Visible Then cboNO.SetFocus
    End If
    Call SetUndisplayBalance
    Call AdjustInfoPosition
End Sub

Private Sub chkPrint_Click()
    SetCodeEnable chkPrint.Value = 0
    SetPatiEnable chkPrint.Value = 0
    vsfPlan.Enabled = chkPrint.Value = 0
    
    cboNO.Text = ""
    chkExtra.Visible = False
    Call RemoveShowItem
    Call ClearBill
    
    mcur合计 = 0: mcur应缴 = 0: txt合计.Text = "0.00": txt本次应缴.Text = "0.00": mint挂号数 = 0
    txt缴款.Text = "0.00": txt缴款.Enabled = chkPrint.Value = 0
    txt找补.Text = "0.00": txt找补.Enabled = chkPrint.Value = 0
        
    Call SetCHKState(chkPrint)
    
    If txtPatientPrint.Visible Then
        txtPatientPrint.Text = ""
        txtPatientPrint.Visible = False
        txtPatientPrint.Locked = False
        Call SetRePrintPatiEnabled(True)
    End If
    
    If chkPrint.Value = 0 Then
        chkPrint.ForeColor = 0
                                
        lbl急.Visible = False
        
        txtFact.Locked = False
        txt号别.Locked = False
        
        txtPatient.Locked = False
        txt年龄.Locked = False
        cbo家庭地址.Locked = False
        cbo户口地址.Locked = False
        padd家庭地址.ControlLock = False
        padd户口地址.ControlLock = False
        txt门诊号.Locked = False
        cbo性别.Locked = False
        cbo付款方式.Locked = False
        cbo费别.Locked = False
        cbo结算方式.Locked = False
        
        chk病历费.Enabled = mbln病历费
        '74017:李南春，2014-6-17，退出挂号重打时，恢复cmdCard的状态
        cmdCard.Enabled = True
        '刷新票据号
        If mbytMode <> 1 And gbytInvoice <> 0 Then Call RefreshFact
    Else
        chkPrint.ForeColor = vbBlue
                
        lbl急.Visible = False
                
        txtFact.Locked = Not (InStr(1, mstrPrivs, ";修改票据号;") > 0) And gblnBill挂号  'True:刘兴洪:20000,增加修改票据号权限
        txt号别.Locked = True
        
        If InStr(1, mstrPrivs, ";修改姓名重打;") > 0 Then
            txtPatientPrint.Width = txtPatient.Width
            txtPatientPrint.Visible = True
        End If
        
        txtPatient.Locked = True
        txt年龄.Locked = True
        cbo家庭地址.Locked = True
        cbo户口地址.Locked = True
        padd家庭地址.ControlLock = True
        padd户口地址.ControlLock = True
        txt门诊号.Locked = True
        cbo性别.Locked = True
        cbo付款方式.Locked = True
        cbo费别.Locked = True
        cbo结算方式.Locked = True
        
        chk病历费.Enabled = False
                
        cboNO.Text = "": txtFact.Text = ""
        
        If cboNO.Enabled And cboNO.Visible Then cboNO.SetFocus
    End If
    Call AdjustInfoPosition
End Sub

Private Sub chk病历费_Click()
    If Not mbln病历费 And mbytInState = 0 Then
        chk病历费.Value = 0: Exit Sub
    End If
    
    '退号
    If mbytInState = 1 And mbytMode = 4 Or chkCancel.Value = 1 Then
        If mblnNotClick Then Exit Sub
        Call IsCheckBackExtra(True)
        Exit Sub
    End If
    '31182:
    If (mbytMode = 0 And mbytInState = 0 And mstrNoIn <> "") And mTy_Para.bln预约接收确定挂号费 = False Then Exit Sub
    
    If Not chk病历费.Enabled Then Exit Sub
    If mblnNotClick Then Exit Sub
    mblnBuyHisBook = True
    Call ShowRegistFromInput
    mblnBuyHisBook = False
End Sub

Private Sub chkExtra_Click()
    If Not mbln病历费 And mbytInState = 0 Then
        chk病历费.Value = 0: Exit Sub
    End If
    
    '退号
    If Not (mbytInState = 1 And mbytMode = 4 Or chkCancel.Value = 1) Then Exit Sub
    If mblnNotClick Then Exit Sub
    Call IsCheckBackExtra
End Sub

Private Function IsCheckBackExtra(Optional ByVal bln病历费 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:退号时检查附费项目是否允许分开退
    '入参:bln病历费-检查病历费
    '返回:成功返回true,否则返回False
    '编制:李南春
    '日期:2018/5/2 11:35:08
    '问题:123874
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strFeeItem As String
    Dim curMoney As Currency, curTotal As Currency, curDiff As Currency
    Dim curAdvance As Currency '预交的缴款
    Dim curInsure As Currency
    Dim curCash As Currency
    Dim i As Long
    Dim strFilter As String
    Dim strItem() As String
    If Not (mbytInState = 1 And mbytMode = 4 Or chkCancel.Value = 1) Then IsCheckBackExtra = True: Exit Function
    strFeeItem = IIf(bln病历费, "病历费", "附加费")
    If Not mrsBillAdvance Is Nothing Then
        mrsBillAdvance.Filter = 0
        If mrsBillAdvance.RecordCount > 0 Then mrsBillAdvance.MoveFirst
        Do While Not mrsBillAdvance.EOF
            If InStr(",7,8,", "," & mrsBillAdvance!性质 & ",") > 0 And (mrsBillAdvance!记录性质 <> 1 And mrsBillAdvance!记录性质 <> 11) Then
                MsgBox "使用三方接口结算的挂号单据,不能将" & strFeeItem & "与挂号费分开退!", vbInformation, gstrSysName
                mblnNotClick = True
                If bln病历费 Then
                    chk病历费.Value = 1
                Else
                    chkExtra.Value = 1
                End If
                mblnNotClick = False
                Exit Function
            End If
            If InStr(",3,", "," & mrsBillAdvance!性质 & ",") > 0 And (MCPAR.不收病历费 = False Or Not bln病历费) Then
                MsgBox "医保个人账户收取" & strFeeItem & "时,不支持" & strFeeItem & "与挂号费分别退!", vbInformation, gstrSysName
                mblnNotClick = True
                If bln病历费 Then
                    chk病历费.Value = 1
                Else
                    chkExtra.Value = 1
                End If
                mblnNotClick = False
                Exit Function
            End If
            mrsBillAdvance.MoveNext
        Loop
    End If
    '如果此时触发了事件,表示本事
    If mrsBill Is Nothing Then IsCheckBackExtra = True: Exit Function
    If mstr附加项目ID <> "" Then
        strFilter = ""
        strItem = Split(mstr附加项目ID, ",")
        For i = 0 To UBound(strItem)
            If strFilter = "" Then
                strFilter = "收费细目ID <> " & strItem(i)
            Else
                strFilter = strFilter & " And 收费细目ID <> " & strItem(i)
            End If
        Next i
    End If
    
    '先取出总金额
    mrsBill.Filter = 0
    If mrsBill.RecordCount > 0 Then mrsBill.MoveFirst
    For i = 1 To mrsBill.RecordCount
        curTotal = curTotal + mrsBill!实收
        mrsBill.MoveNext
    Next
    
    '再取勾选后的金额和项目.有可能是恢复,但不影响
    If chkExtra.Value = 0 And strFilter <> "" Then
        If chk病历费.Value = 1 Then
            mrsBill.Filter = strFilter
        Else
            mrsBill.Filter = "附加标志<>1 And " & strFilter
        End If
    Else
        If chk病历费.Value = 1 Then
            mrsBill.Filter = 0
        Else
            mrsBill.Filter = "附加标志<>1"
        End If
    End If
    If mrsBill.RecordCount > 0 Then mrsBill.MoveFirst
    vsfMoney.Rows = mrsBill.RecordCount + 1
    For i = 1 To mrsBill.RecordCount
        vsfMoney.TextMatrix(i, 0) = mrsBill!项目
        vsfMoney.TextMatrix(i, 1) = Format(mrsBill!应收, "0.00")
        vsfMoney.TextMatrix(i, 2) = Format(mrsBill!实收, "0.00")
        curMoney = curMoney + mrsBill!实收
        mrsBill.MoveNext
    Next
    txt合计.Text = Format(curMoney, "0.00")
    mrsBill.Filter = 0: If mrsBill.RecordCount > 0 Then mrsBill.MoveFirst
    
    '取差额,然后再调用
    curDiff = curTotal - curMoney
    Call Load结算信息(Val(curMoney), Val(curDiff))
    Set连续挂号
    IsCheckBackExtra = True
End Function

Private Sub RecalPay()
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim dblTotal As Double, rsTx As ADODB.Recordset
    On Error GoTo errH
    vsfPay.Rows = 1
    vsfPay.Clear 1
    vsfPay.Rows = 2
    dblTotal = Val(txt合计.Text)
    strSQL = "Select b.名称 As 结算方式, a.冲预交, b.性质" & vbNewLine & _
            "From 病人预交记录 A, 结算方式 B" & vbNewLine & _
            "Where a.结帐id = [1] And a.结算方式 = b.名称 And a.记录性质 = 4" & vbNewLine & _
            "Union" & vbNewLine & _
            "Select '预交金' As 结算方式, a.冲预交, 0 As 性质 From 病人预交记录 A Where a.结帐id = [1] And Mod(a.记录性质, 10) = 1 Order By 性质"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng结帐ID)
    rsTemp.Filter = "性质 <> 1 And 性质 <> 2"
    Do While Not rsTemp.EOF
        If dblTotal > 0 Then
            If dblTotal > Val(rsTemp!冲预交) Then
                vsfPay.TextMatrix(vsfPay.Rows - 1, 0) = rsTemp!结算方式
                vsfPay.TextMatrix(vsfPay.Rows - 1, 1) = Format(Val(rsTemp!冲预交), "0.00")
                vsfPay.RowData(vsfPay.Rows - 1) = Val(rsTemp!性质)
                If Val(rsTemp!性质) = 7 Or Val(rsTemp!性质) = 8 Then
                    strSQL = "Select ID,是否退现 From 医疗卡类别 Where 结算方式=[1]"
                    Set rsTx = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, rsTemp!结算方式)
                    If rsTx.EOF Then
                        strSQL = "Select 编号,是否退现 From 消费卡类别目录 Where 结算方式=[1]"
                        Set rsTx = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, rsTemp!结算方式)
                        If rsTx.EOF Then
                            vsfPay.TextMatrix(vsfPay.Rows - 1, vsfPay.ColIndex("修改")) = "1"
                        Else
                            vsfPay.TextMatrix(vsfPay.Rows - 1, 4) = Nvl(rsTx!编号)
                            vsfPay.TextMatrix(vsfPay.Rows - 1, vsfPay.ColIndex("修改")) = IIf(Val(rsTx!是否退现) = 1, "0", "1")
                        End If
                    Else
                        vsfPay.TextMatrix(vsfPay.Rows - 1, 4) = Nvl(rsTx!ID)
                        vsfPay.TextMatrix(vsfPay.Rows - 1, vsfPay.ColIndex("修改")) = IIf(Val(rsTx!是否退现) = 1, "0", "1")
                    End If
                End If
                If Val(rsTemp!性质) = 0 Or Val(rsTemp!性质) = 3 Then
                    vsfPay.TextMatrix(vsfPay.Rows - 1, vsfPay.ColIndex("修改")) = "1"
                End If
                vsfPay.Rows = vsfPay.Rows + 1
                dblTotal = dblTotal - Val(rsTemp!冲预交)
            Else
                vsfPay.TextMatrix(vsfPay.Rows - 1, 0) = rsTemp!结算方式
                vsfPay.TextMatrix(vsfPay.Rows - 1, 1) = Format(dblTotal, "0.00")
                vsfPay.RowData(vsfPay.Rows - 1) = Val(rsTemp!性质)
                If Val(rsTemp!性质) = 7 Or Val(rsTemp!性质) = 8 Then
                    strSQL = "Select ID,是否退现 From 医疗卡类别 Where 结算方式=[1]"
                    Set rsTx = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, rsTemp!结算方式)
                    If rsTx.EOF Then
                        strSQL = "Select 编号,是否退现 From 消费卡类别目录 Where 结算方式=[1]"
                        Set rsTx = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, rsTemp!结算方式)
                        If rsTx.EOF Then
                            vsfPay.TextMatrix(vsfPay.Rows - 1, vsfPay.ColIndex("修改")) = "1"
                        Else
                            vsfPay.TextMatrix(vsfPay.Rows - 1, 4) = Nvl(rsTx!编号)
                            vsfPay.TextMatrix(vsfPay.Rows - 1, vsfPay.ColIndex("修改")) = IIf(Val(rsTx!是否退现) = 1, "0", "1")
                        End If
                    Else
                        vsfPay.TextMatrix(vsfPay.Rows - 1, 4) = Nvl(rsTx!ID)
                        vsfPay.TextMatrix(vsfPay.Rows - 1, vsfPay.ColIndex("修改")) = IIf(Val(rsTx!是否退现) = 1, "0", "1")
                    End If
                End If
                If Val(rsTemp!性质) = 0 Or Val(rsTemp!性质) = 3 Then
                    vsfPay.TextMatrix(vsfPay.Rows - 1, vsfPay.ColIndex("修改")) = "1"
                End If
                vsfPay.Rows = vsfPay.Rows + 1
                dblTotal = 0
            End If
        End If
        rsTemp.MoveNext
    Loop
    rsTemp.Filter = "性质 = 1 Or 性质 = 2"
    rsTemp.Sort = "性质 Desc"
    Do While Not rsTemp.EOF
        If dblTotal > 0 Then
            If dblTotal > Val(rsTemp!冲预交) Then
                vsfPay.TextMatrix(vsfPay.Rows - 1, 0) = rsTemp!结算方式
                vsfPay.TextMatrix(vsfPay.Rows - 1, 1) = Format(Val(rsTemp!冲预交), "0.00")
                vsfPay.RowData(vsfPay.Rows - 1) = Val(rsTemp!性质)
                vsfPay.TextMatrix(vsfPay.Rows - 1, vsfPay.ColIndex("修改")) = IIf(Val(rsTemp!性质) = 1, "1", "0")
                vsfPay.Rows = vsfPay.Rows + 1
                dblTotal = dblTotal - Val(rsTemp!冲预交)
            Else
                vsfPay.TextMatrix(vsfPay.Rows - 1, 0) = rsTemp!结算方式
                vsfPay.TextMatrix(vsfPay.Rows - 1, 1) = Format(dblTotal, "0.00")
                vsfPay.RowData(vsfPay.Rows - 1) = Val(rsTemp!性质)
                vsfPay.TextMatrix(vsfPay.Rows - 1, vsfPay.ColIndex("修改")) = IIf(Val(rsTemp!性质) = 1, "1", "0")
                vsfPay.Rows = vsfPay.Rows + 1
                dblTotal = 0
            End If
        End If
        rsTemp.MoveNext
    Loop
    Exit Sub
errH:
    If ErrCenter = 1 Then Resume
End Sub
    
 
Private Sub chk病历费_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub cbo费别_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If cbo费别.Locked Then Exit Sub
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        lngIdx = zlControl.CboMatchIndex(cbo费别.Hwnd, KeyAscii)
        If lngIdx = -1 And cbo费别.ListCount > 0 Then lngIdx = 0
        cbo费别.ListIndex = lngIdx
    End If
End Sub

Private Sub cbo结算方式_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If cbo结算方式.Locked Then Exit Sub
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        lngIdx = zlControl.CboMatchIndex(cbo结算方式.Hwnd, KeyAscii)
        If lngIdx = -1 And cbo结算方式.ListCount > 0 Then lngIdx = 0
        cbo结算方式.ListIndex = lngIdx
    End If
End Sub

Private Sub cbo性别_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    
    If cbo性别.Locked Then Exit Sub
    
    If KeyAscii = 13 And cbo性别.ListIndex <> -1 Then
        KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    End If
    
    Call SendMessage(cbo性别.Hwnd, CB_GETDROPPEDSTATE, 0, 0)
    lngIdx = MatchIndex(cbo性别.Hwnd, KeyAscii)
    If lngIdx <> -2 Then cbo性别.ListIndex = lngIdx
    If cbo性别.ListCount > 0 And cbo性别.ListIndex = -1 Then cbo性别.ListIndex = 0
End Sub

Private Sub cbo付款方式_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        If cbo付款方式.Locked Then Exit Sub
        
        lngIdx = zlControl.CboMatchIndex(cbo付款方式.Hwnd, KeyAscii)
        If lngIdx = -1 And cbo付款方式.ListCount > 0 Then lngIdx = 0
        cbo付款方式.ListIndex = lngIdx
    End If
End Sub

Private Sub cmdCancel_Click()
    mblnCancel = True
    If mbytInState > 1 And mbytMode = 1 Then
        Unload Me
        mblnCancel = False
        Exit Sub
    End If
    If mbytInState = 0 And (chkPrint.Value = 1 Or chkCancel.Value = 1 Or chkBooking.Value = 1) Then
        If chkPrint.Value = 1 Then
            chkPrint.Value = 0
        ElseIf chkCancel.Value = 1 Then
            chkCancel.Value = 0
        ElseIf chkBooking.Value = 1 Then
            chkBooking.Value = 0
        End If
    ElseIf mbytMode = 0 And mbytInState = 0 And mstrNoIn <> "" Then '接收预约
        Call ClearBill
        Call SetReceiveState(False)
        
    ElseIf mbytMode = 2 Or mbytInState = 1 Or (mbytInState = 0 And mrsItems Is Nothing) Then
        Unload Me
    Else
        mbln连续挂号 = False
        Call YBIdentifyCancel '取消医保病人身份验证
        Call ClearBill
        
        '刷新票据号
        If mbytMode <> 1 And gbytInvoice <> 0 Then Call RefreshFact
    End If
    mblnCancel = False
End Sub

Private Sub picDetailFee_Resize()
    Err = 0: On Error Resume Next
    With picDetailFee
         
        cbo备注.Top = .ScaleHeight - cbo备注.Height - 50
        lbl摘要.Top = cbo备注.Top + (cbo备注.Height - lbl摘要.Height) \ 2
        
        txt发生时间.Top = IIf(cbo备注.Visible, cbo备注.Top - 20, .ScaleHeight - txt发生时间.Height - 50) - txt发生时间.Height
        lbl发生时间.Top = txt发生时间.Top + (txt发生时间.Height - lbl发生时间.Height) \ 2
        
        cbo预约方式.Top = txt发生时间.Top
        lbl预约方式.Top = lbl发生时间.Top
        
        chkExtra.Top = txt发生时间.Top + (txt发生时间.Height - chkExtra.Height) \ 2
        chk病历费.Top = chkExtra.Top
        vsfMoney.Height = IIf(txt发生时间.Top - vsfMoney.Top - 20 < 0, 0, txt发生时间.Top - vsfMoney.Top - 20)
        
    End With
End Sub

Private Sub picInfo_Resize()
    Dim lntTop As Long
    
    On Error Resume Next
    With picInfo
        lntTop = IIf(stbThis.Visible, stbThis.Height, 0)
        lblPrompt.Top = .ScaleHeight - lntTop - lblPrompt.Height - 50
        vsfPay.Top = lblPrompt.Top - vsfPay.Height - 50
        If mbytMode = 1 And mbytInState = 0 Then
            lntTop = IIf(vsfPay.Visible, vsfPay.Top, lblPrompt.Top - picTotal.Height - 50) - IIf(stbThis.Visible, stbThis.Height, 0)
        Else
            lntTop = IIf(vsfPay.Visible, vsfPay.Top, lblPrompt.Top - picTotal.Height - 50)
        End If
        picTotal.Top = lntTop
        picBal.Top = lntTop
        picDetailFee.Left = vsfPay.Left
        
    End With
    picDetailFee.Height = lntTop - picDetailFee.Top - 20
End Sub
 



Private Sub ClearBill(Optional blnClearPati As Boolean = True, Optional blnClearFact As Boolean = True, Optional ByVal blnClearInsure As Boolean = True)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:清除单据信息
    '入参:blnClearPati-清除病人信息
    '     blnClearFact-清除发票信息
    '     blnClearInsure-清除医保信息
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2009-12-02 10:32:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnIDKind As Boolean, strTemp As String, i As Integer
    
    Call SetShowBalance '68991
    blnIDKind = mblnIDCardKind
    txtSN.Text = ""
    mstrNoIn = ""
    mlng记录ID = 0
    mlng锁号记录ID = 0
    If mbytMode <> 1 Then
        If chkShowAll.Value = 1 Then chkShowAll.Value = 0
    End If
    lbl急.Visible = False
    If blnClearFact Then txtFact.Text = ""
    mblnNoClearPrompt = False
    txt号别.Text = ""                       '调用Change事件加载号别列表
    txt科室.Text = ""
    cbo医生.Clear
    txtIDCard.Text = ""
    mblnAppointmentChange = False
    txt证件.Text = ""
    mstrForceNote = ""
    txt家庭电话.Text = ""
    If mlngOutModeMC > 0 Then cbo医疗类别.ListIndex = 0
    '69338,刘尔旋,挂号完成时未重置先诊疗后结算信息的问题
    mRegistFeeMode = EM_RG_现收
    mPatiChargeMode = EM_先结算后诊疗
    mlng挂号科室ID = 0
    mstr医生姓名 = ""
    mblnViewOriginal = False
    mlng医生ID = 0
    mbln建病案 = False
'    txt摘要.Text = ""
    cbo备注.Text = ""
    mstrPreNO = ""
    mintCancel = 0
    mbln附加费 = False
    mstrPrePriceGrade = ""
    
    txt号别.Locked = False
    txt号别.Enabled = True
    If mbytMode <> 2 Then cbo费别.Locked = False: cbo费别.TabStop = gbln费别
    
    mstr划价NO = ""
    If vsfMoney.Rows < 2 Then
        cmdOK.Visible = True
    Else
        If vsfMoney.RowData(1) = 0 Then
            cmdOK.Visible = True
        End If
    End If
    '问题号:58843
    If blnClearPati Then Set mrsInfo = Nothing '病人信息清空
    Set mobjDelCards = Nothing
    mstr病人家属IDs = ""
    
    Call SetPatiInfoEnabled(False, mrsInfo Is Nothing, Not blnClearPati) '根据参数,如果不要求输姓名,或者号别不建病案,则会清除病人姓名
    
    mblnIDCardKind = False
    
    If blnClearPati Then
        Call ClearPatientInfo
        Call Init费别(True, False)
        Call SetCboDefault(cbo费别)
        Call ClearmobjfrmPatiInfoFace
    Else
        '54537:刘尔旋,2014-02-27,医保病人费别未清空的问题
        If mintInsure <> 0 And mstrYBPati <> "" Then Call SetCboDefault(cbo费别)
        mblnICCard = False
        mblnAddCardItem = False
    End If
    
    If mblnNewCard Then
        mobjfrmPatiInfo.txt卡号 = ""
        mobjfrmPatiInfo.mstrCard = ""
        lblPrompt.Caption = ""
        gCurSendCard.lng收费细目ID = 0
        vsfPay.Height = 2220
        mblnNewCard = False
    End If
    
    '医保改动
    mlng结帐ID = 0
    
    If blnClearPati = False And blnClearInsure = False Then
        '医保病人,连接续挂号时有效
    Else
        mintInsure = 0
        mstrYBPati = ""
        txtPatient.ForeColor = Me.ForeColor
        mobjfrmPatiInfo.txtPatient.ForeColor = Me.ForeColor
        Call SetIdentifyLocked(False)
    End If
    
    cmdComminuty.Enabled = True
    mint社区 = 0
    mstr社区号 = ""
    
    Call ShowMedicareInfo(blnClearPati = False And blnClearInsure = False)
    
    '固定清除预交支付信息
    Call ShowDeposit(False)

    If mblnReSetIDKind And txtPatient.Text = "" Then IDKind.IDKind = IDKind.GetKindIndex("门诊号")
    If blnIDKind And txtPatient.Text = "" Then IDKind.IDKind = IDKind.GetKindIndex("身份证号")
    mblnReSetIDKind = False
    mstr门诊号 = "": txt门诊号.TabStop = True
    
    chk病历费.Enabled = False
    chk病历费.Value = 0
    chk病历费.Enabled = mbln病历费
    If blnClearPati And mbln病历费 Then
        If mbytMode = 0 Or mbytMode = 1 Then chk病历费.Value = IIf(zlDatabase.GetPara("默认购买病历", glngSys, mlngModul, 0) = "1", 1, 0)
    End If
    
'    txt摘要.Text = ""

    Call ClearCardMoney
    Set mrsItems = Nothing
    Set mrsInComes = Nothing
    Call ClearMoney
'    Call SetCboDefault(cbo结算方式)
    Call Load支付方式
    
    If cbo预约方式.Visible Then
        strTemp = zlDatabase.GetPara("缺省预约方式", glngSys, IIf(mblnStation, 1260, mlngModul), "")
        '问题号:112838,焦博,2017/09/05,基础字典表中未设置任何预约方式时会报错
        If cbo预约方式.ListCount <> 0 Then
            For i = 0 To cbo预约方式.ListCount - 1
                If Mid(cbo预约方式.List(i), InStr(cbo预约方式.List(i), ".") + 1) = strTemp Then
                    cbo预约方式.ListIndex = i
                End If
            Next i
            If cbo预约方式.ListIndex < 0 Then cbo预约方式.ListIndex = 0
        End If
    End If
    
    If mbytMode = 0 And mbytInState = 0 And gblnLED Then
        zl9LedVoice.DisplayPatient ""
    End If
    If txt号别.Enabled And txt号别.Visible Then txt号别.SetFocus
End Sub

Private Sub ClearCardMoney()
    With mCurCardPay
        .lng医疗卡类别ID = 0
        .bln消费卡 = False
        .str结算方式 = ""
        .str名称 = ""
        .str刷卡卡号 = ""
        .str刷卡密码 = ""
        .dbl帐户余额 = 0
        .Have挂号费 = False
        .Have卡费 = False
        Set .objCard = Nothing
    End With
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.Hwnd, Me.Name
End Sub

Private Function DefMainCommandBars() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化菜单及工具栏
    '返回:设置成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-01-25 15:29:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cbrControl As CommandBarControl, cbrSubControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup, cbrToolBar As CommandBar
    
    Err = 0: On Error GoTo errHandle
    Set cbsThis.Icons = zlCommFun.GetPubIcons
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    
    cbsThis.VisualTheme = xtpThemeOffice2003
    Set cbsThis.Options.Font = txtPatient.Font
    With cbsThis.Options
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
        .ShowExpandButtonAlways = False
    End With
    cbsThis.EnableCustomization False
    cbsThis.ActiveMenuBar.ModifyStyle &H400000, 0 '去除菜单栏前缀
    '菜单定义
    cbsThis.ActiveMenuBar.Visible = False
    cbsThis.ActiveMenuBar.EnableDocking xtpFlagAlignTop Or xtpFlagHideWrap Or xtpFlagStretched
    
    '工具栏定义
    Set mcbrToolBar = cbsThis.Add("工具栏", xtpBarTop)
    mcbrToolBar.ShowTextBelowIcons = False
    mcbrToolBar.ContextMenuPresent = False
    mcbrToolBar.EnableDocking xtpFlagStretched
    mcbrToolBar.ModifyStyle &H400000, 0 '去除菜单栏前缀
    With mcbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, 2605, "预留")
        Set cbrControl = .Add(xtpControlButton, 2604, "取消预留")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印设置")
        Set cbrControl = .Add(xtpControlButton, 3816, "缴预交")
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新")
        Set cbrControl = .Add(xtpControlButton, 816, "出诊表")
        Set cbrControl = .Add(xtpControlButton, 4006, "历史挂号信息")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
    
    End With
    For Each cbrControl In mcbrToolBar.Controls
        cbrControl.Style = xtpButtonIconAndCaption
    Next
    
    '快键绑定
    With cbsThis.KeyBindings

    End With
    
    '设置不常用菜单
    With cbsThis.Options

    End With
    
    DefMainCommandBars = True
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub InitPanel()
    Dim objPane As Pane
    
    Err = 0: On Error GoTo errHandle
    Set objPane = dkpMain.CreatePane(1, 600, 120, DockLeftOf, Nothing)
    objPane.Options = PaneNoCloseable Or PaneNoHideable Or PaneNoFloatable Or PaneNoCaption
    objPane.Handle = picPlan.Hwnd
    
    Set objPane = dkpMain.CreatePane(2, 700, 400, DockRightOf, Nothing)
    objPane.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoHideable Or PaneNoFloatable
    objPane.Handle = picInfoFrame.Hwnd
    objPane.MaxTrackSize.Width = 500
    objPane.MinTrackSize.Width = 500
    
    With dkpMain
        .SetCommandBars cbsThis
        .VisualTheme = ThemeOffice2003
        .Options.ThemedFloatingFrames = True
        .Options.UseSplitterTracker = True '实时拖动
        .Options.AlphaDockingContext = True
        .Options.HideClient = True
    End With
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub HoldRegNo()
    Dim lngSN        As Long
    Dim blnCan       As Boolean
    Dim strSQL       As String
    Dim datThis      As Date
    Dim datTime As Date
    
    If vsfList.Rows = 0 Or mViewMode = V_普通号分时段 Or vsfList.Visible = False Then Exit Sub
    If mViewMode <> v_专家号分时段 Then
        lngSN = Val(vsfList.TextMatrix(vsfList.Row, vsfList.Col))
    Else
        lngSN = Val(Get时段(vsfList.Row, vsfList.Col, False))
    End If
    If lngSN > 0 Then
        blnCan = True
        If Not mrsSNState Is Nothing Then
            mrsSNState.Filter = "序号=" & lngSN
            If mrsSNState.RecordCount = 0 Then
                blnCan = False
            Else
                blnCan = True
            End If
        End If
    End If
    
    On Error GoTo errH
    If blnCan Then
        If picBookingDate.Visible Then
            Select Case mViewMode
            Case V_普通号:
                datThis = dtpAppointmentDate.Value
            Case Else
                datThis = CDate(Format(dtpAppointmentDate.Value, "yyyy-MM-dd") & " " & Format(Get时段(vsfList.Row, vsfList.Col, True), "hh:mm:ss"))
            End Select
        Else
            If mViewMode <> v_专家号分时段 Then
                datThis = zlDatabase.Currentdate
            Else
                datThis = CDate(Format(zlDatabase.Currentdate, "yyyy-MM-dd") & " " & Format(Get时段(vsfList.Row, vsfList.Col, True), "hh:mm:ss"))
            End If
        End If
        If mViewMode <> v_专家号分时段 Then
            strSQL = "Zl_挂号序号状态_Update('" & vsfPlan.TextMatrix(vsfPlan.Row, GetCol("号别")) & _
                  "',To_Date('" & Format(datThis, "yyyy-MM-dd") & "','YYYY-MM-DD')," & lngSN & _
                  ",3,'" & UserInfo.姓名 & "'," & IIf(mcbrToolBar.Controls.Find(xtpControlButton, 2605).Enabled, "1", "0") & ",Null," & vsfPlan.TextMatrix(vsfPlan.Row, GetCol("记录ID")) & ")"
        Else
            strSQL = "Zl_挂号序号状态_Update('" & vsfPlan.TextMatrix(vsfPlan.Row, GetCol("号别")) & _
                  "',To_Date('" & Format(datThis, "yyyy-MM-dd hh:mm:ss") & "','YYYY-MM-DD hh24:mi:ss')," & lngSN & _
                  ",3,'" & UserInfo.姓名 & "'," & IIf(mcbrToolBar.Controls.Find(xtpControlButton, 2605).Enabled, "1", "0") & ",Null," & vsfPlan.TextMatrix(vsfPlan.Row, GetCol("记录ID")) & ")"
        End If
        
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        
        '刷新状态
        txtSN.Text = ""
        Call vsfPlan_EnterCell
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub ClearPatientInfo()
'功能:清除病人相关信息
    Dim i As Integer
    If Not (mblnNewCard And gblnNewCardNoPop) Then mblnAddCardItem = False
    mblnICCard = False
    mstrPrePati = ""
    txtPatient.Text = ""
    txtPatient.IMEMode = 0
    Call ShowDeposit(False)
    lbl险类.Caption = ""
    lbl险类.Visible = False
    If mbytMode = 1 Then
        vsfPay.Visible = False
    Else
        vsfPay.Visible = True
    End If
    If mbytMode = 1 And mblnIDCardKind Then
        '31182
    Else
        txt年龄.Text = ""
        txt年龄.Tag = ""
        cbo年龄单位.Tag = ""
        Call zlControl.CboLocate(cbo年龄单位, "岁")
        Call txt年龄_Validate(False)
        If gstr性别 <> "无" Then SetCboDefault cbo性别
    End If
    mdbl预交余额 = 0
    For i = 1 To vsfPay.Rows - 1
        If vsfPay.RowData(i) = 0 And vsfPay.TextMatrix(i, 0) <> "" Then
            vsfPay.TextMatrix(i, 6) = 0
        End If
    Next i
    mdbl个帐余额 = 0
    cbo家庭地址.Text = ""
    cbo户口地址.Text = ""
    txtIDCard.Text = ""
    txtIDCard.Tag = ""
    txt证件.Tag = "": txt证件.Text = ""
    txt家庭电话.Text = ""
    '89242:李南春,2015/12/7,读取病人地址信息
    Call zlLoadDefaultAddr(padd家庭地址)
    Call zlLoadDefaultAddr(padd户口地址)
    txt门诊号.Text = ""
    txt出生日期.Text = "____-__-__"
    txt出生时间.Text = "__:__"
    stbThis.Panels(2).Text = ""
    imgPatiPic.Picture = Nothing
    SetCboDefault cbo付款方式
End Sub

Private Sub CopyCboTofrmPatiInfo()
    Dim i As Long
    
    With mobjfrmPatiInfo
        .cbo性别.Clear
        For i = 0 To cbo性别.ListCount - 1
            .cbo性别.AddItem cbo性别.List(i)
            .cbo性别.ItemData(i) = cbo性别.ItemData(i)
        Next
        .cbo年龄单位.Clear
        For i = 0 To cbo年龄单位.ListCount - 1
            .cbo年龄单位.AddItem cbo年龄单位.List(i)
            .cbo年龄单位.ItemData(i) = cbo年龄单位.ItemData(i)
        Next
        .cbo付款方式.Clear
        For i = 0 To cbo付款方式.ListCount - 1
            .cbo付款方式.AddItem cbo付款方式.List(i)
            .cbo付款方式.ItemData(i) = cbo付款方式.ItemData(i)
        Next
        .cbo费别.Clear
        For i = 0 To cbo费别.ListCount - 1
            .cbo费别.AddItem cbo费别.List(i)
            .cbo费别.ItemData(i) = cbo费别.ItemData(i)
        Next
    End With
End Sub

Private Sub CopyInfoTofrmPatiInfo()
    With mobjfrmPatiInfo
        .txtPatient.Text = txtPatient.Text: .txtPatient.MaxLength = txtPatient.MaxLength
        '74428：李南春，2014-7-8，病人姓名颜色处理
        .txtPatient.ForeColor = txtPatient.ForeColor
        If Not mrsInfo Is Nothing And (mbytMode = 2 Or mbytMode = 0 And mbytInState = 0 And mstrNoIn <> "") Then
            '31182:只有预约挂号才会存在
            .txt卡号.Tag = Val(Nvl(mrsInfo!病人ID))
        Else
            .txt卡号.Tag = 0
        End If
        If Not mrsInfo Is Nothing Then
            .mlng病人ID = Val(Nvl(mrsInfo!病人ID))
        Else
            .mlng病人ID = 0
        End If
        .cbo性别.ListIndex = cbo性别.ListIndex
        .cbo年龄单位.ListIndex = cbo年龄单位.ListIndex
        .cbo年龄单位.Tag = .cbo年龄单位.Text
        .txt年龄.Text = txt年龄.Text: .txt年龄.MaxLength = txt年龄.MaxLength
        .txt年龄.Tag = txt年龄.Text
        .cbo家庭地址.Text = cbo家庭地址.Text
        .txtRegLocation.Text = cbo户口地址.Text
        '89242:李南春,2015/12/7,读取病人地址信息
        Call .padd家庭地址.LoadStructAdress(padd家庭地址.value省, padd家庭地址.value市, padd家庭地址.value区县, padd家庭地址.value乡镇, padd家庭地址.value详细地址)
        Call .padd户口地址.LoadStructAdress(padd户口地址.value省, padd户口地址.value市, padd户口地址.value区县, padd户口地址.value乡镇, padd户口地址.value详细地址)
        .txt门诊号.Text = txt门诊号.Text: .txt门诊号.MaxLength = txt门诊号.MaxLength
        .cbo付款方式.ListIndex = cbo付款方式.ListIndex
        .txt家庭电话.Text = txt家庭电话.Text
        .cbo费别.ListIndex = cbo费别.ListIndex
        .cbo费别.Locked = cbo费别.Locked
        .cbo费别.TabStop = cbo费别.TabStop
        .txt出生日期.Tag = txt出生日期.Text
        .txt出生时间.Tag = txt出生时间.Text
        .txt出生日期.Text = txt出生日期.Text
        .txt出生时间.Text = txt出生时间.Text
        .txt身份证号.Text = txtIDCard.Text
        .txt身份证号.Tag = txtIDCard.Text
        .imgPatient.Picture = imgPatiPic.Picture
    End With
    
    Call CopyZJTofrmPatiInfo
End Sub

Private Sub CopyZJTofrmPatiInfo()
    Dim lngRow As Long, lngCol As Long, blnFind As Boolean
    '将证件信息赋值到证件列表中对应的卡类型下面，没有就自动增加
     '身份证不处理
    If IDKind证件.IDKind = IDKind证件.GetKindIndex("身份证号") Then Exit Sub
    With mobjfrmPatiInfo.vsCertificate
        For lngRow = 1 To .Rows - 1
            For lngCol = 0 To .Cols - 1 Step 2
                If .TextMatrix(lngRow, lngCol) = IDKind证件.GetCurCard.名称 Then
                    .TextMatrix(lngRow, lngCol + 1) = txt证件.Text
                    blnFind = True
                    Exit For
                End If
            Next
        Next
        '没找到自动添加
        If Trim(txt证件.Text) <> "" And Not blnFind Then
            blnFind = False '是否找到了空位添加
            For lngRow = 1 To .Rows - 1
                For lngCol = 0 To .Cols - 1 Step 2
                    If .TextMatrix(lngRow, lngCol) = "" And .TextMatrix(lngRow, lngCol + 1) = "" Then
                        .TextMatrix(lngRow, lngCol) = IDKind证件.GetCurCard.名称
                        .TextMatrix(lngRow, lngCol + 1) = txt证件.Text
                        blnFind = True: Exit For
                    End If
                Next
            Next
            
            If Not blnFind Then
                If lngCol = 2 Then
                    .TextMatrix(lngRow, lngCol) = IDKind证件.GetCurCard.名称
                    .TextMatrix(lngRow, lngCol + 1) = txt证件.Text
                Else
                    .Rows = .Rows + 1
                    .TextMatrix(lngRow, 0) = IDKind证件.GetCurCard.名称
                    .TextMatrix(lngRow, 1) = txt证件.Text
                End If
            End If
        End If
    End With
End Sub

Private Sub CopyInfoFromobjfrmPatiInfo()
    Dim lngRow As Long, lngCol As Long
    
    With mobjfrmPatiInfo
        txtPatient.Text = .txtPatient.Text  '调用Change事件
        '74428：李南春，2014-7-8，病人姓名颜色处理
        txtPatient.ForeColor = .txtPatient.ForeColor
        mstrPrePati = txtPatient.Text
        cbo性别.ListIndex = .cbo性别.ListIndex
        txt年龄.Text = .txt年龄.Text
        txt年龄.Tag = txt年龄.Text
        txt家庭电话.Text = .txt家庭电话.Text
        cbo年龄单位.ListIndex = .cbo年龄单位.ListIndex
        txt出生日期.Text = .txt出生日期.Text
        txt出生时间.Text = .txt出生时间.Text
        Call txt年龄_Validate(False)
        
        cbo家庭地址.Text = .cbo家庭地址.Text
        cbo户口地址.Text = .txtRegLocation.Text
        '89242:李南春,2015/12/7,读取病人地址信息
        Call padd家庭地址.LoadStructAdress(.padd家庭地址.value省, .padd家庭地址.value市, .padd家庭地址.value区县, .padd家庭地址.value乡镇, .padd家庭地址.value详细地址)
        Call padd户口地址.LoadStructAdress(.padd户口地址.value省, .padd户口地址.value市, .padd户口地址.value区县, .padd户口地址.value乡镇, .padd户口地址.value详细地址)
        txt门诊号.Text = .txt门诊号.Text
        cbo付款方式.ListIndex = .cbo付款方式.ListIndex
        cbo费别.ListIndex = .cbo费别.ListIndex
        cbo年龄单位.Tag = cbo年龄单位.Text
        txtIDCard.Tag = .txt身份证号.Text
        txtIDCard.Text = .txt身份证号.Text
        imgPatiPic.Picture = .imgPatient.Picture
        
        If Trim(.txtPatiMCNO(0).Text) <> "" Then Call SetCboDefault(cbo医疗类别)
    End With
    '90875:李南春,2016/11/8,医疗卡证件类型
    '从证件列表中找到当前卡类型和卡号
    '身份证不处理
    If IDKind证件.IDKind = IDKind证件.GetKindIndex("身份证号") Then Exit Sub
    With mobjfrmPatiInfo.vsCertificate
        For lngRow = 1 To .Rows - 1
            For lngCol = 0 To .Cols - 1 Step 2
                If .TextMatrix(lngRow, lngCol) = IDKind证件.GetCurCard.名称 Then
                    txt证件.Tag = .TextMatrix(lngRow, lngCol + 1)
                    txt证件.Text = txt证件.Tag
                    Exit For
                End If
            Next
        Next
    End With
End Sub


Private Function LoadCard(blnBoundCard As Boolean, Optional blnNotCardFee As Boolean = False) As Boolean
'功能:刷卡调用
'参数:blnBoundCard-绑定就诊卡,此模式下,病人信息窗口显示并允许录入就诊卡,否则为发新卡模式
'        blnNotCardFee-不收取卡费(只有在点绑定卡并且病人姓名处为空时,才为是绑定卡),问题:38841
'返回:True-未建档,卡费和挂号费一起收,false-已建档,卡费存为划价单

    Dim blnInRange As Boolean
    Dim strCardNo As String
    '90875:李南春,2016/11/8,医疗卡证件类型
    If IDKind.GetCurCard.是否证件 Then Exit Function
    
    mbln发卡 = False '问题号:56599
    '115168:李南春，2017/12/13，保存发卡的医疗卡类型
    mCurSendCard = gCurSendCard
    If Not blnBoundCard Then
        Call ClearmobjfrmPatiInfoFace
    End If
    
    With mobjfrmPatiInfo
        .mbytFun = 1
        Set .mrs家庭地址 = mrs家庭地址
        
        If blnBoundCard Then
            .mstrCard = ""
            Call CopyCboTofrmPatiInfo
            Call CopyInfoTofrmPatiInfo
        
            If .txt门诊号.Text = "" Then .txt门诊号.Text = zlGet门诊号
        Else
            '发新卡,在刷卡时就检查就诊卡是否有，是否在范围内
            blnInRange = True
            .mblnInRange = blnInRange
            .mstrCard = UCase(txtPatient.Text)
            .txt卡号.Text = .mstrCard
            
            mbln发卡 = bln发卡(.txt卡号.Text)
            
            If mbln发卡 = False And InStr(mstrPrivs, ";绑定卡号;") = 0 Then
                MsgBox "你没有绑定卡号的权限，不能绑定该卡！", vbInformation, gstrSysName
                Exit Function
            End If
            
            If Not gblnNewCardNoPop Then
                .txt门诊号.Text = zlGet门诊号
                txt门诊号.Text = .txt门诊号.Text
            End If
        End If
        If Not blnBoundCard And CreatePlugInOK(mlngModul) Then
            If Not zlReadPlugInPati(UCase(txtPatient.Text), mblnBrushPlugin) Then
                .txt卡号.Text = ""
                .txt密码.Text = ""
                .txt验证.Text = ""
                mblnAddCardItem = False
                Exit Function
            End If
        Else
            mblnBrushPlugin = False
        End If
        
        If blnBoundCard Or Not gblnNewCardNoPop Then
            '问题号:53408
            Set mobjfrmPatiInfo.mrsPatiInfo = mrsInfo
            '问题号:56599
            mobjfrmPatiInfo.mbln发卡 = mbln发卡
            .mlng监护人年龄 = mTy_Para.lngN岁以下录入监护人
            .mbln监护人录入 = mTy_Para.bln监护人录入
            If mrsInfo Is Nothing Then
                .mlng病人ID = 0
            Else
                .mlng病人ID = mrsInfo!病人ID
            End If
            Call CloseIDCard '47007
            
            .ShowMe 1, Me
            
            Call NewCardObject '47007
            If .GetmblnCancel = True Then
                .txt卡号.Text = ""
                .txt密码.Text = ""
                .txt验证.Text = ""
                Call CopyCboTofrmPatiInfo
                Call CopyInfoTofrmPatiInfo
                Call NewCardObject
                Exit Function
            End If
            
            Set mrsInfo = Nothing
            Set mrsInfo = mobjfrmPatiInfo.mrsPatiInfo
            mstr门诊号 = mobjfrmPatiInfo.txt门诊号
        Else
            '104238:李南春，2017/2/15，检查卡号是否满足发卡控制限制
            If .txt卡号.Text <> "" And Len(.txt卡号.Text) <> gCurSendCard.lng卡号长度 And Not gCurSendCard.bln严格控制 Then
                Select Case gCurSendCard.byt发卡控制
                    Case 0
                        MsgBox "输入的卡号小于" & gCurSendCard.str卡名称 & "设定的卡号长度，请重新输入！", vbExclamation, gstrSysName
                        If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus
                            Exit Function
                    Case 2
                        If MsgBox("输入的卡号小于" & gCurSendCard.str卡名称 & "设定的卡号长度，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus
                            Exit Function
                        End If
                End Select
            End If
        End If
        '刘兴洪:27493 20100117:lnBoundCard = False
        If blnBoundCard Then
            If .mlng病人ID <> 0 And gbln卡费仅划价 Then
                strCardNo = .mlng病人ID
                Call GetPatient(IDKind.GetCurCard, "-" & strCardNo, True)
                LoadCard = True
                cmdCard.Enabled = False
                Exit Function
            End If
            Call CopyInfoFromobjfrmPatiInfo
            blnInRange = IIf(blnNotCardFee, False, True)
            If .txt卡号.Text <> "" Then
                mbln发卡 = bln发卡(.txt卡号.Text)
            End If
            '31182
            If (mbytMode = 2 Or mbytMode = 0 And mbytInState = 0 And mstrNoIn <> "") And Not mrsInfo Is Nothing Then
                mblnAddCardItem = .txt卡号.Text <> "" And blnInRange And mbln发卡
            Else
                mblnAddCardItem = .txt卡号.Text <> "" And blnInRange And mbln发卡
            End If
            If .txt卡号.Text <> "" Then
                lblPrompt.Caption = gCurSendCard.str短名称 & ":" & .txt卡号.Text & "(" & IIf(mbln发卡, "发卡", "绑定卡") & ")"
                vsfPay.Height = 1755
                lblPrompt.Top = vsfPay.Top + vsfPay.Height + 60
            Else
                lblPrompt.Caption = ""
                vsfPay.Height = 2220
            End If
            Call ReLoadCardFee(True)
            LoadCard = True
        Else
            If .mstrCard <> "" Then
                If gbln卡费仅划价 And Not gblnNewCardNoPop Then     '档案建立成功,绑定就诊卡模式固定不建档
                    Call GetPatient(IDKind.GetCurCard, txtPatient.Text, True)
                Else
                    mblnUnChange = True
                    Call CopyInfoFromobjfrmPatiInfo
                    mblnUnChange = False
                    If Me.ActiveControl Is txtPatient Then
                            If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(txtPatient.Text = "")
                            If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(txtPatient.Text = "")
                    End If
                    If gbln卡费仅划价 Then
                        mblnAddCardItem = False
                    Else
                        mblnAddCardItem = mbln发卡
                    End If
                    lblPrompt.Caption = gCurSendCard.str短名称 & ":" & .mstrCard & "(" & IIf(mbln发卡, "发卡", "绑定卡") & ")"
                    vsfPay.Height = 1755
                    lblPrompt.Top = vsfPay.Top + vsfPay.Height + 60
                End If
                Call ReLoadCardFee
                LoadCard = True
            Else '在弹出窗口选择了取消发新卡
                cmdMore.Enabled = False
            End If
            cmdCard.Enabled = False
        End If
    End With
    Call AdjustInfoPosition
    If CheckIsPrice Or mRegistFeeMode = EM_RG_记帐 Then
        Call SetUndisplayBalance
    Else
        Call SetShowBalance
    End If
    
End Function

Public Sub SetCardDisplay(ByVal strPrompt As String)
    lblPrompt.Caption = strPrompt
    If strPrompt = "" Then
        vsfPay.Height = 2220
    Else
        vsfPay.Height = 1755
        lblPrompt.Top = vsfPay.Top + vsfPay.Height + 60
    End If
    mblnNoClearPrompt = True
End Sub

Private Sub SetmobjfrmPatiInfo()
    Dim i As Long, str过敏 As String
    
    With mobjfrmPatiInfo
    
        .cbo国籍.ListIndex = cbo.FindIndex(.cbo国籍, Nvl(mrsInfo!国籍), True)
        .cbo民族.ListIndex = cbo.FindIndex(.cbo民族, Nvl(mrsInfo!民族), True)
        .cbo婚姻.ListIndex = cbo.FindIndex(.cbo婚姻, Nvl(mrsInfo!婚姻状况), True)
        '76314,李南春，2014-08-06，病人信息正确获取
        .cbo职业.ListIndex = cbo.FindIndex(.cbo职业, Nvl(mrsInfo!职业))
        .txt身份证号.Text = Nvl(mrsInfo!身份证号)
        .txt身份证号.Tag = .txt身份证号.Text
        .txt单位名称.Text = Nvl(mrsInfo!工作单位)
        .txt区域.Text = Trim(Nvl(mrsInfo!区域))
        .txt区域.Tag = .txt区域.Text
        .txt单位名称.Tag = Nvl(mrsInfo!合同单位ID)
        .txt单位电话.Text = Nvl(mrsInfo!单位电话)
        .txt单位邮编.Text = Nvl(mrsInfo!单位邮编)
        .txt家庭电话.Text = Nvl(mrsInfo!家庭电话)
        .txt家庭邮编.Text = Nvl(mrsInfo!家庭地址邮编)
        .txt联系人身份证.Text = Nvl(mrsInfo!联系人身份证号)
        .txtBirthLocation.Text = Nvl(mrsInfo!出生地点)
        .txtRegLocation.Text = Nvl(mrsInfo!户口地址)
        '89242:李南春,2015/12/7,读取病人地址信息
        Call zlReadAddrInfo(.padd户口地址, Val(Nvl(mrsInfo!病人ID)), 0, 4, Nvl(mrsInfo!户口地址))
        .txt户口地址邮编.Text = Nvl(mrsInfo!户口地址邮编)
'        '73609:李南春，2014-8-1，病人信息保存
'        .txtRegLocation.Tag = Nvl(mrsInfo!户口地址邮编)
        '问题号:40005
        .txt联系人电话.Text = Nvl(mrsInfo!联系人电话)
        '84313,李南春,2015/4/27,联系人关系以及其他关系
        .txt其他关系.Text = ""
        .cbo联系人关系.ListIndex = cbo.FindIndex(.cbo联系人关系, Nvl(mrsInfo!联系人关系), True)
        If .cbo联系人关系.ListIndex <> 8 Then .txt其他关系.Text = "": .txt其他关系.Visible = False
        .txt联系人姓名.Text = Nvl(mrsInfo!联系人姓名)
        .txt监护人.Text = Nvl(mrsInfo!监护人)
        .Load健康卡相关信息 (mrsInfo!病人ID)
        .LoadCertificate (mrsInfo!病人ID)
    End With
End Sub

Private Sub ShowPatiInfo()
    Dim i As Integer
    Dim strSimilar As String
    
    If txtPatient.Text = "" Then Exit Sub
    
    With mobjfrmPatiInfo
        .mbytFun = 0
        Set .mrs家庭地址 = mrs家庭地址
        Call CopyCboTofrmPatiInfo
        Call CopyInfoTofrmPatiInfo
                
        If .txt门诊号.Text = "" Then .txt门诊号.Text = zlGet门诊号
'        .txt门诊号.Enabled = mrsInfo Is Nothing
                
        If mlngOutModeMC > 0 Then
            .txtPatiMCNO(0).Enabled = (mstrYBPati = "")
            .txtPatiMCNO(1).Enabled = .txtPatiMCNO(0).Enabled
        End If
    End With
    mobjfrmPatiInfo.mlng监护人年龄 = mTy_Para.lngN岁以下录入监护人
    mobjfrmPatiInfo.mbln监护人录入 = mTy_Para.bln监护人录入
    mobjfrmPatiInfo.mstrPrivs = mstrPrivs
    mobjfrmPatiInfo.mlngModul = mlngModul
    Call CloseIDCard
    mobjfrmPatiInfo.ShowMe 1, Me
    Call NewCardObject
    If mobjfrmPatiInfo.GetmblnCancel = False Then
        '如果是刷卡新建病人档案,则在mobjfrmPatiInfo里点确定时生成病人信息之前处理
        If Trim(mobjfrmPatiInfo.txt身份证号.Text) <> "" And cmdMore.Tag = "" And mobjfrmPatiInfo.cmdOK.Caption Like "返回*" And mobjfrmPatiInfo.txt身份证号.Tag <> Trim(mobjfrmPatiInfo.txt身份证号.Text) Then
            '检查相似病人信息(新增之前检查,以免加入了重复信息！！！)
            With mobjfrmPatiInfo
                strSimilar = SimilarIDs(.txt身份证号.Text)
            End With
            cmdMore.Tag = "已检查"      '在txtPatient_change中清空
            
            If strSimilar <> "" Then
                i = UBound(Split(strSimilar, "|")) + 1
                strSimilar = Replace(strSimilar, "|", vbCrLf)
                If i > 20 Then strSimilar = Mid(strSimilar, 1, 200) & "..."
                
                If MsgBox("在已有的病人信息中发现 " & i & " 个信息相似的病人(身份证号相同): " & vbCrLf & vbCrLf & _
                    strSimilar & vbCrLf & vbCrLf & "登记为新病人请选择[是],提取已有的病人信息请选择[否]？", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    If i = 1 Then
                        txtPatient.Text = "-" & Mid(Split(strSimilar, ",")(0), 4)
                        Call txtPatient_Validate(False)
                    Else
                        txtPatient.SetFocus
                    End If
                    Exit Sub
                End If
            End If
        End If
        
        Call CopyInfoFromobjfrmPatiInfo
    Else
        Call CopyCboTofrmPatiInfo
        Call CopyInfoTofrmPatiInfo
    End If
    
    '74430,冉俊明,2014-7-8,挂号界面显示病人照片的浮动窗体
    If picPatiPicBack.Visible Then Call ShowPatiPic
    
    If cbo结算方式.Enabled And cbo结算方式.Visible Then
        cbo结算方式.SetFocus
    ElseIf chk病历费.Enabled And chk病历费.Visible Then
        chk病历费.SetFocus
    Else
        cmdOK.SetFocus
    End If
End Sub

Private Sub cmdCard_Click()
    Dim blnBound As Boolean
    
    If LoadCard(True, blnBound) Then
        Call ShowRegistFromInput    '可能先绑定卡号返回后再次进入清除卡号
         '问题号:56039,56355
        If Val(zlDatabase.GetPara("挂号发票打印方式", glngSys, mlngModul)) <> 0 Then
           Call ReInitPatiInvoice
        End If
        
        If mobjfrmPatiInfo.txt卡号.Text <> "" Then
            mblnNewCard = True
            Call SetOneCardBalance
        Else
            SetCboDefault cbo结算方式
        End If
    End If
    If cbo结算方式.Enabled And cbo结算方式.Visible Then
        cbo结算方式.SetFocus
    ElseIf chk病历费.Enabled And chk病历费.Visible Then
        chk病历费.SetFocus
    Else
        cmdOK.SetFocus
    End If
    mblnBoundPati = blnBound
    '
    mobjfrmPatiInfo.mblnNewPatient = False
End Sub

Private Sub cmdMore_Click()
    Call ShowPatiInfo
    '
    mobjfrmPatiInfo.mblnNewPatient = False
End Sub

Private Sub cmdLookup_Click()
    frmPatiFind.Show 1, Me
    If frmPatiFind.mlng病人ID <> 0 Then
        Me.Refresh
        txtPatient.Text = "-" & frmPatiFind.mlng病人ID
        Call txtPatient_Validate(False)
    Else
        If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    End If
End Sub

Private Sub dtpAppointmentDate_Change()
    txtSN.Text = ""
    Call ShowPlans
    dtpAppointmentDate.Tag = Format(dtpAppointmentDate.Value, "yyyy-mm-dd HH:MM:SS")
    If txt号别.Text <> "" Then
        If zlCheck限约或限号数(Val(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("记录ID")))) = False Then
            ClearBill (False)
        End If
    End If
    dtpAppointmentDate.SetFocus
End Sub

Private Sub dtpAppointmentDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub Form_Activate()
    Dim lng号别 As Long
    '问题号:57491
    Call picInfoFrame_Resize
    If Not mblnFirst Then Exit Sub
    
    mblnFirst = False
    
    If mblnUnload Then mblnUnload = False: Unload Me: Exit Sub
    
    Call zl权限控制
    
    '医生站挂号时，如果只有一个号，则自动输入
    With vsfPlan
        If .Rows = 2 Then
            lng号别 = GetCol("号别")
            If .TextMatrix(1, lng号别) <> "" And txt号别.Visible And txt号别.Enabled Then
                txt号别.SetFocus
                txt号别.Text = .TextMatrix(.Row, lng号别)
            End If
        End If
    End With
    If mbytInState = 0 And mbytMode = 0 Then
        txtPatient_Change
    End If
    '双屏显示窗体必须在当前窗口显示之后调用显示才能移动窗体
    If mbytMode = 0 And mbytInState = 0 And gblnLED And Trim(txtPatient.Text) = "" Then
        zl9LedVoice.DisplayPatient ""
    End If
    
    If mbytMode = 2 And cmdOK.Visible And cmdOK.Enabled Then
        cmdOK.SetFocus
    End If
    If gCurSendCard.str卡名称 <> "" Then
        cmdCard.ToolTipText = "绑定" & gCurSendCard.str卡名称 & ": F10"
        If mblnSendCard Then cmdCard.ToolTipText = "发" & gCurSendCard.str卡名称 & ": F10"
    End If
    Call picPlan_Resize
    mblnActivate = True
    If mbytMode = 2 And mbytInState = 0 Then
        '102230,调用外挂部件接口
        If Not mrsInfo Is Nothing Then
            If PatiValiedCheckByPlugIn(mlngModul, Val(Nvl(mrsInfo!病人ID)), _
                "<YSXM>" & NeedName(cbo医生.Text) & "</YSXM>") = False Then Unload Me: Exit Sub
        End If
    Else
        Call vsfPlan_EnterCell: If txt号别.Visible And txt号别.Enabled Then txt号别.SetFocus
    End If
    mblnActivate = False
    Call picInfo_Resize
End Sub
Private Sub zl权限控制()
      '刘兴洪 问题:27438 日期:2010-01-13 17:42:32
    If mbytInState <> 0 Then Exit Sub
    If mbytMode = 0 Then
        cmdCard.Visible = InStr(1, mstrPrivs, ";绑定卡号;") > 0
    End If
    Call zlPatiMoveCmdCtrl
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim blnCancel As Boolean
    If mbytInState = 1 Then Exit Sub
    
    Select Case KeyCode
        Case vbKeyF
            If Shift = vbCtrlMask And cmdLookup.Enabled And cmdLookup.Visible Then Call cmdLookup_Click
        Case vbKeyM
            '仅仅ctrl+M
            If Shift <> vbCtrlMask Then Exit Sub
            If ActiveControl Is txtPatient Then Call txtPatient_Validate(False)
            If Shift = vbCtrlMask And cmdMore.Enabled And cmdMore.Visible Then cmdMore_Click
        Case vbKeyF2
            If ActiveControl Is txtPatient Then
                Call txtPatient_Validate(False)
            End If
            If Not blnCancel And cmdOK.Enabled And cmdOK.Visible Then Call cmdOK_Click  '不能设获得焦点,因为保存事件中以此判断是否进行语音报价
        Case vbKeyF3
            If cmdMore.Enabled And cmdMore.Visible Then cmdMore.SetFocus: cmdMore_Click
        Case vbKeyF4
            If Me.ActiveControl Is txtPatient And IDKind.Enabled And txtPatient.Locked Then
                IDKind.ActiveFastKey
            End If
        Case vbKeyF5
            If mcbrToolBar.Controls.Find(xtpControlButton, conMenu_View_Refresh).Visible And mcbrToolBar.Controls.Find(xtpControlButton, conMenu_View_Refresh).Enabled Then RefreshFace
        Case vbKeyF6
            If chkShowAll.Visible And chkShowAll.Enabled Then
                chkShowAll.Value = IIf(chkShowAll.Value = 1, 0, 1)
            End If
        Case vbKeyF7
            If chkPrint.Visible And chkPrint.Enabled Then
                chkPrint.Value = IIf(chkPrint.Value = 1, 0, 1)
                Call chkPrint_Click
            End If
        Case vbKeyF8
            If chkCancel.Enabled And chkCancel.Visible Then
                chkCancel.Value = IIf(chkCancel.Value = 1, 0, 1)
                Call chkCancel_Click
            End If
        Case vbKeyF9
            If txt号别.Enabled And txt号别.Visible Then
                mblnLEDKey = True
                If Not Me.ActiveControl Is txt号别 Then
                    txt号别.SetFocus
                Else
                    Call txt号别_GotFocus 'LED语音报价
                End If
            End If
        Case vbKeyF10
            mbln发卡 = False '问题号:56599
            If ActiveControl Is txtPatient Then Call txtPatient_Validate(False)
            If cmdCard.Visible And cmdCard.Enabled Then Call cmdCard_Click
        Case vbKeyF11
            If txtPatient.Enabled And txtPatient.Visible And Not txtPatient.Locked Then
                If Me.ActiveControl Is txtPatient Then
                    IDKind.IDKind = IIf(IDKind.IDKind = UBound(Split(IDKind.IDkindStr, ";")), IDKind.GetKindIndex("姓名"), IDKind.IDKind + 1)
                Else
                    txtPatient.SetFocus
                End If
            End If
        Case vbKeyF12
            If Shift = vbCtrlMask Then
                chkBooking.Value = IIf(chkBooking.Value = 1, 0, 1)
            Else
                If cboNO.Enabled And cboNO.Visible Then cboNO.SetFocus
            End If
        Case vbKeyAdd
            If mbytInState = 0 And Not mbln病历费 Then Exit Sub
            If (mbytMode = 2 Or mbytMode = 0 And mbytInState = 0 And mstrNoIn <> "") Or chkCancel.Value = 1 Or chkPrint.Value = 1 Or txt号别.Text = "+" Then Exit Sub
            If ActiveControl.Name <> txt号别.Name Then
                chk病历费.Value = IIf(chk病历费.Value = 0, 1, 0)
            End If
        Case 192, 229  '问题:28604:｀
             If Shift <> vbCtrlMask Then
                Exit Sub
             End If
             Call SelectHistoryRegist
    End Select
    
    '74430,冉俊明,2014-7-8,挂号界面显示病人照片的浮动窗体
    If Shift = 2 And KeyCode = vbKeyW Then
         Call ShowPatiPic
    End If
    If Shift = 2 And KeyCode = vbKeyE Then
        Call imgColPlan_Click
    End If
End Sub

Private Sub SelectHistoryRegist()
    '------------------------------------------------------------------------------------------------------------------------
    '功能：选择历次挂号号别
    '编制：刘兴洪
    '日期：2010-08-18 16:14:58
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim lng病人ID As Long, lngPre病人ID As Long, str号别 As String
    Dim blnFind As Boolean, i As Long
    If mbytMode = 2 Then Exit Sub '预约接收不处理
    If mbytInState >= 1 Then Exit Sub  '查阅不处理
    If mrsInfo Is Nothing Then
        lng病人ID = 0
    ElseIf mrsInfo.State <> 1 Then
        lng病人ID = 0
    Else
       lng病人ID = Val(Nvl(mrsInfo!病人ID))
    End If
    lngPre病人ID = lng病人ID
    str号别 = ""
    CloseIDCard
    If frmRegistHistory.ShowRegist(Me, mstrPrivs, mTy_Para.bln允许住院病人挂号, mblnOlnyBJYB, lng病人ID, str号别) = False Then NewCardObject: Exit Sub
    Call CreateMobjIDCard
    If lng病人ID <> lngPre病人ID Then
       '病人不对时,直接读取病人
       Call GetPatient(IDKind.GetCurCard, "-" & lng病人ID, False)
    End If
    
    '查找有此号别没有
    With vsfPlan
       blnFind = False
       For i = 1 To .Rows - 1
           If .TextMatrix(i, .ColIndex("号别")) = str号别 Then
                   .Row = i: .Col = .ColIndex("号别")
                   Call .ShowCell(.Row, .Col)
                   Call vsfPlan_KeyDown(13, 0)
                   blnFind = True: Exit For
           End If
       Next
    End With
    If blnFind = False Then
       Call MsgBox("注意:" & vbCrLf & "    号码为『" & str号别 & "』的号码在当前未进行挂号安排,无法定位!", vbInformation + vbOKOnly, gstrSysName)
       Exit Sub
    End If
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    
    If InStr("':：;；?？", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0: Exit Sub
    ElseIf KeyAscii = Asc("+") Then
        If (mbytMode = 2 Or mbytMode = 0 And mbytInState = 0 And mstrNoIn <> "") Or chkCancel.Value = 1 Or chkPrint.Value = 1 Then KeyAscii = 0
    End If
    If mbytInState = 1 Then Exit Sub
    If InStr("`｀", Chr(KeyAscii)) > 0 Then
        '报请出示就诊卡
         KeyAscii = 0
        If gblnLED Then zl9LedVoice.Speak "#30"  '`为语音报价:有点奇怪:本来应该是192,但不知怎么会成229:32663
    End If
    
End Sub

Private Sub Form_Load()
    Dim lng病历费ID As Long
    Dim strSQL As String, rsTmp As ADODB.Recordset
'    Call InitTimeSect
    '初始化 界面采用的 方式
    InitActionType
    Call zlInitParaSet  '初始化本地参数
    '窗体尺寸限制
    '创建插建
    Call InitCardSquareData
    Call DefMainCommandBars
    Call InitRegist
    Call InitPanel
   ' Call zlInitParaSet  '初始化本地参数
    mblnStartFactUseType = False
    If gblnSharedInvoice Then
        '挂号用门诊票据:42703
        mblnStartFactUseType = zlStartFactUseType("1")
    End If
    Set mrsBillAdvance = Nothing
    mstrPrepayPrivs = ";" & GetPrivFunc(glngSys, 1103) & ";"
    mstrCardPrivs = ";" & GetPrivFunc(glngSys, 1151) & ";"
    mstrSort = ""
    mblnBrushPlugin = False
    Set mobjfrmPatiInfo = New frmPatiInfo
    mobjfrmPatiInfo.mstrPrivs = mstrPrivs
    mobjfrmPatiInfo.mlngModul = mlngModul
    Load mobjfrmPatiInfo
    
    glngOld = 0
    If mbytInState = 0 And mbytMode <> 2 Then
        glngMinW = 15090
        glngMaxW = Screen.Width
        glngMinH = 10605
        glngMaxH = Screen.Height
    Else
        glngMinW = 7600
        glngMaxW = 7600
'        If mbytMode = 2 Then
'            If mbytInState = 0 Then
'                glngMinW = 7600
'                glngMaxW = 7600
'                glngMinH = 10500
'                glngMaxH = 10500
'            Else
'                glngMinH = 9100
'                glngMaxH = 9100
'            End If
'            picInfo.Height = picInfo.Height - 250
'        Else
            glngMinH = 768 * 15 '9400
            glngMaxH = 768 * 15  '9400
            picInfo.Height = picInfo.Height - 350
'        End If
    End If
    
    If Not OS.IsDesinMode Then
        glngOld = GetWindowLong(Me.Hwnd, GWL_WNDPROC)
        Call SetWindowLong(Me.Hwnd, GWL_WNDPROC, AddressOf Custom_WndMessage)
    End If
    
    gblnOk = False
    mblnUnload = False
    mblnFirst = True
    mblnAddCardItem = False
    mblnChange = True
    mstr个人帐户 = ""
    mlng结帐ID = 0
    mintInsure = 0
    mstrYBPati = ""
    mlng磁卡领用ID = 0
    
    cmdComminuty.Visible = False
    If (mbytMode = 0 Or mbytMode = 1) And mbytInState = 0 Then
        Set mobjIDCard = New clsIDCard
        Set mobjICCard = New clsICCard
        Call mobjIDCard.SetParent(Me.Hwnd)
        Call mobjICCard.SetParent(Me.Hwnd)
        Set mobjICCard.gcnOracle = gcnOracle

        '社区接口初始化
        Call CreateCommunity
        
    End If
    
    If mintCancel = 1 Then
        lng病历费ID = 0
        strSQL = "Select 收费细目ID From 收费特定项目 Where 特定项目='病历费'"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        If Not rsTmp.EOF Then
            lng病历费ID = Val(Nvl(rsTmp!收费细目ID))
        End If
        
        If lng病历费ID = 0 Then
            MsgBox "没有发现病历费的收费特定项目，请检查！", vbExclamation, gstrSysName
            mblnUnload = True
        Else
            mstr退费项目IDs = lng病历费ID
        End If
    End If
    
    mstr附加费 = ""
    mstr附加项目ID = ""
    strSQL = "Select zl_Fun_RegCustomName As 附加费 From Dual"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If Not rsTmp.EOF Then
        mstr附加费 = Split(Nvl(rsTmp!附加费) & "|", "|")(0)
        mstr附加项目ID = Split(Nvl(rsTmp!附加费) & "|", "|")(1)
    End If
    
    If mstr附加费 <> "" Then
        chkExtra.Caption = "退" & mstr附加费
    End If

    '初始化数据
    If mbytInState = 0 Then
        mobjfrmPatiInfo.mstrPriceGrade = gstrPriceGrade
    End If
    Call Load支付方式
    Call InitFace
    If mbytInState <> 1 And mbytInState <> 2 And mbytInState <> 3 Then
        Call RestoreWinState(Me, App.ProductName, mbytMode & mbytInState)
        stbThis.Visible = True
    End If
    
    Call InitData
    '问题号:57491
    If mblnUnload Then
        Exit Sub
    End If
    
    Call SetDelBillCtlEnabled
    
    
    If mblnStation And mbytMode = 0 And mTy_Para.bln挂号必须刷卡 Then LoadIdKindStr  '如果是医生工作站挂号并且挂号必须刷卡时需要 重新加载 IDKind的相应信息
    If mblnUnload Then Exit Sub
    
    If mbytMode = 1 Then
        '预约 需要初始化合作单位挂号
        Call InitUnitRegData
    End If
    
    If Me.Height < glngMinH Then Me.Height = glngMinH
    If Me.Width <= glngMinW Then Me.Width = glngMinW
    
    If mbytInState = 1 Or (mbytInState = 0 And mbytMode = 2) Then '查阅时,不能更改窗体大小:25623
        Call zlSetWindowsBroldStyle(Me)
        Call Form_Resize
    End If
    zlControl.PicShowFlat picInfoFrame, -1, , taCenterAlign
    zlControl.PicShowFlat picPlan, -1, , taCenterAlign
    zlControl.PicShowFlat picTotal, -1, , taCenterAlign

    'LED初始化
    If mbytMode <> 1 And mbytInState = 0 And gblnLED Then
        zl9LedVoice.Reset com
        zl9LedVoice.Init UserInfo.编号 & " 挂号员为您服务", mlngModul, gcnOracle
    End If
End Sub

Private Sub InitUnitRegData()
    Dim strSQL As String
    Dim rsTmp   As ADODB.Recordset
    
    strSQL = " Select 1 as 数据  From 临床出诊挂号控制记录 Where 类型=1 And Rownum < 2 "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If rsTmp.EOF Then Exit Sub
    mblnUnitReg = rsTmp.RecordCount > 0
    Set rsTmp = Nothing
    Exit Sub
Hd:
    If ErrCenter() = 1 Then
        Resume
    End If
    SaveErrLog
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If mbytMode <> 2 And mbytInState = 0 And Not mblnUnload And gblnOk And Not mblnCharge And Not mblnStation Then
        If MsgBox("真的要退出吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = 1: Exit Sub
        End If
    End If
End Sub

Private Sub Form_Resize()
    Dim lngSNHeight As Long
    If WindowState = 1 Then Exit Sub
    
    On Error Resume Next
    
    If vsfList.Visible Then
     '*****************************
        lngSNHeight = (picPlan.Height - IIf(picBookingDate.Visible, picBookingDate.Height, 0)) * 1 / 3
        vsfList.Height = lngSNHeight
    End If
    
    txtPatientPrint.Left = txtPatient.Left
    txtPatientPrint.Top = txtPatient.Top
    If mbytMode = 1 Then
        If mbytInState = 0 Then
            cmdOK.Top = lblSum.Top + lblSum.Height + 1150
            cmdCancel.Top = cmdOK.Top
        Else
            picTotal.Width = picBal.Left - picTotal.Left
            lbl合计.Left = picTotal.Width - lbl合计.Width - 150
        End If
    ElseIf mbytMode = 3 Then
        picTotal.Width = picBal.Left - picTotal.Left
        lbl合计.Left = picTotal.Width - lbl合计.Width - 150
    ElseIf mbytMode = 2 Then
        If mbytInState = 1 Then
            picTotal.Width = picBal.Left - picTotal.Left
            lbl合计.Left = picTotal.Width - lbl合计.Width - 150
        End If
    Else
        
    End If
    picTop.Left = Me.ScaleWidth - picTop.Width - 60
    Call AdjustInfoPosition
End Sub

Private Sub AdjustInfoPosition()
    On Error Resume Next
    lblCancel.Left = picTop.ScaleWidth - lblCancel.Width - 150
    lblFree.Left = lblCancel.Left - lblFree.Width - 150
    lbl急.Left = lblFree.Left - lbl急.Width - 150
    lbl险类.Left = lbl急.Left - lbl险类.Width - 180
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call YBIdentifyCancel '取消医保病人身份验证
    
    Call SaveWinState(Me, App.ProductName, mbytMode & mbytInState)
    
    mblnRegReceiveByNo = False '问题号:57423
    mblnViewCancel = False
    mstrNoIn = ""
    mblnNOMoved = False
    mblnUnChange = False
    zl_vsGrid_Para_Save mlngModul, vsfPlan, Me.Caption, "vsfPlan" & mbytMode
    mblnCharge = False
    mblnStation = False
    mstrRoom = ""
    mstrPreNO = ""
    mblnNoneCut = False
    mintCancel = 0
    mstrForceNote = ""
    mblnCenter = False
    mblnViewOriginal = False
    Set mrsALL时间段 = Nothing
    Set mrs时间段 = Nothing
    Set mrsItems = Nothing
    Set mrsInComes = Nothing
    Set mrsPlan = Nothing
    Set mrsInfo = Nothing
    Set mrs费别 = Nothing
    Set mrsDoctor = Nothing
    Set mrsSNState = Nothing
    Set mrsBillAdvance = Nothing
    Set mobjDelCards = Nothing
    Set mobjPayCard = Nothing
    If Not mrs家庭地址 Is Nothing Then
        If mrs家庭地址.State = 1 Then
            On Error Resume Next
            Kill App.Path & "\ZLAddressForRegEvent.Adtg"
            Err.Clear
            mrs家庭地址.Filter = ""
            mrs家庭地址.Save App.Path & "\ZLAddressForRegEvent.Adtg"
        End If
    End If
    Set mrs家庭地址 = Nothing
    
    mbln病历费 = False
    mbln包含病历费 = False
    mlng领用ID = 0
    
    mstrPrePati = ""
    mcur合计 = 0: mint挂号数 = 0
    mcur应缴 = 0
    
    If Not mobjfrmPatiInfo Is Nothing Then Unload mobjfrmPatiInfo
    Set mobjfrmPatiInfo = Nothing
    
    If Not OS.IsDesinMode And glngOld > 0 Then
        Call SetWindowLong(Me.Hwnd, GWL_WNDPROC, glngOld)
    End If
    If Not mobjRegist Is Nothing Then Set mobjRegist = Nothing
    
    'LED初始化
    If mbytMode <> 1 And mbytInState = 0 And gblnLED Then
        zl9LedVoice.DisplayPatient ""
        zl9LedVoice.Reset com
    End If
    mintIDKind = IDKind.IDKind
    If mbytInState = 0 Then
        Call SaveRegInFor(g私有模块, Me.Name, "idkind", mintIDKind)
    End If
    If mbytMode = 1 And mbytInState = 0 Then
        Call zlDatabase.SetPara("预约显示所有号别", IIf(chkShowAll.Value = 1, 1, 0), glngSys, mlngModul, InStr(mstrPrivs, ";参数设置;") > 0)
    End If
    
    Call CloseIDCard
    mbytMode = 0
    mbytInState = 0
    mbln附加费 = False
    mstrPrivs = ""
    '问题号:53408
    mstr门诊号 = ""
    '问题号:56599
    mbln发卡 = False
    Set mobjHealthCard = Nothing
    mblnNotEMPIQuery = False
    '127839：李南春,2018/6/27，清空变量
    mcustomTime = t_普通
    mViewMode = V_普通号
    mblnUnload = False
    mbln连续挂号 = False
    mblnGetBirth = False
End Sub

Private Sub lbl合计_Change()
    Call txt缴款_Change
End Sub

Private Sub picInfoFrame_Resize()
    On Error Resume Next
    With picInfoFrame
        picInfo.Top = 15
        picInfo.Left = 15
        picInfo.Height = .ScaleHeight - picInfo.Top * 2
        picInfo.Width = .ScaleWidth - picInfo.Left * 2
    End With
End Sub

Private Sub picPlan_Resize()
    On Error Resume Next
    sc安排.Width = picPlan.ScaleWidth
    vsfPlan.Width = picPlan.ScaleWidth
    vsfList.Width = picPlan.ScaleWidth
    picBookingDate.Width = picPlan.ScaleWidth
    picSplit.Width = picPlan.ScaleWidth
    picTime.Width = vsfPlan.Width
    
    sc安排.Top = IIf(picBookingDate.Visible, picBookingDate.Height + 30, 0)
    chkShowAll.Top = sc安排.Top + 45
    chkShowAll.Left = sc安排.Width - chkShowAll.Width - 300
    vsfPlan.Top = sc安排.Top + sc安排.Height + 45
    vsfPlan.Height = picPlan.ScaleHeight - vsfPlan.Top - 360 - IIf(vsfList.Visible, vsfList.Height + picSplit.Height + 30, 120) - IIf(picTime.Visible, picTime.Height, 0)
    picSplit.Top = vsfPlan.Top + vsfPlan.Height
    picTime.Top = picSplit.Top + picSplit.Height + 15
    vsfList.Top = IIf(picTime.Visible, picTime.Top + picTime.Height + 60, picTime.Top)
    vsfList.Height = picPlan.ScaleHeight - vsfList.Top - 380
End Sub

Private Sub picSerialInfo_LostFocus()
    picSerialInfo.Visible = False
End Sub

Private Sub picSplit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If vsfPlan.Height + Y < 500 Or vsfList.Height - Y < 500 Then Exit Sub
                
        picSplit.Top = picSplit.Top + Y
        vsfPlan.Height = vsfPlan.Height + Y
        vsfList.Top = vsfList.Top + Y
        vsfList.Height = vsfList.Height - Y
        If picTime.Visible Then
            picTime.Top = picSplit.Top + picSplit.Height + 15
        End If
        Me.Refresh
    End If
End Sub

Private Sub picTime_Resize()
    dtpAppointmentTime.Left = picTime.Width - dtpAppointmentTime.Width - 100
    lbl预约时间.Left = dtpAppointmentTime.Left - lbl预约时间.Width - 20
End Sub

Private Sub stbThis_PanelClick(ByVal Panel As MSComctlLib.Panel)
    If Panel.index = 7 Then
        With picSerialInfo
            .Left = Me.ScaleWidth - .Width - 30
            .Top = Me.ScaleHeight - stbThis.Height - .Height - 30
            .Visible = True
            .SetFocus
        End With
    End If
End Sub

Private Sub txtFact_Change()
    lblFact.Tag = ""
End Sub

Private Sub txtSN_Change()
    If mblnNotChange Then Exit Sub
    If mblnUnChange Then Exit Sub
    If Trim(txtSN.Text) = "" Then Exit Sub
    If vsfPlan.TextMatrix(vsfPlan.Row, GetCol("号别")) <> txt号别.Text And txt号别.Text <> "" Then
        If mlngPreRow <> vsfPlan.Row And mlngPreRow < vsfPlan.Rows And mlngPreRow <> 0 Then
            vsfPlan.Row = mlngPreRow
            lblSN.Tag = "序号不能全选" '输入值时引起的改变不能再全选
            Call zlControl.ControlSetFocus(txtSN)
        End If
    End If
End Sub

Private Sub txt发生时间_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub vsfPay_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim str应缴 As String
    If Col = 1 Then
        str应缴 = txt本次应缴.Text
        txt本次应缴.Text = Format(Val(txt本次应缴.Text) + mdbl原金额 - Val(vsfPay.TextMatrix(Row, 1)), "0.00")
        If txt本次应缴.Text < mcur应缴 Or Val(vsfPay.TextMatrix(Row, 1)) > Val(vsfPay.TextMatrix(Row, 6)) Then
            txt本次应缴.Text = str应缴
            vsfPay.TextMatrix(Row, 1) = Format(mdbl原金额, "0.00")
        Else
            vsfPay.TextMatrix(Row, 1) = Format(Val(vsfPay.TextMatrix(Row, 1)), "0.00")
        End If
    End If
    If Col = 2 Then
        If zlCommFun.ActualLen(vsfPay.TextMatrix(Row, 2)) > 30 Then
            MsgBox "结算号码过长,请重新填写!", vbInformation, gstrSysName
            vsfPay.TextMatrix(Row, 2) = ""
        End If
    End If
End Sub

Private Sub vsfPay_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If vsfPay.TextMatrix(Row, 0) = "" Then Cancel = True
    If mblnCenter Then Cancel = True: Exit Sub
    If mbytMode = 4 Or chkCancel.Value = 1 Then
        '退号
        If Col = 0 And vsfPay.TextMatrix(Row, Col) <> "" Then
            vsfPay.ComboList = "..."
            vsfPay.CellButtonPicture = imgDel
        Else
            Cancel = True
        End If
        Exit Sub
    End If
    If mbytInState = 1 Then Cancel = True: Exit Sub
    If Col = 1 Then
        If Val(vsfPay.TextMatrix(Row, vsfPay.ColIndex("金额修改"))) <> 1 Then
            Cancel = True
        Else
            vsfPay.ComboList = ""
            mdbl原金额 = Val(vsfPay.TextMatrix(Row, 1))
        End If
    End If
    
    If Col = 0 Then
        If mbln连续挂号 Then
            Cancel = True
        Else
            If Val(vsfPay.TextMatrix(Row, vsfPay.ColIndex("修改"))) = 0 Then
                vsfPay.ComboList = "..."
                vsfPay.CellButtonPicture = imgDel
            Else
                Cancel = True
            End If
        End If
    End If
    
    If Col = 2 Then
        If Val(vsfPay.RowData(Row)) = 2 Then
            vsfPay.ComboList = ""
        Else
            If Val(vsfPay.RowData(Row)) <> 8 And Val(vsfPay.RowData(Row)) <> 7 And vsfPay.TextMatrix(Row, 0) Like "*卡*" Then
                vsfPay.ComboList = ""
            Else
                Cancel = True
            End If
        End If
    End If
End Sub

Private Sub vsfPay_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim dblMoney As Double, i As Integer, blnFind As Boolean, rsTemp As ADODB.Recordset
    Dim strSQL As String, str退款操作员 As String
    On Error GoTo errH
    If vsfPay.TextMatrix(Row, Col) = "" Then Exit Sub
    If mbytMode = 4 Or chkCancel.Value = 1 Then
    
        If (Val(vsfPay.RowData(Row)) = 7 Or Val(vsfPay.RowData(Row)) = 8) And Val(vsfPay.TextMatrix(Row, vsfPay.ColIndex("修改"))) = 1 Then
            If InStr(mstrCardPrivs, ";三方退款强制退现;") = 0 Then
                str退款操作员 = zlDatabase.UserIdentifyByUser(Me, "强制退现验证", glngSys, 1151, "三方退款强制退现")
                If str退款操作员 = "" Then
                    MsgBox "录入的操作员验证失败或者录入的操作员不具备强制退现权限，不能强制退现！", vbInformation, gstrSysName
                    Exit Sub
                End If
                mstrForceNote = str退款操作员 & "强制退现:" & vsfPay.TextMatrix(Row, vsfPay.ColIndex("结算方式")) & "," & vsfPay.TextMatrix(Row, 1) & "元"
            Else
                If MsgBox(vsfPay.TextMatrix(Row, 0) & "不支持退现，是否强制退现？", vbQuestion + vbYesNo, gstrSysName) <> vbYes Then Exit Sub
                mstrForceNote = UserInfo.姓名 & "强制退现:" & vsfPay.TextMatrix(Row, vsfPay.ColIndex("结算方式")) & "," & vsfPay.TextMatrix(Row, 1) & "元"
            End If
        End If
    
        dblMoney = Val(vsfPay.TextMatrix(Row, 1))
        vsfPay.RemoveItem Row
        blnFind = False
        For i = 1 To vsfPay.Rows - 1
            If vsfPay.RowData(i) = 1 Then
                vsfPay.TextMatrix(i, 1) = Format(Val(vsfPay.TextMatrix(i, 1)) + dblMoney, "0.00")
                blnFind = True
            End If
            If blnFind Then Exit For
        Next i
        If blnFind = False Then
            strSQL = "Select 名称 From 结算方式 Where 性质=1 Order By 缺省标志 Desc"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
            For i = 1 To vsfPay.Rows - 1
                If vsfPay.TextMatrix(i, 0) = "" Then
                    blnFind = True
                    vsfPay.TextMatrix(i, 0) = Nvl(rsTemp!名称)
                    vsfPay.TextMatrix(i, 1) = Format(dblMoney, "0.00")
                    vsfPay.TextMatrix(i, vsfPay.ColIndex("修改")) = "1"
                    vsfPay.RowData(i) = 1
                End If
                If blnFind Then Exit For
            Next i
            
            If blnFind = False Then
                vsfPay.Rows = vsfPay.Rows + 1
                vsfPay.TextMatrix(vsfPay.Rows - 1, 0) = Nvl(rsTemp!名称)
                vsfPay.TextMatrix(vsfPay.Rows - 1, 1) = Format(dblMoney, "0.00")
                vsfPay.TextMatrix(vsfPay.Rows - 1, vsfPay.ColIndex("修改")) = "1"
                vsfPay.RowData(vsfPay.Rows - 1) = 1
            End If
        End If
    Else
        dblMoney = Val(vsfPay.TextMatrix(Row, 1))
        txt本次应缴.Text = Format(Val(txt本次应缴.Text) + dblMoney, "0.00")
        vsfPay.RemoveItem Row
    End If
    Exit Sub
errH:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub vsfPlan_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim i As Integer
    If vsfPlan.Visible And Me.ActiveControl Is txtSN Then vsfPlan.SetFocus
    With vsfPlan
        If OldRow < vsfPlan.Rows Then
            If OldRow Mod 2 = 1 Then
                For i = 1 To .Cols - 1
                    If .Cell(flexcpBackColor, OldRow, i, OldRow, i) <> &HFF8080 Then .Cell(flexcpBackColor, OldRow, i, OldRow, i) = &H80000005
                Next i
            Else
                For i = 1 To .Cols - 1
                    If .Cell(flexcpBackColor, OldRow, i, OldRow, i) <> &HFF8080 Then .Cell(flexcpBackColor, OldRow, i, OldRow, i) = &HF6F6F6
                Next i
            End If
        End If
        For i = 1 To .Cols - 1
            If .Cell(flexcpBackColor, NewRow, i, NewRow, i) <> &HFF8080 Then .Cell(flexcpBackColor, NewRow, i, NewRow, i) = 16772055
        Next i
    End With
End Sub

Private Sub vsfPlan_AfterSort(ByVal Col As Long, Order As Integer)
    Dim i As Integer, j As Integer
    Dim lngColor As Long
    With vsfPlan
        For j = 1 To .Rows - 1
            If j Mod 2 = 1 Then
                lngColor = &H80000005
            Else
                lngColor = &HF6F6F6
            End If
            For i = 1 To .Cols - 1
                If .Cell(flexcpBackColor, j, i) <> &HFF8080 Then .Cell(flexcpBackColor, j, i) = lngColor
            Next i
        Next j
        For i = 1 To .Cols - 1
            If .Cell(flexcpBackColor, .Row, i, .Row, i) <> &HFF8080 Then .Cell(flexcpBackColor, .Row, i, .Row, i) = 16772055
        Next i
    End With
End Sub

Private Sub vsfplan_DblClick()
    If vsfPlan.MouseRow > 0 Then Call vsfPlan_KeyDown(13, 0)
End Sub

Private Sub SetvsfplanColor()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置挂号号别颜色
    '编制:刘兴洪
    '日期:2010-02-04 14:13:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim PreRedaw As RedrawSettings '
'    PreRedaw = vsfplan.Redraw
'    vsfplan.Redraw = flexRDNone
'    vsfplan.Cell(flexcpBackColor, vsfplan.Row, 0, vsfplan.Row, vsfplan.Cols - 1) = vsfplan.BackColor
'    vsfplan.Cell(flexcpForeColor, vsfplan.Row, 0, vsfplan.Row, vsfplan.Cols - 1) = vsfplan.ForeColor
'    vsfplan.Redraw = PreRedaw
'
End Sub

Private Sub SetvsfplanFiexBackColor(Optional blnCurDate As Boolean = True)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置相关固定列的背景色
    '参数:blnCurDate-是否当前日期列,否则就是预约日期列
    '编制:刘兴洪
    '日期:2010-02-04 14:39:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim PreRedaw As RedrawSettings, i As Long, strSQL As String, strNow As String
    Dim strKey As String, rsTmp As ADODB.Recordset, strColor As String
    Dim j As Long, strPrevKey As String
    Dim DatCur As Date
    With vsfPlan
         .Redraw = flexRDNone
         If blnCurDate Then
             strKey = zlGet当前星期几
             DatCur = zlDatabase.Currentdate
             strPrevKey = zlGet当前星期几(Format(DatCur - 1, "yyyy-mm-dd"))
             For i = 1 To .Rows - 1
                If Format(DatCur, "yyyy-mm-dd") = Format(.TextMatrix(i, .ColIndex("出诊日期")), "yyyy-mm-dd") Then
                    .Cell(flexcpData, 0, .ColIndex(strKey)) = 1 '当前日期
                    .Cell(flexcpBackColor, i, .ColIndex(strKey), i, .ColIndex(strKey)) = &HFF8080
                    .Cell(flexcpFontBold, i, .ColIndex(strKey), i, .ColIndex(strKey)) = True
                Else
                    .Cell(flexcpData, 0, .ColIndex(strPrevKey)) = 1 '昨天日期
                    .Cell(flexcpBackColor, i, .ColIndex(strPrevKey), i, .ColIndex(strPrevKey)) = &HFF8080
                    .Cell(flexcpFontBold, i, .ColIndex(strPrevKey), i, .ColIndex(strPrevKey)) = True
                End If
             Next i
             
            strColor = zlDatabase.GetPara("提前挂号颜色", glngSys, mlngModul, "0")
            strNow = Format(DatCur, "YYYY-MM-DD HH:MM:SS")
            For i = 1 To .Rows - 1
                If .Cell(flexcpData, i, .ColIndex("号别")) = "1" Then
                    For j = 1 To .Cols - 1
                        If .Cell(flexcpData, 0, j) <> 1 Then
                            .Cell(flexcpForeColor, i, j, i, j) = &H8000000C
                        End If
                    Next j
                Else
                    If .TextMatrix(i, .ColIndex("提前时间")) <> "" Then
                        If strNow < Format(.TextMatrix(i, .ColIndex("挂号时间")), "YYYY-MM-DD HH:MM:SS") Then
                            For j = 1 To .Cols - 1
                                If .Cell(flexcpData, 0, j) <> 1 Then
                                    .Cell(flexcpForeColor, i, j, i, j) = strColor
                                End If
                            Next j
                        End If
                    End If
                End If
            Next i
        Else
            DatCur = dtpAppointmentDate.Value
            strKey = zlGet当前星期几(Format(dtpAppointmentDate.Value, "yyyy-mm-dd"))
            strPrevKey = zlGet当前星期几(Format(dtpAppointmentDate.Value - 1, "yyyy-mm-dd"))
            If .ColIndex(strKey) < 0 Then Exit Sub
            For i = 1 To .Cols - 1
                If i <> .ColIndex(mstr当前星期) Then  '以前预约日期列
                    For j = 1 To .Rows - 1
                        If j Mod 2 = 1 Then
                            .Cell(flexcpBackColor, j, i, j, i) = &H80000005
                        Else
                            .Cell(flexcpBackColor, j, i, j, i) = &HF6F6F6
                        End If
                    Next j
                     .Cell(flexcpFontBold, 1, i, .Rows - 1, i) = False
                ElseIf Val(.ColData(.ColIndex(strKey))) = 1 Then    '当前日期的星期几列
                Else
                    .Cell(flexcpData, 0, i) = ""
                    For j = 1 To .Rows - 1
                        If j Mod 2 = 1 Then
                            .Cell(flexcpBackColor, j, i, j, i) = &H80000005
                        Else
                            .Cell(flexcpBackColor, j, i, j, i) = &HF6F6F6
                        End If
                    Next j
                    .Cell(flexcpFontBold, 1, i, .Rows - 1, i) = False
                End If
            Next
            For i = 1 To .Rows - 1
                If Format(DatCur, "yyyy-mm-dd") = Format(.TextMatrix(i, .ColIndex("出诊日期")), "yyyy-mm-dd") Then
                    .Cell(flexcpData, 0, .ColIndex(strKey)) = 1 '当前日期
                    .Cell(flexcpBackColor, i, .ColIndex(strKey), i, .ColIndex(strKey)) = &HFF8080
                    .Cell(flexcpFontBold, i, .ColIndex(strKey), i, .ColIndex(strKey)) = True
                Else
                    .Cell(flexcpData, 0, .ColIndex(strPrevKey)) = 1 '昨天日期
                    .Cell(flexcpBackColor, i, .ColIndex(strPrevKey), i, .ColIndex(strPrevKey)) = &HFF8080
                    .Cell(flexcpFontBold, i, .ColIndex(strPrevKey), i, .ColIndex(strPrevKey)) = True
                End If
             Next i
'            .ColData(.ColIndex(strKey)) = "2"
'            .Cell(flexcpBackColor, 1, .ColIndex(strKey), .Rows - 1, .ColIndex(strKey)) = &HFF8080
'            .Cell(flexcpFontBold, 1, .ColIndex(strKey), .Rows - 1, .ColIndex(strKey)) = True
            If .Rows > 1 Then
                .Cell(flexcpForeColor, 1, GetCol("IDS"), .Rows - 1, .Cols - 1) = vbBlack
            End If
        End If
        mstrCurKey = strKey
        .Redraw = flexRDBuffered
    End With
End Sub

Private Sub SetSnStyle(Optional ByVal bln分时段 As Boolean = False)
'****************************************
'对表格样式进行设置
'****************************************
    Dim i           As Long
    Dim lngWidth    As Long
    Dim X           As Long
    Dim Y           As Long
    Dim j           As Long
    Dim lngHeight   As Long
    Select Case bln分时段
    Case False:
        With vsfList
            
            .FixedCols = 0
            lngWidth = 570
            lngHeight = 375
            For i = 0 To vsfList.Cols - 1
                .ColWidth(i) = lngWidth
                .ColAlignment(i) = 4
            Next
            For i = 0 To vsfList.Rows - 1
                 .RowHeight(i) = lngHeight
            Next
            
        End With
    
    Case True:
        With vsfList
             If .Cols <= 1 Then Exit Sub
             .FixedCols = 1
             .FixedAlignment(0) = flexAlignRightTop
             .ColAlignment(0) = flexAlignRightTop
            lngWidth = 1275
            lngHeight = 800
            For i = 1 To vsfList.Cols - 1
                .ColWidth(i) = lngWidth
                .ColAlignment(i) = 4
            Next
            .ColAlignment(0) = 3
            .ColWidth(0) = lngWidth
            For i = 0 To vsfList.Rows - 1
                 .RowHeight(i) = lngHeight
            Next
           If .Rows > 0 And .Cols > 0 Then
                .Cell(flexcpFontBold, 0, 1, .Rows - 1, .Cols - 1) = True
                .Cell(flexcpFontSize, 0, 1, .Rows - 1, .Cols - 1) = 9
                .Cell(flexcpFontSize, 0, 0, .Rows - 1, 0) = 18
           End If
        End With
    End Select
   If vsfList.Rows >= 1 And vsfList.Cols > 0 Then
       vsfList.Cell(flexcpFontBold, 0, 0, vsfList.Rows - 1, vsfList.Cols - 1) = True
    End If
End Sub

Private Sub LoadTimePlan()
    '***************************************
    '加载时间段
    '***************************************
    Dim i               As Integer
    Dim j               As Integer
    Dim blnPre          As Boolean
    Dim lngThis         As Long
    Dim lngMax          As Long
    Dim datThis         As Date
    Dim lngCurrSn       As Long
    Dim lngMaxSn        As Long '预约的最大使用号
    Dim strSQL          As String
    Dim rs时段统计      As ADODB.Recordset
    Dim str时间点       As String
    Dim lng预约人数     As Long
    Dim lngTatol        As Long '用于分时段 最后重新计算行数
    Dim strMaxDate      As String  '用于分时段保存大预约时间
    Dim lngCols         As Long
    Dim lngRows         As Long
    Dim strData         As String
    Dim strDate         As String
    Dim lng记录ID       As Long
    Dim blnHave         As Boolean
    Dim datMax          As Date
    Dim Datsys          As Date
    Dim bln失约用于挂号 As Boolean
    Dim blnInserted     As Boolean
    Dim lng合作单位人数 As Long
    Dim blnFindSN      As Boolean '是否需要重新定位到上次号别的序号,用于刷新列表时,数据保持
    Dim lngFindSN      As Long '需要查找的序号
    Dim str替诊         As String
    str替诊 = "(替)"
    vsfList.Visible = True
    picSplit.Visible = True
    vsfList.Redraw = False
    mblnStateChange = True
    vsfList.Clear
    '***************************************
    '表格信息设置
    '***************************************
    If dkpMain.Panes(2).Hidden Then
        dkpMain.Panes(2).Hidden = False
        mcbrToolBar.Controls.Find(xtpControlButton, 2605).Visible = InStr(1, mstrPrivs, ";预留号码;") > 0 '36294
        mcbrToolBar.Controls.Find(xtpControlButton, 2604).Visible = InStr(1, mstrPrivs, ";预留号码;") > 0 '36294
        Call Form_Resize
    Else
        mcbrToolBar.Controls.Find(xtpControlButton, 2605).Visible = InStr(1, mstrPrivs, ";预留号码;") > 0 '36294
        mcbrToolBar.Controls.Find(xtpControlButton, 2604).Visible = InStr(1, mstrPrivs, ";预留号码;") > 0 '36294
        Call Form_Resize
    End If
    If mbytMode = 1 Then
        lngMax = Val(vsfPlan.TextMatrix(vsfPlan.Row, GetCol("限约")))
    Else
        lngMax = Val(vsfPlan.TextMatrix(vsfPlan.Row, GetCol("限号"))) '挂将来的号不当成预约,因为已交费,应当成挂号
    End If
    If mbytMode = 1 Then
        lngMax = Val(vsfPlan.TextMatrix(vsfPlan.Row, GetCol("限号")))
    End If
    
    '1.调整位置
    If lngMax > 1000 Then
        vsfList.FontWidth = 4
    Else
        vsfList.FontWidth = 0 '恢复缺省字体
    End If
    '***************************************
    '初始化时间段
    '***************************************
     If InitTimePlan() = False Then vsfList.Redraw = True: Exit Sub
     Datsys = zlDatabase.Currentdate
    '***************************************
    '初始化表格
    '***************************************
     
     If mrs时间段 Is Nothing Then vsfList.Redraw = True: Exit Sub
     'If mrs时间段.RecordCount = 0 Then Exit Sub
 
    '***************************************
    '序号填充
    '***************************************
     With vsfList
        .Rows = 1
        .Cols = 1
        .Clear
     End With
     lngCurrSn = -1
     If mstrPre号别 <> "" Then
        blnFindSN = mstrPre号别 = mtyRegPlanState.str号别
        blnFindSN = blnFindSN And mViewMode = v_专家号分时段 And txtSN.Text <> ""
        If blnFindSN Then lngFindSN = Val(txtSN.Text)
     End If
    Select Case mViewMode
    Case V_普通号分时段:
       
        strSQL = "Select Count(1) As 预约数量, To_Char(开始时间, 'HH24:MI') As 日期" & vbNewLine & _
                "From 临床出诊序号控制" & vbNewLine & _
                "Where Nvl(挂号状态,0) <> 0 And Nvl(是否预约,0) = 1 And 预约顺序号 Is Not Null And 记录id = [1]" & vbNewLine & _
                "Group By 开始时间"
        
        lng记录ID = vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("记录ID"))
        On Error GoTo Hd
        Set rs时段统计 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng记录ID)
        blnHave = False
        
        str时间点 = ""
        With mrs时间段
          datMax = CDate("00:00:00")
          mdatLast = CDate("00:00:00")
          lngRows = -1: lngCols = 0
           Do While Not .EOF
                If IsNull(!预约顺序号) Then
                    If datMax < CDate(Nvl(!开始时间, "00:00:00")) Then datMax = CDate(!开始时间)
                    If mdatLast < CDate(Nvl(!结束时间, "00:00:00")) Then mdatLast = CDate(!结束时间)
                    '预约状态 只填充允许预约的时间段
                    '挂号时不区分都填充
                     rs时段统计.Filter = " 日期='" & Nvl(!开始时间, "_") & "'"
                     If rs时段统计.RecordCount = 0 Then
                        lng预约人数 = 0
                     Else
                        lng预约人数 = rs时段统计!预约数量
                     End If
                     
                     lng合作单位人数 = 0
                     If mbytMode = 1 And mblnUnitReg And Not mrsUnitReg Is Nothing Then
                         mrsUnitReg.Filter = "序号=" & Val(Nvl(!序号))
                         lng合作单位人数 = 0
                         If mrsUnitReg.RecordCount > 0 Then
                            lng合作单位人数 = Val(Nvl(mrsUnitReg!数量))
                         End If
                     End If
                      
                     If Nvl(!限制数量, 0) <> 0 Then
                        If str时间点 <> Nvl(!时间点) Then
                            lngRows = lngRows + 1
                            str时间点 = Nvl(!时间点)
                            If lngRows > vsfList.Rows - 1 Then vsfList.Rows = vsfList.Rows + 1: lngCols = 0
                            If lngCols > vsfList.Cols - 1 Then vsfList.Cols = vsfList.Cols + 1
                            vsfList.TextMatrix(lngRows, 0) = str时间点
                         End If
                        lngCols = lngCols + 1
                        If lngCols > vsfList.Cols - 1 Then vsfList.Cols = vsfList.Cols + 1
                        lng预约人数 = Nvl(!限制数量, 0) - lng预约人数 - lng合作单位人数
                        If vsfPlan.Cell(flexcpData, vsfPlan.Row, vsfPlan.ColIndex("替诊医生")) <> "" And _
                            Format(!详细开始时间, "yyyy-mm-dd hh:mm:ss") >= Format(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("替诊开始时间")), "yyyy-mm-dd hh:mm:ss") And _
                            Format(!详细开始时间, "yyyy-mm-dd hh:mm:ss") <= Format(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("替诊终止时间")), "yyyy-mm-dd hh:mm:ss") Then
                          strData = "预约" & IIf(lng预约人数 < 0, 0, lng预约人数) & "人" & str替诊 & vbCrLf & _
                                                !开始时间 & "-" & !结束时间
                        Else
                          strData = "预约" & IIf(lng预约人数 < 0, 0, lng预约人数) & "人" & vbCrLf & _
                                                !开始时间 & "-" & !结束时间
                        End If
                        vsfList.TextMatrix(lngRows, lngCols) = strData
                        If lng预约人数 <= 0 Then
                             vsfList.Cell(flexcpForeColor, lngRows, lngCols) = vbGreen
                        End If
                        If Format(Datsys, "yyyy-mm-dd") <= Format(dtpAppointmentDate + 1, "yyyy-mm-dd") Then
                              If Format(DateAdd("n", mTy_Para.lng预约限制时间, Datsys), "yyyy-mm-dd hh:mm:ss") > Format(!详细结束时间, "yyyy-mm-dd hh:mm:ss") Then
                                vsfList.Cell(flexcpForeColor, lngRows, lngCols) = vbGrayText
                              End If
                        End If
                        If Format(!详细结束时间, "yyyy-mm-dd") <> Format(dtpAppointmentDate.Value, "yyyy-mm-dd") And dtpAppointmentDate.Visible Then
                            vsfList.Cell(flexcpForeColor, lngRows, lngCols) = vbGrayText
                        End If
                     End If
                 End If
                .MoveNext
          Loop
          .Filter = ""
        End With
        Set rs时段统计 = Nothing
    Case v_专家号分时段:
     '*******************************
     '专家号分时段
     '每行以时间点区分
     '*******************************
     
regHD:
        blnInserted = False
        str时间点 = ""
        With mrs时间段
          mtyRegPlanState.lngLastNO = 0
          lngRows = -1: lngCols = 0
           datMax = CDate("00:00:00")
           Do While Not .EOF
                 If datMax < CDate(Nvl(!开始时间, "00:00:00")) Then datMax = CDate(!开始时间)
                '预约状态 只填充允许预约的时间段
                '挂号时不区分都填充
                If blnFindSN Then
                    If Val(Nvl(!序号)) = lngFindSN And lngFindSN > 0 Then
                          lngCurrSn = lngFindSN
                    End If
                End If
'                If (mbytMode = 1 And Nvl(!是否预约, 0) = 1 Or blnHave) Or mbytMode <> 1 Then
                '78643:李南春,2014/10/16,挂号处预约的挂号安排如果设置了预约号段，只显示预约时段部分
                If ((mbytMode = 1 Or (mbytMode = 0 And chkBooking.Visible And chkBooking.Value = 1)) And Nvl(!是否预约, 0) = 1 Or blnHave) Or _
                    Not (mbytMode = 1 Or (mbytMode = 0 And chkBooking.Visible And chkBooking.Value = 1)) Then
                    If str时间点 <> Nvl(!时间点) Then
                        lngRows = lngRows + 1
                        str时间点 = Nvl(!时间点)
                        If lngRows > vsfList.Rows - 1 Then vsfList.Rows = vsfList.Rows + 1: lngCols = 0
                        If lngCols > vsfList.Cols - 1 Then vsfList.Cols = vsfList.Cols + 1
                        vsfList.TextMatrix(lngRows, 0) = str时间点
                        vsfList.Cell(flexcpForeColor, lngRows, 0, lngRows, 0) = vsfPlan.Cell(flexcpForeColor, vsfPlan.Row, 0, vsfPlan.Row, 0)
                     End If
                    lngCols = lngCols + 1
                      If lngCols > vsfList.Cols - 1 Then vsfList.Cols = vsfList.Cols + 1
                    If vsfPlan.Cell(flexcpData, vsfPlan.Row, vsfPlan.ColIndex("替诊医生")) <> "" And _
                        Format(!详细开始时间, "yyyy-mm-dd hh:mm:ss") >= Format(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("替诊开始时间")), "yyyy-mm-dd hh:mm:ss") And _
                        Format(!详细开始时间, "yyyy-mm-dd hh:mm:ss") <= Format(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("替诊终止时间")), "yyyy-mm-dd hh:mm:ss") Then
                      strData = !序号 & str替诊 & vbCrLf & !开始时间 & "-" & !结束时间
                    Else
                      strData = !序号 & vbCrLf & !开始时间 & "-" & !结束时间
                    End If
                    vsfList.TextMatrix(lngRows, lngCols) = strData
                    
                    Select Case mbytMode
                    Case 0:
                        If chkBooking.Visible And chkBooking.Value = 1 Then
                            If Format(Datsys, "yyyy-mm-dd") <= Format(dtpAppointmentDate + 1, "yyyy-mm-dd") Then
                               If (Format(DateAdd("n", mTy_Para.lng预约限制时间, Datsys), "yyyy-mm-dd hh:mm:ss") > Format(!详细开始时间, "yyyy-mm-dd hh:mm:ss")) Then
                                   vsfList.Cell(flexcpForeColor, lngRows, lngCols) = vbGrayText
                               End If
                             End If
                        ElseIf (Format(Datsys, "yyyy-mm-dd hh:mm:ss") > Format(!详细开始时间, "yyyy-mm-dd hh:mm:ss") And mbytMode = 0) Then
                             vsfList.Cell(flexcpFontUnderline, lngRows, lngCols) = True
                             vsfList.Cell(flexcpForeColor, lngRows, lngCols) = vbGrayText
                        End If
                    Case 1:
                        If Format(Datsys, "yyyy-mm-dd") <= Format(dtpAppointmentDate + 1, "yyyy-mm-dd") Then
                            If (Format(DateAdd("n", mTy_Para.lng预约限制时间, Datsys), "yyyy-mm-dd hh:mm:ss") > Format(!详细开始时间, "yyyy-mm-dd hh:mm:ss")) Then
                                vsfList.Cell(flexcpForeColor, lngRows, lngCols) = vbGrayText
                            End If
                        End If
                    Case Else:
                    End Select
                    If Format(!详细结束时间, "yyyy-mm-dd") <> Format(dtpAppointmentDate.Value, "yyyy-mm-dd") And dtpAppointmentDate.Visible Then
                        vsfList.Cell(flexcpForeColor, lngRows, lngCols) = vbGrayText
                    End If
                End If
                
                '把设置的最大的序号保存到mtyRegPlanState中 方便做对比或者条件限制 'lgf
                If mtyRegPlanState.lngLastNO < Val(Nvl(!序号)) Then
                    With mtyRegPlanState
                        .lngLastNO = Val(Nvl(mrs时间段!序号))
                        .lngLastNO_X = lngRows
                        .lngLastNO_Y = lngCols
                    End With
                    
                End If
                
                .MoveNext
          Loop
          If blnHave = False And vsfList.Rows = 1 And vsfList.Cols = 1 And mrs时间段.RecordCount > 0 Then blnHave = True: mrs时间段.MoveFirst: GoTo regHD
          
          '获取最后一个时段的序号,开始时间,结束时间 'lgf
          mrs时间段.Filter = 0
          If mrs时间段.RecordCount > 0 And mtyRegPlanState.lngLastNO > 0 Then
                mrs时间段.Filter = "序号=" & mtyRegPlanState.lngLastNO
                If mrs时间段.RecordCount > 0 Then
                    mtyRegPlanState.strLastNO_Time = Nvl(!开始时间)
                    mtyRegPlanState.strLastNo_EndTime = Nvl(!结束时间)
                End If
                mrs时间段.Filter = 0
          End If
          If InStr(mstrPrivs, ";加号;") > 0 And mbytMode = 0 Then
            If (Format(Nvl(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("出诊日期")), "3000-01-01"), "yyyy-mm-dd") = Format(dtpAppointmentDate.Value, "yyyy-MM-dd")) Or dtpAppointmentDate.Visible = False Then
                .MoveLast
                For i = 1 To vsfList.Cols - 1
                    If vsfList.TextMatrix(vsfList.Rows - 1, i) = "" Then
                        If blnInserted = False Then
                            vsfList.TextMatrix(vsfList.Rows - 1, i) = " " & vbCrLf & !结束时间 & "以后"
                            vsfList.Cell(flexcpData, vsfList.Rows - 1, i) = "加号"
                            blnInserted = True
                        End If
                    End If
                Next i
                If blnInserted = False Then
                    vsfList.Cols = vsfList.Cols + 1
                    vsfList.TextMatrix(vsfList.Rows - 1, vsfList.Cols - 1) = " " & vbCrLf & !结束时间 & "以后"
                    vsfList.Cell(flexcpData, vsfList.Rows - 1, vsfList.Cols - 1) = "加号"
                End If
            End If
          End If
        End With
    End Select
    dtpAppointmentTime.Tag = Format(datMax, "hh:mm:ss")
    '***************************************
    '序号表格状态设置
    '***************************************
    Call SetSnStyle(True)
    '***************************************
    '序号状态 填充
    '现在挂号状态需要填充的只有一种状态
    '***************************************
     If mViewMode = v_专家号分时段 Then
        If picBookingDate.Visible Or mbytMode = 1 Or mbytMode = 2 Or (mbytMode = 0 And mbytInState = 0 And mstrNoIn <> "") Then             '预约或接收时的日期
            datThis = CDate(Format(dtpAppointmentDate.Value, "yyyy-MM-dd"))
        Else
            datThis = zlDatabase.Currentdate
        End If
         
         If mTy_Para.bln失约用于挂号 Then
            '专家号分时段时  失约的序号用于开放出来挂号
            bln失约用于挂号 = True
            Datsys = DateAdd("n", -1 * mTy_Para.lng预约有效时间, Datsys)
         End If
        
        Set mrsSNState = GetSNState(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("记录ID")))

        If mrsSNState.RecordCount > 0 Then
                For i = 0 To vsfList.Rows - 1
                   For j = 1 To vsfList.Cols - 1
                       If vsfList.TextMatrix(i, j) <> "" And Not vsfList.Cell(flexcpData, i, j) Like "加*" Then
                        '**********************************************
                        '
                        '**********************************************
                          vsfList.Row = i: vsfList.Col = j
                          lngFindSN = Val(Get时段(i, j, False))
                          mrsSNState.Filter = "序号=" & lngFindSN
                          If mrsSNState.RecordCount > 0 Then
                            If lngCurrSn = lngFindSN Then lngCurrSn = -1
                            Select Case mrsSNState!状态
                            Case 1  '已挂
                                  If Nvl(mrsSNState!预约, "0") = "0" Then
                                    vsfList.Cell(flexcpForeColor, i, j) = vbRed
                                  Else
                                    vsfList.Cell(flexcpForeColor, i, j) = &HC000C0
                                  End If
                                  vsfList.Cell(flexcpFontStrikethru, i, j) = True
                            Case 2  '已约
                                vsfList.Cell(flexcpForeColor, i, j) = vbGreen
                            If lngMaxSn < Val(Nvl(mrsSNState!序号)) Then
                                lngMaxSn = Val(Nvl(mrsSNState!序号))
                            End If
                            Case 3  '已留
                              vsfList.Cell(flexcpForeColor, i, j) = vbBlue
                            Case 4  '退号
'                                If mTy_Para.blnReuseCancelNO = False Then
                                    vsfList.Cell(flexcpForeColor, i, j) = vbGrayText
                                    vsfList.Cell(flexcpFontStrikethru, i, j) = True
'                                End If
                            Case 5  '锁号
                                vsfList.Cell(flexcpForeColor, i, j) = vbRed
                            Case 6  '停诊
                                vsfList.Cell(flexcpForeColor, i, j) = vbGrayText
                            End Select
                          End If
                       End If
                   Next
                Next
            
        End If
           If mbytMode = 1 And mblnUnitReg And Not mrsUnitReg Is Nothing Then
            For i = 0 To vsfList.Rows - 1
                For j = 1 To vsfList.Cols - 1
                    If Trim(vsfList.TextMatrix(i, j)) <> "" Then
                        mrsUnitReg.Filter = "序号=" & Get时段(i, j, False)
                        If mrsUnitReg.RecordCount > 0 Then vsfList.Cell(flexcpForeColor, i, j) = &HC000C0
                    End If
                Next
            Next
            mrsUnitReg.Filter = 0
        End If
     End If
     '还有可用序号的情况下，屏蔽加号栏
    If CheckAddAvailable = False Then
        For i = 0 To vsfList.Rows - 1
            For j = 1 To vsfList.Cols - 1
                If vsfList.Cell(flexcpData, i, j) Like "加*" Then
                    vsfList.Cell(flexcpData, i, j) = ""
                    vsfList.TextMatrix(i, j) = ""
                End If
            Next j
        Next i
    End If
    If vsfList.Rows > 1 Then
       vsfList.Cell(flexcpFontBold, 0, 0, vsfList.Rows - 1, 0) = True
    End If
     
    Me.dtpAppointmentTime.Value = Format(Me.dtpAppointmentTime.Tag, "hh:mm:ss")
    vsfList.Redraw = True
    locateSnBy时段 lngCurrSn, True
    mblnStateChange = False
    Exit Sub
Hd:
    If ErrCenter() = 1 Then
        Resume
    End If
    SaveErrLog
End Sub

Private Sub locateSnBy时段(Optional ByVal lngSN As Long = -1, _
    Optional bln强制定位 As Boolean)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:定位到指定的时段
    '入参:lngSN:>0需要定位的序号上,-1:表示按规则取数
    '出参:bln强制定位-强制定位到指定的数据列上
    '编制:刘兴洪
    '日期:2013-12-07 13:01:55
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, j As Long, lngRow As Long, lngCol As Long
    Dim blnFind  As Boolean, blnExit As Boolean, blnMaxSn As Boolean
    Dim lngLastRow As Long, lngLastCol As Long
     lngRow = 0: lngCol = 1
     
    vsfList.HighLight = flexHighlightAlways
    Select Case mViewMode
    Case V_普通号分时段:
         '****************************
         '普通号分时段 序号定位
         '****************************
         vsfList.Redraw = False
         blnMaxSn = True
          For i = 0 To vsfList.Rows - 1
            For j = 1 To vsfList.Cols - 1
                With vsfList
                    If .TextMatrix(i, j) <> "" Then
                        If .Cell(flexcpForeColor, i, j) <> vbGrayText Then
                            If Val(Get时段(i, j, False)) > 0 Then
                                     blnFind = True
                                     lngRow = i: lngCol = j: Exit For
                            End If
                        End If
                        lngLastRow = i
                        lngLastCol = j
                    End If
                End With
            Next
            If blnFind Then Exit For
          Next
         If blnFind Then
           vsfList.Row = lngRow: vsfList.Col = lngCol
            If vsfList.Row > 1 Then
                If vsfList.RowIsVisible(vsfList.Row) = False Then
                     vsfList.TopRow = vsfList.Row - 1
                End If
            End If
        Else
            vsfList.Row = lngLastRow: vsfList.Col = lngLastCol
            If vsfList.Row > 1 Then
                If vsfList.RowIsVisible(vsfList.Row) = False Then
                     vsfList.TopRow = vsfList.Row - 1
                End If
            End If
           vsfList.HighLight = flexHighlightAlways
        End If
        
        dtpAppointmentTime.Value = IIf(blnFind, CDate(Get时段(lngRow, lngCol, True)), CDate(mdatLast))
        vsfList.Redraw = True
    Case v_专家号分时段:
        blnMaxSn = True
        With vsfList
            For i = 0 To .Rows - 1
                For j = 1 To .Cols - 1
                    If .TextMatrix(i, j) <> "" Then
                        '预留
                        If .Cell(flexcpForeColor, i, j) = vbBlue Then
                            If lngSN <> -1 Then
                                 If lngSN = Val(Get时段(i, j, False)) Then
                                    .Row = i: .Col = j
                                     blnFind = True
                                     lngRow = i: lngCol = j
                                     blnMaxSn = False
                                     dtpAppointmentTime.Value = CDate(Get时段(i, j, True))
                                     blnExit = True: Exit For
                                 End If
                             End If
                        End If
                         If .Cell(flexcpForeColor, i, j) <> vbRed _
                             And .Cell(flexcpForeColor, i, j) <> vbBlue _
                             And .Cell(flexcpForeColor, i, j) <> vbGrayText Then
                             
                            If blnMaxSn = True _
                                And .Cell(flexcpForeColor, i, j) <> vbGreen _
                                And .Cell(flexcpForeColor, i, j) <> &HC000C0 Then
                                If Not mTy_Para.bln随机序号选择 Or lngSN = -1 Then  '66788
                                    blnFind = True
                                    lngRow = i: lngCol = j
                                    blnMaxSn = False
                                    If mbytMode <> 1 Then
                                        blnExit = True: Exit For  '45768
                                    End If
                                End If
                             End If
                             
                             If lngSN <> -1 Then
                                 If lngSN = Val(Get时段(i, j, False)) Then
                                    .Row = i: .Col = j
                                     blnFind = True
                                    lngRow = i: lngCol = j
                                    blnMaxSn = False
                                     dtpAppointmentTime.Value = CDate(Get时段(i, j, True))
                                     blnExit = True: Exit For
                                 End If
                             End If
                         End If
                    End If
                Next
                If blnExit Then Exit For '45768
            Next
        End With
        
        If blnFind And blnMaxSn = False Then
            If bln强制定位 Then mblnNotClick = True
            vsfList.Row = lngRow: vsfList.Col = lngCol
            mblnNotClick = False
        Else
            vsfList.HighLight = flexHighlightAlways
        End If
        If blnFind = False And blnMaxSn And Me.dtpAppointmentTime.Tag <> "" Then
            dtpAppointmentTime.Value = Format(CDate(Me.dtpAppointmentTime.Tag), "hh:mm:ss")
        Else
            dtpAppointmentTime.Value = Format(CDate(Get时段(lngRow, lngCol, True)), "hh:mm:ss")
            If dtpAppointmentDate.Visible Then
                txt发生时间.Text = Format(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("出诊日期")), "yyyy-MM-dd") & " " & Format(Get时段(lngRow, lngCol, True), "hh:mm:ss")
            Else
                txt发生时间.Text = Format(zlDatabase.Currentdate, "yyyy-MM-dd") & " " & Format(Get时段(lngRow, lngCol, True), "hh:mm:ss")
            End If
        End If
        If bln强制定位 = False Then Call vsfList_DblClick
    Case Else: Exit Sub
    End Select
End Sub
Private Function Get时段(ByVal lngRow As Long, ByVal lngCol As Long, Optional ByVal blnTime As Boolean = False, Optional ByVal blnLastTime As Boolean = False) As String
    '*****************************************************************
    '功能说明:在挂号专家号分时时 获取 序号,或者 开始时间
    '参数:  blntime 是否获取时间 是则获取时间  否则返回序号
    '*****************************************************************
    Dim strResult       As String, i As Long
    If lngRow > vsfList.Rows - 1 Or lngCol > vsfList.Cols - 1 Then
        Exit Function
    End If
     If vsfList.TextMatrix(lngRow, lngCol) = "" Then
        Exit Function
    End If
    
    If blnTime Then
        i = IIf(blnLastTime = False, 0, 1)
        If InStr(vsfList.TextMatrix(lngRow, lngCol), "-") > 0 Then
            Get时段 = Split(Split(Replace(vsfList.TextMatrix(lngRow, lngCol), "(替)", ""), vbCrLf)(1), "-")(i)
        Else
            Get时段 = Split(Split(Replace(vsfList.TextMatrix(lngRow, lngCol), "(替)", ""), vbCrLf)(1), "以")(i)
        End If
        Exit Function
    End If
    If mViewMode = v_专家号分时段 Then
       strResult = Split(Replace(vsfList.TextMatrix(lngRow, lngCol), "(替)", ""), vbCrLf)(0)
    ElseIf mViewMode = V_普通号分时段 Then
       strResult = Replace(Replace(Split(Replace(vsfList.TextMatrix(lngRow, lngCol), "(替)", ""), vbCrLf)(0), "预约", ""), "人数", "")
    End If
    Get时段 = strResult
End Function

Private Sub ClearRegState()
    Dim strSQL As String, rsTemp As ADODB.Recordset
    '初始化状态变量信息
    'lgf 2012-10-30
   '初始化状态变量信息
    With mtyRegPlanState
        .str号别 = "" '选中的号别
        .lngLastNO = 0 '最后的一个序号
        .strLastNO_Time = "" '最后一个时段开始时间
        .strLastNo_EndTime = "" '最有一个时段结束时间
        .blnAdditionalNumber = False '是否已经追加序号 '追加序号的特点(挂出去的序号,序号大于设置的最大序号,或者时间大于或者等于,最后一个时段的结束时间)
        .lngSelX = 0 '选中的行
        .lngSelY = 0 '选中的列
        .lngSelNO = 0  '选中的序号
        .strSelTime = ""   '选中的序号对应时段的开始时间
        .bln序号控制 = False    '序号控制
        .lng限号数 = 0             '限号数
        .lng限约数 = 0             '限约数
        .lngLastNO_X = 0 '最后一个序号的位置
        .lngLastNO_Y = 0
        '.lngPlanRow = 0 '号别所在行
    End With
    '73767
    If mTy_Para.bln失约用于挂号 = True And mTy_Para.lng预约有效时间 <> 0 Then
        '问题号:110549,焦博,2017/07/21,SQL性能问题
        strSQL = "Select 1" & vbNewLine & _
                " From 病人挂号记录 A, 临床出诊序号控制 B" & vbNewLine & _
                " Where a.预约时间 < Sysdate + 1 / 24 / 60 * " & mTy_Para.lng预约有效时间 & " And a.预约时间 > Trunc(Sysdate) And a.记录性质 = 2 And" & vbNewLine & _
                "       a.出诊记录Id = b.记录Id And a.出诊记录Id = [1] And (a.号序 = b.序号 Or to_Char(a.号序) = b.备注) And Nvl(b.挂号状态,0) = 2 And rownum < 2"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("记录ID"))))
        If Not rsTemp.EOF Then
            Call zlDatabase.ExecuteProcedure("zl_挂号序号状态_出诊_DELETE(" & Val(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("记录ID"))) & ")", Me.Caption)
        End If
    End If
End Sub
 
Private Sub vsfPlan_EnterCell()
    Dim i           As Integer
    Dim j           As Integer
    Dim blnPre      As Boolean
    Dim lngThis     As Long
    Dim lngMax      As Long
    Dim datThis     As Date
    Dim lngCurrSn   As Long
    Dim lngMaxSn    As Long '预约的最大使用号
    Dim lngRow      As Long
    Dim lngCol      As Long
    Dim blnChk      As Boolean
    Dim sngTime     As Single
    Dim DatCur      As Date
    If Me.Visible = False Then GoTo regTab
    
    '125595:李南春，2018/5/16，预约接收出诊记录定位问题
    If mblnChangeByCode Then mlngRow = vsfPlan.Row: Exit Sub
    sngTime = Timer
    If Format(sngTime, "0.000") - Format(msngTime, "0.000") < 0.1 And mblnManualInput = False Then
        mblnChangeByCode = True
        If mlngRow <> 0 Then vsfPlan.Select mlngRow, vsfPlan.ColIndex("IDS")
        mblnChangeByCode = False
        Exit Sub
    End If
    msngTime = Timer
    mlngRow = vsfPlan.Row
    
    Call SetvsfplanColor
    '接收时仍要处理,但不显示,因为可能需要修改序号
    If mbytInState <> 0 Then
        Exit Sub
    End If
   
    dtpAppointmentTime.MaxDate = CDate("23:59:59")
    dtpAppointmentTime.MinDate = CDate("00:00:00")
    
    DatCur = zlDatabase.Currentdate
    If mbytMode = 1 Or chkBooking.Value = 1 Then
        txt发生时间.Text = Format(Format(dtpAppointmentDate.Value, "yyyy-mm-dd" & " " & Format(dtpAppointmentTime.Value, "hh:mm:ss")), "yyyy-mm-dd hh:mm:ss")
    Else
        txt发生时间.Text = Format(DatCur, "yyyy-mm-dd hh:mm:ss")
    End If
    
    
    '暂时只处理分时段这种情况,主要处理,分时段中各个时间,例如时段的序号和时段的时间对不上等情况,
    '初始化变量信息
    Call ClearRegState
    
    mtyRegPlanState.str号别 = vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("号别"))
    
    '*****************************
    '获取使用那种流程处理挂号
    '******************************
    If mcustomTime = t_时段 Then
         GetActiveView
         If mcustomTime = t_普通 Then
            dtpAppointmentTime.Enabled = False
            dtpAppointmentTime.Visible = False
         Else
           If (mViewMode = V_普通号分时段 Or mViewMode = v_专家号分时段) Then
               dtpAppointmentTime.Enabled = False
              
           ElseIf (mbytMode = 1 Or (chkBooking.Visible And chkBooking.Value = 1)) And (mViewMode = V_普通号 Or mViewMode = v_专家号) Then
                dtpAppointmentTime.Enabled = True
                Call SetDefaultRegistTime
           ElseIf mbytMode = 0 Then
               dtpAppointmentTime.Enabled = False
           End If
           
         End If
        If mViewMode = V_普通号分时段 Or mViewMode = v_专家号分时段 Then
           If mbytMode = 1 And mblnUnitReg Then
                '如果是预约同时分配了挂号合作单位信息的话序号先加载 合作单位号信息
                LoadUnitReg (Val(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("记录ID"))))
            End If
           '*************************************************
           '如果存在分时段的情况 使用分时段的处理方法
           '*************************************************
           LoadTimePlan
           SetDefaultRegistTime
           Call picPlan_Resize
'           vsfList.Height = picPlan.ScaleHeight - vsfList.Top - 350
'           vsfPlan.Height = picPlan.ScaleHeight - IIf(picBookingDate.Visible, picBookingDate.Height + 30, 0) - 360 - IIf(vsfList.Visible, vsfList.Height + picSplit.Height + 30, 0)
'           Call locateSnBy时段(, True)
           vsfPlan.ShowCell vsfPlan.Row, vsfPlan.Col
           If InStr(vsfList.TextMatrix(vsfList.Row, vsfList.Col), "替") > 0 Then
                vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("替诊医生")) = vsfPlan.Cell(flexcpData, vsfPlan.Row, vsfPlan.ColIndex("替诊医生"))
           Else
                vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("替诊医生")) = ""
           End If
           Exit Sub
        End If
    Else
         If vsfPlan.TextMatrix(vsfPlan.Row, GetCol("序号控制")) <> "" Then
                mViewMode = v_专家号
         Else
                mViewMode = V_普通号
         End If
    End If
    
    If vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("替诊开始时间")) <> "" And Not (mViewMode = V_普通号分时段 Or mViewMode = v_专家号分时段) Then
        If CDate(txt发生时间.Text) >= CDate(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("替诊开始时间"))) And CDate(txt发生时间.Text) <= CDate(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("替诊终止时间"))) Then
            vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("替诊医生")) = vsfPlan.Cell(flexcpData, vsfPlan.Row, vsfPlan.ColIndex("替诊医生"))
        Else
            vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("替诊医生")) = ""
        End If
    End If
    
    If mbytMode = 1 And mblnUnitReg Then
        '如果是预约同时分配了挂号合作单位信息的话序号先加载 合作单位号信息
        LoadUnitReg (Val(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("记录ID"))))
    End If
    vsfList.Redraw = False
    vsfList.Clear
    If mbytMode = 1 Then
        lngMax = Val(vsfPlan.TextMatrix(vsfPlan.Row, GetCol("限约")))
        If lngMax = 0 Then lngMax = Val(vsfPlan.TextMatrix(vsfPlan.Row, GetCol("限号")))
    Else
        lngMax = Val(vsfPlan.TextMatrix(vsfPlan.Row, GetCol("限号"))) '挂将来的号不当成预约,因为已交费,应当成挂号
    End If
    If lngMax > 0 And vsfPlan.TextMatrix(vsfPlan.Row, GetCol("序号控制")) <> "" Then
        If mbytMode = 1 Then
              lngMax = Val(vsfPlan.TextMatrix(vsfPlan.Row, GetCol("限号")))  '预约开放出来,用户选择:问题????
        End If
        If lngMax = 0 Then GoTo regTab
        '1.调整位置
        If lngMax > 1000 Then
            vsfList.FontWidth = 4
        Else
            vsfList.FontWidth = 0 '恢复缺省字体
        End If
        'mblnNotClick = True
        If (lngMax \ SNCOLS) * SNCOLS = lngMax Then
            vsfList.Rows = lngMax \ SNCOLS
        Else
            vsfList.Rows = lngMax \ SNCOLS + 1
        End If
        'mblnNotClick = False
        vsfList.Cols = SNCOLS
        If dkpMain.Panes(2).Hidden Then
            dkpMain.Panes(2).Hidden = False
            mcbrToolBar.Controls.Find(xtpControlButton, 2605).Visible = InStr(1, mstrPrivs, ";预留号码;") > 0
            mcbrToolBar.Controls.Find(xtpControlButton, 2604).Visible = InStr(1, mstrPrivs, ";预留号码;") > 0
            Call Form_Resize
        Else
            mcbrToolBar.Controls.Find(xtpControlButton, 2605).Visible = InStr(1, mstrPrivs, ";预留号码;") > 0
            mcbrToolBar.Controls.Find(xtpControlButton, 2604).Visible = InStr(1, mstrPrivs, ";预留号码;") > 0
            Call Form_Resize
        End If
                                
        '填充序号
        lngThis = 1
        For i = 0 To vsfList.Rows - 1
            For j = 0 To vsfList.Cols - 1
                vsfList.TextMatrix(i, j) = lngThis
                lngThis = lngThis + 1
                If lngThis > lngMax Then Exit For
            Next
            If lngThis > lngMax Then Exit For
        Next
             
        If picBookingDate.Visible Or mbytMode = 1 Or mbytMode = 2 Or (mbytMode = 0 And mbytInState = 0 And mstrNoIn <> "") Then             '预约或接收时的日期
            datThis = CDate(Format(dtpAppointmentDate.Value, "yyyy-MM-dd"))
        Else
            datThis = CDate(Format(zlDatabase.Currentdate, "yyyy-MM-dd"))
        End If
        
        
        Set mrsSNState = GetSNState(Val(vsfPlan.TextMatrix(vsfPlan.Row, GetCol("记录ID"))))
        lngMaxSn = 0
        For i = 0 To mrsSNState.RecordCount - 1
            If mrsSNState!序号 <= lngMax Then
                If (mrsSNState!序号 \ SNCOLS) * SNCOLS = mrsSNState!序号 Then
                   lngRow = (mrsSNState!序号 \ SNCOLS) - 1
                   lngRow = IIf(lngRow < 0, 0, lngRow) '问题号:51843
                Else
                    lngRow = (mrsSNState!序号 \ SNCOLS)
                End If
                    lngCol = (mrsSNState!序号 - 1) Mod SNCOLS
                    lngCol = IIf(lngCol < 0, 0, lngCol) '问题号:51843
                Select Case mrsSNState!状态
                    Case 1  '已挂
                       If Nvl(mrsSNState!预约, "0") = "0" Then
                          vsfList.Cell(flexcpForeColor, lngRow, lngCol) = vbRed
                          '用于序号定位最大的有效号后
                          If lngMaxSn < Val(Nvl(mrsSNState!序号)) Then
                            lngMaxSn = Val(Nvl(mrsSNState!序号))
                          End If
                       Else
                          '预约接收
                          vsfList.Cell(flexcpForeColor, lngRow, lngCol) = &HC000C0
                       End If
                    Case 2  '已约
                          vsfList.Cell(flexcpForeColor, lngRow, lngCol) = vbGreen
                    Case 3  '已留
                      vsfList.Cell(flexcpForeColor, lngRow, lngCol) = vbBlue
                    Case 4  '退号
'                        If mTy_Para.blnReuseCancelNO = False Then
                            vsfList.Cell(flexcpForeColor, lngRow, lngCol) = vbGrayText
                            vsfList.Cell(flexcpFontStrikethru, lngRow, lngCol) = True
'                        End If
                    Case 5  '锁号
                        vsfList.Cell(flexcpForeColor, lngRow, lngCol) = vbRed
                End Select
            End If
            mrsSNState.MoveNext
        Next
        
        If mbytMode = 1 And mblnUnitReg And Not mrsUnitReg Is Nothing Then
            For i = 0 To vsfList.Rows - 1
                For j = 0 To vsfList.Cols - 1
                    If Trim(vsfList.TextMatrix(i, j)) <> "" Then
                        mrsUnitReg.Filter = "序号=" & vsfList.TextMatrix(i, j)
                        If mrsUnitReg.RecordCount > 0 Then
                            vsfList.Cell(flexcpForeColor, i, j) = &HC000C0
                            If lngMaxSn < Val(Trim(vsfList.TextMatrix(i, j))) Then lngMaxSn = Val(Trim(vsfList.TextMatrix(i, j)))
                        End If
                    End If
                Next
            Next
            mrsUnitReg.Filter = 0
        End If
        
        If Trim(txtSN.Text) = "" Then  '定时刷新时保持已输的不变
           lngCurrSn = GetCurrSN(IIf(mbytMode = 0, lngMaxSn, -1))
           txtSN.Text = lngCurrSn
        Else
            lngCurrSn = Val(txtSN.Text)
            '处理问题编号：38779
            If lngMax < lngCurrSn Then lngCurrSn = GetCurrSN(IIf(mbytMode = 1, lngMaxSn, -1))
        End If
    Else
regTab:
        If mbytMode = 0 Or mbytMode = 1 Then
            mblnUnChange = True
            txtSN.Tag = ""
            txtSN.Text = ""
            mblnUnChange = False
        End If
        Set mrsSNState = Nothing
        vsfList.Visible = False
        picSplit.Visible = False
        mcbrToolBar.Controls.Find(xtpControlButton, 2605).Visible = False
        mcbrToolBar.Controls.Find(xtpControlButton, 2604).Visible = False
        Call Form_Resize
    End If
    vsfList.Redraw = True
    SetSnStyle
    Call LocateSN(lngCurrSn)
    Call picPlan_Resize
End Sub

Private Sub LoadUnitReg(ByVal lng记录ID As Long)
 '加载挂号合作单位控制信息
    Dim strSQL As String
        
    strSQL = "Select 名称 As 合作单位, 控制方式, 序号, 数量 From 临床出诊挂号控制记录 Where 记录id = [1] And 类型 = 1"
    
    On Error GoTo Hd
    Set mrsUnitReg = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng记录ID)
    Exit Sub
Hd:
    If ErrCenter() = 1 Then
        Resume
    End If
    SaveErrLog
End Sub

Private Sub LocateSN(lngCurrSn As Long)
'功能:定位到指定序号上
'     如果不是在输号别或序号,则序号表获得焦点
    Dim lngRow          As Long
    Dim i               As Long
    Dim j               As Long
    Dim blnHave         As Boolean
    If lngCurrSn = 0 Then Exit Sub
   
    If mViewMode = V_普通号 Or mViewMode = v_专家号 Then
        '************************************************
        '不分时段 序号定位还是按照以前的方式
        '************************************************
        If (lngCurrSn \ SNCOLS) * SNCOLS = lngCurrSn Then
            lngRow = (lngCurrSn - 1) \ SNCOLS
        Else
            lngRow = (lngCurrSn \ SNCOLS)
        End If
        If Not vsfList.RowIsVisible(lngRow) Then
            If lngRow >= 1 Then  '保留上一行可见
                vsfList.TopRow = lngRow - 1
            Else
                vsfList.TopRow = lngRow
            End If
        End If
        '问题号:52335
        mblnNotClick = True
        vsfList.Row = lngRow
        vsfList.RowSel = vsfList.Row
        vsfList.Col = (lngCurrSn - 1) Mod SNCOLS
        vsfList.ColSel = vsfList.Col
        '问题号:52335
        mblnNotClick = False
     
    ElseIf mViewMode = v_专家号分时段 Then
        '*******************************************
        '专家号分时段 序号定位
        '*******************************************
        For i = 0 To vsfList.Rows - 1
            For j = 1 To vsfList.Cols - 1
               If vsfList.TextMatrix(i, j) <> "" Then
                    If lngCurrSn = Val(Get时段(i, j, False)) Then
                     If Not vsfList.RowIsVisible(i) Then
                        If lngRow >= 1 Then  '保留上一行可见
                             vsfList.TopRow = i - 1
                        Else
                             vsfList.TopRow = i
                        End If
                      End If
 
                      vsfList.Row = i
                      vsfList.Col = j
                  
'                     vsflist.ColSel = vsflist.Col
'                     vsflist.RowSel = vsflist.Row
                     blnHave = True
                     dtpAppointmentTime.Value = CDate(Get时段(i, j, True))
                     Exit For
                      
                     
                    End If
                End If
            Next
            If blnHave Then Exit For
        Next
    End If
    Call vsfList_EnterCell
    If vsfList.Visible And vsfList.Enabled _
                And Not Me.ActiveControl Is txt号别 And Not Me.ActiveControl Is txtSN _
                And Not Me.ActiveControl Is dtpAppointmentDate And Not Me.ActiveControl Is vsfPlan Then Call vsfList.SetFocus     '焦点在号别正在连续输入
End Sub

Private Function GetSNState(lng记录ID As Long, Optional lngSN As Long) As ADODB.Recordset
    Dim strSQL           As String
    On Error GoTo errH

    strSQL = "    " & vbNewLine & " Select A.序号,Decode(是否停诊,1,6,Nvl(A.挂号状态,0)) As 状态,A.操作员姓名,Decode(A.挂号状态,2,1,0) as 预约,To_Char(B.出诊日期,'hh24:mi:ss') as 日期  "
    strSQL = strSQL & vbNewLine & " From 临床出诊序号控制 A, 临床出诊记录 B "
    strSQL = strSQL & vbNewLine & " Where B.ID=[1] And B.ID=A.记录ID"
    strSQL = strSQL & vbNewLine & IIf(lngSN > 0, " And A.序号=[2]", "")
    Set GetSNState = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng记录ID, lngSN)

    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Sub vsfPlan_LeaveCell()
    Call SetvsfplanColor
End Sub

Private Sub vsfPlan_KeyDown(KeyCode As Integer, Shift As Integer)
    '选择号别进行挂号
    If KeyCode = 13 Then
        
        If CheckNoValied(vsfPlan.Row) = False Then
             txt号别.Text = "": txt号别.SetFocus: Exit Sub
        End If
        vsfPlan.Tag = vsfPlan.TextMatrix(vsfPlan.Row, GetCol("号别"))
        If txt号别.Visible And txt号别.Enabled Then txt号别.SetFocus
        If txt号别.Text = vsfPlan.Tag Then
            Call txt号别_Change
        Else
            txt号别.Text = vsfPlan.Tag
        End If
    vsfPlan.Tag = ""
    Call locateSnBy时段(, True)
'    If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus
    End If
End Sub

Private Sub vsfplan_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If vsfPlan.MouseRow = 0 Then
        vsfPlan.MousePointer = flexCustom
    Else
        vsfPlan.MousePointer = flexArrow
    End If
End Sub

Private Sub vsfplan_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim intCol As Integer, intRow As Integer
    
    If mTy_Para.bln点击列头排序 = False Then Exit Sub
    intCol = vsfPlan.MouseCol
    intRow = vsfPlan.MouseRow
    If intRow = 1 And intCol >= 1 And intCol <= vsfPlan.Cols - 1 Then
'        If vsfPlan.ColData(intCol) = "" Then Exit Sub
'        vsfPlan.ColData(intCol) = (Val(vsfPlan.ColData(intCol)) + 1) Mod 2
'        mstrSort = vsfPlan.TextMatrix(1, intCol) & IIf(vsfPlan.ColData(intCol) = 1, " Desc", "")
'        Call ShowPlans(mstrSort)
    End If
End Sub

Private Sub vsfplan_SelChange()
    If vsfPlan.Rows = 2 Then Exit Sub
    vsfPlan.RowSel = vsfPlan.Row
End Sub

Private Function CheckAddAvailable() As Boolean
'-----------------------------------------------------------------------------------------------------------------------
'功能:检查当前选择的号别加号是否可用
'返回:可用返回True,不可用返回False
'编制:刘尔旋
'日期:2014-01-15
'备注:
'-----------------------------------------------------------------------------------------------------------------------
    Dim i As Integer, j As Integer
    Dim intTotal As Integer, intUse As Integer
    If vsfList.Visible = False Then Exit Function
    intTotal = 0
    intUse = 0
    '只对分时段进行处理
    If mViewMode = V_普通号分时段 Or mViewMode = v_专家号分时段 Or mViewMode = V_普通号 Or mViewMode = v_专家号 Then
        With vsfList
            For j = 1 To .Cols - 1
                For i = 0 To .Rows - 1
                    If .TextMatrix(i, j) <> "" And Not .Cell(flexcpData, i, j) Like "加*" Then
                        intTotal = intTotal + 1
                        If .Cell(flexcpForeColor, i, j) <> vbBlack Then
                            intUse = intUse + 1
                        End If
                    End If
                Next i
            Next j
        End With
        If intUse = intTotal Then CheckAddAvailable = True: Exit Function
        CheckAddAvailable = False
        Exit Function
    End If
End Function

Private Sub vsfList_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewRow > vsfList.Rows - 1 Or NewCol > vsfList.Cols - 1 Then Exit Sub
End Sub

Private Sub vsfList_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    If mblnStateChange Then Exit Sub
    '问题号:52203
    '问题号:52335
   
    If mblnNotClick Then Exit Sub
    If (mViewMode = V_普通号分时段 Or mViewMode = v_专家号分时段 Or mViewMode = v_专家号) And mTy_Para.bln随机序号选择 = False _
        And Not (mbytMode = 1 Or chkBooking.Value = 1 And chkBooking.Visible) And vsfList.Cell(flexcpForeColor, NewRow, NewCol) <> vbBlue Then
        Cancel = True
        Exit Sub
    End If
    If vsfList.TextMatrix(NewRow, NewCol) = "" Then Cancel = True
    If vsfList.Cell(flexcpForeColor, NewRow, NewCol) <> vbBlack And vsfList.Cell(flexcpForeColor, NewRow, NewCol) <> vbBlue Then Cancel = True
    If Not CheckAddAvailable And mbytMode = 0 Then
        If vsfList.Cell(flexcpData, NewRow, NewCol) Like "加*" Then Cancel = True
    End If
'    'vsflist.Cell(flexcpBackColor, OldRow, OldCol) = vbWhite
'    'vsflist.Cell(flexcpBackColor, NewRow, NewCol) = &HECBAAA
End Sub

Private Sub vsfList_DblClick()
    Dim lngSN       As Long
    Dim datThis     As Date
    Dim strTmp      As String
    If mViewMode = V_普通号 Or mViewMode = v_专家号 Then
        '*************************************************
        '普通号和没有分时段的专家号 按照以前处理方法
        '*************************************************
        lngSN = Val(vsfList.TextMatrix(vsfList.Row, vsfList.Col))
        If Not mrsSNState Is Nothing And lngSN > 0 Then
            mrsSNState.Filter = "序号=" & lngSN & " And 状态 <> 0"
            If mrsSNState.RecordCount > 0 Then
                If mrsSNState!状态 = 3 And mrsSNState!操作员姓名 = UserInfo.姓名 Then
                    '自已预留的可以直接用来挂号
                    vsfPlan.Tag = vsfPlan.TextMatrix(vsfPlan.Row, GetCol("号别"))
                    txt号别.Text = vsfPlan.Tag
                    txtSN.Text = lngSN
                    mstrPre号别 = txt号别.Text
                    mlngPreRow = vsfPlan.Row
                    vsfPlan.Tag = ""
                  If mcustomTime = t_普通 Or dtpAppointmentTime.Enabled = False Then
                    If txtSN.Enabled And txtSN.Visible Then txtSN.SetFocus
                  ElseIf dtpAppointmentTime.Visible And dtpAppointmentTime.Enabled Then
                     dtpAppointmentTime.SetFocus
                  End If
                    If txtSN.Enabled And txtSN.Visible Then
                        txtSN.SetFocus
                    Else
                        If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
                    End If
                End If
            Else
                If vsfList.CellForeColor = &HC000C0 Then Exit Sub
                vsfPlan.Tag = vsfPlan.TextMatrix(vsfPlan.Row, GetCol("号别"))
                txt号别.Text = vsfPlan.Tag
                txtSN.Text = lngSN
                vsfPlan.Tag = ""
                mstrPre号别 = txt号别.Text
                mlngPreRow = vsfPlan.Row
                If mcustomTime = t_普通 Or dtpAppointmentTime.Enabled = False Then
                    If txtSN.Enabled And txtSN.Visible Then txtSN.SetFocus
                ElseIf dtpAppointmentTime.Visible And dtpAppointmentTime.Enabled Then
                     dtpAppointmentTime.SetFocus
                End If
                If txtSN.Enabled And txtSN.Visible Then
                    txtSN.SetFocus
                Else
                    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
                End If
            End If
        End If
        Exit Sub
    End If
    
    If mViewMode = V_普通号 Or mViewMode = v_专家号 Then Exit Sub
    
    '*************************************************
    '分时段 按照新的方式来处理
    '*************************************************
    
    Select Case mViewMode
    Case V_普通号分时段:
        If vsfList.CellForeColor = vbGrayText Then Exit Sub
        If vsfList.TextMatrix(vsfList.Row, vsfList.Col) = "" Then Exit Sub
        If Val(Get时段(vsfList.Row, vsfList.Col, False)) = 0 Then Exit Sub
        strTmp = Format(dtpAppointmentDate.Value, "yyyy-MM-dd") & " " & Get时段(vsfList.Row, vsfList.Col, True)
        txt发生时间.Text = Format(strTmp, "yyyy-mm-dd hh:mm:ss")
        datThis = CDate(Format(strTmp, "hh:mm:ss"))
        dtpAppointmentTime.Value = datThis
        dtpAppointmentTime.Tag = strTmp
        vsfPlan.Tag = vsfPlan.TextMatrix(vsfPlan.Row, GetCol("号别"))
        txt号别.Text = vsfPlan.Tag
        txtSN.Text = ""
        vsfPlan.Tag = ""
        
        '保存序号
        mtyRegPlanState.lngSelNO = 0
        mtyRegPlanState.lngSelX = vsfList.Row
        mtyRegPlanState.lngSelY = vsfList.Col
        mtyRegPlanState.strSelTime = Get时段(vsfList.Row, vsfList.Col, True)
        mstrPre号别 = txt号别.Text
        mlngPreRow = vsfPlan.Row
        If InStr(vsfList.TextMatrix(vsfList.Row, vsfList.Col), "替") > 0 Then
            vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("替诊医生")) = vsfPlan.Cell(flexcpData, vsfPlan.Row, vsfPlan.ColIndex("替诊医生"))
        Else
            vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("替诊医生")) = ""
        End If
        If txtSN.Enabled And txtSN.Visible Then
            txtSN.SetFocus
        Else
            If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
        End If
    Case v_专家号分时段:
        '**********************************************
        '如果序号为已挂或者已约的不允许选择
        '
        '**********************************************
        If vsfList.TextMatrix(vsfList.Row, vsfList.Col) = "" Then Exit Sub
        If vsfList.CellForeColor = vbRed Or vsfList.CellForeColor = vbGreen Or vsfList.CellForeColor = vbGrayText Or vsfList.CellForeColor = &HC000C0 Then Exit Sub  '--And .CellForeColor <> vbBlue
        If dtpAppointmentDate.Visible Then
            strTmp = Format(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("出诊日期")), "yyyy-MM-dd") & " " & Get时段(vsfList.Row, vsfList.Col, True)
        Else
            strTmp = Format(zlDatabase.Currentdate, "yyyy-MM-dd") & " " & Get时段(vsfList.Row, vsfList.Col, True)
        End If
        txt发生时间.Text = Format(strTmp, "yyyy-mm-dd hh:mm:ss")
        datThis = CDate(strTmp)
        dtpAppointmentTime.Value = Get时段(vsfList.Row, vsfList.Col, True)
        dtpAppointmentTime.Tag = strTmp
        vsfPlan.Tag = vsfPlan.TextMatrix(vsfPlan.Row, GetCol("号别"))
        txt号别.Text = vsfPlan.Tag
        
        mblnNotChange = True
        txtSN.Text = Get时段(vsfList.Row, vsfList.Col, False)
        If txtSN.Text = "加号" Then txtSN.Text = ""
        mtyRegPlanState.lngSelNO = Val(txtSN.Text)
        mtyRegPlanState.lngLastNO_X = vsfList.Row
        mtyRegPlanState.lngLastNO_Y = vsfList.Col
        mtyRegPlanState.strSelTime = Get时段(vsfList.Row, vsfList.Col, True)
        mblnNotChange = False
        
        mstrPre号别 = txt号别.Text
        mlngPreRow = vsfPlan.Row
        vsfPlan.Tag = ""
        If InStr(vsfList.TextMatrix(vsfList.Row, vsfList.Col), "替") > 0 Then
            vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("替诊医生")) = vsfPlan.Cell(flexcpData, vsfPlan.Row, vsfPlan.ColIndex("替诊医生"))
        Else
            vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("替诊医生")) = ""
        End If
        If txtSN.Enabled And txtSN.Visible Then txtSN.SetFocus
    Case Else
        Exit Sub
    End Select
     
End Sub

Private Sub vsfList_EnterCell()
'处理是否允许预留
    '***************************************
    '这里处理预留号
    '预留号处理情况为
    '专家号不分时段 以前的处理方式
    '专家号 分时段 新处理方式
    '普通号分时段 不允许预留
    '***************************************
    If mViewMode = V_普通号分时段 Then
        mcbrToolBar.Controls.Find(xtpControlButton, 2605).Enabled = False
        mcbrToolBar.Controls.Find(xtpControlButton, 2604).Enabled = False
        Exit Sub
    End If
    If vsfList.Row <> -1 Then
         '问题号:52335
         If vsfList.Cols > vsfList.Col And vsfList.Rows > vsfList.Row Then
            If vsfList.TextMatrix(vsfList.Row, vsfList.Col) <> "" Then
              ' vsflist.CellBackColor = &HECBAAA
                'vsflist.Cell(flexcpBackColor, vsflist.Row, vsflist.Col) = &HECBAAA
            Else
                Exit Sub
            End If
         End If
    End If
    mcbrToolBar.Controls.Find(xtpControlButton, 2605).Enabled = True
    mcbrToolBar.Controls.Find(xtpControlButton, 2604).Enabled = False
    If Not mrsSNState Is Nothing Then
        '问题号:52335
        If vsfList.Cols > vsfList.Col And vsfList.Rows > vsfList.Row Then
            Select Case mViewMode
            Case v_专家号:
                mrsSNState.Filter = "序号=" & Val(vsfList.TextMatrix(vsfList.Row, vsfList.Col))
            Case v_专家号分时段:
                mrsSNState.Filter = "序号=" & Val(Get时段(vsfList.Row, vsfList.Col, False))
            End Select
        End If
        If mrsSNState.RecordCount > 0 Then
            mrsSNState.MoveFirst
            If Val(Nvl(mrsSNState!状态)) = 3 Then
                If mrsSNState!状态 = 3 And mrsSNState!操作员姓名 = UserInfo.姓名 Then
                    '取消预留
                    mcbrToolBar.Controls.Find(xtpControlButton, 2605).Enabled = False
                    mcbrToolBar.Controls.Find(xtpControlButton, 2604).Enabled = True
                Else
                    mcbrToolBar.Controls.Find(xtpControlButton, 2605).Enabled = False
                    mcbrToolBar.Controls.Find(xtpControlButton, 2604).Enabled = False
                    '64184:刘尔旋,2014-03-20,选择预留号码
                    If Me.ActiveControl Is vsfList Then
                        Select Case mViewMode
                            Case v_专家号:
                                MsgBox Val(vsfList.TextMatrix(vsfList.Row, vsfList.Col)) & "号已被" & mrsSNState!操作员姓名 & "预留!无法选择.", vbInformation, gstrSysName
                            Case v_专家号分时段:
                                MsgBox Val(Get时段(vsfList.Row, vsfList.Col, False)) & "号已被" & mrsSNState!操作员姓名 & "预留!无法选择.", vbInformation, gstrSysName
                        End Select
                        txt号别_KeyPress (13)
                    End If
                End If
            End If
        End If
    Else
        mcbrToolBar.Controls.Find(xtpControlButton, 2605).Enabled = False
        mcbrToolBar.Controls.Find(xtpControlButton, 2604).Enabled = False
    End If
    If mcbrToolBar.Controls.Find(xtpControlButton, 2605).Enabled Then
        If vsfList.Row >= vsfList.Rows Then Exit Sub
        If vsfList.Col >= vsfList.Cols Then Exit Sub
        If vsfList.Cell(flexcpForeColor, vsfList.Row, vsfList.Col) <> vbBlack Then mcbrToolBar.Controls.Find(xtpControlButton, 2605).Enabled = False
    End If
End Sub

Private Sub vsflist_KeyDown(KeyCode As Integer, Shift As Integer)
     If mTy_Para.bln随机序号选择 Then Exit Sub
     If KeyCode <> 13 Then KeyCode = 0
End Sub

Private Sub vsflist_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call vsfList_DblClick
End Sub

Private Sub picPatiPicBack_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MovePatiPic
End Sub

Private Sub txtIDCard_Change()
        txtIDCard.Tag = ""
End Sub

Private Sub txtIDCard_GotFocus()
    zlControl.TxtSelAll txtIDCard
End Sub

Private Sub txtIDCard_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txtIDCard_Validate(Cancel As Boolean)
    Dim strbirthday As String, strAge As String, strSex As String, strErrInfo As String, strInfo As String
    
    On Error GoTo errH
    If chkCancel.Value = 1 Then Exit Sub
    If txtIDCard.Tag = txtIDCard.Text Then Exit Sub
    If Trim(txtIDCard.Text) = "" Then Exit Sub
    
    '81103,冉俊明,2014-12-26,录入身份证号后,出生日期、年龄、性别的同步关联检查和调整
    If txtIDCard.Visible And txtIDCard.Enabled And Not mobjfrmPatiInfo.mobjPubPatient Is Nothing Then
        'CheckPatiIdcard(ByVal strIdcard As String, Optional strBirthday As String, _
        '    Optional strAge As String, Optional strSex As String, Optional strErrInfo As String) As Boolean
        '功能：身份证号码合法性校验
        '入参：strIdCard 身份证号码
        '出参：strBirthday  函数返回True为出生日期
        '         strAge 函数返回True为年龄
        '         strSex 函数返回True为性别
        '         strErrInfo 函数返回False为错误信息
        '返回：True/False  身份证合法返回True(可从strBirthday，strSex获取出生日期和性别)，
        '       否则返回False(可从strErrInfo获取详细错误信息)
        If mobjfrmPatiInfo.mobjPubPatient.CheckPatiIdcard(Trim(txtIDCard.Text), strbirthday, strAge, strSex, strErrInfo) Then
            '新病人或调整无业务数据的已有病人信息时提示是否调整不一致的基本信息
            If strSex <> NeedName(cbo性别.Text) Then strInfo = "性别"
            If strAge <> Trim(txt年龄.Text) & cbo年龄单位 Then strInfo = strInfo & IIf(strInfo = "", "年龄", "、年龄")
            
            If strInfo <> "" Then
                If Trim(txtPatient.Text) = "" Then '67213,先输入身份号再输入姓名时,不应该提醒,而是直接由身份证计算性别、年龄
                    Call zlControl.CboLocate(cbo性别, strSex)
                    txt年龄.Text = ReCalcOld(CDate(strbirthday), cbo年龄单位)
                    txt出生日期.Text = Format(strbirthday, "yyyy-mm-dd")
                    Call txt出生日期_Validate(False)
                Else
                    If MsgBox("输入的" & strInfo & "与身份证号的" & strInfo & "不一致，" & _
                            "将根据身份证号修改" & strInfo & "，是否继续？", vbInformation + vbYesNo, gstrSysName) = vbYes Then
                        Call zlControl.CboLocate(cbo性别, strSex)
                        txt年龄.Text = ReCalcOld(CDate(strbirthday), cbo年龄单位)
                        txt出生日期.Text = Format(strbirthday, "yyyy-mm-dd")
                        Call txt出生日期_Validate(False)
                    Else
                        If txtIDCard.Enabled And txtIDCard.Visible Then txtIDCard.SetFocus
                        Cancel = True: Exit Sub
                    End If
                End If
            End If
        Else
            MsgBox strErrInfo, vbInformation, gstrSysName
            If txtIDCard.Enabled And txtIDCard.Visible Then txtIDCard.SetFocus
            Cancel = True: Exit Sub
        End If
    End If
    
    '新输入的,肯定又要去查找一次,看病人信息中是否存在该身份证号的病人:
    Call GetPatient(IDKind.GetCurCard, txtIDCard.Text, False, True, Cancel)
    Call ReLoadCardFee(True, True)
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume

End Sub

Private Sub txtPatientPrint_GotFocus()
    Call zlControl.TxtSelAll(txtPatientPrint)
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txtPatientPrint_KeyPress(KeyAscii As Integer)
    If txt号别.Text = "" Then KeyAscii = 0: Exit Sub
    If txtPatientPrint.Text <> "" And KeyAscii = vbKeyReturn Then
        If cbo性别.Enabled And cbo性别.Visible Then
            cbo性别.SetFocus
        Else
            If cmdOK.Enabled And cmdOK.Visible Then cmdOK.SetFocus
        End If
    End If
End Sub

Private Sub txtPatientPrint_Validate(Cancel As Boolean)
    txtPatientPrint.Text = Trim(txtPatientPrint.Text)
End Sub

Private Sub txtSN_GotFocus()
    If (Not mTy_Para.bln随机序号选择) And mbytMode <> 1 Then
        zlCommFun.PressKey vbKeyTab
        Exit Sub
    End If
    Call zlControl.TxtSelAll(txtSN)
End Sub
Private Sub txtSN_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If InStr("1234567890" & Chr(8), Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        ElseIf txt号别.Text = "" Or mrsSNState Is Nothing Then
            KeyAscii = 0
        End If
    Else
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub txtSN_Validate(Cancel As Boolean)
'检查输入的序号的有效性
    Dim i As Long, j As Long, blnHave As Boolean
    Dim lngSN As Long
    Dim bln失效 As Boolean
    Dim bln
    Dim blnLock As Boolean
    Dim blnLocateSn As Boolean
    Dim lngLocateSnX As Long
    Dim lngLocateSnY As Long
    Dim lngRow As Long, lngCol As Long
    If mblnNotChange Then Exit Sub
    If Val(txtSN.Text) = 0 Then txtSN.Text = ""
    If Trim(txtSN.Text) = "" Then Exit Sub
    If txtSN.Tag = txtSN.Text Then Exit Sub '接收预约时没有变则不用检查
    If Not IsNumeric(txtSN.Text) Then
        Cancel = True
        Call zlControl.TxtSelAll(txtSN)
        Exit Sub
    End If
    
    If Not vsfList.Visible Then Exit Sub
    If mViewMode = V_普通号 Or mViewMode = v_专家号 Then
        '**********************************************
        '不分时段 判断 按照以前的方法
        '**********************************************
        
        lngSN = Val(txtSN.Text)
        For i = 0 To vsfList.Rows - 1
            For j = 0 To vsfList.Cols - 1
                If lngSN = Val(vsfList.TextMatrix(i, j)) Then
                    lngRow = i
                    lngCol = j
                    blnHave = True
                    Exit For
                End If
            Next
            If blnHave Then Exit For
        Next
        
        If Not blnHave Then
            If Not CheckAddAvailable Then
                MsgBox "该号别还有未使用序号，你不能使用加号序号！", vbInformation, gstrSysName
                txtSN.Text = ""
                Exit Sub
            End If
            If InStr(mstrPrivs, ";加号;") <= 0 Then
                MsgBox lngSN & "号超过最大限号数!你没有满号后继续挂号的权限.", vbInformation, gstrSysName
                Cancel = True
                txtSN.Text = ""
            Else
                If MsgBox(lngSN & "号超过最大限号数!你确定要使用吗?", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Cancel = True
                    Call zlControl.TxtSelAll(txtSN)
                Else
                    If mbytMode = 0 Then
                        With vsfList
                            For i = 0 To .Rows - 1
                                For j = 0 To .Cols - 1
                                    If .Cell(flexcpData, i, j) Like "加*" Then .Select i, j
                                Next j
                            Next i
                        End With
                    End If
                End If
            End If
        ElseIf Not mrsSNState Is Nothing Then
            mrsSNState.Filter = "序号=" & lngSN
            If mrsSNState.RecordCount > 0 Then
                If mrsSNState!状态 = 1 Or mrsSNState!状态 = 2 Then
                    Cancel = True
                    MsgBox lngSN & "号已被" & IIf(mrsSNState!状态 = 1, "使用", "预约") & "!请重新输入一个号.", vbInformation, gstrSysName
                    Call zlControl.TxtSelAll(txtSN)
                ElseIf mrsSNState!状态 = 3 Then
                    If mrsSNState!操作员姓名 = UserInfo.姓名 Then
                        If MsgBox(lngSN & "号是预留号!你确定要使用吗?", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            Cancel = True
                            txtSN.Text = ""
                            Call zlControl.TxtSelAll(txtSN)
                        Else
                            Call LocateSN(lngSN)
                        End If
                    Else
                        Cancel = True
                        MsgBox lngSN & "号已被" & mrsSNState!操作员姓名 & "预留!请重新输入一个号.", vbInformation, gstrSysName
                        txtSN.Text = ""
                        Call zlControl.TxtSelAll(txtSN)
                    End If
                ElseIf mrsSNState!状态 = 4 Then
                    If mTy_Para.blnReuseCancelNO = False Then
                        Cancel = True
                        MsgBox lngSN & "号已被退号,无法再次使用" & "!请重新输入一个号.", vbInformation, gstrSysName
                        Call zlControl.TxtSelAll(txtSN)
                    End If
                ElseIf mrsSNState!状态 = 5 Then
                    Cancel = True
                    MsgBox lngSN & "号已被自助机锁定,无法使用" & "!请重新输入一个号.", vbInformation, gstrSysName
                    Call zlControl.TxtSelAll(txtSN)
                ElseIf mrsSNState!状态 = 6 Then
                    Cancel = True
                    MsgBox lngSN & "号已被停诊,无法使用" & "!请重新输入一个号.", vbInformation, gstrSysName
                    Call zlControl.TxtSelAll(txtSN)
                End If
            Else
                If blnHave And vsfList.Cell(flexcpForeColor, lngRow, lngCol) = &HC000C0 Then
                    Cancel = True
                    MsgBox lngSN & "号不可用!请重新输入一个号.", vbInformation, gstrSysName
                    Call zlControl.TxtSelAll(txtSN)
                Else
                    Call LocateSN(lngSN)
                End If
            End If
        End If
    Else
        '*****************************************************
        '分时段 处理方法
        '只对专家号进行验证
        '普通号分时段 不对序号进行验证
        '*****************************************************
        If mViewMode <> v_专家号分时段 Then Exit Sub
        lngSN = Val(txtSN.Text)
        For i = 0 To vsfList.Rows - 1
            For j = 1 To vsfList.Cols - 1
                If lngSN = Val(Get时段(i, j, False)) Then
                    lngLocateSnX = i
                    lngLocateSnY = j
                    blnHave = True
                    blnLock = vsfList.Cell(flexcpForeColor, i, j) = vbRed And vsfList.Cell(flexcpFontStrikethru, i, j) = False
                    bln失效 = vsfList.Cell(flexcpForeColor, i, j) = vbGrayText
                    Exit For
                End If
            Next
            If blnHave Then Exit For
        Next
        If blnLock Then
            MsgBox lngSN & "号已经被锁定!请输入其他号进行挂号.", vbInformation, gstrSysName
            Cancel = True
            txtSN.Text = ""
        End If
        If bln失效 Then
            MsgBox lngSN & "号已经失效!请输入有效号进行挂号.", vbInformation, gstrSysName
            Cancel = True
            txtSN.Text = ""
        End If
        If Not blnHave Then
            If Not CheckAddAvailable Then
                MsgBox "该号别还有未使用序号，你不能使用加号序号！", vbInformation, gstrSysName
                txtSN.Text = ""
                Call locateSnBy时段(-1)
                Exit Sub
            End If
            If InStr(mstrPrivs, ";加号;") <= 0 Then
                MsgBox lngSN & "号超过最大限号数!你没有满号后继续挂号的权限.", vbInformation, gstrSysName
                Cancel = True
                txtSN.Text = ""
            Else
                If MsgBox(lngSN & "号超过最大限号数!你确定要使用吗?", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Cancel = True
                    Call zlControl.TxtSelAll(txtSN)
                Else
                    If mbytMode = 0 Then
                        With vsfList
                            For i = 0 To .Rows - 1
                                For j = 0 To .Cols - 1
                                    If .Cell(flexcpData, i, j) Like "加*" Then .Select i, j
                                Next j
                            Next i
                        End With
                    End If
                End If
            End If
        ElseIf Not mrsSNState Is Nothing Then
            mrsSNState.Filter = "序号=" & lngSN
            If mrsSNState.RecordCount > 0 Then
                If mrsSNState!状态 = 1 Or mrsSNState!状态 = 2 Then
                    Cancel = True
                    MsgBox lngSN & "号已被" & IIf(mrsSNState!状态 = 1, "使用", "预约") & "!请重新输入一个号.", vbInformation, gstrSysName
                    Call zlControl.TxtSelAll(txtSN)
                ElseIf mrsSNState!状态 = 3 Then
                    If mrsSNState!操作员姓名 = UserInfo.姓名 Then
                        If MsgBox(lngSN & "号是预留号!你确定要使用吗?", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            Cancel = True
                            txtSN.Text = ""
                            Call zlControl.TxtSelAll(txtSN)
                        Else
                            Call locateSnBy时段(lngSN)
                        End If
                    Else
                        Cancel = True
                        MsgBox lngSN & "号已被" & mrsSNState!操作员姓名 & "预留!请重新输入一个号.", vbInformation, gstrSysName
                        txtSN.Text = ""
                        Call zlControl.TxtSelAll(txtSN)
                    End If
                ElseIf mrsSNState!状态 = 4 Then
                    If mTy_Para.blnReuseCancelNO = False Then
                        Cancel = True
                        MsgBox lngSN & "号已被退号,无法再次使用" & "!请重新输入一个号.", vbInformation, gstrSysName
                        Call zlControl.TxtSelAll(txtSN)
                    End If
                ElseIf mrsSNState!状态 = 5 Then
                    Cancel = True
                    MsgBox lngSN & "号已被锁定," & "!请重新输入一个号.", vbInformation, gstrSysName
                    Call zlControl.TxtSelAll(txtSN)
                ElseIf mrsSNState!状态 = 6 Then
                    Cancel = True
                    MsgBox lngSN & "号已被停诊," & "!请重新输入一个号.", vbInformation, gstrSysName
                    Call zlControl.TxtSelAll(txtSN)
                End If
                If Cancel = False Then Call locateSnBy时段(lngSN)
            Else
                If blnHave And vsfList.Cell(flexcpForeColor, lngLocateSnX, lngLocateSnY) = &HC000C0 Then
                    Cancel = True
                    MsgBox lngSN & "号不可用!请重新输入一个号.", vbInformation, gstrSysName
                    Call zlControl.TxtSelAll(txtSN)
                Else
                    Call locateSnBy时段(lngSN)
                End If
            End If
        End If
    End If
End Sub

Private Sub txt出生日期_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txt出生日期.Text = "____-__-__" Then
           zlCommFun.PressKey (vbKeyTab) '跳过时间
           zlCommFun.PressKey (vbKeyTab)
       Else
           zlCommFun.PressKey (vbKeyTab)
       End If
    End If
End Sub

Private Sub txt出生日期_Validate(Cancel As Boolean)
    If txt出生日期.Tag <> txt出生日期.Text Then
        With mobjfrmPatiInfo '更正出生日期
            .txt出生日期.Text = txt出生日期.Text
            txt出生日期.Tag = txt出生日期.Text
            .txt年龄.Text = txt年龄.Text
            .txt年龄.Tag = txt年龄.Text
            txt年龄.Tag = txt年龄.Text
            .cbo年龄单位.Visible = cbo年龄单位.Visible
            If .cbo年龄单位.ListCount <> 0 Then .cbo年龄单位.ListIndex = cbo年龄单位.ListIndex
        End With
        Call ShowRegistFromInput
    End If
End Sub

Private Sub txt出生时间_Change()
    Dim str出生时间 As String
    '76669，李南春,2014-8-18,病人年龄更新
    If IsDate(txt出生日期.Text) And mblnChange Then
        str出生时间 = txt出生日期.Text & IIf(IsDate(txt出生时间.Text), " " & txt出生时间.Text, "")
        txt年龄.Text = ReCalcOld(CDate(str出生时间), cbo年龄单位)
        txt年龄.Tag = txt年龄.Text
    End If
End Sub

Private Sub txt出生时间_GotFocus()
    zlControl.TxtSelAll txt出生时间
End Sub

Private Sub txt出生时间_KeyPress(KeyAscii As Integer)
    If Not IsDate(txt出生日期.Text) Then
        KeyAscii = 0
        txt出生时间.Text = "__:__"
    End If
    If KeyAscii = 13 Then
        KeyAscii = 0
        zlCommFun.PressKey (vbKeyTab)
    End If
End Sub


Private Sub txt出生时间_Validate(Cancel As Boolean)
    If txt出生时间.Text <> "__:__" And Not IsDate(txt出生时间.Text) Then
        txt出生时间.SetFocus
        Cancel = True
    End If
End Sub

Private Sub txt出生日期_Change()
    Dim str出生时间 As String
    
    If IsDate(txt出生日期.Text) And mblnChange Then
        mblnChange = False
        txt出生日期.Text = Format(CDate(txt出生日期.Text), "yyyy-mm-dd") '0002-02-02自动转换为2002-02-02,否则,看到的是2002,实际值却是0002
        mblnChange = True
        
        str出生时间 = txt出生日期.Text & IIf(IsDate(txt出生时间.Text), " " & txt出生时间.Text, "")
        txt年龄.Text = ReCalcOld(CDate(str出生时间), cbo年龄单位)
        txt年龄.Tag = txt年龄.Text
        mblnGetBirth = False
    End If
End Sub
Private Sub txt出生日期_GotFocus()
    zlControl.TxtSelAll txt出生日期
End Sub

Private Sub txt出生日期_LostFocus()
    If txt出生日期.Text <> "____-__-__" And Not IsDate(txt出生日期.Text) Then
      If txt出生日期.Enabled And txt出生日期.Visible Then txt出生日期.SetFocus
    End If
End Sub

Private Sub txtPatient_Change()
    If txtPatient.Locked Then Exit Sub
    If mblnUnChange Or mbytInState = 1 Then Exit Sub
    
    '74430,冉俊明,2014-7-8,挂号界面显示病人照片的浮动窗体
    picPatiPicBack.Visible = False: cmdPatiPic.Enabled = txtPatient.Text <> ""
    
    mblnBoundPati = False
    mblnUnChange = True
    txt门诊号.Enabled = txtPatient.Text <> "" And InStr(mstrPrivs, ";建立病案;") > 0
    cmdMore.Enabled = txtPatient.Text <> "" And InStr(mstrPrivs, ";建立病案;") > 0
    cmdMore.Tag = ""    '用于判断是否进入病人信息编辑后提取过已有病人
    cmdCard.Enabled = Not mblnNewCard   'txtPatient.Text <> "" And
    cmdCard.Enabled = cmdCard.Enabled And Not (mblnStation And mTy_Para.bln挂号必须刷卡)
    
    If Me.ActiveControl Is txtPatient Then
        If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(txtPatient.Text = "")
        If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(txtPatient.Text = "")
        IDKind.SetAutoReadCard (txtPatient.Text = "")
    End If
    
    If Trim(txtPatient.Text) = "" Then
        '清除病人时则清除所有病人信息
        If mstr门诊号 = "" Then '才自动刷入了门诊号时不清除
            Call ClearPatientInfo
            Call Init费别(True, False) '恢复缺省费别
            Set mrsInfo = Nothing
            Call ClearmobjfrmPatiInfoFace(Not (mblnNewCard And gblnNewCardNoPop))
        End If
    End If
    mblnUnChange = False
    '还原文本框颜色
    txtPatient.ForeColor = Me.ForeColor
End Sub

Private Sub txtPatient_GotFocus()

    Call zlControl.TxtSelAll(txtPatient)
    
    'LED语音报价
    If gblnLED And mbytMode <> 1 And mbytInState = 0 And txt号别.Text <> "" And txtPatient.Text = "" Then
        zl9LedVoice.Speak "#4" '请问你的姓名
    End If
        
    If txtPatient.Text = "" And Not txtPatient.Locked Then
        If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(True)
        If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(True)
        IDKind.SetAutoReadCard True
    End If
    Call zlCommFun.OpenIme(True)
End Sub
Private Sub zlInusreIdentify()
    '------------------------------------------------------------------------------------------------------------------------
    '功能：医保身份验卡
    '编制：刘兴洪
    '日期：2010-07-14 11:32:08
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim lng病人ID As Long
    Dim str病人类型 As String
    Dim rsTmp As ADODB.Recordset
    Dim cur余额 As Currency
    Dim curMoney As Currency
    Dim i As Integer
    Dim curPayed As Currency
    Dim curTotal As Currency
    If mrsInfo Is Nothing Then
        lng病人ID = 0
        str病人类型 = ""
    Else
        lng病人ID = Val(Nvl(mrsInfo!病人ID))
        str病人类型 = Nvl(mrsInfo!病人类型)
    End If
    '52867
    Call SetShowBalance
    If gblnLED Then zl9LedVoice.Speak "#50"

    If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(False)
    If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(False)
    IDKind.SetAutoReadCard False
    '68991
    Dim strAdvance As String    '结算模式(0-先结算后诊疗或1-先诊疗后结算)|挂号费收取方式(0-现收或1-记帐)
    Dim varData As Variant
    mstrYBPati = gclsInsure.Identify(3, lng病人ID, mintInsure, strAdvance)
    mRegistFeeMode = EM_RG_现收: mPatiChargeMode = EM_先结算后诊疗
    If txtPatient.Text = "" And Not txtPatient.Locked Then
        If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(True)
        If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(True)
        IDKind.SetAutoReadCard True
    End If
    
    If mstrYBPati = "" Then
        '修改问题：38917 作者：冉勇
        If Not txtPatient.Enabled Then txtPatient.Enabled = True
         mstrYBPati = "": mintInsure = 0: txtPatient.SetFocus
         Exit Sub
    End If
    
    '空或：0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码);8病人ID
    If UBound(Split(mstrYBPati, ";")) >= 8 Then
        If IsNumeric(Split(mstrYBPati, ";")(8)) Then lng病人ID = Val(Split(mstrYBPati, ";")(8))
    End If
        
    If lng病人ID = 0 Then
        mstrYBPati = "": mintInsure = 0: txtPatient.SetFocus
        Exit Sub
    End If
        
    '问题:29283
    '  -- 参数:调用场合-1-挂号;2-收费
    '  --        病人id_In-病人ID(未建档的,传入零)
    '  --        卡号_In: 刷卡卡号;未刷卡时,为空
    '  --         刷卡方式_In:  1-普能刷卡;2-医保刷卡
    If mbytMode <> 1 Then
        If zlPatiCardCheck(1, lng病人ID, CStr(Split(mstrYBPati, ";")(0)), 2) = False Then
            Call ClearmobjfrmPatiInfoFace: ClearPatientInfo
            mstrYBPati = "": mintInsure = 0: txtPatient.SetFocus
            Exit Sub
        End If
    End If
    Call initInsurePara(lng病人ID)
    txtPatient.Text = "-" & lng病人ID
    Call SetIdentifyLocked(False)
    Call txtPatient_Validate(False)    '其中的Setfocus调用使本事件(txtPatient_KeyPress)执行完后,不会再次自动执行txtPatient_Validate
    '74428：李南春，2014-7-8，病人姓名显示颜色处理
    If mblnUnload Then
        mblnUnload = False
        Exit Sub
    End If
    Call SetPatiColor(txtPatient, str病人类型, vbRed)
    mobjfrmPatiInfo.txtPatient.ForeColor = txtPatient.ForeColor
    Call SetIdentifyLocked(True)
    '68991
    If strAdvance <> "" Then
        varData = Split(strAdvance & "|", "|")
        mPatiChargeMode = IIf(Val(varData(0)) = 1, EM_先诊疗后结算, EM_先结算后诊疗)
        mRegistFeeMode = IIf(Val(varData(1)) = 1, EM_RG_记帐, EM_RG_现收)
    End If
    
    Dim dbl家属余额 As Double
    Set rsTmp = GetMoneyInfo(lng病人ID, , , 1, , , True)
    cur余额 = 0: stbThis.Panels(4).ToolTipText = ""
    Do While Not rsTmp.EOF
        cur余额 = cur余额 + Val(Nvl(rsTmp!预交余额))
        cur余额 = cur余额 - Val(Nvl(rsTmp!费用余额))
        If Val(Nvl(rsTmp!家属)) = 1 Then
            dbl家属余额 = Val(Nvl(rsTmp!预交余额)) - Val(Nvl(rsTmp!费用余额))
        End If
        rsTmp.MoveNext
    Loop
    
    mcur个帐余额 = gclsInsure.SelfBalance(lng病人ID, CStr(Split(mstrYBPati, ";")(1)), 10, mcur个帐透支, mintInsure)
    mdbl个帐余额 = mcur个帐余额
    stbThis.Panels(3).Text = "个人帐户余额:" & Format(mcur个帐余额, "0.00")
    Call CalcYBMoney
    Call initInsurePara(lng病人ID)
    
    Call ShowMedicareInfo(Not mRegistFeeMode = EM_RG_记帐)
    Call ShowDeposit(False)
    If cur余额 > 0 Then
        Call ShowDeposit(Not mRegistFeeMode = EM_RG_记帐)
        mdbl预交余额 = cur余额
        curTotal = GetRegistMoney(True)
        curPayed = 0
        For i = 1 To vsfPay.Rows - 1
            curPayed = curPayed + Val(vsfPay.TextMatrix(i, 1))
        Next i
        For i = 1 To vsfPay.Rows - 1
            If vsfPay.RowData(i) = 0 And vsfPay.TextMatrix(i, 0) <> "" Then
                vsfPay.TextMatrix(i, 6) = mdbl预交余额
                If gblnPrePayPriority Then
                    If mdbl预交余额 > curTotal - curPayed Then
                        vsfPay.TextMatrix(i, 1) = Format(curTotal - curPayed, "0.00")
                    Else
                        vsfPay.TextMatrix(i, 1) = Format(mdbl预交余额, "0.00")
                    End If
                    Call Set连续挂号
                End If
            End If
        Next i
        stbThis.Panels(4).Text = "门诊预交余额:" & Format(cur余额, "0.00")
        If Round(dbl家属余额, 6) <> 0 Then stbThis.Panels(4).ToolTipText = "含家属预交：" & Format(dbl家属余额, "0.00")
        
        '医生站挂号缺省使用预交款
        curMoney = GetRegistMoney
    End If
    
    If MCPAR.使用个人帐户 Then
        If mstr个人帐户 = "" Then MsgBox "挂号场合未设置个人帐户结算，病人帐户不能支付！", vbInformation, gstrSysName
    End If
    
    '68991
    If mRegistFeeMode = EM_RG_记帐 Or CheckIsPrice Then
        Call SetUndisplayBalance
    End If
    
End Sub

 
Private Sub txtPatient_KeyDown(KeyCode As Integer, Shift As Integer)
    '问题号:44114
    If KeyCode = 38 And 1 < IDKind.IDKind And IDKind.IDKind <= IDKind.ListCount Then '小键盘上方向键
        IDKind.IDKind = IDKind.IDKind - 1
    ElseIf KeyCode = 40 And IDKind.IDKind < IDKind.ListCount Then '小键盘下方向键
        IDKind.IDKind = IDKind.IDKind + 1
    End If
End Sub
Private Sub txtPatient_KeyPress(KeyAscii As Integer)
    Dim lng病人ID As Long, blnCard As Boolean
    
    '问题:51488
    '空格读卡
'    If IDKind.GetCurCard.是否刷卡 = False And KeyAscii = Asc(" ") And mbytInState = 0 Then
'        KeyAscii = 0: Call IDKind_Click(IDKind.GetCurCard): Exit Sub
'    End If
    
    If (KeyAscii = Asc("/") Or KeyAscii = Asc("／") Or KeyAscii = Asc("、") Or KeyAscii = Asc("、")) And Trim(txtPatient.Text) = "" Then
        '预约接收时,如果单据号输入的是"/"或"、"(全角和半角),则自动弹出小窗口,供预约挂号用"
        KeyAscii = 0:        Call ShowBookSeled
        Call CreateMobjIDCard
        Exit Sub
    End If
    If SetBrushCard(txtPatient, KeyAscii) = True Then Exit Sub
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If mbytMode <> 1 And Not gblnPrice And Trim(txtPatient.Text) = "" And mobjfrmPatiInfo.mstrCard = "" Then
            '医保身份验卡
            Call zlInusreIdentify
        Else
            zlCommFun.PressKey (vbKeyTab)
        End If
    ElseIf InStr(1, "'[]+", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0 '住院病人不允许再挂号，特殊字符过滤在Form_KeyPress中进行
    Else
        If txtPatient.Text = "" Then gsngStartTime = Timer
        gblnLen = False
        If IDKind.GetCurCard Is Nothing Then Exit Sub
        If IDKind.GetCurCard.名称 = "门诊号" Then
            If KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then
                If Not (IsNumeric(Chr(KeyAscii)) Or Chr(KeyAscii) = "-") Then KeyAscii = 0: Exit Sub
            End If
        ElseIf IDKind.GetCurCard.名称 = "姓名" Or IDKind.GetCurCard.名称 = "姓名或就诊卡" Then
            blnCard = zlCommFun.InputIsCard(txtPatient, KeyAscii, gCurSendCard.str卡号密文 <> "")
            mblnCard = blnCard
            If blnCard And Len(txtPatient.Text) = gCurSendCard.lng卡号长度 - 1 And KeyAscii <> 8 Then
                txtPatient.Text = txtPatient.Text & Chr(KeyAscii): txtPatient.SelStart = Len(txtPatient.Text)
                KeyAscii = 0
                gblnLen = True
                gsngStartTime = Timer
                Call txtPatient_Validate(False)
                mblnCard = False
                '刘兴洪:27494  20100117
                If Replace(txtPatient.Text, vbCrLf, "") = "" Then
                    DoEvents: txtPatient.SetFocus
                End If
            End If
        ElseIf IDKind.GetCurCard.接口序号 <> 0 Then
            '42947
            blnCard = zlCommFun.InputIsCard(txtPatient, KeyAscii, IDKind.ShowPassText)
            mblnCard = blnCard
            If blnCard And Len(txtPatient.Text) = IDKind.GetCardNoLen - 1 And KeyAscii <> 8 Then
                txtPatient.Text = txtPatient.Text & Chr(KeyAscii): txtPatient.SelStart = Len(txtPatient.Text)
                KeyAscii = 0
                gblnLen = True
                gsngStartTime = Timer
                Call txtPatient_Validate(False)
                mblnCard = False
                '刘兴洪:27494  20100117
                If Replace(txtPatient.Text, vbCrLf, "") = "" Then
                    DoEvents: txtPatient.SetFocus
                End If
            End If
        
        Else
            txtPatient.PasswordChar = IIf(IDKind.ShowPassText, "*", "")
            '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
            txtPatient.IMEMode = 0
        End If
    End If
End Sub

Private Sub txtPatient_LostFocus()
    If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(False)
    If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(False)
    IDKind.SetAutoReadCard False
    Call zlCommFun.OpenIme
End Sub

Private Sub txtPatient_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        glngTXTProc = GetWindowLong(txtPatient.Hwnd, GWL_WNDPROC)
        Call SetWindowLong(txtPatient.Hwnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txtPatient_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Call SetWindowLong(txtPatient.Hwnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub


Public Sub txtPatient_Validate(Cancel As Boolean)
    Dim blnTmp As Boolean
    Dim strTemp As String, lng卡类别ID As Long
    If txtPatient.Locked And mblnOnVilidate = False Then Exit Sub
    If mstrPrePati = txtPatient.Text Then
        '自动门诊号处理
        If txt门诊号.Text = "" Then
            If txt门诊号.Enabled And txt门诊号.Visible Then
                txt门诊号.TabStop = True
                If gbln自动门诊号 Or mblnStation Then
                    If txt号别.Text <> "" And mbln建病案 And txt门诊号.Text = "" And txtPatient.Text <> "" Then
                        txt门诊号.Text = zlGet门诊号
                        mintNOLength = Len(txt门诊号.Text)  '用来判断修改门诊号时的异常输入
                        txt门诊号.TabStop = False
                    End If
                End If
            End If
        End If
        If mblnOnVilidate = False Then Exit Sub
    End If
        
    '上次挂号的费用情况,新号时清除
    txt缴款.Text = "0.00": txt找补.Text = "0.00"
    txt合计.Text = Format(mcur合计 + GetRegistMoney, "0.00"): mint挂号数 = 0
    
    Call Set连续挂号
    If mbytMode = 0 And txt缴款.Enabled = False Then txt缴款.Enabled = True
    
    '更换病人或不输入病人后,清除挂号累计,预约时不输缴款,一直保持累计
    If Not (mTy_Para.byt缴款方式 = 1 And mbytMode <> 1) Then mcur合计 = 0: mcur应缴 = 0
    
    If txtPatient.Text <> "" Then
        txtPatient.Text = Trim(txtPatient.Text)
        strTemp = txtPatient.Text
        If (Left(txtPatient.Text, 1) = "*" Or Left(txtPatient.Text, 1) = "-") And IsNumeric(Mid(txtPatient.Text, 2)) Then blnTmp = True
        
        Call GetPatient(IDKind.GetCurCard, txtPatient.Text, mblnCard)
        
        '69730,刘尔旋,2014-01-23,对医生工作站启用了挂号必须刷卡参数的检查
        If mblnStation And mbytMode = 0 And mTy_Para.bln挂号必须刷卡 Then
            If mrsInfo Is Nothing Then
                MsgBox "没有找到该卡对应的病人信息，请检查该卡是否正确！", vbInformation, gstrSysName
                txtPatient.Text = ""
                txtPatient.SetFocus
                Cancel = True
                Exit Sub
            End If
        End If
        
        '问题号:58843
        If mblnStation Then
            If Not mrsInfo Is Nothing Then mstrPrePati = txtPatient.Text
            SetPatiInfoEnabled vsfPlan.TextMatrix(vsfPlan.Row, GetCol("病案")) <> "", mrsInfo Is Nothing
        End If
        
        
        '病人预约单据接收提醒
        If Not mblnStation And Not mrsInfo Is Nothing And mbytMode = 0 Then
            If mbytInState = 0 And mstrNoIn <> "" Then Exit Sub
            If zlExistsTodaysAppointment(mrsInfo!病人ID) Then Exit Sub
        End If
        
        
        If Not IDKind.GetCurCard.名称 Like "姓名*" Then
            lng卡类别ID = IDKind.GetCurCard.接口序号
            If lng卡类别ID <> IDKind.GetDefaultCardTypeID And lng卡类别ID > 0 Then
                mblnCard = False
            End If
            '刘兴洪:65945,不能以缺省卡作为发卡依据,如果是门诊号就有问题.
          ' If lng卡类别ID <= 0 Then lng卡类别ID = IDKind.GetDefaultCardTypeID

        End If
 
        If mblnCard Or (IsCardType(IDKind, "IC卡号") _
            Or (gCurSendCard.lng卡类别ID = lng卡类别ID And lng卡类别ID > 0)) And Not blnTmp And lblPrompt.Caption = "" Then
            mblnCard = False
            mbln发卡 = True '问题号:56599
            If mrsInfo Is Nothing Then
                If mblnStation Or mbytMode = 1 Then '医生站或预约时不支持发卡,因为发卡要收费
                    Cancel = True: txtPatient.Text = "": Exit Sub
                Else
                    If mTy_Para.bln允许住院病人挂号 = False Then
                        If PatiExist(UCase(txtPatient.Text)) Then
                            MsgBox "发现该持卡病人在院,或该病人信息目前不可用!不能以此卡挂号!", vbInformation, gstrSysName
                            Cancel = True: txtPatient.Text = "":  Exit Sub
                        End If
                    End If
                    If IsCardType(IDKind, "IC卡") Then mblnICCard = True
                    
                    '如果卡费和挂号费一起收取则此时没有档案,保存挂号单时再建档.否则卡费存为划价单,此时已建档
                    If LoadCard(False) Then
                        mblnNewCard = True
                        '问题:29283
                        '  -- 参数:调用场合-1-挂号;2-收费
                        '  --        病人id_In-病人ID(未建档的,传入零)
                        '  --        卡号_In: 刷卡卡号;未刷卡时,为空
                        '  --         刷卡方式_In:  1-普能刷卡;2-医保刷卡
                        '问题:And mbytMode <> 1 :40482
                        If mstrYBPati = "" And mbytMode <> 1 Then
                            If zlPatiCardCheck(1, 0, Trim(mobjfrmPatiInfo.txt卡号.Text), 1) = False Then
                                Call ClearmobjfrmPatiInfoFace: ClearPatientInfo
                                Cancel = True: txtPatient.Text = "":  Exit Sub
                            End If
                        End If
                        
                        Call ShowRegistFromInput    '重新加载卡费信息
                        txtPatient.PasswordChar = ""
                    Else
                        txtPatient.PasswordChar = ""
                        Cancel = True: txtPatient.Text = "": Exit Sub
                    End If
                End If
            Else
                '问题:29283
                '  -- 参数:调用场合-1-挂号;2-收费
                '  --        病人id_In-病人ID(未建档的,传入零)
                '  --        卡号_In: 刷卡卡号;未刷卡时,为空
                '  --         刷卡方式_In:  1-普能刷卡;2-医保刷卡
                'And mbytMode <> 1:40482
                If mstrYBPati = "" And mbytMode <> 1 Then
                    If zlPatiCardCheck(1, Val(Nvl(mrsInfo!病人ID)), strTemp, 1) = False Then
                        Call ClearmobjfrmPatiInfoFace: ClearPatientInfo
                        Set mrsInfo = Nothing: txt门诊号.Enabled = True
                        Cancel = True: txtPatient.Text = "":  Exit Sub
                    End If
               End If
                 '就诊卡密码检查
                If Mid(gstrCardPass, 1, 1) = "1" And mstrPassWord <> "" Then
                    '54501
                    If Not zlCommFun.VerifyPassWord(Me, "" & mstrPassWord, mrsInfo!姓名, mrsInfo!性别, "" & mrsInfo!年龄) Then
                        txt门诊号.Enabled = True: Set mrsInfo = Nothing
                        Cancel = True: txtPatient.Text = "":  Exit Sub
                    End If
                End If
            End If
        Else
                '问题:29283
                '  -- 参数:调用场合-1-挂号;2-收费
                '  --        病人id_In-病人ID(未建档的,传入零)
                '  --        卡号_In: 刷卡卡号;未刷卡时,为空
                '  --         刷卡方式_In:  1-普能刷卡;2-医保刷卡
                'And mbytMode <> 1:40482
                If mstrYBPati = "" And mbytMode <> 1 Then
                    If mrsInfo Is Nothing Then
                        If Trim(mobjfrmPatiInfo.txt卡号.Text) <> "" Then    '读取有卡号的病人时没有加载卡号到界面
                            strTemp = Trim(mobjfrmPatiInfo.txt卡号.Text)
                        End If
                    
                        If zlPatiCardCheck(1, 0, strTemp, 1) = False Then
                            Call ClearmobjfrmPatiInfoFace: ClearPatientInfo
                            Set mrsInfo = Nothing: txt门诊号.Enabled = True
                            Cancel = True: txtPatient.Text = "":  Exit Sub
                        End If
                    Else
                        If zlPatiCardCheck(1, Val(Nvl(mrsInfo!病人ID)), "", 1) = False Then
                            Call ClearmobjfrmPatiInfoFace: ClearPatientInfo
                            Set mrsInfo = Nothing: txt门诊号.Enabled = True
                            Cancel = True: txtPatient.Text = "":  Exit Sub
                        End If
                    End If
               End If
               mblnCard = False
        End If
        
        If Not mrsInfo Is Nothing And gblnPrice And mbytMode = 0 And txt缴款.Enabled Then txt缴款.Enabled = False
        
        
        If mbytMode <> 2 Then
            If Not mrsInfo Is Nothing And InStr(1, mstrPrivs, ";调整病人费别;") = 0 And Not mblnStation Then
                cbo费别.Locked = True: cbo费别.TabStop = False
            Else
                cbo费别.Locked = False: cbo费别.TabStop = gbln费别
            End If
        End If
        '其中通过cbo费别_Click事件会调用ShowRegistFromInput
        Call Init费别((mbytMode = 2 Or mbytMode = 0 And mbytInState = 0 And mstrNoIn <> "") Or mrsInfo Is Nothing, Not mrsInfo Is Nothing Or mblnNewCard)

        If txtPatient.Text = "" And mstr门诊号 <> "" Then '使用输入的门诊号作为缺省号
            Cancel = True
            If IDKind.IDKind = IDKind.GetKindIndex("门诊号") Then
                IDKind.IDKind = IDKind.GetKindIndex("姓名")
                mblnReSetIDKind = True
            End If
            txt门诊号.Text = mstr门诊号
            Call txtPatient_GotFocus 'LED:请问姓名
            Exit Sub
        End If
        
        '新输入的病人
        If mrsInfo Is Nothing And (Not mblnNewCard Or gblnNewCardNoPop) And Not mblnBrushPlugin Then
            If mblnIDCardKind And mbytMode = 1 Then
                    '不清除年龄,因为在输入的时候,已经根据身份证号读取出来了:31182
            Else
                txt年龄.Text = ""
                Call zlControl.CboLocate(cbo年龄单位, "岁")
                If gstr性别 <> "无" Then
                    Call SetCboDefault(cbo性别)
                Else
                    cbo性别.ListIndex = -1
                End If
                txtIDCard.Text = "": txtIDCard.Tag = ""
                txt证件.Text = "": txt证件.Tag = ""
            End If
            cbo家庭地址.Text = ""
            cbo户口地址.Text = ""
            txt家庭电话.Text = ""
            '89242:李南春,2015/12/7,读取病人地址信息
            Call zlLoadDefaultAddr(padd家庭地址)
            Call zlLoadDefaultAddr(padd户口地址)
            '新病人保持输入的门诊号
            If Not (txt门诊号.Text <> "" And mstr门诊号 = txt门诊号.Text) Then txt门诊号.Text = ""
            Call SetCboDefault(cbo付款方式)
            If mbytMode <> 2 Then Call SetCboDefault(cbo费别)
            Call ClearmobjfrmPatiInfoFace(Not (mblnNewCard And gblnNewCardNoPop))
            Call zlQueryEMPIPatiInfo
        End If
        
        '门诊医生站挂号，或本地参数设置自动产生门诊号
        If txt门诊号.Enabled And txt门诊号.Visible Then
            txt门诊号.TabStop = True
            If gbln自动门诊号 Or mblnStation Then
                If txt号别.Text <> "" And mbln建病案 And txt门诊号.Text = "" And txtPatient.Text <> "" Then
                    txt门诊号.Text = zlGet门诊号
                    mintNOLength = Len(txt门诊号.Text)  '用来判断修改门诊号时的异常输入
                    txt门诊号.TabStop = False
                End If
            End If
        End If
        If mblnStartFactUseType Then
            Call ReInitPatiInvoice
        End If
        If mblnNewCard Then
             '29396
            If gblnNewCardNoPop And mblnCard And Not mblnBrushPlugin Then
                Cancel = True: txtPatient.SetFocus
            ElseIf txt门诊号.Text = "" And txt门诊号.Enabled And txt门诊号.Visible Then
                txt门诊号.SetFocus
            ElseIf cbo结算方式.Enabled And cbo结算方式.Visible Then
                cbo结算方式.SetFocus
            ElseIf chk病历费.Enabled And chk病历费.Visible Then
                chk病历费.SetFocus
            ElseIf txt缴款.Enabled And txt缴款.Visible And mTy_Para.byt缴款方式 = 1 Then
                txt缴款.SetFocus
            Else
                cmdOK.SetFocus
            End If
        ElseIf Not mrsInfo Is Nothing Then
            '89242:李南春,2015/12/7,读取病人地址信息
            If mblnStructAdress Then
                If padd家庭地址.CheckNullValue <> "" And padd家庭地址.Enabled And padd家庭地址.Visible And padd家庭地址.TabStop Then
                    padd家庭地址.SetFocus
                ElseIf padd户口地址.CheckNullValue <> "" And padd户口地址.Enabled And padd户口地址.Visible And padd户口地址.TabStop Then
                    padd户口地址.SetFocus
                End If
            Else
                If cbo家庭地址.Text = "" And cbo家庭地址.Enabled And cbo家庭地址.Visible And cbo家庭地址.TabStop Then
                     cbo家庭地址.SetFocus
                End If
            End If
            If txt门诊号.Enabled And txt门诊号.Visible And IsNull(mrsInfo!门诊号) And txt门诊号.TabStop Then
                 txt门诊号.SetFocus
            ElseIf cbo结算方式.Enabled And cbo结算方式.Visible Then
                 cbo结算方式.SetFocus
            ElseIf chk病历费.Enabled And chk病历费.Visible Then
                 chk病历费.SetFocus
            ElseIf txt缴款.Enabled And txt缴款.Visible And mTy_Para.byt缴款方式 = 1 Then
                txt缴款.SetFocus
            Else
                 If cmdOK.Enabled And cmdOK.Visible Then cmdOK.SetFocus
            End If
        Else
            If txtPatient.Text = "" And txtPatient.Enabled And txtPatient.Visible Then Cancel = True
        End If
        
    Else '为空表示不想输入病人信息
        Call ClearPatientInfo
        If mbytMode <> 2 Then Call SetCboDefault(cbo费别)
        Call ShowRegistFromInput
        
        Call ClearmobjfrmPatiInfoFace(Not (mblnNewCard And gblnNewCardNoPop))
        
        If cbo费别.Enabled And cbo费别.Visible Then
             cbo费别.SetFocus
        ElseIf cbo结算方式.Enabled And cbo结算方式.Visible Then
             cbo结算方式.SetFocus
        ElseIf chk病历费.Enabled And chk病历费.Visible Then
             chk病历费.SetFocus
        Else
             cmdOK.SetFocus
        End If
    End If
    Call ReLoadCardFee(True, True)
    Call Led欢迎信息
    
    If CheckIsPrice Or mRegistFeeMode = EM_RG_记帐 Then
        Call SetUndisplayBalance
    Else
        Call SetShowBalance
    End If
    
    mstrPrePati = txtPatient.Text
End Sub

Private Sub Led欢迎信息()
    Dim strInfo As String, lngPatient As Long
    'LED初始化
    If mbytMode = 0 And mbytInState = 0 And gblnLED Then
        zl9LedVoice.Speak "#1"
        
        strInfo = Trim(txtPatient.Text)
        If Not mrsInfo Is Nothing Then
            If mrsInfo.State = 1 Then strInfo = strInfo & " " & mrsInfo!性别 & " " & mrsInfo!年龄: lngPatient = Val("" & mrsInfo!病人ID)
        End If
        zl9LedVoice.DisplayPatient strInfo, lngPatient
    End If
End Sub

Private Sub txt号别_Validate(Cancel As Boolean)
    '清除上一张单据号
    If mbytInState = 0 And chkCancel.Value = 0 Then
        If cboNO.ListIndex <> -1 Then cboNO.ListIndex = -1
    End If
    mstrPre号别 = Trim(txt号别.Text) '53299
    If Trim(txt号别.Text) = "" Then Exit Sub
    mlngPreRow = vsfPlan.Row
    If CheckNoValied(vsfPlan.Row) = False Then
        mstrPre号别 = "" '53299
        mlngPreRow = 0
        Cancel = True
         txt号别.Text = "": txt号别.SetFocus: Exit Sub
    End If
End Sub

 
Private Sub txt家庭电话_GotFocus()
    zlCommFun.OpenIme False
End Sub

Private Sub txt家庭电话_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txt家庭电话_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txt家庭电话, KeyAscii, m文本式
End Sub

Private Sub txt家庭电话_Validate(Cancel As Boolean)
    If mobjfrmPatiInfo Is Nothing Then Exit Sub
    With mobjfrmPatiInfo
        .txt家庭电话.Text = txt家庭电话.Text
    End With
End Sub

Private Sub txt缴款_Change()
    Dim cur应缴 As Currency
    If Val(txt缴款.Text) = 0 Then
        txt找补.Text = "0.00"
    Else
        cur应缴 = mcur应缴 + GetRegistMoney
        txt找补.Text = Format(Val(txt缴款.Text) - Val(txt本次应缴.Text), "0.00")
    End If
End Sub

Private Sub txt缴款_GotFocus()
    Dim cur应缴 As Currency
    
    '只以缴款作为收费结束条件时,必须输入缴款或0
    If mTy_Para.byt缴款方式 = 1 Then
        If Val(txt缴款.Text) = 0 And Me.ActiveControl Is txt缴款 Then
            txt缴款.Text = ""
        End If
    End If
    Call zlControl.TxtSelAll(txt缴款)
    
    'LED语音报价
     If Not (mintInsure <> 0 And mstrYBPati <> "") Then
        cur应缴 = mcur应缴 + GetRegistMoney
        If gblnLED And mbytMode <> 1 And mbytInState = 0 Then
            zl9LedVoice.Speak "#21 " & Format(cur应缴, "0.00")
        End If
    End If
End Sub

Private Sub txt缴款_KeyPress(KeyAscii As Integer)
    Dim cur应缴 As Currency
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txt缴款.Text = "" Then
            If GetRegistMoney = 0 Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
        End If
        If mTy_Para.byt缴款方式 = 1 And txt缴款.Text = "" Then Exit Sub
'        If Val(txt缴款.Text) <> 0 Then
'            If Val(txt找补.Text) < 0 Then
'                MsgBox "缴款金额不足。", vbInformation, gstrSysName
'                Call zlControl.TxtSelAll(txt缴款): Exit Sub
'            End If
'        End If
        Call zlCommFun.PressKey(vbKeyTab)
        
        'LED显示
         If Not (mintInsure <> 0 And mstrYBPati <> "") Then
            If gblnLED And mbytMode <> 1 And mbytInState = 0 And Val(txt找补.Text) >= 0 Then
                cur应缴 = mcur应缴 + GetRegistMoney
                zl9LedVoice.DispCharge Format(cur应缴, "0.00"), txt缴款.Text, txt找补.Text
                zl9LedVoice.Speak "#22 " & txt缴款.Text
                zl9LedVoice.Speak "#23 " & txt找补.Text
                zl9LedVoice.Speak "#3"
                txt缴款.Tag = "1"
            End If
        End If
    Else
        If KeyAscii = Asc(".") And InStr(txt缴款.Text, ".") > 0 Then KeyAscii = 0:  Exit Sub
        If InStr("0123456789." & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txt缴款_LostFocus()
    txt缴款.Text = Format(Val(txt缴款.Text), "0.00")
End Sub

Private Sub txt缴款_Validate(Cancel As Boolean)
    txt缴款.Text = Format(Val(txt缴款.Text), "0.00")
End Sub

Private Sub txt科室_Change()
    If Not mrsInfo Is Nothing Then
        If mlng挂号科室ID > 0 And txt科室.Text <> "" Then
            mobjfrmPatiInfo.chk复诊.Value = IIf(Check复诊(mrsInfo!病人ID, mlng挂号科室ID), 1, 0)
        End If
    End If
End Sub

Private Sub txt门诊号_GotFocus()
    If InStr(";" & mstrPrivs & ";", ";允许修改门诊号;") > 0 Then
        '允许修改门诊号是才全部选中
        Call zlControl.TxtSelAll(txt门诊号)
    End If
End Sub

Private Sub txt门诊号_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        
        If txt门诊号.Enabled And txt门诊号.Visible And mintNOLength > 0 And mblnCheckNOValidity Then
        '如果手工输入了异常的门诊号则提示
            If Len(txt门诊号.Text) > mintNOLength + 1 Then
                MsgBox "注意,输入的门诊号过大,请确认是否输入正常!", vbInformation, gstrSysName
                txt门诊号.SetFocus
                txt门诊号.SelStart = 0: txt门诊号.SelLength = Len(txt门诊号.Text)
                Exit Sub
            End If
        End If
        
        Call zlCommFun.PressKey(vbKeyTab)
    ElseIf KeyAscii = 32 Then
        KeyAscii = 0
        If txt门诊号.Text = "" Then
            txt门诊号.Text = zlGet门诊号
            mintNOLength = Len(txt门诊号.Text)      '用来判断修改门诊号时的异常输入
        End If
        If ActiveControl Is txt门诊号 Then
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    ElseIf InStr("1234567890" & Chr(8), Chr(KeyAscii)) = 0 Or InStr(";" & mstrPrivs & ";", ";允许修改门诊号;") = 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txt门诊号_Validate(Cancel As Boolean)
    '如果病人有门诊号,则不可清除
    If txt门诊号.Text = "" Then
        If Not mrsInfo Is Nothing Then
            txt门诊号.Text = Nvl(mrsInfo!门诊号)
        End If
    End If
End Sub

Private Sub txt年龄_LostFocus()
    zlCommFun.OpenIme
End Sub

Private Sub txt年龄_Validate(Cancel As Boolean)
    Dim strBirth As String
    If txt年龄.Locked Then Exit Sub
    txt年龄.Text = Trim(txt年龄.Text)
    If Not IsNumeric(txt年龄.Text) And Trim(txt年龄.Text) <> "" Then
        cbo年龄单位.ListIndex = -1: cbo年龄单位.Visible = False: txt年龄.Width = 1320
    ElseIf cbo年龄单位.Visible = False Then
        cbo年龄单位.ListIndex = 0: cbo年龄单位.Visible = True: txt年龄.Width = 600
    End If
    '69026,冉俊明,2014-8-8,检查输入年龄
    If txt年龄.Visible And Trim(txt年龄.Text <> "") Then
        If mobjfrmPatiInfo.mobjPubPatient Is Nothing Then Exit Sub
        If mobjfrmPatiInfo.mobjPubPatient.CheckPatiAge(Trim(txt年龄.Text) & IIf(cbo年龄单位.Visible, cbo年龄单位.Text, "")) = False Then
            Cancel = True: txt年龄.SetFocus: Exit Sub
        End If
    End If
    
    If txt年龄.Tag <> txt年龄.Text Then
        With mobjfrmPatiInfo '更正出生日期
            .txt年龄.Text = txt年龄.Text
            .txt年龄.Tag = txt年龄.Text
            If .cbo年龄单位.ListCount = 0 Then CopyCboTofrmPatiInfo
            .cbo年龄单位.ListIndex = cbo年龄单位.ListIndex
            .cbo年龄单位.Visible = cbo年龄单位.Visible
            If Not IsDate(txt出生日期.Text) Then mblnGetBirth = True
            .mblnChange = False
            '125451：控制是否允许通过年龄计算出生日期
            If mblnGetBirth Then
    '                .txt出生日期.Text = ReCalcBirth(.txt年龄.Text, .cbo年龄单位.Text)
                If mobjfrmPatiInfo.mobjPubPatient.ReCalcBirthDay(Trim(txt年龄.Text) & IIf(cbo年龄单位.Visible, cbo年龄单位.Text, ""), strBirth) Then
                    .txt出生日期.Text = Format(strBirth, "yyyy-mm-dd")
                    .txt出生时间.Text = Format(strBirth, "hh:mm")
                End If
            End If
            .mblnChange = True
        End With
        
        txt年龄.Tag = txt年龄.Text
        '89130:李南春,2015/10/13,更新出生日期
        mblnChange = False
        txt出生日期.Text = mobjfrmPatiInfo.txt出生日期.Text
        txt出生时间.Text = mobjfrmPatiInfo.txt出生时间.Text
        mblnChange = True
        Call ShowRegistFromInput
        Call ReLoadCardFee(, True)
    End If
End Sub
Private Function CheckNoValied(ByVal lngRow As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '功能：检查指定行的号别是否有效
    '返回：有效,返回true,否则返回False
    '编制：刘兴洪
    '日期：2010-08-17 16:00:11
    '说明：31922
    '------------------------------------------------------------------------------------------------------------------------
    If InStr(1, mstrPrivs, ";临时挂号;") > 0 Or mblnStation Or mbytMode <> 0 Then
        CheckNoValied = True: Exit Function
    End If
    With vsfPlan
        If Val(.Cell(flexcpData, lngRow, .ColIndex("号别"))) = 1 Then
            '31922
            '不能挂此号
            MsgBox "号别『" & .TextMatrix(lngRow, .ColIndex("号别")) & "』不在有效范围内或你权限不足,不能挂号,请检查!", vbInformation + vbOKOnly + vbDefaultButton1
            Exit Function
        End If
    End With
    CheckNoValied = True
End Function

Private Sub txt号别_Change()
'功能：根据输入号别显示内容
    Dim strInfo As String, i As Integer
    Dim blnChkLimit As Boolean
    
    '清除上一张单据号
    mlng挂号科室ID = 0
    txt科室.Text = ""
    txtSN.Text = ""
    mlngPreRow = 0
        
    If mbytInState = 1 Then Exit Sub
    If chkCancel.Value = 1 Or chkPrint.Value = 1 Then Exit Sub
    If mblnUnChange Then Exit Sub
    
    '刷新号别直接从缓存中读取数据
    If vsfPlan.Tag = "" Then
        mblnManualInput = True
        Call ShowPlans(, Len(txt号别) > 0 And IsNumeric(Trim(txt号别.Text)), False)
        mblnManualInput = False
    End If
    
    If Trim(txt号别.Text) = "" Then
        chk病历费.Enabled = mbln病历费
        lblFree.Visible = False
        Exit Sub
    End If
    
    '上次挂号的缴款情况,新号时清除
    txt缴款.Text = "0.00": txt找补.Text = "0.00"
    
    If txt号别.Text = "+" Then '仅购病历
        txtSN.Text = ""
        txtSN.Enabled = False
        
        mlng挂号科室ID = UserInfo.部门ID
        If Not mrsInfo Is Nothing Then
            Call Init费别(mobjfrmPatiInfo.chk复诊.Value = 0, True)
        Else
            Call Init费别(True, mblnNewCard)
        End If
        Call ShowRegistFromInput
        Call zlCommFun.PressKey(vbKeyTab)
    ElseIf (IsNumeric(Trim(txt号别.Text)) And Len(Trim(txt号别.Text)) = gint号长 Or vsfPlan.Rows = 2) Or vsfPlan.Tag <> "" Then
        If vsfPlan.Tag = "" Then
            If vsfPlan.Rows = 2 And Trim(txt号别.Text) <> vsfPlan.TextMatrix(vsfPlan.Row, GetCol("号别")) Then
                '当前号别列表只有一行时，如果没有输完整号别，不自动匹配，除非按回车
                Exit Sub
            End If
            '定位表格中的号别
            For i = 1 To vsfPlan.Rows - 1
                If Trim(vsfPlan.TextMatrix(i, GetCol("号别"))) = Trim(txt号别.Text) Then
                    If CheckNoValied(i) = False Then
                         txt号别.Text = "": txt号别.SetFocus: Exit Sub
                    End If
                    Call vsfPlan_LeaveCell
                    vsfPlan.Row = i: vsfPlan.RowSel = i
                    vsfPlan.Col = vsfPlan.ColIndex("IDS"): vsfPlan.ColSel = vsfPlan.Cols - 1
                    Call vsfPlan_EnterCell
                    SetGridTop i
                    Exit For
                End If
            Next
            '号表中无安排时要求重输
            If i = vsfPlan.Rows Then
                txt号别.Text = "": txt号别.SetFocus: Exit Sub
            End If
        End If
        
        '病案权限控制
        If vsfPlan.TextMatrix(vsfPlan.Row, GetCol("病案")) <> "" Then
            If InStr(mstrPrivs, ";建立病案;") = 0 Then
                MsgBox "该号别要求给病人建立门诊病案，但你没有建立病案的权限。不能继续挂号！", vbInformation, gstrSysName
                txt号别 = "": txt号别.SetFocus: Exit Sub
            End If
            Call SetPatiInfoEnabled(True, mrsInfo Is Nothing) '问题号:58843
            If mrs家庭地址 Is Nothing And Not mblnStructAdress Then Call Load家庭地址
        Else
            Call SetPatiInfoEnabled(False, mrsInfo Is Nothing) '问题号:58843
        End If
        
        If mbytMode = 1 Then
            blnChkLimit = vsfPlan.TextMatrix(vsfPlan.Row, GetCol("限约")) <> ""
            If blnChkLimit = False Then
                blnChkLimit = vsfPlan.TextMatrix(vsfPlan.Row, GetCol("限号")) <> ""
            End If
        Else
            blnChkLimit = vsfPlan.TextMatrix(vsfPlan.Row, GetCol("限号")) <> ""
        End If
        '限号控制
        If chkCancel.Value = 0 And blnChkLimit And Not mblnFinishReg Then
            '问题:26962 日期:2009-12-25 11:46:30
            If zlCheck限约或限号数(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("记录ID"))) = False Then Exit Sub
        End If
        
        '确定当前序号
        txtSN.Enabled = vsfPlan.TextMatrix(vsfPlan.Row, GetCol("序号控制")) <> ""
        If txtSN.Enabled And vsfList.Tag = "" And vsfList.Visible Then
            txtSN.Text = GetCurrSN(, Not mTy_Para.bln随机序号选择)
            If Val(txtSN.Text) = 0 Then
                txtSN.Text = ""
                If CheckArangement = False Then Exit Sub
            Else
                Call LocateSN(Val(txtSN.Text))
            End If
        End If
        Dim blnCancel As Boolean
        
        '装入挂号内容
        '费别事件中调用ShowRegistFromInput
        mstrPre费别 = ""
        
        '72168
        mlng挂号科室ID = Abs(vsfPlan.RowData(vsfPlan.Row))
        If Not mrsInfo Is Nothing Then
            Call Init费别(mobjfrmPatiInfo.chk复诊.Value = 0, True)
        Else
            Call Init费别(True, mblnNewCard)
        End If
        
        If CheckIsPrice Or mRegistFeeMode = EM_RG_记帐 Then
            Call SetUndisplayBalance
        Else
            Call SetShowBalance
        End If
        Call zlCommFun.PressKey(vbKeyTab)
    End If
    
End Sub

Private Function GetCurrSN(Optional ByVal lngCurMaxSN As Long = -1, Optional ByVal blnGetLapseNO As Boolean = False) As Long
'功能:获取当前号别的最大可用序号
'     全部都用完时返回0
'    blngetlapseNo:是否从无效号以后开始算
'     lngCurMaxSN-当明最大使用号
    Dim i           As Integer
    Dim j           As Integer
    Dim lngMaxSn    As Long
    Dim lngSN       As Long
    Dim intStart    As Integer
    Dim lngTmp      As Long
    Dim blnUnitReg  As Boolean
    Dim lngMaxLapse As Long '最大无效号码
    If mbytMode = 1 And mblnUnitReg And Not mrsUnitReg Is Nothing Then
        blnUnitReg = True
    End If
    
'    If (mViewMode = v_专家号 Or mViewMode = v_专家号分时段) And Not mTy_Para.bln随机序号选择 And blnGetLapseNO Then
'        lngMaxLapse = GetMaxLapseNO
'    End If
    
    mtyRegPlanState.lngSelNO = 0
    mtyRegPlanState.lngSelX = 0
    mtyRegPlanState.lngSelY = 0
    mtyRegPlanState.strSelTime = ""
   
   If Not mrsSNState Is Nothing Or blnUnitReg Then
ReGet:
        If mrsSNState Is Nothing And mbytMode = 1 Then Set mrsSNState = GetSNState(Val(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("记录ID"))))
        mrsSNState.Filter = ""
        If mrsSNState.RecordCount > 0 Or blnUnitReg Then
        
            If lngCurMaxSN = -1 And mViewMode = v_专家号分时段 Then
                With vsfList
                    i = vsfList.Row
                    j = vsfList.Col
                    If .TextMatrix(i, j) <> "" Then
                        If .Cell(flexcpForeColor, i, j) <> vbRed And .Cell(flexcpForeColor, i, j) <> vbBlue And .Cell(flexcpForeColor, i, j) <> vbGreen And .Cell(flexcpForeColor, i, j) <> vbGrayText And .Cell(flexcpForeColor, i, j) <> &HC000C0 Then
                           lngTmp = Val(Get时段(i, j, False))
                           mrsSNState.Filter = "序号=" & lngTmp & " And 状态 <> 0"
                            If mrsSNState.RecordCount = 0 And lngTmp > lngMaxLapse Then
                                    GetCurrSN = lngTmp
                                    mtyRegPlanState.lngSelNO = lngTmp
                                    mtyRegPlanState.lngSelX = i
                                    mtyRegPlanState.lngSelY = j
                                    mtyRegPlanState.strSelTime = Get时段(i, j, True)
                                    Exit Function
                            End If
                        End If
                    End If
                End With
            End If
            
            
           If lngCurMaxSN = -1 And mViewMode = v_专家号 Then
               lngTmp = 0
               mrsSNState.Filter = "预约=0 and 状态=1"
                Do While Not mrsSNState.EOF
                   If lngTmp < Val(mrsSNState!序号) Then lngTmp = Val(mrsSNState!序号)
                   mrsSNState.MoveNext
                Loop
                mrsSNState.Filter = 0
               If lngTmp <> 0 Then lngCurMaxSN = lngTmp
            End If
            
            
            intStart = IIf(mViewMode = v_专家号分时段 Or mViewMode = V_普通号分时段, 1, 0)
            For i = 0 To vsfList.Rows - 1
                For j = intStart To vsfList.Cols - 1
                    Select Case mViewMode
                    Case V_普通号, v_专家号:
                        lngSN = Val(vsfList.TextMatrix(i, j))
                        If vsfList.Cell(flexcpForeColor, i, j) = &HC000C0 And mbytMode = 1 Then
                            lngSN = -1
                        End If
                        
                    Case v_专家号分时段:
                        With vsfList
                            If .Cell(flexcpForeColor, i, j) = vbGrayText Or .Cell(flexcpForeColor, i, j) = &HC000C0 Then
                                lngSN = -1
                            Else
                               lngSN = IIf(Trim(.TextMatrix(i, j)) = "", -1, Val(Get时段(i, j, False)))
                               If lngSN < lngMaxLapse And mTy_Para.bln随机序号选择 = False Then lngSN = -1
                               
                               '如果如果已经是最后一个序号了,需要检查是否存在加号,以及是否随机序号选择,随机序号选择,时 允许选择已经退号的序号 'lgf
                               If lngSN = mtyRegPlanState.lngLastNO And lngSN > 0 And mtyRegPlanState.blnAdditionalNumber And Not mTy_Para.bln随机序号选择 Then lngSN = -1
                            End If
                        End With
                    Case Else
                       Exit Function
                    End Select
                    '73411:默认序号的问题
                    If lngSN > -1 Then
                        mrsSNState.Filter = "序号=" & lngSN & " And 状态 <> 0"
                        '问题号:52335
                        If mrsSNState.RecordCount = 0 Then
                            lngMaxSn = lngSN
                            mblnStateChange = True
                            vsfList.Select i, j
                            mblnStateChange = False
                            mtyRegPlanState.lngSelNO = lngSN
                            mtyRegPlanState.lngSelX = i
                            mtyRegPlanState.lngSelY = j
                            If mViewMode = V_普通号分时段 Or mViewMode = v_专家号分时段 Then
                                '只有分时段,才存在时段
                                mtyRegPlanState.strSelTime = Get时段(i, j, True)
                            End If
                            Exit For
                        End If
                    End If
                    
                Next
                
                If lngMaxSn = lngSN Then Exit For
            Next
            If lngCurMaxSN > 0 And lngMaxSn = 0 Then
                '刘兴洪:???
                '主要是解决预约最大+1后,还有预约的情况,所以又从1开始检查是否有未选择的.
                '如:预约从5开始;到了7已经是最大号了,因此再从1开始取.
               ' lngCurMaxSN = -1: GoTo ReGet:
            End If
            GetCurrSN = lngMaxSn
        Else
            Select Case mViewMode
                Case v_专家号分时段:
                     vsfList.Redraw = False
                    For i = 0 To vsfList.Rows - 1
                        For j = 1 To vsfList.Cols - 1
                            If (vsfList.Cell(flexcpForeColor, i, j) = vbBlue Or vsfList.Cell(flexcpForeColor, i, j) = vbBlack) And vsfList.TextMatrix(i, j) <> "" Then
                                GetCurrSN = Val(Get时段(i, j, False))
                                mtyRegPlanState.lngSelNO = GetCurrSN
                                mtyRegPlanState.lngSelX = i
                                mtyRegPlanState.lngSelY = j
                                mtyRegPlanState.strSelTime = Get时段(i, j, True)
                                vsfList.Redraw = True
                                Exit Function
                            End If
                        Next
                    Next
                    vsfList.Redraw = True
                Case Else:
                  If mbytMode = 1 And mblnUnitReg And Not mrsUnitReg Is Nothing Then
                      mrsUnitReg.Filter = "序号=1"
                      If mrsUnitReg.RecordCount = 0 Then GetCurrSN = 1
                      mrsUnitReg.Filter = 0
                  Else
                    GetCurrSN = 1
                  End If
            End Select
        End If
    End If

End Function


Private Sub txt号别_GotFocus()
    Call zlControl.TxtSelAll(txt号别)
    
    If gblnLED And mbytMode <> 1 And mbytInState = 0 And txt号别.Text = "" And mblnLEDKey Then
        zl9LedVoice.Speak "#14" '请问你挂什么科
    End If
    mblnLEDKey = False
End Sub

Private Sub txt号别_KeyDown(KeyCode As Integer, Shift As Integer)
'上下移动号别,以便快速选择
    Select Case KeyCode
        Case vbKeyUp
            If vsfPlan.Row - 1 >= vsfPlan.FixedRows Then
                KeyCode = 0
                vsfPlan_LeaveCell
                vsfPlan.Row = vsfPlan.Row - 1
                vsfPlan_EnterCell
            End If
        Case vbKeyDown
            If vsfPlan.Row + 1 <= vsfPlan.Rows - 1 Then
                KeyCode = 0
                vsfPlan_LeaveCell
                vsfPlan.Row = vsfPlan.Row + 1
                vsfPlan_EnterCell
            End If
    End Select
End Sub

Private Sub txt号别_KeyPress(KeyAscii As Integer)
    '上次挂号的缴款情况,新号时清除
    txt缴款.Text = "0.00": txt找补.Text = "0.00"
    txt合计.Text = Format(mcur合计 + GetRegistMoney, "0.00")
    Call Set连续挂号
    
    If KeyAscii = Asc("/") And Trim(txt号别.Text) = "" Then
        '预约接收时,如果单据号输入的是"/",则自动弹出小窗口,供预约挂号用"
        KeyAscii = 0:        Call ShowBookSeled
        Exit Sub
    End If
    
    If KeyAscii = Asc("+") Then
        If mbytInState = 0 And (Not mbln病历费 Or picBookingDate.Visible Or mblnStation) Then
            KeyAscii = 0: Exit Sub '预约时不允许单独买病历
        End If
        '问题:27493
    ElseIf KeyAscii = Asc("-") Then
        KeyAscii = 0
        If chkShowAll.Enabled And chkShowAll.Visible Then
            If chkShowAll.Value = 0 Then
                chkShowAll.Value = 1
            Else
                chkShowAll.Value = 0
            End If
        End If
    ElseIf KeyAscii = Asc(".") Then
        '相关于按回退键
        KeyAscii = 0: zlCommFun.PressKey vbKeyBack
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        If CheckNoValied(vsfPlan.Row) = False Then
             txt号别.Text = "": txt号别.SetFocus: Exit Sub
        End If
        
        vsfPlan.Tag = vsfPlan.TextMatrix(vsfPlan.Row, GetCol("号别"))
        If vsfPlan.Tag <> "" Then
            If txt号别.Text <> vsfPlan.Tag Then
                txt号别.Text = vsfPlan.Tag  '自动调用change事件
            Else
                Call txt号别_Change
            End If
            vsfPlan.Tag = ""
        End If
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If InStr("1234567890+ABCDEFGHIJKLMNOPQRSTUVWXYZ" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    '问题号:110228,焦博,2017/07/20,挂号时过滤号别刷新不出来
    If txt号别.SelLength > 0 Then
        Set mrsPlan = Nothing
    End If
End Sub

Private Sub txt科室_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub txt科室_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub txt年龄_GotFocus()
    'txt年龄.IMEMode = vbIMEOff
    Call zlCommFun.OpenIme(True)
    Call zlControl.TxtSelAll(txt年龄)
End Sub

Private Sub txt年龄_KeyPress(KeyAscii As Integer)
    Dim blnTab As Boolean
    
    If txt年龄.Locked Then Exit Sub
    If KeyAscii = vbKeyReturn Then
        If txtPatient.Text <> "" And txt年龄.Text = "" And gbln年龄 Then Exit Sub
        
        If cbo年龄单位.Visible = False And IsNumeric(txt年龄.Text) Then
            Call txt年龄_Validate(False)
            If cbo年龄单位.Visible And cbo年龄单位.Enabled Then cbo年龄单位.SetFocus
        Else
            Call zlCommFun.PressKey(vbKeyTab)
        End If
        If Not IsNumeric(txt年龄.Text) And cbo年龄单位.Visible Then Call zlCommFun.PressKey(vbKeyTab)
    Else
        '仅仅限制几个 指定的特殊的字符
        If InStr("~·！@#￥%……&*（）——-+=|、？、。，~`!#$%^&*()-_=+|\/?<>,/<>", UCase(Chr(KeyAscii))) > 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txt年龄_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        glngTXTProc = GetWindowLong(txt年龄.Hwnd, GWL_WNDPROC)
        Call SetWindowLong(txt年龄.Hwnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txt年龄_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Call SetWindowLong(txt年龄.Hwnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Function bln发卡(ByVal strCardNo As String) As Boolean
'---------------------------------------------------------------------------------------------------------------------------------------------
'功能:判断当前是否为卡发操作 (不是发卡操作既是绑定卡操作)
'入参:
'编制:56599
'日期:2012-12-12 14:55:36
'---------------------------------------------------------------------------------------------------------------------------------------------
    Dim bln是否发卡 As Boolean
    '115168:李南春，2017/12/13，保存发卡的医疗卡类型
    If mCurSendCard.lng卡类别ID = 0 Then mCurSendCard = gCurSendCard
    '89572:李南春,2015/10/20,挂号发卡获取票据领用ID
    If mCurSendCard.bln严格控制 = True Then
        mlng磁卡领用ID = CheckUsedBill(5, IIf(mlng磁卡领用ID > 0, mlng磁卡领用ID, mCurSendCard.lng共用批次), strCardNo, mCurSendCard.lng卡类别ID)
        bln是否发卡 = IIf(mlng磁卡领用ID <= 0, False, True)
        If mCurSendCard.bln自制卡 = False Then
            bln是否发卡 = (mCurSendCard.bln是否发卡 = True)
        End If
    Else
        bln是否发卡 = mbln发卡
        If mblnAlwaysSend Then bln是否发卡 = True
        If mCurSendCard.bln自制卡 = False Then
            bln是否发卡 = (mCurSendCard.bln是否发卡 = True)
        End If
    End If
    bln发卡 = bln是否发卡
    mbln发卡 = bln是否发卡
End Function

Private Sub ClearmobjfrmPatiInfoFace(Optional blnClearCard As Boolean = True)
    Dim i As Integer
            
    With mobjfrmPatiInfo
        Call CopyCboTofrmPatiInfo '如果窗体没有Load,此时会Load调用Form_load事件
                
        .chk复诊.Value = 0
        .txt门诊号.Text = "": .txt门诊号.MaxLength = txt门诊号.MaxLength
        SetCboDefault .cbo费别
        SetCboDefault .cbo性别
            
        .txtPatiMCNO(0).Text = ""
        .txtPatiMCNO(0).Tag = ""
        .txtPatiMCNO(1).Text = ""
        
        If blnClearCard Then
            .mstrCard = ""
            .txt卡号.Text = ""
            If mblnNoClearPrompt = False Then lblPrompt.Caption = "": gCurSendCard.lng收费细目ID = 0: vsfPay.Height = 2220
            mblnNewCard = False
            mblnAddCardItem = False
        End If
        .txt密码.Text = ""
        .txt验证.Text = ""
        If mbytMode = 1 And mblnIDCardKind Then
            '31182:因为在读取身份证时,已经赋值不能再清空
        Else
            .txt年龄.Text = "": .txt年龄.MaxLength = txt年龄.MaxLength
            .txt年龄.Tag = ""
            .txt出生日期.Text = "____-__-__"
            .txt出生时间.Text = "__:__"
            Call zlControl.CboLocate(.cbo年龄单位, "岁")
            .cbo年龄单位.Tag = .cbo年龄单位.Text
            .txt身份证号.Text = ""
            .txt身份证号.Tag = ""
        End If
        .txtPatient.Text = "": .txtPatient.MaxLength = txtPatient.MaxLength
        
        SetCboDefault .cbo付款方式
        SetCboDefault .cbo国籍
        SetCboDefault .cbo民族
        SetCboDefault .cbo婚姻
        SetCboDefault .cbo职业
        
        
        .txt单位名称.Text = ""
        .txt单位名称.Tag = ""
        .txt单位电话.Text = ""
        .txt单位邮编.Text = ""
        .txt区域.Text = ""
        .cbo家庭地址.Text = ""
        .txt家庭邮编.Text = ""
        .txt家庭电话.Text = ""
        .txt过敏反应.Text = ""
        '问题号:40005
        .txt联系人电话.Text = ""
        .cbo联系人关系.ListIndex = -1
        .txtMobile = ""
        .txt联系人身份证.Text = ""
        .txt联系人姓名.Text = ""
        .txtBirthLocation.Text = ""
        .txtRegLocation.Text = ""
        .txt户口地址邮编.Text = ""
        '89242:李南春,2015/12/7,清空病人地址信息
        .padd家庭地址.Value = ""
        .padd户口地址.Value = ""
        '82649:李南春,2015/2/13,清除监护人信息
        .txt监护人.Text = ""
        For i = 1 To .msh过敏.Rows - 1
            .msh过敏.TextMatrix(i, 0) = ""
            .msh过敏.TextMatrix(i, 1) = "" '问题号:56599
            .msh过敏.RowData(i) = 0
        Next
        '问题号:56599
        .msh过敏.Rows = 2
        .Clear健康档案
        If .mblnNewPatient = False Then
            '74430,冉俊明,2014-7-7,挂号中的病人信息编辑功能中提供采集照片功能
            .imgPatient.Picture = Nothing
        End If
    End With
End Sub

Private Function LoadzlIDKindPatiInfor(objPati As zlIDKind.PatiInfor) As Boolean
    'IDKind_Read事件中,新病人加载信息到发卡界面
    ClearmobjfrmPatiInfoFace True
Call SetCboDefault(cbo医疗类别)
      Call zlControl.CboLocate(cbo性别, objPati.性别)
      
         
    With mobjfrmPatiInfo
        .txtPatient.Text = txtPatient.Text: .txtPatient.MaxLength = txtPatient.MaxLength
        
             
          If 1 = 1 Then
        Else
            .txt卡号.Tag = 0
        End If
        If Not mrsInfo Is Nothing Then
            .mlng病人ID = Val(Nvl(mrsInfo!病人ID))
        Else
            .mlng病人ID = 0
        End If
        
        
        .cbo性别.ListIndex = cbo性别.ListIndex
        .cbo年龄单位.ListIndex = cbo年龄单位.ListIndex
        .txt年龄.Text = txt年龄.Text: .txt年龄.MaxLength = txt年龄.MaxLength
        .txt年龄.Tag = txt年龄.Text
        .cbo家庭地址.Text = cbo家庭地址.Text
        .txtRegLocation = cbo户口地址.Text
        '89242:李南春,2015/12/7,读取病人地址信息
        Call .padd家庭地址.LoadStructAdress(padd家庭地址.value省, padd家庭地址.value市, padd家庭地址.value区县, padd家庭地址.value乡镇, padd家庭地址.value详细地址)
        Call .padd户口地址.LoadStructAdress(padd户口地址.value省, padd户口地址.value市, padd户口地址.value区县, padd户口地址.value乡镇, padd户口地址.value详细地址)
        .txt门诊号.Text = txt门诊号.Text: .txt门诊号.MaxLength = txt门诊号.MaxLength
        .cbo付款方式.ListIndex = cbo付款方式.ListIndex
        .cbo费别.ListIndex = cbo费别.ListIndex
        .cbo费别.Locked = cbo费别.Locked
        .cbo费别.TabStop = cbo费别.TabStop
        '问题号:40005
        If Not mrsInfo Is Nothing Then
            .txt联系人身份证.Text = Nvl(mrsInfo!联系人身份证号)
            .txt联系人姓名.Text = Nvl(mrsInfo!联系人姓名)
            .txt联系人电话.Text = Nvl(mrsInfo!联系人电话)
            .cbo联系人关系.ListIndex = cbo.FindIndex(.cbo联系人关系, Nvl(mrsInfo!联系人关系), True)
            If .cbo联系人关系.ListIndex = -1 And Nvl(mrsInfo!联系人关系) <> "" Then
                .cbo联系人关系.ListIndex = 8: .txt其他关系.Text = Nvl(mrsInfo!联系人关系)
            End If
        End If
    End With
    
     With mobjfrmPatiInfo
        txtPatient.Text = .txtPatient.Text  '调用Change事件
        
        cbo性别.ListIndex = .cbo性别.ListIndex
        txt年龄.Text = .txt年龄.Text
        txt年龄.Tag = txt年龄.Text
        cbo年龄单位.ListIndex = .cbo年龄单位.ListIndex
        Call txt年龄_Validate(False)
        
        cbo家庭地址.Text = .cbo家庭地址.Text
        cbo户口地址.Text = .txtRegLocation.Text
        '89242:李南春,2015/12/7,读取病人地址信息
        Call padd家庭地址.LoadStructAdress(.padd家庭地址.value省, .padd家庭地址.value市, .padd家庭地址.value区县, .padd家庭地址.value乡镇, .padd家庭地址.value详细地址)
        Call padd户口地址.LoadStructAdress(.padd户口地址.value省, .padd户口地址.value市, .padd户口地址.value区县, .padd户口地址.value乡镇, .padd户口地址.value详细地址)
        txt门诊号.Text = .txt门诊号.Text
        cbo付款方式.ListIndex = .cbo付款方式.ListIndex
        cbo费别.ListIndex = .cbo费别.ListIndex
        
         
    End With
     
End Function
Private Function isCheckInputIDCard(ByVal strInput As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '功能：检查当前输入的是否身份证号
    '入参：strInput-输入的值
    '返回:如果是身分证号,则返回true,否则返回False
    '编制：刘兴洪
    '日期：2010-07-14 16:37:51
    '说明：31182
    '      自动识别身份证,主要从三个方面来确定
    '      a.前缀为".":暂没用
    '      b.前缀后的字符长度为15位或18位(因为身份证目前只有15位和18位区分)
    '      c.前缀后中根据身份证取出来出生日期，看取出的值是否为身份证.
    '------------------------------------------------------------------------------------------------------------------------
    Dim strTemp As String, strDate As String
    'If Left(strInput, 1) = "." Then Exit Function
    If Len(strTemp) = 15 Or Len(strTemp) = 18 Then Exit Function '本身包含了识别符的.因此需要在原身份证前+1位
    strDate = zlCommFun.GetIDCardDate(strInput)
    If strDate = "" Then Exit Function
    If IsDate(strDate) = False Then Exit Function
    isCheckInputIDCard = True
End Function

Private Sub cbo户口地址_GotFocus()
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub cbo户口地址_LostFocus()
    Call zlCommFun.OpenIme
End Sub

Private Sub GetPatient(objCard As zlIDKind.Card, ByVal strInput As String, ByVal blnCard As Boolean, Optional blnInputIDCard As Boolean = False, _
                        Optional ByRef Cancel As Boolean, Optional ByRef blnCertificate As Boolean = False)
    '功能：获取病人信息
    '参数：blnCard=是否就诊卡刷卡
    '
    '         blnInputIDCard-是否身份证刷卡
    '出参:Cancel-为true表示返回的放弃读取病人信息
    Dim strSQL As String, lng卡类别ID As Long, strPassWord As String, strErrMsg As String
    Dim lng病人ID As Long, blnHavePassWord As Boolean
    Dim rsTmp As ADODB.Recordset, strTemp As String
    Dim blnSame As Boolean, blnCancel As Boolean
    Dim cur余额 As Currency, curMoney As Currency
    Dim strInputInfo As String '保存传入的输入文本 避免在使用身份证号 对病人进行查找后 被替换成"-" 病人ID的情况
    Dim i As Integer, strPati As String, strRegist As String, rsRegist As ADODB.Recordset
    Dim vRect As RECT, str非在院 As String
    Dim bln医保号 As Boolean, dbl家属余额 As Double
    Dim intMsg As VbMsgBoxResult, blnReload As Boolean
    Dim blnOtherType As Boolean '非法卡类别
    Dim lngRow As Long, lngCol As Long
    
    strInputInfo = strInput
    lbl险类.Caption = ""
    lbl险类.Visible = False
    
    On Error GoTo errH
    bln医保号 = False
    
    If objCard Is Nothing Then Set objCard = IDKind.GetCurCard
     
    strSQL = "Select  A.病人ID,A.门诊号,A.住院号,A.就诊卡号,A.费别,A.医疗付款方式,A.姓名,A.性别,A.年龄,A.出生日期,A.出生地点,A.身份证号,A.其他证件,A.身份,A.职业,A.民族,A.病人类型, " & _
             "A.国籍,A.籍贯,A.区域,A.学历,A.婚姻状况,A.家庭地址,A.家庭电话,A.家庭地址邮编,A.监护人,A.联系人姓名,A.联系人关系,A.联系人地址,A.联系人电话,A.户口地址, " & _
             "A.户口地址邮编,A.Email,A.QQ,A.合同单位id,A.工作单位,A.单位电话,A.单位邮编,A.单位开户行,A.单位帐号,A.担保人,A.担保额,A.担保性质,A.就诊时间,A.就诊状态, " & _
             "A.就诊诊室,A.住院次数,A.当前科室id,A.当前病区id,A.当前床号,A.入院时间,A.出院时间,A.在院,A.IC卡号,A.健康号,A.医保号,A.险类,A.查询密码,A.登记时间,A.停用时间,A.锁定,A.联系人身份证号, " & _
             "B.名称 险类名称,A.查询密码 As 卡验证码,A.结算模式,A.手机号 From 病人信息 A,保险类别 B  Where A.险类 = B.序号(+) And A.停用时间 is NULL "
                 
    If mTy_Para.bln允许住院病人挂号 = False Then
        str非在院 = " And Not Exists(Select 1 From 病案主页 Where 病人ID=A.病人ID   And 主页ID<>0 And 主页ID=A.主页ID And Nvl(病人性质,0)=0 And 出院日期 is Null)"
    End If
   

    If blnCard And objCard.名称 Like "姓名*" And mstrYBPati = "" And InStr("-+*.", Left(strInput, 1)) = 0 Then     '刷卡
        If IDKind.Cards.按缺省卡查找 And Not IDKind.GetfaultCard Is Nothing Then
            lng卡类别ID = IDKind.GetfaultCard.接口序号
        ElseIf IDKind.GetCurCard.接口序号 > 0 Then
            lng卡类别ID = IDKind.GetCurCard.接口序号
        Else
            If lng卡类别ID = 0 Then lng卡类别ID = -1
        End If
        
        '短名|完成名|刷卡标志|卡类别ID|卡号长度|缺省标志(1-当前缺省;0-非缺省)|是否存在帐户(1-存在帐户;0-不存在帐户);…
        If gobjSquare.objSquareCard.zlGetPatiID(lng卡类别ID, strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then lng病人ID = 0

        If IDKind.IsMobileNO(strInput) And lng病人ID = 0 Then
            If gobjSquare.objSquareCard.zlGetPatiID("手机号", strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then lng病人ID = 0
        End If
        If lng病人ID <= 0 Then GoTo NewPati:
        strInput = "-" & lng病人ID
        blnHavePassWord = True
        '72868,冉俊明,2014-5-20,在门诊挂号管理的参数设置中未勾选“允许住院病人挂号”的参数，但是在院病人依然能够直接通过门诊挂号管理刷卡挂号
        strSQL = strSQL & " And A.病人ID=[2] " & str非在院
        mstr门诊号 = "": txt门诊号.TabStop = True
    ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then
        '门诊号
        strSQL = strSQL & " And A.门诊号=[2]" & str非在院
        If InStr(mstrPrivs, ";建立病案;") > 0 Then
            mstr门诊号 = Mid(strInput, 2) '记录输入的门诊号
            txt门诊号.TabStop = False
        End If
    ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then
        '病人ID
        strSQL = strSQL & " And A.病人ID=[2]" & _
        IIf(mstrYBPati <> "", "", str非在院)
        If mstrYBPati = "" Then mstr门诊号 = "": txt门诊号.TabStop = True
    ElseIf blnInputIDCard Then  '单独的身份证识别
        strInput = UCase(strInput)
        If gobjSquare.objSquareCard.zlGetPatiID("身份证", strInput, False, lng病人ID, strPassWord, strErrMsg, , , , , , , mblnUserCancel) = False Then lng病人ID = 0
        strInput = "-" & lng病人ID
        strSQL = strSQL & " And A.病人ID=[2] " & str非在院
        mstr门诊号 = "": txt门诊号.TabStop = True
        blnHavePassWord = True
    ElseIf blnCertificate Then
        strInput = UCase(strInput)
        If gobjSquare.objSquareCard.zlGetPatiID(objCard.名称, strInput, False, lng病人ID, strPassWord, strErrMsg, , , , , , blnCertificate) = False Then Exit Sub
        strInput = "-" & lng病人ID
        strSQL = strSQL & " And A.病人ID=[2] " & str非在院
        mstr门诊号 = "": txt门诊号.TabStop = True
        blnHavePassWord = True
    ElseIf objCard.名称 Like "姓名*" And IDKind.IsMobileNO(strInput) = True Then
        If gobjSquare.objSquareCard.zlGetPatiID("手机号", strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then Exit Sub
        strInput = "-" & lng病人ID
        strSQL = strSQL & " And A.病人ID=[2] " & str非在院
        mstr门诊号 = "": txt门诊号.TabStop = True
    Else
        Select Case objCard.名称
            Case "姓名", "姓名或就诊卡"
                '姓名
                blnSame = False
                If Not mrsInfo Is Nothing Then
                    mstr门诊号 = "": txt门诊号.TabStop = True
                    If txtPatient.Text = mrsInfo!姓名 Then blnSame = True
                End If
                If Not blnSame Then
                    If Not gblnSeekName Or gblnSeekName And Len(txtPatient.Text) < 2 Or mstr门诊号 <> "" Or mblnNewCard Then
                        Set mrsInfo = Nothing: Exit Sub
                    Else
                         '问题号:50485
                        strPati = _
                            " Select /*+Rule */distinct 1 as 排序ID,A.病人ID as ID,A.病人ID,A.姓名,A.性别,A.年龄,A.门诊号,A.出生日期,A.身份证号,A.家庭地址,A.工作单位,decode(b.卡号,Null,Null,'√') As 是否有医疗卡,A.手机号,A.就诊时间" & _
                            " From 病人信息 A, 病人医疗卡信息 B " & _
                            " Where Rownum <101 And a.病人ID=b.病人ID(+) And b.状态(+)=0 And B.卡类别ID(+)=[3]  And A.停用时间 is NULL And A.姓名 Like [1]" & str非在院 & _
                            IIf(gintNameDays = 0, "", " And Nvl(A.就诊时间,A.登记时间)>Trunc(Sysdate-[2])")
                            
                        strPati = strPati & " Union ALL " & _
                                "Select 0,0 as ID,-NULL,'[新病人]',NULL,NULL,-NULL,To_Date(NULL),NULL,NULL,NULL,NULL,NULL,To_Date(NULL) From Dual"
                        strPati = strPati & " Order by 排序ID,姓名"
                            
                        vRect = zlControl.GetControlRect(txtPatient.Hwnd)
                        Set rsTmp = zlDatabase.ShowSQLSelect(Me, strPati, 0, "病人查找", 1, "", "请选择病人", False, False, True, vRect.Left, vRect.Top, txtPatient.Height, blnCancel, False, True, strInput & "%", gintNameDays, Val(zlDatabase.GetPara("缺省医疗卡类别", glngSys, mlngModul, 0)))
                        If Not rsTmp Is Nothing Then
                            If rsTmp!ID = 0 Then '当作新病人
                                Set mrsInfo = Nothing
                                '82859:李南春,2015/4/8,病人基本信息调整
                                If mbytInState = 0 Then SetPatiInfoEnabled vsfPlan.TextMatrix(vsfPlan.Row, GetCol("病案")) <> "", mrsInfo Is Nothing
                                Exit Sub
                            Else '以病人ID读取
                                strInput = rsTmp!病人ID
                                strSQL = strSQL & " And A.病人ID=[1]"
                            End If
                        Else '取消选择
                            txtPatient.Text = ""
                            Set mrsInfo = Nothing: Exit Sub
                        End If
                    End If
                Else
                    '同一个病人时需要重新读取预交款信息
                    If mbytMode <> 1 And mstrYBPati = "" Then
                        Set rsTmp = GetMoneyInfo(mrsInfo!病人ID, , , 1, , , True)
                        cur余额 = 0: dbl家属余额 = 0: stbThis.Panels(4).ToolTipText = ""
                        Do While Not rsTmp.EOF
                            cur余额 = cur余额 + Val(Nvl(rsTmp!预交余额))
                            cur余额 = cur余额 - Val(Nvl(rsTmp!费用余额))
                            If Val(Nvl(rsTmp!家属)) = 1 Then
                                dbl家属余额 = Val(Nvl(rsTmp!预交余额)) - Val(Nvl(rsTmp!费用余额))
                            End If
                            rsTmp.MoveNext
                        Loop
                        If cur余额 > 0 Then
                            Call ShowDeposit(True): Call ShowMedicareInfo(False)
                            mdbl预交余额 = cur余额
                            For i = 1 To vsfPay.Rows - 1
                                If vsfPay.RowData(i) = 0 And vsfPay.TextMatrix(i, 0) <> "" Then
                                    vsfPay.TextMatrix(i, 6) = mdbl预交余额
                                End If
                            Next i
                            stbThis.Panels(4).Text = "门诊预交余额:" & mdbl预交余额
                            If Round(dbl家属余额, 6) <> 0 Then stbThis.Panels(4).ToolTipText = "含家属预交：" & Format(dbl家属余额, "0.00")
                            
                            '医生站挂号缺省使用预交款
                            curMoney = GetRegistMoney
                            '77786,冉俊明,2014-9-2,勾选优先使用预交款缴款,挂号时,没有默认减少冲减
                            '74550,冉俊明,2014-7-2,在病人来院就诊,医生在门诊医生站挂号时能够选择结算方式(包含性质为7的一卡通结算)
                            If (gblnPrePayPriority Or (mblnStation And Not mblnStationPrice And Not cbo结算方式.Visible)) And cur余额 >= curMoney Then
'                                txt预交支付.Text = Format(curMoney, "0.00")
                            Else
'                                txt预交支付.Text = "0.00"
                            End If
                        End If
                    End If
                    Call zlQueryEMPIPatiInfo
                    Exit Sub
                End If
            Case "医保号"
                strInput = UCase(strInput)
                mstr门诊号 = "": txt门诊号.TabStop = True
                bln医保号 = True
                If mblnOlnyBJYB And zlCommFun.ActualLen(strInput) >= 9 Then
                    '仅北京医保才有效:见问题:问题:26982
                    strSQL = strSQL & " And A.医保号 like [3] " & str非在院
                    strTemp = Left(strInput, 9) & "%"
                Else
                    strSQL = strSQL & " And A.医保号=[1]" & str非在院
                End If
            Case "手机号"
                If IDKind.IsMobileNO(strInput) = False Then Exit Sub
                If gobjSquare.objSquareCard.zlGetPatiID("手机号", strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then Exit Sub
                strInput = "-" & lng病人ID
                strSQL = strSQL & " And A.病人ID=[2] " & str非在院
                mstr门诊号 = "": txt门诊号.TabStop = True
            Case "身份证号", "身份证", "二代身份证"
                strInput = UCase(strInput)
                If gobjSquare.objSquareCard.zlGetPatiID("身份证", strInput, False, lng病人ID, strPassWord, strErrMsg, , , , , , , mblnUserCancel) = False Then lng病人ID = 0
                strInput = "-" & lng病人ID
                strSQL = strSQL & " And A.病人ID=[2] " & str非在院
                mstr门诊号 = "": txt门诊号.TabStop = True
                blnHavePassWord = True
            Case "IC卡号", "IC卡"
                strInput = UCase(strInput)
                mstr门诊号 = "": txt门诊号.TabStop = True
                If gobjSquare.objSquareCard.zlGetPatiID("IC卡", strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then lng病人ID = 0
                strInput = "-" & lng病人ID
                strSQL = strSQL & " And A.病人ID=[2] " & str非在院
                blnHavePassWord = True
            Case "门诊号"
                If Not IsNumeric(strInput) Then strInput = "0"
                '72868,冉俊明,2014-5-20,在门诊挂号管理的& str非在院参数设置中未勾选“允许住院病人挂号”的参数，但是在院病人依然能够直接通过门诊挂号管理刷卡挂号
                strSQL = strSQL & " And A.门诊号=[1]" & str非在院
                If InStr(mstrPrivs, ";建立病案;") > 0 Then
                    mstr门诊号 = strInput
                    txt门诊号.TabStop = False
                End If

             Case Else
                '其他类别的,获取相关的病人ID
                If objCard.接口序号 > 0 Then
                    lng卡类别ID = objCard.接口序号
                    If gobjSquare.objSquareCard.zlGetPatiID(lng卡类别ID, strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then lng病人ID = 0
                    blnOtherType = True
                Else
                    If gobjSquare.objSquareCard.zlGetPatiID(objCard.名称, strInput, False, lng病人ID, _
                        strPassWord, strErrMsg) = False Then lng病人ID = 0
                End If
                'If lng病人ID <= 0 Then GoTo NotFoundPati:
                '72868,冉俊明,2014-5-20,在门诊挂号管理的参数设置中未勾选“允许住院病人挂号”的参数，但是在院病人依然能够直接通过门诊挂号管理刷卡挂号
                strSQL = strSQL & " And A.病人ID=[2]" & str非在院
                strInput = "-" & lng病人ID
                blnHavePassWord = True
        End Select
    End If
    If blnInputIDCard And Not mrsInfo Is Nothing Then
        If mrsInfo.State <> 1 Then GoTo ReadPati:
        '原来有病人,又按身份证读取,则可能存在补身份证的情况:
        '1.如果未找到,则是补份证
        '2.如果找到了,则以新的病人为准(通过提示来选择)
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strInput, Mid(strInput, 2), strTemp)
        If rsTmp.EOF Then
            mobjfrmPatiInfo.txt身份证号 = txtIDCard.Text
            Call zlQueryEMPIPatiInfo
            Exit Sub
        End If
        If Nvl(rsTmp!姓名) <> Trim(txtPatient.Text) And Trim(txtPatient.Text) <> "" Then
            If gbln身份证唯一 Then
                intMsg = MsgBox("注意:" & vbCrLf & _
                                 "      录入的身份证号的姓名为『" & Nvl(rsTmp!姓名) & " 』与录入姓名『" & Trim(txtPatient.Text) & " 』" & vbCrLf & _
                                 "      不一致,是否以『" & Nvl(rsTmp!姓名) & " 』为准进行挂号？", vbQuestion + vbYesNo, gstrSysName)
                If intMsg = vbNo Then intMsg = vbCancel
            Else
            
                intMsg = MsgBox("注意:" & vbCrLf & _
                                 "      录入的身份证号的姓名为『" & Nvl(rsTmp!姓名) & " 』与录入姓名『" & Trim(txtPatient.Text) & " 』" & vbCrLf & _
                                 "      不一致,请检查!   " & vbCrLf & _
                                 "【是】表示以身份证查找的病人进行挂号" & vbCrLf & _
                                 "【否】表示以输入的姓名进行挂号,身份证号更改为当前录入的身份证号" & vbCrLf & _
                                 "【取消】表示身份证号录入错误,重新录入身份证号" & vbCrLf & _
                                "", vbQuestion + vbYesNoCancel, gstrSysName)
            End If
            If intMsg = vbCancel Then
              
                Cancel = True: Exit Sub
            End If
            If intMsg = vbYes Then
                Set mrsInfo = rsTmp
                txtPatient.Text = Nvl(rsTmp!姓名)
                blnReload = True
            End If
            If intMsg = vbNo Then
                mobjfrmPatiInfo.txt身份证号 = txtIDCard.Text
            End If
        End If
    Else
ReadPati:
        Set mrsInfo = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strInput, Mid(strInput, 2), strTemp)
    End If
    
    '82859:李南春,2015/4/8,病人基本信息调整
    If mbytInState = 0 Then SetPatiInfoEnabled vsfPlan.TextMatrix(vsfPlan.Row, GetCol("病案")) <> "", True
        
    strInput = strInputInfo
    Call ClearmobjfrmPatiInfoFace(IIf(mblnNewCard, False, True))
    If blnInputIDCard Then mobjfrmPatiInfo.txt身份证号.Text = strInput
    If Not mrsInfo.EOF Then
         '在发卡时 当操作员 使用病人的医疗卡提取病人信息时 发现新发的卡和病人拥有的卡是同种类型的情况下
         '使用原来的卡 不再发卡给病人
         If mblnNewCard And mbytMode = 0 And blnCard And lng卡类别ID = gCurSendCard.lng卡类别ID Then
              mblnNewCard = False
              Call ClearmobjfrmPatiInfoFace(IIf(mblnNewCard, False, True))
         End If
        '31182:检查用身份证查找的病人是否与输入的姓名一致
        If mbytMode = 1 Or mbytMode = 2 Then
            Call zlAutoCalcBackLists(Val(Nvl(mrsInfo!病人ID))) '自动加入黑名单
        End If
        If blnInputIDCard Then
                If Nvl(mrsInfo!姓名) <> Trim(txtPatient.Text) And Trim(txtPatient.Text) <> "" Then
                    If gbln身份证唯一 Then
                        intMsg = MsgBox("注意:" & vbCrLf & _
                                         "      录入的身份证号的姓名为『" & Nvl(mrsInfo!姓名) & " 』与录入姓名『" & Trim(txtPatient.Text) & " 』" & vbCrLf & _
                                         "      不一致,是否以『" & Nvl(mrsInfo!姓名) & " 』为准进行挂号？", vbQuestion + vbYesNo, gstrSysName)
                        If intMsg = vbNo Then intMsg = vbCancel
                    Else
                    
                            intMsg = MsgBox("注意:" & vbCrLf & _
                                             "      录入的身份证号的姓名为『" & Nvl(mrsInfo!姓名) & " 』与录入姓名『" & Trim(txtPatient.Text) & " 』" & vbCrLf & _
                                             "      不一致,请检查!   " & vbCrLf & _
                                             "【是】表示以身份证查找的挂号对象 " & vbCrLf & _
                                             "【否】表示以输入的姓名为挂号对象，需要重新建立病人档案" & vbCrLf & _
                                             "【取消】表示身份证号录入错误,重新录入身份证号" & vbCrLf & _
                                            "", vbQuestion + vbYesNoCancel, gstrSysName)
                    End If
                    If intMsg = vbCancel Then
                        Cancel = True: Exit Sub
                    End If
                    If intMsg = vbNo Then GoTo NewPati:
                    blnReload = True
                End If
        End If
        
        If blnCertificate Then
            If Nvl(mrsInfo!姓名) <> Trim(txtPatient.Text) And Trim(txtPatient.Text) <> "" Then
                intMsg = MsgBox("注意:" & vbCrLf & _
                                 "      录入的证件号码的姓名为『" & Nvl(mrsInfo!姓名) & " 』与录入姓名『" & Trim(txtPatient.Text) & " 』" & vbCrLf & _
                                 "      的信息不一致,是否以证件查找的姓名为挂号对象？   " & vbCrLf & _
                                "", vbQuestion + vbYesNo, gstrSysName)
                If intMsg = vbNo Then
                    Cancel = True: Exit Sub
                End If
            End If
        End If
        
        '102230,调用外挂部件接口
        If (mbytMode = 0 Or mbytMode = 1) And mbytInState = 0 _
            And Not (mbytMode = 0 And mbytInState = 0 And mstrNoIn <> "") Then
            If PatiValiedCheckByPlugIn(mlngModul, Val(Nvl(mrsInfo!病人ID)), _
                "<YSXM>" & NeedName(cbo医生.Text) & "</YSXM>") = False Then
                Set mrsInfo = Nothing: txtPatient.Text = ""
                Cancel = True:  Exit Sub
            End If
        End If
        
        If Not IsNull(mrsInfo!险类名称) Then
            lbl险类.Caption = "" & mrsInfo!险类名称
            lbl险类.Visible = True
        End If
        
        txtPatient.Text = Nvl(mrsInfo!姓名) '会调用Change事件
        '在调用txtPatient_Change事件后在门诊号和病人姓名都为空的情况下 无法识别该病人信息 出现错误
        '对这类数据库数据错误不再进行后续的处理
        If mrsInfo Is Nothing Then Cancel = True: Exit Sub
        '74428：李南春，2014-7-8，病人姓名显示颜色处理
        Call SetPatiColor(txtPatient, Nvl(mrsInfo!病人类型), IIf(IsNull(mrsInfo!险类名称), txtPatient.ForeColor, vbRed))

        '113999:李南春,2017/11/14,根据发卡性质进行控制
        If Check发卡性质(Val(Nvl(mrsInfo!病人ID)), IIf(mCurSendCard.lng卡类别ID = 0, gCurSendCard.lng卡类别ID, mCurSendCard.lng卡类别ID), Trim(mobjfrmPatiInfo.txt卡号) <> "") = True Then
            cmdCard.Enabled = True
        Else
            cmdCard.Enabled = gCurSendCard.lng发卡性质 <> 1
            mobjfrmPatiInfo.mstrCard = ""
            mobjfrmPatiInfo.txt卡号.Text = ""
            mobjfrmPatiInfo.txt密码.Text = ""
            mobjfrmPatiInfo.txt验证.Text = ""
            If mblnNoClearPrompt = False Then lblPrompt.Caption = ""
            mblnNewCard = False
            mblnAddCardItem = False
        End If
        cmdCard.Enabled = cmdCard.Enabled And Not (mblnStation And mTy_Para.bln挂号必须刷卡)
        mstrPassWord = strPassWord
        If Not blnHavePassWord Then mstrPassWord = Nvl(mrsInfo!卡验证码)
        cbo性别.ListIndex = cbo.FindIndex(cbo性别, Nvl(mrsInfo!性别), True) '年龄在后面根据出生日期算
        cbo家庭地址.Text = IIf(Nvl(mrsInfo!家庭地址) = "", Nvl(mrsInfo!户口地址), Nvl(mrsInfo!家庭地址))
        cbo户口地址.Text = Nvl(mrsInfo!户口地址)
        '89242:李南春,2015/12/7,读取病人地址信息
        Call zlReadAddrInfo(padd家庭地址, Val(Nvl(mrsInfo!病人ID)), 0, 3, cbo家庭地址.Text)
        Call zlReadAddrInfo(padd户口地址, Val(Nvl(mrsInfo!病人ID)), 0, 4, cbo户口地址.Text)
        txtPatient.PasswordChar = ""
        
        '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
        txtPatient.IMEMode = 0
        Call zlControl.CboSetIndex(cbo费别.Hwnd, cbo.FindIndex(cbo费别, "" & mrsInfo!费别, True))
        
        If Not blnInputIDCard Or blnReload Or txt门诊号.Text = "" Then
            txt门诊号.Text = Nvl(mrsInfo!门诊号, mstr门诊号)
'            txt门诊号.Enabled = (Val(txt门诊号.Text) = 0)
        End If
        
        If txt门诊号.Text = "" And txt门诊号.Enabled And gbln自动门诊号 Then
            txt门诊号.Text = zlGet门诊号
        End If
        
        If blnReload Then
            txtIDCard.Text = Nvl(mrsInfo!身份证号, txtIDCard.Text) '身份证号:31182
            txtIDCard.Tag = Nvl(mrsInfo!身份证号, txtIDCard.Text)  '以便反过来再查
        Else
            If Not blnInputIDCard Then
                txtIDCard.Text = Nvl(mrsInfo!身份证号)
                txtIDCard.Tag = Nvl(mrsInfo!身份证号)
            Else
                txtIDCard.Tag = txtIDCard.Text
            End If
        End If
    
        '医疗付款方式
        If Not IsNull(mrsInfo!医疗付款方式) Then
            cbo付款方式.ListIndex = cbo.FindIndex(cbo付款方式, mrsInfo!医疗付款方式, True)
        ElseIf mstrYBPati <> "" Then
            cbo付款方式.ListIndex = cbo.FindIndex(cbo付款方式, "1", True)
        End If
        If Not IsNull(mrsInfo!医保号) And mlngOutModeMC <> 0 Then Call SetCboDefault(cbo医疗类别)
        
        If Not blnInputIDCard Or blnReload Then
            txt出生日期.Text = Format(IIf(IsNull(mrsInfo!出生日期), "____-__-__", mrsInfo!出生日期), "YYYY-MM-DD")
            If Not IsNull(mrsInfo!出生日期) Then
                txt年龄.Text = ReCalcOld(CDate(mrsInfo!出生日期), cbo年龄单位, mrsInfo!病人ID) '根据出生日期重算年龄
                
                txt出生时间.Text = Format(mrsInfo!出生日期, "HH:MM")
            Else
                txt出生时间.Text = "__:__"
                txt出生日期.Text = ReCalcBirth(txt年龄.Text, cbo年龄单位.Text)
            End If
        End If
        
        '详细病人信息设置
        txt证件.Tag = "": txt证件.Text = ""
        Call CopyInfoTofrmPatiInfo
        With mobjfrmPatiInfo
    
            If mblnOlnyBJYB And bln医保号 Then
                .txtPatiMCNO(0).Text = strInput
            Else
                .txtPatiMCNO(0).Text = "" & Nvl(mrsInfo!医保号)
            End If
            .txtPatiMCNO(0).Tag = "" & Nvl(mrsInfo!医保号)
            .txtPatiMCNO(1).Text = .txtPatiMCNO(0).Text
            If Not blnInputIDCard Or blnReload Then
                Call LoadOldData("" & mrsInfo!年龄, .txt年龄, .cbo年龄单位)
                .mblnChange = False
                .txt出生日期.Text = Format(IIf(IsNull(mrsInfo!出生日期), "____-__-__", mrsInfo!出生日期), "YYYY-MM-DD")
                .mblnChange = True
                
                If Not IsNull(mrsInfo!出生日期) Then
                    .txt年龄.Text = ReCalcOld(CDate(.txt出生日期.Text), .cbo年龄单位, mrsInfo!病人ID) '根据出生日期重算年龄
                    .txt年龄.Tag = .txt年龄.Text
                    If CDate(.txt出生日期.Text) - CDate(mrsInfo!出生日期) <> 0 Then .txt出生时间.Text = Format(mrsInfo!出生日期, "HH:MM")
                Else
                    .txt出生时间.Text = "__:__"
                    .mblnChange = False
                    .txt出生日期.Text = ReCalcBirth(.txt年龄.Text, .cbo年龄单位.Text)
                    .mblnChange = True
                End If
            End If
            
            Call SetmobjfrmPatiInfo
            '90875:李南春,2016/8/19,从证件列表中获取当前证件类型的号码
            If IDKind证件.IDKind <> IDKind证件.GetKindIndex("身份证号") Then
                With mobjfrmPatiInfo.vsCertificate
                    For lngRow = 1 To .Rows - 1
                        For lngCol = 0 To .Cols - 1 Step 2
                            If .TextMatrix(lngRow, lngCol) = IDKind证件.GetCurCard.名称 Then
                                txt证件.Tag = .TextMatrix(lngRow, lngCol + 1)
                                txt证件.Text = txt证件.Tag
                                Exit For
                            End If
                        Next
                    Next
                End With
            End If
                
            txt年龄.Text = .txt年龄.Text
            txt年龄.Tag = txt年龄.Text
            cbo年龄单位.ListIndex = .cbo年龄单位.ListIndex
            cbo年龄单位.Tag = cbo年龄单位.Text
            Call txt年龄_Validate(False)
            
            If mlng挂号科室ID > 0 Then .chk复诊.Value = IIf(Check复诊(mrsInfo!病人ID, mlng挂号科室ID), 1, 0)
            If mbytMode = 1 And Not blnInputIDCard Then
                .txt身份证号 = txtIDCard.Text
            End If
            .mstr身份证号 = Nvl(mrsInfo!身份证号)
            imgPatiPic.Picture = .imgPatient.Picture
            txt家庭电话.Text = .txt家庭电话
            .mstr出生日期 = .txt出生日期.Text
            .mstr出生时间 = .txt出生时间.Text
            .mstr年龄单位 = IIf(IsNumeric(txt年龄.Text), cbo年龄单位.Text, "")
            .mstr年龄 = txt年龄.Text
            .mstr性别 = NeedName(cbo性别.Text)
            .mstr姓名 = txtPatient.Text
            .mstr身份证号 = txtIDCard.Text
            mstr出生日期 = .txt出生日期.Text
            .txtMobile.Text = Nvl(mrsInfo!手机号)
        End With
        mstr年龄单位 = IIf(IsNumeric(txt年龄.Text), cbo年龄单位.Text, "")
        mstr年龄 = txt年龄.Text
        mstr性别 = NeedName(cbo性别.Text)
        mstr姓名 = txtPatient.Text
        
        '病人预交款信息
        If mbytMode <> 1 And mstrYBPati = "" Then
            Set rsTmp = GetMoneyInfo(mrsInfo!病人ID, , , 1, , , True)
            cur余额 = 0: dbl家属余额 = 0: stbThis.Panels(4).ToolTipText = ""
            Do While Not rsTmp.EOF
                cur余额 = cur余额 + Val(Nvl(rsTmp!预交余额))
                cur余额 = cur余额 - Val(Nvl(rsTmp!费用余额))
                If Val(Nvl(rsTmp!家属)) = 1 Then
                    dbl家属余额 = Val(Nvl(rsTmp!预交余额)) - Val(Nvl(rsTmp!费用余额))
                End If
                rsTmp.MoveNext
            Loop
            If cur余额 > 0 Then
                Call ShowMedicareInfo(False): Call ShowDeposit(True)
                stbThis.Panels(4).Text = "门诊预交余额:" & Format(cur余额, "0.00")
                stbThis.Panels(4).AutoSize = sbrContents
                
                mdbl预交余额 = cur余额
                For i = 1 To vsfPay.Rows - 1
                    If vsfPay.RowData(i) = 0 And vsfPay.TextMatrix(i, 0) <> "" Then
                        vsfPay.TextMatrix(i, 6) = mdbl预交余额
                    End If
                Next i
                If Round(dbl家属余额, 6) <> 0 Then stbThis.Panels(4).ToolTipText = "含家属预交：" & Format(dbl家属余额, "0.00")
                
                '医生站挂号缺省使用预交款
                curMoney = GetRegistMoney
                '77786,冉俊明,2014-9-2,勾选优先使用预交款缴款,挂号时,没有默认减少冲减
                '74550,冉俊明,2014-7-2,在病人来院就诊,医生在门诊医生站挂号时能够选择结算方式(包含性质为7的一卡通结算)
                If (gblnPrePayPriority Or (mblnStation And Not mblnStationPrice And Not cbo结算方式.Visible)) And cur余额 >= curMoney Then
'                    txt预交支付.Text = Format(curMoney, "0.00")
                Else
'                    txt预交支付.Text = "0.00"
                End If
            Else
                Call ShowDeposit(False)
            End If
        End If
        mstr门诊号 = "": txt门诊号.TabStop = True
        mblnIDCardKind = False
        Call zlQueryEMPIPatiInfo
    Else
NewPati:
        txt门诊号.Enabled = True
        
        '82859:李南春,2015/4/8,病人基本信息调整
        If mbytInState = 0 Then SetPatiInfoEnabled vsfPlan.TextMatrix(vsfPlan.Row, GetCol("病案")) <> "", mrsInfo Is Nothing
        
        mblnIDCardKind = False
        If objCard.名称 Like "姓名*" And blnCard = False Then
            lng卡类别ID = 0
        End If
        If Not (mblnCard Or IsCardType(IDKind, "IC卡") _
            Or (gCurSendCard.lng卡类别ID = lng卡类别ID And lng卡类别ID > 0)) And blnInputIDCard = False And lng卡类别ID <= 0 Then txtPatient.Text = ""    '刷卡时不能清除,因为如果是发新卡要以此传入弹出窗体
        
        If lng病人ID = 0 And lng卡类别ID <> gCurSendCard.lng卡类别ID Then
            If lng卡类别ID <= 0 And Not IDKind.GetfaultCard Is Nothing Then lng卡类别ID = IDKind.GetfaultCard.接口序号
            If lng卡类别ID <> 0 And lng卡类别ID <> gCurSendCard.lng卡类别ID Then
                Call InitSendCardPreperty(mlngModul, lng卡类别ID)
                 
                 cmdCard.ToolTipText = "绑定" & gCurSendCard.str卡名称 & ": F10"
            End If
           If lng卡类别ID <= 0 And blnOtherType Then Cancel = True: txtPatient.Text = ""
        End If
            
        If isCheckInputIDCard(strInput) Then
            Dim str年龄单位 As String, str年龄 As String
            txtIDCard.Text = strInput     '身份证号:31182
            txtIDCard.Tag = strInput
            
            strTemp = zlGetIDCardSex(strInput)
            zlControl.CboLocate cbo性别, strTemp
            zlControl.CboLocate mobjfrmPatiInfo.cbo性别, strTemp
            
            mobjfrmPatiInfo.txt身份证号 = strInput
            mobjfrmPatiInfo.txt出生日期 = zlCommFun.GetIDCardDate(strInputInfo)
            If txt年龄.Text = "" Then
                str年龄 = zlGetIDCardAge(mobjfrmPatiInfo.txt出生日期, str年龄单位)
                If str年龄单位 <> "" Then
                    zlControl.CboLocate cbo年龄单位, str年龄单位
                    txt年龄.Text = str年龄
                     zlControl.CboLocate mobjfrmPatiInfo.cbo年龄单位, str年龄单位
                      mobjfrmPatiInfo.txt年龄.Text = str年龄
                      mobjfrmPatiInfo.txt年龄.Tag = str年龄
                End If
            End If
            '67213:李南春,2014/10/23,保存身份证上的信息
            mblnIDCardKind = IDKind.IDKind = IDKind.GetKindIndex("身份证号")
            If mblnIDCardKind Then
                IDKind.IDKind = IDKind.GetKindIndex("姓名")
            End If
            mblnIDCardKind = blnInputIDCard Or IDKind.IDKind = IDKind.GetKindIndex("身份证号")
        End If
        Set mrsInfo = Nothing
    End If
    
    If Not mrsInfo Is Nothing Then
        If mrsInfo.State = 1 Then
            If mrsInfo.RecordCount <> 0 Then
                If Not IsNull(mrsInfo!病人ID) Then
                    strRegist = "Select a.号别, b.名称 As 部门, a.执行人 As 医生, d.名称 As 项目, a.登记时间" & vbNewLine & _
                                "From 病人挂号记录 A, 部门表 B, 门诊费用记录 C, 收费项目目录 D " & vbNewLine & _
                                "Where a.病人id = [1] And a.记录状态 = 1 And a.No = c.No(+) And c.记录性质 = 4 And c.序号 = 1 And c.收费细目Id = d.Id And a.记录性质 = 1 And a.执行部门id = b.Id" & vbNewLine & _
                                "Order By a.登记时间 Desc"
                    Set rsRegist = zlDatabase.OpenSQLRecord(strRegist, Me.Caption, mrsInfo!病人ID)
                    If Not rsRegist.EOF Then
                        stbThis.Panels(2).Text = "上次挂号:" & "号码:" & rsRegist!号别 & ",科室:" & rsRegist!部门 & IIf(IsNull(rsRegist!医生), "", ",医生:" & rsRegist!医生) & ",项目:" & rsRegist!项目 & ",时间:" & Format(rsRegist!登记时间, "yyyy-mm-dd hh:mm:ss")
                    End If
                End If
            End If
        End If
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub zlQueryEMPIPatiInfo()
    '功能：从EMPI平台获取病人信息
    '日期：2016/10/9 10:47:13
    '编制：李南春
    '说明：101170
    Dim rsTmp As ADODB.Recordset, lng病人ID As Long, strDiff As String, strMsgInfo As String
    Dim strSQL As String
    If mblnNotEMPIQuery Then Exit Sub
    If CreatePlugInOK(mlngModul) = False Then Exit Sub
    If Trim(txtPatient.Text) = "" Then Exit Sub
    If mbytMode <> 0 And mbytMode <> 2 Or mbytInState <> 0 Or chkCancel.Value = 1 Then Exit Sub

    On Error GoTo Errhand
    If zlInitMEPIPati(rsTmp) = False Then Exit Sub
    
    If mrsInfo Is Nothing Then
        lng病人ID = 0
    ElseIf mrsInfo.State = 0 Then
        lng病人ID = 0
    Else
        lng病人ID = Val(Nvl(mrsInfo!病人ID))
    End If
    
    With rsTmp
        .AddNew
        !病人ID = lng病人ID
        !门诊号 = txt门诊号.Text
        !医保号 = mobjfrmPatiInfo.txtPatiMCNO(0).Text
        !身份证号 = mobjfrmPatiInfo.txt身份证号.Text
        !姓名 = txtPatient.Text
        !性别 = zlStr.NeedName(cbo性别.Text)
        If IsDate(txt出生日期.Text) Then
            !出生日期 = Format(txt出生日期.Text & " " & IIf(IsDate(txt出生时间.Text), txt出生时间.Text, "00:00"), "YYYY-MM-DD HH:MM")
        Else
            !出生日期 = ""
        End If
        !出生地点 = mobjfrmPatiInfo.txtBirthLocation.Text
        !国籍 = zlStr.NeedName(mobjfrmPatiInfo.cbo国籍.Text)
        !民族 = zlStr.NeedName(mobjfrmPatiInfo.cbo民族.Text)
        !职业 = zlStr.NeedName(mobjfrmPatiInfo.cbo职业.Text)
        !工作单位 = mobjfrmPatiInfo.txt单位名称.Text
        !婚姻状况 = zlStr.NeedName(mobjfrmPatiInfo.cbo婚姻.Text)
        !家庭电话 = mobjfrmPatiInfo.txt家庭电话.Text
        !联系人电话 = mobjfrmPatiInfo.txt联系人电话.Text
        !单位电话 = mobjfrmPatiInfo.txt单位电话.Text
        !家庭地址 = cbo家庭地址.Text
        !家庭地址邮编 = mobjfrmPatiInfo.txt家庭邮编.Text
        !户口地址 = cbo户口地址.Text
        !户口地址邮编 = mobjfrmPatiInfo.txt户口地址邮编.Text
        !单位邮编 = mobjfrmPatiInfo.txt单位邮编.Text
        !联系人姓名 = mobjfrmPatiInfo.txt联系人姓名.Text
        !联系人关系 = zlStr.NeedName(mobjfrmPatiInfo.cbo联系人关系.Text)
        .Update
    End With
    'EMPI没有找到病人信息,直接返回
    Dim rsOut As New ADODB.Recordset
    Err = 0: On Error Resume Next
    mlngEMPI病人ID = 0
    If gobjPlugIn.EMPI_QueryPatiInfo(glngSys, mlngModul, rsTmp, rsOut) = False Then
        Call zlPlugInErrH(Err, "EMPI_QueryPatiInfo")
        Err.Clear: Set mobjfrmPatiInfo.mrsEMPIOut = Nothing: Exit Sub
    End If
    Err.Clear: On Error GoTo 0
    Set mobjfrmPatiInfo.mrsEMPIOut = rsOut
    If mobjfrmPatiInfo.mrsEMPIOut Is Nothing Then Exit Sub
    If mobjfrmPatiInfo.mrsEMPIOut.RecordCount = 0 Then Exit Sub
    mobjfrmPatiInfo.mrsEMPIOut.MoveFirst
    On Error Resume Next
    With mobjfrmPatiInfo.mrsEMPIOut
        '104905:李南春,2017/1/12,根据EMPI传回的病人ID，查找病人
        '接收查阅退号肯定有病人ID
        mlngEMPI病人ID = Val(Nvl(!病人ID))
        If lng病人ID <> mlngEMPI病人ID And mlngEMPI病人ID <> 0 Then
            mblnNotEMPIQuery = True
            Call GetPatient(IDKind.GetCurCard, "-" & mlngEMPI病人ID, False)
            mblnNotEMPIQuery = False
            If mrsInfo.EOF Then
                lng病人ID = 0
            Else
                lng病人ID = mlngEMPI病人ID
            End If
        End If
        
        mobjfrmPatiInfo.mstrPlugChange = ""
        If Nvl(!医保号) <> "" Then
            mobjfrmPatiInfo.txtPatiMCNO(0).Text = Nvl(!医保号)
            mobjfrmPatiInfo.txtPatiMCNO(1).Text = mobjfrmPatiInfo.txtPatiMCNO(0).Text
        End If
        If mbln基本信息调整 Or lng病人ID = 0 Then
            If Nvl(!身份证号) <> "" Then txtIDCard.Text = Nvl(!身份证号)
            If Nvl(!姓名) <> "" Then txtPatient.Text = Nvl(!姓名): mstrPrePati = Nvl(!姓名)
            If Nvl(!性别) <> "" Then cbo性别.ListIndex = cbo.FindIndex(cbo性别, Nvl(!性别), True)
            If Nvl(!出生日期) <> "" Then
                txt出生日期.Text = Format(Nvl(!出生日期), "YYYY-MM-DD")
                txt出生时间.Text = Format(Nvl(!出生日期), "HH:MM")
            End If
        Else
            If Nvl(!姓名) <> "" And txtPatient.Text <> Nvl(!姓名) Then strDiff = ",姓名"
            If Nvl(!性别) <> "" And cbo性别.ListIndex <> cbo.FindIndex(cbo性别, Nvl(!性别), True) Then strDiff = strDiff & ",性别"
            If Nvl(!出生日期) <> "" And Format(Nvl(!出生日期), "YYYY-MM-DD HH:MM:SS") <> Format(txt出生日期.Text & " " & txt出生时间.Text, "YYYY-MM-DD HH:MM:SS") Then strDiff = strDiff & ",出生日期"
            If Nvl(!身份证号) <> "" And txtIDCard.Text <> Nvl(!身份证号) Then strDiff = strDiff & ",身份证号"
        End If
        If InStr(";" & mstrPrivs & ";", ";允许修改门诊号;") > 0 And Exist门诊号(Nvl(!门诊号), lng病人ID) = False Then
            If Nvl(!门诊号) <> "" Then txt门诊号.Text = Nvl(!门诊号)
        Else
            If Nvl(!门诊号) <> "" And txt门诊号.Text <> Nvl(!门诊号) Then strDiff = strDiff & ",门诊号"
        End If
        If Nvl(!出生地点) <> "" Then mobjfrmPatiInfo.txtBirthLocation.Text = Nvl(!出生地点)
        If Nvl(!国籍) <> "" Then mobjfrmPatiInfo.cbo国籍.ListIndex = cbo.FindIndex(mobjfrmPatiInfo.cbo国籍, Nvl(!国籍), True)
        If Nvl(!民族) <> "" Then mobjfrmPatiInfo.cbo民族.ListIndex = cbo.FindIndex(mobjfrmPatiInfo.cbo民族, Nvl(!民族), True)
        If Nvl(!职业) <> "" Then mobjfrmPatiInfo.cbo职业.ListIndex = cbo.FindIndex(mobjfrmPatiInfo.cbo职业, Nvl(!职业))
        If Nvl(!工作单位) <> "" Then mobjfrmPatiInfo.txt单位名称.Text = Nvl(!工作单位)
        If Nvl(!婚姻状况) <> "" Then mobjfrmPatiInfo.cbo婚姻.ListIndex = cbo.FindIndex(mobjfrmPatiInfo.cbo婚姻, Nvl(!婚姻状况), True)
        If Nvl(!家庭电话) <> "" Then txt家庭电话.Text = Nvl(!家庭电话)
        If Nvl(!联系人电话) <> "" Then mobjfrmPatiInfo.txt联系人电话.Text = Nvl(!联系人电话)
        If Nvl(!单位电话) <> "" Then mobjfrmPatiInfo.txt单位电话.Text = Nvl(!单位电话)
        If Nvl(!家庭地址) <> "" Then cbo家庭地址.Text = Nvl(!家庭地址): padd家庭地址.Value = Nvl(!家庭地址)
        If Nvl(!家庭地址邮编) <> "" Then mobjfrmPatiInfo.txt家庭邮编.Text = Nvl(!家庭地址邮编)
        If Nvl(!户口地址) <> "" Then cbo户口地址.Text = Nvl(!户口地址): padd户口地址.Value = Nvl(!户口地址)
        If Nvl(!户口地址邮编) <> "" Then mobjfrmPatiInfo.txt户口地址邮编.Text = Nvl(!户口地址邮编)
        If Nvl(!单位邮编) <> "" Then mobjfrmPatiInfo.txt单位邮编.Text = Nvl(!单位邮编)
        If Nvl(!联系人姓名) <> "" Then mobjfrmPatiInfo.txt联系人姓名.Text = Nvl(!联系人姓名)
        If Nvl(!联系人关系) <> "" Then mobjfrmPatiInfo.cbo联系人关系.ListIndex = cbo.FindIndex(mobjfrmPatiInfo.cbo联系人关系, Nvl(!联系人关系), True)
    End With
    Err = 0: On Error GoTo 0
    Call CopyInfoTofrmPatiInfo
    If lng病人ID <> 0 Then
        If strDiff <> "" Then strDiff = Mid(strDiff, 2)
        If mobjfrmPatiInfo.mstrPlugChange <> "" Then mobjfrmPatiInfo.mstrPlugChange = Mid(mobjfrmPatiInfo.mstrPlugChange, 2)
        If strDiff <> "" Then
            strMsgInfo = "病人的 " & strDiff & " 与EMPI信息不一致，因您不具有相应的权限或与其他病人信息冲突，本次不会进行更新。"
        End If
        If mobjfrmPatiInfo.mstrPlugChange <> "" Then
            If strMsgInfo <> "" Then strMsgInfo = strMsgInfo & vbNewLine
            strMsgInfo = strMsgInfo & "病人的 " & mobjfrmPatiInfo.mstrPlugChange & " 根据EMPI信息进行了调整,请注意检查！"
        End If
        If strMsgInfo <> "" Then MsgBox strMsgInfo, vbInformation, gstrSysName
        mobjfrmPatiInfo.mstrPlugChange = ""
    End If
    Exit Sub
Errhand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function zlSaveEMPIPatiInfo(ByVal blnNewPati As Boolean, ByVal lngPatiID As Long, ByVal lngClinicID As Long, ByRef strErrMsg As String) As Boolean
    '功能:上传病人信息到EMPI平台,如果平台信息保存失败，连同HIS数据一起回退
    '参数: In-lngPatiID 病人ID,lngClinicID 挂号ID
    '      Out-strErrMsg 错误信息，函数返回失败有效
    '返回:True-EMPI平台保存信息成功,False-保存失败
    '编制:李南春
    '说明:101170
    Dim blnCharge As Boolean, lngRet As Long
    If CreatePlugInOK(mlngModul) = False Then zlSaveEMPIPatiInfo = True: Exit Function
    If mbytMode <> 0 And mbytMode <> 2 Or mbytInState <> 0 Then zlSaveEMPIPatiInfo = True: Exit Function
    
    On Error GoTo Errhand
    If mobjfrmPatiInfo.mrsEMPIOut Is Nothing Then
        'EMPI没有病人信息，需要新建
        On Error Resume Next
        lngRet = gobjPlugIn.EMPI_AddPatiInfo(glngSys, mlngModul, lngPatiID, 0, lngClinicID, strErrMsg)
        Call zlPlugInErrH(Err, "EMPI_AddPatiInfo")
        If lngRet = 0 And Err.Number <> 438 Then Err.Clear: Exit Function
        Err.Clear: On Error GoTo Errhand
    Else
        '判断平台回传的信息是否发生改变
        With mobjfrmPatiInfo.mrsEMPIOut
            If InStr(";" & mstrPrivs & ";", ";允许修改门诊号;") > 0 And Exist门诊号(Nvl(!门诊号), lngPatiID) = False Then
                If txt门诊号.Text <> Nvl(!门诊号) Then blnCharge = True: GoTo EMPIModify
            End If
            If mobjfrmPatiInfo.txtPatiMCNO(0).Text <> Nvl(!医保号) Then blnCharge = True: GoTo EMPIModify
            If mbln基本信息调整 Or blnNewPati Then
                If txtIDCard.Text <> Nvl(!身份证号) Then blnCharge = True: GoTo EMPIModify
                If txtPatient.Text <> Nvl(!姓名) Then blnCharge = True: GoTo EMPIModify
                If cbo性别.ListIndex <> cbo.FindIndex(cbo性别, Nvl(!性别), True) Then blnCharge = True: GoTo EMPIModify
                If Format(txt出生日期.Text, "YYYY-MM-DD") <> Format(Nvl(!出生日期), "YYYY-MM-DD") Then blnCharge = True: GoTo EMPIModify
                If Format(txt出生时间.Text, "HH:MM") <> Format(Nvl(!出生日期), "HH:MM") Then blnCharge = True: GoTo EMPIModify
            End If
            If mobjfrmPatiInfo.txtBirthLocation.Text <> Nvl(!出生地点) Then blnCharge = True: GoTo EMPIModify
            If mobjfrmPatiInfo.cbo国籍.ListIndex <> cbo.FindIndex(mobjfrmPatiInfo.cbo国籍, Nvl(!国籍), True) Then blnCharge = True: GoTo EMPIModify
            If mobjfrmPatiInfo.cbo民族.ListIndex <> cbo.FindIndex(mobjfrmPatiInfo.cbo民族, Nvl(!民族), True) Then blnCharge = True: GoTo EMPIModify
            If mobjfrmPatiInfo.cbo职业.ListIndex <> cbo.FindIndex(mobjfrmPatiInfo.cbo职业, Nvl(!职业)) Then blnCharge = True: GoTo EMPIModify
            If mobjfrmPatiInfo.txt单位名称.Text <> Nvl(!工作单位) Then blnCharge = True: GoTo EMPIModify
            If mobjfrmPatiInfo.cbo婚姻.ListIndex <> cbo.FindIndex(mobjfrmPatiInfo.cbo婚姻, Nvl(!婚姻状况), True) Then blnCharge = True: GoTo EMPIModify
            If txt家庭电话.Text <> Nvl(!家庭电话) Then blnCharge = True: GoTo EMPIModify
            If mobjfrmPatiInfo.txt联系人电话.Text <> Nvl(!联系人电话) Then blnCharge = True: GoTo EMPIModify
            If mobjfrmPatiInfo.txt单位电话.Text <> Nvl(!单位电话) Then blnCharge = True: GoTo EMPIModify
            If cbo家庭地址.Text <> Nvl(!家庭地址) Then blnCharge = True: GoTo EMPIModify
            If mobjfrmPatiInfo.txt家庭邮编.Text <> Nvl(!家庭地址邮编) Then blnCharge = True: GoTo EMPIModify
            If cbo户口地址.Text <> Nvl(!户口地址) Then blnCharge = True: GoTo EMPIModify
            If mobjfrmPatiInfo.txt户口地址邮编.Text <> Nvl(!户口地址邮编) Then blnCharge = True: GoTo EMPIModify
            If mobjfrmPatiInfo.txt单位邮编.Text <> Nvl(!单位邮编) Then blnCharge = True: GoTo EMPIModify
            If mobjfrmPatiInfo.txt联系人姓名.Text <> Nvl(!联系人姓名) Then blnCharge = True: GoTo EMPIModify
            If mobjfrmPatiInfo.cbo联系人关系.ListIndex <> cbo.FindIndex(mobjfrmPatiInfo.cbo联系人关系, Nvl(!联系人关系), True) Then blnCharge = True: GoTo EMPIModify
        End With
    End If
EMPIModify:
    If blnCharge Then
        On Error Resume Next
        lngRet = gobjPlugIn.EMPI_ModifyPatiInfo(glngSys, mlngModul, lngPatiID, 0, lngClinicID, strErrMsg)
        Call zlPlugInErrH(Err, "EMPI_AddPatiInfo")
        If lngRet = 0 And Err.Number <> 438 Then Err.Clear: Exit Function
        Err.Clear: On Error GoTo 0
    End If
    zlSaveEMPIPatiInfo = True
    Exit Function
Errhand:
    strErrMsg = Err.Description
    Call SaveErrLog
End Function

Private Sub ShowDeposit(ByVal blnShow As Boolean)
'功能：显示/隐藏预交支付信息
    Dim i As Integer
    Dim intIndex As Integer
    If mbln连续挂号 Then Exit Sub
    If gblnPrice Then blnShow = False
    stbThis.Panels(4).Visible = blnShow

    If Not blnShow Then
        mdbl预交余额 = 0
        For i = 1 To vsfPay.Rows - 1
            If vsfPay.RowData(i) = 0 And vsfPay.TextMatrix(i, 0) <> "" Then
                vsfPay.TextMatrix(i, 6) = 0
            End If
        Next i
        stbThis.Panels(4).Text = "门诊预交余额:0.00"
        For i = 1 To vsfPay.Rows - 1
            If vsfPay.RowData(i) = 0 And vsfPay.TextMatrix(i, 0) <> "" Then vsfPay.RowHidden(i) = True
        Next i
    Else
        For i = 1 To vsfPay.Rows - 1
            If vsfPay.RowData(i) = 0 And vsfPay.TextMatrix(i, 0) <> "" Then vsfPay.RowHidden(i) = False
        Next i
    End If
End Sub

Private Sub ShowMedicareInfo(ByVal blnShow As Boolean)
'功能：显示/隐藏医保个人帐户支付信息
    Dim i As Integer
    If gblnPrice Then blnShow = False
    stbThis.Panels(3).Visible = blnShow
    If Not blnShow Then
        mdbl个帐余额 = 0
        stbThis.Panels(3).Text = "0.00"
        For i = 1 To vsfPay.Rows - 1
            If vsfPay.RowData(i) = 3 And vsfPay.TextMatrix(i, 0) <> "" Then vsfPay.RowHidden(i) = True
        Next i
    Else
        If MCPAR.使用个人帐户 Then
            For i = 1 To vsfPay.Rows - 1
                If vsfPay.RowData(i) = 3 And vsfPay.TextMatrix(i, 0) <> "" Then vsfPay.RowHidden(i) = False
            Next i
        End If
    End If
End Sub

Private Sub txtFact_GotFocus()
    zlControl.TxtSelAll txtFact
End Sub

Private Sub txtFact_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab)
    ElseIf Not ((KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or InStr("0123456789" & Chr(8), Chr(KeyAscii)) > 0) Then
        KeyAscii = 0
    ElseIf Len(txtFact.Text) = txtFact.MaxLength And KeyAscii <> 8 And txtFact.SelLength <> Len(txtFact) Then
        KeyAscii = 0
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub timPlan_Timer()
    If DateAdd("n", glngInterval, mDatLastRefresh) <= Now Then
        If chkPrint.Value = 1 Or chkCancel.Value = 1 Or chkBooking.Value = 1 Or vsfPlan.Enabled = False Then Exit Sub
        '自动定时刷新,不是正在挂号时,不是正在选择序号时
        If mcbrToolBar.Controls.Find(xtpControlButton, conMenu_View_Refresh).Enabled And mcbrToolBar.Controls.Find(xtpControlButton, conMenu_View_Refresh).Visible And txt号别.Text = "" And Not Me.ActiveControl Is vsfList Then RefreshFace
        mDatLastRefresh = Now
    End If
End Sub

Private Sub SetGridTop(intRow As Integer)
    Dim intRows As Integer
    intRows = vsfPlan.Height \ vsfPlan.RowHeight(1) - 2
    If vsfPlan.TopRow + intRows > intRow Then Exit Sub
    vsfPlan.TopRow = intRow
End Sub

Private Sub Load家庭地址()
    Dim strSQL As String, strFile As String
    Dim fld As Field, rsCheck As ADODB.Recordset
    Dim fso As Scripting.FileSystemObject
    Dim rsNew As ADODB.Recordset
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    strFile = App.Path & "\ZLAddressForRegEvent.Adtg"
    
    Set mrs家庭地址 = New ADODB.Recordset
    
    On Error Resume Next
    If fso.FileExists(strFile) Then
        mrs家庭地址.Open strFile, "Provider=MSPersist", adOpenKeyset, adLockOptimistic, adCmdFile   '仅Update时才锁定
    End If
    Err.Clear
    On Error GoTo errH
    
    If mrs家庭地址.State = 0 Then
        strSQL = "Select '系统' As 类别, 家庭地址 As 名称, Null As 简码, 1 As 次数" & vbNewLine & _
                "From 病人信息" & vbNewLine & _
                "Where 1 = 0" & vbNewLine & _
                "Union" & vbNewLine & _
                "Select '系统' As 类别, 名称, 简码, 1 As 次数 From 地区"

        Call zlDatabase.OpenRecordset(mrs家庭地址, strSQL, Me.Caption)            '必须是adUseClient才能建索引
        
        If Not mrs家庭地址.EOF Then
            '创建索引:名称,简码
            Set fld = mrs家庭地址.Fields(1)
            fld.Properties("Optimize") = True
            Set fld = mrs家庭地址.Fields(2)
            fld.Properties("Optimize") = True
            
            If fso.FileExists(strFile) Then
                Kill strFile
            End If
            mrs家庭地址.Save strFile, adPersistADTG
        End If
        mrs家庭地址.Close
        mrs家庭地址.Open strFile, "Provider=MSPersist", adOpenKeyset, adLockOptimistic, adCmdFile   '仅Update时才锁定
    Else
        strSQL = "Select '系统' As 类别, 家庭地址 As 名称, Null As 简码, 1 As 次数" & vbNewLine & _
                "From 病人信息" & vbNewLine & _
                "Where 1 = 0" & vbNewLine & _
                "Union" & vbNewLine & _
                "Select '系统' As 类别, 名称, 简码, 1 As 次数 From 地区 Where 1 = 0"
        Set rsCheck = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        If rsCheck.Fields(1).DefinedSize > mrs家庭地址.Fields(1).DefinedSize Or rsCheck.Fields(2).DefinedSize > mrs家庭地址.Fields(2).DefinedSize Then
            If fso.FileExists(strFile) Then
                Kill strFile
            End If
            strSQL = "Select '系统' As 类别, 家庭地址 As 名称, Null As 简码, 1 As 次数" & vbNewLine & _
                    "From 病人信息" & vbNewLine & _
                    "Where 1 = 0" & vbNewLine & _
                    "Union" & vbNewLine & _
                    "Select '系统' As 类别, 名称, 简码, 1 As 次数 From 地区"
            Set rsNew = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
            rsNew.Save strFile, adPersistXML
            mrs家庭地址.Close
            mrs家庭地址.Open strFile, "Provider=MSPersist", adOpenKeyset, adLockOptimistic, adCmdFile   '仅Update时才锁定
        End If
    End If
    
    lbl家庭地址.ToolTipText = "请定期备份本机[家庭地址]数据文件:" & strFile
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub cbo家庭地址_GotFocus()
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub cbo家庭地址_LostFocus()
    Call zlCommFun.OpenIme
End Sub

Private Sub cbo家庭地址_KeyDown(KeyCode As Integer, Shift As Integer)
    '此过程处理本机缓存数据的删除,以及按下拉键时弹出下拉列表
    '下拉列表弹出时,如果按下删除键时,则删除缓存记录
    
    Dim str家庭地址 As String
    
    If KeyCode = vbKeyDelete Then
        str家庭地址 = cbo家庭地址.Text
        
        If Not mrs家庭地址 Is Nothing And mTy_Para.byt家庭地址联想 = 1 Then
            If mrs家庭地址.State = 1 And str家庭地址 <> "" Then
                If cbo家庭地址.SelText = str家庭地址 And SendMessage(cbo家庭地址.Hwnd, CB_GETDROPPEDSTATE, 0, 0) = True Then
                    mrs家庭地址.Filter = "名称='" & str家庭地址 & "'"
                    If Not mrs家庭地址.EOF Then
                        mrs家庭地址.Delete adAffectCurrent
                        mrs家庭地址.Update
                    End If
                End If
            End If
        End If
    ElseIf KeyCode = vbKeyDown And cbo家庭地址.Text <> "" Then
        If SendMessage(cbo家庭地址.Hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 Then Call SendMessage(cbo家庭地址.Hwnd, CB_SHOWDROPDOWN, True, 0&)
    End If
End Sub

Private Sub cbo家庭地址_KeyUp(KeyCode As Integer, Shift As Integer)
    '此时text中已接收输入的信息
    '此事件处理删除和退格键,删除部分输入项目后,下拉列表数据中做对应的数据筛选
    '如果全部文字都删除了,则清空下拉列表数据
        
    Dim str家庭地址 As String, i As Long
    Dim lng位置 As Long
    
    If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
        If mrs家庭地址 Is Nothing Then Exit Sub
        If mTy_Para.byt家庭地址联想 = 0 Then Exit Sub
        
        str家庭地址 = cbo家庭地址.Text                      '此时,如果选择了部分文字,则选择的文字已经被删除
        lng位置 = cbo家庭地址.SelStart
        
        If mrs家庭地址.State = 1 And Len(str家庭地址) > 1 Then
            
            If InStr(1, "ABCDEFGHIJKLMNOPQRSTUVWXYZ", UCase(Left(str家庭地址, 1))) > 0 Then
                mrs家庭地址.Filter = "简码 like '" & gstrLike & UCase(str家庭地址) & "*'"
            Else
                mrs家庭地址.Filter = "名称 Like '" & gstrLike & str家庭地址 & "*'"
            End If
            
            If Not mrs家庭地址.EOF Then
                
                If mrs家庭地址.RecordCount <> cbo家庭地址.ListCount Then
                    Call SendMessage(cbo家庭地址.Hwnd, CB_RESETCONTENT, 0, 0)
                    mrs家庭地址.Sort = "次数 Desc,名称"
                    For i = 1 To mrs家庭地址.RecordCount
                        AddComboItem cbo家庭地址.Hwnd, CB_ADDSTRING, 0, mrs家庭地址!名称
                        mrs家庭地址.MoveNext
                    Next
                    If SendMessage(cbo家庭地址.Hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 Then Call SendMessage(cbo家庭地址.Hwnd, CB_SHOWDROPDOWN, True, 0&)
                                        
                    cbo家庭地址.Text = str家庭地址
                    cbo家庭地址.SelStart = lng位置
                End If
            Else
                Call SendMessage(cbo家庭地址.Hwnd, CB_SHOWDROPDOWN, False, 0&)
            End If
        ElseIf str家庭地址 = "" Then
            cbo家庭地址.Clear
            Call SendMessage(cbo家庭地址.Hwnd, CB_SHOWDROPDOWN, False, 0&)
        End If
    End If
End Sub

Private Sub cbo家庭地址_KeyPress(KeyAscii As Integer)
    Dim i As Long
    Dim str简码 As String
    Dim str家庭地址 As String
    Dim lng中间输入点 As Long
    Dim strTemp As String
    
    If mrs家庭地址 Is Nothing Then Exit Sub
    
    If mTy_Para.byt家庭地址联想 = 0 Then
        If KeyAscii = 13 Then Call zlCommFun.PressKey(vbKeyTab)
        Exit Sub
    End If
    '用本地缓存匹配输入
    If KeyAscii <> 13 And KeyAscii <> vbKeyF4 And KeyAscii <> vbKeyEscape And _
        KeyAscii <> vbKeyBack And KeyAscii <> 26 And KeyAscii <> 3 And KeyAscii <> 22 Then   '26表示ctrl+z,3-ctrl+c,22-ctrl+v
            
        If mrs家庭地址.State = 0 Or cbo家庭地址.Text = "" Then  '输第一个字时不匹配
            Exit Sub
        End If
       
        '选中中间部分文本再输入的情况
        If cbo家庭地址.SelText <> "" And (cbo家庭地址.SelStart + cbo家庭地址.SelLength) <> Len(cbo家庭地址.Text) Then
            lng中间输入点 = cbo家庭地址.SelStart + 1
            cbo家庭地址.Text = Mid(cbo家庭地址.Text, 1, cbo家庭地址.SelStart) & Chr(KeyAscii) & Mid(cbo家庭地址.Text, cbo家庭地址.SelStart + cbo家庭地址.SelLength + 1)
            cbo家庭地址.SelText = ""
            str家庭地址 = cbo家庭地址.Text
        Else
            '输入点在尾部,或在中间时,后面的已选中
            If cbo家庭地址.SelStart = Len(cbo家庭地址.Text) Or (cbo家庭地址.SelStart + cbo家庭地址.SelLength) = Len(cbo家庭地址.Text) Then
                str家庭地址 = Mid(cbo家庭地址.Text, 1, cbo家庭地址.SelStart) & Chr(KeyAscii)
            Else
                str家庭地址 = Mid(cbo家庭地址.Text, 1, cbo家庭地址.SelStart) & Chr(KeyAscii) & Mid(cbo家庭地址.Text, cbo家庭地址.SelStart + 1)
                lng中间输入点 = cbo家庭地址.SelStart + 1
            End If
        End If
         
        
        If InStr(1, "ABCDEFGHIJKLMNOPQRSTUVWXYZ", UCase(Left(str家庭地址, 1))) > 0 Then
            mrs家庭地址.Filter = "简码 like '" & gstrLike & UCase(str家庭地址) & "*'"
        Else
            mrs家庭地址.Filter = "名称 Like '" & gstrLike & str家庭地址 & "*'"
        End If
        
        If Not mrs家庭地址.EOF Then
            If mrs家庭地址.RecordCount <> cbo家庭地址.ListCount Then
                Call SendMessage(cbo家庭地址.Hwnd, CB_RESETCONTENT, 0, 0)
                mrs家庭地址.Sort = "次数 Desc,名称"
                For i = 1 To mrs家庭地址.RecordCount
                    AddComboItem cbo家庭地址.Hwnd, CB_ADDSTRING, 0, mrs家庭地址!名称
                    mrs家庭地址.MoveNext
                Next
                If SendMessage(cbo家庭地址.Hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 Then Call SendMessage(cbo家庭地址.Hwnd, CB_SHOWDROPDOWN, True, 0&)
            End If
            
            i = KeyAscii    '用来后面判断是否是按退格删除键
            KeyAscii = 0
            cbo家庭地址.Text = str家庭地址
            cbo家庭地址.SelStart = Len(cbo家庭地址.Text)

            mrs家庭地址.MoveFirst   '如果不是输入的简码,相同则取下一个更多的
            If mrs家庭地址!名称 = str家庭地址 And i <> vbKeyBack Then
                mrs家庭地址.MoveNext
            End If
            If Not mrs家庭地址.EOF Then
                If InStr(1, mrs家庭地址!名称, str家庭地址) > 0 Or mrs家庭地址!简码 = UCase(str家庭地址) Then    '输入内容属于已有内容的一部分,则选中缓存多余文字
                    i = Len(cbo家庭地址.Text)
                    strTemp = cbo家庭地址.Text
                    cbo家庭地址.Text = mrs家庭地址!名称
                    If InStr(1, mrs家庭地址!名称, str家庭地址) > 0 Then '问题:31570
                        i = InStr(1, cbo家庭地址.Text, strTemp) + Len(strTemp) - 1
                    End If
                    cbo家庭地址.SelStart = i
                    cbo家庭地址.SelLength = Len(cbo家庭地址.Text) - cbo家庭地址.SelStart
                    If mrs家庭地址.RecordCount = 1 Then Exit Sub
                End If
            End If
            
        '没有找到匹配的缓存数据时,需清除下拉列表数据
        Else
            Call SendMessage(cbo家庭地址.Hwnd, CB_RESETCONTENT, 0, 0)
            If SendMessage(cbo家庭地址.Hwnd, CB_GETDROPPEDSTATE, 0, 0) = 1 Then Call SendMessage(cbo家庭地址.Hwnd, CB_SHOWDROPDOWN, False, 0&)
            KeyAscii = 0
            cbo家庭地址.Text = str家庭地址
            cbo家庭地址.SelStart = Len(cbo家庭地址.Text)
        End If
        
        If lng中间输入点 > 0 Then cbo家庭地址.SelStart = lng中间输入点: cbo家庭地址.SelText = ""
        
    ElseIf KeyAscii = 13 Then
        'a.在没有选中任何文字,且输入内容为空,光标为于末端时,确认输入,并保存信息到本地缓存
        Call SendMessage(cbo家庭地址.Hwnd, CB_SHOWDROPDOWN, False, 0&)
        
        If cbo家庭地址.Text = "" Then
            If gbln家庭地址 And txtPatient.Text <> "" Then
                Exit Sub
            Else
                Call zlCommFun.PressKey(vbKeyTab): Exit Sub
            End If
        End If
        
        '下拉列表弹出时按回车,则定位到末尾
        If cbo家庭地址.SelText = cbo家庭地址.Text Then
            cbo家庭地址.SelStart = Len(cbo家庭地址.Text):
            Exit Sub
       End If
        If mrs家庭地址.State = 0 Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
        If zlCommFun.ActualLen(cbo家庭地址.Text) > 100 Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
        'a.非下拉状态下按回车,没有选择文本
        If cbo家庭地址.SelText = "" Then
            str家庭地址 = cbo家庭地址.Text
            mrs家庭地址.Filter = "名称='" & str家庭地址 & "'"
            If mrs家庭地址.EOF Then
                str简码 = Mid(zlCommFun.zlGetSymbol(str家庭地址), 1, 10)
                If str简码 <> UCase(str家庭地址) Then
                    With mrs家庭地址
                        .AddNew
                        !类别 = "用户"
                        !名称 = str家庭地址
                        !简码 = str简码
                        !次数 = 1
                        .Update                 '在窗体Unload中save
                    End With
                End If
            Else
                mrs家庭地址!次数 = mrs家庭地址!次数 + 1
                mrs家庭地址.Update
                
                If zlCommFun.IsCharAlpha(str家庭地址) Then
                    If mrs家庭地址.RecordCount = 1 Then
                        cbo家庭地址.Text = mrs家庭地址!名称
                    Else
                        Call SendMessage(cbo家庭地址.Hwnd, CB_SHOWDROPDOWN, True, 0&)
                        Exit Sub
                    End If
                End If
            End If
            
            Call zlCommFun.PressKey(vbKeyTab)
        Else
                Call zlCommFun.PressKey(vbKeyTab)
        End If
    End If
End Sub

Private Function CheckMCOutMode(ByVal strMCCode As String) As Boolean
    Dim rsTmp As ADODB.Recordset, strSQL As String
 
    strSQL = "Select 1 From 保险类别 Where 外挂=1 And 序号=[1]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strMCCode)

    CheckMCOutMode = rsTmp.RecordCount > 0
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Sub Init结算方式(ByVal str性质 As String, Optional ByVal objCards As Cards)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化结算方式
    '入参:str性质-结算方式的性质,多个用逗号分离
    '                   1-现金结算方式,2-其他非医保结算,
    '                   3-医保个人帐户,4-医保各类统筹,5-代收款项,6-费用折扣,
    '                   7-一卡通结算,8-结算卡结算)
    '出参:objCards-将相关的结算方式返回给卡对象
    '编制:刘兴洪
    '日期:2013-10-24 10:41:51
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, i As Long, objCard As Card
    Dim rsTmp As ADODB.Recordset
    If str性质 = "" Then
        str性质 = ",1,2,3,4,5,6,7,8,"
    Else
        str性质 = "," & str性质 & ","
    End If
    
    strSQL = _
        "Select B.编码,B.名称,Nvl(B.性质,1) as 性质,Nvl(A.缺省标志,0) as 缺省" & _
        " From 结算方式应用 A,结算方式 B" & _
        " Where A.应用场合=[1] And B.名称=A.结算方式 And Instr([2] ,','||B.性质||',')>0" & _
        " Order by B.编码"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, "挂号", str性质)
    
    cbo结算方式.Clear
    Do While Not rsTmp.EOF
        If Not objCards Is Nothing Then
            Set objCard = New Card
            With objCard
                .接口序号 = 0
                .名称 = Nvl(rsTmp!名称)
                .结算方式 = Nvl(rsTmp!名称)
                .接口编码 = Val(Nvl(rsTmp!性质))
                .启用 = False
            End With
            objCards.Add objCard
        End If
        cbo结算方式.AddItem rsTmp!编码 & "-" & rsTmp!名称
        If rsTmp!名称 = gstr结算方式 Then
            For i = 0 To cbo结算方式.ListCount - 1
                cbo结算方式.ItemData(i) = 0
            Next
            cbo结算方式.ItemData(cbo结算方式.NewIndex) = 1
            cbo结算方式.ListIndex = cbo结算方式.NewIndex
        End If
        
        If rsTmp!缺省 = 1 Then
            If cbo结算方式.ListIndex = -1 Then
                cbo结算方式.ItemData(cbo结算方式.NewIndex) = 1
                cbo结算方式.ListIndex = cbo结算方式.NewIndex
            End If
        End If
        rsTmp.MoveNext
    Loop
    
    If cbo结算方式.ListCount > 0 And cbo结算方式.ListIndex = -1 Then
        cbo结算方式.ListIndex = 0
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub InitFace()
    '功能:初始化界面控件
    Dim i As Long, Control As Object
    
    '68991
    mRegistFeeMode = EM_RG_现收
    mPatiChargeMode = EM_先结算后诊疗
    
    lblPrompt.Caption = ""
    vsfPay.Height = 2220
    Call ClearMoney
    
    
    If mTy_Para.bln点击列头排序 Then
       vsfPlan.ExplorerBar = flexExSortShow
    Else
       vsfPlan.ExplorerBar = flexExNone
    End If
    If mbytInState = 0 Then
        Call InitInputMaxLen
        If mbytMode = 0 And Not mblnStation Then
            chkShowAll.Visible = True
        End If
        
        If InStr(mstrPrivs, ";重打票据;") = 0 Then
            chkPrint.Visible = False
        End If
        If InStr(";" & mstrPrepayPrivs & ";", ";门诊预交;") = 0 Then
            mcbrToolBar.Controls.Find(xtpControlButton, 3816).Visible = False
            mcbrToolBar.Controls.Find(xtpControlButton, 3816).Enabled = False
        End If
        '权限修改 问题：37798 作者：冉勇
        If InStr(mstrPrivs, ";预约挂号;") = 0 Then chkBooking.Visible = False
        
        lblFree.Left = lblCancel.Left: lblFree.Height = lblCancel.Height
        lblFree.Visible = False
        
        txtFact.Locked = Not (InStr(1, mstrPrivs, ";修改票据号;") > 0) And gblnBill挂号  '刘兴洪:20000,增加修改票据号权限
        timPlan.Enabled = glngInterval > 0 And Not mblnStation And (mbytMode = 0 Or mbytMode = 1)
        If timPlan.Enabled Then mDatLastRefresh = Now
    
        Call SetPatiInfoEnabled(False, mrsInfo Is Nothing)  '问题号:58843
        
        cbo医生.Enabled = False
        cbo费别.Enabled = (gbln费别 Or mblnStation) And mbytMode <> 2
        cbo结算方式.Enabled = gbln结算方式 And mbytMode <> 1
        txt家庭电话.Enabled = gbln电话
        lblIDCard.Visible = True
        If IDKind证件.IDKind = IDKind证件.GetKindIndex("身份证号") Then
            txtIDCard.Visible = True: txt证件.Visible = False
        Else
            txtIDCard.Visible = False: txt证件.Visible = True
        End If
        Call SetPicTimeObjectVisible
        
          If mbytMode = 1 Then
            '预约挂号
            chkPrint.Visible = False: chkCancel.Visible = False: chkBooking.Visible = False
            cboNO.Width = cbo医生.Width
            cmdPatiPic.Left = chkBooking.Left
            txtSN.Width = txtSN.Width + cmdPatiPic.Width + 60
            '问题:26964
            chkShowAll.Visible = Not mblnStation: mblnUnChkClick = True
            If Val(zlDatabase.GetPara("预约显示所有号别", glngSys, mlngModul, 1, Array(chkShowAll), InStr(mstrPrivs, ";参数设置;") > 0)) = 1 Then
                chkShowAll.Value = 1
            Else
                chkShowAll.Value = 0
            End If
            mblnUnChkClick = False

            picBookingDate.Visible = True
            lbl摘要.Visible = True: cbo备注.Visible = True
            lbl预约方式.Visible = True: cbo预约方式.Visible = True
            '-----------------------------------------------------------------------------------------
            vsfPay.Visible = False
            lbl应缴.Visible = False
            lblIDCard.Visible = True
            If IDKind证件.IDKind = IDKind证件.GetKindIndex("身份证号") Then
                txtIDCard.Visible = True: txt证件.Visible = False
            Else
                txtIDCard.Visible = False: txt证件.Visible = True
            End If
            txt家庭电话.Visible = True: lbl家庭电话.Visible = True
            cmdCard.Visible = False: cmdYb.Visible = False
            '-----------------------------------------------------------------------------------------
            Call SetUndisplayBalance
        ElseIf mbytMode = 2 Then
            '接收预约
            '隐藏号别安排部份(但要读并填写数据)
            vsfPlan.Visible = False: vsfList.Visible = False
            picSplit.Visible = False
            cmdCard.Visible = InStr(1, mstrPrivs, ";绑定卡号;") > 0   '绑定卡号:31182
            cmdYb.Visible = True   '预约接收时,可以刷医保 '问题:31182
            
            lbl摘要.Visible = True: cbo备注.Visible = True
            cbo备注.Enabled = False: cbo预约方式.Enabled = False
            lbl预约方式.Visible = True: cbo预约方式.Visible = True
'            mcbrToolBar.Visible = False
            dkpMain.DestroyPane dkpMain.Panes(1)
            Call SetReceiveState(True)
            Me.Width = glngMinW: Me.Height = glngMinH
            Me.WindowState = 0
        Else
            '正常挂号
            If InStr(mstrPrivs, ";退号;") = 0 Then
                chkCancel.Visible = False
                lblNO.Left = lblNO.Left + chkCancel.Width
                cboNO.Left = cboNO.Left + chkCancel.Width
            End If
            cmdYb.Visible = True
            picBookingDate.Visible = False
        End If
        
        '初始化序号状态表格
        vsfList.Cols = SNCOLS
        For i = 0 To vsfList.Cols - 1
            vsfList.ColWidth(i) = 570
            vsfList.ColAlignment(i) = 4
        Next
        vsfList.RowHeightMin = 500
        
        '取安排表
        Call SetPlanGrid
    
    Else
        If mbytMode = 1 Then '查看预约单时无结算相关信息
            lbl摘要.Visible = True: cbo备注.Visible = True
            lbl预约方式.Visible = True: cbo预约方式.Visible = True
            Call SetUndisplayBalance
            lblIDCard.Visible = True:  IDKind证件.Visible = True
            If IDKind证件.IDKind = IDKind证件.GetKindIndex("身份证号") Then
                txtIDCard.Visible = True: txt证件.Visible = False
            Else
                txtIDCard.Visible = False: txt证件.Visible = True
            End If
            txt家庭电话.Visible = True: lbl家庭电话.Visible = True
            cmdCard.Visible = False: cmdYb.Visible = False
            If mbytInState = 1 And (mbytMode = 1 Or mbytMode = 3) Then
                lbl发生时间.Visible = True: txt发生时间.Visible = True
            End If
            vsfPay.Visible = False
            stbThis.Visible = False
        Else
            lbl发生时间.Visible = True: txt发生时间.Visible = True
            vsfPay.Height = 1500
        End If
        If mbytMode = 2 Then
            vsfPay.Visible = False
        End If
'        Frame3.Visible = False
'        mcbrToolBar.Visible = False
        dkpMain.DestroyPane dkpMain.Panes(1)
        cmdYb.Visible = False
        stbThis.Visible = False
        
        Call SetPatiEnable(False): Call SetCodeEnable(False)
        cboNO.Locked = True
        chkBooking.Enabled = False
        chk病历费.Enabled = False
        cbo预约方式.Enabled = False
        cbo备注.Enabled = False
        txt发生时间.Enabled = False
        txtFact.Enabled = False
        
        
'        picInfo.Enabled = False
'
'        Set picBal.Container = Me
'        picBal.Top = picBal.Top + 450
'        Set vsfPay.Container = Me
'        vsfPay.Top = vsfPay.Top + 450
'        Set vsfMoney.Container = Me
'        vsfMoney.Top = vsfMoney.Top + 520
        vsfPlan.Visible = False: vsfList.Visible = False
        picSplit.Visible = False
        lblCancel.Visible = mblnViewCancel
        chkCancel.Visible = False: chkPrint.Visible = False: chkBooking.Visible = False
        cmdLookup.Visible = False: cmdMore.Visible = False: cmdCard.Visible = False
                
        cmdOK.Visible = False
        lbl缴款.Visible = False: txt缴款.Visible = False
        lbl找补.Visible = False: txt找补.Visible = False
        txt本次应缴.Visible = False
        lbl应缴.Visible = False
        If mbytMode <> 0 Then
            lblSum.Visible = False: txt合计.Visible = False
            picTotal.Visible = True
        End If
        Call SetUndisplayBalance
        
        If Not (Me.mbytInState = 1 And (mbytMode = 3 Or mbytMode = 4)) Then
            cmdCancel.Caption = "退出(&X)"
            
        Else
            If mbytMode = 3 Then
                cmdOK.Visible = True
                vsfPay.Visible = False
            End If
        End If
        
        If mbytMode = 4 Then
            '设置退号时 , 相关控件的属性
'            chk病历费.Enabled = True
'            Set chk病历费.Container = Me
'            chk病历费.Top = chk病历费.Top + 480
'            chk病历费.Caption = "退病历费"
'            chkExtra.Enabled = True
'            Set chkExtra.Container = Me
'            chkExtra.Top = chk病历费.Top
'            Set cbo备注.Container = Me
'            cbo备注.Top = chk病历费.Top + 450
            lbl预约方式.Visible = False
            cbo预约方式.Visible = False
            cbo费别.Enabled = False
            cbo结算方式.Enabled = False
            cbo结算方式.Visible = False
            vsfMoney.Enabled = False
            cmdCancel.Visible = True
            cmdOK.Visible = True
            cmdOK.Top = 900
            cmdCancel.Top = 900
        Else
            cbo结算方式.Visible = False
            If lbl应缴.Visible Then
                cmdCancel.Top = lbl应缴.Top + lbl应缴.Height + 180
            Else
                cmdCancel.Top = lblSum.Top + lblSum.Height + 180
            End If
            If mbytMode <> 3 Then
                cmdCancel.Left = (picBal.ScaleWidth - cmdCancel.Width) / 2
            Else
                cmdOK.Top = cmdCancel.Top
            End If
        End If
        
        Me.Width = glngMinW: Me.Height = glngMinH
        
        Me.WindowState = 0
        If chkCancel.Value = 1 Or mbytMode = 4 Then
            chkExtra.Caption = "退附加费"
        Else
            chkExtra.Caption = "附加费"
        End If
    End If
      
    Call Set备注Enabled
    lbl险类.Visible = False
    '74430,冉俊明,2014-7-8,挂号界面显示病人照片的浮动窗体
    picPatiPicBack.Left = Me.ScaleWidth - picPatiPicBack.Width
    picPatiPicBack.Top = 0
    picPatiPicBack.Visible = False: cmdPatiPic.Enabled = False
    
    If mbytMode <> 0 And mbytMode <> 1 And mbytMode <> 2 Then cmdPatiPic.Visible = False
'    If mbytMode = 1 Or mbytMode = 2 Then cmdPatiPic.Left = picCode.Width - 300
 
    If mblnStructAdress Then
        padd家庭地址.Visible = True: padd户口地址.Visible = True
        padd家庭地址.ShowTown = mblnShowTown: padd户口地址.ShowTown = mblnShowTown
        cbo家庭地址.Visible = False: padd家庭地址.MaxLength = glngMax家庭地址
        
        padd家庭地址.Top = cbo家庭地址.Top: padd家庭地址.Left = cbo家庭地址.Left
        lbl家庭地址.Top = padd家庭地址.Top
        
        cbo户口地址.Visible = False: padd户口地址.MaxLength = glngMax户口地址
        padd户口地址.Top = padd家庭地址.Top + padd家庭地址.Height + 20: padd户口地址.Left = cbo户口地址.Left
        lbl户口地址.Top = padd户口地址.Top
        picDetailFee.Top = padd户口地址.Top + padd户口地址.Height + 50
    Else
        lbl家庭地址.Top = cbo家庭地址.Top + (cbo家庭地址.Height - lbl家庭地址.Height) \ 2
        lbl户口地址.Top = cbo户口地址.Top + (cbo户口地址.Height - lbl户口地址.Height) \ 2
        picDetailFee.Top = cbo户口地址.Top + cbo户口地址.Height + 50
    End If
    
End Sub

Private Sub Set备注Enabled()
    '--------------------------
    '备注控件的位置以及可用性的调整
    '挂号,退号时 需要调动大小以及位置
    '--------------------------
   Dim Control As Object
   Me.cbo备注.Visible = mbytInState <= 0
   Me.lbl摘要.Visible = True
   If mbytInState <= 0 Or (mbytInState = 1 And (mbytMode = 3 Or mbytMode = 4)) Then
        '执行 或者退预约时
        Me.cbo备注.Visible = True
        Me.cbo预约方式.Enabled = IIf(mbytInState = 1 And mbytMode = 3 Or mbytMode = 4, False, True)
        Me.cbo备注.Enabled = IIf(mbytInState = 1, False, True)
        Me.cbo备注.Visible = True
   Else
        Me.cbo备注.Visible = True: Me.cbo备注.Enabled = IIf(mbytInState = 1, False, True)
   End If
 
   If (mbytMode = 4 Or mbytMode = 3) And mbytInState = 1 Then
        Me.cmdOK.Visible = True: Me.cmdOK.Enabled = True
        Me.cboNO.Locked = True: Me.cboNO.TabStop = False
        Me.cmdCancel.TabIndex = Me.cmdOK.TabIndex - 1
  End If
End Sub
Private Sub zlInitParaSet()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化参数设置
    '编制:刘兴洪
    '日期:2009-12-25 11:27:09
    '问题:26962
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTmp          As String
    Dim lngTmp          As Long
    Err = 0: On Error GoTo Errhand:
    If mblnStation Then zlDatabase.ClearParaCache    '医生站时 读取参数 不从缓存中读取避免立即修改参数不能生效
    strTmp = zlDatabase.GetPara("预约限制时间", glngSys, mlngModul, "1|60")
    With mTy_Para
        .bln挂号生成队列 = Val(zlDatabase.GetPara("排队叫号模式", glngSys, 1113)) <> 0 And mblnStation = False
        '问题:31182
        .int同科限约数 = Val(zlDatabase.GetPara("病人同科限约N个号", glngSys, mlngModul, 0))
        .int同科限挂数 = Val(Split(zlDatabase.GetPara("病人同科限挂N个号", glngSys, mlngModul, 0) & "|", "|")(0))
        .bln同科限挂急诊 = Split(zlDatabase.GetPara("病人同科限挂N个号", glngSys, mlngModul, 0) & "|", "|")(1) = "1"
        .int病人挂号科室数 = Val(zlDatabase.GetPara("病人挂号科室限制", glngSys, mlngModul, 0))
        .int病人预约科室数 = Val(zlDatabase.GetPara("病人预约科室数", glngSys, mlngModul, 0))
        .lng预约有效时间 = Val(zlDatabase.GetPara("预约有效时间", glngSys, mlngModul, 0))
        .int预约失效次数 = Val(zlDatabase.GetPara("预约失约次数", glngSys, mlngModul, 0))
        .bln预约接收确定挂号费 = zlDatabase.GetPara("预约接收确定挂号费", glngSys, mlngModul, 0) = "1"
        .bln允许住院病人挂号 = zlDatabase.GetPara("允许住院病人挂号", glngSys, mlngModul, 0) = "1"
        .bln预约不产生门诊号 = Val(zlDatabase.GetPara("预约不生成门诊号", glngSys, mlngModul, 0)) = 1   '36028
        .bln点击列头排序 = Val(zlDatabase.GetPara("允许列头排序", glngSys, mlngModul, 0)) = 1   '43847
        .bln随机序号选择 = Val(zlDatabase.GetPara("随机序号选择", glngSys, mlngModul, 0)) = 1   '43847
        .bln失约用于挂号 = Val(zlDatabase.GetPara("失约用于挂号", glngSys, mlngModul, 0)) = 1
        .bln退号审核 = Val(zlDatabase.GetPara("退号审核", glngSys, mlngModul, 0)) = 1
        .lngN天取消预约 = Val(zlDatabase.GetPara("N天内不能取消预约号", glngSys, mlngModul, 0))
        .lng预约限制时间 = Val(Split(strTmp, "|")(1))
        .lng预约缺省天数 = Val(Split(strTmp, "|")(0))
          '参数为门诊医生工作站的参数设置中设置
        .bln挂号必须刷卡 = Val(zlDatabase.GetPara("挂号必须刷卡", glngSys, 1260, 0)) = 1     '38603
        .byt家庭地址联想 = Val(Nvl(zlDatabase.GetPara("家庭地址输入方式", glngSys, mlngModul, 1)))
        lngTmp = Val(zlDatabase.GetPara("N岁以下必须录入监护人", glngSys, mlngModul, 0))
        .bln监护人录入 = lngTmp > 0
        .lngN岁以下录入监护人 = lngTmp
        .bln严格按时段挂号 = Val(zlDatabase.GetPara("严格按时段挂号", glngSys, mlngModul, 0)) = 1   '62467
        .blnReuseCancelNO = Val(zlDatabase.GetPara("已退序号允许挂号", glngSys, mlngModul, 1)) = 1
        .int专家号挂号限制 = Val(zlDatabase.GetPara("专家号挂号限制", glngSys, , 0))
        .int专家号预约限制 = Val(zlDatabase.GetPara("专家号预约限制", glngSys, , 0))
        .bln禁止输入年龄 = Val(zlDatabase.GetPara("禁止输入年龄", glngSys, mlngModul, 0)) = 1
        .byt缴款方式 = Val(zlDatabase.GetPara("挂号缴款输入控制", glngSys, mlngModul, 0))
        .byt接收模式 = Val(zlDatabase.GetPara("预约接收模式", glngSys, mlngModul, 0))
    End With
    If mTy_Para.lng预约限制时间 <= 0 Then mTy_Para.lng预约限制时间 = 60
    mblnCheckNOValidity = Val(Nvl(zlDatabase.GetPara("门诊号有效性检查", glngSys, mlngModul, 1), 1)) = 1
    mSortType = Val(zlDatabase.GetPara("缺省排序方式", glngSys, mlngModul, 0))
    mblnStructAdress = Val(zlDatabase.GetPara(251, glngSys)) <> 0 '病人地址结构化录入
    mblnShowTown = Val(zlDatabase.GetPara(252, glngSys)) <> 0 '乡镇地址结构化录入
    Exit Sub
Errhand:
    If ErrCenter = 1 Then Resume
End Sub

Private Function zlGet当前星期几(Optional strDate As String = "") As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取当日是星期几
    '编制:刘兴洪
    '日期:2010-02-04 14:42:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset, bln当前日期 As Boolean, strTemp As String
    bln当前日期 = False
    If strDate = "" Then
        bln当前日期 = True
        If mstr当前星期 <> "" Then zlGet当前星期几 = mstr当前星期: Exit Function
        strSQL = "Select Decode(To_Char(Sysdate,'D'),'1','日','2','一','3','二','4','三','5','四','6','五','7','六',NULL) as 星期  From dual"
        strDate = "1990-01-01"
    Else
        strSQL = "Select Decode(To_Char([1],'D'),'1','日','2','一','3','二','4','三','5','四','6','五','7','六','') As 星期 From dual"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CDate(strDate))
    If rsTemp.EOF = True Then
        Exit Function
    End If
    strTemp = Nvl(rsTemp!星期)
    If bln当前日期 Then mstr当前星期 = strTemp
    zlGet当前星期几 = strTemp
End Function

Private Sub InitData()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long, strTemp As String
    Dim Curdate As Date, arrTmp As Variant
    
    '初始基本数据
     On Error GoTo errH
    
    If mbytInState = 0 Then
        Call GetRegInFor(g私有模块, Me.Name, "idkind", strTemp)
        mintIDKind = Val(strTemp)
    End If
    If mintIDKind > 0 And mintIDKind <= IDKind.ListCount Then IDKind.IDKind = mintIDKind
    
    cbo年龄单位.AddItem "岁"
    cbo年龄单位.AddItem "月"
    cbo年龄单位.AddItem "天"
    cbo年龄单位.ListIndex = 0
    
    mblnOlnyBJYB = False: mlngOutModeMC = 0
    If mbytMode = 0 And Not mblnStation Then '预约和接收不支持,门诊医生站不支持
        arrTmp = Split(GetSetting("ZLSOFT", "公共全局", "本地支持的医保", ""), ",")
        strTemp = ""
        For i = 0 To UBound(arrTmp)
            If IsNumeric(arrTmp(i)) Then
                strTemp = strTemp & "," & Val(arrTmp(i))
                If CheckMCOutMode(arrTmp(i)) Then mlngOutModeMC = Val(arrTmp(i)): Exit For
            End If
        Next
        If strTemp <> "" Then strTemp = Mid(strTemp, 2)
        mblnOlnyBJYB = strTemp = "920"  '见问题:问题:26982
    End If
    
      '加载取消预约挂号所需的 常用取消原因
     cbo备注.Clear
    
    'txtIDCard.Width = cbo家庭地址.Width '31182
    mobjfrmPatiInfo.mlngOutModeMC = mlngOutModeMC
    If mlngOutModeMC = 0 Then
        cbo医疗类别.Enabled = False
'        If mbytMode = 1 Or mbytMode = 4 Then
'            cbo家庭地址.Width = txt家庭电话.Width
'        Else
'            cbo家庭地址.Width = (cbo医疗类别.Left + cbo医疗类别.Width - cbo家庭地址.Left)
'        End If
        'txtIDCard.Width = cbo家庭地址.Width '31182
    Else
        cbo医疗类别.Enabled = True
        strSQL = _
            "Select 编码,名称,Nvl(缺省标志,0) as 缺省 From 医疗类别 Order by 编码"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        cbo医疗类别.AddItem ""
        For i = 1 To rsTmp.RecordCount
            cbo医疗类别.AddItem rsTmp!编码 & "-" & rsTmp!名称
            If rsTmp!缺省 = 1 Then
                cbo医疗类别.ItemData(cbo医疗类别.NewIndex) = 1
            End If
            rsTmp.MoveNext
        Next
        cbo医疗类别.ListIndex = 0
    End If
    
    '问题:26955
    If mbytInState = 0 Then
        zlComboxLoadFromSQL "Select 编码,名称,简码,缺省标志 From 预约方式 ", cbo预约方式
        strTemp = zlDatabase.GetPara("缺省预约方式", glngSys, IIf(mblnStation, 1260, mlngModul), "")
        '问题号:112838,焦博,2017/09/05,基础字典表中未设置任何预约方式时会报错
        If cbo预约方式.ListCount <> 0 Then
            For i = 0 To cbo预约方式.ListCount - 1
                If Mid(cbo预约方式.List(i), InStr(cbo预约方式.List(i), ".") + 1) = strTemp Then
                    cbo预约方式.ListIndex = i
                End If
            Next i
            If cbo预约方式.ListIndex < 0 Then cbo预约方式.ListIndex = 0
        End If
    End If
    
    If Not mblnStation Then
        strSQL = "Select Count(1) As 原因 From 常用退号原因"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        mbln退号原因 = Val(Nvl(rsTmp!原因)) <> 0
    End If
    
    If mbytMode = 4 Then Call SetDelMemo("")
    
    If mbytInState = 0 Then
        If mbytMode = 0 Then
            Set mrsOneCard = GetOneCard
            mblnOneCard = mrsOneCard.RecordCount > 0
        End If
        
        '病历费数据:接收时不需要
        mbln病历费 = True
        If mbytMode <> 2 Then
            mbln病历费 = Not zlGetSpecialItemFee("病历费") Is Nothing
            If Not mbln病历费 Then chk病历费.Visible = False
        End If
        
        If mbytMode = 0 Or mbytMode = 1 Then chk病历费.Value = IIf(zlDatabase.GetPara("默认购买病历", glngSys, mlngModul, 0) = "1", 1, 0)
        
        '结算方式:预约时不需要
        If mbytMode <> 1 Then
            Call Load支付方式
            If cbo结算方式.ListCount = 0 Then
                '74550,冉俊明,2014-7-2,在病人来院就诊,医生在门诊医生站挂号时能够选择结算方式(包含性质为7的一卡通结算)
                If mblnStation Or mblnStationPrice Then
                    cbo结算方式.Visible = False: txt缴款.Left = txt本次应缴.Left: txt缴款.Width = txt本次应缴.Width '隐藏
                End If
            End If
        End If
            
        '费别:接收时不允许再更改
        If Not Init费别(True, False) Then mblnUnload = True: Exit Sub
        If cbo费别.ListCount = 0 Then
            MsgBox "费别等级未设置，请先到费别管理中设置费别！", vbInformation, gstrSysName
            mblnUnload = True: Exit Sub
        End If
    
        '性别
        strSQL = "Select '性别' as 类别,编码,名称,简码,Nvl(缺省标志,0) as 缺省 From 性别 Union All " & _
                 " Select '医疗付款方式' as 类别,编码,名称,简码,Nvl(缺省标志,0) as 缺省 From 医疗付款方式 " & _
                 " Order by 类别,编码"
        Set rsTmp = New ADODB.Recordset
        Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
        rsTmp.Filter = "类别='性别'"
        
        mblnNotChange = True
        cbo性别.Clear
        Do While Not rsTmp.EOF
            cbo性别.AddItem rsTmp!编码 & "-" & rsTmp!名称
            If rsTmp!名称 = gstr性别 Then
                For i = 0 To cbo性别.ListCount - 1
                    cbo性别.ItemData(i) = 0
                Next
                cbo性别.ItemData(cbo性别.NewIndex) = 1
                cbo性别.ListIndex = cbo性别.NewIndex
            End If
            
            If rsTmp!缺省 = 1 And cbo性别.ListIndex = -1 Then
                cbo性别.ItemData(cbo性别.NewIndex) = 1
                cbo性别.ListIndex = cbo性别.NewIndex
            End If
            
            rsTmp.MoveNext
        Loop
        If gstr性别 = "无" Then cbo性别.ListIndex = -1
        mblnNotChange = False
        
        '医疗付款方式
        rsTmp.Filter = "类别='医疗付款方式'"
        cbo付款方式.Clear
        Do While Not rsTmp.EOF
            cbo付款方式.AddItem rsTmp!编码 & "-" & rsTmp!名称
            If rsTmp!名称 = gstr付款方式 Then
                For i = 0 To cbo付款方式.ListCount - 1
                    cbo付款方式.ItemData(i) = 0
                Next
                cbo付款方式.ItemData(cbo付款方式.NewIndex) = 1
                cbo付款方式.ListIndex = cbo付款方式.NewIndex
            End If
            If rsTmp!缺省 = 1 Then
                If cbo付款方式.ListIndex = -1 Then
                    cbo付款方式.ItemData(cbo付款方式.NewIndex) = 1
                    cbo付款方式.ListIndex = cbo付款方式.NewIndex
                End If
            End If
            rsTmp.MoveNext
        Loop
        If cbo付款方式.ListIndex = -1 And cbo付款方式.ListCount > 0 Then cbo付款方式.ListIndex = 0
        
        If cbo家庭地址.Enabled And Not mblnStructAdress Then
            Call Load家庭地址
        End If
        Set mobjfrmPatiInfo.mrsBaseDict = GetBaseDict   '用于挂号病人窗体的字典初始
        Set mrsDoctor = New ADODB.Recordset
        If Not mblnStation Then Call GetAll医生
         
                
        'A.接收
        If mbytMode = 2 Then
            If ReadBooking(mstrNoIn) = False Then
                mblnUnload = True
                Exit Sub
            Else
                If mrsInfo Is Nothing And mbytMode = 2 Then cbo费别.Enabled = True
            End If
            '预约接收
            If CheckIsPrice Then
                Call SetUndisplayBalance
            Else
                Call SetShowBalance
            End If
            
        'B.挂号或预约
        Else
            '挂号日期,ShowPlans中的vsfplan_EnterCell会用到日期
            Curdate = zlDatabase.Currentdate
            
            If mbytMode = 1 Then
                If Curdate < gdatRegistTime Then
                    dtpAppointmentDate.Value = Format(gdatRegistTime + mTy_Para.lng预约缺省天数, "yyyy-MM-dd " & gstr上班时间)
                    dtpAppointmentDate.MinDate = Format(gdatRegistTime, "yyyy-MM-dd 00:00")
                Else
                    dtpAppointmentDate.Value = Format(Curdate + mTy_Para.lng预约缺省天数, "yyyy-MM-dd " & gstr上班时间)
                    dtpAppointmentDate.MinDate = Format(Curdate, "yyyy-MM-dd 00:00")  '27781:当前加一小时
                End If
            End If
        
            Call ShowPlans
        
            '用于判断的最大号别长度 GetMaxLen
            gint号长 = 5
            If Not mrsPlan Is Nothing Then
                If mrsPlan.State = 1 Then
                    gint号长 = 1
                    mrsPlan.MoveFirst
                    For i = 1 To mrsPlan.RecordCount
                        If Len(mrsPlan!号别) > gint号长 Then gint号长 = Len(mrsPlan!号别)
                        mrsPlan.MoveNext
                    Next
                End If
            Else
                gint号长 = GetMaxLen
            End If
        End If
        '79619:李南春,2014/11/13,显示缺省的挂号摘要
        strSQL = "Select 编码,名称,简码 " & _
                 " From 常用挂号摘要 " & _
                 " Where Nvl(缺省标志,0)=1"
        Set rsTmp = New ADODB.Recordset
        Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
        If Not rsTmp.EOF Then
            cbo备注.Text = rsTmp!名称
        End If
        '刷新票据号
        If mbytMode <> 1 And Not mblnStation And gbytInvoice <> 0 And Not mblnStartFactUseType Then
            If Not RefreshFact Then mblnUnload = True: Exit Sub
        End If
    Else '查看
        Call ReadBill(mstrNoIn)
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub Set结算方式Eanbled()
    '设置结算方式的enabled属性
     If mbytInState = 0 Then    '0-执行,1-查阅
        cbo结算方式.Enabled = gbln结算方式 And mbytMode <> 1
     End If
End Sub
Private Sub SetShowBalance()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示结算控件
    '编制:刘兴洪
    '日期:2013-12-24 15:49:21
    '问题:68991
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnVisible As Boolean
    '74522:李南春,2014-6-27,医生工作站挂号不显示结算方式等信息
    If mbytInState = 1 Or mblnStation Or mbytInState = 0 And mbytMode = 1 Then Exit Sub
    '显示结算方式
    blnVisible = True
    lblFact.Visible = blnVisible: txtFact.Visible = blnVisible
    If blnVisible Then
        cbo结算方式.Visible = True
        txt缴款.Left = cbo结算方式.Left + cbo结算方式.Width + 30
        txt缴款.Width = 1305
        vsfPay.Visible = True
    Else
        cbo结算方式.Visible = False
        txt缴款.Left = txt本次应缴.Left
        txt缴款.Width = txt本次应缴.Width
        vsfPay.Visible = False
    End If
    lbl应缴.Visible = blnVisible: txt本次应缴.Visible = blnVisible
    lbl缴款.Visible = blnVisible: txt缴款.Visible = blnVisible
    lbl找补.Visible = blnVisible: txt找补.Visible = blnVisible
    
    lblSum.Caption = "合计"
    lblTotal.Caption = lblSum.Caption
    lblSum.Visible = blnVisible: txt合计.Visible = blnVisible
    picTotal.Visible = Not blnVisible
End Sub
Private Sub SetUndisplayBalance()
    '设置不显示结算相关信息
    Dim blnVisible As Boolean
    Dim blnPrice As Boolean
    
    If (mbytInState = 0 Or mbytInState = 1 And mbytMode = 4 Or chkCancel.Value = 1) Then
        blnPrice = CheckIsPrice
        If mRegistFeeMode = EM_RG_记帐 Or blnPrice Then
            '68991:挂号费采用记帐方式,不应该显示结算的相关信息
            cbo结算方式.Visible = False
            txt缴款.Left = txt本次应缴.Left
            txt缴款.Width = txt本次应缴.Width
            lbl应缴.Visible = False: txt本次应缴.Visible = False
            lbl缴款.Visible = False: txt缴款.Visible = False
            lbl找补.Visible = False: txt找补.Visible = False
            lblFact.Visible = False: txtFact.Visible = False
            lblSum.Caption = IIf(blnPrice, "划价", "记帐")
            lblTotal.Caption = lblSum.Caption
            vsfPay.Visible = False
            lblSum.Visible = False: txt合计.Visible = False
            picTotal.Visible = True
            Exit Sub
        End If
    End If
    
    If mbytInState = 1 And mbytMode = 0 And mRegistFeeMode = EM_RG_记帐 Then
        cbo结算方式.Visible = False
        txt缴款.Left = txt本次应缴.Left
        txt缴款.Width = txt本次应缴.Width
        lbl应缴.Visible = False: txt本次应缴.Visible = False
        lbl缴款.Visible = False: txt缴款.Visible = False
        lbl找补.Visible = False: txt找补.Visible = False
        lblFact.Visible = False: txtFact.Visible = False
        lblSum.Caption = "记帐"
        lblTotal.Caption = lblSum.Caption
        vsfPay.Visible = False
        lblSum.Visible = False: txt合计.Visible = False
        picTotal.Visible = True
        picTotal.Width = picBal.Left - picTotal.Left - 30
        lbl合计.Left = picTotal.Width - lbl合计.Width - 60
        Exit Sub
    End If
    If mbytInState = 1 And mbytMode = 4 Or chkCancel.Value = 1 Then
        '刘兴洪:退号,只需要显示退号方式
        blnVisible = True
        lblFact.Visible = blnVisible: txtFact.Visible = blnVisible
        If blnVisible Then
            cbo结算方式.Visible = False
            txt缴款.Left = cbo结算方式.Left + cbo结算方式.Width + 30
            txt缴款.Width = 1305
        Else
            cbo结算方式.Visible = False
            txt缴款.Left = txt本次应缴.Left
            txt缴款.Width = txt本次应缴.Width
        End If
        lbl应缴.Visible = blnVisible: txt本次应缴.Visible = blnVisible
        lbl应缴.ForeColor = vbRed: txt本次应缴.ForeColor = vbRed
        lbl缴款.Visible = Not blnVisible: txt缴款.Visible = Not blnVisible
        lbl找补.Visible = Not blnVisible: txt找补.Visible = Not blnVisible
        lbl应缴.Caption = "退款": txt本次应缴.ToolTipText = "本次退款=累计实缴金额-累计退个人帐户-累计退冲预交额"
        lblSum.Visible = blnVisible: txt合计.Visible = blnVisible
        picTotal.Visible = Not blnVisible
    ElseIf mbytInState = 0 Then
        blnVisible = mbytInState = 0 Or mbytInState = 1 And mbytMode <> 0
        If blnVisible Then
            cbo结算方式.Visible = True
            txt缴款.Left = cbo结算方式.Left + cbo结算方式.Width + 30
            txt缴款.Width = 1305
        Else
            cbo结算方式.Visible = False
            txt缴款.Left = txt本次应缴.Left
            txt缴款.Width = txt本次应缴.Width
        End If
        If mbytMode = 1 Then
            cbo结算方式.Visible = False
            txt缴款.Left = txt本次应缴.Left
            txt缴款.Width = txt本次应缴.Width
            lblFact.Visible = False: txtFact.Visible = False
            lbl缴款.Visible = False: txt缴款.Visible = False
            lbl找补.Visible = False: txt找补.Visible = False
            txt本次应缴.Visible = False
            lblSum.Visible = False: txt合计.Visible = False
            picTotal.Visible = True
        Else
            lblFact.Visible = blnVisible: txtFact.Visible = blnVisible
            lbl缴款.Visible = blnVisible: txt缴款.Visible = blnVisible
            lbl找补.Visible = blnVisible: txt找补.Visible = blnVisible
            txt本次应缴.Visible = blnVisible
            lblSum.Visible = blnVisible: txt合计.Visible = blnVisible
            picTotal.Visible = Not blnVisible
        End If
        lbl应缴.ForeColor = lbl缴款.ForeColor: txt本次应缴.ForeColor = &H108000
        lbl应缴.Caption = "应缴"
        txt本次应缴.ToolTipText = "本次应缴合计 = 累计实缴金额 - 累计个人帐户支付 - 累计冲预交额"
    ElseIf mblnViewCancel Then
        '显示退的数据
        blnVisible = True
        cbo结算方式.Visible = True
        txt缴款.Left = cbo结算方式.Left + cbo结算方式.Width + 30
        txt缴款.Width = 1590
        lbl应缴.Visible = False: txt本次应缴.Visible = False
        lbl应缴.ForeColor = vbRed: txt本次应缴.ForeColor = vbRed
        lbl缴款.Visible = Not blnVisible: txt缴款.Visible = Not blnVisible
        lbl找补.Visible = Not blnVisible: txt找补.Visible = Not blnVisible
        lbl应缴.Caption = "退款"
        txt本次应缴.ToolTipText = "本次退款=累计实缴金额-累计退个人帐户-累计退冲预交额"
        lblSum.Visible = False: txt合计.Visible = False
        picTotal.Visible = True
    End If
End Sub
 
Private Sub SetPicTimeObjectVisible()
    If mbytMode <> 0 And mbytMode <> 1 Then
        picTime.Visible = False: Exit Sub
    End If
    If mbytMode = 0 Then
        lblRegTotal(0).Visible = True
        lblRegTotal(1).Visible = True
        lbl预约时间.Visible = chkBooking.Value = 1
        dtpAppointmentTime.Visible = chkBooking.Value = 1
    Else
        lblRegTotal(0).Visible = False
        lblRegTotal(1).Visible = False
        lbl预约时间.Visible = True
        dtpAppointmentTime.Visible = True
    End If
End Sub

Private Sub SetPlanGrid()
    Dim i As Integer, strHead As String
    
    '133363:李南春,2018/11/1,增加列“剩余可挂合计”，用于显示限号号别的剩余可挂数量
    '每列增加ColData属性，用于控制列是否可见
    '初始安排表
    '列名,对齐,列宽,是否可选(1-固定,-1-不能选,0-可选)
    If mbytMode = 1 Then
        strHead = "RowNum,1,285,-1|IDS,1,0,-1|号类,1,500,1|号别,1,750,1|科室,1,1200,1|项目,1,1500,0|医生,1,700,0|替诊医生,1,1000,0|时段,1,750,0|剩余,1,500,-1|已挂,1,500,-1|限号,1,500,-1|已约,1,500,0|限约,1,500,0" & _
            "|日,4,450,0|一,4,450,0|二,4,450,0|三,4,450,0|四,4,450,0|五,4,450,0|六,4,450,0" & _
            "|病案,4,500,0|分诊,4,500,0|序号控制,4,850,0|出诊日期,1,1100,0|记录ID,1,0,-1|提前时间,1,0,-1|挂号时间,1,0,-1|号源ID,1,0,-1|是否独占,1,0,-1|替诊开始时间,1,0,-1|替诊终止时间,1,0,-1|替诊医生姓名,1,0,-1|替诊医生ID,1,0,-1|分时段,1,0,-1|结束时间,1,0,-1"
    Else
        strHead = "RowNum,1,285,-1|IDS,1,0,-1|号类,1,500,1|号别,1,750,1|科室,1,1200,1|项目,1,1500,0|医生,1,700,0|替诊医生,1,1000,0|时段,1,750,0|剩余,1,500,0|已挂,1,500,0|限号,1,500,0|已约,1,500,0|限约,1,500,0" & _
            "|日,4,450,0|一,4,450,0|二,4,450,0|三,4,450,0|四,4,450,0|五,4,450,0|六,4,450,0" & _
            "|病案,4,500,0|分诊,4,500,0|序号控制,4,850,0|出诊日期,1,1100,0|记录ID,1,0,-1|提前时间,1,0,-1|挂号时间,1,0,-1|号源ID,1,0,-1|是否独占,1,0,-1|替诊开始时间,1,0,-1|替诊终止时间,1,0,-1|替诊医生姓名,1,0,-1|替诊医生ID,1,0,-1|分时段,1,0,-1|结束时间,1,0,-1"
    End If

    With vsfPlan
        .Redraw = flexRDNone
        .Clear: .Rows = 2
        .Cols = UBound(Split(strHead, "|")) + 1
        For i = 0 To .Cols - 1
            .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
            .ColAlignment(i) = Split(Split(strHead, "|")(i), ",")(1)
            .ColData(i) = Val(Split(Split(strHead, "|")(i), ",")(3))
            .ColKey(i) = .TextMatrix(0, i)
            If Not Visible Then .ColWidth(i) = Split(Split(strHead, "|")(i), ",")(2)
            .FixedAlignment(i) = flexAlignCenterCenter
        Next
        .TextMatrix(0, GetCol("RowNum")) = ""
        
        If Not Visible Then Call RestoreFlexState(vsfPlan, App.ProductName & "\" & Me.Name)
        If mbytMode <> 0 Then
            .ColHidden(.ColIndex("剩余")) = True
        End If
        If mbytMode = 1 Then
            .ColHidden(.ColIndex("已挂")) = True: .ColHidden(.ColIndex("限号")) = True
        End If
        .ColHidden(.ColIndex("IDS")) = True
        .ColHidden(.ColIndex("提前时间")) = True: .ColHidden(.ColIndex("挂号时间")) = True
        .ColHidden(.ColIndex("号源ID")) = True: .ColHidden(.ColIndex("是否独占")) = True
        .ColHidden(.ColIndex("替诊开始时间")) = True: .ColHidden(.ColIndex("替诊终止时间")) = True
        .ColHidden(.ColIndex("替诊医生姓名")) = True: .ColHidden(.ColIndex("替诊医生ID")) = True
        .ColHidden(.ColIndex("分时段")) = True
        .RowHeight(0) = 500
        .RowData(0) = 0
        .Redraw = flexRDBuffered
    End With
End Sub

Private Function zlCheck限约或限号数(ByVal lng记录ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查限约数或限号数是否合法
    '入参:str号别-号别
    '出参:
    '返回:合法,返回ture,否则返回False
    '编制:刘兴洪
    '日期:2009-12-30 15:15:57
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, lngTemp As Long, strSQL As String, Curdate As Date
    Dim lng限约数 As Long, lng限号数 As Long, lng已挂数 As Long, lng已约数 As Long, lng剩余预约数 As Long
    Dim lng失约数 As Long
    Dim lng已接收 As Long
    Dim bln分时段 As Boolean
    Dim strMsg As String, int控制方式 As Integer
    Dim lng合作单位数量 As Long
    Dim blnHaveUnitreg As Boolean
    Dim i As Integer, j As Integer
    Err = 0: On Error GoTo Errhand:
    lng限约数 = 0: lng限号数 = 0: lng已挂数 = 0: lng已约数 = 0: lng剩余预约数 = 0
    mbln加号 = False
    If picBookingDate.Visible Then
        Curdate = CDate(Format(dtpAppointmentDate.Value, IIf(bln分时段, "yyyy-MM-dd HH:mm:ss", "yyyy-MM-dd")))
    Else
        Curdate = CDate(Format(zlDatabase.Currentdate, "yyyy-MM-dd"))
    End If
    strSQL = "Select Nvl(限号数, 0) As 限号数, 限约数, Nvl(已挂数, 0) As 已挂数, Nvl(已约数, 0) As 已约数, Nvl(其中已接收, 0) As 已接收" & vbNewLine & _
            "From 临床出诊记录 Where Id = [1]"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng记录ID)
    
    If mbytMode = 0 Then
        lng失约数 = Get失约号(lng记录ID, Curdate)
    End If
    
    If Not rsTmp.EOF Then
        lng限约数 = Val(Nvl(rsTmp!限约数)): lng限号数 = Val(Nvl(rsTmp!限号数))
        lng已挂数 = Val(Nvl(rsTmp!已挂数)): lng已约数 = Val(Nvl(rsTmp!已约数)) - Val(Nvl(rsTmp!已接收))
        lng已接收 = Val(Nvl(rsTmp!已接收))
        If lng已约数 < 0 Then lng已约数 = 0
        lng剩余预约数 = IIf(lng限号数 - lng已挂数 - lng已约数 <= 0, 0, lng限约数 - lng已约数): If lng剩余预约数 < 0 Then lng剩余预约数 = 0
        If lng限约数 = 0 And IsNull(rsTmp!限约数) Then lng限约数 = lng限号数
        lng已约数 = lng已约数 - lng失约数
    End If
    If lng限号数 <= 0 Then
        '不作限制:返回
        zlCheck限约或限号数 = True: Exit Function
    End If
    If (mbytMode = 1 Or chkBooking.Value = 1) And mblnUnitReg And Not mrsUnitReg Is Nothing Then
        mrsUnitReg.Filter = 0
        If mrsUnitReg.RecordCount <> 0 Then
            int控制方式 = Val(Nvl(mrsUnitReg!控制方式))
        End If
        If mViewMode = V_普通号 And mrsUnitReg.RecordCount > 0 Then
           If Val(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("是否独占"))) = 0 Then
               lng合作单位数量 = 0
           Else
                If int控制方式 = 1 Then
                    Do While Not mrsUnitReg.EOF
                        lng合作单位数量 = lng合作单位数量 + Int(Val(Nvl(mrsUnitReg!数量)) * lng限约数 / 100)
                        mrsUnitReg.MoveNext
                    Loop
                Else
                    Do While Not mrsUnitReg.EOF
                        lng合作单位数量 = lng合作单位数量 + Val(Nvl(mrsUnitReg!数量))
                        mrsUnitReg.MoveNext
                    Loop
                End If
                mrsUnitReg.MoveFirst
           End If
        End If
        If mViewMode = V_普通号分时段 And mrsUnitReg.RecordCount > 0 Then
            If Val(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("是否独占"))) = 0 Then
                lng合作单位数量 = 0
            Else
                If int控制方式 = 1 Then
                    Do While Not mrsUnitReg.EOF
                        lng合作单位数量 = lng合作单位数量 + Int(Val(Nvl(mrsUnitReg!数量)) * lng限约数 / 100)
                        mrsUnitReg.MoveNext
                    Loop
                ElseIf int控制方式 = 2 Then
                    Do While Not mrsUnitReg.EOF
                        lng合作单位数量 = lng合作单位数量 + Val(Nvl(mrsUnitReg!数量))
                        mrsUnitReg.MoveNext
                    Loop
                End If
                mrsUnitReg.MoveFirst
            End If
        End If
        If (mViewMode = v_专家号 Or mViewMode = v_专家号分时段) And mrsUnitReg.RecordCount > 0 Then
            If int控制方式 = 3 Then
                Do While Not mrsUnitReg.EOF
                    lng合作单位数量 = lng合作单位数量 + Val(Nvl(mrsUnitReg!数量))
                    mrsUnitReg.MoveNext
                Loop
                mrsUnitReg.MoveFirst
            Else
                If Val(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("是否独占"))) = 0 Then
                    lng合作单位数量 = 0
                Else
                    If int控制方式 = 1 Then
                        Do While Not mrsUnitReg.EOF
                            lng合作单位数量 = lng合作单位数量 + Int(Val(Nvl(mrsUnitReg!数量)) * lng限约数 / 100)
                            mrsUnitReg.MoveNext
                        Loop
                    ElseIf int控制方式 = 2 Then
                        Do While Not mrsUnitReg.EOF
                            lng合作单位数量 = lng合作单位数量 + Val(Nvl(mrsUnitReg!数量))
                            mrsUnitReg.MoveNext
                        Loop
                    End If
                    mrsUnitReg.MoveFirst
                End If
            End If
        End If
       '排除已经挂出的合作单位号
       strSQL = "Select Count(1) As 已约数 From 病人挂号记录 Where 记录状态 = 1 And 出诊记录ID = [1] And 合作单位 Is Not Null "
       Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng记录ID)
       If Not rsTmp.EOF Then
            lng合作单位数量 = lng合作单位数量 - Val(rsTmp!已约数)
       End If
       If lng合作单位数量 < 0 Then lng合作单位数量 = 0
    End If
    
    '************************************************************************
    '限号数-已约数-已挂数>0  | 限号数>限约数 |如果限约数=0 限约数=限号数
    '达到限号数或者预约时达到限约数
    '   操作员拥有加号权限 提示 让操作员自己选择是否继续挂号或者预约
    '   操作员没有加号权限 则禁止操作员继续挂号或者预约
    '************************************************************************
    
    
    'mbytMode:0-挂号,1-预约,2-接收,预约有两种模式:0-挂号,此时预约要收费,1-预约,不收费
    Select Case mbytMode
    Case 1:  '预约
         If lng限号数 - lng已挂数 - lng已约数 - lng合作单位数量 > 0 Then
            '----------------------------------------------
            '挂号+预约数 没有达到限号数
            '----------------------------------------------
            
             If lng已约数 + lng已接收 + lng合作单位数量 >= lng限约数 Then
                If InStr(mstrPrivs, ";加号;") > 0 Then  '需要提示:
                     If MsgBox("该号别该天已达到限约数" & lng限约数 & IIf(lng合作单位数量 > 0, "(其中包含挂号合作单位分配数量[" & lng合作单位数量 & "])", "") & " ，你是否继续预约?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        If mbytMode <> 2 Then txt号别 = "": If txt号别.Enabled And txt号别.Visible Then txt号别.SetFocus
                        Exit Function
                    End If
                    mbln加号 = True
                Else
                    MsgBox "该号别该天已达到限约数 " & lng限约数 & IIf(lng合作单位数量 > 0, "(其中包含挂号合作单位分配数量[" & lng合作单位数量 & "])", "") & "，不能再预约！", vbInformation, gstrSysName
                    If mbytMode <> 2 Then txt号别 = "": If txt号别.Enabled And txt号别.Visible Then txt号别.SetFocus
                    Exit Function
                End If
            End If
        Else
          '------------------------------------------
           '已挂数+已约数 达到了限号数
           '操作员拥有加号码权限 让操作员选择是否继续
           '------------------------------------------
           If InStr(mstrPrivs, ";加号;") > 0 Then
                                If MsgBox("该号别今天已达到限号数 " & lng限号数 & IIf(lng合作单位数量 > 0, "(其中包含挂号合作单位分配数量[" & lng合作单位数量 & "])", "") & "，你是否继续预约?", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
                    If mbytMode <> 2 Then txt号别 = "": If txt号别.Enabled And txt号别.Visible Then txt号别.SetFocus
                    Exit Function
                End If
                mbln加号 = True
           Else
                                        MsgBox "该号别今天已达到限号数 " & lng限号数 & IIf(lng合作单位数量 > 0, "(其中包含挂号合作单位分配数量[" & lng合作单位数量 & "])", "") & "不能再预约！", vbInformation, gstrSysName
                    If mbytMode <> 2 Then txt号别 = "": If txt号别.Enabled And txt号别.Visible Then txt号别.SetFocus
                    Exit Function
                
           End If
        End If
    Case Else '挂号,接收
        If mbytMode = 0 And chkBooking.Value = 0 Then
            '挂号
            If lng已挂数 + lng已约数 >= lng限号数 Then
                If InStr(mstrPrivs, ";加号;") > 0 Then
                    If MsgBox("该号别今天已达到限号数 " & lng限号数 & "，你是否继续挂号?", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
                         If mbytMode <> 2 Then txt号别 = "": If txt号别.Enabled And txt号别.Visible Then txt号别.SetFocus
                         Exit Function
                    End If
                    If mbytMode = 0 Then
                        With vsfList
                            For i = 0 To .Rows - 1
                                For j = 0 To .Cols - 1
                                    If .Cell(flexcpData, i, j) Like "加*" Then .Select i, j
                                Next j
                            Next i
                        End With
                    End If
                    mbln加号 = True
                Else
                    MsgBox "该号别今天已达到限号数 " & lng限号数 & "不能再挂号！", vbInformation, gstrSysName
                    If mbytMode <> 2 Then txt号别 = "": If txt号别.Enabled And txt号别.Visible Then txt号别.SetFocus
                    Exit Function
                End If
            End If
        Else
            '接收
            If lng已挂数 >= lng限号数 Then
                If InStr(mstrPrivs, ";加号;") > 0 Then
                    If MsgBox("该号别今天已达到限号数 " & lng限号数 & "，你是否继续挂号?", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
                         If mbytMode <> 2 Then txt号别 = "": If txt号别.Enabled And txt号别.Visible Then txt号别.SetFocus
                         Exit Function
                    End If
                    If mbytMode = 0 Then
                        With vsfList
                            For i = 0 To .Rows - 1
                                For j = 0 To .Cols - 1
                                    If .Cell(flexcpData, i, j) Like "加*" Then .Select i, j
                                Next j
                            Next i
                        End With
                    End If
                    mbln加号 = True
                Else
                    MsgBox "该号别今天已达到限号数 " & lng限号数 & "不能再挂号！", vbInformation, gstrSysName
                    If mbytMode <> 2 Then txt号别 = "": If txt号别.Enabled And txt号别.Visible Then txt号别.SetFocus
                    Exit Function
                End If
            End If
        End If
    End Select
    zlCheck限约或限号数 = True
   
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function

Private Function GetHave(ByVal lng记录ID As Long) As String
    '功能:取指定号别限号数及已挂数
    '返回:"限号数;已挂数;剩余预约数"或"限约数;已约数;剩余预约数"
    '刘兴洪 问题:26962 日期:2009-12-25 11:46:30 Modify:剩余预约数
    Dim rsTmp As ADODB.Recordset, lngTemp As Long
    Dim strSQL As String, Curdate As Date
    
    GetHave = "0;0;0"
    If picBookingDate.Visible Then
        Curdate = CDate(Format(dtpAppointmentDate.Value, "yyyy-MM-dd"))
    Else
        Curdate = CDate(Format(zlDatabase.Currentdate, "yyyy-MM-dd"))
    End If
    
    strSQL = "Select Nvl(限号数, 0) As 限号数, Nvl(已挂数, 0) - Nvl(其中已接收, 0) As 已挂数, Nvl(限约数, 0) As 限约数, Nvl(已约数, 0) As 已约数" & vbNewLine & _
            "From 临床出诊记录" & vbNewLine & _
            "Where ID = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng记录ID)
    
    If Not rsTmp.EOF Then
        lngTemp = Val(Nvl(rsTmp!限约数)) - Val(Nvl(rsTmp!已约数))
        If lngTemp < 0 Then lngTemp = 0
        If mbytMode = 1 Then
            GetHave = rsTmp!限约数 & ";" & rsTmp!已约数 & ";" & lngTemp
        Else
            GetHave = rsTmp!限号数 & ";" & rsTmp!已挂数 & ";" & lngTemp
        End If
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function ShowPlans(Optional strSort As String, Optional blnCache As Boolean, Optional ByVal blnAutoUpdate As Boolean = True, Optional ByVal blnShowStop As Boolean = False) As Boolean
'功能：读取当日安排内容
'blnCache:仅当输入号别未达到最大长度时才缓存,主要是考虑限号时刻在变
    Dim strTime As String, strState As String
    Dim strSQL As String, strIF As String
    Dim i As Integer, k As Integer, rsDays As ADODB.Recordset
    Dim DateThis As Date, strZero As String, intGap As Integer, intSurplusTotal As Integer
    Dim str挂号安排 As String, strDays As String
    Dim str挂号安排计划 As String, str号源IDs As String
    Dim str排序         As String, strSpan As String
    Dim dat开始时间 As Date, dat结束时间 As Date, datNow As Date
    Dim str不当班号码 As String, rs不当班 As ADODB.Recordset, str不当班SQL As String
    Dim strStationSql As String, str不当班IF As String
    Dim str号源IDSQL As String, str不当班号码SQL As String
    
    On Error GoTo errH
    If mblnUnload Then Exit Function
    Select Case mSortType
        Case by号别:
                str排序 = "号别,出诊日期 Desc,挂号时间"
        Case by科室:
                str排序 = "科室,项目,已挂,出诊日期 Desc,挂号时间"
        Case by科室and已挂数:
                str排序 = "科室,已挂,出诊日期 Desc,挂号时间"
        Case Else:
             str排序 = "号别,出诊日期 Desc,挂号时间"
    End Select
    
    If strSort = "" Then strSort = IIf(mstrSort = "", str排序, mstrSort)
    If InStr(1, strSort, str排序) = 0 Then strSort = strSort & "," & str排序
    If blnCache Then blnCache = Not mrsPlan Is Nothing
    
    If Not blnCache Then
        datNow = zlDatabase.Currentdate
        If picBookingDate.Visible Or mbytMode = 1 Or (mbytMode = 2 Or mbytMode = 0 And mbytInState = 0 And mstrNoIn <> "") Then
            DateThis = dtpAppointmentDate.Value
        Else
            DateThis = zlDatabase.Currentdate
        End If
        
        strSQL = "Select a.Id, b.号码 As 号别, a.出诊日期, b.号类, b.科室id, c.名称 As 科室, a.项目id, d.名称 As 项目, a.替诊医生id, a.医生id, a.替诊医生姓名, a.医生姓名 As 医生, Nvl(a.已挂数, 0) As 已挂," & vbNewLine & _
                "       Nvl(a.已约数, 0) As 已约, a.限号数 As 限号, a.限约数 As 限约, Nvl(b.是否建病案, 0) As 病案, Nvl(d.项目特性, 0) As 急诊, Decode(a.分诊方式,1,'指定',2,'动态',3,'平均',NULL) As 分诊," & vbNewLine & _
                "       a.是否序号控制 As 序号控制, a.上班时段 As 排班, a.号源id, a.提前挂号时间 As 提前时间, a.开始时间 As 挂号时间, a.终止时间 As 结束时间, Nvl(a.是否独占,0) As 是否独占, a.替诊开始时间 , a.替诊终止时间, a.是否分时段 As 分时段 " & vbNewLine & _
                "From 临床出诊记录 A, 临床出诊号源 B, 部门表 C, 收费项目目录 D, 人员表 E" & vbNewLine & _
                "Where (a.出诊日期 = [6] Or a.出诊日期 = [8]) And a.号源id = b.Id  And Nvl(C.撤档时间,To_Date('3000-01-01','YYYY-MM-DD')) > Sysdate And b.科室id = c.Id And a.项目id = d.Id And Nvl(a.是否锁定, 0) = 0 " & vbNewLine & _
                "       And a.医生id = e.Id(+) And (d.撤档时间 is NULL Or d.撤档时间=To_Date('3000-01-01','YYYY-MM-DD')) And (c.站点='" & gstrNodeNo & "' Or c.站点 is Null) " & _
                "       And Nvl(a.是否发布,0) = 1 "
        
        If mbytMode = 1 Or chkBooking.Value = 1 Then
            If Format(DateThis, "yyyy-mm-dd") = Format(datNow, "yyyy-mm-dd") Then
                strSQL = strSQL & "       And Not Exists (Select 1 From 临床出诊记录 Where Id=a.Id And 终止时间 < [9])"
            Else
                strSQL = strSQL & "       And Not Exists (Select 1 From 临床出诊记录 Where Id=a.Id And 终止时间 < [6])"
            End If
        Else
            strSQL = strSQL & "       And Not Exists (Select 1 From 临床出诊记录 Where Id=a.Id And 终止时间 < [5])"
        End If
        
        If Not blnShowStop And chkShowAll.Value <> 1 Then
            strSQL = strSQL & " And (a.开始时间 < Nvl(a.停诊开始时间,a.终止时间) Or a.终止时间 > Nvl(a.停诊终止时间,a.开始时间) Or Exists (Select 1 From 临床出诊序号控制 C,临床出诊记录 D Where D.ID=A.ID And C.记录ID=D.ID And Nvl(C.是否停诊,0) = 0 And D.是否序号控制 =1 And D.是否分时段 = 1 And C.开始时间 <> C.终止时间)) "
        End If
        
        If chkShowAll.Value <> 1 Then
            strSQL = strSQL & " And [5] Not Between Nvl(a.停诊开始时间,a.终止时间) And Nvl(a.停诊终止时间,a.开始时间) "
        End If
        
        If (mbytMode = 0 Or mbytMode = 1) And mstrNoIn = "" Then
            '挂号和预约时不能显示已停用或删除的号源
            strSQL = strSQL & " And Nvl(b.是否删除, 0) = 0 And (b.撤档时间 >= To_Date('3000-01-01', 'yyyy-mm-dd') Or b.撤档时间 Is Null)"
        End If
        
        If gstr挂号科室ID <> "" Then
            '本地参数确定了的挂号科室
            strIF = " And Instr(','||[4]||',',','||b.科室ID||',')>0"
        End If
        
        '按输入的号别过滤：仅号别输入过程中才过滤,这时的ActiveControl一定是txt号别
        If Trim(txt号别.Text) <> "" And Trim(txt号别.Text) <> "+" And ActiveControl Is txt号别 And mblnReadBooking = False Then
            If IsNumeric(Trim(txt号别.Text)) Then
                strIF = strIF & " And b.号码 Like [2]"
            Else
                strIF = strIF & " And (zlSpellCode(e.姓名) Like [2] or c.简码 Like [2])"
            End If
        End If
        
        strSQL = strSQL & strIF
   
        If chkShowAll.Value = 1 Then
            '所有安排
            If mbytMode = 1 Or chkBooking.Value = 1 Then
                strTime = strTime & " And Nvl(a.预约控制,0) <> 1 "
            End If
        Else
            '该部分语句取现在所对应的时间段
            If mbytMode = 1 Or chkBooking.Value = 1 Then
                strTime = strTime & " And Nvl(a.预约控制,0) <> 1 "
            Else
                strTime = _
                " And [5] Between Nvl(a.提前挂号时间, a.开始时间) And a.终止时间  "
            End If
        End If
               
        If mbytMode <> 1 Then
            If InStr(mstrPrivs, ";挂免费号;") = 0 And mbytMode = 0 Then
                strZero = "" & _
                "   And Exists(Select 1 From 收费价目" & _
                                " Where 收费细目id = d.Id And [5] Between 执行日期 And Nvl(终止日期, To_Date('3000-01-01', 'YYYY-MM-DD'))" & _
                                " Group By 价格等级 Having Nvl(Sum(现价), 0) <> 0)"
            End If
            
            If InStr(mstrPrivs, ";挂收费号;") = 0 And mbytMode = 0 Then
                strZero = strZero & _
                "   And Exists(Select 1 From 收费价目" & _
                                " Where 收费细目id = d.Id And [5] Between 执行日期 And Nvl(终止日期, To_Date('3000-01-01', 'YYYY-MM-DD'))" & _
                                " Group By 价格等级 Having Nvl(Sum(现价), 0) = 0)"
            End If
            
            strSQL = strSQL & strZero
        End If
        
        Dim strWhere As String
        If mbytMode = 1 Or (mbytMode = 0 And chkBooking.Visible And chkBooking.Value = 1) Then
            '预约挂号
            '预约  根据是否采用了分时段
            ' 判断是否限制 了只有在当前时间段 才出来
            If mcustomTime = t_普通 Then
                strSQL = strSQL & strTime
            Else
                strSQL = IIf(chkShowAll.Value = 0, strSQL & strTime, strSQL)
            End If
            strSQL = strSQL & IIf(chkShowAll.Value = 1, "", " And (a.限约数 > 0 Or a.限约数 Is Null)")
            strSQL = strSQL & IIf(chkShowAll.Value = 1, "", " And Nvl(a.预约控制,0) <> 1 ")
            strSQL = strSQL & " And Sysdate + zl_Fun_GetAppointmentDays + Decode(Nvl(B.预约天数," & gint预约天数 & "),0,15,Nvl(B.预约天数," & gint预约天数 & ")" & ") > [5] "
        Else
            '挂号
            strSQL = strSQL & strTime
            If chkShowAll.Value = 1 Then
                '显示不当班号别
                If Trim(txt号别.Text) <> "" And Trim(txt号别.Text) <> "+" And ActiveControl Is txt号别 And mblnReadBooking = False Then
                    If IsNumeric(Trim(txt号别.Text)) Then
                        str不当班IF = " And a.号码 Like [2]"
                    Else
                        str不当班IF = " And (zlSpellCode(a.医生姓名) Like [2] or c.简码 Like [2])"
                    End If
                End If
                
                strSQL = strSQL & vbNewLine & _
                    "Union All" & vbNewLine & _
                    "Select 0 As 记录id, a.号码 As 号别, Null As 出诊日期, a.号类, a.科室id, c.名称 As 科室, a.项目id, d.名称 As 项目, Null As 替诊医生id, a.医生id," & vbNewLine & _
                    "       Null As 替诊医生姓名, a.医生姓名 As 医生, 0 As 已挂, 0 As 已约, Null As 限号, Null As 限约, Nvl(a.是否建病案, 0) As 病案," & vbNewLine & _
                    "       Nvl(d.项目特性, 0) As 急诊, Null As 分诊, 0 As 序号控制, Null As 排班, a.Id As 号源id, Null As 提前时间, Null As 挂号时间, Null As 结束时间," & vbNewLine & _
                    "       Null As 是否独占, Null As 替诊开始时间, Null As 替诊终止时间, 0 As 分时段 " & vbNewLine & _
                    "From 临床出诊号源 A, 部门表 C, 收费项目目录 D" & vbNewLine & _
                    "Where a.科室id = c.Id And a.项目id = d.Id" & str不当班IF & vbNewLine & _
                    "      And Nvl(a.撤档时间, To_Date('3000-01-01', 'YYYY-MM-DD')) > [5]" & vbNewLine & _
                    "      And Nvl(c.撤档时间, To_Date('3000-01-01', 'YYYY-MM-DD')) > [5]" & vbNewLine & _
                    "      And Nvl(d.撤档时间, To_Date('3000-01-01', 'YYYY-MM-DD')) > [5]" & vbNewLine & _
                    "      And Exists(Select 1 From 临床出诊安排 M, 临床出诊表 N Where m.号源id = a.Id And m.出诊id = n.Id And n.发布时间 Is Not Null)" & vbNewLine & _
                    "      And Not Exists(Select 1" & vbNewLine & _
                    "                     From 临床出诊记录" & vbNewLine & _
                    "                     Where 号源id = a.Id And (出诊日期 = [6] Or 出诊日期 = [8])" & vbNewLine & _
                    "                           And [5] Between 开始时间 And 终止时间" & vbNewLine & _
                    "                           And (开始时间 < Nvl(停诊开始时间, 终止时间) Or 终止时间 > Nvl(停诊终止时间, 开始时间))" & vbNewLine & _
                    "                           And Nvl(是否锁定, 0) = 0 And Nvl(是否发布, 0) = 1)" & vbNewLine
                
                If mbytMode <> 1 Then
                    If InStr(mstrPrivs, ";挂免费号;") = 0 Then
                        strZero = "" & _
                        "   And Exists(Select 1 From 收费价目" & _
                                " Where 收费细目id = d.Id And [5] Between 执行日期 And Nvl(终止日期, To_Date('3000-01-01', 'YYYY-MM-DD'))" & _
                                " Group By 价格等级 Having Nvl(Sum(现价), 0) <> 0)"
                    End If
                    
                    If InStr(mstrPrivs, ";挂收费号;") = 0 Then
                        strZero = strZero & _
                        "   And Exists(Select 1 From 收费价目" & _
                                " Where 收费细目id = d.Id And [5] Between 执行日期 And Nvl(终止日期, To_Date('3000-01-01', 'YYYY-MM-DD'))" & _
                                " Group By 价格等级 Having Nvl(Sum(现价), 0) = 0)"
                    End If
                    strSQL = strSQL & strZero
                End If
            End If
        End If
              
        strSQL = strSQL & " Order by " & strSort
        
        Set mrsPlan = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, _
            UserInfo.姓名, Trim(txt号别.Text) & "%", mstrRoom, gstr挂号科室ID, DateThis, CDate(Format(DateThis, "yyyy-MM-dd")), _
            CDate(Format(DateThis, "yyyy-MM-dd")) + 1 - 1 / 24 / 60 / 60, CDate(Format(DateThis - 1, "yyyy-MM-dd")), datNow, gdatRegistTime)

        mblnNotClick = True
        cboTime.Clear
        cboTime.AddItem "所有"
        strSpan = ""
        Do While Not mrsPlan.EOF
            If InStr(strSpan, "," & mrsPlan!排班 & ",") = 0 Then
                strSpan = strSpan & "," & mrsPlan!排班 & ","
                cboTime.AddItem Nvl(mrsPlan!排班)
            End If
            mrsPlan.MoveNext
        Loop
        cboTime.ListIndex = 0
        If mrsPlan.RecordCount <> 0 Then mrsPlan.MoveFirst
        mblnNotClick = False
    Else
       '缓存从筛选
        If picBookingDate.Visible Or mbytMode = 1 Or (mbytMode = 2 Or mbytMode = 0 And mbytInState = 0 And mstrNoIn <> "") Then
            DateThis = dtpAppointmentDate.Value
        Else
            DateThis = zlDatabase.Currentdate
        End If
        If txt号别.Text = "" Then
            mrsPlan.Filter = IIf(cboTime.ListIndex = 0, "", "排班='" & cboTime.Text & "'")
        Else
            mrsPlan.Filter = "号别 like '" & txt号别.Text & "*'" & IIf(cboTime.ListIndex = 0, "", " And 排班='" & cboTime.Text & "'")
        End If
    End If
    
    With vsfPlan
        .Redraw = flexRDNone
        If Not mrsPlan.EOF Then
            '获取所有号源ID，以及不当班号别
            Do While Not mrsPlan.EOF
                If zlCommFun.ActualLen(str号源IDs & "," & Nvl(mrsPlan!号源ID)) > 4000 Then
                    str号源IDSQL = str号源IDSQL & vbNewLine & _
                        " Union All Select Column_Value From Table(f_Num2list('" & Mid(str号源IDs, 2) & "'))"
                    str号源IDs = ""
                End If
                str号源IDs = str号源IDs & "," & Nvl(mrsPlan!号源ID)
                
                If Val(Nvl(mrsPlan!ID)) = 0 Then
                    If zlCommFun.ActualLen(str不当班号码 & "," & Nvl(mrsPlan!号别)) > 4000 Then
                        str不当班号码SQL = str不当班号码SQL & vbNewLine & _
                            " Union All Select Column_Value From Table(f_Str2list('" & Mid(str不当班号码, 2) & "'))"
                        str不当班号码 = ""
                    End If
                    str不当班号码 = str不当班号码 & "," & Nvl(mrsPlan!号别)
                End If
                mrsPlan.MoveNext
            Loop
            If str号源IDs <> "" Then
                str号源IDSQL = str号源IDSQL & vbNewLine & _
                    " Union All Select Column_Value From Table(f_Num2list('" & Mid(str号源IDs, 2) & "'))"
            End If
            If str不当班号码 <> "" Then
                str不当班号码SQL = str不当班号码SQL & vbNewLine & _
                    " Union All Select Column_Value From Table(f_Str2list('" & Mid(str不当班号码, 2) & "'))"
            End If
            If str号源IDSQL <> "" Then str号源IDSQL = Mid(str号源IDSQL, 14)
            If str不当班号码SQL <> "" Then str不当班号码SQL = Mid(str不当班号码SQL, 14)
            
            If str不当班号码SQL <> "" Then
                str不当班SQL = "Select Count(1) As 数量, 号别" & vbNewLine & _
                                "From 病人挂号记录" & vbNewLine & _
                                "Where 出诊记录ID Is Null And 记录性质 = 1 And 记录状态 = 1 And 发生时间 >= [1] And 发生时间 <= [2]" & vbNewLine & _
                                "      And 号别 In (" & str不当班号码SQL & ")" & vbNewLine & _
                                "Group By 号别"
                Set rs不当班 = zlDatabase.OpenSQLRecord(str不当班SQL, Me.Caption, CDate(Format(DateThis, "yyyy-MM-dd")), CDate(Format(DateThis, "yyyy-MM-dd")) + 1 - 1 / 24 / 60 / 60)
            End If
            mrsPlan.MoveFirst
            If mbytMode = 0 Then
                dat开始时间 = CDate(Format(DateThis, "yyyy-mm-dd")) - 1
                dat结束时间 = CDate(Format(DateThis, "yyyy-mm-dd")) + 5
            Else
                datNow = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd"))
                If CDate(Format(DateThis, "yyyy-mm-dd")) - datNow >= 3 Then
                    dat开始时间 = CDate(Format(DateThis, "yyyy-mm-dd")) - 3
                    dat结束时间 = CDate(Format(DateThis, "yyyy-mm-dd")) + 3
                Else
                    dat开始时间 = CDate(Format(datNow, "yyyy-mm-dd")) - 1
                    dat结束时间 = CDate(Format(datNow, "yyyy-mm-dd")) + 5
                End If
            End If
            strDays = "Select 号源id, To_Char(出诊日期,'DD') As 日期, To_Char(出诊日期, 'D') As 周天, 上班时段" & vbNewLine & _
                    "From 临床出诊记录" & vbNewLine & _
                    "Where 号源id In (" & str号源IDSQL & ") And 出诊日期 Between [1] And [2]" & vbNewLine & _
                    "Order By 周天"
            Set rsDays = zlDatabase.OpenSQLRecord(strDays, Me.Caption, dat开始时间, dat结束时间)
            If Not rsDays.EOF Then
                Do While Not rsDays.EOF
                    Select Case Val(Nvl(rsDays!周天))
                        Case 1
                            .TextMatrix(i, .ColIndex("日")) = "日" & vbCrLf & rsDays!日期 & ""
                        Case 2
                            .TextMatrix(i, .ColIndex("一")) = "一" & vbCrLf & rsDays!日期 & ""
                        Case 3
                            .TextMatrix(i, .ColIndex("二")) = "二" & vbCrLf & rsDays!日期 & ""
                        Case 4
                            .TextMatrix(i, .ColIndex("三")) = "三" & vbCrLf & rsDays!日期 & ""
                        Case 5
                            .TextMatrix(i, .ColIndex("四")) = "四" & vbCrLf & rsDays!日期 & ""
                        Case 6
                            .TextMatrix(i, .ColIndex("五")) = "五" & vbCrLf & rsDays!日期 & ""
                        Case 7
                            .TextMatrix(i, .ColIndex("六")) = "六" & vbCrLf & rsDays!日期 & ""
                    End Select
                    rsDays.MoveNext
                Loop
                rsDays.MoveFirst
            End If
            
            .ToolTipText = "共 " & mrsPlan.RecordCount & " 条安排"
            .Clear 1
            mblnChangeByCode = True
            .Rows = 2
            If mrsPlan.RecordCount <> 0 Then
                .Rows = mrsPlan.RecordCount + 1
            End If
            mblnChangeByCode = False

            mrsPlan.MoveFirst
            For i = 1 To mrsPlan.RecordCount
                .RowData(i) = IIf(mrsPlan!急诊 = 1, -1, 1) * mrsPlan!科室ID
                .TextMatrix(i, .ColIndex("IDS")) = mrsPlan!ID & "," & mrsPlan!项目ID & "," & IIf(IsNull(mrsPlan!医生ID), 0, mrsPlan!医生ID)
                .Cell(flexcpData, i, .ColIndex("IDS")) = Nvl(mrsPlan!ID)
                .TextMatrix(i, .ColIndex("号类")) = IIf(IsNull(mrsPlan!号类), "", mrsPlan!号类)
                .TextMatrix(i, .ColIndex("号别")) = mrsPlan!号别
                .TextMatrix(i, .ColIndex("科室")) = mrsPlan!科室
                .TextMatrix(i, .ColIndex("项目")) = mrsPlan!项目
                .TextMatrix(i, .ColIndex("出诊日期")) = Format(Nvl(mrsPlan!出诊日期), "yyyy-mm-dd")
                .TextMatrix(i, .ColIndex("医生")) = Nvl(mrsPlan!医生)
                .TextMatrix(i, .ColIndex("已约")) = Nvl(mrsPlan!已约)
                .TextMatrix(i, .ColIndex("限约")) = Nvl(mrsPlan!限约)
                If Not rs不当班 Is Nothing And Val(Nvl(mrsPlan!ID)) = 0 Then
                    rs不当班.Filter = "号别=" & "'" & mrsPlan!号别 & "'"
                    If Not rs不当班.EOF Then
                        .TextMatrix(i, .ColIndex("已挂")) = Nvl(rs不当班!数量)
                    End If
                Else
                    .TextMatrix(i, .ColIndex("已挂")) = Nvl(mrsPlan!已挂)
                End If
                .TextMatrix(i, .ColIndex("限号")) = Nvl(mrsPlan!限号)
                If Val(.TextMatrix(i, .ColIndex("限号"))) <> 0 And mbytMode = 0 Then
                    .TextMatrix(i, .ColIndex("剩余")) = Val(.TextMatrix(i, .ColIndex("限号"))) - Val(.TextMatrix(i, .ColIndex("已挂")))
                    intSurplusTotal = intSurplusTotal + Val(.TextMatrix(i, .ColIndex("剩余")))
                End If
                .TextMatrix(i, .ColIndex("提前时间")) = Format(mrsPlan!提前时间, "yyyy-MM-dd hh:mm:ss")
                .TextMatrix(i, .ColIndex("挂号时间")) = Format(mrsPlan!挂号时间, "yyyy-MM-dd hh:mm:ss")
                .TextMatrix(i, .ColIndex("结束时间")) = Format(mrsPlan!结束时间, "yyyy-MM-dd hh:mm:ss")
                .TextMatrix(i, .ColIndex("时段")) = Nvl(mrsPlan!排班)
                .TextMatrix(i, .ColIndex("是否独占")) = Nvl(mrsPlan!是否独占)
                If Nvl(mrsPlan!替诊医生姓名) <> "" Then
                    .TextMatrix(i, .ColIndex("替诊医生")) = ""
                    .Cell(flexcpData, i, .ColIndex("替诊医生")) = Nvl(mrsPlan!替诊医生姓名) & "(" & Format(Nvl(mrsPlan!替诊开始时间), "hh:mm") & "-" & Format(Nvl(mrsPlan!替诊终止时间), "hh:mm") & ")"
                    .TextMatrix(i, .ColIndex("替诊开始时间")) = Format(mrsPlan!替诊开始时间, "yyyy-MM-dd hh:mm:ss")
                    .TextMatrix(i, .ColIndex("替诊终止时间")) = Format(mrsPlan!替诊终止时间, "yyyy-MM-dd hh:mm:ss")
                    .TextMatrix(i, .ColIndex("替诊医生姓名")) = Nvl(mrsPlan!替诊医生姓名)
                    .TextMatrix(i, .ColIndex("替诊医生ID")) = Nvl(mrsPlan!替诊医生id)
                End If
                
                rsDays.Filter = "号源id=" & Val(mrsPlan!号源ID)
                Do While Not rsDays.EOF
                    Select Case Val(Nvl(rsDays!周天))
                    Case 1
                        If Nvl(rsDays!上班时段) = Nvl(mrsPlan!排班) Then
                            .TextMatrix(i, .ColIndex("日")) = Left(Nvl(mrsPlan!排班), 1)
                            .Cell(flexcpData, i, .ColIndex("日")) = Nvl(mrsPlan!排班)
                        Else
                            If .Cell(flexcpData, i, .ColIndex("日")) <> Nvl(mrsPlan!排班) Then
                                If .TextMatrix(i, .ColIndex("日")) = "" Then
                                    .TextMatrix(i, .ColIndex("日")) = Left(Nvl(rsDays!上班时段), 1)
                                    .Cell(flexcpData, i, .ColIndex("日")) = Nvl(rsDays!上班时段)
                                Else
                                    .TextMatrix(i, .ColIndex("日")) = .TextMatrix(i, .ColIndex("日")) & "/" & Left(Nvl(rsDays!上班时段), 1)
                                    .Cell(flexcpData, i, .ColIndex("日")) = .Cell(flexcpData, i, .ColIndex("日")) & "/" & Nvl(rsDays!上班时段)
                                End If
                            End If
                        End If
                    Case 2
                        If Nvl(rsDays!上班时段) = Nvl(mrsPlan!排班) Then
                            .TextMatrix(i, .ColIndex("一")) = Left(Nvl(mrsPlan!排班), 1)
                            .Cell(flexcpData, i, .ColIndex("一")) = Nvl(mrsPlan!排班)
                        Else
                            If .Cell(flexcpData, i, .ColIndex("一")) <> Nvl(mrsPlan!排班) Then
                                If .TextMatrix(i, .ColIndex("一")) = "" Then
                                    .TextMatrix(i, .ColIndex("一")) = Left(Nvl(rsDays!上班时段), 1)
                                    .Cell(flexcpData, i, .ColIndex("一")) = Nvl(rsDays!上班时段)
                                Else
                                    .TextMatrix(i, .ColIndex("一")) = .TextMatrix(i, .ColIndex("一")) & "/" & Left(Nvl(rsDays!上班时段), 1)
                                    .Cell(flexcpData, i, .ColIndex("一")) = .Cell(flexcpData, i, .ColIndex("一")) & "/" & Nvl(rsDays!上班时段)
                                End If
                            End If
                        End If
                    Case 3
                        If Nvl(rsDays!上班时段) = Nvl(mrsPlan!排班) Then
                            .TextMatrix(i, .ColIndex("二")) = Left(Nvl(mrsPlan!排班), 1)
                            .Cell(flexcpData, i, .ColIndex("二")) = Nvl(mrsPlan!排班)
                        Else
                            If .Cell(flexcpData, i, .ColIndex("二")) <> Nvl(mrsPlan!排班) Then
                                If .TextMatrix(i, .ColIndex("二")) = "" Then
                                    .TextMatrix(i, .ColIndex("二")) = Left(Nvl(rsDays!上班时段), 1)
                                    .Cell(flexcpData, i, .ColIndex("二")) = Nvl(rsDays!上班时段)
                                Else
                                    .TextMatrix(i, .ColIndex("二")) = .TextMatrix(i, .ColIndex("二")) & "/" & Left(Nvl(rsDays!上班时段), 1)
                                    .Cell(flexcpData, i, .ColIndex("二")) = .Cell(flexcpData, i, .ColIndex("二")) & "/" & Nvl(rsDays!上班时段)
                                End If
                            End If
                        End If
                    Case 4
                        If Nvl(rsDays!上班时段) = Nvl(mrsPlan!排班) Then
                            .TextMatrix(i, .ColIndex("三")) = Left(Nvl(mrsPlan!排班), 1)
                            .Cell(flexcpData, i, .ColIndex("三")) = Nvl(mrsPlan!排班)
                        Else
                            If .Cell(flexcpData, i, .ColIndex("三")) <> Nvl(mrsPlan!排班) Then
                                If .TextMatrix(i, .ColIndex("三")) = "" Then
                                    .TextMatrix(i, .ColIndex("三")) = Left(Nvl(rsDays!上班时段), 1)
                                    .Cell(flexcpData, i, .ColIndex("三")) = Nvl(rsDays!上班时段)
                                Else
                                    .TextMatrix(i, .ColIndex("三")) = .TextMatrix(i, .ColIndex("三")) & "/" & Left(Nvl(rsDays!上班时段), 1)
                                    .Cell(flexcpData, i, .ColIndex("三")) = .Cell(flexcpData, i, .ColIndex("三")) & "/" & Nvl(rsDays!上班时段)
                                End If
                            End If
                        End If
                    Case 5
                        If Nvl(rsDays!上班时段) = Nvl(mrsPlan!排班) Then
                            .TextMatrix(i, .ColIndex("四")) = Left(Nvl(mrsPlan!排班), 1)
                            .Cell(flexcpData, i, .ColIndex("四")) = Nvl(mrsPlan!排班)
                        Else
                            If .Cell(flexcpData, i, .ColIndex("四")) <> Nvl(mrsPlan!排班) Then
                                If .TextMatrix(i, .ColIndex("四")) = "" Then
                                    .TextMatrix(i, .ColIndex("四")) = Left(Nvl(rsDays!上班时段), 1)
                                    .Cell(flexcpData, i, .ColIndex("四")) = Nvl(rsDays!上班时段)
                                Else
                                    .TextMatrix(i, .ColIndex("四")) = .TextMatrix(i, .ColIndex("四")) & "/" & Left(Nvl(rsDays!上班时段), 1)
                                    .Cell(flexcpData, i, .ColIndex("四")) = .Cell(flexcpData, i, .ColIndex("四")) & "/" & Nvl(rsDays!上班时段)
                                End If
                            End If
                        End If
                    Case 6
                        If Nvl(rsDays!上班时段) = Nvl(mrsPlan!排班) Then
                            .TextMatrix(i, .ColIndex("五")) = Left(Nvl(mrsPlan!排班), 1)
                            .Cell(flexcpData, i, .ColIndex("五")) = Nvl(mrsPlan!排班)
                        Else
                            If .Cell(flexcpData, i, .ColIndex("五")) <> Nvl(mrsPlan!排班) Then
                                If .TextMatrix(i, .ColIndex("五")) = "" Then
                                    .TextMatrix(i, .ColIndex("五")) = Left(Nvl(rsDays!上班时段), 1)
                                    .Cell(flexcpData, i, .ColIndex("五")) = Nvl(rsDays!上班时段)
                                Else
                                    .TextMatrix(i, .ColIndex("五")) = .TextMatrix(i, .ColIndex("五")) & "/" & Left(Nvl(rsDays!上班时段), 1)
                                    .Cell(flexcpData, i, .ColIndex("五")) = .Cell(flexcpData, i, .ColIndex("五")) & "/" & Nvl(rsDays!上班时段)
                                End If
                            End If
                        End If
                    Case 7
                        If Nvl(rsDays!上班时段) = Nvl(mrsPlan!排班) Then
                            .TextMatrix(i, .ColIndex("六")) = Left(Nvl(mrsPlan!排班), 1)
                            .Cell(flexcpData, i, .ColIndex("六")) = Nvl(mrsPlan!排班)
                        Else
                            If .Cell(flexcpData, i, .ColIndex("六")) <> Nvl(mrsPlan!排班) Then
                                If .TextMatrix(i, .ColIndex("六")) = "" Then
                                    .TextMatrix(i, .ColIndex("六")) = Left(Nvl(rsDays!上班时段), 1)
                                    .Cell(flexcpData, i, .ColIndex("六")) = Nvl(rsDays!上班时段)
                                Else
                                    .TextMatrix(i, .ColIndex("六")) = .TextMatrix(i, .ColIndex("六")) & "/" & Left(Nvl(rsDays!上班时段), 1)
                                    .Cell(flexcpData, i, .ColIndex("六")) = .Cell(flexcpData, i, .ColIndex("六")) & "/" & Nvl(rsDays!上班时段)
                                End If
                            End If
                        End If
                    End Select
                    rsDays.MoveNext
                Loop
                .TextMatrix(i, .ColIndex("病案")) = IIf(mrsPlan!病案 = 1, "√", "")
                .TextMatrix(i, .ColIndex("分诊")) = Nvl(mrsPlan!分诊)
                .TextMatrix(i, .ColIndex("序号控制")) = IIf(mrsPlan!序号控制 = 1, "√", "")
                .TextMatrix(i, .ColIndex("记录ID")) = Nvl(mrsPlan!ID)
                .TextMatrix(i, .ColIndex("号源ID")) = Nvl(mrsPlan!号源ID)
                .TextMatrix(i, .ColIndex("分时段")) = Val(Nvl(mrsPlan!分时段))
                .Cell(flexcpData, i, .ColIndex("号别")) = ""
                If InStr(mstrPrivs, ";临时挂号;") = 0 And chkShowAll.Value = 1 Then
                    If Val(Nvl(mrsPlan!ID)) = 0 Or DateThis < CDate(IIf(.TextMatrix(i, .ColIndex("提前时间")) = "", IIf(.TextMatrix(i, .ColIndex("挂号时间")) = "", "3000-01-01", .TextMatrix(i, .ColIndex("挂号时间"))), .TextMatrix(i, .ColIndex("提前时间")))) Then
                        .Cell(flexcpData, i, .ColIndex("号别")) = "1"
                        .Cell(flexcpForeColor, i, GetCol("IDS"), i, .Cols - 1) = &H8000000C
                    End If
                End If
                
                If mrsPlan!号别 = txt号别.Text And k = 0 And (Nvl(mrsPlan!ID) = mlng记录ID Or mlng记录ID = 0) Then k = i
                '问题 43847
                If k = 0 And mrsPlan!号别 = mstrPreNO And (mSortType = by号别 Or txt号别.Text = "") Then k = i
                mrsPlan.MoveNext
            Next
            lblRegTotal(1).Caption = intSurplusTotal
        Else
            Set mrsPlan = Nothing
            Call SetPlanGrid
            .ToolTipText = ""
        End If
        zl_vsGrid_Para_Restore mlngModul, vsfPlan, Me.Caption, "vsfPlan" & mbytMode
        If k <> 0 Then
            mblnChangeByCode = True
            .Row = k
            mblnChangeByCode = False
            '53299
            mlngPreRow = k
            Call SetGridTop(k)
        Else
'            .Row = .FixedRows + 1
        End If
        Call SetvsfplanFiexBackColor
        If picBookingDate.Visible Or mbytMode = 1 Or (mbytMode = 2 Or mbytMode = 0 And mbytInState = 0 And mstrNoIn <> "") Then
            Call SetvsfplanFiexBackColor(False)
        End If
        .Col = 0: .ColSel = .Cols - 1
        '70193:刘尔旋,2014-2-18,号别自动定位错误的问题
        If vsfPlan.Row = 1 Then
            vsfPlan.Select 1, 1
            If txt号别.Visible And txt号别.Enabled Then txt号别.SetFocus
        End If
        If vsfPlan.Rows = 3 Then Call vsfPlan_EnterCell
        If k <> 0 And k = vsfPlan.RowSel Then
            For i = 0 To .Cols - 1
                If .Cell(flexcpBackColor, k, i, k, i) <> &HFF8080 Then .Cell(flexcpBackColor, k, i, k, i) = 16772055
            Next i
        End If
        .Redraw = flexRDBuffered
    End With
    ShowPlans = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Set mrsPlan = Nothing
End Function
Private Function zlRePrintRegistered() As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '功能：重打
    '返回:重打成功,返回true,否则返回False
    '编制：刘兴洪
    '日期：2010-07-02 10:49:06
    '说明：主要是重新整理代码
    '------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, str年龄 As String, str性别 As String, str出生日期 As String
    Dim lng结帐ID As Long, lng病人ID As Long, intInsure As Integer
    Dim strNO As String, blnVirtualPrint As Boolean
    
    If cboNO.Tag = "" Then
        MsgBox "未输入挂号单据，不能重打！", vbInformation, gstrSysName
        Exit Function
    End If
    strNO = cboNO.Tag
    If CheckBillExistReplenishData(strNO) Then
        MsgBox "选择的挂号记录进行了医保补充结算，不允许重打操作！", vbInformation, gstrSysName
        Exit Function
    End If
    lng结帐ID = GetBill结帐ID(strNO, 4, lng病人ID)
    intInsure = ExistInsure(strNO)
    If intInsure <> 0 Then
        blnVirtualPrint = gclsInsure.GetCapability(support医保接口打印票据, lng病人ID, intInsure)
    End If
    
    Dim blnStartFactUseType  As Boolean, strUseType As String
    If gblnSharedInvoice Then
        blnStartFactUseType = zlStartFactUseType("1")
        If blnStartFactUseType Then
            strUseType = zl_GetInvoiceUserType(lng病人ID, 0, intInsure)
        End If
    End If
    
    
    If txtPatientPrint.Visible Then
        If txtPatientPrint.Text = "" Then
            MsgBox "姓名为空,请输入姓名！", vbInformation, gstrSysName
            If txtPatientPrint.Enabled Then txtPatientPrint.SetFocus
            Exit Function
        End If
        str年龄 = Trim(txt年龄.Text): str性别 = NeedName(cbo性别.Text)
        If IsNumeric(str年龄) Then str年龄 = str年龄 & cbo年龄单位.Text
        If txtPatient.Text <> txtPatientPrint.Text Or mstr年龄 & mstr年龄单位 <> str年龄 Or mstr性别 <> str性别 Then
            If zlExistOperationData(Val(txtPatientPrint.Tag), cboNO.Tag) Then
                MsgBox "注意:" & vbCrLf & "该病人已经发生医嘱业务数据,不能调整病人的基本信息,请在『病人信息管理』中调整!" & vbCrLf & "点击确定后恢复修改的病人信息。", vbOKOnly + vbDefaultButton1, gstrSysName
                txt年龄.Text = mstr年龄
                If mstr年龄单位 <> "" Then cbo年龄单位.ListIndex = cbo.FindIndex(cbo年龄单位, mstr年龄单位, True): cbo年龄单位.Visible = True: txt年龄.Width = 600
                str年龄 = Trim(txt年龄.Text): str性别 = NeedName(cbo性别.Text)
                If IsNumeric(str年龄) Then str年龄 = str年龄 & cbo年龄单位.Text
                cbo性别.ListIndex = cbo.FindIndex(cbo性别, mstr性别, True)
                txtPatient.Text = mstr姓名
                Exit Function
            End If
            str出生日期 = "NULL"
            '35544
            If str年龄 <> mstr年龄 Then
                If IsNumeric(CStr(txt年龄.Text)) Then
                    str出生日期 = ReCalcBirth(txt年龄.Text, cbo年龄单位.Text)
                    If IsDate(str出生日期) = False Then
                        str出生日期 = "NULL"
                    Else
                        str出生日期 = "to_date('" & str出生日期 & "','yyyy-mm-dd')"
                    End If
                End If
            End If
            'Zl_病人费用记录_Update
            strSQL = "Zl_病人费用记录_Update("
            '  No_In       门诊费用记录.NO%Type,
            strSQL = strSQL & "'" & strNO & "',"
            '  记录性质_In 门诊费用记录.记录性质%Type,
            strSQL = strSQL & "" & 4 & ","
            '  开单人_In   门诊费用记录.开单人%Type,
            strSQL = strSQL & "" & "Null" & ","
            '  发生时间_In 门诊费用记录.发生时间%Type,
            strSQL = strSQL & "" & "Null" & ","
            '  姓名_In     门诊费用记录.姓名%Type := Null,
            strSQL = strSQL & "'" & txtPatientPrint.Text & "',"
            '  来源_In     Integer := 1,
            strSQL = strSQL & "" & 1 & ","
            '  年龄_In     门诊费用记录.年龄%Type := Null,
            strSQL = strSQL & "" & IIf(str年龄 = "", "NULL", "'" & str年龄 & "'") & ","
            '  性别_In     门诊费用记录.性别%Type := Null
            strSQL = strSQL & "" & IIf(str性别 = "", "NULL", "'" & str性别 & "'") & ","
            '  出生日期_In 病人信息.出生日期%Type := Null
            strSQL = strSQL & "" & str出生日期 & ")"
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        End If
    End If
   '问题:53037
    If Not RePrintBill(Me, 3, strNO, lng结帐ID, intInsure, blnVirtualPrint, strUseType, True) Then Exit Function

    zlRePrintRegistered = True
End Function

Private Function GetTotal(ByVal strNO As String) As Double
    Dim strSQL As String, rsTmp As ADODB.Recordset
    strSQL = "Select Sum(结帐金额) As 总金额 From 门诊费用记录 Where No = [1] And 记录性质 = 4"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
    If Not rsTmp.EOF Then GetTotal = Val(Nvl(rsTmp!总金额))
End Function


Private Function zlExcuteDelRegistered() As Boolean
    '------------------------------------ ------------------------------------------------------------------------------------
    '功能：挂号退号
    '返回：退号成功,返回true,否则返回False
    '编制：刘兴洪
    '日期：2010-07-02 10:53:29
    '说明：重新整理代码时,加上此过程
    '------------------------------------------------------------------------------------------------------------------------
    Dim rsOneCard1 As ADODB.Recordset, rsOneCard2 As ADODB.Recordset, objICCard As Object
    Dim blnPromptClear As Boolean, strSQL As String, strNO As String, lngCard结帐ID As Long
    Dim strSQLCard As String, intMsgReturn As Integer, bln退费重打 As Boolean, blnTrans As Boolean
    Dim bytTogetherDo As Byte, dblTotal As Double                            '0-无附加操作,1-删除门诊号
    Dim strAdvance  As String, strCardNo As String, lng结帐ID As Long
    Dim blnNotCommit As Boolean
    Dim Curdate As Date '问题号:56599
    Dim str操作 As String '问题号:56599
    Dim str卡号 As String '问题号:56599
    Dim rs医疗卡类别 As Recordset '问题号:56599
    Dim cllPro As Collection, cllBillBalance As Collection, dblThreeMoney As Double
    Dim cllUpdate As Collection, cllThreeIns As Collection, strErrMsg As String
    Dim byt退费类型 As Byte '0-全退 1-退挂号费 2-退病历费
    Dim i As Long, curMoney As Currency, dbl预交支付 As Double
    Dim curChkMoney As Currency
    Dim blnCardReprint As Boolean
    Dim objCard As Card, str结算方式 As String, str现金 As String, dbl现金 As Double
    Dim strBalance As String, strDelCardNo As String, str原结算方式 As String
    Dim strInvoice As String, lng病人ID As Long, lng领用ID As Long
    Dim bln记帐 As Boolean, bln结帐 As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim bln原样退 As Boolean, strTemp As String, strBackNote As String
    Dim bln医保原样退 As Boolean
    Dim dbl预交 As Double
    Dim rsInvoice As ADODB.Recordset
    Dim strBackInvoice As String, blnReprint As Boolean
    Dim dblCheckThreeMoney As Double
    
    Set cllPro = New Collection
 
    
    strNO = cboNO.Tag
    If strNO = "" Then
        MsgBox "未输入挂号单据，不能退号！", vbInformation, gstrSysName
        Exit Function
    End If
    If CheckBillExistReplenishData(strNO) Then
        MsgBox "选择的挂号记录进行了医保补充结算，不允许进行退号操作！", vbInformation, gstrSysName
        Exit Function
    End If
    If cbo备注.Text <> "" And cbo备注.Tag = "" And mbln退号原因 And cbo备注.Enabled And cbo备注.Visible Then
        If cbo备注.Text <> mstr原摘要 Then
            MsgBox "请在摘要中选择正确的退号原因!", vbInformation, gstrSysName
            cbo备注.SetFocus
            Exit Function
        End If
    End If
    '68991
    lng结帐ID = GetBill结帐ID(strNO, 4, lng病人ID, bln记帐)
    If zlCheckIsAllowBackSN(strNO, bln记帐, bln结帐) = False Then Exit Function
    
    If Not bln记帐 Then
        '问题:51527
        Call zlReadRegThreeBalance(strNO, cllBillBalance, objCard)
        bln原样退 = True
        If Not objCard Is Nothing Then
            bln原样退 = False
            For i = 1 To vsfPay.Rows - 1
                If vsfPay.RowData(i) = 1 Then
                    strTemp = vsfPay.TextMatrix(i, 0)
                End If
                If Val(vsfPay.TextMatrix(i, 4)) = objCard.接口序号 Then
                    bln原样退 = True
                End If
            Next i
        End If
        If bln原样退 = False Then
            str结算方式 = strTemp
            str原结算方式 = objCard.结算方式
            Set mCurCardPay.objCard = Nothing
            mCurCardPay.lng医疗卡类别ID = 0
        End If
    Else
        str结算方式 = ""
    End If
    
    
    blnPromptClear = True
    If vsfMoney.Tag = "卡费" Then   '处理挂号费和卡费没有分离以前的
        If MsgBox("当前要退号的单据费用中包含就诊卡费,将一起退费!" & vbCrLf & _
            "你确实要进行退号吗?", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
           cboNO.Text = "": cboNO.SetFocus: Exit Function
        End If
    Else
        strDelCardNo = ExistCardFee(strNO, lngCard结帐ID, str卡号)
        If strDelCardNo <> "" Then
            '问题号:56599
            If str卡号 <> "" Then
                '113613：李南春，2018/1/18，退卡时检查当前卡是否允许退卡
                strSQL = "Select Nvl(是否自制,0) As 是否自制,zl1_EX_ReFundCard_Check([1],[2],A.卡类别ID,[3]) as 验证" & _
                "           From 病人医疗卡信息 A,医疗卡类别 B " & _
                "           Where A.卡号=[3] And A.卡类别ID =B.ID "
                Set rs医疗卡类别 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngModul, lng病人ID, str卡号)
                If rs医疗卡类别.EOF = False Then
                    If Nvl(rs医疗卡类别!验证) <> "" Then
                        If Not objCard Is Nothing Then
                            If mCurCardPay.lng医疗卡类别ID <> 0 And mCurCardPay.bln消费卡 = False And objCard.是否全退 Then
                                MsgBox Nvl(rs医疗卡类别!验证) & "，不能单独退挂号费！", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName
                                cboNO.Text = "": cboNO.SetFocus: Exit Function
                            End If
                        End If
                        If MsgBox(Nvl(rs医疗卡类别!验证) & "，是否单独退挂号费？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                            cboNO.Text = "": cboNO.SetFocus: Exit Function
                        End If
                        str操作 = "仅退号"
                    ElseIf rs医疗卡类别!是否自制 = 0 Then '院外卡
                        str操作 = zl9ComLib.zlCommFun.ShowMsgbox(gstrSysName, "卡号:" & str卡号 & "卡为院外卡发,请选择退卡或取消绑定操作", "退卡,取消绑定", Me, vbQuestion)
                    End If
                End If
            End If
            
            '问题号:56599
            If str操作 <> "" Then
                 Select Case str操作
                    Case "退卡"
                        'Zl_医疗卡记录_Delete
                        strSQLCard = "Zl_医疗卡记录_Delete("
                        '      单据号_In     住院费用记录.No%Type,
                        strSQLCard = strSQLCard & "'" & strDelCardNo & "',"
                        '      操作员编号_In 住院费用记录.操作员编号%Type,
                        strSQLCard = strSQLCard & "'" & UserInfo.编号 & "',"
                        '      操作员姓名_In 住院费用记录.操作员姓名%Type,
                        strSQLCard = strSQLCard & "'" & UserInfo.姓名 & "')"
                    Case "取消绑定"
                        Curdate = zlDatabase.Currentdate
                        'Zl_医疗卡变动_Insert
                         strSQLCard = "Zl_医疗卡变动_Insert("
                        '      变动类型_In   Number,
                        '发卡类型=1-发卡(或11绑定卡);2-换卡;3-补卡(13-补卡停用);4-退卡(或14取消绑定); ５-密码调整(只记录);6-挂失(16取消挂失)
                        strSQLCard = strSQLCard & "" & 14 & ","
                        '      病人id_In     住院费用记录.病人id%Type,
                        strSQLCard = strSQLCard & "" & lng病人ID & ","
                        '      卡类别id_In   病人医疗卡信息.卡类别id%Type,
                        strSQLCard = strSQLCard & "" & mCurCardPay.lng医疗卡类别ID & ","
                        '      原卡号_In     病人医疗卡信息.卡号%Type,
                        strSQLCard = strSQLCard & "NULL,"
                        '      医疗卡号_In   病人医疗卡信息.卡号%Type,
                        strSQLCard = strSQLCard & str卡号 & ","
                        '      变动原因_In   病人医疗卡变动.变动原因%Type,
                        strSQLCard = strSQLCard & "'取消卡号绑定',"
                        '      密码_In       病人信息.卡验证码%Type,
                        strSQLCard = strSQLCard & "NULL,"
                        '      操作员姓名_In 住院费用记录.操作员姓名%Type,
                        strSQLCard = strSQLCard & "NULL,"
                        '      变动时间_In   住院费用记录.登记时间%Type,
                        strSQLCard = strSQLCard & "to_date('" & Format(Curdate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
                        '      Ic卡号_In     病人信息.Ic卡号%Type := Null,
                        strSQLCard = strSQLCard & "NULL,"
                        '      挂失方式_In   病人医疗卡变动.挂失方式%Type := Null
                        strSQLCard = strSQLCard & "NULL)"
                 End Select
            Else
                If str操作 = "仅退号" Then
                    intMsgReturn = vbNo
                Else
                     '116278:李南春,2017/12/15，不支持部分退的三方卡，退号必须同时退卡
                    If Not objCard Is Nothing Then
                        If mCurCardPay.lng医疗卡类别ID <> 0 And mCurCardPay.bln消费卡 = False And objCard.是否全退 Then
                            intMsgReturn = MsgBox("该病人挂号时发过卡,退号必须同时退卡,是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName)
                            If intMsgReturn = vbNo Then Exit Function
                        Else
                            intMsgReturn = MsgBox("该病人挂号时发过卡,退号同时退卡吗？", vbQuestion + vbYesNoCancel + vbDefaultButton2, gstrSysName)
                        End If
                    Else
                        intMsgReturn = MsgBox("该病人挂号时发过卡,退号同时退卡吗？", vbQuestion + vbYesNoCancel + vbDefaultButton2, gstrSysName)
                    End If
                End If
                If intMsgReturn = vbYes Then
                    strSQLCard = "zl_医疗卡记录_DELETE('" & strDelCardNo & "','" & UserInfo.编号 & "','" & UserInfo.姓名 & "')"
                ElseIf intMsgReturn = vbNo Then
                    strSQLCard = ""
                    lngCard结帐ID = 0
                    bln退费重打 = gbln退费重打
                    blnCardReprint = gbln退费重打
                    blnPromptClear = False
                Else
                    Exit Function
                End If
            End If
        End If
    End If
    
    '问题:51527
    dblThreeMoney = 0
    If mCurCardPay.lng医疗卡类别ID <> 0 Then
        dblThreeMoney = zlGetRegThreeMoney(lng结帐ID, lngCard结帐ID, cllBillBalance)
    End If
    dblCheckThreeMoney = zlGetRegThreeMoney(lng结帐ID, lngCard结帐ID, cllBillBalance)
    
    bytTogetherDo = 0
    '全退
    If mintCancel = 0 And mbln主费用 = True Then
        If Not (mbln包含病历费 And chk病历费.Value = 0) And Not (mbln附加费 And chkExtra.Value = 0) Then
            '如果挂号单的登记日期-病人信息的登记日期在挂号单有效天数之内,则提示是否删除门诊号   txt发生时间
            If txt门诊号.Text <> "" And blnPromptClear Then
                If Check挂号时建档(strNO, txt发生时间.Text) Then
                    Select Case gbyt清除门诊信息    '35176
                    Case 0  '不清除
                    Case 1  '清除
                           bytTogetherDo = 1
                    Case 2  '提示清除
                        If MsgBox("退号后要清除与该病人相关的门诊号信息吗!", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                           bytTogetherDo = 1
                        End If
                    End Select
                End If
            End If
        End If
    End If
    
    dbl预交支付 = 0
    For i = 1 To vsfPay.Rows - 1
        If vsfPay.RowData(i) = 0 And vsfPay.TextMatrix(i, 0) <> "" Then
            dbl预交支付 = dbl预交支付 + Val(vsfPay.TextMatrix(i, 1))
        End If
    Next i
    
    '如果退费涉及预交款,则需要刷卡验证
    If gbyt预存款退费验卡 <> 0 And dbl预交支付 <> 0 Then
        If mrsBill.RecordCount <> 0 Then mrsBill.MoveFirst
        If Not zlDatabase.PatiIdentify(Me, glngSys, Nvl(mrsBill!病人ID, 0), dbl预交支付, _
                            mlngModul, 1, IDKind.GetCurCard.接口序号, , True, , , (gbyt预存款退费验卡 = 2)) Then Exit Function
    End If
    
    Select Case mintCancel
    Case 0
        If mbln主费用 Then
            If ((mbln包含病历费 And chk病历费.Value = 1) Or mbln包含病历费 = False) And ((mbln附加费 And chkExtra.Value = 1) Or mbln附加费 = False) Then
                '本次退费金额计算.
                For i = 1 To vsfMoney.Rows - 1
                    curMoney = Val(vsfMoney.TextMatrix(i, 2)) + curMoney
                Next
                
                byt退费类型 = 0
            ElseIf ((mbln包含病历费 And chk病历费.Value = 0) Or mbln包含病历费 = False) And ((mbln附加费 And chkExtra.Value = 0) Or mbln附加费 = False) Then
                If bln记帐 = False Then
                    If dblCheckThreeMoney <> 0 Then
                        MsgBox "使用三方接口结算的挂号单据,不能将挂号费分开退!", vbInformation, gstrSysName
                        Exit Function
                    End If
                    
                    If mintInsure <> 0 And MCPAR.不收病历费 = False Then
                        MsgBox "使用医保接口结算的挂号单据,不能将挂号费分开退!", vbInformation, gstrSysName
                        Exit Function
                    End If
                    
                    If mstr划价NO <> "" Then
                        MsgBox "挂号产生划价单时,不支持挂号费分别退!", vbInformation, gstrSysName
                        Exit Function
                    End If
                End If
                
                '本次退费金额计算.
                For i = 1 To vsfMoney.Rows - 1
                    curMoney = Val(vsfMoney.TextMatrix(i, 2)) + curMoney
                Next
                
                byt退费类型 = 1
                bln退费重打 = gbln退费重打
            ElseIf mbln包含病历费 And chk病历费.Value = 1 Then
                If mbln附加费 And chkExtra.Value = 0 Then
                    If bln记帐 = False Then
                        If dblCheckThreeMoney <> 0 Then
                            MsgBox "使用三方接口结算的挂号单据,不能将挂号费分开退!", vbInformation, gstrSysName
                            Exit Function
                        End If
                        
                        If mintInsure <> 0 Then
                            MsgBox "使用医保接口结算的挂号单据,不能将挂号费分开退!", vbInformation, gstrSysName
                            Exit Function
                        End If
                        
                        If mstr划价NO <> "" Then
                            MsgBox "挂号产生划价单时,不支持挂号费分别退!", vbInformation, gstrSysName
                            Exit Function
                        End If
                    End If
                End If
                '本次退费金额计算.
                For i = 1 To vsfMoney.Rows - 1
                    curMoney = Val(vsfMoney.TextMatrix(i, 2)) + curMoney
                Next
                
                byt退费类型 = 4
                bln退费重打 = gbln退费重打
            ElseIf mbln附加费 And chkExtra.Value = 1 Then
                If mbln包含病历费 And chk病历费.Value = 0 Then
                    If bln记帐 = False Then
                        If dblCheckThreeMoney <> 0 Then
                            MsgBox "使用三方接口结算的挂号单据,不能将挂号费分开退!", vbInformation, gstrSysName
                            Exit Function
                        End If
                        
                        If mintInsure <> 0 And MCPAR.不收病历费 = False Then
                            MsgBox "使用医保接口结算的挂号单据,不能将挂号费分开退!", vbInformation, gstrSysName
                            Exit Function
                        End If
                        
                        If mstr划价NO <> "" Then
                            MsgBox "挂号产生划价单时,不支持挂号费分别退!", vbInformation, gstrSysName
                            Exit Function
                        End If
                    End If
                End If
                '本次退费金额计算.
                For i = 1 To vsfMoney.Rows - 1
                    curMoney = Val(vsfMoney.TextMatrix(i, 2)) + curMoney
                Next
                
                byt退费类型 = 5
                bln退费重打 = gbln退费重打
            End If
        Else
            If (mbln包含病历费 And chk病历费.Value = 1) And (mbln附加费 And chkExtra.Value = 1) Then
                MsgBox "已经冲销的挂号单据,不能将病历费与附加费一起退!", vbInformation, gstrSysName
                Exit Function
            End If
            
            If (mbln包含病历费 And chk病历费.Value = 1) Then
                If bln记帐 = False Then
                    If dblCheckThreeMoney <> 0 Then
                        MsgBox "使用三方接口结算的挂号单据,不能将病历费与挂号费分开退!", vbInformation, gstrSysName
                        Exit Function
                    End If
                    
                    If mintInsure <> 0 And MCPAR.不收病历费 = False Then
                        MsgBox "使用医保接口结算的挂号单据,不能将挂号费分开退!", vbInformation, gstrSysName
                        Exit Function
                    End If
                    
                    If mstr划价NO <> "" Then
                        MsgBox "挂号产生划价单时,不支持病历费与挂号费分别退!", vbInformation, gstrSysName
                        Exit Function
                    End If
                End If
                
                '本次退费金额计算.
                For i = 1 To vsfMoney.Rows - 1
                    curMoney = Val(vsfMoney.TextMatrix(i, 2)) + curMoney
                Next
                
                byt退费类型 = 2
                bln退费重打 = gbln退费重打
            End If
            
            If (mbln附加费 And chkExtra.Value = 1) Then
                If bln记帐 = False Then
                    If dblCheckThreeMoney <> 0 Then
                        MsgBox "使用三方接口结算的挂号单据,不能将挂号费与" & mstr附加费 & "分开退!", vbInformation, gstrSysName
                        Exit Function
                    End If
                    
                    If mintInsure <> 0 Then
                        MsgBox "使用医保接口结算的挂号单据,不能将挂号费分开退!", vbInformation, gstrSysName
                        Exit Function
                    End If
                    
                    If mstr划价NO <> "" Then
                        MsgBox "挂号产生划价单时,不支持挂号费与" & mstr附加费 & "分别退!", vbInformation, gstrSysName
                        Exit Function
                    End If
                End If
                
                '本次退费金额计算.
                For i = 1 To vsfMoney.Rows - 1
                    curMoney = Val(vsfMoney.TextMatrix(i, 2)) + curMoney
                Next
                
                byt退费类型 = 3
                bln退费重打 = gbln退费重打
            End If
        End If
    Case 1
        If bln记帐 = False Then
            If dblCheckThreeMoney <> 0 Then
                MsgBox "使用三方接口结算的挂号单据,不能将病历费与挂号费分开退!", vbInformation, gstrSysName
                Exit Function
            End If
            
            If mintInsure <> 0 And MCPAR.不收病历费 = False Then
                MsgBox "使用医保接口结算的挂号单据,不能将挂号费分开退!", vbInformation, gstrSysName
                Exit Function
            End If
            
            If mstr划价NO <> "" Then
                MsgBox "挂号产生划价单时,不支持病历费与挂号费分别退!", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        
        '本次退费金额计算.
        For i = 1 To vsfMoney.Rows - 1
            curMoney = Val(vsfMoney.TextMatrix(i, 2)) + curMoney
        Next
        
        byt退费类型 = 2
        bln退费重打 = gbln退费重打
    Case 2
        If bln记帐 = False Then
            If dblCheckThreeMoney <> 0 Then
                MsgBox "使用三方接口结算的挂号单据,不能将挂号费与" & mstr附加费 & "分开退!", vbInformation, gstrSysName
                Exit Function
            End If
            
            If mintInsure <> 0 Then
                MsgBox "使用医保接口结算的挂号单据,不能将挂号费与" & mstr附加费 & "分开退!", vbInformation, gstrSysName
                Exit Function
            End If
            
            If mstr划价NO <> "" Then
                MsgBox "挂号产生划价单时,不支持挂号费与" & mstr附加费 & "分别退!", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        
        '本次退费金额计算.
        For i = 1 To vsfMoney.Rows - 1
            curMoney = Val(vsfMoney.TextMatrix(i, 2)) + curMoney
        Next
        
        byt退费类型 = 3
        bln退费重打 = gbln退费重打
    End Select
    
    bln医保原样退 = True
    If mintInsure <> 0 Then
        Call initInsurePara(lng病人ID)
        If bln记帐 = False Then
            If mstr个人帐户 = "" Then
                strSQL = "Select 名称 From 结算方式 Where 性质=3"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
                If Not rsTmp.EOF Then
                    mstr个人帐户 = Nvl(rsTmp!名称)
                End If
            End If
            bln医保原样退 = False
            For i = 1 To vsfPay.Rows - 1
                If (vsfPay.TextMatrix(i, 0) = mstr个人帐户) And vsfPay.TextMatrix(i, 0) <> "" And vsfPay.RowHidden(i) = False Then
                    bln医保原样退 = True
                End If
            Next i
            strAdvance = IIf(mstr个人帐户 <> "", mstr个人帐户, "个人帐户")
            If bln医保原样退 = True Then
                If gclsInsure.GetCapability(support门诊结算作废, , mintInsure, strAdvance) Then
                    strAdvance = ""     '向过程传入不允许退的结算方式,空表示全部允许
                End If
            End If
            If MCPAR.医保接口打印票据 Then
                 If zlGetInvoiceGroupUseID(lng领用ID) = False Then Exit Function
                 strInvoice = GetNextBill(lng领用ID)
            End If
        End If
    ElseIf bln记帐 = False Then
        Set rsOneCard1 = GetOneCardBalance(mlng结帐ID)
        
        If rsOneCard1.RecordCount > 0 Then
            If mbln包含病历费 And chk病历费.Value = 0 Then
                '不允许部分退
                MsgBox "使用一卡通接口进行扣款,不能将病历费与挂号费分开退!", vbInformation, gstrSysName
                Exit Function
            End If
            If mbln附加费 And chkExtra.Value = 0 Then
                '不允许部分退
                MsgBox "使用一卡通接口进行扣款,不能将病历费与" & mstr附加费 & "分开退!", vbInformation, gstrSysName
                Exit Function
            End If
            On Error Resume Next
            Set objICCard = CreateObject("zlICCard.clsICCard")
            On Error GoTo 0
            If objICCard Is Nothing Then
                MsgBox "一卡通接口创建失败,不能进行退费!请检查接口文件.", vbInformation, gstrSysName
                Exit Function
            End If
            strCardNo = objICCard.Read_Card(Me)
            If strCardNo = "" Then Exit Function
            If strCardNo <> rsOneCard1!单位帐号 Then
                MsgBox "当前卡号与扣款卡号不一致!不能进行退费.", vbInformation, gstrSysName
                Exit Function
            End If
                    
            If lngCard结帐ID <> 0 Then
                Set rsOneCard2 = GetOneCardBalance(lngCard结帐ID)
            End If
        End If
        '检查三方结算
        If Not mCurCardPay.objCard Is Nothing Then
            If mCurCardPay.objCard.接口序号 <> 0 Then
                If IsCheckCancelValied(lng结帐ID, lngCard结帐ID, cllBillBalance, dblThreeMoney, mCurCardPay.objCard.是否退款验卡) = False Then Exit Function
            End If
        End If
    End If
    
    If byt退费类型 = 0 Then
        '获取收回票据
        strSQL = _
        "   Select A.号码" & vbNewLine & _
        "   From 票据使用明细 A" & vbNewLine & _
        "   Where A.性质 = 1 And a.原因 <> 6 " & vbNewLine & _
        "           And A.打印id = (Select Max(ID) From 票据打印内容 Where 数据性质 = [2] And NO = [1])" & vbNewLine & _
        "Minus" & vbNewLine & _
        "Select A.号码" & vbNewLine & _
        "From 票据使用明细 A" & vbNewLine & _
        "Where A.性质 = 2 And a.原因 <> 6 " & vbNewLine & _
        "   And A.打印id = (Select Max(ID) From 票据打印内容 Where 数据性质 = [2] And NO = [1])" & vbNewLine & _
        "Order By 号码"
        Set rsInvoice = zlDatabase.OpenSQLRecord(strSQL, "获取收回票据", strNO, 4)
        Do While Not rsInvoice.EOF
            strBackInvoice = strBackInvoice & "," & rsInvoice!号码
            rsInvoice.MoveNext
        Loop
        If strBackInvoice <> "" Then strBackInvoice = Mid(strBackInvoice, 2)
    Else
        If gblnBill挂号 Then
            If frmReInvoice.ShowMe(Me, strNO, dblTotal, CDbl(curMoney), strBackInvoice, blnReprint) = False Then Exit Function
            If blnReprint = False Then bln退费重打 = False
        End If
    End If
    
    strBalance = ""
    str现金 = ""
    dbl现金 = 0
    With vsfPay
        For i = 1 To .Rows - 1
            If .RowData(i) = 1 And .TextMatrix(i, 0) <> "" Then
                str现金 = .TextMatrix(i, 0)
                dbl现金 = Val(.TextMatrix(i, 1))
                Exit For
            End If
        Next i
        dbl预交 = 0
        For i = 1 To .Rows - 1
            If .RowData(i) = 0 And .TextMatrix(i, 0) <> "" And Val(.TextMatrix(i, 1)) <> 0 Then
                dbl预交 = dbl预交 + Val(.TextMatrix(i, 1))
                Exit For
            End If
        Next i
        If str现金 = "" Then
            strSQL = "Select 名称 From 结算方式 Where 性质=1 Order By 缺省标志 Desc"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
            If rsTmp.EOF Then
                str现金 = "现金"
            Else
                str现金 = Nvl(rsTmp!名称)
            End If
        End If
        For i = 1 To .Rows - 1
            If .TextMatrix(i, 0) <> "" Then
                If .TextMatrix(i, 0) = mstr个人帐户 Then
                    If Val(.TextMatrix(i, 1)) <> 0 Then
                        If InStr(strAdvance, mstr个人帐户) <> 0 Then
                            dbl现金 = dbl现金 + Val(.TextMatrix(i, 1))
                        Else
                            strBalance = strBalance & "|" & .TextMatrix(i, 0) & "," & Val(.TextMatrix(i, 1)) & ",0"
                        End If
                    End If
                Else
                    If .RowData(i) = 7 Or .RowData(i) = 8 Then
                        strBalance = strBalance & "|" & mrsBillAdvance!结算方式 & "," & Val(.TextMatrix(i, 1)) & ",1"
                    Else
                        If .RowData(i) <> 0 And .TextMatrix(i, 0) <> "" And .TextMatrix(i, 0) <> str现金 Then
                            strBalance = strBalance & "|" & .TextMatrix(i, 0) & "," & Val(.TextMatrix(i, 1)) & ",0"
                        End If
                    End If
                End If
            End If
        Next i
        If str现金 <> "" And dbl现金 <> 0 Then
            strBalance = strBalance & "|" & str现金 & "," & dbl现金 & ",0"
        End If
        If strBalance <> "" Then strBalance = Mid(strBalance, 2)
    End With
        
    cmdOK.Enabled = False      '防止打印弹出设置打印机的非模态窗体及医保结算延迟
    On Error GoTo errH
    If mstr划价NO <> "" And bln记帐 = False Then
        strSQL = "zl_门诊划价记录_Delete('" & mstr划价NO & "')"
        zlAddArray cllPro, strSQL
    End If
    If strSQLCard <> "" Then zlAddArray cllPro, strSQLCard   '分离时退卡
    
    If mstrForceNote = "" And str原结算方式 <> "" And str结算方式 <> str原结算方式 Then
        strBackNote = objCard.名称 & "退现"
    Else
        strBackNote = mstrForceNote
    End If
    'zl_病人挂号记录_Delete
    strSQL = "zl_病人挂号记录_出诊_DELETE("
    '  单据号_In       门诊费用记录.No%Type,
    strSQL = strSQL & "'" & strNO & "',"
    '  操作员编号_In   门诊费用记录.操作员编号%Type,
    strSQL = strSQL & "'" & UserInfo.编号 & "',"
    '  操作员姓名_In   门诊费用记录.操作员姓名%Type,
    strSQL = strSQL & "'" & UserInfo.姓名 & "',"
    '  摘要_In         门诊费用记录.摘要%Type := Null, --预约取消时 填写 存放预约取消原因
    strSQL = strSQL & "" & IIf(Me.cbo备注.Text <> "", "'" & Me.cbo备注.Text & "'", " NULL ") & ","
    '  删除门诊号_In   Number := 0,
    strSQL = strSQL & "" & bytTogetherDo & ","
    '  非原样退结算_In Varchar2 := Null,
    If strAdvance <> "" Or str结算方式 <> str原结算方式 Then
        If strAdvance <> "" Then str原结算方式 = str原结算方式 & "," & strAdvance
        If Left(str原结算方式, 1) = "," Then str原结算方式 = Mid(str原结算方式, 2)
    End If
    strSQL = strSQL & IIf(str原结算方式 = "" Or bln记帐, "NULL", "'" & str原结算方式 & "'") & ","
    '  退费类型_In     In Number := 0, --0-全退 1-退挂号费 2-退病历费 3-退附加费 4-退挂号&病历费 5-退挂号&附加费
    strSQL = strSQL & "" & byt退费类型 & ","
    '  退指定结算_In   病人预交记录.结算方式%Type := Null
    strSQL = strSQL & IIf(str结算方式 = "" Or bln记帐, "NULL", "'" & str结算方式 & "'") & ","
    '  退号重用_In   Number := 1
    strSQL = strSQL & IIf(mTy_Para.blnReuseCancelNO, 1, 0) & ","
    '  结算方式_In   Varchar2 := Null
    strSQL = strSQL & "'" & strBalance & "',"
    '  退预交_In     病人预交记录.冲预交%Type := Null
    strSQL = strSQL & "" & ZVal(dbl预交) & ","
    strSQL = strSQL & "'" & strBackInvoice & "','"
    '  交易说明_In   病人预交记录.交易说明%Type := Null
    strSQL = strSQL & strBackNote & "')"
    zlAddArray cllPro, strSQL
    
    blnNotCommit = False
    '需要处理零费用结帐
    '退号
    Err = 0: On Error GoTo Errhand:
    blnTrans = True
    zlExecuteProcedureArrAy cllPro, Me.Caption, True
    If mintInsure <> 0 And Not (MCPAR.不收病历费 = True And mintCancel = 1) Then
        '68991
        '挂号收取方式(0或1)|挂号单号
        Dim strAdvanceTemp As String
        If bln记帐 Then strAdvanceTemp = "1|" & strNO
        If Not gclsInsure.RegistDelSwap(mlng结帐ID, mintInsure, strAdvanceTemp) Then
            gcnOracle.RollbackTrans: cmdOK.Enabled = True: Exit Function
        End If
        
        blnNotCommit = True
    ElseIf Not rsOneCard1 Is Nothing And bln记帐 = False Then
        If rsOneCard1.RecordCount > 0 Then
            If Not objICCard.ReturnSwap(Nvl(rsOneCard1!单位帐号), Nvl(rsOneCard1!医院编码), "" & rsOneCard1!结算号码, Nvl(rsOneCard1!金额)) Then
                gcnOracle.RollbackTrans
                MsgBox "一卡通退费交易调用失败！退费操作失败！", vbExclamation, gstrSysName
                cmdOK.Enabled = True: Exit Function
            End If
            If Not rsOneCard2 Is Nothing Then
                If rsOneCard2.RecordCount > 0 Then
                    If Not objICCard.ReturnSwap(Nvl(rsOneCard2!单位帐号), Nvl(rsOneCard2!医院编码), "" & rsOneCard2!结算号码, Nvl(rsOneCard2!金额)) Then
                        gcnOracle.RollbackTrans
                        MsgBox "一卡通退卡费交易调用失败！退费操作失败！", vbExclamation, gstrSysName
                        cmdOK.Enabled = True: Exit Function
                    End If
                End If
            End If
        End If
    End If
    '三方交易
    '退费
    If mCurCardPay.lng医疗卡类别ID <> 0 And bln记帐 = False And dblThreeMoney <> 0 Then
        If CallBackBalanceInterface(cllBillBalance, lng结帐ID, lngCard结帐ID, dblThreeMoney, cllUpdate, cllThreeIns, strErrMsg) = False Then
             gcnOracle.RollbackTrans
             If strErrMsg <> "" Then
                MsgBox strErrMsg, vbExclamation + vbOKOnly, gstrSysName
             Else
                MsgBox "调用第三方接口交易失败,此次退费操作失败!", vbExclamation + vbOKOnly, gstrSysName
            End If
             Exit Function
        End If
        If Not cllBillBalance Is Nothing Then
            zlExecuteProcedureArrAy cllUpdate, Me.Caption, True, True
        End If
    End If
    '单独处理退卡金额
    If mCurCardPay.lng医疗卡类别ID = 0 And bln记帐 = False And dblThreeMoney = 0 And strSQLCard <> "" Then
        strSQL = " Select A.实收金额 From 住院费用记录 A,病人预交记录 B,结算方式 C Where A.记录性质=5 And A.NO =(Select Max(NO) From 住院费用记录 where 结帐ID=[1] and  记录性质=5  )  And A.记录状态=2 " & _
                 "        And a.结帐ID=b.结帐ID And b.结算方式=c.名称 And c.性质 In (7,8)"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngCard结帐ID)
        If Not rsTmp.EOF Then
            If CallBackBalanceInterface(cllBillBalance, 0, lngCard结帐ID, -1 * Val(Nvl(rsTmp!实收金额)), cllUpdate, cllThreeIns, strErrMsg) = False Then
                gcnOracle.RollbackTrans
                If strErrMsg <> "" Then
                   MsgBox strErrMsg, vbExclamation + vbOKOnly, gstrSysName
                Else
                   MsgBox "调用第三方接口交易失败,此次退费操作失败!", vbExclamation + vbOKOnly, gstrSysName
                End If
                cmdOK.Enabled = True: Exit Function
            End If
            If Not cllBillBalance Is Nothing Then
                zlExecuteProcedureArrAy cllUpdate, Me.Caption, True, True
            End If
        End If
    End If
    
    gcnOracle.CommitTrans
    blnTrans = False
    '问题号:58567
    If Not cllThreeIns Is Nothing Then
        Err = 0: On Error GoTo ErrOthers:
        zlExecuteProcedureArrAy cllThreeIns, Me.Caption
    End If
    '继续执行
ResumeExecute:
    '问题:31634
    If mintInsure > 0 Then Call gclsInsure.BusinessAffirm(交易Enum.Busi_RegistDelSwap, True, mintInsure)
    cmdOK.Enabled = True      '防止打印弹出设置打印机的非模态窗体及医保结算延迟
    blnTrans = False
    If gblnBillPrint Then
        Err = 0: On Error Resume Next
        Call gobjBillPrint.zlEraseBill_Reg("'" & strNO & "'")
        If Err <> 0 Then
            Err = 0
        End If
        On Error GoTo errH
    End If
    If bln退费重打 And Not bln记帐 And (byt退费类型 <> 0 Or blnCardReprint) Then Call RePrintBill(Me, 2, strNO, lng结帐ID, mintInsure, MCPAR.医保接口打印票据, mstrUseType, bln退费重打 And Not bln记帐 And (byt退费类型 <> 0 Or blnCardReprint))
    
    If bln医保原样退 = True And strAdvance <> "" And mintInsure <> 0 And Not bln记帐 Then
        MsgBox "医保不支持[" & strAdvance & "]回退,退为" & str现金 & "." & vbCrLf & vbCrLf & _
            "退款共计:" & Format(GetCashMoney(cboNO.Tag), "0.00") & " 元.", vbInformation, gstrSysName
    End If
    mstr划价NO = "": vsfMoney.Tag = ""
    zlExcuteDelRegistered = True
    Exit Function
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    '问题:31634
    If mintInsure > 0 And blnNotCommit Then Call gclsInsure.BusinessAffirm(交易Enum.Busi_RegistDelSwap, False, mintInsure)
    Call SaveErrLog
    timPlan.Enabled = glngInterval > 0 And mbytInState = 0 And (mbytMode = 0 Or mbytMode = 1)
    cmdOK.Enabled = True
    Exit Function
ErrOthers:
  gcnOracle.RollbackTrans:
  If ErrCenter = 1 Then Resume
  GoTo ResumeExecute:
   Exit Function
Errhand:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
    cmdOK.Enabled = True
    Exit Function
End Function

Private Function CheckServeRange(intType As Integer, lng收费细目ID As Long, Optional intRow As Integer = 0) As Boolean
'功能:检查收费项目的服务对象,intType:0-门诊调用;1-住院调用
    Dim strSQL As String, rsTmp As ADODB.Recordset
    strSQL = "Select 名称,Nvl(服务对象,0) As 服务对象 From 收费项目目录 Where ID = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "CheckServeRange", lng收费细目ID)
    If rsTmp.EOF Then
        MsgBox "不能确定" & IIf(intRow = 0, "", "第" & intRow & "行") & "收费项目的服务对象,请检查项目是否正确录入!"
        Exit Function
    Else
        Select Case intType
        Case 0
            If Val(rsTmp!服务对象) = 2 Or Val(rsTmp!服务对象) = 0 Then
                MsgBox IIf(intRow = 0, "", "第" & intRow & "行") & "收费项目[" & rsTmp!名称 & "]不适用于门诊,请检查!"
                Exit Function
            End If
        Case 1
            If Val(rsTmp!服务对象) = 1 Or Val(rsTmp!服务对象) = 0 Then
                MsgBox IIf(intRow = 0, "", "第" & intRow & "行") & "收费项目[" & rsTmp!名称 & "]不适用于住院,请检查!"
                Exit Function
            End If
        Case Else
            If Val(rsTmp!服务对象) = 0 Then
                MsgBox IIf(intRow = 0, "", "第" & intRow & "行") & "收费项目[" & rsTmp!名称 & "]不适用于病人,请检查!"
                Exit Function
            End If
        End Select
    End If
    CheckServeRange = True
End Function

Private Function CheckInputValied() As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '功能：检查输入的有效性
    '返回：数据合法,,返回True,否则返回False
    '编制：刘兴洪
    '日期：2010-07-02 11:15:29
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim dtDate As Date, lngSN As Long, i As Long, j As Long
    Dim blnHave As Boolean, blnPrice As Boolean '建档病人存为划价单
    Dim dt预约  As Date, lng记录ID As Long, lng项目id As Long, rsTemp As ADODB.Recordset
    Dim blnCheckDat   As Boolean, lngTmp As Long
    Dim rsReserve As New ADODB.Recordset, strSQL As String, strErrInfo As String
    Dim bytMode As Byte, rsCheck As ADODB.Recordset, dat预约时间 As Date
    Dim strResult As String, bln专家号 As Boolean
    Dim dbl预交支付 As Double
    
    blnPrice = gblnPrice And Not mrsInfo Is Nothing And mbytMode = 0 And picBookingDate.Visible = False And mstrNoIn = ""
    dtDate = zlDatabase.Currentdate
    
    If vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("记录ID")) = "" Then
        MsgBox "当前未选择任何挂号安排,请先选择一个挂号安排后再继续!", vbInformation, gstrSysName
        Exit Function
    End If
    
    '82859:李南春,2015/4/8,病人基本信息调整
    '87876:李南春,2015/8/31,判断是不是新病人挂号
    With mobjfrmPatiInfo
        If Not mrsInfo Is Nothing And .mlng病人ID > 0 And mbln基本信息调整 And (.mstr年龄 & .mstr年龄单位 <> IIf(IsNumeric(txt年龄.Text), txt年龄.Text & cbo年龄单位.Text, txt年龄.Text) Or .mstr性别 <> NeedName(cbo性别.Text) Or .mstr姓名 <> txtPatient.Text Or _
            .mstr身份证号 <> txtIDCard.Text Or .mstr出生日期 <> txt出生日期.Text Or .mstr出生时间 <> txt出生时间.Text) Then
            If MsgBox("病人基本信息已发生改变，是否继续？", vbInformation + vbYesNo, gstrSysName) = vbNo Then
                '记录病人原始信息
                txtPatient.Text = .mstr姓名:  cbo性别.ListIndex = cbo.FindIndex(cbo性别, .mstr性别, True)
                txt年龄.Text = .mstr年龄: Call txt年龄_Validate(False)
                If .mstr年龄单位 <> "" Then cbo年龄单位.ListIndex = cbo.FindIndex(cbo年龄单位, .mstr年龄单位, True): cbo年龄单位.Visible = True: txt年龄.Width = 600
                txt出生日期.Text = IIf(.mstr出生日期 = "", "____-__-__", .mstr出生日期): txt出生时间.Text = IIf(.mstr出生时间 = "", "__:__", .mstr出生时间)
                txtIDCard.Text = .mstr身份证号
                .txt身份证号.Text = .mstr身份证号
                Exit Function
            Else
                '记录病人新的信息
                .mstr姓名 = txtPatient.Text: .mstr性别 = NeedName(cbo性别.Text)
                .mstr年龄 = txt年龄.Text: .mstr年龄单位 = NeedName(cbo年龄单位.Text)
                .mstr出生日期 = txt出生日期.Text: .mstr出生时间 = txt出生时间.Text
                .mstr身份证号 = txtIDCard.Text
            End If
        End If
    End With
    
    If txt门诊号.Enabled And txt门诊号.Visible And mintNOLength > 0 And mblnCheckNOValidity Then
    '如果手工输入了异常的门诊号则提示
        If Len(txt门诊号.Text) > mintNOLength + 1 Then
            MsgBox "注意,输入的门诊号过大,请确认是否输入正常!", vbInformation, gstrSysName
            txt门诊号.SetFocus
            txt门诊号.SelStart = 0: txt门诊号.SelLength = Len(txt门诊号.Text)
            Exit Function
        End If
    End If
    
    '检查单据数据有效性
    If txtPatient.Text = "" Then
        If picBookingDate.Visible Then        '预约挂号时必须要有病人信息
            MsgBox "预约挂号时必须输入病人信息。", vbInformation, gstrSysName
            If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
            Exit Function
        End If
        
        If txt门诊号.Text <> "" Then
            MsgBox "必须输入病人姓名！", vbInformation, gstrSysName
            txtPatient.SetFocus: Exit Function
        End If
    Else
        
        If CheckTextLength("姓名", txtPatient) = False Then Exit Function
        If CheckTextLength("年龄", txt年龄) = False Then Exit Function
        
        If mblnStructAdress Then
            If Not CheckStructAddr(padd家庭地址, padd家庭地址.MaxLength) Then Exit Function
            If Not CheckStructAddr(padd户口地址, padd户口地址.MaxLength) Then Exit Function
        Else
            If zlCommFun.ActualLen(cbo家庭地址.Text) > glngMax家庭地址 Then
                MsgBox "现住址输入过长，只允许输入" & glngMax家庭地址 & "个字符或" & glngMax家庭地址 \ 2 & "个汉字，请检查!", vbInformation, gstrSysName
                cbo家庭地址.SetFocus: Exit Function
            End If
            
            If zlCommFun.ActualLen(cbo户口地址.Text) > glngMax户口地址 Then
                MsgBox "户口地址输入过长，只允许输入" & glngMax户口地址 & "个字符或" & glngMax户口地址 \ 2 & "个汉字，请检查!", vbInformation, gstrSysName
                cbo户口地址.SetFocus: Exit Function
            End If
        End If
    
        If txt年龄.Enabled And txt年龄.Text = "" And Not (gblnAutoAddName And txtPatient.Text = "新病人") Then
            MsgBox "必须输入病人年龄！", vbInformation, gstrSysName
            txt年龄.SetFocus: Exit Function
        End If
        
        If mTy_Para.bln禁止输入年龄 Then
            '禁止输入年龄的情况,检查是否录入出生日期
            If txt出生日期.Enabled And IsDate(txt出生日期.Text) = False And Not (gblnAutoAddName And txtPatient.Text = "新病人") Then
                MsgBox "必须输入病人出生日期！", vbInformation, gstrSysName
                txt出生日期.SetFocus: Exit Function
            End If
            If mobjfrmPatiInfo.mobjPubPatient Is Nothing Then Exit Function
            If mobjfrmPatiInfo.mobjPubPatient.CheckPatiAge(Trim(txt年龄.Text) & IIf(cbo年龄单位.Visible, cbo年龄单位.Text, ""), _
                IIf(txt出生日期.Text = "____-__-__", "", txt出生日期.Text) & _
                IIf(txt出生时间.Text = "__:__", "", " " & txt出生时间.Text)) = False Then
                If txt出生日期.Enabled And txt出生日期.Visible Then txt出生日期.SetFocus
                Exit Function
            End If
        End If
        
        If cbo性别.Enabled And cbo性别.ListIndex = -1 Then
            MsgBox "必须输入病人性别！", vbInformation, gstrSysName
            cbo性别.SetFocus: Exit Function
            Exit Function
        End If
        '89242:李南春,2015/12/10,必须输入
        If mblnStructAdress Then
            If padd家庭地址.Visible And padd家庭地址.Enabled And padd家庭地址.Value = "" And gbln家庭地址 And Not mblnStation And Not (gblnAutoAddName And txtPatient.Text = "新病人") Then
                MsgBox "必须输入病人现住址！", vbInformation, gstrSysName
                If padd家庭地址.Enabled And padd家庭地址.Visible Then
                    padd家庭地址.SetFocus: Exit Function
                End If
            End If
        Else
            If cbo家庭地址.Visible And cbo家庭地址.Enabled And cbo家庭地址.Text = "" And gbln家庭地址 And Not mblnStation And Not (gblnAutoAddName And txtPatient.Text = "新病人") Then
                MsgBox "必须输入病人现住址！", vbInformation, gstrSysName
                If cbo家庭地址.Enabled And cbo家庭地址.Visible Then
                    cbo家庭地址.SetFocus: Exit Function
                End If
            End If
        End If
        If txt家庭电话.Visible And txt家庭电话.Enabled And txt家庭电话.Text = "" And gbln电话 And Not mblnStation And Not (gblnAutoAddName And txtPatient.Text = "新病人") Then
            MsgBox "必须输入病人联系电话！", vbInformation, gstrSysName
            If txt家庭电话.Enabled And txt家庭电话.Visible Then
                txt家庭电话.SetFocus: Exit Function
            End If
        End If
    End If
    
    '69026,冉俊明,2014-8-11,年龄有效性检查
    If txt年龄.Enabled And txt年龄.Visible And Trim(txt年龄.Text <> "") Then
        If mobjfrmPatiInfo.mobjPubPatient Is Nothing Then Exit Function
        If mobjfrmPatiInfo.mobjPubPatient.CheckPatiAge(Trim(txt年龄.Text) & IIf(cbo年龄单位.Visible, cbo年龄单位.Text, "")) = False Then
            txt年龄.SetFocus: Exit Function
        End If
    End If
    '必须建病案检查,预约时可以不管
    If mbytMode <> 1 And txt号别.Text <> "+" And mbln建病案 And txt门诊号.Text = "" Then
        MsgBox "使用当前号别时必须给病人建立门诊病案！", vbInformation, gstrSysName
        If txt门诊号.Enabled Then
            txt门诊号.SetFocus
        ElseIf txtPatient.Enabled And txtPatient.Text = "" Then
            txtPatient.SetFocus
        End If
        Exit Function
    End If
    
     '主要检查新病人这种方式
    If mintInsure = 0 And mbytInState = 0 And (mbytMode = 0 Or mbytMode = 2) And txtPatient.Text = "" Then
         '主要检查新病人这种方式
         If zlPatiCardCheck(1, 0, "", 1) = False Then
             Call ClearmobjfrmPatiInfoFace: ClearPatientInfo
             Set mrsInfo = Nothing
             If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
             Exit Function
         End If
     End If
    '医生检查
    If cbo医生.ListIndex = -1 And cbo医生.Enabled Then
        MsgBox "不能确定输入的医生,请重新输入或选择正确的医生！", vbInformation, gstrSysName
        If cbo医生.Enabled And cbo医生.Visible Then cbo医生.SetFocus
        Exit Function
    End If
    '134429：李南春，2019/1/12，检查当前出诊记录ID信息是否一致
    lng记录ID = Val(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("记录ID")))
    lng记录ID = IIf(txt号别.Text = "+", 0, lng记录ID)
    lng项目id = Val(Split(vsfPlan.TextMatrix(vsfPlan.Row, GetCol("IDS")), ",")(1))
    If lng记录ID <> 0 Then
        strSQL = "Select a.号码, b.项目id, b.科室id, b.医生姓名" & _
                " From 临床出诊号源 a, 临床出诊记录 b " & _
                " where a.Id = b.号源id And b.id = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng记录ID)
        If rsTemp.EOF Then
            MsgBox "号别信息错误，请重新选择。", vbInformation, gstrSysName
            Exit Function
        Else
            strErrInfo = ""
            If txt号别.Text <> Nvl(rsTemp!号码) Then strErrInfo = "、号别"
            If mlng挂号科室ID <> Val(Nvl(rsTemp!科室ID)) Then strErrInfo = strErrInfo & "、科室"
            If lng项目id <> Val(Nvl(rsTemp!项目ID)) Then strErrInfo = strErrInfo & "、挂号项目"
            If strErrInfo <> "" Then
                MsgBox "挂号信息(" & Mid(strErrInfo, 2) & ")不一致，请重新选择。", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End If
    
    If dtpAppointmentDate.Visible And (mbytMode = 1 Or chkBooking.Value = 1) Then '２7781
        dtDate = DateAdd("n", mTy_Para.lng预约限制时间, dtDate)
        Select Case mcustomTime
        Case t_普通:
            dt预约 = dtpAppointmentDate.Value
        Case t_时段:
            If Format(dtpAppointmentDate.Value, "yyyy-MM-dd") <> Format(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("挂号时间")), "yyyy-MM-dd") Then
                If Format(dtpAppointmentTime.Value, "hh:mm:ss") >= Format(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("挂号时间")), "hh:mm:ss") Then
                    dt预约 = CDate(Format(dtpAppointmentDate.Value - 1, "yyyy-MM-dd") & " " & Format(dtpAppointmentTime.Value, "hh:mm:ss"))
                Else
                    dt预约 = CDate(Format(dtpAppointmentDate.Value, "yyyy-MM-dd") & " " & Format(dtpAppointmentTime.Value, "hh:mm:ss"))
                End If
            Else
                dt预约 = CDate(Format(dtpAppointmentDate.Value, "yyyy-MM-dd") & " " & Format(dtpAppointmentTime.Value, "hh:mm:ss"))
            End If
        End Select
        Select Case mViewMode
        Case V_普通号分时段:
            If Format(dtpAppointmentDate.Value, "yyyy-MM-dd") & " " & Trim(Get时段(vsfList.Row, vsfList.Col, True, True)) < Format(dtDate, "yyyy-MM-dd hh:mm:ss") Then
                 blnCheckDat = True
            End If
        Case Else:
            If dt预约 < dtDate Then     '27781
                  blnCheckDat = True
            End If
        End Select
        If blnCheckDat Then
            MsgBox "当前预约时间,小于了" & Format(dtDate, "yyyy-mm-dd HH:MM") & " ,不能预约!"
             If mcustomTime = t_普通 Then
                    If dtpAppointmentDate.Enabled Then dtpAppointmentDate.SetFocus
             Else
                    If dtpAppointmentTime.Enabled Then
                        dtpAppointmentTime.SetFocus
                    ElseIf dtpAppointmentTime.Enabled Then
                        dtpAppointmentDate.SetFocus
                    End If
             End If
             Exit Function
        End If
        
        If dtpAppointmentTime.Enabled Then
            '问题:51408
            With vsfPlan
                lng记录ID = Val(.TextMatrix(.Row, .ColIndex("记录ID")))
            End With
            
            If Check有效时间段(lng记录ID, dt预约) = False Then
                  MsgBox "当前预约时间," & Format(dt预约, "yyyy-mm-dd HH:MM") & " ,不存在挂号安排或者已经被停诊!", vbOKOnly + vbInformation, gstrSysName
                  If dtpAppointmentDate.Enabled And dtpAppointmentDate.Visible Then dtpAppointmentDate.SetFocus
                  Exit Function
            End If
        End If
    End If
    
    '81103,冉俊明,2014-12-26,录入身份证号后,出生日期、年龄、性别的同步关联检查和调整
    If Trim(txtIDCard.Text) <> "" Then
        Dim strbirthday As String, strAge As String, strSex As String, strInfo As String
        If txtIDCard.Visible And txtIDCard.Enabled And Not mobjfrmPatiInfo.mobjPubPatient Is Nothing Then
            'CheckPatiIdcard(ByVal strIdcard As String, Optional strBirthday As String, _
            '    Optional strAge As String, Optional strSex As String, Optional strErrInfo As String) As Boolean
            '功能：身份证号码合法性校验
            '入参：strIdCard 身份证号码
            '出参：strBirthday  函数返回True为出生日期
            '         strAge 函数返回True为年龄
            '         strSex 函数返回True为性别
            '         strErrInfo 函数返回False为错误信息
            '返回：True/False  身份证合法返回True(可从strBirthday，strSex获取出生日期和性别)，
            '       否则返回False(可从strErrInfo获取详细错误信息)
            If mobjfrmPatiInfo.mobjPubPatient.CheckPatiIdcard(Trim(txtIDCard.Text), strbirthday, strAge, strSex, strErrInfo) Then
                If strSex <> NeedName(cbo性别.Text) Then strInfo = "性别"
                If strAge <> Trim(txt年龄.Text) & cbo年龄单位 Then strInfo = strInfo & IIf(strInfo = "", "年龄", "、年龄")
                
                If strInfo <> "" Then
                    If MsgBox("输入的" & strInfo & "与身份证号的" & strInfo & "不一致，" & _
                            "将根据身份证号修改" & strInfo & "，是否继续？", vbInformation + vbYesNo, gstrSysName) = vbYes Then
                        Call zlControl.CboLocate(cbo性别, strSex)
                        txt年龄.Text = ReCalcOld(CDate(strbirthday), cbo年龄单位)
                        txt出生日期.Text = Format(strbirthday, "yyyy-mm-dd")
                        Call txt出生日期_Validate(False)
                    Else
                        If txtIDCard.Enabled And txtIDCard.Visible Then txtIDCard.SetFocus
                        Exit Function
                    End If
                End If
            Else
                MsgBox strErrInfo, vbInformation, gstrSysName
                If txtIDCard.Enabled And txtIDCard.Visible Then txtIDCard.SetFocus
                Exit Function
            End If
        End If
    End If
    
    '费别检查
    If Not (mbytMode = 2 Or mbytMode = 0 And mbytInState = 0 And mstrNoIn <> "") And cbo费别.ListIndex = -1 Then
        MsgBox "不能确定病人的费别,不能挂号！", vbInformation, gstrSysName
        If cbo费别.Enabled And cbo费别.Visible Then cbo费别.SetFocus
        Exit Function
    End If
    
    If Not (mbytMode = 2 Or mbytMode = 0 And mbytInState = 0 And mstrNoIn <> "") And cbo费别.ItemData(cbo费别.ListIndex) = 2 And Not mrsInfo Is Nothing Then
        MsgBox "该病人不是新病人,不能使用仅限初诊的费别！", vbInformation, gstrSysName
        Call SetCboDefault(cbo费别): Exit Function
    End If
    
    If mbytMode = 1 Or chkBooking.Value = 1 Then
        If vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("序号控制")) <> "" Then
            If vsfList.Cell(flexcpForeColor, vsfList.Row, vsfList.Col) <> vbBlack Or vsfList.Cell(flexcpFontStrikethru, vsfList.Row, vsfList.Col) = True Then
                MsgBox "当前序号控制的号别全部序号均不可用,无法预约！", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End If
    
    '74550,冉俊明,2014-7-2,在病人来院就诊,医生在门诊医生站挂号时能够选择结算方式(包含性质为7的一卡通结算)
    If mbytMode <> 1 And (mblnStation And Not mblnStationPrice And cbo结算方式.Visible = True) Then
        If cbo结算方式.ListIndex = -1 And Not blnPrice Then
            MsgBox "不能确定挂号费用的结算方式,不能挂号！", vbInformation, gstrSysName
            If cbo结算方式.Enabled And cbo结算方式.Visible Then cbo结算方式.SetFocus
            Exit Function
        End If
    End If
    If mlngOutModeMC > 0 And cbo医疗类别.Visible Then
        If mobjfrmPatiInfo.txtPatiMCNO(0).Text <> "" Then
            If cbo医疗类别.ListIndex <= 0 Then
                MsgBox "请确定该医保病人的医疗类别！", vbInformation, gstrSysName
                If cbo医疗类别.Visible And cbo医疗类别.Enabled Then cbo医疗类别.SetFocus
                Exit Function
            End If
        ElseIf cbo医疗类别.ListIndex > 0 Then
            MsgBox "确定了医疗类别,但是未输入医保号！", vbInformation, gstrSysName
            If cmdMore.Enabled Then Call cmdMore_Click
            Exit Function
        End If
    End If
    If cbo付款方式.ListIndex = -1 And cbo付款方式.Enabled And cbo付款方式.Visible And cbo付款方式.Locked = False Then
        MsgBox "请选择病人的医疗付款方式!", vbInformation, gstrSysName
        cbo付款方式.SetFocus
        Exit Function
    End If
    If mstr社区号 <> "" Then
        If Trim(txt门诊号.Text) = "" Then
            MsgBox "已验证身份的社区病人要求建档,门诊号不能为空！", vbInformation, gstrSysName
            If txt门诊号.Enabled And txt门诊号.Visible Then txt门诊号.SetFocus
            Exit Function
        End If
    End If
    '检查挂号项目输入是否正确
    If Not (mbytMode = 2 Or mbytMode = 0 And mbytInState = 0 And mstrNoIn <> "") Then
        If txt号别.Text <> "+" Then
            If Trim(txt科室.Text) = "" Or Trim(txt号别.Text) = "" Then
                MsgBox "挂号项目未正确输入，请检查！", vbInformation, gstrSysName
                txt号别.SetFocus: Exit Function
            Else
                For i = 1 To vsfPlan.Rows - 1
                    If vsfPlan.TextMatrix(i, GetCol("号别")) = txt号别.Text Then
                        Exit For
                    End If
                Next
                If i = vsfPlan.Rows Then
                    MsgBox "挂号项目未正确输入，请检查！", vbInformation, gstrSysName
                    txt号别.SetFocus: Exit Function
                End If
            End If
        ElseIf mrsItems Is Nothing Then
            MsgBox "挂号项目未正确输入，请检查！", vbInformation, gstrSysName
            txt号别.SetFocus: Exit Function
        End If
    End If
    If cbo备注.Visible And cbo备注.Enabled Then
        If zlCommFun.ActualLen(cbo备注.Text) > 200 Then
            MsgBox "摘要内容过多，最多允许 " & 100 & " 个汉字或 " & 200 & " 个字符。", vbInformation, gstrSysName
            cbo备注.SetFocus: Exit Function
        End If
    End If
    
    '序号
    If txtSN.Visible Then
        lngSN = Val(txtSN.Text)
        
        If Trim(txtSN.Text) <> "" And Val(txtSN.Tag) <> Val(txtSN.Text) Then  '如果是接收预约时没有变则不用检查
            If Not IsNumeric(txtSN.Text) Then
                MsgBox "挂号序号要求是数字，请检查！", vbInformation, gstrSysName
               If txtSN.Enabled And txtSN.Visible Then txtSN.SetFocus
               Exit Function
            ElseIf vsfList.Visible Then
                
                For i = 0 To vsfList.Rows - 1
                    For j = 0 To vsfList.Cols - 1
                        If mViewMode = v_专家号 Then
                            If lngSN = Val(vsfList.TextMatrix(i, j)) Then blnHave = True: Exit For
                        ElseIf mViewMode = v_专家号分时段 Then
                            If lngSN = Val(Get时段(i, j, False)) Then blnHave = True: Exit For
                        End If
                    Next
                    If blnHave Then Exit For
                Next
                If Not blnHave Then
                    If InStr(mstrPrivs, ";加号;") <= 0 Then
                        MsgBox lngSN & "号超过最大限号数!你没有满号后继续挂号的权限.", vbInformation, gstrSysName
                        If txtSN.Visible And txtSN.Enabled Then txtSN.SetFocus: Exit Function
                    End If
                End If
            End If
        End If
        '68659,刘尔旋,2014-01-10,挂号时处理预留号与限号数的关系
        If mbytMode = 0 And vsfPlan.TextMatrix(vsfPlan.Row, GetCol("序号控制")) <> "" Then
            strSQL = "Select Count(1) As 预留数 From 临床出诊序号控制 Where 记录ID = [1] And 挂号状态 = 3 "
            Set rsReserve = zlDatabase.OpenSQLRecord(strSQL, "查询挂号预留数", vsfPlan.TextMatrix(vsfPlan.Row, GetCol("记录ID")))
            If Val(Nvl(rsReserve!预留数)) <> 0 Then
                With vsfPlan
                    If Val(.TextMatrix(.Row, GetCol("限号"))) <= Val(Nvl(rsReserve!预留数)) + Val(.TextMatrix(.Row, GetCol("已挂"))) Then
                        If InStr(mstrPrivs, ";加号;") = 0 Then
                            MsgBox "该号别已经没有剩余可用号!(其中有" & Val(Nvl(rsReserve!预留数)) & "个预留号被使用)你没有继续挂号的权限.", vbInformation, gstrSysName
                            CheckInputValied = False
                            Exit Function
                        Else
                            If MsgBox("该号别已经没有剩余可用号!(其中有" & Val(Nvl(rsReserve!预留数)) & "个预留号被使用)你是否要继续挂号?", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
                                CheckInputValied = False
                                Exit Function
                            End If
                        End If
                    End If
                End With
            End If
        End If
    End If
    '使用打折费别的检查
    If mblnNoneCut And Not (mbytMode = 2 Or mbytMode = 0 And mbytInState = 0 And mstrNoIn <> "") Then
        For i = 1 To vsfMoney.Rows - 1
            If Val(vsfMoney.TextMatrix(i, 2)) <> Val(vsfMoney.TextMatrix(i, 1)) Then
                MsgBox "你没有权限给病人使用当前的打折费别""" & NeedName(cbo费别.Text) & """，请选择其他不打折的费别。", vbInformation, gstrSysName
                If cbo费别.Enabled And cbo费别.Visible Then cbo费别.SetFocus
                Exit Function
            End If
        Next
    End If
    
    strSQL = "Select Zl_临床出诊限制_Check([1],[2],[3]) As 适用性检查 From Dual"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("记录ID"))), NeedName(cbo性别.Text), txt年龄.Text & IIf(cbo年龄单位.Visible, cbo年龄单位.Text, ""))
    If rsTemp.EOF Then
        MsgBox "当前选择的病人不适用该号别!", vbInformation, gstrSysName
        Exit Function
    Else
        If Val(Mid(Nvl(rsTemp!适用性检查), 1, 1)) <> 0 Then
            MsgBox "当前选择的病人不适用该号别!" & vbCrLf & "原因:" & Mid(Nvl(rsTemp!适用性检查), InStr(Nvl(rsTemp!适用性检查), "|") + 1), vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    '服务对象检查
    If Not mrsItems Is Nothing Then
        mrsItems.Filter = ""
        Do While Not mrsItems.EOF
            If Val(Nvl(mrsItems!项目ID)) <> 0 Then
                If CheckServeRange(0, Val(Nvl(mrsItems!项目ID))) = False Then Exit Function
            End If
            mrsItems.MoveNext
        Loop
        mrsItems.MoveFirst
    End If
    
    '********************************************
    ' 对专家号和分时段的这种情况
    ' 需要对有效时间进行限制
    '********************************************
    If mcustomTime = t_时段 Then
        If (mViewMode <> V_普通号 And mViewMode <> V_普通号分时段 And mbytMode = 1 And dtpAppointmentTime.Visible) Or (mbytMode = 0 And chkBooking.Value = 1 And chkBooking.Visible) Then
            If Check有效号别(vsfPlan.TextMatrix(vsfPlan.Row, _
                                            GetCol("号别")), CDate(Format(dtpAppointmentDate.Value, "yyyy-MM-dd") & " " & Format(dtpAppointmentTime.Value, "HH:mm:00 ")), True) = False Then
                Exit Function
            End If
        ElseIf mbytMode = 0 And mViewMode = v_专家号分时段 Then
            If vsfList.TextMatrix(vsfList.Row, vsfList.Col) <> "" Then
            '-----------------------------------------------
            '挂号 检查 时间是否在工作时间内
            '-----------------------------------------------
                If Format(CDate(Format(dtDate, "hh:mm:ss")), "hh:mm:ss") < Format(CDate(Get时段(vsfList.Row, vsfList.Col, True)), "hh:mm:ss") Then
                    If Check有效号别(vsfPlan.TextMatrix(vsfPlan.Row, _
                                                    GetCol("号别")), CDate(Format(dtDate, "yyyy-MM-dd") & " " & Format(dtpAppointmentTime.Value, "HH:mm:00 ")), False) = False Then
                        Exit Function
                    End If
                End If
            End If
        End If
    End If
    
    With vsfPay
        For i = 1 To .Rows - 1
            If .TextMatrix(i, 0) <> "" And .RowData(i) = 0 Then
                dbl预交支付 = Val(.TextMatrix(i, 1))
                Exit For
            End If
        Next i
    End With
    
    
    If Val(dbl预交支付) <> 0 Then
        mstr病人家属IDs = ""
        If Not zlDatabase.PatiIdentify(Me, glngSys, Nvl(mrsInfo!病人ID, 0), Val(dbl预交支付), mlngModul, 1, _
                                    IDKind.GetCurCard.接口序号, IIf(-1 * gdbl预存款消费验卡 >= Val(dbl预交支付), False, True), True, mstr病人家属IDs, _
                                    (gdbl预存款消费验卡 <> 0), (gdbl预存款消费验卡 = 2)) Then Exit Function
    End If
    
    If mbytMode >= 0 And mbytMode <= 2 And Not mrsInfo Is Nothing Then
        strSQL = "Select Zl_Fun_病人挂号记录_Check([1],[2],[3],[4],[5],[6]) As 检查结果 From Dual"
        Select Case mbytMode
            Case 0
                If mstrNoIn <> "" Then
                    bytMode = 2
                    dat预约时间 = CDate(Format(dtpAppointmentDate.Value, "yyyy-mm-dd"))
                Else
                    bytMode = mbytMode
                    If chkBooking.Value = 1 Then
                        dat预约时间 = CDate(Format(dtpAppointmentDate.Value, "yyyy-mm-dd"))
                    Else
                        dat预约时间 = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd"))
                    End If
                End If
            Case 1, 2
                bytMode = mbytMode
                dat预约时间 = CDate(Format(dtpAppointmentDate.Value, "yyyy-mm-dd"))
        End Select
        bln专家号 = vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("医生")) <> ""
        Set rsCheck = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, bytMode, Val(Nvl(mrsInfo!病人ID)), Trim(txt号别.Text), _
                                                Val(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("记录ID"))), dat预约时间, IIf(bln专家号, 1, 0))
        If Not rsCheck.EOF Then
            strResult = Nvl(rsCheck!检查结果)
            If Val(Mid(strResult, 1, 1)) <> 0 Then
                MsgBox Mid(strResult, 3), vbInformation, gstrSysName
                Exit Function
            End If
        Else
            MsgBox "有效性检查失败,无法继续！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    If CheckArangement() = False Then Exit Function
    
    If mbytMode = 2 Then
        If zlCheck限约或限号数(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("记录ID"))) = False Then Exit Function
    End If
    
    If Len(Trim(mobjfrmPatiInfo.txt密码.Text)) <= 0 And Len(Trim(mobjfrmPatiInfo.txt卡号.Text)) > 0 Then
        If mobjfrmPatiInfo.zl_Get设置默认发卡密码 = False Then
            Call cmdMore_Click
            Exit Function
        End If
    End If
    
    CheckInputValied = True
End Function

Private Function Check有效时间段(lng记录ID As Long, datTime As Date) As Boolean
    Dim strSQL As String, rsTemp As ADODB.Recordset
    If mViewMode = v_专家号分时段 Then Check有效时间段 = True: Exit Function
    With vsfPlan
        '序号控制分时段号,不检查序号时间是否在出诊记录时间内
        If .TextMatrix(.Row, .ColIndex("序号控制")) <> "" And Val(.TextMatrix(.Row, .ColIndex("分时段"))) = 1 Then
            strSQL = "Select 1 From 临床出诊记录 Where ID=[1] And [2] Between Nvl(停诊开始时间,To_Date('3000-01-01', 'yyyy-mm-dd')) And Nvl(停诊终止时间,To_Date('3000-01-01', 'yyyy-mm-dd')) "
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng记录ID, datTime)
            If rsTemp.EOF Then
                Check有效时间段 = True
            Else
                Check有效时间段 = False
            End If
            Exit Function
        End If
    End With
    strSQL = "Select 1 From 临床出诊记录 Where ID=[1] And [2] Between 开始时间 And 终止时间 "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng记录ID, datTime)
    
    If rsTemp.EOF Then
        Check有效时间段 = False
    Else
        strSQL = "Select 1 From 临床出诊记录 Where ID=[1] And [2] Between Nvl(停诊开始时间,To_Date('3000-01-01', 'yyyy-mm-dd')) And Nvl(停诊终止时间,To_Date('3000-01-01', 'yyyy-mm-dd')) "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng记录ID, datTime)
        If rsTemp.EOF Then
            Check有效时间段 = True
        Else
            Check有效时间段 = False
        End If
    End If
    
End Function

'检查安排序号数据是否合法
Private Function CheckArangement() As Boolean
    Dim str号别 As Long, strChkTime As String
    Dim lngSN As Long, i As Long, j As Long
    Dim blnExit As Boolean
    
    If mViewMode = V_普通号 Or mViewMode = v_专家号 Or mbytMode = 2 Then CheckArangement = True: Exit Function
     
    Select Case mViewMode
        Case V_普通号分时段
        '暂时不处理,以后有需求进行补充
        Case v_专家号分时段
            lngSN = Val(txtSN.Text)
            If lngSN = 0 Then
                If mTy_Para.bln严格按时段挂号 And InStr(mstrPrivs, ";加号;") = 0 Then
                    MsgBox "该号别的时段已经使用完成,不能再进行挂号!", vbInformation + vbOKOnly, gstrSysName
                    Exit Function
                End If
                CheckArangement = True: Exit Function
            End If
            If vsfList.TextMatrix(vsfList.Row, vsfList.Col) Like "加*" Then CheckArangement = True: Exit Function
            If lngSN = Val(Get时段(vsfList.Row, vsfList.Col)) Then CheckArangement = True: Exit Function
            With vsfList
                For i = 0 To .Rows - 1
                    For j = 1 To .Cols - 1
                       If .TextMatrix(i, j) <> "" Then
                            If lngSN = Val(Get时段(i, j, False)) Then
                               .Row = i: .Col = j
                                dtpAppointmentTime.Value = CDate(Get时段(i, j, True))
                                blnExit = True: Exit For
                            End If
                        End If
                    Next
                    If blnExit Then Exit For
                Next
            End With
        Case Else
        CheckArangement = True
        Exit Function
    End Select
    CheckArangement = True
    Exit Function
Errhand:
    If ErrCenter() = 1 Then
        Resume
    End If
    SaveErrLog
End Function

Private Function CheckPayStyleValied(ByRef lngRow As Long) As Boolean
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim int性质 As Integer, i As Integer
    If cbo结算方式.Text = "" Then
        If GetRegistMoney = 0 Then
            CheckPayStyleValied = True
            Exit Function
        Else
            Exit Function
        End If
    End If
    If cbo结算方式.Visible = False Then Exit Function
    
    If mbln连续挂号 Then
        For i = 1 To vsfPay.Rows - 1
            If vsfPay.TextMatrix(i, 0) = NeedName(cbo结算方式.Text) Then
                If Val(txt缴款.Text) <> 0 Then
                    If Val(txt缴款.Text) < Val(txt本次应缴.Text) Then
                        MsgBox "输入的缴款金额不足,请重新输入!", vbInformation, gstrSysName
                        If txt缴款.Enabled And txt缴款.Visible Then txt缴款.SetFocus
                        Exit Function
                    End If
                End If
                lngRow = i
                CheckPayStyleValied = True
                Exit Function
            End If
        Next i
    End If
    
    If Val(txt缴款.Text) = 0 And Val(txt本次应缴.Text) = 0 Then CheckPayStyleValied = True: Exit Function
    
    For i = 1 To vsfPay.Rows - 1
        If vsfPay.TextMatrix(i, 0) = NeedName(cbo结算方式.Text) And cbo结算方式.Enabled And ((Val(txt缴款.Text) <> 0 And Val(txt本次应缴.Text) = 0) Or (Val(txt缴款.Text) = 0 And Val(txt本次应缴.Text) <> 0) Or (Val(txt缴款.Text) <> 0 And Val(txt本次应缴.Text) <> 0)) Then
            If Val(vsfPay.TextMatrix(i, 1)) <> 0 Then
                MsgBox "已经存在" & NeedName(cbo结算方式.Text) & "的结算方式,不能再使用该结算方式支付!", vbInformation, gstrSysName
                If cbo结算方式.Enabled And cbo结算方式.Visible Then cbo结算方式.SetFocus
                Exit Function
            Else
                lngRow = i
                CheckPayStyleValied = True
                Exit Function
            End If
        End If
        If vsfPay.TextMatrix(i, 0) = "" Then lngRow = i
    Next i
    If lngRow = 0 Then lngRow = vsfPay.Rows: vsfPay.Rows = vsfPay.Rows + 1
    
    If NeedName(cbo结算方式.Text) = "预交金" Then
        If Val(txt缴款.Text) > Val(txt本次应缴.Text) Then
            MsgBox "使用预交金支付金额不能超过本次挂号金额!", vbInformation, gstrSysName
            Exit Function
        End If
        CheckPayStyleValied = True: Exit Function
    End If
    
    strSQL = "Select 性质 From 结算方式 Where 名称=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, NeedName(cbo结算方式.Text))
    If rsTemp.EOF Then
        strSQL = "Select 8 As 性质 From 医疗卡类别 Where 名称=[1] Union Select 8 As 性质 From 消费卡类别目录 Where 名称=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, NeedName(cbo结算方式.Text))
        If rsTemp.EOF Then
            MsgBox "不能确定当前选择的结算方式,请检查!", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    If txt缴款.Visible And txt缴款.Enabled And mTy_Para.byt缴款方式 = 2 Then
        If Val(txt本次应缴.Text) <> 0 And Val(txt缴款.Text) = 0 Then
            MsgBox "请输入缴款金额！", vbInformation, gstrSysName
            txt缴款.SetFocus
            Exit Function
        End If
    End If
    
    If Val(rsTemp!性质) <> 1 And Val(txt缴款.Text) > Val(txt本次应缴.Text) Then
        MsgBox "使用非现金结算方式时不能超过本次挂号金额!", vbInformation, gstrSysName
        Exit Function
    End If
    
    If Val(rsTemp!性质) = 3 Or Val(rsTemp!性质) = 8 Or Val(rsTemp!性质) = 7 Then
        For i = 1 To vsfPay.Rows - 1
            If (Val(vsfPay.RowData(i)) = 3 Or Val(vsfPay.RowData(i)) = 7 Or Val(vsfPay.RowData(i)) = 8) And Val(vsfPay.TextMatrix(i, 1)) <> 0 Then
                MsgBox "目前只允许一种接口的支付方式,不能再使用" & NeedName(cbo结算方式.Text) & "!", vbInformation, gstrSysName
                Exit Function
            End If
        Next i
    End If
    CheckPayStyleValied = True
End Function

Private Function Get性质(ByVal str支付方式 As String, ByRef str结算方式 As String) As Integer
    On Error GoTo errHandle
    Dim strSQL As String, rsTemp As ADODB.Recordset
    If str支付方式 = "预交金" Then Get性质 = 0: str结算方式 = str支付方式: Exit Function
    strSQL = "Select 性质 From 结算方式 Where 名称=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str支付方式)
    If Not rsTemp.EOF Then
        Get性质 = Val(rsTemp!性质)
        str结算方式 = str支付方式
    Else
        strSQL = "Select 结算方式 From 医疗卡类别 Where 名称=[1] Union Select 结算方式 From 消费卡类别目录 Where 名称=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str支付方式)
        If Not rsTemp.EOF Then
            Get性质 = 8
            str结算方式 = Nvl(rsTemp!结算方式)
        End If
    End If
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub RestorePay()
    Dim i As Integer
    Dim j As Integer
    Dim dblDiff As Double
    With vsfPay
        For i = 1 To .Rows - 1
            If (.RowData(i) = 1 Or .RowData(i) = 2) And .TextMatrix(i, 0) = NeedName(cbo结算方式.Text) Then
                If Val(.TextMatrix(i, 7)) = 0 Then
                    txt本次应缴.Text = Format(Val(.TextMatrix(i, 1)), "0.00")
                    For j = 0 To .Cols - 1
                        .TextMatrix(i, j) = ""
                    Next j
                    .RowData(i) = ""
                Else
                    txt本次应缴.Text = Format(Val(.TextMatrix(i, 1)), "0.00")
                    .TextMatrix(i, 1) = Format(.TextMatrix(i, 7), "0.00")
                End If
            End If
        Next i
    End With
    mbln连续挂号 = mblnPre连续
End Sub

Private Function PrivCheck() As Boolean
    '挂号权限检查
    '挂免费号以及挂收费号的检查
    Dim dblMoney As Double
    Dim i As Integer
    
    On Error GoTo Errhand
    If mbytMode <> 0 Then PrivCheck = True: Exit Function
    If zlStr.IsHavePrivs(mstrPrivs, "挂免费号") And zlStr.IsHavePrivs(mstrPrivs, "挂收费号") Then PrivCheck = True: Exit Function
    
    '统计挂号项目金额
    If Not mrsItems Is Nothing Then
        For i = 1 To mrsItems.RecordCount
            dblMoney = 0
            If Not mrsInComes Is Nothing Then
                mrsInComes.Filter = "项目ID=" & mrsItems!项目ID
                Do While Not mrsInComes.EOF
                    dblMoney = dblMoney + Val(Nvl(mrsInComes!应收))
                    mrsInComes.MoveNext
                Loop
            End If
            Exit For
        Next
    End If
        
    If zlStr.IsHavePrivs(mstrPrivs, "挂免费号") = False Then
        If RoundEx(dblMoney, 5) = 0 Then
            MsgBox "你没有挂免费号的权限，不能为该病人挂当前号别！", vbInformation, gstrSysName
            Exit Function
        End If
    ElseIf zlStr.IsHavePrivs(mstrPrivs, "挂收费号") = False Then
        If RoundEx(dblMoney, 5) <> 0 Then
            MsgBox "你没有挂收费号的权限，不能为该病人挂当前号别！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    PrivCheck = True
    Exit Function
Errhand:
    If ErrCenter() = 1 Then
        Resume
    End If
    SaveErrLog
End Function

Private Sub SaveInvoiceNotify(ByVal blnPrice As Boolean, ByRef blnSlipPrint As Boolean, ByRef blnNoPrint As Boolean, _
                                ByRef blnPrintBooking As Boolean, ByRef blnCodePrint As Boolean)
    If mbytMode = 0 Or mbytMode = 2 Then
        '挂号及挂号接收
        Select Case Val(zlDatabase.GetPara("挂号凭条打印方式", glngSys, mlngModul))
            Case 0    '不打印
                blnSlipPrint = False
            Case 1    '自动打印
                If InStr(mstrPrivs, ";挂号凭条打印;") > 0 Then
                    blnSlipPrint = True
                Else
                    blnSlipPrint = False
                    MsgBox "你没有挂号凭条打印的权限，请联系管理员！", vbInformation, gstrSysName
                End If
            Case 2    '选择打印
                If MsgBox("要打印挂号凭条吗？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                    If InStr(mstrPrivs, ";挂号凭条打印;") > 0 Then
                        blnSlipPrint = True
                    Else
                        blnSlipPrint = False
                        MsgBox "你没有挂号凭条打印的权限，请联系管理员！", vbInformation, gstrSysName
                    End If
                Else
                    blnSlipPrint = False
                End If
        End Select
    End If
    
    If blnPrice Then
        blnNoPrint = True
        If mbytMode = 1 And mblnStation And InStr(1, gstrPrivsStation, ";预约挂号单;") > 0 Then    '医生站调用
            Select Case Val(zlDatabase.GetPara("预约挂号单打印方式", glngSys, 1260))    '使用医生站的相关参数
            Case 0    '不打印
            Case 1    '自助动打印
                blnPrintBooking = True
            Case 2    '选择打印
                If MsgBox("要打印挂号预约单吗？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                    blnPrintBooking = True
                End If
            End Select
        End If
    ElseIf mbytMode <> 1 Then
        If mRegistFeeMode = EM_RG_记帐 Then
            blnNoPrint = True
        Else
            If Not gblnPrintFree Then blnNoPrint = (GetRegistMoney(False) = 0)
            
            If Not blnNoPrint And txt号别.Text = "+" And Not mblnAddCardItem And gbytInvoice <> 0 Then
                If MsgBox("当前病人只购买病历，要打印票据吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    blnNoPrint = True
                End If
            End If
            If Not blnNoPrint Then
                If gbytInvoice = 0 Then
                    blnNoPrint = True
                ElseIf gbytInvoice = 2 Then
                    If Not (txt号别.Text = "+" And Not mblnAddCardItem) Then    '前面已提示过了,不再提示
                        If MsgBox("要打印挂号票据吗？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                            blnNoPrint = True
                        End If
                    End If
                End If
            End If
        End If
    ElseIf mbytMode = 1 Then
        Select Case Val(zlDatabase.GetPara("预约挂号单打印方式", glngSys, mlngModul))
        Case 0    '不打印
        Case 1    '自助动打印
            blnPrintBooking = True
        Case 2    '选择打印
            If MsgBox("要打印挂号预约单吗？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                blnPrintBooking = True
            End If
        End Select
        blnNoPrint = True
    End If
    
    If Not mblnStation And mbytMode <> 1 Then
        Select Case gByt打印病人条码
        Case 0: blnCodePrint = False
        Case 1: blnCodePrint = True
        Case 2:
               If MsgBox("是否需要打印病人条码？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    blnCodePrint = True
               Else
                    blnCodePrint = False
               End If
        End Select
    End If
End Sub

Private Sub Get时间(ByVal Datsys As Date, ByVal lngSN As Long, ByVal bln追加时段 As Boolean, ByRef str登记时间 As String, _
                    ByRef str发生时间 As String, ByRef dat发生时间 As Date, ByRef bln达到限号数 As Boolean)
    '在获取了可用序号后  才对发生时间进行处理
    str登记时间 = "To_Date('" & Format(Datsys, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
    If mcustomTime = t_时段 Then    '该时段代表只要所有安排中有一个安排设置了时段该mcustomTime的值都被设置为了t_时段（也就是说该条件成立）
        If dtpAppointmentTime.Visible = True And mbytMode <> 2 Then
            If picBookingDate.Visible And dtpAppointmentTime.Visible Then
                If lngSN <> 0 And (mViewMode = V_普通号分时段 Or mViewMode = v_专家号分时段) Then
                    mrs时间段.Filter = "序号=" & lngSN
                    If Not mrs时间段.EOF Then
                        str发生时间 = "To_Date('" & Format(mrs时间段!详细开始时间, "yyyy-MM-dd hh:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                        dat发生时间 = CDate(Format(mrs时间段!详细开始时间, "yyyy-MM-dd hh:mm:ss"))
                    Else
                        str发生时间 = "To_Date('" & Format(dtpAppointmentDate.Value, "yyyy-MM-dd") & " " & Format(dtpAppointmentTime.Value, "HH:mm:00 ") & "','YYYY-MM-DD HH24:MI:SS')"
                        dat发生时间 = CDate(Format(dtpAppointmentDate.Value, "yyyy-MM-dd") & " " & Format(dtpAppointmentTime.Value, "HH:mm:00 "))
                    End If
                    mrs时间段.Filter = ""
                Else
                    str发生时间 = "To_Date('" & Format(dtpAppointmentDate.Value, "yyyy-MM-dd") & " " & Format(dtpAppointmentTime.Value, "HH:mm:00 ") & "','YYYY-MM-DD HH24:MI:SS')"
                    dat发生时间 = CDate(Format(dtpAppointmentDate.Value, "yyyy-MM-dd") & " " & Format(dtpAppointmentTime.Value, "HH:mm:00 "))
                End If
            Else
                str发生时间 = "To_Date('" & Format(Datsys, "yyyy-MM-dd") & " " & Format(dtpAppointmentTime.Value, "HH:mm:00 ") & "','YYYY-MM-DD HH24:MI:SS')"
                dat发生时间 = CDate(Format(Datsys, "yyyy-MM-dd") & " " & Format(dtpAppointmentTime.Value, "HH:mm:00 "))
            End If
        ElseIf picBookingDate.Visible Or mbytMode = 2 Or (mbytMode = 0 And mbytInState = 0 And mstrNoIn <> "") Then
            str发生时间 = "To_Date('" & Format(dtpAppointmentDate.Value, "yyyy-MM-dd HH:mm:00") & "','YYYY-MM-DD HH24:MI:SS')"
            dat发生时间 = CDate(Format(dtpAppointmentDate.Value, "yyyy-MM-dd HH:mm:00"))
        Else
            str发生时间 = str登记时间
            dat发生时间 = Datsys
            If mbytMode = 0 Then
                If Format(dat发生时间, "yyyy-mm-dd hh:mm:ss") < Format(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("挂号时间")), "yyyy-MM-dd hh:mm:ss") Then
                    dat发生时间 = CDate(Format(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("挂号时间")), "yyyy-MM-dd hh:mm:ss"))
                    str发生时间 = "To_Date('" & Format(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("挂号时间")), "yyyy-MM-dd hh:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                End If
            End If
        End If
        
        If vsfList.Row < vsfList.Rows And vsfList.Col < vsfList.Cols Then
            If mbytMode = 0 And mViewMode = v_专家号分时段 And dtpAppointmentTime.Visible = False And mstrNoIn = "" Then
                If vsfList.TextMatrix(vsfList.Row, vsfList.Col) <> "" Then
                    If lngSN <> 0 Then
                        mrs时间段.Filter = "序号=" & lngSN
                        If Not mrs时间段.EOF Then
                            str发生时间 = "To_Date('" & Format(mrs时间段!详细开始时间, "yyyy-MM-dd hh:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                            dat发生时间 = CDate(Format(mrs时间段!详细开始时间, "yyyy-MM-dd hh:mm:ss"))
                        Else
                            If Format(Datsys, "hh:mm:ss") < Format(dtpAppointmentTime.Value, "hh:mm:ss") Then
                                str发生时间 = "To_Date('" & Format(Datsys, "yyyy-MM-dd") & " " & Format(dtpAppointmentTime.Value, "hh:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                                dat发生时间 = CDate(Format(Datsys, "yyyy-MM-dd") & " " & Format(dtpAppointmentTime.Value, "hh:mm:ss"))
                            End If
                        End If
                        mrs时间段.Filter = ""
                    Else
                        If Format(Datsys, "hh:mm:ss") < Format(dtpAppointmentTime.Value, "hh:mm:ss") Then
                            str发生时间 = "To_Date('" & Format(Datsys, "yyyy-MM-dd") & " " & Format(dtpAppointmentTime.Value, "hh:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                            dat发生时间 = CDate(Format(Datsys, "yyyy-MM-dd") & " " & Format(dtpAppointmentTime.Value, "hh:mm:ss"))
                        End If
                    End If
                End If
            End If
        End If
        If Not mrs时间段 Is Nothing And dtpAppointmentTime.Visible = False And (mViewMode = v_专家号分时段 Or mViewMode = V_普通号分时段) Then
            mrs时间段.MoveLast
            With vsfPlan
                bln达到限号数 = (Val(.TextMatrix(.Row, .ColIndex("限号"))) - (Val(.TextMatrix(.Row, .ColIndex("已挂"))) + Val(.TextMatrix(.Row, .ColIndex("已约"))) - Get失约号(.TextMatrix(.Row, .ColIndex("号别")), Datsys))) <= 0
            End With
            If bln追加时段 Or mbln加号 Or _
                (CDate(CStr(DatePart("h", CStr(mrs时间段!开始时间))) & ":" & CStr(DatePart("n", CStr(mrs时间段!开始时间))) & ":" & CStr(DatePart("s", CStr(mrs时间段!开始时间)))) <= CDate(Format(CStr(DatePart("h", CStr(Datsys))) & ":" & CStr(DatePart("n", CStr(Datsys))) & ":" & CStr(DatePart("s", CStr(Datsys))), "hh:mm:ss")) And bln达到限号数 = False) Then
                If CDate(CStr(DatePart("h", CStr(mrs时间段!结束时间))) & ":" & CStr(DatePart("n", CStr(mrs时间段!结束时间))) & ":" & CStr(DatePart("s", CStr(mrs时间段!结束时间)))) > CDate(Format(CStr(DatePart("h", CStr(Datsys))) & ":" & CStr(DatePart("n", CStr(Datsys))) & ":" & CStr(DatePart("s", CStr(Datsys))), "hh:mm:ss")) Then
                    str发生时间 = "To_Date('" & Format(Datsys, "yyyy-MM-dd") & " " & Format(mrs时间段!结束时间, "hh:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                    dat发生时间 = CDate(Format(Datsys, "yyyy-MM-dd") & " " & Format(mrs时间段!结束时间, "hh:mm:ss"))
                Else
                    str发生时间 = "To_Date('" & Format(Datsys, "yyyy-MM-dd") & " " & Format(Datsys, "hh:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                    dat发生时间 = CDate(Format(Datsys, "yyyy-MM-dd") & " " & Format(Datsys, "hh:mm:ss"))
                End If
            End If
        End If
    Else    '该分支代表当所有安排中没有一个设置了时段的情况
        If picBookingDate.Visible Or mbytMode = 2 Or (mbytMode = 0 And mbytInState = 0 And mstrNoIn <> "") Then
            str发生时间 = "To_Date('" & Format(dtpAppointmentDate.Value, "yyyy-MM-dd HH:mm:00") & "','YYYY-MM-DD HH24:MI:SS')"
            dat发生时间 = CDate(Format(dtpAppointmentDate.Value, "yyyy-MM-dd HH:mm:00"))
        Else
            str发生时间 = str登记时间
            dat发生时间 = Datsys
        End If
    End If
End Sub
Private Function SaveRegister_预约接收(ByVal lng病人ID As Long, ByVal dtSysDate As Date, str门诊号 As String, ByVal str卡号 As String, _
    ByVal rsCardFee As ADODB.Recordset, ByVal str发生时间 As String, ByVal str登记时间 As String, _
    ByVal blnPrice As Boolean, ByVal blnNoPrint As Boolean, ByVal lngSN As Long, _
    ByVal str结算方式 As String, cur现金 As Currency, cur卡费 As Currency, cur预交 As Currency, cur个帐 As Currency, _
    ByRef lng结帐ID As Long, _
    ByRef cllPro As Collection, ByRef cllProAfter As Collection, ByRef lngCard结帐ID As Long, ByRef bln发卡 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:预约接收,不以接收时的价格为准
    '入参:blnPrice-是否存为划价单
    '     blnNoPrint-发票打印标志(true-不打印,false-打印)
    '     lng结帐ID-结帐ID
    '出参:cllPro-返回数据保存集
    '     cllProAfter-最后完成执行的相关SQL集
    '     bln发卡-是否当前发卡
    '     lngCard结帐ID-返回卡数据的结帐ID
    '     lng结帐ID-结帐ID
    '返回:成功获取数据保存集,返回true,否则返回False
    '编制:刘兴洪
    '日期:2017-11-06 11:01:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim k As Integer, i As Integer, j As Integer, int价格父号 As Integer
    Dim blnHaveBookFee As Boolean '是否存病历费
    Dim dblTotalRegFee As Double
    Dim str划价NO As String, str年龄 As String, strRoom As String
    Dim strNO As String, str费别 As String
    Dim lng序号 As Long
    Dim strSQL As String, dblTemp As Double
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    
    '115168:李南春，2017/12/13，保存发卡的医疗卡类型
    If mCurSendCard.lng卡类别ID = 0 Then mCurSendCard = gCurSendCard
    If cllPro Is Nothing Then Set cllPro = New Collection
    If cllProAfter Is Nothing Then Set cllProAfter = New Collection
    '非预约接收处理，直接返回True
    If Not (mbytMode = 2 Or mbytMode = 0 And mbytInState = 0 And mstrNoIn <> "") Then SaveRegister_预约接收 = True: Exit Function
    If mlng记录ID <> 0 And mlng记录ID <> Val(vsfPlan.TextMatrix(vsfPlan.Row, GetCol("记录ID"))) Then
        MsgBox "未找到预约单据的出诊排班记录，不能接收单据。", vbInformation, gstrSysName
        Exit Function
    End If
    '获取分诊诊室
    If mbytMode <> 1 And txt号别.Text <> "+" And vsfPlan.TextMatrix(vsfPlan.Row, GetCol("分诊")) <> "" Then  '预约时不分诊
         strRoom = GetRoom(vsfPlan.TextMatrix(vsfPlan.Row, GetCol("记录ID")))
    End If

    str费别 = NeedName(cbo费别.Text)
    str年龄 = Trim(txt年龄.Text)
    If IsNumeric(str年龄) Then str年龄 = str年龄 & cbo年龄单位.Text
    strNO = cboNO.Text
    If mTy_Para.bln预约接收确定挂号费 Then  '预约接受，以新价格为准
        If blnPrice Then
            dblTotalRegFee = GetRegistMoney(True, False)
            '挂号费存为零且保存为划价单，才产生划价NO
            If dblTotalRegFee <> 0 Then str划价NO = zlDatabase.GetNextNo(13)
        End If
        mrsItems.Filter = ""
        k = 1: mrsItems.MoveFirst
        For i = 1 To mrsItems.RecordCount
            int价格父号 = k
            mrsInComes.Filter = "项目ID=" & mrsItems!项目ID
            For j = 1 To mrsInComes.RecordCount
                If mrsItems!性质 = 4 Then   '读费用集时已限制仅有一行,不支持设置多个收入项目,为了保持与就诊卡管理中一致
                    '
                Else
                    'Zl_病人预约挂号记录_Update
                    strSQL = "Zl_病人预约挂号记录_Update("
                    strSQL = strSQL & "'" & strNO & "',"
                    strSQL = strSQL & "" & k & ","
                    strSQL = strSQL & "" & IIf(int价格父号 = k, "NULL", int价格父号) & ","
                    strSQL = strSQL & "" & IIf(mrsItems!性质 = 2, 1, "NULL") & ","
                    strSQL = strSQL & "'" & mrsItems!类别 & "',"
                    strSQL = strSQL & "'" & mrsItems!项目ID & "',"
                    strSQL = strSQL & "" & Val(Nvl(mrsItems!数次)) & ","
                    strSQL = strSQL & "" & Val(Nvl(mrsInComes!单价)) & ","
                    strSQL = strSQL & "" & Val(Nvl(mrsInComes!收入项目ID)) & ","
                    strSQL = strSQL & "'" & Trim(Nvl(mrsInComes!收据费目)) & "',"
                    strSQL = strSQL & "" & IIf(blnPrice And dblTotalRegFee <> 0, 0, mrsInComes!应收) & ","
                    strSQL = strSQL & "" & IIf(blnPrice And dblTotalRegFee <> 0, 0, mrsInComes!实收) & ","
                    strSQL = strSQL & "" & IIf(mrsItems!性质 = 3, 1, IIf(mrsItems!性质 = 4, 2, 0)) & ","
                    strSQL = strSQL & "" & ZVal(Nvl(mrsItems!保险大类id, 0)) & ","
                    strSQL = strSQL & "" & ZVal(Nvl(mrsItems!保险项目否, 0)) & ","
                    strSQL = strSQL & "" & ZVal(Nvl(mrsInComes!统筹金额, 0)) & ","
                    strSQL = strSQL & "'" & Trim(Nvl(mrsItems!保险编码)) & "',"
                    strSQL = strSQL & "" & mlng挂号科室ID & ","
                    strSQL = strSQL & "" & IIf(mrsItems!执行科室ID = 0, mlng挂号科室ID, mrsItems!执行科室ID) & ","
                    strSQL = strSQL & "" & "NULL" & ","
                    strSQL = strSQL & "" & IIf(mrsItems!性质 = 1 Or mrsItems!性质 = 2, 1, 0) & ")" '挂号项目才传True
                    
                    Call zlAddArray(cllPro, strSQL)
                    If blnPrice And dblTotalRegFee <> 0 Then
                        strSQL = _
                        "zl_门诊划价记录_Insert('" & str划价NO & "'," & k & "," & lng病人ID & ",NULL," & _
                                 IIf(str门诊号 = "", "NULL", str门诊号) & ",'" & NeedCode(cbo付款方式.Text) & "'," & _
                                 "'" & txtPatient.Text & "','" & NeedName(cbo性别.Text) & "','" & str年龄 & "'," & _
                                 "'" & str费别 & "',NULL," & mlng挂号科室ID & "," & _
                                 IIf(mblnStation, mlng挂号科室ID, UserInfo.部门ID) & ",'" & UserInfo.姓名 & "'," & IIf(mrsItems!性质 = 2, 1, "NULL") & "," & _
                                 mrsItems!项目ID & ",'" & mrsItems!类别 & "','" & mrsItems!计算单位 & "'," & _
                                 "NULL,1," & mrsItems!数次 & ",NULL," & IIf(mrsItems!执行科室ID = 0, mlng挂号科室ID, mrsItems!执行科室ID) & "," & IIf(int价格父号 = k, "NULL", int价格父号) & "," & _
                                 mrsInComes!收入项目ID & ",'" & mrsInComes!收据费目 & "'," & mrsInComes!单价 & "," & _
                                 mrsInComes!应收 & "," & mrsInComes!实收 & "," & str发生时间 & "," & str登记时间 & ",NULL,'" & UserInfo.姓名 & "','挂号:" & strNO & "')"
                        Call zlAddArray(cllPro, strSQL)
                    End If
                End If
                k = k + 1
                mrsInComes.MoveNext
            Next
            mrsItems.MoveNext
        Next
    Else    '以预约时的价格为准
        If blnPrice Then
            dblTotalRegFee = GetRegistMoney(True, False)
            '挂号费存为零且保存为划价单，才产生划价NO
            If dblTotalRegFee <> 0 Then str划价NO = zlDatabase.GetNextNo(13)
        End If
        blnHaveBookFee = False
        mrsBill.Sort = "序号 "
        mrsBill.MoveFirst
        lng序号 = 0
        Do While Not mrsBill.EOF
            'Zl_病人预约挂号记录_Update
            strSQL = "Zl_病人预约挂号记录_Update("
            '  单据号_In     门诊费用记录.NO%Type,
            strSQL = strSQL & "'" & mrsBill!NO & "',"
            '  序号_In       门诊费用记录.序号%Type,
            strSQL = strSQL & "" & mrsBill!序号 & ","
            '  价格父号_In   门诊费用记录.价格父号%Type,
            strSQL = strSQL & "" & IIf(Val(Nvl(mrsBill!价格父号)) = 0, "NULL", mrsBill!价格父号) & ","
            '  从属父号_In   门诊费用记录.从属父号%Type,
            strSQL = strSQL & "" & IIf(Val(Nvl(mrsBill!从属父号)) = 0, "NULL", mrsBill!从属父号) & ","
            '  收费类别_In   门诊费用记录.收费类别%Type,
            strSQL = strSQL & "'" & mrsBill!收费类别 & "',"
            '  收费细目id_In 门诊费用记录.收费细目id%Type,
            strSQL = strSQL & "'" & mrsBill!收费细目ID & "',"
            '  数次_In       门诊费用记录.数次%Type,
            strSQL = strSQL & "" & Val(Nvl(mrsBill!数次)) & ","
            '  标准单价_In   门诊费用记录.标准单价%Type,
            strSQL = strSQL & "" & Val(Nvl(mrsBill!标准单价)) & ","
            '  收入项目id_In 门诊费用记录.收入项目id%Type,
            strSQL = strSQL & "" & Val(Nvl(mrsBill!收入项目ID)) & ","
            '  收据费目_In   门诊费用记录.收据费目%Type,
            strSQL = strSQL & "'" & Trim(Nvl(mrsBill!收据费目)) & "',"
            '  应收金额_In   门诊费用记录.应收金额%Type,
            strSQL = strSQL & "" & IIf(blnPrice And dblTotalRegFee <> 0, 0, Val(mrsBill!应收)) & ","
            '  实收金额_In   门诊费用记录.实收金额%Type,
            dblTemp = GetActualMoney(str费别, mrsBill!收入项目ID, mrsBill!应收, mrsBill!收费细目ID)
            strSQL = strSQL & "" & IIf(blnPrice And dblTotalRegFee <> 0, 0, dblTemp) & ","
            '  病历费_In Number, --该条记录是否病历工本费
            If chk病历费.Value = 0 And Val(Nvl(mrsBill!附加标志)) = 1 Then
                strSQL = strSQL & "3,"
            Else
                strSQL = strSQL & "" & Val(Nvl(mrsBill!附加标志)) & ","
            End If
            If Val(Nvl(mrsBill!附加标志)) = 1 Then blnHaveBookFee = True
            '  保险大类id_In 门诊费用记录.保险大类id%Type,
            strSQL = strSQL & "" & ZVal(Nvl(mrsBill!保险大类id, 0)) & ","
            '  保险项目否_In 门诊费用记录.保险项目否%Type,
            strSQL = strSQL & "" & ZVal(Nvl(mrsBill!保险项目否, 0)) & ","
            '  统筹金额_In   门诊费用记录.统筹金额%Type,
            strSQL = strSQL & "" & ZVal(Nvl(mrsBill!统筹金额, 0)) & ","
            '  保险编码_In   门诊费用记录.保险编码%Type,
            strSQL = strSQL & "'" & Trim(Nvl(mrsBill!保险编码)) & "',"
            '  病人科室id_In 门诊费用记录.病人科室id%Type,
            strSQL = strSQL & "" & Val(mrsBill!病人科室id) & ","
            '  执行部门id_In 门诊费用记录.执行部门id%Type
            strSQL = strSQL & "" & Val(Nvl(mrsBill!执行部门id)) & ","
            '摘要_In       门诊费用记录.摘要%Type := Null
            strSQL = strSQL & "" & IIf(str划价NO <> "", "'划价:" & str划价NO & "'", "NULL") & ")"
            
            lng序号 = Val(Nvl(mrsBill!序号))
            Call zlAddArray(cllPro, strSQL)
            If blnPrice And dblTotalRegFee <> 0 Then
                strSQL = _
                "zl_门诊划价记录_Insert('" & str划价NO & "'," & mrsBill!序号 & "," & lng病人ID & ",NULL," & _
                         IIf(str门诊号 = "", "NULL", str门诊号) & ",'" & NeedCode(cbo付款方式.Text) & "'," & _
                         "'" & txtPatient.Text & "','" & NeedName(cbo性别.Text) & "','" & str年龄 & "'," & _
                         "'" & str费别 & "',NULL," & mlng挂号科室ID & "," & _
                         IIf(mblnStation, mlng挂号科室ID, UserInfo.部门ID) & ",'" & UserInfo.姓名 & "'," & "NULL" & "," & _
                         mrsBill!收费细目ID & ",'" & mrsBill!收费类别 & "',Null," & _
                         "NULL,1," & Val(Nvl(mrsBill!数次)) & ",NULL," & IIf(mrsBill!执行部门id = 0, mlng挂号科室ID, mrsBill!执行部门id) & "," & IIf(Val(Nvl(mrsBill!价格父号)) = 0, "NULL", mrsBill!价格父号) & "," & _
                         Val(Nvl(mrsBill!收入项目ID)) & ",'" & Trim(Nvl(mrsBill!收据费目)) & "'," & Val(Nvl(mrsBill!标准单价)) & "," & _
                         Val(mrsBill!应收) & "," & dblTemp & "," & str发生时间 & "," & str登记时间 & ",NULL,'" & UserInfo.姓名 & "','挂号:" & strNO & "')"
                Call zlAddArray(cllPro, strSQL)
            End If
            mrsBill.MoveNext
        Loop
        
        If lng序号 = 0 Then lng序号 = 1
         
        If blnHaveBookFee = False And Not mrsItems Is Nothing Then
            If blnPrice And dblTotalRegFee <> 0 And str划价NO = "" Then str划价NO = zlDatabase.GetNextNo(13)
            mrsItems.Filter = "性质 = 3"
            Do While Not mrsItems.EOF
                mrsInComes.Filter = "项目ID=" & mrsItems!项目ID
                If mrsInComes.RecordCount = 0 Then
                    MsgBox "未找到病历费,可能因并发原因造成数据提取失败，请重新提取病历数据!", vbInformation + vbDefaultButton1, gstrSysName
                    Exit Function
                End If
                lng序号 = lng序号 + 1
                
                strSQL = "Zl_病人预约挂号记录_Update("
                strSQL = strSQL & "'" & strNO & "',"
                strSQL = strSQL & "" & lng序号 & ","
                strSQL = strSQL & "NULL,"
                strSQL = strSQL & "" & IIf(mrsItems!性质 = 2, 1, "NULL") & ","
                strSQL = strSQL & "'" & mrsItems!类别 & "',"
                strSQL = strSQL & "'" & mrsItems!项目ID & "',"
                strSQL = strSQL & "" & Val(Nvl(mrsItems!数次)) & ","
                strSQL = strSQL & "" & Val(Nvl(mrsInComes!单价)) & ","
                strSQL = strSQL & "" & Val(Nvl(mrsInComes!收入项目ID)) & ","
                strSQL = strSQL & "'" & Trim(Nvl(mrsInComes!收据费目)) & "',"
                strSQL = strSQL & "" & IIf(blnPrice And dblTotalRegFee <> 0, 0, mrsInComes!应收) & ","
                strSQL = strSQL & "" & IIf(blnPrice And dblTotalRegFee <> 0, 0, mrsInComes!实收) & ","
                
                strSQL = strSQL & "" & IIf(mrsItems!性质 = 3, 1, IIf(mrsItems!性质 = 4, 2, 0)) & ","
                strSQL = strSQL & "" & ZVal(Nvl(mrsItems!保险大类id, 0)) & ","
                strSQL = strSQL & "" & ZVal(Nvl(mrsItems!保险项目否, 0)) & ","
                strSQL = strSQL & "" & ZVal(Nvl(mrsInComes!统筹金额, 0)) & ","
                strSQL = strSQL & "'" & Trim(Nvl(mrsItems!保险编码)) & "',"
                strSQL = strSQL & "" & mlng挂号科室ID & ","
                strSQL = strSQL & "" & IIf(mrsItems!执行科室ID = 0, mlng挂号科室ID, mrsItems!执行科室ID) & ","
                '摘要_In       门诊费用记录.摘要%Type := Null
                strSQL = strSQL & "" & IIf(str划价NO <> "", "'划价:" & str划价NO & "'", "NULL") & ")"
                
                Call zlAddArray(cllPro, strSQL)
                If blnPrice And dblTotalRegFee <> 0 Then
                    strSQL = _
                    "zl_门诊划价记录_Insert('" & str划价NO & "'," & mrsBill!序号 + i & "," & lng病人ID & ",NULL," & _
                             IIf(str门诊号 = "", "NULL", str门诊号) & ",'" & NeedCode(cbo付款方式.Text) & "'," & _
                             "'" & txtPatient.Text & "','" & NeedName(cbo性别.Text) & "','" & str年龄 & "'," & _
                             "'" & str费别 & "',NULL," & mlng挂号科室ID & "," & _
                             IIf(mblnStation, mlng挂号科室ID, UserInfo.部门ID) & ",'" & UserInfo.姓名 & "'," & "NULL" & "," & _
                             mrsBill!收费细目ID & ",'" & mrsBill!收费类别 & "',Null," & _
                             "NULL,1," & Val(Nvl(mrsBill!数次)) & ",NULL," & IIf(mrsBill!执行部门id = 0, mlng挂号科室ID, mrsBill!执行部门id) & "," & IIf(Val(Nvl(mrsBill!价格父号)) = 0, "NULL", mrsBill!价格父号) & "," & _
                             Val(Nvl(mrsBill!收入项目ID)) & ",'" & Trim(Nvl(mrsBill!收据费目)) & "'," & Val(Nvl(mrsBill!标准单价)) & "," & _
                             Val(mrsBill!应收) & "," & GetActualMoney(str费别, mrsBill!收入项目ID, mrsBill!应收, mrsBill!收费细目ID) & "," & str发生时间 & "," & str登记时间 & ",NULL,'" & UserInfo.姓名 & "','挂号:" & strNO & "')"
                    Call zlAddArray(cllPro, strSQL)
                End If
                mrsItems.MoveNext
            Loop
        End If
        If Not mrsItems Is Nothing Then
            
            If blnPrice And dblTotalRegFee <> 0 And str划价NO = "" Then str划价NO = zlDatabase.GetNextNo(13)
            
            '处理药事服务费等
            mrsItems.Filter = "性质 = 5"
            Do While Not mrsItems.EOF
                lng序号 = lng序号 + 1
                mrsInComes.Filter = "项目ID=" & mrsItems!项目ID
                strSQL = "Zl_病人预约挂号记录_Update("
                strSQL = strSQL & "'" & strNO & "',"
                strSQL = strSQL & "" & lng序号 & ","
                strSQL = strSQL & "NULL,"
                strSQL = strSQL & "" & IIf(mrsItems!性质 = 2, 1, "NULL") & ","
                strSQL = strSQL & "'" & mrsItems!类别 & "',"
                strSQL = strSQL & "'" & mrsItems!项目ID & "',"
                strSQL = strSQL & "" & Val(Nvl(mrsItems!数次)) & ","
                strSQL = strSQL & "" & Val(Nvl(mrsInComes!单价)) & ","
                strSQL = strSQL & "" & Val(Nvl(mrsInComes!收入项目ID)) & ","
                strSQL = strSQL & "'" & Trim(Nvl(mrsInComes!收据费目)) & "',"
                strSQL = strSQL & "" & IIf(blnPrice And dblTotalRegFee <> 0, 0, mrsInComes!应收) & ","
                strSQL = strSQL & "" & IIf(blnPrice And dblTotalRegFee <> 0, 0, mrsInComes!实收) & ","
                strSQL = strSQL & "" & IIf(mrsItems!性质 = 3, 1, IIf(mrsItems!性质 = 4, 2, 0)) & ","
                strSQL = strSQL & "" & ZVal(Nvl(mrsItems!保险大类id, 0)) & ","
                strSQL = strSQL & "" & ZVal(Nvl(mrsItems!保险项目否, 0)) & ","
                strSQL = strSQL & "" & ZVal(Nvl(mrsInComes!统筹金额, 0)) & ","
                strSQL = strSQL & "'" & Trim(Nvl(mrsItems!保险编码)) & "',"
                strSQL = strSQL & "" & mlng挂号科室ID & ","
                strSQL = strSQL & "" & IIf(mrsItems!执行科室ID = 0, mlng挂号科室ID, mrsItems!执行科室ID) & ","
                '摘要_In       门诊费用记录.摘要%Type := Null
                strSQL = strSQL & "" & IIf(str划价NO <> "", "'划价:" & str划价NO & "'", "NULL") & ")"
                Call zlAddArray(cllPro, strSQL)
                If blnPrice And dblTotalRegFee <> 0 Then
                    strSQL = _
                    "zl_门诊划价记录_Insert('" & str划价NO & "'," & mrsBill!序号 + i & "," & lng病人ID & ",NULL," & _
                     IIf(str门诊号 = "", "NULL", str门诊号) & ",'" & NeedCode(cbo付款方式.Text) & "'," & _
                     "'" & txtPatient.Text & "','" & NeedName(cbo性别.Text) & "','" & str年龄 & "'," & _
                     "'" & str费别 & "',NULL," & mlng挂号科室ID & "," & _
                     IIf(mblnStation, mlng挂号科室ID, UserInfo.部门ID) & ",'" & UserInfo.姓名 & "'," & "NULL" & "," & _
                     mrsBill!收费细目ID & ",'" & mrsBill!收费类别 & "',Null," & _
                     "NULL,1," & Val(Nvl(mrsBill!数次)) & ",NULL," & IIf(mrsBill!执行部门id = 0, mlng挂号科室ID, mrsBill!执行部门id) & "," & IIf(Val(Nvl(mrsBill!价格父号)) = 0, "NULL", mrsBill!价格父号) & "," & _
                     Val(Nvl(mrsBill!收入项目ID)) & ",'" & Trim(Nvl(mrsBill!收据费目)) & "'," & Val(Nvl(mrsBill!标准单价)) & "," & _
                     Val(mrsBill!应收) & "," & GetActualMoney(str费别, mrsBill!收入项目ID, mrsBill!应收, mrsBill!收费细目ID) & "," & str发生时间 & "," & str登记时间 & ",NULL,'" & UserInfo.姓名 & "','挂号:" & strNO & "')"
                    Call zlAddArray(cllPro, strSQL)
                End If
                mrsItems.MoveNext
            Loop
        End If
    End If
    
    '--预约接收
    strSQL = "" & _
    "Zl_预约挂号接收_出诊_insert('" & strNO & "','" & IIf(blnNoPrint, "", txtFact.Text) & "',Null," & _
    lng结帐ID & ",'" & strRoom & "'," & ZVal(lng病人ID) & "," & IIf(str门诊号 = "", "NULL", str门诊号) & ",'" & txtPatient.Text & "'," & _
    "'" & NeedName(cbo性别.Text) & "','" & str年龄 & "','" & NeedCode(cbo付款方式.Text) & "'," & _
    "'" & str费别 & "','" & str结算方式 & "'," & cur现金 - cur卡费 & "," & cur预交 & "," & cur个帐 & "," & _
    str发生时间 & "," & ZVal(lngSN) & ",'" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & IIf(mTy_Para.bln挂号生成队列, 1, 0) & "," & _
    str登记时间 & ","  '问题号:48350
    '卡类别id_In   病人预交记录.卡类别id%Type := Null,
    strSQL = strSQL & "" & IIf(mCurCardPay.Have挂号费 And mCurCardPay.lng医疗卡类别ID <> 0 And mCurCardPay.bln消费卡 = False, mCurCardPay.lng医疗卡类别ID, "NULL") & ","
    '结算卡序号_In 病人预交记录.结算卡序号%Type := Null,
    strSQL = strSQL & "" & IIf(mCurCardPay.Have挂号费 And mCurCardPay.lng医疗卡类别ID <> 0 And mCurCardPay.bln消费卡, mCurCardPay.lng医疗卡类别ID, "NULL") & ","
    '卡号_In       病人预交记录.卡号%Type := Null,
    strSQL = strSQL & "" & IIf(mCurCardPay.Have挂号费 And mCurCardPay.str刷卡卡号 <> "", "'" & mCurCardPay.str刷卡卡号 & "'", "NULL") & ","
    '交易流水号_In 病人预交记录.交易流水号%Type := Null,
    strSQL = strSQL & " NULL,"
    '交易说明_In   病人预交记录.交易说明%Type := Null
    strSQL = strSQL & " NULL,"
    '险类_In       病人挂号记录.险类%Type := Null,
    strSQL = strSQL & "" & IIf(mintInsure = 0, "Null", mintInsure) & ","
    '结算模式_In   Number := 0,
    strSQL = strSQL & "" & IIf(mPatiChargeMode = EM_先诊疗后结算, 1, 0) & ","
    '记帐费用_In Number:=0
    strSQL = strSQL & "" & IIf(mRegistFeeMode = EM_RG_记帐, 1, 0) & ","
    '冲预交病人ids_In Varchar2 := Null
    strSQL = strSQL & "'" & lng病人ID & "," & mstr病人家属IDs & "'," '79868,冉俊明,2015-6-15,使用家属预交
    '三方调用_In      Number := 0,
    strSQL = strSQL & "" & 0 & ","
    '更新交款余额_In  Number := 1,
    strSQL = strSQL & "" & 1 & ","
    '摘要_In          病人挂号记录.摘要%Type := Null
    strSQL = strSQL & "'" & cbo备注.Text & "',"
    strSQL = strSQL & IIf(str划价NO = "", "Null", "'" & str划价NO & "'") & ")"
    Call zlAddArray(cllPro, strSQL)
    
    '预约挂号接收
    strSQL = "" & _
           " Select D.科室id, C.项目id, C.医生id, C.医生姓名,D.号码 " & _
           " From 门诊费用记录 A, 病人挂号记录 B, 临床出诊记录 C, 临床出诊号源 D " & _
           " Where A.记录性质 = 4 And A.记录状态 = 0 And A.NO = [1] And A.序号 = 1 And A.NO = B.NO And B.出诊记录ID = C.ID And C.号源ID = D.ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
    '问题:31187:将挂号汇总单独出来
    If rsTemp.EOF = False Then
        strSQL = "zl_病人挂号汇总_Update("
        '  医生姓名_In   挂号安排.医生姓名%Type,
        strSQL = strSQL & "'" & Nvl(rsTemp!医生姓名) & "',"
        '  医生id_In     挂号安排.医生id%Type,
        strSQL = strSQL & "" & ZVal(Val(Nvl(rsTemp!医生ID))) & ","
        '  收费细目id_In 门诊费用记录.收费细目id%Type,
        strSQL = strSQL & "" & Val(Nvl(rsTemp!项目ID)) & ","
        '  执行部门id_In 门诊费用记录.执行部门id%Type,
        strSQL = strSQL & "" & Val(Nvl(rsTemp!科室ID)) & ","
        '  发生时间_In   门诊费用记录.发生时间%Type,
        strSQL = strSQL & "" & str发生时间 & ","
        '  预约标志_In   Number := 0  --是否为预约接收:0-非预约挂号; 1-预约挂号,2-预约接收
        strSQL = strSQL & "2" & ","
        '  号码_In       挂号安排.号码%Type := Null
        strSQL = strSQL & "'" & IIf(txt号别.Text = "+", "", txt号别.Text) & "',0,"
        strSQL = strSQL & "" & vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("记录ID")) & ")"
        Call zlAddArray(cllProAfter, strSQL)
    End If

    SaveRegister_预约接收 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Private Sub SaveData(Optional blnCall结束挂号 As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存数据
    '参数:blnCall结束挂号-true结束挂号按钮调用(否则为确认按钮调用)
    '编制:刘兴洪
    '日期:2009-12-02 16:08:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim lng病人ID As Long, str门诊号 As String, lng结帐ID As Long, lngCard结帐ID As Long, lngSN As Long
    Dim str登记时间 As String, str发生时间 As String, strNO As String, strRoom As String, strInfo As String, strTmp As String
    Dim bytType As Byte, str费别 As String, str年龄 As String, strPatiInforXML As String
    Dim str卡号 As String, str密码 As String, str出生日期 As String
    Dim strSQL As String, strFact As String, strAdvance As String, strMCAccount As String
    Dim str联系电话 As String, int原结算模式 As Integer, RegistFeeMode As EM_REGISTFEE_MODE
    Dim blnSlipPrint As Boolean, blnNoDoc As Boolean, blnCodePrint As Boolean
    Dim cur现金 As Currency, cur个帐 As Currency, cur预交 As Currency, str身份证号 As String, dblThreeSwap As Double  '三方支付额
    Dim curOneCard As Currency, dblOneCardBalance As Double, strFilter As String
    Dim strCardNo As String, intCardType As Integer, strTransFlow As String
    Dim rsOneCard1 As ADODB.Recordset, rsOneCard2 As ADODB.Recordset, bln发卡 As Boolean   '标识挂号中，是否同时进行了发卡或绑卡操作
    Dim objICCard As Object, dblPaySum As Double, str结算方式, blnPrice As Boolean   '建档病人存为划价单
    Dim strStyle As String, rsCheck As ADODB.Recordset, dbl费用金额 As Double, dbl结帐金额 As Double
    Dim int价格父号 As Integer, intMsgReturn As Integer, blnNotCommit As Boolean, dat发生时间 As Date
    Dim blnNoPrint As Boolean, cur应缴 As Currency, cur卡费 As Currency, bln达到限号数 As Boolean, blnEnterPrint As Boolean
    Dim i As Long, j As Long, k As Long, blnAfterRefresh As Boolean, cllProBefor As Collection   '事务前执行数据
    Dim blnCancel As Boolean, str划价NO As String, strCardBillNO As String, cllPro As Collection   '正常事务中执行的数据
    Dim blnNew As Boolean, blnPati As Boolean, blnTrans As Boolean, cllProAfter As Collection   '接口调用后执行数据
    Dim byt复诊 As Byte, blnPrintBooking As Boolean, bln连续 As Boolean
    Dim rsTmp As ADODB.Recordset, rsSNCheck As ADODB.Recordset
    Dim Datsys As Date, lngRow As Long, blnInsertHisBook As Boolean
    Dim cllCardPro As Collection, cllTheeSwap As Collection, str时点 As String, int性质 As Integer
    Dim bln追加时段 As Boolean    '用于标识,是否用于时段已经,挂号或者过期,但是没有达到限号数的情况,
    Dim dbl找补 As Double, curCard As Currency, cur连续 As Currency
    Dim lng就诊ID As Long, str医生姓名 As String, lng医生ID As Long, strErrMsg As String
    Dim lng记录ID As Long
    Dim dblTotal(0 To 1) As Double, dblTemp As Double
   
    Err = 0: On Error GoTo ErrGo:
    mobjfrmPatiInfo.mstrFirstCode = ""
    If chkPrint.Value = 1 Then    '重打
        If zlRePrintRegistered = False Then Exit Sub
    ElseIf chkCancel.Value = 1 Or (mbytInState = 1 And mbytMode = 4) Then    '退号
        If zlExcuteDelRegistered = False Then Exit Sub
        If mbytInState = 1 And (mbytMode = 4 Or mbytMode = 3) Then mblnOk = True: Unload Me: Exit Sub
    Else
        '115168:李南春，2017/12/13，保存发卡的医疗卡类型
        If mCurSendCard.lng卡类别ID = 0 Then mCurSendCard = gCurSendCard
        If cbo结算方式.Visible Then
            If CheckPayStyleValied(lngRow) = False Then Exit Sub
            mblnPre连续 = mbln连续挂号
            mbln连续挂号 = False
            If lngRow <> 0 Then
                If Val(txt缴款.Text) = 0 Then
                    int性质 = Get性质(NeedName(cbo结算方式.Text), strStyle)
                    If mTy_Para.byt缴款方式 = 1 And int性质 <> 7 And int性质 <> 8 Then
                        With vsfPay
                            .TextMatrix(lngRow, 0) = NeedName(cbo结算方式.Text)
                            If Val(txt本次应缴.Text) <> 0 Then
                                .TextMatrix(lngRow, 1) = Format(Val(.TextMatrix(lngRow, 1)) + Val(txt本次应缴.Text) - mcur应缴, "0.00")
                            End If
                            .RowData(lngRow) = int性质: .TextMatrix(lngRow, .ColIndex("结算方式")) = strStyle
                        End With
                        cur连续 = Val(txt本次应缴.Text) - mcur应缴: txt本次应缴.Text = "0.00": txt缴款.Text = ""
                        mbln连续挂号 = True
                    Else
                        With vsfPay
                            .TextMatrix(lngRow, 0) = NeedName(cbo结算方式.Text)
                            .TextMatrix(lngRow, 1) = Format(Val(txt本次应缴.Text), "0.00")
                            .RowData(lngRow) = int性质
                            .TextMatrix(lngRow, .ColIndex("结算方式")) = strStyle
                        End With
                        txt本次应缴.Text = "0.00": txt缴款.Text = ""
                    End If
                Else
                    mcur合计 = 0
                    mcur应缴 = 0
                    If Get性质(NeedName(cbo结算方式.Text), strStyle) = 1 And Val(txt本次应缴.Text) < Val(txt缴款.Text) Then
                        With vsfPay
                            .TextMatrix(lngRow, 0) = NeedName(cbo结算方式.Text)
                            .TextMatrix(lngRow, 1) = Format(Val(txt本次应缴.Text), "0.00")
                            .RowData(lngRow) = Get性质(NeedName(cbo结算方式.Text), strStyle)
                            .TextMatrix(lngRow, .ColIndex("结算方式")) = strStyle
                        End With
                        dbl找补 = Val(txt缴款.Text) - Val(txt本次应缴.Text)
                        txt本次应缴.Text = "0.00": txt找补.Text = Format(dbl找补, "0.00"): txt缴款.Text = ""
                    Else
                        With vsfPay
                            .TextMatrix(lngRow, 0) = NeedName(cbo结算方式.Text)
                            .TextMatrix(lngRow, 1) = Format(Val(txt缴款.Text), "0.00")
                            .RowData(lngRow) = Get性质(NeedName(cbo结算方式.Text), strStyle)
                            .TextMatrix(lngRow, .ColIndex("结算方式")) = strStyle
                        End With
                        If Val(txt本次应缴.Text) = Val(txt缴款.Text) Then
                            txt本次应缴.Text = "0.00": txt缴款.Text = ""
                        Else
                            txt本次应缴.Text = Format(Val(txt本次应缴.Text) - Val(txt缴款.Text), "0.00"): txt缴款.Text = ""
                            If cbo结算方式.ListCount > 0 Then cbo结算方式.ListIndex = 0
                            If cbo结算方式.Visible And cbo结算方式.Enabled Then cbo结算方式.SetFocus
                            Exit Sub
                        End If
                    End If
                End If
            End If
        End If
        '是否保存为划价单
        blnPrice = CheckIsPrice
        txtPatient.Text = Trim(txtPatient.Text): txt年龄.Text = Trim(txt年龄.Text)
        If txtSN.Visible Then
            If Val(txtSN.Text) = 0 Then txtSN.Text = ""
            lngSN = Val(txtSN.Text)
        End If
        '134429:李南春，2019/1/12，只能以选中行号进行判断
        If mlngPreRow <> vsfPlan.Row And mlngPreRow < vsfPlan.Rows And mlngPreRow <> 0 And txt号别.Text <> "" And txt号别.Text <> "+" Then
            MsgBox "挂号类别选项发生了变化，请重新选择挂号号别", vbInformation, gstrSysName
            Exit Sub
        End If
        lng记录ID = Val(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("记录ID")))
        lng记录ID = IIf(txt号别.Text = "+", 0, lng记录ID)
        '相关数据检查
        If CheckInputValied = False Then RestorePay: Exit Sub
        If CheckNoValied(vsfPlan.Row) = False Then RestorePay: Exit Sub
        If PrivCheck() = False Then RestorePay: Exit Sub
        If Not mrsItems Is Nothing And Not mrsInComes Is Nothing Then
            mrsItems.Filter = "性质=4"
            If mrsItems.RecordCount > 0 Then
                If Not mrsItems.EOF Then
                    mrsInComes.Filter = "项目ID=" & mrsItems!项目ID
                    If Not mrsInComes.EOF Then
                        '问题号:110224,焦博,2017/06/20
                        If gCurSendCard.rs卡费 Is Nothing Then
                            MsgBox "卡费的收费项目未正确设置，请检查后重试！", vbInformation, gstrSysName
                            RestorePay
                            Exit Sub
                        End If
                    End If
                End If
            End If
            mrsItems.Filter = ""
            mrsInComes.Filter = ""
        End If

        strMCAccount = Trim(mobjfrmPatiInfo.txtPatiMCNO(0).Text)
        If mlngOutModeMC = 920 And strMCAccount <> "" Then
            If strMCAccount <> mobjfrmPatiInfo.txtPatiMCNO(0).Tag Then
                If CheckExistsMCNO(strMCAccount) Then
                    If cmdMore.Enabled Then Call cmdMore_Click
                    RestorePay
                    Exit Sub
                End If
            End If
            strMCAccount = UCase(strMCAccount)
        End If
        
        '保存创建病人信息
        If Not mobjfrmPatiInfo Is Nothing Then
            If Not mobjfrmPatiInfo.SaveAfterArrList Then Exit Sub
        End If
        
        '102230,调用外挂部件接口
        If mbytMode = 0 And mbytInState = 0 Then
            If mrsInfo Is Nothing Then
                strPatiInforXML = GetPatiInforXML
                If PatiValiedCheckByPlugIn(mlngModul, 0, strPatiInforXML) = False Then
                    Call RestorePay: Exit Sub
                End If
            End If
        End If
        
        '票据打印提醒
        Call SaveInvoiceNotify(blnPrice, blnSlipPrint, blnNoPrint, blnPrintBooking, blnCodePrint)
        
        '票据号码检查
        If mbytMode <> 1 And Not blnNoPrint Then
            If gblnBill挂号 Then
                If Trim(txtFact.Text) = "" Then
                    MsgBox "必须输入一个有效的票据号码！", vbInformation, gstrSysName
                    Call RestorePay
                    txtFact.SetFocus: Exit Sub
                End If

InvoiceHandle:
                mlng领用ID = CheckUsedBill(IIf(gblnSharedInvoice, 1, 4), IIf(mlng领用ID > 0, mlng领用ID, glng挂号ID), txtFact.Text, IIf(mblnStartFactUseType, mstrUseType, ""))
                If mlng领用ID <= 0 Then
                    Select Case mlng领用ID
                    Case 0    '操作失败
                    Case -1
                        MsgBox "你没有自用和共用的挂号票据,请先领用一批票据或设置本地共用票据！", vbInformation, gstrSysName
                    Case -2
                        MsgBox "本地的共用票据已经用完,请先领用一批票据或重新设置本地共用票据！", vbInformation, gstrSysName
                    Case -3
                        MsgBox "票据号码不在当前有效领用范围内,请重新输入！", vbInformation, gstrSysName
                        txtFact.SetFocus
                    End Select
                    Call RestorePay
                    Exit Sub
                End If
            
                '并发操作检查,票号是否已用
                If CheckBillRepeat(mlng领用ID, IIf(gblnSharedInvoice, 1, 4), txtFact.Text) Then
                    If txtFact.Locked = False And txtFact.Tag <> Trim(txtFact.Text) Then
                        MsgBox "票据号""" & txtFact.Text & """已经被使用，请重新输入。", vbInformation, gstrSysName
                        Call RestorePay
                        zlControl.ControlSetFocus txtFact: Exit Sub
                    Else
                        Call RefreshFact
                        If txtFact.Text = "" Then
                            Call RestorePay
                            zlControl.ControlSetFocus txtFact: Exit Sub
                        Else
                            MsgBox "当前票据号已经被使用，已重新获取票据号:" & txtFact.Text, vbInformation, gstrSysName
                            GoTo InvoiceHandle
                        End If
                    End If
                End If
            Else
                If Len(txtFact.Text) <> gbytFactLength And txtFact.Text <> "" Then
                    MsgBox "票据号码长度应该为 " & gbytFactLength & " 位！", vbInformation, gstrSysName
                    Call RestorePay
                    txtFact.SetFocus: Exit Sub
                End If
            End If
        End If
        timPlan.Enabled = False
        
        If mRegistFeeMode <> EM_RG_记帐 Then
            '记帐不进行语音提示
            If Not (mintInsure <> 0 And mstrYBPati <> "") Then
                If gblnLED And mbytMode <> 1 And mbytInState = 0 And txt缴款.Tag = "" Then
                    cur应缴 = mcur应缴 + GetRegistMoney
                    zl9LedVoice.Speak "#21 " & Format(cur应缴, "0.00")
                End If
            End If
        End If
        txt缴款.Tag = ""
        '----------------
        Set cllPro = New Collection: Set cllProAfter = New Collection: Set cllProBefor = New Collection
        Datsys = zlDatabase.Currentdate
        str费别 = NeedName(cbo费别.Text)
        str年龄 = Trim(txt年龄.Text)
        If IsNumeric(str年龄) Then str年龄 = str年龄 & cbo年龄单位.Text
        
        '挂号费用信息
        If Not blnPrice Then
            curCard = 0
            mstrCard结算方式 = ""
            If Not mrsItems Is Nothing Then
                mrsItems.Filter = "性质=4"
                If mrsItems.RecordCount > 0 Then
                    Do While Not mrsItems.EOF
                        mrsInComes.Filter = "项目ID=" & mrsItems!项目ID
                        Do While Not mrsInComes.EOF
                            curCard = curCard + mrsInComes!实收
                            mrsInComes.MoveNext
                        Loop
                        mrsItems.MoveNext
                    Loop
                End If
                mrsItems.Filter = ""
            End If
            
            '137473:李南春，2019/1/31，三方卡(消费卡)支付时区分支付费用
            Call ClearCardMoney
            With vsfPay
                For i = 1 To .Rows - 1
                    If .TextMatrix(i, 0) <> "" Then
                        If .RowData(i) = 0 Then
                            cur预交 = Val(.TextMatrix(i, 1))
                        Else
                            If .RowData(i) = 3 Then
                                cur个帐 = Val(.TextMatrix(i, 1))
                            Else
                                If strFilter = "" Then
                                    strFilter = "结算方式='" & .TextMatrix(i, 4) & "'"
                                Else
                                    strFilter = strFilter & " Or 结算方式='" & .TextMatrix(i, 4) & "'"
                                End If
                                If curCard <> 0 Then
                                    If Val(.TextMatrix(i, .ColIndex("金额"))) = curCard Then
                                        mstrCard结算方式 = .TextMatrix(i, .ColIndex("结算方式"))
                                        curCard = 0
                                        If .RowData(i) = 7 Or .RowData(i) = 8 Then
                                            mCurCardPay.Have卡费 = True
                                        End If
                                    ElseIf Val(.TextMatrix(i, 1)) > curCard Then
                                        str结算方式 = str结算方式 & "|" & .TextMatrix(i, 4) & "," & Val(.TextMatrix(i, 1)) - curCard - Val(.TextMatrix(i, 7)) & "," & .TextMatrix(i, 2) & "," & IIf(.RowData(i) = 7 Or .RowData(i) = 8, 1, 0)
                                        mstrCard结算方式 = .TextMatrix(i, 4)
                                        curCard = 0
                                        If .RowData(i) = 7 Or .RowData(i) = 8 Then
                                            mCurCardPay.Have卡费 = True
                                            mCurCardPay.Have挂号费 = True
                                        End If
                                    End If
                                Else
                                    str结算方式 = str结算方式 & "|" & .TextMatrix(i, 4) & "," & .TextMatrix(i, 1) - Val(.TextMatrix(i, 7)) & "," & .TextMatrix(i, 2) & "," & IIf(.RowData(i) = 7 Or .RowData(i) = 8, 1, 0)
                                    If .RowData(i) = 7 Or .RowData(i) = 8 Then
                                        mCurCardPay.Have挂号费 = True
                                    End If
                                End If
                                cur现金 = cur现金 + Val(.TextMatrix(i, 1)) - Val(.TextMatrix(i, 7))
                            End If
                        End If
                    End If
                Next i
                If curCard > 0 Then
                    MsgBox "录入的结算方式不足支付卡费,请至少录入一项金额大于(" & Format(curCard, "0.00") & ")的结算方式!", vbInformation, gstrSysName
                    Call RestorePay: Exit Sub
                End If
                If cur个帐 > mdbl个帐余额 + mcur个帐透支 Then
                    MsgBox "医保帐户余额不足,不能使用医保支付!", vbInformation, gstrSysName
                    Call RestorePay: Exit Sub
                End If
                If str结算方式 <> "" Then
                    str结算方式 = Mid(str结算方式, 2)
                Else
                    Get性质 NeedName(cbo结算方式.Text), strStyle
                    str结算方式 = strStyle & ",0,,0"
                End If
            End With
            
            If mblnOneCard And cur现金 <> 0 And mRegistFeeMode <> EM_RG_记帐 Then
                mrsOneCard.Filter = strFilter
                If mrsOneCard.RecordCount > 0 Then
                    If mstrYBPati <> "" Then
                        MsgBox "不支持医保病人使用一卡通支付！", vbInformation, gstrSysName
                        Call RestorePay: Exit Sub
                    End If
                    If mobjICCard Is Nothing Then
                        MsgBox "使用一卡通支付必须先读卡！", vbInformation, gstrSysName
                        Call RestorePay: Exit Sub
                    End If
                    For i = 1 To vsfPay.Rows - 1
                        If vsfPay.TextMatrix(i, 0) = Nvl(mrsOneCard!结算方式) Then
                            curOneCard = mobjICCard.GetSpare
                            If curOneCard < Val(vsfPay.TextMatrix(i, 1)) Then
                                MsgBox "卡上余额" & Format(curOneCard, "0.00") & ",本次要求支付金额" & Format(Val(vsfPay.TextMatrix(i, 1)), "0.00"), vbInformation, gstrSysName
                                Call RestorePay: Exit Sub
                            Else
                                curOneCard = Val(vsfPay.TextMatrix(i, 1))
                            End If
                            Exit For
                        End If
                    Next i
                End If
            End If
            If mRegistFeeMode <> EM_RG_记帐 Then
                For i = 1 To vsfPay.Rows - 1
                    strSQL = "Select ID,名称,结算方式 From 医疗卡类别 Where 名称= [1] "
                    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, vsfPay.TextMatrix(i, 0))
                    If Not rsTemp.EOF Then
                        mCurCardPay.lng医疗卡类别ID = rsTemp!ID
                        mCurCardPay.bln消费卡 = False
                        mCurCardPay.str结算方式 = rsTemp!结算方式
                        mCurCardPay.str名称 = rsTemp!名称
                        If CheckBrushCard(CDbl(vsfPay.TextMatrix(i, 1))) = False Then RestorePay: Exit Sub
                        dblThreeSwap = CDbl(vsfPay.TextMatrix(i, 1))
                        Exit For
                    Else
                        strSQL = "Select 编号,名称,结算方式 From 消费卡类别目录 Where 名称= [1] "
                        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, vsfPay.TextMatrix(i, 0))
                        If Not rsTemp.EOF Then
                            mCurCardPay.lng医疗卡类别ID = rsTemp!编号
                            mCurCardPay.bln消费卡 = True
                            mCurCardPay.str结算方式 = rsTemp!结算方式
                            mCurCardPay.str名称 = rsTemp!名称
                            If CheckBrushCard(CDbl(vsfPay.TextMatrix(i, 1))) = False Then RestorePay: Exit Sub
                            Exit For
                        End If
                    End If
                Next i
            End If
        End If
        
        If (mbytMode = 2 Or mbytMode = 0 And mbytInState = 0 And mstrNoIn <> "") Then
            strNO = cboNO.Text
        Else
            strNO = zlDatabase.GetNextNo(12)
            mstr连续挂号_挂号NO = mstr连续挂号_挂号NO & "," & strNO
        End If

        If mbytMode <> 1 Then
            lng结帐ID = zlDatabase.GetNextId("病人结帐记录")
        End If
        byt复诊 = Val(mobjfrmPatiInfo.chk复诊.Value)
        '获取分诊诊室
        If mbytMode <> 1 And txt号别.Text <> "+" And vsfPlan.TextMatrix(vsfPlan.Row, GetCol("分诊")) <> "" Then  '预约时不分诊
            strRoom = GetRoom(lng记录ID)
        End If

        '挂号病人信息处理:新发卡,绑定卡,以及建病案的新旧病人
        If mblnAddCardItem Or Trim(txt门诊号.Text) <> "" Or (txtIDCard.Text <> "" And mbytMode = 1) Then
            str门诊号 = txt门诊号.Text
            If mrsInfo Is Nothing Then
                bytType = 1
                lng病人ID = zlDatabase.GetNextNo(1)
                int原结算模式 = 0
            Else
                If IsNull(mrsInfo!门诊号) Then
                    bytType = 2
                Else
                    bytType = 3
                End If
                lng病人ID = mrsInfo!病人ID
                int原结算模式 = Val(Nvl(mrsInfo!结算模式))
            End If
            blnPati = True
        ElseIf Not mrsInfo Is Nothing Then
            lng病人ID = mrsInfo!病人ID
            int原结算模式 = Val(Nvl(mrsInfo!结算模式))
        End If
        
        If zlIsAllowPatiChargeFeeMode(lng病人ID, int原结算模式) = False Then RestorePay: Exit Sub
        
        If Trim(mobjfrmPatiInfo.txt卡号.Text) <> "" Then    '读取有卡号的病人时没有加载卡号到界面
            str卡号 = Trim(mobjfrmPatiInfo.txt卡号.Text)
            str密码 = zlCommFun.zlStringEncode(Trim(mobjfrmPatiInfo.txt密码.Text))
        End If

        '门诊号检查
        If IsValiedMzNo(lng病人ID, str门诊号) = False Then RestorePay: Exit Sub

        If mViewMode <> V_普通号 Then
            Set mrsSNState = GetSNState(lng记录ID, CDate(Format(IIf(picBookingDate.Visible, dtpAppointmentDate.Value, Datsys), "yyyy-MM-dd")))
        End If

        '序号检查
        If Trim(txtSN.Text) <> "" And Val(txtSN.Tag) <> Val(txtSN.Text) And Not mrsSNState Is Nothing Then
            strSQL = "Select Nvl(挂号状态,0) As 状态,操作员姓名 From 临床出诊序号控制 Where 记录ID = [1] AND 序号 = [2] "
            Set rsSNCheck = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng记录ID, lngSN)
            mrsSNState.Filter = "序号=" & lngSN
            If rsSNCheck.RecordCount > 0 Then
                If rsSNCheck!状态 = 1 Or rsSNCheck!状态 = 2 Or rsSNCheck!状态 = 4 Or ((rsSNCheck!状态 = 5 Or rsSNCheck!状态 = 3) And rsSNCheck!操作员姓名 <> UserInfo.姓名) Then
                    Call vsfPlan_EnterCell
                    lngSN = GetCurrSN(, True)   '自动取下一个
                    If lngSN = 0 Then
                        MsgBox "序号" & Trim(txtSN.Text) & "已经被挂出请选择别的号进行挂号。", vbInformation, gstrSysName
                        Call RestorePay: Exit Sub
                    Else
                        If IsDate(mtyRegPlanState.strSelTime) And mtyRegPlanState.lngSelNO = lngSN And Format(dtpAppointmentTime.Value, "hh:mm:00") <> Format(mtyRegPlanState.strSelTime, "hh:mm:00") Then
                            dtpAppointmentTime.Value = CDate(mtyRegPlanState.strSelTime)
                        End If
                    End If
                End If
            End If
        End If
        '对在操作员有随机序号选择权限,启用序号控制,没有设置时段这种情况下
        '对操作员直接挂出最后一个序号这种情况需要特殊处理
        '因为前面已经检查过限制条件 这里不在进行限制条件的检查 这里直接把启用序号控制且没有启用时段的安排并且序号为空的这种情况进行特殊处理
        If vsfPlan.TextMatrix(vsfPlan.Row, GetCol("序号控制")) <> "" And mViewMode = v_专家号 And lngSN = 0 Then
            mbln加号 = True
        ElseIf mViewMode = v_专家号分时段 And lngSN = 0 And mbln加号 = False Then
            '这里是对专家号分时段情况 在序号在不明原因的情况合作序号被操作员误操作删掉的情况下 进行检查 处理 恢复序号或者 提示
            mrsSNState.Filter = 0
            i = vsfList.Row: j = vsfList.Col

            If (mtyRegPlanState.lngSelX <> vsfList.Row Or mtyRegPlanState.lngSelY <> vsfList.Col) And IsDate(mtyRegPlanState.strSelTime) Then
                '如果选择的序号时段正确,但是没有序号的情况
                mblnStateChange = True
                i = mtyRegPlanState.lngSelX
                j = mtyRegPlanState.lngSelY
                If vsfList.Cell(flexcpData, vsfList.Row, vsfList.Col) Like "加*" Then
                    i = vsfList.Row
                    j = vsfList.Col
                End If
                vsfList.Select i, j
                dtpAppointmentTime.Value = CDate(mtyRegPlanState.strSelTime)
                mblnStateChange = False
            End If
            With vsfList
                If Val(vsfPlan.TextMatrix(vsfPlan.Row, GetCol("限号"))) <= mrsSNState.RecordCount And InStr(mstrPrivs, ";加号;") <= 0 Then
                    '加号 是否有加号权限
                    MsgBox lngSN & "号超过最大限号数!你没有满号后继续挂号的权限.", vbInformation, gstrSysName
                    RestorePay
                    Exit Sub
                End If
                If vsfList.TextMatrix(vsfList.Row, vsfList.Col) <> "" And .Cell(flexcpForeColor, i, j) <> vbRed _
                   And .Cell(flexcpForeColor, i, j) <> vbBlue And .Cell(flexcpForeColor, i, j) <> vbGrayText _
                   And .Cell(flexcpForeColor, i, j) <> &HC000C0 And .Cell(flexcpForeColor, i, j) <> vbGreen _
                   Then
                    If Format(Get时段(i, j, True), "hh:mm:00") <> Format(dtpAppointmentTime.Value, "hh:mm:ss") Then
                        dtpAppointmentTime.Value = CDate(Get时段(i, j, True))
                    End If
                    lngSN = GetCurrSN(, True)
                    If lngSN = 0 Then mbln加号 = True
                Else
                    '存在过期的时段,此时没有达到限号数,此时没有达到限号数,增加的号,发生时间,为最后一个时段的结束时间
                    bln追加时段 = True
                End If
            End With
        End If
        
        Call Get时间(Datsys, lngSN, bln追加时段, str登记时间, str发生时间, dat发生时间, bln达到限号数)
        
        If mbytMode <> 2 And mstrNoIn = "" And mViewMode <> v_专家号分时段 Then
            If vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("挂号时间")) <> "" Then
                If Not (CDate(Format(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("挂号时间")), "yyyy-mm-dd hh:mm:ss")) > zlDatabase.Currentdate And CDate(Format(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("提前时间")), "yyyy-mm-dd hh:mm:ss")) < zlDatabase.Currentdate) Then
                    If Check有效时间段(lng记录ID, dat发生时间) = False Then
                        If chkShowAll.Value = 1 Then
                            If MsgBox("当前挂号号别不当班或者已经被停诊,你是否要继续挂号？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                Call RestorePay: Exit Sub
                            End If
                        Else
                            MsgBox "当前挂号号别不当班或者已经被停诊,不能继续挂号！", vbInformation, gstrSysName
                            Call RestorePay: Exit Sub
                        End If
                    End If
                End If
            Else
                If chkShowAll.Value = 1 Then
                    If MsgBox("当前挂号号别不当班或者已经被停诊,你是否要继续挂号？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Call RestorePay: Exit Sub
                    End If
                Else
                    MsgBox "当前挂号号别不当班或者已经被停诊,不能继续挂号！", vbInformation, gstrSysName
                    Call RestorePay: Exit Sub
                End If
            End If
        End If
        
        With vsfPlan
            If .TextMatrix(.Row, .ColIndex("替诊开始时间")) <> "" Then
                If .TextMatrix(.Row, .ColIndex("替诊医生")) <> "" And dat发生时间 >= CDate(.TextMatrix(.Row, .ColIndex("替诊开始时间"))) And dat发生时间 <= CDate(.TextMatrix(.Row, .ColIndex("替诊终止时间"))) Then
                    str医生姓名 = .TextMatrix(.Row, .ColIndex("替诊医生姓名"))
                    lng医生ID = .TextMatrix(.Row, .ColIndex("替诊医生ID"))
                Else
                    str医生姓名 = mstr医生姓名
                    lng医生ID = mlng医生ID
                End If
            Else
                str医生姓名 = mstr医生姓名
                lng医生ID = mlng医生ID
            End If
        End With
        
        If cbo预约方式.Visible And (mbytMode = 1 Or chkBooking.Value = 1) Then
            strSQL = "Select Zl_Fun_Get临床出诊预约状态([1],[2],[3],[4],[5],[6]) As 预约检查 From Dual"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng记录ID, dat发生时间, lngSN, NeedName(cbo预约方式.Text, , "."), "", IIf(chkBooking.Value = 1, 1, 0))
            If rsTemp.EOF Then
                MsgBox "当前选择的号码无法预约,请选择其他号码!", vbInformation, gstrSysName
                If cbo预约方式.Enabled And cbo预约方式.Visible Then cbo预约方式.SetFocus
                Call RestorePay: Exit Sub
            Else
                If Val(Mid(Nvl(rsTemp!预约检查), 1, 1)) <> 0 Then
                    MsgBox "当前选择的号码无法预约,请选择其他号码!" & vbCrLf & "原因:" & Mid(Nvl(rsTemp!预约检查), InStr(Nvl(rsTemp!预约检查), "|") + 1), vbInformation, gstrSysName
                    If cbo预约方式.Enabled And cbo预约方式.Visible Then cbo预约方式.SetFocus
                    Call RestorePay: Exit Sub
                End If
            End If
        End If
        
        '137272:李南春,2019/2/20,对序号进行锁号,如果序号不可用则返回一个有效的序号
        If ReserveRegNo(lng记录ID, lngSN, str发生时间, Datsys) = False Then Exit Sub
        
        str卡号 = Trim(str卡号)
        If blnPati Then
            With mobjfrmPatiInfo
                If .txt出生时间 = "__:__" Then
                    str出生日期 = IIf(IsDate(.txt出生日期.Text), "TO_Date('" & .txt出生日期.Text & "','YYYY-MM-DD')", "NULL")
                Else
                    str出生日期 = IIf(IsDate(.txt出生日期.Text), "TO_Date('" & .txt出生日期.Text & " " & .txt出生时间.Text & "','YYYY-MM-DD HH24:MI:SS')", "NULL")
                End If
                str联系电话 = Trim(txt家庭电话.Text)
                str身份证号 = Trim(txtIDCard.Text)
                strSQL = _
                "zl_挂号病人病案_INSERT(" & bytType & "," & lng病人ID & "," & IIf(str门诊号 = "", "NULL", str门诊号) & "," & _
                         IIf(str卡号 = "" Or mCurSendCard.bln就诊卡 = False, "NULL", "'" & str卡号 & "'") & ",'" & str密码 & "','" & txtPatient.Text & "','" & NeedName(cbo性别.Text) & "'," & _
                         "'" & str年龄 & "','" & str费别 & "','" & NeedName(cbo付款方式.Text) & "'," & _
                         "'" & NeedName(.cbo国籍.Text) & "','" & NeedName(.cbo民族.Text) & "','" & NeedName(.cbo婚姻.Text) & "'," & _
                         "'" & NeedName(.cbo职业.Text, True) & "','" & str身份证号 & "','" & .txt单位名称.Text & "'," & _
                         Val(.txt单位名称.Tag) & ",'" & .txt单位电话.Text & "','" & .txt单位邮编.Text & "','" & IIf(mblnStructAdress, padd家庭地址.Value, cbo家庭地址.Text) & "'," & _
                         "'" & str联系电话 & "','" & .txt家庭邮编.Text & "'," & str登记时间 & ",''," & str出生日期 & ",'" & strMCAccount & _
                         "', " & IIf(str卡号 = "", "NULL", "'" & IIf(mblnICCard, str卡号, "") & "'") & "," & ZVal(mintInsure) & "," & _
                         IIf(Trim(.txt区域.Text) = "", "NULL,", "'" & Trim(.txt区域.Text) & "',") & _
                          "'" & IIf(mblnStructAdress, Trim(padd户口地址.Value), Trim(cbo户口地址.Text)) & "','" & Trim(mobjfrmPatiInfo.txt户口地址邮编.Text) & "'," & IIf(Trim(mobjfrmPatiInfo.txt联系人身份证.Text) = "", "NULL,", "'" & Trim(mobjfrmPatiInfo.txt联系人身份证.Text) & "',") & _
                         IIf(Trim(mobjfrmPatiInfo.txt联系人姓名.Text) = "", "NULL,", "'" & Trim(mobjfrmPatiInfo.txt联系人姓名.Text) & "',") & _
                         IIf(Trim(mobjfrmPatiInfo.txt联系人电话.Text) = "", "NULL,", "'" & Trim(mobjfrmPatiInfo.txt联系人电话.Text) & "',") & _
                         IIf(NeedName(mobjfrmPatiInfo.cbo联系人关系.Text) = "", "NULL,", "'" & NeedName(mobjfrmPatiInfo.cbo联系人关系.Text) & "',")
                strSQL = strSQL & IIf(Trim(mobjfrmPatiInfo.txt监护人.Text) = "", "NULL,", "'" & Trim(mobjfrmPatiInfo.txt监护人.Text) & "',")  'lgf
                strSQL = strSQL & IIf(Trim(mobjfrmPatiInfo.txtBirthLocation.Text) = "", "NULL,", "'" & Trim(mobjfrmPatiInfo.txtBirthLocation.Text) & "',")
                strSQL = strSQL & "'" & mobjfrmPatiInfo.txtMobile.Text & "')"
                Call zlAddArray(cllProBefor, strSQL)
                
                '90875:李南春,2016/11/8,医疗卡证件类型
                If AddCertificate(lng病人ID, cllProBefor, Datsys) = False Then Exit Sub
                
                '89242:李南春,2015/12/7,更新病人地址信息
                If mblnStructAdress Then
                    If padd家庭地址.Value <> "" Then
                       strSQL = "zl_病人地址信息_update(1," & lng病人ID & ",NULL,3,'" & padd家庭地址.value省 & "','" & _
                           padd家庭地址.value市 & "','" & padd家庭地址.value区县 & "','" & padd家庭地址.value乡镇 & "','" & _
                           padd家庭地址.value详细地址 & "','" & padd家庭地址.Code & "')"
                    Else
                       strSQL = "zl_病人地址信息_update(2," & lng病人ID & ",NULL,3)"
                    End If
                    Call zlAddArray(cllProBefor, strSQL)
                    If padd户口地址.Value <> "" Then
                       strSQL = "zl_病人地址信息_update(1," & lng病人ID & ",NULL,4,'" & padd户口地址.value省 & "','" & _
                           padd户口地址.value市 & "','" & padd户口地址.value区县 & "','" & padd户口地址.value乡镇 & "','" & _
                           padd户口地址.value详细地址 & "','" & padd户口地址.Code & "')"
                    Else
                       strSQL = "zl_病人地址信息_update(2," & lng病人ID & ",NULL,4)"
                    End If
                    Call zlAddArray(cllProBefor, strSQL)
                End If
                If mobjfrmPatiInfo.txt联系人姓名.Text <> "" And NeedName(mobjfrmPatiInfo.cbo联系人关系.Text) = "其他" Then
                    strSQL = "Zl_病人信息从表_Update("
                    '病人ID_In 病人信息从表.病人Id%Type
                    strSQL = strSQL & "" & lng病人ID & ","
                    '信息名_In 病人信息从表.信息名%Type0
                    strSQL = strSQL & "'联系人附加信息',"
                    '信息值_In 病人信息从表.信息值%Type
                    strSQL = strSQL & "'" & mobjfrmPatiInfo.txt其他关系.Text & "',"
                    '就诊Id_In 病人信息从表.就诊Id%Type
                    strSQL = strSQL & "'')"
                    Call zlAddArray(cllProBefor, strSQL)
                End If
        
                If mlngOutModeMC > 0 And cbo医疗类别.ListIndex > 0 Then
                    strInfo = cbo医疗类别.Text: strInfo = Mid(strInfo, 1, InStr(1, strInfo, "-") - 1)
                    strSQL = "zl_就诊登记记录_UPDATE(" & mlngOutModeMC & "," & lng病人ID & ",0," & str登记时间 & ",0,'" & strInfo & "')"
                    Call zlAddArray(cllProBefor, strSQL)
                End If

                If mstr社区号 <> "" And mint社区 <> 0 Then
                    strSQL = "Zl_病人社区信息_Insert(" & lng病人ID & "," & mint社区 & ",'" & mstr社区号 & "',1," & str登记时间 & ")"
                    Call zlAddArray(cllProBefor, strSQL)
                End If
            End With
        End If

        strSQL = "Select ID as 就诊ID From 病人挂号记录 Where 记录状态 = 1 And NO=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
        If rsTemp.RecordCount > 0 Then lng就诊ID = Val(Nvl(rsTemp!就诊ID))
        Err = 0: On Error GoTo ErrFirt:
        '先保存病人信息,然后再处理其他,避免造成并发问题(主要是病人ID为重复
        zlExecuteProcedureArrAy cllProBefor, Me.Caption, True

        '101170:李南春,2016/10/13,保存HIS数据要提交EMPI数据，失败后所有数据都要回退
        If zlSaveEMPIPatiInfo(bytType = 1, lng病人ID, lng就诊ID, strErrMsg) = False Then
            gcnOracle.RollbackTrans
            If strErrMsg = "" Then strErrMsg = "向EMPI平台上传病人信息失败！"
            MsgBox strErrMsg, vbInformation, gstrSysName
            Exit Sub
        End If
        gcnOracle.CommitTrans

        Err = 0: On Error GoTo ErrGo:
        If mobjfrmPatiInfo.mblnSavePati = False Then
            Call mobjfrmPatiInfo.SavePatiPic(lng病人ID)
            If CreatePlugInOK(mlngModul) And mobjfrmPatiInfo.mlngPlugInHwnd <> 0 Then  '保存插件附加信息
                On Error Resume Next
                Call gobjPlugIn.PatiInfoSaveAfter(lng病人ID)
                Call zlPlugInErrH(Err, "PatiInfoSaveAfter")
                Err.Clear: On Error GoTo 0
            End If
        End If
        mobjfrmPatiInfo.mblnSavePati = False
        
        RegistFeeMode = mRegistFeeMode
        If mRegistFeeMode <> EM_RG_记帐 Then
            RegistFeeMode = EM_RG_现收
            If str结算方式 = "" Then RegistFeeMode = EM_RG_划价
        End If
        
        '处理卡费
        cur卡费 = 0                 '挂号同时发卡，必定只用现金结算，不涉及医保及预交款
        mCurSendCard.dbl应收金额 = 0
        mCurSendCard.dbl实收金额 = 0
        If Not mrsItems Is Nothing Then
            mrsItems.Filter = "性质=4"
            If mrsItems.RecordCount > 0 Then
                bln发卡 = True
                mrsInComes.Filter = "项目ID=" & mrsItems!项目ID
                Do While Not mrsInComes.EOF
                    cur卡费 = cur卡费 + mrsInComes!实收
                    mCurSendCard.dbl应收金额 = mrsInComes!应收 + mCurSendCard.dbl应收金额
                    mrsInComes.MoveNext
                Loop
                mCurSendCard.dbl实收金额 = cur卡费
                Call AddCardDataSQL(lng病人ID, Datsys, cllPro, lngCard结帐ID, (mRegistFeeMode = EM_RG_记帐), mrsItems!项目ID)
            ElseIf str卡号 <> "" Then
                '问题: 42947 绑定卡,也需要处理发卡记录
                bln发卡 = True    '问题号:56599
                Call AddCardDataSQL(lng病人ID, Datsys, cllPro, lngCard结帐ID)
            End If
        ElseIf str卡号 <> "" Then
            '问题: 42947 绑定卡,也需要处理发卡记录
            bln发卡 = True    '问题号:56599
            Call AddCardDataSQL(lng病人ID, Datsys, cllPro, lngCard结帐ID)
        End If
        
        '产生费用记录SQL语句
        '------------------------------------------------------------------------------
        If (mbytMode = 2 Or mbytMode = 0 And mbytInState = 0 And mstrNoIn <> "") Then
            '预约接收
            If SaveRegister_预约接收(lng病人ID, Datsys, str门诊号, str卡号, gCurSendCard.rs卡费, str发生时间, str登记时间, blnPrice, blnNoPrint, lngSN, _
                str结算方式, cur现金, cur卡费, cur预交, cur个帐, lng结帐ID, cllPro, cllProAfter, lngCard结帐ID, bln发卡) = False Then Exit Sub
        Else
            If mobjfrmPatiInfo.txt支付密码 <> "" And mobjfrmPatiInfo.txt身份证号 <> "" And mbytMode <> 1 Then    '专门针对【二代身份证】这种情况进行绑定
                bln发卡 = True    '问题号:56999
                Call AddSQL绑定卡(lng病人ID, Val(mobjfrmPatiInfo.txt支付密码.Tag), mobjfrmPatiInfo.txt身份证号, zlCommFun.zlStringEncode(mobjfrmPatiInfo.txt支付密码), Datsys, mblnICCard, cllPro)
            End If
            If txt号别.Text = "+" Then lngSN = 0
            
            dblTotal(0) = GetRegistMoney(True, False)
            dblTotal(1) = GetCardMoney  '卡费
            If dblTotal(0) <> 0 And blnPrice Then
                '挂号费存为零且保存为划价单，才产生划价NO
                   str划价NO = zlDatabase.GetNextNo(13)
            End If
            
            mrsItems.Filter = ""
            k = 1: mrsItems.MoveFirst
            For i = 1 To mrsItems.RecordCount
                int价格父号 = k
                mrsInComes.Filter = "项目ID=" & mrsItems!项目ID
                For j = 1 To mrsInComes.RecordCount
                    '卡费
                    If mrsItems!性质 = 4 Then   '读费用集时已限制仅有一行,不支持设置多个收入项目,为了保持与就诊卡管理中一致
                        '
                    Else
                        '挂号收费数据
                        strSQL = _
                        "zl_病人挂号记录_出诊_INSERT(" & ZVal(lng记录ID) & "," & ZVal(lng病人ID) & "," & IIf(str门诊号 = "", "NULL", str门诊号) & ",'" & txtPatient.Text & "','" & NeedName(cbo性别.Text) & "'," & _
                                 "'" & str年龄 & "','" & NeedCode(cbo付款方式.Text) & "','" & str费别 & "','" & strNO & "'," & _
                                 "'" & IIf(blnNoPrint, "", txtFact.Text) & "'," & k & "," & IIf(int价格父号 = k, "NULL", int价格父号) & "," & IIf(mrsItems!性质 = 2, 1, "NULL") & "," & _
                                 "'" & mrsItems!类别 & "'," & mrsItems!项目ID & "," & mrsItems!数次 & "," & mrsInComes!单价 & "," & _
                                 mrsInComes!收入项目ID & ",'" & mrsInComes!收据费目 & "','" & str结算方式 & "'," & _
                                 IIf(blnPrice And dblTotal(0) <> 0, 0, mrsInComes!应收) & "," & IIf(blnPrice And dblTotal(0) <> 0, 0, mrsInComes!实收) & "," & _
                                 mlng挂号科室ID & "," & IIf(mblnStation, mlng挂号科室ID, UserInfo.部门ID) & "," & IIf(mrsItems!执行科室ID = 0, mlng挂号科室ID, mrsItems!执行科室ID) & ",'" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & _
                                 str发生时间 & "," & str登记时间 & "," & _
                                 "'" & str医生姓名 & "'," & ZVal(lng医生ID) & "," & IIf(mrsItems!性质 = 3, 1, IIf(mrsItems!性质 = 4, 2, 0)) & "," & IIf(lbl急.Visible, 1, 0) & "," & _
                                 "'" & IIf(txt号别.Text = "+", "", txt号别.Text) & "','" & strRoom & "'," & ZVal(lng结帐ID) & "," & IIf(blnNoPrint, "NULL", ZVal(mlng领用ID)) & "," & _
                                 ZVal(IIf(mbytMode <> 1 And k = 1, cur预交, 0)) & "," & ZVal(IIf(mbytMode <> 1 And k = 1 And Not blnPrice, cur现金 - cur卡费, 0)) & "," & _
                                 ZVal(IIf(mbytMode <> 1 And k = 1, cur个帐, 0)) & "," & ZVal(Nvl(mrsItems!保险大类id, 0)) & "," & _
                                 ZVal(Nvl(mrsItems!保险项目否, 0)) & "," & ZVal(Nvl(mrsInComes!统筹金额, 0)) & "," & _
                                 "'" & IIf(str划价NO <> "", "划价:" & str划价NO, Me.cbo备注.Text) & "'," & IIf(mbytMode = 1, 1, 0) & "," & IIf(gblnSharedInvoice, 1, 0) & ",'" & mrsItems!保险编码 & "'," & byt复诊 & "," & ZVal(lngSN) & "," & ZVal(mint社区) & "," & _
                                 IIf(mbytMode = 2 Or chkBooking.Value = 1 Or mbytMode = 1, 1, 0) & "," & IIf(mbytMode = 1 Or chkBooking.Value = 1, "'" & Mid(cbo预约方式.Text, InStr(cbo预约方式.Text, ".") + 1) & "'", "NULL") & "," & _
                                 IIf(mTy_Para.bln挂号生成队列, 1, 0) & ","
                        
                        '卡类别id_In   病人预交记录.卡类别id%Type := Null,
                        strSQL = strSQL & "" & IIf(mCurCardPay.Have挂号费 And mCurCardPay.lng医疗卡类别ID <> 0 And mCurCardPay.bln消费卡 = False, mCurCardPay.lng医疗卡类别ID, "NULL") & ","
                        '结算卡序号_In 病人预交记录.结算卡序号%Type := Null,
                        strSQL = strSQL & "" & IIf(mCurCardPay.Have挂号费 And mCurCardPay.lng医疗卡类别ID <> 0 And mCurCardPay.bln消费卡, mCurCardPay.lng医疗卡类别ID, "NULL") & ","
                        '卡号_In       病人预交记录.卡号%Type := Null,
                        strSQL = strSQL & "" & IIf(mCurCardPay.Have挂号费 And mCurCardPay.str刷卡卡号 <> "", "'" & mCurCardPay.str刷卡卡号 & "'", "NULL") & ","
                        '交易流水号_In 病人预交记录.交易流水号%Type := Null,
                        strSQL = strSQL & " NULL,"
                        '交易说明_In   病人预交记录.交易说明%Type := Null,
                        strSQL = strSQL & " NULL,"
                        '合作单位_In   病人预交记录.合作单位%Type := Null
                        strSQL = strSQL & " NULL,"
                        '  操作类型_In   Number:=0
                        strSQL = strSQL & IIf(mbln加号, "1", "0") & ","
                        '  险类_IN       病人挂号记录.险类%type:=null,
                        strSQL = strSQL & IIf(mintInsure = 0, "NULL", mintInsure) & ","
                        '  结算模式_IN   NUMBER :=0,
                        strSQL = strSQL & IIf(mPatiChargeMode = EM_先诊疗后结算, 1, 0) & ","
                        '  记帐费用_IN Number:=0,
                        strSQL = strSQL & IIf(mRegistFeeMode = EM_RG_记帐, 1, 0) & ","
                        '  退号重用_IN Number:=1,
                        strSQL = strSQL & IIf(mTy_Para.blnReuseCancelNO, 1, 0) & ","
                        '  冲预交病人ids_In Varchar2 := Null
                        strSQL = strSQL & "'" & lng病人ID & "," & mstr病人家属IDs & "'," '79868,冉俊明,2015-6-15,使用家属预交
                        '  修正病人费别_In  Number := 0,
                        strSQL = strSQL & 0 & ","
                        '  预约顺序号_In    临床出诊序号控制.预约顺序号%Type := Null,
                        strSQL = strSQL & "Null,"
                        '  修正病人年龄_In  Number := 0,
                        strSQL = strSQL & "0,'"
                        '  收费单_In        病人挂号记录.收费单%Type := Null
                        strSQL = strSQL & str划价NO & "')"
                        
                        Call zlAddArray(cllPro, strSQL)
                        
                        If Trim(IIf(txt号别.Text = "+", "", txt号别.Text)) <> "" And k = 1 Then
                            If Nvl(vsfPlan.TextMatrix(vsfPlan.Row, GetCol("医生"))) = "" Then blnNoDoc = True
                            strSQL = "zl_病人挂号汇总_Update("
                            '  医生姓名_In   挂号安排.医生姓名%Type,
                            strSQL = strSQL & IIf(blnNoDoc, "Null,", "'" & str医生姓名 & "',")
                            '  医生id_In     挂号安排.医生id%Type,
                            strSQL = strSQL & "" & IIf(blnNoDoc, "0,", ZVal(lng医生ID) & ",")
                            '  收费细目id_In 门诊费用记录.收费细目id%Type,
                            strSQL = strSQL & "" & Val(Nvl(mrsItems!项目ID)) & ","
                            '  执行部门id_In 门诊费用记录.执行部门id%Type,
                            strSQL = strSQL & "" & IIf(Val(Nvl(mrsItems!执行科室ID)) = 0, mlng挂号科室ID, Val(Nvl(mrsItems!执行科室ID))) & ","
                            '  发生时间_In   门诊费用记录.发生时间%Type,
                            strSQL = strSQL & "" & str发生时间 & ","
                            '  预约标志_In   Number := 0  --是否为预约接收:0-非预约挂号; 1-预约挂号,2-预约接收,3-收费预约
                            strSQL = strSQL & Decode(mbytMode, 1, 1, 2, 2, IIf(chkBooking.Value = 1, 3, 0)) & ","
                            '  号码_In       挂号安排.号码%Type := Null
                            strSQL = strSQL & "'" & IIf(txt号别.Text = "+", "", txt号别.Text) & "',0,"
                            strSQL = strSQL & "" & ZVal(lng记录ID) & ")"
                            Call zlAddArray(cllProAfter, strSQL)
                        End If
                        '门诊医生站挂号时,如果是现金支付则生成划价单,此时应收/实收填写为0,摘要填写为挂号单据号
                        If blnPrice And dblTotal(0) <> 0 Then
                            strSQL = _
                            "zl_门诊划价记录_Insert('" & str划价NO & "'," & k & "," & lng病人ID & ",NULL," & _
                                     IIf(str门诊号 = "", "NULL", str门诊号) & ",'" & NeedCode(cbo付款方式.Text) & "'," & _
                                     "'" & txtPatient.Text & "','" & NeedName(cbo性别.Text) & "','" & str年龄 & "'," & _
                                     "'" & str费别 & "',NULL," & mlng挂号科室ID & "," & _
                                     IIf(mblnStation, mlng挂号科室ID, UserInfo.部门ID) & ",'" & UserInfo.姓名 & "'," & IIf(mrsItems!性质 = 2, 1, "NULL") & "," & _
                                     mrsItems!项目ID & ",'" & mrsItems!类别 & "','" & mrsItems!计算单位 & "'," & _
                                     "NULL,1," & mrsItems!数次 & ",NULL," & IIf(mrsItems!执行科室ID = 0, mlng挂号科室ID, mrsItems!执行科室ID) & "," & IIf(int价格父号 = k, "NULL", int价格父号) & "," & _
                                     mrsInComes!收入项目ID & ",'" & mrsInComes!收据费目 & "'," & mrsInComes!单价 & "," & _
                                     mrsInComes!应收 & "," & mrsInComes!实收 & "," & str发生时间 & "," & str登记时间 & ",NULL,'" & UserInfo.姓名 & "','挂号:" & strNO & "')"
                            Call zlAddArray(cllPro, strSQL)
                        End If
                    End If
                    k = k + 1
                    mrsInComes.MoveNext
                Next
                mrsItems.MoveNext
            Next
        End If
 
        cmdOK.Enabled = False      '防止打印弹出设置打印机的非模态窗体及医保结算延迟
        '执行处理
        Err = 0: On Error GoTo ErrFirt:
        If cllPro.Count > 0 Then
            '问题:31187 在事务当中处理过程数据
            Err = 0: On Error GoTo ErrFirt:
            zlExecuteProcedureArrAy cllPro, Me.Caption, True
            
            '金额检查
            If lng结帐ID <> 0 Then
                strSQL = "Select Sum(结帐金额) As 费用金额 From 门诊费用记录 Where 记录性质=4 And 结帐ID=[1]"
                Set rsCheck = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng结帐ID)
                If Not rsCheck.EOF Then
                    dbl费用金额 = Val(Nvl(rsCheck!费用金额))
                    strSQL = "Select Sum(冲预交) As 结帐金额 From 病人预交记录 Where 结帐ID=[1]"
                    Set rsCheck = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng结帐ID)
                    If Not rsCheck.EOF Then
                        If dbl费用金额 <> Val(Nvl(rsCheck!结帐金额)) Then
                            gcnOracle.RollbackTrans
                            MsgBox "结算信息与费用信息保存不一致，请重新提取数据再试!", vbInformation, gstrSysName
                            cmdOK.Enabled = True: RestorePay: Exit Sub
                        End If
                    Else
                        If dbl费用金额 <> 0 Then
                            gcnOracle.RollbackTrans
                            MsgBox "结算信息与费用信息保存不一致，请重新提取数据再试!", vbInformation, gstrSysName
                            cmdOK.Enabled = True: RestorePay: Exit Sub
                        End If
                    End If
                End If
            End If

            Err = 0: On Error GoTo errH:
            blnTrans = True
            If curOneCard <> 0 And mRegistFeeMode <> EM_RG_记帐 Then
                If Not (curOneCard = cur卡费 And cur卡费 <> 0) Then    '不只是卡费时
                    If Not mobjICCard.PaymentSwap(curOneCard - cur卡费, dblOneCardBalance, intCardType, Val("" & mrsOneCard!医院编码), strCardNo, strTransFlow, lng结帐ID, lng病人ID) Then
                        gcnOracle.RollbackTrans
                        MsgBox "一卡通结算挂号费失败", vbInformation, gstrSysName
                        RestorePay
                        cmdOK.Enabled = True: Exit Sub
                    Else
                        strSQL = "zl_一卡通结算_Update(" & lng结帐ID & ",'" & mrsOneCard!结算方式 & "','" & strCardNo & "','" & intCardType & "','" & strTransFlow & "'," & dblOneCardBalance & ")"
                        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
                    End If
                End If

                If cur卡费 <> 0 Then
                    dblOneCardBalance = 0
                    strTransFlow = ""
                    If Not mobjICCard.PaymentSwap(cur卡费, dblOneCardBalance, intCardType, Val("" & mrsOneCard!医院编码), strCardNo, strTransFlow, lngCard结帐ID, lng病人ID) Then
                        gcnOracle.RollbackTrans
                        MsgBox "一卡通结算卡费失败", vbInformation, gstrSysName
                        RestorePay
                        cmdOK.Enabled = True: Exit Sub
                    Else
                        strSQL = "zl_一卡通结算_Update(" & lngCard结帐ID & ",'" & mrsOneCard!结算方式 & "','" & strCardNo & "','" & intCardType & "','" & strTransFlow & "'," & dblOneCardBalance & ")"
                        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
                    End If
                End If
            End If

            '医保改动
            blnNotCommit = False
            If mintInsure <> 0 And mstrYBPati <> "" Then
                strAdvance = ""
                If mRegistFeeMode = EM_RG_记帐 Or mPatiChargeMode = EM_先诊疗后结算 Then
                    strAdvance = IIf(mPatiChargeMode = EM_先诊疗后结算, "1", "0") & "|" & IIf(mRegistFeeMode = EM_RG_记帐, "1", "0") & "|" & strNO
                End If
                If Not gclsInsure.RegistSwap(lng结帐ID, cur个帐, mintInsure, strAdvance) Then
                    gcnOracle.RollbackTrans: cmdOK.Enabled = True: RestorePay: Exit Sub
                End If
                blnNotCommit = True
            End If
            zlExecuteProcedureArrAy cllProAfter, Me.Caption, True, True
            Set cllCardPro = New Collection: Set cllTheeSwap = New Collection
            If Not mPatiChargeMode = EM_先诊疗后结算 Then
                If zlInterfacePrayMoney(lngCard结帐ID, lng结帐ID, cllCardPro, cllTheeSwap, dblThreeSwap) = False Then
                    gcnOracle.RollbackTrans: If cmdOK.Enabled = False Then cmdOK.Enabled = True: RestorePay: Exit Sub
                End If
                '修正三方交易
                zlExecuteProcedureArrAy cllCardPro, Me.Caption, True, True
            End If
            gcnOracle.CommitTrans
            
            Call zlExcPatiInfo(lng病人ID, lng就诊ID, strNO)
            
            Err = 0: On Error GoTo OthersCommit:
            zlExecuteProcedureArrAy cllTheeSwap, Me.Caption, True, False
OthersCommit:
            gcnOracle.CommitTrans
            '写卡操作
            If bln发卡 And mCurSendCard.bln是否写卡 Then Call WriteCard(lng病人ID)
            If mintInsure > 0 Then Call gclsInsure.BusinessAffirm(交易Enum.Busi_RegistSwap, True, mintInsure)
            Dim strOutPut As String
            Call zlExcuteUploadSwap(lng病人ID, strOutPut, mobjICCard)
            
            blnTrans = False
            On Error GoTo 0
            '医保结算后的语音报价
            If mintInsure <> 0 And mstrYBPati <> "" And Not blnPrice And mRegistFeeMode <> EM_RG_记帐 Then
                '如果是医保病人,需要重新获取本次结算的现收金额
                cur应缴 = GetActualCash(lng结帐ID)
                If gblnLED And mbytMode <> 1 And mbytInState = 0 Then
                    zl9LedVoice.Speak "#21 " & Format(cur应缴, "0.00")
                    txt找补.Text = Format(Val(txt缴款.Text) - cur应缴, "0.00")
                End If
            End If
        End If
        If str卡号 <> "" Then
            Call zlCommitPlugInpati(str卡号)
        End If
        '消息传送:
        Call SendMsgModule(strNO)
        '打印单据
        If mbytMode <> 1 And Not blnNoPrint Then
RePrint:
            Dim strNotValiedNos As String
            If Not gobjTax Is Nothing And gblnTax Then
                Call TaxInterface(1, "'" & strNO & "'", "")
            Else
                If mRegistFeeMode <> EM_RG_记帐 Then
                    blnEnterPrint = True
                    Call frmPrint.ReportPrint(1, strNO, "", mlng领用ID, mlngShareUseID, txtFact.Text, Datsys, txt缴款.Text, txt找补.Text, , mintInsure <> 0 And MCPAR.医保接口打印票据, False, mstrUseType)
                    If gblnBill挂号 Then
                        If zlIsNotSucceedPrintBill(4, strNO, strNotValiedNos) = True Then
                            If MsgBox("挂号单号为[" & strNotValiedNos & "]票据打印未成功,是否重新进行票据打印!", vbYesNo + vbDefaultButton1 + vbQuestion, gstrSysName) = vbYes Then GoTo RePrint:
                        End If
                    End If
                End If
            End If
        ElseIf blnPrintBooking And mbytMode = 1 Then
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1111_1", Me, "NO=" & strNO, 2)
        End If
        
        If mbytMode <> 1 And gblnPrintCase Then
            '新增病人的情况 问题号：42452 修改人:王吉
            If chk病历费.Value = 1 And blnPati = True And bytType = 1 Then
                Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1111_2", Me, "病人ID=" & lng病人ID, 2)
            ElseIf chk病历费.Value = 1 Or Trim(txt号别.Text) = "+" Then
                Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1111_2", Me, "病人ID=" & lng病人ID, 2)
            End If
        End If
        
        If blnSlipPrint Then
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1111_3", Me, "NO=" & strNO, 2)
            If Not blnEnterPrint Then
                strSQL = "Zl_凭条打印记录_Update(4,'" & strNO & "',1,'" & UserInfo.姓名 & "','发票号:" & txtFact.Text & "')"
                zlDatabase.ExecuteProcedure strSQL, "凭条打印记录"
            End If
        End If
        
        If blnCodePrint Then
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1111_2", Me, "病人ID=" & lng病人ID, 2)
        End If
        
        If CreatePlugInOK(mlngModul) Then
            On Error Resume Next
            strSQL = "Select ID From 病人挂号记录 Where no=[1] And Rownum<2"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
            If Not rsTemp.EOF Then Call gobjPlugIn.OutPatiRegisterAfter(lng病人ID, Nvl(rsTemp!ID))
            Err.Clear
        End If
        
        cmdOK.Enabled = True
        '预约接收后退出
        If mbytMode = 2 Then
            If Not gblnBill挂号 And Not blnNoPrint And mRegistFeeMode <> EM_RG_记帐 Then
                If gblnSharedInvoice Then
                    zlDatabase.SetPara "当前收费票据号", txtFact.Text, glngSys, 1121
                Else
                    zlDatabase.SetPara "当前挂号票据号", txtFact.Text, glngSys, mlngModul
                End If
            End If
            gblnOk = True:
            mblnUnload = True
            Call ClearBill
            mblnUnload = False
            Unload Me: Exit Sub
        ElseIf mbytMode = 0 And mbytInState = 0 And mstrNoIn <> "" Then
            Call SetReceiveState(False)
            cmdYb.Visible = mblnRegReceiveByNo
            blnAfterRefresh = True
        End If

        '加入单据历史记录(所有类型单据)
        If strNO <> "" Then
            For i = 0 To cboNO.ListCount - 1
                strNO = strNO & "," & cboNO.List(i)
            Next
            cboNO.Clear
            For i = 0 To UBound(Split(strNO, ","))
                cboNO.AddItem Split(strNO, ",")(i)
                If i = 9 Then Exit For    '只显示10个
            Next
            If cboNO.ListCount > 0 Then cboNO.ListIndex = 0
        End If
        blnNew = True: strFact = txtFact.Text
        If blnNoPrint Then blnNew = False    '不打印时,非严格控制的票据不增加号
    End If
    gblnOk = True
    Call SetControlChk
    '保存病人及累计信息的条件:参数要求输缴款后才结束,当前未输缴款,并且是非医保病人,输入了姓名,
    '并且本地参数要求输姓名(否则ClearBill中调用SetPatiInfoEnabled时会清除姓名)
    '刘兴洪:26602
    ' 现增加对医保病人进行连续挂号,医保病人条件为:
    '   1.参数要求输入缴款金额后，终止连续收费
    '   2.需要参数:support连续挂号
    Dim blnClearInsure As Boolean
    blnClearInsure = True
    If mintInsure <> 0 And mstrYBPati <> "" Then
        bln连续 = gclsInsure.GetCapability(support连续挂号, lng病人ID, mintInsure)
        bln连续 = mTy_Para.byt缴款方式 = 1 And mbytMode <> 1 And Val(txt缴款.Text) = 0 And txtPatient.Text <> "" And bln连续
        blnClearInsure = Not bln连续
        Dim cur找补 As Currency, cur缴款 As Currency

        If blnCall结束挂号 Then
            If mstr连续挂号_挂号NO <> "" Then mstr连续挂号_挂号NO = Mid(mstr连续挂号_挂号NO, 2)
            If mstr连续挂号_就诊卡NO <> "" Then mstr连续挂号_就诊卡NO = Mid(mstr连续挂号_就诊卡NO, 2)
            txt本次应缴.Visible = False: lbl应缴.Visible = False: lbl缴款.Visible = False: txt缴款.Visible = False: lbl找补.Visible = False: txt找补.Visible = False
            lblSum.Visible = False: txt合计.Visible = False
            picTotal.Visible = True
            If frmYbPayFeeShow.zlShowPayWindows(Me, gclsInsure, gblnLED, txtPatient.Text, cbo性别.Text, txt年龄.Text & cbo年龄单位.Text, lng病人ID, mintInsure, mstr连续挂号_挂号NO, mstr连续挂号_就诊卡NO, mcur合计 + GetRegistMoney, mcur应缴 + cur应缴, cur缴款, cur找补) Then
                txt本次应缴.Text = Format(mcur应缴 + cur应缴, "0.00")
                txt缴款.Text = Format(cur缴款, "0.00")
                txt找补.Text = Format(cur找补, "0.00")
                bln连续 = False
            End If
            txt本次应缴.Visible = True: lbl应缴.Visible = True: lbl缴款.Visible = True: txt缴款.Visible = True: lbl找补.Visible = True: txt找补.Visible = True
            lblSum.Visible = True: txt合计.Visible = True
            picTotal.Visible = False
        End If
    Else
        bln连续 = mTy_Para.byt缴款方式 = 1 And mbytMode <> 1 And Val(txt缴款.Text) = 0 And mstrYBPati = "" And txtPatient.Text <> ""
    End If
    
    If Not mbln连续挂号 Then
        mcur合计 = 0: mcur应缴 = 0: mint挂号数 = 0
        mstrPrePati = "": mstr连续挂号_挂号NO = "": mstr连续挂号_就诊卡NO = ""
        lng病人ID = 0
        mblnFinishReg = True
        Call ClearBill(, Not blnNoPrint)
        mblnFinishReg = False
    Else
        If Not blnPrice Then
            mcur合计 = mcur合计 + GetRegistMoney
            mcur应缴 = mcur应缴 + cur连续
        End If
        mstrPrePati = txtPatient.Text
        '
        Call ClearBill(False, Not blnNoPrint, False)  '根据参数,如果不要求输姓名,或者号别不建病案,则会清除病人姓名
        mint挂号数 = mint挂号数 + 1
        '刘兴洪:如果是医保病人,需要重新获取余额
        If mintInsure <> 0 And mstrYBPati <> "" Then
            mcur个帐余额 = gclsInsure.SelfBalance(lng病人ID, CStr(Split(mstrYBPati, ";")(1)), 10, mcur个帐透支, mintInsure)
            stbThis.Panels(3).Text = "个人帐户余额:" & Format(mcur个帐余额, "0.00")
            mdbl个帐余额 = mcur个帐余额
        End If
    End If

    '刷新票据号
    If mbytMode <> 1 And Not mblnStation And Not blnPrice Then
        If blnNoPrint = False Then Call RefreshFact
    End If

    '对于输入的信息病人或刚建信息的病人下一张单子时保留病人信息(本地参数要求病人姓名时)
    If lng病人ID > 0 And chkCancel.Value = 0 And txtPatient.Enabled Then
        Call GetPatient(IDKind.GetCurCard, "-" & lng病人ID, False)
    End If

    timPlan.Enabled = glngInterval > 0 And mbytInState = 0 And (mbytMode = 0 Or mbytMode = 1)

    '刷新当前序号,ClearBill中已调用txt号别_change
    If txt号别.Enabled And txt号别.Visible Then txt号别.SetFocus
    mblnRegReceiveByNo = False
    If blnAfterRefresh Then
        Call RefreshFace
    End If
    Exit Sub
ErrFirt:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
    timPlan.Enabled = glngInterval > 0 And mbytInState = 0 And (mbytMode = 0 Or mbytMode = 1)
    cmdOK.Enabled = True
    Exit Sub
errH:
    '问题:31634
    If mintInsure > 0 And blnNotCommit Then Call gclsInsure.BusinessAffirm(交易Enum.Busi_RegistSwap, False, mintInsure)
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    timPlan.Enabled = glngInterval > 0 And mbytInState = 0 And (mbytMode = 0 Or mbytMode = 1)
    cmdOK.Enabled = True
    mbln加号 = False
    Exit Sub
ErrGo:
    If ErrCenter() = 1 Then
        Resume
    End If
    timPlan.Enabled = glngInterval > 0 And mbytInState = 0 And (mbytMode = 0 Or mbytMode = 1)
    cmdOK.Enabled = True
End Sub

Private Function GetPatiInforXML() As String
    Dim strPatiInforXML As String, str年龄 As String, str出生日期 As String, str身份证号 As String
    
    strPatiInforXML = strPatiInforXML & "<XM>" & Trim(txtPatient.Text) & "</XM>" & vbCrLf
    strPatiInforXML = strPatiInforXML & "<XB>" & NeedName(cbo性别.Text) & "</XB>" & vbCrLf
    str年龄 = Trim(txt年龄.Text)
    If IsNumeric(str年龄) Then str年龄 = str年龄 & cbo年龄单位.Text
    strPatiInforXML = strPatiInforXML & "<NL>" & str年龄 & "</NL>" & vbCrLf
    If IsDate(txt出生日期.Text) Then
        str出生日期 = Format(txt出生日期.Text & IIf(txt出生时间 = "__:__", "", " " & txt出生时间.Text), "yyyy-mm-dd HH:mm:ss")
    End If
    strPatiInforXML = strPatiInforXML & "<CSRQ>" & str出生日期 & "</CSRQ>" & vbCrLf
    strPatiInforXML = strPatiInforXML & "<YBH>" & mobjfrmPatiInfo.txtPatiMCNO(0).Text & "</YBH>" & vbCrLf
    If txtIDCard.Text <> "" And txtIDCard.Visible Then str身份证号 = Trim(txtIDCard.Text)
    strPatiInforXML = strPatiInforXML & "<SFZH>" & str身份证号 & "</SFZH>"
    strPatiInforXML = strPatiInforXML & "<YSXM>" & NeedName(cbo医生.Text) & "</YSXM>"
    
    GetPatiInforXML = strPatiInforXML
End Function

Private Sub SetControlChk()
    mstrPreNO = txt号别.Text
    cboNO.Tag = ""
    If chkCancel.Value = 1 Then chkCancel.Value = 0
    If chkPrint.Value = 1 Then chkPrint.Value = 0
    If chkBooking.Value = 1 Then
        chkBooking.Tag = "保存"
        chkBooking.Value = 0
        chkBooking.Tag = ""
    End If
End Sub

Private Sub zlExcPatiInfo(ByVal lng病人ID As Long, ByVal lng就诊ID As Long, ByVal strNO As String)
    Dim cllPro As Collection, Datsys As Date
    Dim strSQL As String, rsTemp As ADODB.Recordset
    '82072:李南春,2015/1/23,血型和RH增加一条有就诊ID的记录
    '.,所以将病人信息从表转移到这里
    
    On Error GoTo Errhand
    If lng病人ID > 0 And Not ((mbytMode = 2 Or mbytMode = 0 And mbytInState = 0 And mstrNoIn <> "") And mTy_Para.bln预约接收确定挂号费 = False) Then
        Set cllPro = New Collection
        Datsys = zlDatabase.Currentdate
        If lng就诊ID = 0 Then
            strSQL = "Select ID as 就诊ID From 病人挂号记录 Where 记录状态 = 1 And NO=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
            If rsTemp.RecordCount > 0 Then lng就诊ID = Nvl(rsTemp!就诊ID, 0)
        End If
        Call mobjfrmPatiInfo.Add健康卡相关信息(lng病人ID, cllPro, lng就诊ID)
        '保存病人信息中的证件
        Call mobjfrmPatiInfo.AddCertificate(lng病人ID, cllPro, Datsys)
        zlExecuteProcedureArrAy cllPro, Me.Caption
    End If
    Exit Sub
Errhand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function WriteCard(lng病人ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:写卡
    '入参:lng病人ID - 病人ID
    '编制:王吉
    '问题:56599
    '日期:2012-12-17 15:07:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strExpend As String
    On Error GoTo ErrHandl:
    '115168:李南春，2017/12/13，保存发卡的医疗卡类型
    If mCurSendCard.lng卡类别ID = 0 Then mCurSendCard = gCurSendCard
    If mCurSendCard.bln是否写卡 = False Then Exit Function
    If Not gobjSquare.objSquareCard Is Nothing Then
        WriteCard = gobjSquare.objSquareCard.zlBandCardArfter(Me, mlngModul, mCurSendCard.lng卡类别ID, lng病人ID, strExpend)
    Else
        WriteCard = False
    End If
    Exit Function
ErrHandl:
    WriteCard = False
    If ErrCenter() = 1 Then Resume
End Function


Private Sub SetOneCardBalance()
    Dim curOneCard As Currency, strName As String
    
    If mblnOneCard And Not mobjICCard Is Nothing Then
        curOneCard = mobjICCard.GetSpare(strName)
        If curOneCard <> 0 Then
           mrsOneCard.Filter = "名称='" & strName & "'"
           If mrsOneCard.RecordCount > 0 Then
                strName = mrsOneCard!结算方式
                If NeedName(cbo结算方式) <> strName Then zlControl.CboLocate cbo结算方式, strName
           End If
        End If
    End If
End Sub

Private Function RefreshFact() As Boolean
    '刷新发票号
    '说明：
    '   24363:主要是解决自动生成的号是否被用户更改：
    '    1.更改的票据号需要检查是否重复，重复后直接返回不更改发票号
    '    2.并发操作，不更改的情况下，检查是否重复，如果重复，自动取下一个号码！
    Dim strFact As String
    
    If mblnStationPrice Then Exit Function
    'lblFact.tag主要是检查发票号是否手工输入的.手工输入的,发票号为空,否则是自动产生的发票号
    If (lblFact.Tag <> "" And txtFact.Text <> "") Or Trim(txtFact.Text) = "" Then
        If gblnBill挂号 Then
            mlng领用ID = CheckUsedBill(IIf(gblnSharedInvoice, 1, 4), IIf(mlng领用ID > 0, mlng领用ID, glng挂号ID), , IIf(mblnStartFactUseType, mstrUseType, ""))
            If mlng领用ID <= 0 Then
                Select Case mlng领用ID
                    Case 0 '操作失败
                    Case -1
                        MsgBox "你没有自用和共用的挂号票据,请先领用一批票据或设置本地共用票据！", vbInformation, gstrSysName
                    Case -2
                        MsgBox "本地的共用票据已经用完,请先领用一批票据或重新设置本地共用票据！", vbInformation, gstrSysName
                End Select
                txtFact.Text = "": txtFact.Tag = "":  Exit Function
            End If
            
            '严格：取下一个号码
            txtFact.Text = GetNextBill(mlng领用ID)
        Else
            '松散：取下一个号码
            If gblnSharedInvoice Then
                strFact = zlDatabase.GetPara("当前收费票据号", glngSys, 1121)
            Else
                strFact = zlDatabase.GetPara("当前挂号票据号", glngSys, mlngModul)
            End If
            txtFact.Text = zlStr.Increase(strFact)
        End If
        txtFact.Tag = txtFact.Text: lblFact.Tag = txtFact.Tag
    End If
    RefreshFact = True
End Function

Private Function GetBookingNO(ByVal strInput As String) As String
    Dim rsTmp As ADODB.Recordset, strSQL As String
 
    If Len(strInput) = 8 And InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ", UCase(Mid(strInput, 1, 1))) > 0 And IsNumeric(Mid(strInput, 2)) Then
        strInput = UCase(strInput)
        strSQL = " And A.NO = [1]"
    Else
        strSQL = " And  (B.就诊卡号 = [1] Or B.Ic卡号 = [1] Or B.身份证号 = [1]" & IIf(IsNumeric(strInput), " Or B.门诊号 = [1]", "") & ")"
    End If
    
    strSQL = "" & _
    "Select Min(A.NO) NO" & vbNewLine & _
    "From 门诊费用记录 A, 病人信息 B" & vbNewLine & _
    "Where A.记录性质 = 4 And A.记录状态 = 0 And A.病人id = B.病人id(+)  " & _
                IIf(mTy_Para.int预约失效次数 > 0, "  And A.发生时间 between trunc(sysdate) and  trunc(sysdate)+1-1/24/60/60 ", _
            "  And ((nvl(A.加班标志,0) =0 And A.发生时间 > Trunc(Sysdate) - [2]) or  (nvl(A.加班标志,0) =1 And A.发生时间 > Trunc(Sysdate) - [3])  ) ") & strSQL
    On Error GoTo errH
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strInput, gSysPara.Sy_Reg.bytNODaysGeneral, gSysPara.Sy_Reg.bytNoDayseMergency)

    GetBookingNO = "" & rsTmp!NO
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub SetReceiveState(Optional blnReceive As Boolean)
    '------------------------------------------------------------------------------------------------------------------------
    '功能：设置接收预约时的状态,以及状态恢复
    '编制：刘兴洪
    '日期：2010-07-14 10:27:10
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    If dkpMain.Panes(1).Hidden = False Then '存在号别列表,屏蔽选择号别
        picPlan.Enabled = Not blnReceive
        mcbrToolBar.Controls.Find(xtpControlButton, conMenu_View_Refresh).Enabled = Not blnReceive   '刷新
        mcbrToolBar.Controls.Find(xtpControlButton, 2605).Enabled = Not blnReceive   '预留号码
        mcbrToolBar.Controls.Find(xtpControlButton, 2604).Enabled = Not blnReceive   '预留号码
    End If
    
    cboNO.Locked = blnReceive       '单据号
        
    chkPrint.Visible = Not blnReceive   '重打
    chkCancel.Visible = Not blnReceive    '退号
    chkBooking.Visible = Not blnReceive And InStr(1, mstrPrivs, ";预约挂号;") > 0 '预约
    If mobjCommunity Is Nothing Then
        cmdComminuty.Visible = False
    Else
        cmdComminuty.Visible = Not blnReceive  '社区病人
    End If
    cmdLookup.Visible = Not blnReceive          '查找病人
    cmdMore.Visible = True            '输入更多的病人信息
    lbl医疗类别.Visible = True
    cbo医疗类别.Visible = True
    
    cmdCard.Visible = InStr(1, mstrPrivs, ";绑定卡号;") > 0   '绑定卡号:31182:Not blnReceive And
    
    If mbytMode = 0 And mbytInState = 0 Then
        cmdYb.Visible = True
    Else
        cmdYb.Visible = blnReceive   '预约接收时,可以刷医保 '问题:31182
    End If
    
    lblIDCard.Visible = True
    If IDKind证件.IDKind = IDKind证件.GetKindIndex("身份证号") Then
        txtIDCard.Visible = True: txt证件.Visible = False
    Else
        txtIDCard.Visible = False: txt证件.Visible = True
    End If
    stbThis.Visible = True
    
    txt号别.Enabled = Not blnReceive '接收时不可再更改号别,但允许改序号
    cbo结算方式.Enabled = blnReceive Or gbln结算方式
    
    '55985:刘尔旋,2014-02-17,预约接受时允许修改费别和购买病历
    If InStr(1, mstrPrivs, ";允许修改费别;") > 0 And mTy_Para.bln预约接收确定挂号费 = True Then
        cbo费别.Enabled = True
        chk病历费.Enabled = True
    Else
        cbo费别.Enabled = Not blnReceive '可以选择结算方式
        chk病历费.Enabled = Not blnReceive '接收时不可再加收病历费
    End If
    
    txtSN.Locked = blnReceive
    
    If blnReceive Then
         '确定序号控制
         If GetCol("序号控制") >= 0 Then
            txtSN.Enabled = vsfPlan.TextMatrix(vsfPlan.Row, GetCol("序号控制")) <> ""
        End If
        If Not txtSN.Enabled And txtSN.Text <> "" Then txtSN.Text = ""
    End If
    Call zlPatiMoveCmdCtrl
    
End Sub

Private Function ReadBooking(ByVal strNO As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '功能：读取预约挂号单数据
    '入参：strNO-预约挂号单据号
    '返回:读取成功,返回True,否则返回False
    '编制：刘兴洪
    '日期：2010-07-16 16:21:45
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    On Error GoTo errH
    '非预约的,不处理
    If Not (chkCancel.Value = 0 And chkPrint.Value = 0 And chkBooking.Value = 0 And Not mblnStation) Then Exit Function
    mstrNoIn = strNO
    If mstrNoIn = "" Then
        MsgBox "没有找到待接收的预约挂号单！", vbInformation, gstrSysName
       ' mblnUnload = True
        cboNO.SetFocus: Exit Function
    End If
    
    mblnReadBooking = True
    If ReadBill(mstrNoIn, True) = False Then mblnReadBooking = False: Exit Function
    If mblnUnload Then mstrNoIn = "": Exit Function
    strSQL = "Select 出诊记录ID From 病人挂号记录 Where NO=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrNoIn)
    If Not rsTemp.EOF Then
        mlng记录ID = Val(Nvl(rsTemp!出诊记录ID))
    Else
        mlng记录ID = 0
    End If
    
    If Not txt发生时间.Text Like "____*" Then
        dtpAppointmentDate.Value = CDate(txt发生时间.Text) '此时没有自动调用change事件
    End If
    If txt门诊号.Text = "" And gbln自动门诊号 Then
        txt门诊号.Text = zlGet门诊号
    End If
    
    chkShowAll.Value = 1
    Call ShowPlans
    mblnReadBooking = False
    
    
    '定位号表,如果没有则不允许接收
    For i = 1 To vsfPlan.Rows - 1
        If Val(vsfPlan.TextMatrix(i, GetCol("记录ID"))) = mlng记录ID Then
            mblnChangeByCode = True
            vsfPlan.Row = i
            mblnChangeByCode = False
            vsfPlan_EnterCell
            Exit For
        End If
    Next
    If mbytMode = 2 Or mbytMode = 0 And mbytInState = 0 And mstrNoIn <> "" Then
        '预约接收
        If CheckIsPrice Then
            Call SetUndisplayBalance
        Else
            Call SetShowBalance
        End If
    End If
    
    If mbln建病案 And InStr(mstrPrivs, ";建立病案;") = 0 And txt门诊号.Text = "" Then
        MsgBox "该号别要求给病人建立门诊病案，但你没有建立病案的权限。不能接收。", vbInformation, gstrSysName
        mblnUnload = True: Exit Function
    End If
    cboNO.Text = mstrNoIn
    Call SetReceiveState(True)
    
    
    If gbytInvoice <> 0 Then Call RefreshFact
    If cbo结算方式.Enabled And cbo结算方式.Visible Then cbo结算方式.SetFocus
    If txt号别.Text <> "" Then
         Call ShowRegistFromInput
    End If
    '68216
    If Val(txtSN.Tag) <> 0 Then '
        txtSN.Text = txtSN.Tag
        locateSnBy时段 Val(txtSN.Tag), True
    End If
    ReadBooking = True
    Exit Function
errH:
    mblnReadBooking = False
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Sub ShowBookSeled()
    '------------------------------------------------------------------------------------------------------------------------
    '功能：根据快键,进入预约挂号接收小窗体,以选择具体的预约挂号单
    '编制：刘兴洪
    '日期：2010-07-16 16:34:39
    '说明：31182
    '------------------------------------------------------------------------------------------------------------------------
    Dim rsInfor As ADODB.Recordset
    Dim strOutNo As String
    Dim frmNew As frmSelRegist
    Dim blnExit As Boolean
    If mbytInState = 1 Then Exit Sub
    If InStr(1, mstrPrivs, ";接收预约;") = 0 Then Exit Sub
    If Not (chkCancel.Value = 0 And chkPrint.Value = 0 And chkBooking.Value = 0 And Not mblnStation) Then Exit Sub
    If mbytMode = 1 Or mbytMode = 2 Then Exit Sub
    Call CloseIDCard    '47007
    Set frmNew = New frmSelRegist
    If frmNew.ShowRegist(Me, mstrPrivs, mblnOlnyBJYB, mTy_Para.int预约失效次数, strOutNo, rsInfor) = False Then
        blnExit = True
    End If
    If Not frmNew Is Nothing Then Unload frmNew
    Set frmNew = Nothing
    Call NewCardObject
    If blnExit Then Exit Sub
    Call ReadBooking(strOutNo)
End Sub
Private Sub CloseIDCard()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:关闭自助读卡功能
    '编制:刘兴洪
    '日期:2012-03-09 16:26:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If Not mobjIDCard Is Nothing Then
        mobjIDCard.SetEnabled (False)
        Set mobjIDCard = Nothing
    End If
    If Not mobjICCard Is Nothing Then
        mobjICCard.SetEnabled (False)
        Set mobjICCard = Nothing
    End If
End Sub
Private Sub NewCardObject()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化新的卡对象
    '编制:刘兴洪
    '日期:2012-03-09 16:28:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mobjIDCard Is Nothing Then
        Set mobjIDCard = New clsIDCard
        Call mobjIDCard.SetParent(Me.Hwnd)
    End If
    If mobjICCard Is Nothing Then
        Set mobjICCard = New clsICCard
        Call mobjICCard.SetParent(Me.Hwnd)
    End If
End Sub

Private Sub cboNO_KeyPress(KeyAscii As Integer)
    Dim strOper As String, vDate As Date
    Dim str划价NO As String, strNO As String
    Dim blnEnableDel As Boolean, i As Long
    If KeyAscii = Asc("/") And Trim(cboNO.Text) = "" Then
        '预约接收时,如果单据号输入的是"/",则自动弹出小窗口,供预约挂号用"
        KeyAscii = 0:
        Call ShowBookSeled
        Exit Sub
    End If
    
      If KeyAscii = 13 And Trim(cboNO.Text) <> "" Then
        KeyAscii = 0
        cboNO.Text = Trim(cboNO.Text)
        
        If chkCancel.Value = 0 And chkPrint.Value = 0 And chkBooking.Value = 0 And Not mblnStation Then
            'A.接收预约挂号单
            'cboNO.Text = GetFullNO(cboNO.Text, 12) '不能自动补全单据号,因为输入的可能是门诊号,身份证等
            mblnRegReceiveByNo = True '问题号:57423
            strNO = cboNO.Text
            Call ClearBill
            '问题:38503
            If InStr(1, mstrPrivs, ";接收预约;") = 0 Then Exit Sub
            mstrNoIn = GetBookingNO(strNO)
            Call ReadBooking(mstrNoIn)        '必须要mstrNoIn值
        ElseIf chkCancel.Value = 1 Or chkPrint.Value = 1 Then
            'B.退号或重打
            cboNO.Text = GetFullNO(cboNO.Text, 12)
            strNO = cboNO.Text
            '是否已转入后备数据表中,注意此处不能加frmRegistFilter.mblnNOMoved条件判断,因为收费窗口和医生工作站窗口会调用这个窗体.
            If zlDatabase.NOMoved("门诊费用记录", strNO, , "4") Then
                If Not ReturnMovedExes(strNO, 4, Me.Caption) Then Exit Sub
                mblnNOMoved = False
            End If
            If InStr(1, mstrPrivs, ";强制退号;") = 0 Then
                    '单据操作权限检查,时间限制,不用检查挂号单有效天数
                    If Not ReadBillInfo(1, strNO, 4, strOper, vDate) Then
                        cboNO.Text = "": cboNO.SetFocus: Exit Sub
                    End If
                    If Not BillOperCheck(1, strOper, vDate, IIf(chkCancel.Value = 1, "退号", "重打")) Then
                        cboNO.Text = "": cboNO.SetFocus: Exit Sub
                    End If
            End If
            
            '单据退号权限
            If chkCancel.Value = 1 Then
                If mblnStation Then '门诊医生站退号检查
                    If Not StationDelete(strNO, str划价NO) Then
                        cboNO.Text = "": cboNO.SetFocus: Exit Sub
                    End If
                ElseIf InStr(1, mstrPrivs, ";强制退号;") = 0 Then
                    If CheckPriceHaveFee(strNO, str划价NO) Then Exit Sub
                    '检查挂号单是否已执行
                    blnEnableDel = (InStr(mstrPrivs, ";下医嘱后退号;") > 0)
                    If CheckExecuted(strNO, blnEnableDel) Then
                        MsgBox "挂号单" & strNO & "已经被医生接诊或下过医嘱,不能退号！", vbInformation, gstrSysName
                        cboNO.Text = "": cboNO.SetFocus: Exit Sub
                    End If
                    
                    '是否发生过费用,但未退费
                    If InStr(1, mstrPrivs, ";收费后退号;") = 0 Then
                        If ExistFee(strNO) Then
                            MsgBox strNO & "挂号单的病人已经产生了费用,须先退费才能退号.", vbInformation, gstrSysName
                            cboNO.Text = "": cboNO.SetFocus: Exit Sub
                        End If
                    End If
                End If
                mintInsure = ExistInsure(strNO)
                mlng结帐ID = GetBill结帐ID(strNO, 4)
            End If
            
            If Not ReadBill(strNO) Then
                MsgBox "没有发现你输入的挂号单据！", vbInformation, gstrSysName
                cboNO.Text = "": cboNO.SetFocus
            Else
                mstr划价NO = str划价NO
                If txtPatientPrint.Text <> "" And txtPatientPrint.Locked = False And txtPatientPrint.Visible Then
                    txtPatientPrint.SetFocus
                Else
                    If cmdOK.Visible And cmdOK.Enabled Then cmdOK.SetFocus
                End If
            End If
        End If
    Else
        If chkCancel.Value = 1 Or chkPrint.Value = 1 Then
            Call SetNOInputLimit(cboNO, KeyAscii)
        Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
            If InStr("0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Beep: Exit Sub
        End If
    End If
End Sub

Private Sub cbo户口地址_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Function ReadBill(strNO As String, Optional blnGetBooking As Boolean = False) As Boolean
    '功能：根据单据号读取挂号单据并显示在界面上
    '调用: 查看,退号,接收预约
    'blnGetBooking-是否是预约接收 因为在门诊挂号使用“/” 提取预约单据时 缺少对限制时间的检查 所以增加可选参数 在通过"/"提取的预约单据时 传入
       ' Dim rsBill As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Integer
    Dim curMoney As Currency
    Dim Datsys      As Date
    Dim datTmp      As Date
    Dim blnChk      As Boolean
    Dim bytState    As Byte, strTable As String
    Dim blnNotClick As Boolean
    Dim bln消费卡   As Boolean
    Dim dblTotal    As Double, dblBalance As Double
    Dim cllBillBalance As Collection
    Dim objCard As Card, rsTx As ADODB.Recordset
    Dim strWhere As String, str结帐IDs As String
    
    On Error GoTo errH
    
    Set mrsBill = Nothing
    If mbytInState <= 1 Then
        If mbytMode = 4 Then
            bytState = 1
        Else
            bytState = IIf(mbytMode <> 0 Or mbytMode = 0 And mbytInState = 0 And mstrNoIn <> "", 0, IIf(mblnViewCancel, 2, 1))
        End If
        
        If mblnViewOriginal Then bytState = 3
        
        If mintCancel = 1 Then
            strTable = ",Table(f_str2list([5])) M "
        ElseIf mintCancel = 2 Then
            strTable = ",Table(f_str2list([4])) M "
        Else
            strTable = ""
        End If
        strWhere = IIf(mblnStation, " And A.执行人=[3]", "") '医生站，需限制执行人
        strWhere = strWhere & IIf(mintCancel = 1 Or mintCancel = 2, "And A.收费细目ID = M.Column_Value", "") '读取指定项目
   
        If mbytMode = 0 And chkCancel.Value = 0 Then
            strSQL = " " & _
            "   Select  a.No, a.实际票号, Nvl(a.价格父号, a.序号) As 序号, a.从属父号, a.标识号,D.病人类型, a.病人id, a.付款方式, d.医疗付款方式, f.医疗类别,a.姓名, a.性别, a.年龄, " & _
            "            d.身份证号, d.家庭电话, d.家庭地址, d.出生日期, d.户口地址, a.费别,  a.加班标志, Nvl(a.附加标志, 0) As 附加标志, a.计算单位 As 号别, b.名称 As 项目,a.执行部门id, " & _
            "           c.名称 As 科室, nvl(a.应收金额,0)+nvl(J.应收,0) As 应收,nvl(a.实收金额,0)+nvl(J.实收,0) As 实收, G.退号审核人, g.退号审核时间, a.执行人,a.执行人, a.发生时间, a.操作员姓名, a.结帐id, a.摘要, a.结论, " & _
            "           Decode(g.号序, Null, a.发药窗口, To_Char(g.号序)) as 号序,a.收费细目id,a.收入项目id, a.价格父号, a.收费类别, a.数次, a.标准单价, a.收据费目, a.保险大类id, " & _
            "           a.保险项目否, a.统筹金额, a.保险编码, a.病人科室id,Nvl(a.记帐费用, 0)  As 记帐费用,Nvl(G.险类, 0)  As 险类 " & _
            "   From 病人挂号记录 G,门诊费用记录  A, 收费项目目录 B, 部门表 C, 病人信息 D, 就诊登记记录 F, " & _
            "          (  Select B1.NO,A1.序号,sum(A1.应收金额) as 应收,sum(A1.实收金额) as 实收 " & _
            "             From 门诊费用记录 A1,病人挂号记录 B1 " & _
            "             Where b1.收费单=A1.No and a1.记录性质=1 and A1.记录状态 in (0,1,3) and b1.NO=[2] And b1.记录状态(+)=Decode([1],0,1,[1]) " & _
            "             group by B1.NO,A1.序号  ) J " & strTable & _
            "   Where  G.No=A.No and a.记录性质=4 And a.记录状态 = [1] and G.NO=[2] And g.记录状态(+)=Decode([1],0,1,[1]) " & strWhere & _
            "          And a.收费细目id = b.Id And a.执行部门id = c.Id And a.病人id = d.病人id(+)  " & _
            "          And G.登记时间 = f.就诊时间(+) And G.病人id = F.病人id(+)  And (C.站点='" & gstrNodeNo & "' Or B.站点 is Null)" & _
            "          And A.NO=J.No(+) and A.序号=J.序号(+)"
        Else
              strSQL = "" & _
            " Select A.NO,A.实际票号,Nvl(A.价格父号,A.序号) as 序号,A.从属父号,A.标识号,D.病人类型," & _
            "           A.病人ID,A.付款方式 ,D.医疗付款方式,F.医疗类别,A.姓名,A.性别,A.年龄,D.身份证号,D.家庭电话 ,D.家庭地址, D.出生日期,D.户口地址,A.费别,A.加班标志," & _
            "           Nvl(A.附加标志,0) as 附加标志,A.计算单位 as 号别,B.名称 as 项目,A.执行部门ID,C.名称 as 科室," & _
            "           " & IIf(bytState = 2, "-1*", "") & "Sum(应收金额) as 应收," & IIf(bytState = 2, "-1*", "") & "Sum(实收金额) as 实收,e.退号审核人,e.退号审核时间," & _
            "           A.执行人,A.发生时间,A.操作员姓名,A.结帐ID,A.摘要,A.结论,Decode(E.号序, Null, A.发药窗口, To_Char(E.号序)) 号序,A.收费细目ID,A.收入项目ID,  A.价格父号, A.收费类别," & _
            "           A.数次, A.标准单价, A.收据费目, A.保险大类id, A.保险项目否, A.统筹金额, A.保险编码, A.病人科室id, " & _
            "           max(nvl(A.记帐费用,0)) as 记帐费用,Max(nvl(E.险类,0)) as  险类" & _
            " From 门诊费用记录 A,病人挂号记录 E,就诊登记记录 F,收费项目目录 B,部门表 C,病人信息 D" & strTable & _
            " Where A.NO=E.NO(+) And A.病人ID=D.病人ID(+) And A.记录性质=4 And A.记录状态=[1] And E.记录状态(+)=Decode([1],0,1,[1])  " & _
            "       And E.登记时间=F.就诊时间(+) And E.病人ID=F.病人ID(+)" & strWhere & _
            "       And A.NO=[2] And A.收费细目ID=B.ID And A.执行部门ID=C.ID" & _
            "       And (C.站点='" & gstrNodeNo & "' Or B.站点 is Null)" & IIf(mbytMode = 0 And chkCancel.Value = 0, " And e.收费单 Is Null ", "") & _
            " Group by A.NO,A.实际票号,Nvl(A.价格父号,A.序号),A.从属父号,A.标识号,D.病人类型,A.病人ID,A.付款方式,D.医疗付款方式,F.医疗类别,A.姓名,A.性别,D.身份证号,D.家庭电话," & _
            "           A.年龄,D.家庭地址,D.户口地址,A.费别,A.加班标志,A.附加标志,A.计算单位,B.名称,C.名称,A.执行部门ID,A.执行人,A.发生时间,A.操作员姓名,A.结帐ID,A.摘要,A.结论,Decode(E.号序, Null, A.发药窗口, To_Char(E.号序)),E.退号审核人, E.退号审核时间,A.收费细目ID,A.收入项目ID, A.价格父号, A.收费类别," & _
            "           A.数次, A.标准单价, A.收据费目, A.保险大类id, A.保险项目否, A.统筹金额, A.保险编码, A.病人科室id, D.出生日期" & _
            " "
        End If
        strSQL = strSQL & " Order By 序号 "
        If mblnNOMoved Then
            strSQL = Replace(strSQL, "门诊挂号记录", "H门诊挂号记录")
            strSQL = Replace(strSQL, "门诊费用记录", "H门诊费用记录")
        End If
        Set mrsBill = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, bytState, strNO, UserInfo.姓名, mstr附加项目ID, mstr退费项目IDs)
   Else
        strSQL = "" & _
        "   Select a.No, Null As 实际票号, 0 As 序号, Null As 从属父号, a.门诊号 as 标识号, a.病人id, Null As 付款方式, Null 医疗付款方式, f.医疗类别, a.姓名, a.性别, a.年龄," & _
        "          d.身份证号, d.家庭电话, d.家庭地址, d.费别, a.急诊 As 加班标志, Nvl(A.附加标志,0) as 附加标志, a.号别, b.名称 As 项目, a.执行部门id, c.名称 As 科室, 0  As 应收, 0 As 实收, a.执行人," & _
        "          a.发生时间, a.操作员姓名, Null As 结帐ID, a.摘要, a.预约方式 As 结论, a.号序,a.退号审核人,a.退号审核时间, 0 as 收费细目ID,0 as 收入项目ID,D.病人类型," & _
        "          0 as 记帐费用,Nvl(A.险类,0) as  险类,D.出生日期,D.户口地址" & _
        "   From 病人挂号记录 A, 收费项目目录 B,挂号安排 E, 部门表 C, 病人信息 D, 就诊登记记录 F  " & _
        "   Where E.项目id = b.Id And a.号别=e.号码 And a.执行部门id = c.Id And a.记录性质 = 2 And a.记录状态 = [1] And a.病人id = d.病人id(+) And " & _
        "       A.No=[2] and  a.登记时间 = f.就诊时间(+) And a.病人ID=f.病人ID(+)  " & _
        "       And (c.站点 ='" & gstrNodeNo & "' Or b.站点 Is Null)" & IIf(mblnStation, " And A.执行人=[3]", "") & vbNewLine & _
        "   Union All " & vbNewLine & _
        "   Select a.No, Null As 实际票号, 0 As 序号, Null As 从属父号, a.门诊号 as 标识号, a.病人id, Null As 付款方式, Null 医疗付款方式, f.医疗类别, a.姓名, a.性别, a.年龄," & _
        "          d.身份证号, d.家庭电话, d.家庭地址, d.费别, a.急诊 As 加班标志, Nvl(A.附加标志,0) as 附加标志, a.号别, b.名称 As 项目, a.执行部门id, c.名称 As 科室, 0  As 应收, 0 As 实收, a.执行人," & _
        "          a.发生时间, a.操作员姓名, Null As 结帐ID, a.摘要, a.预约方式 As 结论, a.号序,a.退号审核人,a.退号审核时间, 0 as 收费细目ID,0 as 收入项目ID,D.病人类型," & _
        "          0 as 记帐费用,Nvl(A.险类,0) as  险类,D.出生日期,D.户口地址" & _
        "   From 病人挂号记录 A, 收费项目目录 B,挂号安排 E, 部门表 C, 病人信息 D, 就诊登记记录 F ,收费从属项目 G " & _
        " Where E.项目id = G.主项Id And a.号别=e.号码 And a.执行部门id = c.Id And a.记录性质 = 2 And a.记录状态 = [1] And a.病人id = d.病人id(+) And " & _
        "        G.从项ID=b.Id And A.No=[2] and  a.登记时间 = f.就诊时间(+) And a.病人ID=f.病人ID(+)  " & _
        "        And (c.站点 ='" & gstrNodeNo & "' Or b.站点 Is Null)" & IIf(mblnStation, " And A.执行人=[3]", "")
        
        Set mrsBill = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(mbytInState), strNO, UserInfo.姓名)
    End If
    
    If mrsBill.EOF Then
        If mbytMode = 4 And mbytInState = 1 Then
            MsgBox "没有找到单据号为[" & mstrNoIn & "]的单据!", vbOKOnly, Me.Caption
        End If
        Exit Function
    End If
    mlng结帐ID = Val(Nvl(mrsBill!结帐ID))
    '------------------------------------
     ' 对接收 或者取消预约 的检查
     '------------------------------------
    Select Case mbytMode
    Case 2:
     '--接收
        If mbytMode = 2 And mTy_Para.lng预约有效时间 <> 0 Then
chkBooking:
            blnChk = True
            Datsys = DateAdd("n", 1 * mTy_Para.lng预约有效时间, zlDatabase.Currentdate)
            If Format(Datsys, "yyyy-MM-dd hh:mm:ss") > Format(mrsBill!发生时间, "yyyy-MM-dd hh:mm:ss") Then
                datTmp = DateAdd("n", -1 * mTy_Para.lng预约有效时间, CDate(Format(mrsBill!发生时间, "yyyy-MM-dd hh:mm:ss")))
                MsgBox "该预约号已过预约最后接收时间 " & Format(datTmp, "yyyy-MM-dd hh:mm:00") & ",不能接收", vbInformation, Me.Caption
                mblnUnload = True
                Exit Function
            End If
        End If
    Case 3:
         '--取消预约
         '----------------------
         '取消预约
         '限制参数:1. N天内不能取消预约号
         '        2.退号审核
         '   参数1.限制在取消预约必须在预约时间的N天外
         '   如果取消预约在N天内
         '    <1> 退号审核为真 时 审核的预约号 能够取消 否则不能
         '    <2> 退号审核为假 时 不能取消预约
         '----------------------
         If mTy_Para.lngN天取消预约 > 0 Then
            Datsys = zlDatabase.Currentdate
            datTmp = DateAdd("d", -1 * mTy_Para.lngN天取消预约, CDate(Format(mrsBill!发生时间, "yyyy-MM-dd hh:mm:ss")))
            '预约时间-K >datSys
            If Format(Datsys, "yyyy-MM-dd hh:mm:ss") > Format(datTmp, "yyyy-MM-dd hh:mm:ss") Then
                Select Case mTy_Para.bln退号审核
                Case False:
                ' 严格控制不能取消预约
                 MBox "该预约号已经超过最后取消预约时间" & Format(datTmp, "yyyy-MM-dd hh:mm:ss") & ",不能取消预约!"
                 mblnUnload = True
                 Exit Function
                Case True:
                  If Nvl(mrsBill!退号审核人, "") = "" Then
                    MBox "该单据号为" & Nvl(mrsBill!NO) & "的预约单没有经过退号审核!不能取消预约!"
                    mblnUnload = True
                    Exit Function
                  End If
                End Select
            End If
         End If
    Case Else:
    End Select
    
    If mbytMode = 0 And mbytInState = 0 And mstrNoIn <> "" Then
        '102230,调用外挂部件接口
        If PatiValiedCheckByPlugIn(mlngModul, Val(Nvl(mrsBill!病人ID)), _
            "<YSXM>" & NeedName(cbo医生.Text) & "</YSXM>") = False Then Exit Function
    End If
    
    If blnGetBooking And mbytMode <> 2 And mTy_Para.lng预约有效时间 <> 0 And blnChk = False Then GoTo chkBooking
    Call RemoveShowItem
    Call ClearMoney
    cboNO.Text = mrsBill!NO
    cboNO.Tag = mrsBill!NO
    txtFact.Text = Nvl(mrsBill!实际票号)
    cbo备注.Text = Nvl(mrsBill!摘要)
    
    mbln包含病历费 = False
    mbln附加费 = False
    mbln主费用 = False
    If mrsBill.RecordCount = 1 And Nvl(mrsBill!附加标志, 0) = 1 Then
        '单独收取病历费
        mblnUnChange = True
        txt号别.Text = "+"
        txtSN.Text = ""
        mblnUnChange = False
        chk病历费.Enabled = False
        mbln包含病历费 = True
        If mintCancel = 0 And chkCancel.Value = 1 Then
            chk病历费.Value = 1
        End If
    Else
        '正常挂号,包括购买病历
        vsfMoney.Tag = ""
        mrsBill.MoveFirst
        For i = 1 To mrsBill.RecordCount
            If Nvl(mrsBill!从属父号, 0) = 0 And Nvl(mrsBill!附加标志, 0) = 0 Then
                '只可能有一行
                mblnUnChange = True
                txt号别.Text = Nvl(mrsBill!号别)
                If Not IsNull(mrsBill!号序) Then txtSN.Text = IIf(IsNumeric(mrsBill!号序), mrsBill!号序, "")
                txtSN.Tag = txtSN.Text
                mblnUnChange = False
                If InStr("," & mstr附加项目ID & ",", "," & Nvl(mrsBill!收费细目ID) & ",") > 0 Then
                    mbln附加费 = True
                Else
                    mbln主费用 = True
                End If
                
                txt科室.Text = Nvl(mrsBill!科室)
                If (mbytMode = 2 Or mbytMode = 0 And mbytInState = 0 And mstrNoIn <> "") And mlng挂号科室ID = 0 Then mlng挂号科室ID = Nvl(mrsBill!执行部门id)
                cbo医生.Clear
                If Not IsNull(mrsBill!执行人) Then
                    cbo医生.AddItem mrsBill!执行人
                    cbo医生.ListIndex = 0
                End If
           
                lbl急.Visible = Nvl(mrsBill!加班标志, 0) = 1
            ElseIf Nvl(mrsBill!附加标志, 0) = 1 Then
                blnNotClick = mblnNotClick
                mblnNotClick = True
                '只可能有一行
                chk病历费.Value = 1
                mbln包含病历费 = True
                mblnNotClick = blnNotClick
                
            ElseIf Nvl(mrsBill!附加标志, 0) = 2 Then
                '标志包含就诊卡费
                vsfMoney.Tag = "卡费"
            End If
            mrsBill.MoveNext
         Next
        mrsBill.MoveFirst
    End If
    Call AdjustInfoPosition
    If chkPrint.Value <> 1 Then
        If mbln包含病历费 = True Then
            chk病历费.Enabled = mintCancel = 0
        End If
        If mbln附加费 = True Then
            mblnNotClick = True
            chkExtra.Value = 1
            mblnNotClick = False
            chkExtra.Enabled = mintCancel = 0
            chkExtra.Visible = mintCancel = 0
            chkExtra.Top = chk病历费.Top
            lbl预约方式.Visible = Not mbln附加费
            cbo预约方式.Visible = Not mbln附加费
        Else
            chkExtra.Visible = False
        End If
    End If
    If mbln包含病历费 Then chk病历费.Enabled = True
    
    mrsBill.MoveFirst
    Do While Not mrsBill.EOF
        dblTotal = dblTotal + Val(Nvl(mrsBill!实收))
        mrsBill.MoveNext
    Loop
    mrsBill.MoveFirst
    
    If (mbytMode = 2 Or mbytMode = 0 And mbytInState = 0 And mstrNoIn <> "") And Not IsNull(mrsBill!病人ID) Then
        mblnNotEMPIQuery = True
        Call GetPatient(IDKind.GetCurCard, "-" & mrsBill!病人ID, False)
    End If
    If mrsBill.RecordCount <> 0 And mrsBill.EOF Then mrsBill.MoveFirst
    txtPatient.Text = Nvl(mrsBill!姓名)
    '74428：李南春，2014-7-8，病人姓名显示颜色处理
    Call SetPatiColor(txtPatient, Nvl(mrsBill!病人类型), IIf(Val(mrsBill!险类) = 0, txtPatient.ForeColor, vbRed))
    If txtPatientPrint.Visible Then
        txtPatientPrint.Text = txtPatient.Text
        txtPatientPrint.Tag = Val(Nvl(mrsBill!病人ID))
        txtPatientPrint.ForeColor = txtPatient.ForeColor
        If Val(Nvl(mrsBill!病人ID)) <> 0 Then
            '如果是建档病人,则按以下规则更改姓名:
            '  1.只有挂号时建档的且才允许修改
            If Not CheckCanModifyName(cboNO.Text) And zlExistOperationData(Val(Nvl(mrsBill!病人ID)), cboNO.Text) Then
                txtPatientPrint.Locked = True
                Call SetRePrintPatiEnabled(False)
            Else
                txtPatientPrint.Locked = False
                Call SetRePrintPatiEnabled(True)
            End If
        End If
        '问题:53037
        ReInitPatiInvoice True
    End If
    
    If (mbytMode = 2 Or mbytMode = 0 And mbytInState = 0 And mstrNoIn <> "") Then mstrPrePati = txtPatient.Text
    
    
    Call LoadOldData("" & mrsBill!年龄, txt年龄, cbo年龄单位)
    mstr年龄 = txt年龄.Text
    mstr年龄单位 = IIf(IsNumeric(txt年龄.Text), cbo年龄单位.Text, "")
    cbo家庭地址.Text = Nvl(mrsBill!家庭地址)
    cbo户口地址.Text = Nvl(mrsBill!户口地址)
    '89242:李南春,2015/12/7,读取病人地址信息
    Call zlReadAddrInfo(padd家庭地址, Val(Nvl(mrsBill!病人ID)), 0, 3, cbo家庭地址.Text)
    Call zlReadAddrInfo(padd户口地址, Val(Nvl(mrsBill!病人ID)), 0, 4, cbo户口地址.Text)
    txtIDCard.Text = Nvl(mrsBill!身份证号): txt家庭电话.Text = Nvl(mrsBill!家庭电话)
    mblnNotChange = True
    cbo性别.ListIndex = cbo.FindIndex(cbo性别, Nvl(mrsBill!性别), True)
    If cbo性别.ListIndex = -1 Then
        cbo性别.AddItem Nvl(mrsBill!性别), 0
        cbo性别.ListIndex = cbo性别.NewIndex
    End If
    mblnNotChange = False
    mstr性别 = NeedName(cbo性别.Text)
    mstr姓名 = txtPatient.Text
    If mrsBill.RecordCount <> 0 And mrsBill.EOF Then mrsBill.MoveFirst
    txt门诊号.Text = Nvl(mrsBill!标识号)
    mRegistFeeMode = IIf(Val(Nvl(mrsBill!记帐费用)) = 1, EM_RG_记帐, EM_RG_现收)
    '103974:李南春,2016/12/16，查看、接收、退号等操作时，不换算年龄
    '也不换算出生日期
    mblnChange = False
    txt出生日期.Text = Format(IIf(IsNull(mrsBill!出生日期), "____-__-__", mrsBill!出生日期), "YYYY-MM-DD")
    If Not IsNull(mrsBill!出生日期) Then
        txt出生时间.Text = Format(mrsBill!出生日期, "HH:MM")
    Else
        txt出生时间.Text = "__:__"
    End If
    mblnChange = True
    
    '90875:李南春,2016/11/8,医疗卡证件类型
    If txtIDCard.Text = "" Then
        strSQL = "Select B.名称,A.卡号 from 病人医疗卡信息 A,医疗卡类别 B,证件类型 C " & _
                "Where A.卡类别ID=B.ID And B.名称=C.名称 And A.病人ID=[1]  Order by C.编码 "
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "缺省的证件卡号", Val(Nvl(mrsBill!病人ID)))
        If Not rsTmp.EOF Then
            IDKind证件.IDKind = IDKind证件.GetKindIndex(Nvl(rsTmp!名称))
            txt证件.Text = Nvl(rsTmp!卡号): txt证件.Tag = txt证件.Text
        End If
    End If
    
    '医疗付款方式
    If Not IsNull(mrsBill!医疗付款方式) Then
        cbo付款方式.ListIndex = cbo.FindIndex(cbo付款方式, mrsBill!医疗付款方式, True)
        If cbo付款方式.ListIndex = -1 Then
            cbo付款方式.AddItem mrsBill!医疗付款方式, 0
            cbo付款方式.ListIndex = cbo付款方式.NewIndex
        End If
    ElseIf Not IsNull(mrsBill!付款方式) Then
        cbo付款方式.AddItem Get医疗付款方式(Val(mrsBill!付款方式)), 0
        cbo付款方式.ListIndex = cbo付款方式.NewIndex
    Else
        cbo付款方式.ListIndex = -1
    End If
    
    cbo费别.ListIndex = cbo.FindIndex(cbo费别, Nvl(mrsBill!费别), True)
    If cbo费别.ListIndex = -1 Then
        cbo费别.AddItem Nvl(mrsBill!费别), 0
        cbo费别.ListIndex = cbo费别.NewIndex
    End If
    
    If mlngOutModeMC > 0 Then
        cbo医疗类别.ListIndex = cbo.FindIndex(cbo医疗类别, "" & mrsBill!医疗类别, True)
        If cbo医疗类别.ListIndex = -1 And Not IsNull(mrsBill!医疗类别) Then
            cbo医疗类别.AddItem "" & mrsBill!医疗类别, 0
            cbo医疗类别.ListIndex = cbo医疗类别.NewIndex
        Else
            cbo医疗类别.ListIndex = 0
        End If
    End If
    Set mobjDelCards = New Cards
    '134708:李南春,2018/12/14,清空一卡通对象
    Set mobjPayCard = Nothing
    Dim bln退号处理 As Boolean
    
    If mbytInState = 1 And mbytMode = 4 Or chkCancel.Value = 1 Then
        bln退号处理 = True
        '退号时,获取结算时相应的信息
         If Not zlReadRegThreeBalance(strNO, cllBillBalance, mobjPayCard) Then
         '卡类别ID,卡号,是否消费卡(1-是;0-否),交易流水号,交易说明,strNO,结帐ID
             SetDelBillCtlEnabled (False)
         Else
            If Not cllBillBalance Is Nothing Then
                bln消费卡 = Val(cllBillBalance(1)(2)) = 1
                Call SetDelBillCtlEnabled(True)
            End If
         End If
    End If
    '查阅病人挂号信息时,结算方式也调整为医疗卡名称
    If mbytInState = 1 And mbytMode = 0 Then
        Call zlReadRegThreeBalance(strNO, cllBillBalance, mobjPayCard)
    End If
    '68991
    If Val(Nvl(mrsBill!记帐费用)) <> 0 Then
        '是否医保刷卡
        mRegistFeeMode = EM_RG_记帐
        If mintInsure = 0 Then mintInsure = Val(Nvl(mrsBill!险类))
        Call SetUndisplayBalance
    Else
        mRegistFeeMode = EM_RG_现收
        If mintInsure = 0 Then mintInsure = ExistInsure(strNO)
    End If
    
    If mintInsure <> 0 Then Call initInsurePara(mrsBill!病人ID)
    
    If chkCancel.Value = 1 Or (mbytInState = 1 And mbytMode = 4) Then
        strSQL = "Select 结帐ID From 门诊费用记录 where NO = [1] and 记录性质 = 4 "
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
        Do While Not rsTmp.EOF
            If InStr("," & str结帐IDs & ",", "," & Val(Nvl(rsTmp!结帐ID)) & ",") = 0 Then
                str结帐IDs = str结帐IDs & "," & Val(Nvl(rsTmp!结帐ID))
            End If
            rsTmp.MoveNext
        Loop
        If str结帐IDs <> "" Then str结帐IDs = Mid(str结帐IDs, 2)
    Else
        str结帐IDs = mlng结帐ID
    End If
    
'    txt预交支付.Tag = ""
    '结算方式:可能包含医保支付部份
    strSQL = "Select Mod(A.记录性质,10) as 记录性质,B.性质,A.结算方式," & _
        IIf(bytState = 2, "-1*", "") & "Sum(A.冲预交) as 金额, A.结算号码 ,Nvl(Nvl(C.名称, D.名称), A.结算方式) As 名称 " & _
        " From " & IIf(mblnNOMoved, "H", "") & "病人预交记录 A,结算方式 B,医疗卡类别 C,消费卡类别目录 D" & _
        " Where A.结算方式=B.名称(+) And A.卡类别ID=C.ID(+) And a.结算卡序号 = D.编号(+) " & _
        "   And a.结帐id in (Select /* +cardinality(M,10) */ M.Column_Value From Table(f_Str2list([1])) M)" & _
        " Group by Mod(A.记录性质,10),B.性质,A.结算方式,A.结算号码,C.名称,D.名称" & _
        " Having Sum(A.冲预交) <> 0" & _
        " Order by Mod(A.记录性质,10),B.性质,A.结算方式"
    Set mrsBillAdvance = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str结帐IDs)
'    vsfPay.Clear 1
'    vsfPay.Rows = 2
'    vsfPay.RowHidden(1) = False
    Call Load结算信息(dblTotal, 0)
    
    If mrsBillAdvance.RecordCount <> 0 Then mrsBillAdvance.MoveFirst
    For i = 1 To mrsBillAdvance.RecordCount
        If mrsBillAdvance!记录性质 = 1 Or mrsBillAdvance!记录性质 = 11 Then
        Else
            Select Case Val(Nvl(mrsBillAdvance!性质))
            Case 3 '医保个人账户
                '74428：李南春，2014-7-8，病人姓名显示颜色处理
                Call SetPatiColor(txtPatient, Nvl(mrsBill!病人类型), vbRed)
            Case 7, 8    '一卡通相关
                If mobjPayCard Is Nothing Then
                    If bln退号处理 Then
                        Set objCard = New Card
                        With objCard
                            .接口序号 = 0
                            .名称 = Nvl(mrsBillAdvance!结算方式)
                            .结算方式 = Nvl(mrsBillAdvance!结算方式)
                            .接口编码 = Val(Nvl(mrsBillAdvance!性质))   ' 记录性质
                            .启用 = False
                        End With
                        mobjDelCards.Add objCard
                        cbo结算方式.ListIndex = -1
                    Else
                        cbo结算方式.ListIndex = cbo.FindIndex(cbo结算方式, mrsBillAdvance!结算方式, True)
                    End If
                    If cbo结算方式.ListIndex = -1 Then
                        cbo结算方式.AddItem mrsBillAdvance!结算方式, 0
                        cbo结算方式.ListIndex = cbo结算方式.NewIndex
                    End If
                    txt本次应缴.Text = Format(mrsBillAdvance!金额, "0.00")
                Else
                  cbo结算方式.Clear
                   If mobjPayCard.是否退现 Then
                        '支持退现，需要加入相关现金和非医保类的结算方式
                        Call Init结算方式("1,2", mobjDelCards)
                   End If
                   cbo结算方式.AddItem IIf(Nvl(mobjPayCard.名称) = "", mrsBillAdvance!结算方式, Nvl(mobjPayCard.名称))
                   mobjDelCards.Add mobjPayCard
                   Set mCurCardPay.objCard = mobjPayCard
                   mCurCardPay.lng医疗卡类别ID = mobjPayCard.接口序号
                   If (mobjPayCard.启用 Or cbo结算方式.ListIndex < 0 Or mobjPayCard.是否退现 = False) Then
                        cbo结算方式.ListIndex = cbo结算方式.NewIndex
                    End If
                End If
            Case Else '1,2或其他
                If mobjPayCard Is Nothing Then
                    If bln退号处理 Then
                        Set objCard = New Card
                        With objCard
                            .接口序号 = 0
                            .名称 = Nvl(mrsBillAdvance!结算方式)
                            .结算方式 = Nvl(mrsBillAdvance!结算方式)
                            .接口编码 = Val(Nvl(mrsBillAdvance!性质))   ' 记录性质
                            .启用 = False
                        End With
                        mobjDelCards.Add objCard
                        cbo结算方式.ListIndex = -1
                    Else
                        cbo结算方式.ListIndex = cbo.FindIndex(cbo结算方式, mrsBillAdvance!结算方式, True)
                    End If
                    If cbo结算方式.ListIndex = -1 Then
                        cbo结算方式.AddItem mrsBillAdvance!结算方式, 0
                        cbo结算方式.ListIndex = cbo结算方式.NewIndex
                    End If
                Else
                  cbo结算方式.Clear
                   If mobjPayCard.是否退现 Then
                        '支持退现，需要加入相关现金和非医保类的结算方式
                        Call Init结算方式("1,2", mobjDelCards)
                   End If
                   mobjDelCards.Add mobjPayCard
                    cbo结算方式.AddItem IIf(Nvl(mobjPayCard.结算方式) = "", mrsBillAdvance!结算方式, Nvl(mobjPayCard.结算方式))
                    If (mobjPayCard.启用 Or cbo结算方式.ListIndex < 0 Or mobjPayCard.是否退现 = False) Then
                        cbo结算方式.ListIndex = cbo结算方式.NewIndex
                    End If
                End If
                txt本次应缴.Text = Format(mrsBillAdvance!金额, "0.00")
            End Select
        End If
        mrsBillAdvance.MoveNext
    Next
    
    If bln退号处理 And Not mobjPayCard Is Nothing Then
        '退号:允许退现,允许更改结算方式
        If mobjPayCard.是否退现 Then cbo结算方式.Enabled = True
    End If
    
    
    txt发生时间.Text = Format(mrsBill!发生时间, "yyyy-MM-dd HH:mm:ss")
    cbo备注.Text = Nvl(mrsBill!摘要)
    lbl险类.Visible = False
    mblnNotChange = True
    zlControl.CboSetText cbo备注, Nvl(mrsBill!摘要)
    mblnNotChange = False
    mstr原摘要 = Nvl(mrsBill!摘要)
    '问题:26955
    zlAddComboItem cbo预约方式, Nvl(mrsBill!结论)
        
    mrsBill.MoveFirst
    vsfMoney.Rows = mrsBill.RecordCount + 1
    For i = 1 To mrsBill.RecordCount
        vsfMoney.TextMatrix(i, 0) = mrsBill!项目
        vsfMoney.TextMatrix(i, 1) = Format(mrsBill!应收, "0.00")
        vsfMoney.TextMatrix(i, 2) = Format(mrsBill!实收, "0.00")
        curMoney = curMoney + mrsBill!实收
        mrsBill.MoveNext
    Next
    mrsBill.MoveFirst
    txt合计.Text = Format(curMoney, "0.00")
    lbl合计.Caption = Format(curMoney, "0.00")
    Call Set连续挂号
    If txt门诊号.Text = "" And mbytMode = 2 And gbln自动门诊号 Then
        txt门诊号.Text = zlGet门诊号
    End If
    mbln建病案 = zlIsCreatePatiArchives(txt号别.Text)   '36131
    mblnNotEMPIQuery = False
    Call zlQueryEMPIPatiInfo
    ReadBill = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Load结算信息(ByVal dblTotal As Double, Optional ByVal dblDiff As Double) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据支付信息加载退费时的退费信息
    '入参:dblTotal-本次退费总金额，dblDiff-取消病历费或附加费的差额
    '返回:成功返回true,否则返回False
    '编制:李南春
    '日期:2018/5/2 11:35:08
    '问题:123874
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim dblBalance As Double
    Dim strSQL As String, rsTx As ADODB.Recordset
    On Error GoTo errH
    Call InitVsfPay(mbln连续挂号)
    If mrsBillAdvance Is Nothing Then Exit Function
    If mrsBillAdvance.RecordCount > 0 Then mrsBillAdvance.MoveFirst
    For i = 1 To mrsBillAdvance.RecordCount
        If dblTotal <> 0 Then
            If dblTotal < Val(mrsBillAdvance!金额) Then
                dblBalance = dblTotal
                dblTotal = 0
            Else
                If FormatEx(Val(mrsBillAdvance!金额), 6) >= FormatEx(dblDiff, 6) And dblDiff <> 0 And Val(Nvl(mrsBillAdvance!记录性质)) <> 5 Then
                    dblBalance = Val(mrsBillAdvance!金额) - dblDiff: dblDiff = 0
                Else
                    dblBalance = Val(mrsBillAdvance!金额)
                End If
                dblTotal = dblTotal - dblBalance
            End If
            If mrsBillAdvance!记录性质 = 1 Or mrsBillAdvance!记录性质 = 11 Then
                With vsfPay
                    .TextMatrix(.Rows - 1, 0) = "冲预交"
                    .TextMatrix(.Rows - 1, 1) = Format(dblBalance, "0.00")
                    .TextMatrix(.Rows - 1, .ColIndex("修改")) = "1"
                    .TextMatrix(.Rows - 1, 2) = Nvl(mrsBillAdvance!结算号码)
                    .RowData(.Rows - 1) = 0
                End With
            Else
                With vsfPay
                    '问题号:116146,焦博,2017/11/24,退号时,结算方式显示的是医疗卡的结算方式，统一调整为医疗卡名称
                    If mobjPayCard Is Nothing Then
                        .TextMatrix(.Rows - 1, 0) = mrsBillAdvance!结算方式
                    Else
                        .TextMatrix(.Rows - 1, 0) = IIf(Nvl(mobjPayCard.名称) <> "" And (Val(Nvl(mrsBillAdvance!性质, -1)) = 7 Or Val(Nvl(mrsBillAdvance!性质, -1)) = 8), Nvl(mobjPayCard.名称), mrsBillAdvance!结算方式)
                    End If
                    .TextMatrix(.Rows - 1, 1) = Format(dblBalance, "0.00")
                    .TextMatrix(.Rows - 1, 2) = Nvl(mrsBillAdvance!结算号码)
                    .RowData(.Rows - 1) = Val(Nvl(mrsBillAdvance!性质, -1))
                    If Val(Nvl(mrsBillAdvance!性质, -1)) = 7 Or Val(Nvl(mrsBillAdvance!性质, -1)) = 8 Then
                        strSQL = "Select ID,是否退现 From 医疗卡类别 Where 结算方式=[1]"
                        Set rsTx = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mrsBillAdvance!结算方式)
                        If rsTx.EOF Then
                            strSQL = "Select 编号,是否退现 From 消费卡类别目录 Where 结算方式=[1]"
                            Set rsTx = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mrsBillAdvance!结算方式)
                            If rsTx.EOF Then
                                vsfPay.TextMatrix(.Rows - 1, .ColIndex("修改")) = "1"
                            Else
                                vsfPay.TextMatrix(.Rows - 1, 4) = Nvl(rsTx!编号)
                                vsfPay.TextMatrix(.Rows - 1, .ColIndex("修改")) = IIf(Val(rsTx!是否退现) = 1, "0", "1")
                            End If
                        Else
                            vsfPay.TextMatrix(.Rows - 1, 4) = Nvl(rsTx!ID)
                            vsfPay.TextMatrix(.Rows - 1, .ColIndex("修改")) = IIf(Val(rsTx!是否退现) = 1, "0", "1")
                        End If
                    End If
                    If Val(Nvl(mrsBillAdvance!性质, -1)) = 1 Or Val(Nvl(mrsBillAdvance!性质, -1)) = 3 Then
                        .TextMatrix(.Rows - 1, .ColIndex("修改")) = "1"
                    End If
                End With
            End If
        End If
        mrsBillAdvance.MoveNext
        vsfPay.Rows = vsfPay.Rows + 1
    Next i
    
    Load结算信息 = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function zlIsCreatePatiArchives(ByVal str号码 As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查当前号别是否建档
    '入参:str号码-安排号码
    '返回:需建档,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-03-03 11:15:42
    '问题:36131
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    On Error GoTo errHandle
    strSQL = " Select max(病案必须) as 建档 From 挂号安排 where 号码=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str号码)
    zlIsCreatePatiArchives = Val(Nvl(rsTemp!建档)) = 1
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function CheckCanModifyName(ByVal strNO As String) As Boolean
'功能:检查挂号单是否可以修改姓名,如果不是挂号时建的档,就不能修改.
    Dim rsTmp As ADODB.Recordset, strSQL As String
 
    strSQL = "Select 1" & vbNewLine & _
            "From 门诊费用记录 A, 病人信息 B" & vbNewLine & _
            "Where A.NO = [1] And A.记录性质 = 4 And A.登记时间 = B.登记时间 And A.病人id = B.病人id"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
    CheckCanModifyName = rsTmp.RecordCount > 0
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Sub RemoveShowItem()
    '性别
    If cbo性别.ListCount > 0 Then
        If Not cbo性别.List(0) Like "*-*" Then
            cbo性别.RemoveItem 0
            SetCboDefault cbo性别
        End If
    End If
    '付款方式
    If cbo付款方式.ListCount > 0 Then
        If Not cbo付款方式.List(0) Like "*-*" Then
            cbo付款方式.RemoveItem 0
            SetCboDefault cbo付款方式
        End If
    End If
    '费别
    If cbo费别.ListCount > 0 Then
        If Not cbo费别.List(0) Like "*-*" Then
            cbo费别.RemoveItem 0
            SetCboDefault cbo费别
        End If
    End If
    
    '结算方式
    If cbo结算方式.ListCount > 0 Then
        If Not cbo结算方式.List(0) Like "*-*" Then
            cbo结算方式.RemoveItem 0
            SetCboDefault cbo结算方式
        End If
    End If
End Sub
Private Function GetCol(strName As String) As Long
   GetCol = vsfPlan.ColIndex(strName)
End Function

Private Sub SetPatiInfoEnabled(Optional ByVal blnUse As Boolean, Optional ByVal blnNewPati As Boolean, Optional ByVal blnReservePati As Boolean)
'功能：设置病人输入使能状态
    Dim blnEnabled As Boolean, lng病人ID As Long
    '82859:李南春,2015/4/8,病人基本信息调整
    If Not blnNewPati Then
        If mrsInfo.RecordCount > 0 Then lng病人ID = Val(Nvl(mrsInfo!病人ID))
    End If
    mbln基本信息调整 = Not (lng病人ID <> 0 And InStr(1, ";" & GetPrivFunc(glngSys, 9003) & ";", ";基本信息调整;") = 0)
    txtPatient.Enabled = gbln病人 Or blnUse
    If mblnStation Then
        blnEnabled = (gbln付款方式 Or blnUse) And blnNewPati
        cbo性别.Enabled = blnEnabled And mbln基本信息调整 '问题号:58843
        txt年龄.Enabled = blnEnabled And mbln基本信息调整 And Not mTy_Para.bln禁止输入年龄 '问题号:58843
        cbo年龄单位.Enabled = blnEnabled And mbln基本信息调整 And Not mTy_Para.bln禁止输入年龄 '问题号:58843
        cbo家庭地址.Enabled = gbln家庭地址 Or blnUse '问题号:58843
        cbo户口地址.Enabled = blnUse
        '89242:李南春,2015/12/7,读取病人地址信息
        padd家庭地址.Enabled = gbln家庭地址 Or blnUse: padd家庭地址.ControlLock = Not (gbln家庭地址 Or blnUse)
        padd户口地址.Enabled = blnUse: padd户口地址.ControlLock = Not blnUse
        cbo付款方式.Enabled = blnEnabled '问题号:58843
        txt家庭电话.Enabled = blnEnabled
    Else
        '刘兴洪:66032(更改王吉的问题58843)
        cbo性别.Enabled = mbln基本信息调整 And (gbln性别 Or blnUse)
        txt年龄.Enabled = mbln基本信息调整 And (gbln年龄 Or blnUse) And Not mTy_Para.bln禁止输入年龄
        cbo年龄单位.Enabled = mbln基本信息调整 And (gbln年龄 Or blnUse) And Not mTy_Para.bln禁止输入年龄
        txtIDCard.Enabled = mbln基本信息调整
        cbo家庭地址.Enabled = gbln家庭地址 Or blnUse
        cbo户口地址.Enabled = blnUse
        padd家庭地址.Enabled = gbln家庭地址 Or blnUse: padd家庭地址.ControlLock = Not (gbln家庭地址 Or blnUse)
        padd户口地址.Enabled = blnUse: padd户口地址.ControlLock = Not blnUse
        cbo付款方式.Enabled = gbln付款方式 Or blnUse
        If cbo付款方式.Enabled Then
            If mbytMode = 2 And gintPriceGradeStartType >= 2 Then
                cbo付款方式.Enabled = mTy_Para.bln预约接收确定挂号费
            End If
        End If
        txt出生时间.Enabled = mbln基本信息调整 And blnUse
        txt出生日期.Enabled = mbln基本信息调整 And blnUse
        txt家庭电话.Enabled = mbln基本信息调整 And (gbln电话 Or blnUse)
    End If
    
    cbo医疗类别.Enabled = blnUse
    cmdLookup.Enabled = txtPatient.Enabled And Not txtPatient.Locked
    cmdLookup.Enabled = cmdLookup.Enabled And Not (mblnStation And mTy_Para.bln挂号必须刷卡)
    If Not txtPatient.Enabled And Not blnReservePati Then
        mstrPrePati = ""
        txtPatient.Text = ""
        txt门诊号.Text = ""
    End If
    
    'If Not txt年龄.Enabled  Then txt年龄.Text = ""
    'If Not cbo家庭地址.Enabled Then cbo家庭地址.Text = ""
    
    If Not cbo性别.Enabled And gstr性别 <> "无" And txtPatient.Text <> mstrPrePati And mrsInfo Is Nothing Then
        Call SetCboDefault(cbo性别)
    ElseIf gstr性别 = "无" And txtPatient.Text <> mstrPrePati Then
        cbo性别.ListIndex = -1
    End If
    If cbo付款方式.ListIndex = -1 Then Call SetCboDefault(cbo付款方式)
End Sub

Private Sub Fill医生(ByVal lng科室ID As Long)
'功能：根据科室读取并绑定医生下拉列表
    Dim strSQL As String
        
    On Error GoTo errH
    If mrsDoctor.State = 1 Then
        mrsDoctor.Filter = "部门id=" & lng科室ID
        
        Do While Not mrsDoctor.EOF
            cbo医生.AddItem IIf(IsNull(mrsDoctor!简码), "", mrsDoctor!简码 & "-") & mrsDoctor!姓名
            cbo医生.ItemData(cbo医生.NewIndex) = mrsDoctor!ID
            mrsDoctor.MoveNext
        Loop
        If cbo医生.ListCount > 0 Then
            cbo医生.ListIndex = 0
            cbo医生.TabStop = gbln医生 And Not mblnStation
            
            mstr医生姓名 = Mid(cbo医生.Text, InStr(1, cbo医生.Text, "-") + 1)
            mlng医生ID = cbo医生.ItemData(cbo医生.ListIndex)
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub GetAll医生()
    Dim strSQL As String
    On Error GoTo errH
    
    strSQL = "Select a.Id, a.姓名, Upper(a.简码) As 简码,b.部门id,a.编号" & _
            " From 人员表 a, 部门人员 b, 人员性质说明 c" & _
            " Where a.Id = b.人员id And a.Id = c.人员id And c.人员性质 = [1] And (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null) " & _
            " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
            " Order By a.简码 Desc"
    Set mrsDoctor = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, "医生")
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function GetRoom(lng记录ID As Long) As String
'功能：根据号别的分诊方式获取号别的诊室
    Dim strSQL As String, strRoomIDs As String
    Dim rsTmp As ADODB.Recordset, rsRoom As ADODB.Recordset
    
    On Error GoTo errH
            
    strSQL = "Select ID,Nvl(分诊方式,0) as 分诊 From 临床出诊记录 Where ID = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng记录ID)
    
    If rsTmp.EOF Then Exit Function
    If rsTmp!分诊 = 0 Then Exit Function '不分诊
    
    '处理分诊
    If rsTmp!分诊 = 1 Then
        '指定诊室
        strSQL = "Select B.名称 As 门诊诊室 From 临床出诊诊室记录 A,门诊诊室 B Where A.诊室ID=B.ID And A.记录ID = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CLng(rsTmp!ID))
        If Not rsTmp.EOF Then GetRoom = rsTmp!门诊诊室
    ElseIf rsTmp!分诊 = 2 Then
        '动态分诊：该个号别当天挂号未诊数最少的诊室   //todo未考虑预约挂号
        strSQL = _
            " Select 门诊诊室,Sum(NUM) as NUM From (" & _
                " Select B.名称 As 门诊诊室,0 as NUM From 临床出诊诊室记录 A,门诊诊室 B Where A.诊室ID = B.ID And 记录ID=[1]" & _
                " Union ALL" & _
                " Select 诊室,Count(诊室) as NUM From 病人挂号记录" & _
                " Where Nvl(执行状态,0)=0 And 记录性质=1 and 记录状态=1 and  发生时间 Between Trunc(Sysdate) And Sysdate And 出诊记录ID = [2]" & _
                " And 诊室 IN (Select D.名称 As 门诊诊室 From 临床出诊诊室记录 C,门诊诊室 D Where C.记录ID=[1] And C.诊室ID = D.ID )" & _
                " Group by 诊室)" & _
            " Group by 门诊诊室 Order by Num"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CLng(rsTmp!ID), lng记录ID)
        If Not rsTmp.EOF Then GetRoom = rsTmp!门诊诊室
    ElseIf rsTmp!分诊 = 3 Then
        '平均分诊：当前分配=1表示下次应取的当前诊室
        strSQL = "Select * From 临床出诊诊室记录 Where 记录ID=" & rsTmp!ID
'        strSQL = "Select A.记录ID,B.名称 As 门诊诊室,A.当前分配 From 临床出诊诊室记录 A,门诊诊室 B Where A.诊室ID=B.ID And A.记录ID=" & rsTmp!ID
        Set rsTmp = New ADODB.Recordset
        Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption, adOpenDynamic, adLockOptimistic)
        If Not rsTmp.EOF Then
            Do While Not rsTmp.EOF
                If IIf(IsNull(rsTmp!当前分配), 0, rsTmp!当前分配) = 1 Then
                    strRoomIDs = rsTmp!诊室ID
                    rsTmp!当前分配 = 0
                    rsTmp.MoveNext
                    If rsTmp.EOF Then rsTmp.MoveFirst
                    rsTmp!当前分配 = 1
                    rsTmp.Update
                    Exit Do
                End If
                rsTmp.MoveNext
            Loop
            '处理第一次平均分配
            If strRoomIDs = "" Then
                rsTmp.MoveFirst
                strRoomIDs = rsTmp!诊室ID
                rsTmp.MoveNext
                If rsTmp.EOF Then rsTmp.MoveFirst
                rsTmp!当前分配 = 1
                rsTmp.Update
            End If
        End If
        If strRoomIDs <> "" Then
            strSQL = "Select 名称 From 门诊诊室 Where ID = [1]"
            Set rsRoom = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strRoomIDs)
            If Not rsRoom.EOF Then
                GetRoom = rsRoom!名称
            End If
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetActualCash(ByVal lng结帐ID As Long) As Currency
'功能：获取本次挂号医保结算后现金支付部份金额
'200510byZT
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    '刘兴洪:26242
    '   原因是没有加上就诊卡费(就诊卡费用是另外一个结帐ID,需要用收费用时间来处理

    strSQL = "" & _
    "   Select Sum(冲预交) As 金额 " & _
    "   From 病人预交记录 A, 结算方式 B " & _
    "   Where A.结算方式 = B.名称 And B.性质 = 1 And " & _
    "         (A.收款时间, A.病人id) In (Select 收款时间, 病人id From 病人预交记录 Where 记录性质 = 4 And 结帐id = [1])"
    
    
    'strSQL = "" & _
    "   Select A.冲预交 as 金额" & _
    "   From 病人预交记录 A,结算方式 B" & _
    "   Where A.结算方式=B.名称 And B.性质=1 And A.记录性质=4 And A.结帐ID=[1] " & _
    "   "
    
    '加上卡费处理
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng结帐ID)
    If Not rsTmp.EOF Then
        GetActualCash = Nvl(rsTmp!金额, 0)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Init费别(bln初诊 As Boolean, Optional blnKeepIndex As Boolean) As Boolean
'参数：bln初诊=是否允许仅限初诊的项目
'      blnKeepIndex=是否保持原有的费别选择
    Dim strSQL As String, i As Integer
    Dim strKeep As String
    Dim str缺省费别 As String
    
    On Error GoTo errH
    
    strKeep = cbo费别.Text      '病人以前的费别,有可能现在的系统中已没有该费别了
    If strKeep <> "" Then strKeep = Mid(strKeep, InStr(1, strKeep, "-") + 1)
    str缺省费别 = gstr费别      '本地缺省费别,如果为空,后面再取系统缺省
    
    '72168,冉俊明,2014/4/22,挂号时通过挂号科室确定可选费别
    If mrs费别 Is Nothing Then '首次调用该函数时[bln初诊]为true
        Set mrs费别 = New ADODB.Recordset
        '费别:身份唯一性项目(包含了缺省费别),可以是初诊,不管有效期间及科室
        strSQL = "Select a.编码, a.名称, a.简码, Nvl(a.仅限初诊, 0) As 初诊," & _
                "       Nvl(a.缺省标志, 0) As 缺省, Nvl(b.科室id, 0) As 科室id" & _
                " From 费别 A, 费别适用科室 B" & _
                " Where a.名称 = b.费别(+) And a.属性 = 1" & _
                "      And Trunc(Sysdate) Between Nvl(a.有效开始, To_Date('1900-01-01', 'YYYY-MM-DD'))" & _
                "                         And Nvl(a.有效结束, To_Date('3000-01-01', 'YYYY-MM-DD'))" & _
                "      And Nvl(a.服务对象, 3) In (1, 3)" & _
                " Order By a.编码"
        Call zlDatabase.OpenRecordset(mrs费别, strSQL, Me.Caption)
    End If
    
    If mrs费别 Is Nothing Then Exit Function
    If bln初诊 Then
        mrs费别.Filter = "科室id=" & mlng挂号科室ID & " or 科室id=0"   'adFilterNone
    Else                        '不允许仅限初诊的项目
        mrs费别.Filter = "(初诊=0 and 科室id=" & mlng挂号科室ID & ") or (初诊=0 and 科室id=0)"
    End If
    If mrs费别.RecordCount > 0 Then mrs费别.MoveFirst
    
    cbo费别.Clear: mstrPre费别 = ""
    Do While Not mrs费别.EOF
        cbo费别.AddItem mrs费别!编码 & "-" & mrs费别!名称
        '记录初诊项目:不会是本地缺省及系统缺省
        cbo费别.ItemData(cbo费别.NewIndex) = IIf(mrs费别!初诊 = 1, 2, 0)
        
        If str缺省费别 = "" Then    '没有本地缺省时取系统缺省
            If mrs费别!缺省 = 1 Then str缺省费别 = mrs费别!名称
        End If
        mrs费别.MoveNext
    Loop
    
    If blnKeepIndex And Not mrsInfo Is Nothing Then
        If Not mrsInfo.EOF Then Call zlControl.CboLocate(cbo费别, Nvl(mrsInfo!费别))
    End If
    If blnKeepIndex And strKeep <> "" Then Call zlControl.CboLocate(cbo费别, strKeep)

    If cbo费别.ListIndex = -1 Then Call zlControl.CboLocate(cbo费别, str缺省费别)
    
    If cbo费别.ListIndex = -1 Then If cbo费别.ListCount > 0 Then cbo费别.ListIndex = 0
    If cbo费别.ListIndex <> -1 Then cbo费别.ItemData(cbo费别.ListIndex) = 1
            
    Init费别 = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function PatiExist(strCard As String) As Boolean
'功能：判断是否确实存在该卡号的持卡病人,因为住院病人不能在此刷卡
    Dim rsTmp As ADODB.Recordset
    Dim strSQL  As String
    
    On Error GoTo errH
    
    strSQL = "Select a.就诊卡号 " & vbNewLine & _
             "From 病人信息 A, 病人医疗卡信息 B, 医疗卡类别 C " & vbNewLine & _
             "Where a.就诊卡号 = b.卡号 And c.特定项目 = '就诊卡' And b.卡类别id = c.Id And a.在院 = 1 And b.卡号 = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strCard)
    PatiExist = Not rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function SetIdentifyLocked(blnLocked As Boolean)
'功能：设置医保身份验证后不允许修改的信息条目
    txtPatient.Locked = blnLocked
    cbo性别.Locked = blnLocked
    cbo性别.TabStop = Not blnLocked
    txt年龄.Locked = blnLocked
    txt年龄.TabStop = Not blnLocked
    cbo年龄单位.Locked = blnLocked
    cbo年龄单位.TabStop = Not blnLocked
    cbo付款方式.Locked = blnLocked
    cbo付款方式.TabStop = Not blnLocked
    cmdLookup.Enabled = IIf(Not blnLocked, txtPatient.Enabled, Not blnLocked)
    cmdLookup.Enabled = cmdLookup.Enabled And Not (mblnStation And mTy_Para.bln挂号必须刷卡)
    
    If blnLocked Then
        txtPatient.BackColor = &HE0E0E0
    Else
        txtPatient.BackColor = &HFFFFFF
    End If
    txt年龄.BackColor = txtPatient.BackColor
    cbo性别.BackColor = txtPatient.BackColor
    cbo年龄单位.BackColor = txtPatient.BackColor
    cbo付款方式.BackColor = txtPatient.BackColor
    
    With mobjfrmPatiInfo
        .txtPatient.Locked = blnLocked
        .cbo性别.Locked = blnLocked
        .txt年龄.Locked = blnLocked
        .cbo年龄单位.Locked = blnLocked
        .cbo付款方式.Locked = blnLocked
    End With
    
End Function

Private Sub ClearMoney()
    Dim blnDraw As Boolean, i As Long
    Dim j As Long, blnFinish As Boolean
    With vsfMoney
        blnDraw = .Redraw
        .Redraw = False
        For i = 1 To .Rows - 1
            .RowData(i) = 0
            .TextMatrix(i, 0) = "": .ColAlignment(0) = 1
            .TextMatrix(i, 1) = "": .ColAlignment(1) = 7
            .TextMatrix(i, 2) = "": .ColAlignment(2) = 7
        Next
        .Rows = 2
        .Row = 1: .TopRow = 1
        .Col = 0: .ColSel = .Cols - 1
        .Redraw = blnDraw
    End With
    If mbln连续挂号 Then
        cbo结算方式.Enabled = False
        Call InitVsfPay(True)
        txt本次应缴.Text = Format(mcur应缴, "0.00")
    Else
        cbo结算方式.Enabled = gbln结算方式
        Call InitVsfPay
        If mblnCancel Then
            mcur合计 = 0
            mcur应缴 = 0
            txt合计.Text = "0.00"
            txt本次应缴.Text = "0.00"
        End If
    End If
End Sub

Private Sub InitVsfPay(Optional ByVal bln连续挂号 As Boolean)
    Dim i As Long, j As Long
    Dim blnFinish As Boolean, blnDraw As Boolean
    
    If bln连续挂号 Then
        With vsfPay
            For i = 1 To .Rows - 1
                If .RowData(i) = 0 And .TextMatrix(i, 0) = "预交金" Then
                    .TextMatrix(i, 0) = "预交金"
                    .TextMatrix(i, 1) = "0.00"
                    .TextMatrix(i, 2) = ""
                    .TextMatrix(i, 3) = 1
                    .TextMatrix(i, 4) = "预交金"
                    .TextMatrix(i, 5) = 1
                    .TextMatrix(i, 6) = 0
                    If mdbl预交余额 = 0 Then .RowHidden(i) = True
                End If
                If .RowData(i) = 3 And .TextMatrix(i, 0) <> "" And mstr个人帐户 <> "" Then
                    .TextMatrix(i, 0) = mstr个人帐户
                    .TextMatrix(i, 1) = "0.00"
                    .TextMatrix(i, 2) = ""
                    .TextMatrix(i, 3) = 1
                    .TextMatrix(i, 4) = mstr个人帐户
                    .TextMatrix(i, 5) = 1
                    .TextMatrix(i, 6) = 0
                    If mdbl个帐余额 = 0 Then .RowHidden(i) = True
                End If
                If (.RowData(i) = 1 Or .RowData(i) = 2) And .TextMatrix(i, 0) = NeedName(cbo结算方式.Text) Then
                    .TextMatrix(i, 7) = Val(.TextMatrix(i, 1))
                Else
                    If .RowData(i) <> 0 And .RowData(i) <> 3 And .TextMatrix(i, 0) <> "" Then
                        .TextMatrix(i, 0) = "删除"
                        blnFinish = False
                    End If
                End If
            Next i
            Do While blnFinish = False
                For i = 1 To .Rows - 1
                    If .TextMatrix(i, 0) = "删除" Then
                        .RemoveItem i
                        Exit For
                    End If
                    If i = .Rows - 1 Then blnFinish = True
                Next i
            Loop
        End With
    Else

        With vsfPay
            blnDraw = .Redraw
            .Redraw = False
            For i = 1 To .Rows - 1
                .RowData(i) = 0
                For j = 0 To .Cols - 1
                    .TextMatrix(i, j) = ""
                Next j
            Next
            .Rows = 2
            .Row = 1: .TopRow = 1
            .Col = 0: .ColSel = .Cols - 1
            '加载预交金
            .TextMatrix(.Rows - 1, 0) = "预交金"
            .TextMatrix(.Rows - 1, 1) = "0.00"
            .TextMatrix(.Rows - 1, 2) = ""
            .TextMatrix(.Rows - 1, 3) = 1
            .TextMatrix(.Rows - 1, 4) = "预交金"
            .TextMatrix(.Rows - 1, 5) = 1
            .TextMatrix(.Rows - 1, 6) = 0
            .RowData(.Rows - 1) = 0
            If mdbl预交余额 = 0 Then .RowHidden(.Rows - 1) = True
            .Rows = .Rows + 1
            '加载个人帐户
            If mstr个人帐户 <> "" Then
                .TextMatrix(.Rows - 1, 0) = mstr个人帐户
                .TextMatrix(.Rows - 1, 1) = "0.00"
                .TextMatrix(.Rows - 1, 2) = ""
                .TextMatrix(.Rows - 1, 3) = 1
                .TextMatrix(.Rows - 1, 4) = mstr个人帐户
                .TextMatrix(.Rows - 1, 5) = 1
                .TextMatrix(.Rows - 1, 6) = 0
                .RowData(.Rows - 1) = 3
                If mdbl个帐余额 = 0 Then .RowHidden(.Rows - 1) = True
                .Rows = .Rows + 1
            End If
            .Redraw = blnDraw
        End With
    End If
End Sub

Private Sub CalcYBMoney()
'功能：计算并显示当前医保病人个人帐户可以支持的金额
    Dim cur合计 As Currency
    Dim strInfo As String, i As Long, j As Long, lng病人ID As Long
    Dim curTotal As Currency
    Dim lngYBRow As Long, lngYJRow As Long
    
    If mRegistFeeMode = EM_RG_记帐 Then Exit Sub
    
    For i = 1 To vsfPay.Rows - 1
        If vsfPay.RowData(i) = 3 And vsfPay.TextMatrix(i, 0) <> "" Then lngYBRow = i
        If vsfPay.RowData(i) = 0 And vsfPay.TextMatrix(i, 0) <> "" Then lngYJRow = i
    Next i
    
    cur合计 = GetRegistMoney(True)
    curTotal = cur合计
    If MCPAR.不收病历费 = True Then
        cur合计 = cur合计 - mcur病历
    End If
    If mstrYBPati <> "" Then lng病人ID = Val(Split(mstrYBPati, ";")(8))
    
    '计算并显示个人帐户支付金额
    '要求医保支持个人帐户支付及ZLHIS允许使用个人帐户
    If mintInsure <> 0 And mstr个人帐户 <> "" Then
        If gclsInsure.GetCapability(support挂号使用个人帐户, lng病人ID, mintInsure) Then
            If mdbl个帐余额 - cur合计 >= -1 * mcur个帐透支 Then
                vsfPay.TextMatrix(lngYBRow, 1) = Format(cur合计, "0.00")
                vsfPay.TextMatrix(lngYBRow, 6) = cur合计
            Else
                If mblnStation Then
                    vsfPay.TextMatrix(lngYBRow, 1) = "0.00"
                ElseIf mcur个帐透支 = 0 And mdbl个帐余额 > 0 Then
                    vsfPay.TextMatrix(lngYBRow, 1) = Format(mdbl个帐余额, "0.00")
                    vsfPay.TextMatrix(lngYBRow, 6) = mdbl个帐余额
                Else
                    vsfPay.TextMatrix(lngYBRow, 1) = "0.00"
                End If
            End If
        Else
            vsfPay.TextMatrix(lngYBRow, 1) = "0.00"
        End If
    Else
        If lngYBRow <> 0 Then vsfPay.TextMatrix(lngYBRow, 1) = "0.00"
    End If
    
    If gblnPrePayPriority And mdbl预交余额 >= Val(curTotal - Val(vsfPay.TextMatrix(lngYBRow, 1))) Then
        vsfPay.TextMatrix(lngYJRow, 1) = Format(curTotal - Val(vsfPay.TextMatrix(lngYBRow, 1)), "0.00")
    End If
    If lngYJRow <> 0 Then vsfPay.TextMatrix(lngYJRow, 6) = mdbl预交余额
    
    '获取医保统筹相关内容
    If mintInsure <> 0 And mstrYBPati <> "" And Not mrsItems Is Nothing Then
        mrsItems.MoveFirst
        For i = 1 To mrsItems.RecordCount
            mrsInComes.Filter = "项目ID=" & mrsItems!项目ID
            For j = 1 To mrsInComes.RecordCount
                strInfo = gclsInsure.GetItemInsure(lng病人ID, mrsItems!项目ID, mrsInComes!实收, True, mintInsure)
                If strInfo <> "" Then
                    mrsItems!保险项目否 = Val(Split(strInfo, ";")(0))
                    mrsItems!保险大类id = Val(Split(strInfo, ";")(1))
                    mrsItems!保险编码 = CStr(Split(strInfo, ";")(3))
                    mrsInComes!统筹金额 = Format(Val(Split(strInfo, ";")(2)), "0.00")
                End If
                mrsInComes.MoveNext
            Next
            mrsItems.MoveNext
        Next
    End If
    Call Set连续挂号
End Sub

Private Sub ReCalc预约接收发卡()
    '------------------------------------------------------------------------------------------------------------------------
    '功能：重新计算预约接收发卡的卡费用信息
    '编制：刘兴洪
    '日期：2010-07-16 09:38:54
    '说明：31182
    '------------------------------------------------------------------------------------------------------------------------
    Dim blnExitLoop As Boolean, i As Long, j As Long, lngRow As Long, lng病人ID As Long
    Dim str费别 As String, cur应收 As Currency, cur实收  As Currency, cur合计 As Currency
    Dim cur病历 As Currency, rsItems As ADODB.Recordset, rsIncomes As ADODB.Recordset
    
     '31182:预约接收时,也要收取相应的卡费
    '删除卡费等
    Do While True
       blnExitLoop = True
       For j = 1 To vsfMoney.Rows - 1
             If vsfMoney.RowData(j) <> 0 Then
                vsfMoney.RemoveItem j:
                blnExitLoop = False
                Exit For
             End If
       Next
       If blnExitLoop Then Exit Do
    Loop
    Set mrsItems = Nothing
    Set mrsInComes = Nothing
    str费别 = NeedName(cbo费别.Text)
    If mrsBill Is Nothing Then Exit Sub
    
    mrsBill.MoveFirst
    Call ReadRegistPrice(mrsBill!收费细目ID, mbln包含病历费, mblnAddCardItem, str费别, rsItems, rsIncomes, 0, mintInsure, _
        txt号别.Text, 10, mlng挂号科室ID, mobjfrmPatiInfo.mstrPriceGrade, _
        IIf(dtpAppointmentDate.Visible Or mbytMode = 2, Format(dtpAppointmentDate.Value, "yyyy-mm-dd") & " 23:59:59", ""), gCurSendCard.lng收费细目ID)
    
    If mintInsure <> 0 Then
        If MCPAR.挂号检查项目 = True Then
            If gclsInsure.CheckItem(mintInsure, 2, 0, rsItems) = False Then
                MsgBox "医保病人收费项目检查失败，不能继续挂号！", vbInformation, gstrSysName
                Call ClearBill: Exit Sub
            End If
        End If
    End If
    If mrsInfo Is Nothing Then
        lng病人ID = 0
    Else
        If mrsInfo.RecordCount = 0 Then
            lng病人ID = 0
        Else
            lng病人ID = Val(Nvl(mrsInfo!病人ID))
        End If
    End If
    Call ReadRegistPrice(0, False, mblnAddCardItem, str费别, mrsItems, mrsInComes, lng病人ID, mintInsure, _
            txt号别.Text, mbytMode, , mobjfrmPatiInfo.mstrPriceGrade, _
    IIf(dtpAppointmentDate.Visible Or mbytMode = 2, Format(dtpAppointmentDate.Value, "yyyy-mm-dd") & " 23:59:59", ""), gCurSendCard.lng收费细目ID)
    
    '显示卡费数据
     If Not mrsItems Is Nothing Then
         vsfMoney.Redraw = False
         cur应收 = 0: cur实收 = 0
         For j = 1 To vsfMoney.Rows - 1
             If vsfMoney.RowData(j) = 0 Then    '回为读取单据的时候,没有加入RowData数据
                 cur实收 = Val(vsfMoney.TextMatrix(j, 2))
                cur合计 = cur合计 + cur实收
             End If
         Next
         lngRow = vsfMoney.Rows - 1
         vsfMoney.Rows = vsfMoney.Rows + mrsItems.RecordCount
         mrsItems.MoveFirst
        
         For i = 1 To mrsItems.RecordCount
             vsfMoney.RowData(lngRow + i) = mrsItems!项目ID
             vsfMoney.TextMatrix(lngRow + i, 0) = mrsItems!项目名称
             mrsInComes.Filter = "项目ID=" & mrsItems!项目ID
            cur应收 = 0: cur实收 = 0
             For j = 1 To mrsInComes.RecordCount
                 cur应收 = cur应收 + mrsInComes!应收
                 cur实收 = cur实收 + mrsInComes!实收
                 If mrsItems!性质 = 3 Then cur病历 = cur病历 + mrsInComes!实收
                 mrsInComes.MoveNext
             Next
             vsfMoney.TextMatrix(lngRow + i, 1) = Format(cur应收, "0.00")
             vsfMoney.TextMatrix(lngRow + i, 2) = Format(cur实收, "0.00")
             cur合计 = cur合计 + cur实收
             mrsItems.MoveNext
         Next
         vsfMoney.Redraw = True
     End If
End Sub

Private Sub ShowAcceptFromInput()
    Dim lng项目id As Long, bln病历 As Boolean, str费别 As String
    Dim cur应收 As Currency, cur实收 As Currency, cur合计 As Currency, cur病历 As Currency
    Dim lngRow As Long, i As Long, j As Long
    Dim dblMoney As Double
    
    If mbytMode = 2 And Not mrsBill Is Nothing Then
            mrsBill.MoveFirst
            '如果预约时,没有建立档案,接收时可以更改费别,
            If Nvl(mrsBill!费别) <> NeedName(cbo费别.Text) Then
                '费别不一致 需要重新计算
                str费别 = NeedName(cbo费别.Text)
                mrsBill.MoveFirst
                vsfMoney.Rows = mrsBill.RecordCount + 1
                For i = 1 To mrsBill.RecordCount
                    vsfMoney.TextMatrix(i, 0) = mrsBill!项目
                    vsfMoney.TextMatrix(i, 1) = Format(mrsBill!应收, "0.00")
'                    dblMoney = Val(Nvl(mrsBill!实收))
                    cur实收 = GetActualMoney(str费别, mrsBill!收入项目ID, mrsBill!应收, mrsBill!收费细目ID)
                    vsfMoney.TextMatrix(i, 2) = Format(cur实收, "0.00")
                    cur合计 = cur合计 + cur实收
                    mrsBill.MoveNext
                Next
                txt合计.Text = Format(cur合计, "0.00")
                lbl合计.Caption = txt合计.Text
            Else
                mrsBill.MoveFirst
                vsfMoney.Rows = mrsBill.RecordCount + 1
                For i = 1 To mrsBill.RecordCount
                    vsfMoney.TextMatrix(i, 0) = mrsBill!项目
                    vsfMoney.TextMatrix(i, 1) = Format(mrsBill!应收, "0.00")

                    vsfMoney.TextMatrix(i, 2) = Format(mrsBill!实收, "0.00")
                    cur合计 = cur合计 + mrsBill!实收
                    mrsBill.MoveNext
                Next
                txt合计.Text = Format(cur合计, "0.00")
                lbl合计.Caption = txt合计.Text
            End If
        End If
        '问题:31182
        cur合计 = Val(txt合计.Text)
        Call ReCalc预约接收发卡
          '60171 预约接收时,需要重新计算卡费和挂号费,此时不存在连续挂号
        If Not mrsItems Is Nothing Then
            cur合计 = GetRegistMoney
        End If
End Sub

Private Sub ShowRegistFromInput()
    '功能：根据当前界面输入的号别,读取挂号费用集,显示在表格中
    Dim lng项目id As Long, bln病历 As Boolean, str费别 As String
    Dim cur应收 As Currency, cur实收 As Currency, cur合计 As Currency, cur病历 As Currency
    Dim lngRow As Long, i As Long, j As Long
    Dim dblMoney As Double, rsItems As ADODB.Recordset, rsIncomes As ADODB.Recordset
    Dim str记录ID As String, strTemp As String
    Dim strReadSQL As String, rsRead As ADODB.Recordset
    Dim str医生姓名 As String, lng病人ID As Long
    If mblnReadBooking Then Exit Sub
    If mblnBuyHisBook = False Then
        If (mbytMode = 2 Or mbytMode = 0 And mbytInState = 0 And mstrNoIn <> "") And mTy_Para.bln预约接收确定挂号费 = False Then
            If mbytMode = 2 And Not mrsBill Is Nothing Then
                mrsBill.MoveFirst
                '如果预约时,没有建立档案,接收时可以更改费别,
                If Nvl(mrsBill!费别) <> NeedName(cbo费别.Text) Then
                    '费别不一致 需要重新计算
                    str费别 = NeedName(cbo费别.Text)
                    mrsBill.MoveFirst
                    vsfMoney.Rows = mrsBill.RecordCount + 1
                    For i = 1 To mrsBill.RecordCount
                        vsfMoney.TextMatrix(i, 0) = mrsBill!项目
                        vsfMoney.TextMatrix(i, 1) = Format(mrsBill!应收, "0.00")
    '                    dblMoney = Val(Nvl(mrsBill!实收))
                        cur实收 = GetActualMoney(str费别, mrsBill!收入项目ID, mrsBill!应收, mrsBill!收费细目ID)
                        vsfMoney.TextMatrix(i, 2) = Format(cur实收, "0.00")
                        cur合计 = cur合计 + cur实收
                        mrsBill.MoveNext
                    Next
                    txt合计.Text = Format(cur合计, "0.00")
                    lbl合计.Caption = txt合计.Text
                Else
                    mrsBill.MoveFirst
                    vsfMoney.Rows = mrsBill.RecordCount + 1
                    For i = 1 To mrsBill.RecordCount
                        vsfMoney.TextMatrix(i, 0) = mrsBill!项目
                        vsfMoney.TextMatrix(i, 1) = Format(mrsBill!应收, "0.00")
    
                        vsfMoney.TextMatrix(i, 2) = Format(mrsBill!实收, "0.00")
                        cur合计 = cur合计 + mrsBill!实收
                        mrsBill.MoveNext
                    Next
                    txt合计.Text = Format(cur合计, "0.00")
                    lbl合计.Caption = txt合计.Text
                End If
            End If
            '问题:31182
            cur合计 = Val(txt合计.Text)
            Call ReCalc预约接收发卡
              '60171 预约接收时,需要重新计算卡费和挂号费,此时不存在连续挂号
            If Not mrsItems Is Nothing Then
                cur合计 = GetRegistMoney
            End If
            GoTo CalcOther:
            Exit Sub
        End If
    End If
    If chkCancel.Value = 1 Then Exit Sub
    If chkPrint.Value = 1 Then Exit Sub

    Call ClearMoney

    '读取挂号费用
    If txt号别.Text = "+" Then    '仅购病历
        lng项目id = 0
        bln病历 = True

        chk病历费.Enabled = False
        chk病历费.Value = 0

        mbln建病案 = False
        mlng挂号科室ID = UserInfo.部门ID
        mstr医生姓名 = "": mlng医生ID = 0
        txt科室.Text = ""
        cbo医生.Clear
        cbo医生.Enabled = False
        lbl急.Visible = False
    ElseIf txt号别.Text <> "" Then
        '134441:李南春，2019/1/12，不相等说明确定了号别后列表选择又发生了改变
        If mlngPreRow <> 0 Then vsfPlan.Row = mlngPreRow
        If vsfPlan.Row > vsfPlan.Rows - 1 Then
            lngRow = 0
        Else
            lngRow = vsfPlan.Row
        End If
        If lngRow = 0 Then
            mbln建病案 = False
            mlng挂号科室ID = 0
            mstr医生姓名 = ""
            mlng医生ID = 0
            
            If mbytMode <> 2 Then
                chk病历费.Enabled = False
                chk病历费.Value = 0
            End If
            txt科室.Text = ""
            cbo医生.Clear
            lbl急.Visible = False
            Exit Sub
        End If

        lng病人ID = 0
        If Not mrsInfo Is Nothing Then
            If Not mrsInfo.EOF Then lng病人ID = Val(Nvl(mrsInfo!病人ID))
        End If
        str记录ID = ""
        strTemp = vsfPlan.Cell(flexcpData, lngRow, vsfPlan.ColIndex("IDS"))
        If Val(strTemp) <> 0 Then
            str记录ID = "2|" & Val(strTemp)
        End If
        If str记录ID = "" Then str记录ID = "3|" & vsfPlan.TextMatrix(lngRow, vsfPlan.ColIndex("号别"))
        
        lng项目id = Val(Split(vsfPlan.TextMatrix(lngRow, GetCol("IDS")), ",")(1))
        strReadSQL = "Select Zl_Custom_Getregeventitem([1],[2],[3],[4],[5],[6],[7]) As 项目ID From Dual"
        Set rsRead = zlDatabase.OpenSQLRecord(strReadSQL, Me.Caption, lng病人ID, txtPatient.Text, txtIDCard.Text, _
                                            CDate(IIf(IsDate(txt出生日期.Text) = False, "3000-01-01", txt出生日期.Text)), NeedName(cbo性别.Text), txt年龄.Text & IIf(cbo年龄单位.Visible, cbo年龄单位.Text, ""), str记录ID)
        If Not rsRead.EOF Then
            If Val(Nvl(rsRead!项目ID)) <> 0 Then lng项目id = Val(Nvl(rsRead!项目ID))
        End If
        bln病历 = chk病历费.Value = 1

        If mbytMode <> 2 Then chk病历费.Enabled = True
        mbln建病案 = vsfPlan.TextMatrix(lngRow, GetCol("病案")) <> ""
        lbl急.Visible = vsfPlan.RowData(lngRow) < 0
        cbo医生.Enabled = False
       
        mlng挂号科室ID = Abs(vsfPlan.RowData(lngRow))
        str医生姓名 = NeedName(cbo医生.Text)
        mstr医生姓名 = vsfPlan.TextMatrix(lngRow, GetCol("医生"))
        mlng医生ID = CLng(Split(vsfPlan.TextMatrix(lngRow, GetCol("IDS")), ",")(2))

        txt科室.Text = vsfPlan.TextMatrix(lngRow, GetCol("科室"))
        cbo医生.Clear
        cbo医生.TabStop = False
        If mstr医生姓名 <> "" Then
            cbo医生.AddItem mstr医生姓名
            cbo医生.ItemData(cbo医生.NewIndex) = mlng医生ID
            cbo医生.ListIndex = 0
        ElseIf Not mblnStation Then     '如果要求输医生,号别没有确定医生,列出科室可选医生
            cbo医生.Enabled = gbln医生
            If gbln医生 Then
                Call Fill医生(mlng挂号科室ID)
                zlControl.CboLocate cbo医生, str医生姓名
                mstr医生姓名 = NeedName(cbo医生.Text)
                If mstr医生姓名 = "" Then
                    mlng医生ID = 0
                Else
                    mlng医生ID = cbo医生.ItemData(cbo医生.ListIndex)
                End If
            End If
        End If
        
    End If
    Call AdjustInfoPosition
    str费别 = NeedName(cbo费别.Text)
    Set mrsItems = Nothing
    Set mrsInComes = Nothing
    Call ReadRegistPrice(lng项目id, bln病历, mblnAddCardItem, str费别, rsItems, rsIncomes, 0, mintInsure, _
        txt号别.Text, 10, mlng挂号科室ID, mobjfrmPatiInfo.mstrPriceGrade, _
        IIf(dtpAppointmentDate.Visible Or mbytMode = 2, Format(dtpAppointmentDate.Value, "yyyy-mm-dd") & " 23:59:59", ""), gCurSendCard.lng收费细目ID)
    
    If mintInsure <> 0 Then
        If MCPAR.挂号检查项目 = True Then
            If gclsInsure.CheckItem(mintInsure, 2, 0, rsItems) = False Then
                MsgBox "医保病人收费项目检查失败，不能继续挂号！", vbInformation, gstrSysName
                mblnUnload = True
                Call ClearBill: Exit Sub
            End If
        End If
    End If
    
    Call ReadRegistPrice(lng项目id, bln病历, mblnAddCardItem, str费别, mrsItems, mrsInComes, lng病人ID, _
        mintInsure, txt号别.Text, mbytMode, , mobjfrmPatiInfo.mstrPriceGrade, _
    IIf(dtpAppointmentDate.Visible Or mbytMode = 2, Format(dtpAppointmentDate.Value, "yyyy-mm-dd") & " 23:59:59", ""), gCurSendCard.lng收费细目ID)

    '显示挂号费用
    If Not mrsItems Is Nothing Then
        vsfMoney.Redraw = False
        vsfMoney.Rows = mrsItems.RecordCount + 1
        mrsItems.MoveFirst
        For i = 1 To mrsItems.RecordCount
            If mrsItems!性质 = 4 Then
                vsfMoney.RowData(i) = mrsItems!项目ID
            End If
            vsfMoney.TextMatrix(i, 0) = mrsItems!项目名称

            cur应收 = 0: cur实收 = 0
            mrsInComes.Filter = "项目ID=" & mrsItems!项目ID
            For j = 1 To mrsInComes.RecordCount
                cur应收 = cur应收 + mrsInComes!应收
                cur实收 = cur实收 + mrsInComes!实收
                If mrsItems!性质 = 3 Then cur病历 = cur病历 + mrsInComes!实收
                mrsInComes.MoveNext
            Next

            vsfMoney.TextMatrix(i, 1) = Format(cur应收, "0.00")
            vsfMoney.TextMatrix(i, 2) = Format(cur实收, "0.00")
            cur合计 = cur合计 + cur实收
            mcur病历 = cur病历
            mrsItems.MoveNext
        Next
        vsfMoney.Redraw = True

    End If

CalcOther:
    '预交款支付重新设置
    '77786,冉俊明,2014-9-2,勾选优先使用预交款缴款,挂号时,没有默认减少冲减
    '74550,冉俊明,2014-7-2,在病人来院就诊,医生在门诊医生站挂号时能够选择结算方式(包含性质为7的一卡通结算)
    If (gblnPrePayPriority Or (mblnStation And Not mblnStationPrice And Not cbo结算方式.Visible)) And mdbl预交余额 >= cur合计 And mblnAddCardItem = False Then
        For i = 1 To vsfPay.Rows - 1
            If vsfPay.RowData(i) = 0 And vsfPay.TextMatrix(i, 0) <> "" Then
                vsfPay.TextMatrix(i, 1) = Format(cur合计, "0.00")
                vsfPay.TextMatrix(i, 6) = mdbl预交余额
            End If
        Next i
    Else
        For i = 1 To vsfPay.Rows - 1
            If vsfPay.RowData(i) = 0 And vsfPay.TextMatrix(i, 0) <> "" Then
                vsfPay.TextMatrix(i, 1) = Format(0, "0.00")
                vsfPay.TextMatrix(i, 6) = mdbl预交余额
            End If
        Next i
    End If
    
    '卡费和挂号费用一起收时,禁用预交款
    If mblnAddCardItem Then ShowDeposit (False)
    
    
    '计算并显示个人帐户支付额
    Call CalcYBMoney
     
    '显示累加费用
    txt合计.Text = Format(cur合计 + mcur合计, "0.00")
    lbl合计.Caption = txt合计.Text
    
    Call Set连续挂号
    '显示挂免费号,不算病历费
    If Me.Visible Then
        lblFree.Visible = (cur合计 - cur病历) = 0
    Else
        lblFree.Visible = False
    End If
End Sub

Private Sub txt找补_GotFocus()
    Call zlControl.TxtSelAll(txt找补)
End Sub

Private Sub YBIdentifyCancel()
'功能：取消医保病人身份验证
    Dim lng病人ID As Long
    
    If mbytInState = 0 And mintInsure <> 0 And mstrYBPati <> "" And txtPatient.Text <> "" Then
        If UBound(Split(mstrYBPati, ";")) >= 8 Then
            If IsNumeric(Split(mstrYBPati, ";")(8)) And Val(Split(mstrYBPati, ";")(8)) <> 0 Then
                lng病人ID = Val(CLng(Split(mstrYBPati, ";")(8)))
            End If
        End If
        If lng病人ID <> 0 Then
            Call gclsInsure.IdentifyCancel(3, lng病人ID, mintInsure)
        End If
    End If
End Sub



Private Function StationDelete(ByVal strNO As String, Optional str划价NO As String) As Boolean
'功能：检查指定的挂号单是否允许退号(未收费,待接诊)
'返回：str划价NO=同时要删除的划价单
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, lng病人ID As Long
    
    On Error GoTo errH
    
    '1-执行人及病人状态判断
    strSQL = "Select 病人ID,执行人,执行状态 From 病人挂号记录 Where NO=[1] and 记录性质=1 and 记录状态=1"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
    If rsTmp.EOF Then
        MsgBox "指定的挂号单不存在，该单据可能已经退号。", vbInformation, gstrSysName
        Exit Function
    End If
    If Nvl(rsTmp!执行状态, 0) <> 2 Then
        MsgBox "该病人" & Decode(Nvl(rsTmp!执行状态, 0), 0, "不处于直接挂号就诊状态", 1, "已经完成就诊") & "，不能退号。", vbInformation, gstrSysName
        Exit Function
    End If
    If Nvl(rsTmp!执行人) <> UserInfo.姓名 Then
        MsgBox "该病人不是在你这儿挂的号，不能退号。", vbInformation, gstrSysName
        Exit Function
    End If
    lng病人ID = rsTmp!病人ID
    
    '2-挂号金额判断:有现金等其他非预交结算的不是医生站挂号
    strSQL = "Select Sum(冲预交) as 金额 From 病人预交记录 A,结算方式 B " & _
            " Where A.结算方式=B.名称 And A.记录性质=4 And A.记录状态=1 And A.NO=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
    If Not rsTmp.EOF Then
        If Nvl(rsTmp!金额, 0) > 0 Then
            MsgBox "该挂号采用了其他结算方式，不能在这里退号。", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    '3-收费判断
    strSQL = "Select NO,记录状态 From 门诊费用记录 " & _
            " Where 记录性质=1 And 病人ID=[1] And 记录状态 IN(0,1,3) And 序号=1 And 摘要 Like [2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, "%" & strNO & "%")
    If Not rsTmp.EOF Then
        If Nvl(rsTmp!记录状态, 0) = 1 Then
            MsgBox "该挂号单对应的费用已经被单独收费，不能退号。", vbInformation, gstrSysName
            Exit Function
        ElseIf Nvl(rsTmp!记录状态, 0) = 0 Then
            str划价NO = rsTmp!NO
        End If
    End If
    
    '4-医嘱判断
    strSQL = "Select Count(*) as Num From 病人医嘱记录 Where 挂号单=[1] And 医嘱状态<>4"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
    If Nvl(rsTmp!Num, 0) > 0 Then
        MsgBox "病人已经下达医嘱，不能退号。", vbInformation, gstrSysName
        Exit Function
    End If
    
    StationDelete = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Function Check复诊(ByVal lng病人ID As Long, ByVal lng执行部门ID As Long) As Boolean
'功能:判断病人是否再次到“相同临床性质的临床科室”挂号
'     包括挂过号的,或住过院的,复诊不好确定时间
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    strSQL = "Select Zl1_Fun_Getreturnvisit([1],[2]) As 复诊标志 From Dual"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, lng执行部门ID)
    Check复诊 = Val(Nvl(rsTmp!复诊标志)) = 1
End Function

Private Sub Set连续挂号()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重新计算应缴款合计数
    '编制:刘兴洪
    '日期:2009-12-02 12:02:51
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '计算缴款合计给文本框
    Dim strSQL As String, rsTemp As ADODB.Recordset
'    strSQL = "Select 性质" & vbNewLine & _
'                        "From 结算方式" & vbNewLine & _
'                        "Where 名称 = [1] And Rownum < 2" & vbNewLine & _
'                        "Union" & vbNewLine & _
'                        "Select a.性质" & vbNewLine & _
'                        "From 结算方式 A, 医疗卡类别 B" & vbNewLine & _
'                        "Where b.名称 = [1] And a.名称 = b.结算方式 And Rownum < 2" & vbNewLine & _
'                        "Union" & vbNewLine & _
'                        "Select a.性质 From 结算方式 A, 消费卡类别目录 B Where b.名称 = [1] And a.名称 = b.结算方式 And Rownum < 2"
'    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, cbo结算方式.Text)
'    If rsTemp.RecordCount <> 0 Then
'        If Val(Nvl(rsTemp!性质)) <> 7 And Val(Nvl(rsTemp!性质)) <> 8 Then
'            txt本次应缴.Text = Format(mcur应缴 + GetRegistMoney, "0.00")
'        Else
'            txt本次应缴.Text = Format(GetRegistMoney(False, True), "0.00")
'        End If
'    Else
    txt本次应缴.Text = Format(mcur应缴 + GetRegistMoney - GetPayedMoney, "0.00")
'    End If
'    cmd结束挂号.Visible = mint挂号数 > 0 And mintInsure <> 0     '医保病人才会增加结算挂号按钮
'    txt缴款.Enabled = Not cmd结束挂号.Visible
'    txt找补.Enabled = Not cmd结束挂号.Visible
End Sub

Private Function GetPayedMoney() As Double
    '获取已交款金额
    Dim i As Integer
    Dim dblReturn As Double
    If mbytMode = 4 Or chkCancel.Value = 1 Then Exit Function
    For i = 1 To vsfPay.Rows - 1
        dblReturn = dblReturn + (Val(vsfPay.TextMatrix(i, 1)) - Val(vsfPay.TextMatrix(i, 7)))
    Next i
    GetPayedMoney = dblReturn
End Function

Private Sub zlPatiMoveCmdCtrl()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据按钮状态,移动病人信息的相关按钮
    '编制:刘兴洪
    '日期:2010-01-15 10:02:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim sngLeft As Single
    sngLeft = cmdLookup.Left
    If cmdLookup.Visible Then sngLeft = sngLeft + cmdLookup.Width + 45
    If cmdCard.Visible Then
       cmdCard.Left = sngLeft: sngLeft = sngLeft + cmdCard.Width + 45
    End If
    If cmdMore.Visible Then
       cmdMore.Left = sngLeft: sngLeft = sngLeft + cmdMore.Width + 45
    End If
    If cmdComminuty.Visible Then
       cmdComminuty.Left = sngLeft: sngLeft = sngLeft + cmdComminuty.Width + 45
    End If
    If cmdYb.Visible Then cmdYb.Left = sngLeft + 45
End Sub

Private Function IsCheckReservationSameDept(ByVal lng科室ID As Long, ByVal strConditions As String, ByVal str预约时间 As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '功能：检查预约挂号是否在同一科室中已经存在预约
    '入参：strConditions: 比如:病人ID=...或身份证号=
    '出参：
    '返回：存在返回true,否则返回False
    '编制：刘兴洪
    '日期：2010-03-17 09:44:11
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim varData As Variant, strWhere As String
    On Error GoTo Hd
    varData = Split(strConditions, "=")
    Select Case varData(0)
    Case "病人ID"
            strWhere = " And A.病人ID=[2]"
    Case "身份证号"
            strWhere = " And B.身份证号=[3]"
     Case "就诊卡号"
            strWhere = " And B.就诊卡号=[3]"
    Case Else
            strWhere = strConditions
    End Select
    strSQL = "" & _
    "   Select  1 " & _
    "   From 门诊费用记录  A,病人信息 B " & _
    "   Where A.病人ID=B.病人ID And A.记录性质=4 and 记录状态=0  " & _
    "               and A.发生时间 between [4]  and [5]  and A.病人科室ID+0=[1] " & strWhere
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "检查预约挂号是否已经挂过", lng科室ID, Val(varData(1)), CStr(varData(1)), CDate(str预约时间), CDate(str预约时间) + 1 - 1 / 24 / 60 / 60)
    IsCheckReservationSameDept = (rsTemp.RecordCount <> 0)
    Exit Function
Hd:
    If ErrCenter() = 1 Then Resume
    SaveErrLog
End Function
Private Sub SetRePrintPatiEnabled(ByVal blnEdit As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:是否允许修改病人信息值
    '编制:刘兴洪
    '日期:2011-01-31 10:33:04
    '问题:35544
    '---------------------------------------------------------------------------------------------------------------------------------------------
    txt年龄.Locked = Not blnEdit
    cbo年龄单位.Locked = Not blnEdit
    cbo性别.Locked = Not blnEdit
    SetPatiEnable blnEdit
    txt年龄.Enabled = blnEdit And Not mTy_Para.bln禁止输入年龄
    cbo年龄单位.Enabled = blnEdit And Not mTy_Para.bln禁止输入年龄
    cbo性别.Enabled = blnEdit
    cbo付款方式.Enabled = Not blnEdit And Not mblnStation    '56263
    cbo家庭地址.Enabled = Not blnEdit
    cbo户口地址.Enabled = Not blnEdit
    padd家庭地址.Enabled = Not blnEdit: padd家庭地址.ControlLock = blnEdit
    padd户口地址.Enabled = Not blnEdit: padd户口地址.ControlLock = blnEdit
    cbo费别.Enabled = Not blnEdit
    cbo结算方式.Enabled = Not blnEdit
    txt门诊号.Enabled = Not blnEdit
    txt家庭电话.Enabled = Not blnEdit
    txtIDCard.Enabled = Not blnEdit
    '74017:李南春，2014-6-17，挂号重打不允许编辑更多病人信息界面的内容
    cmdCard.Enabled = False
End Sub
Public Function zlGet门诊号() As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:是否生成门诊号
    '返回:门诊号,如果未生成,则返回空
    '编制:刘兴洪
    '日期:2011-02-28 15:27:22
    '问题:36028
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mTy_Para.bln预约不产生门诊号 And mbytMode = 1 Then Exit Function
    If gbln自动门诊号 Or mblnStation Or mbln建病案 Then     '要求根据参数来设置 好别要求建立病案的 必须产生门诊号 以便建立病案
        zlGet门诊号 = zlDatabase.GetNextNo(3)
    Else
        zlGet门诊号 = ""
    End If
End Function

Private Function zlCommitPlugInpati(ByVal strCardNo As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:提交插件数据
    '返回:成功,返回true,否则的返回False
    '编制:刘兴洪
    '日期:2011-07-22 14:13:11
    '问题:40012
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsPatiInfor As ADODB.Recordset, str年龄 As String, str出生日期 As String
    Err = 0: On Error GoTo errHandle
    If CreatePlugInOK(mlngModul) = False Then zlCommitPlugInpati = True: Exit Function
    If mblnNotQuery = False Then zlCommitPlugInpati = True:  Exit Function
    If Not zlInitPati(rsPatiInfor) Then Exit Function
    str年龄 = Trim(txt年龄.Text)
    If IsNumeric(str年龄) Then str年龄 = str年龄 & cbo年龄单位.Text
    With mobjfrmPatiInfo
        If .txt出生时间 = "__:__" Then
            str出生日期 = IIf(IsDate(.txt出生日期.Text), .txt出生日期.Text, "")
        Else
            str出生日期 = IIf(IsDate(.txt出生日期.Text), "" & .txt出生日期.Text & " " & .txt出生时间.Text & "", "")
        End If
        rsPatiInfor.AddNew
        rsPatiInfor!姓名 = .txtPatient.Text
        rsPatiInfor!性别 = NeedName(cbo性别.Text)
        '89242:李南春,2015/12/7,读取病人地址信息
        If mblnStructAdress Then
            rsPatiInfor!家庭地址 = IIf(Trim(.padd家庭地址.Value) = "", padd家庭地址.Value, .padd家庭地址.Value)
        Else
            rsPatiInfor!家庭地址 = IIf(Trim(.cbo家庭地址.Text) = "", cbo家庭地址.Text, .cbo家庭地址.Text)
        End If
        rsPatiInfor!费别 = NeedName(cbo费别.Text)
        rsPatiInfor!身份证号 = Trim(.txt身份证号.Text)
        rsPatiInfor!医疗付款方式 = NeedName(cbo付款方式.Text)
        rsPatiInfor!医保号 = .txtPatiMCNO(0).Text
        rsPatiInfor!年龄 = str年龄
        rsPatiInfor!国籍 = NeedName(.cbo国籍.Text)
        rsPatiInfor!民族 = NeedName(.cbo民族.Text)
        rsPatiInfor!婚姻状况 = NeedName(.cbo婚姻.Text)
        rsPatiInfor!职业 = NeedName(.cbo职业.Text, True)
        rsPatiInfor!出生日期 = IIf(str出生日期 <> "", CDate(str出生日期), Null)
        rsPatiInfor!工作单位 = .txt单位名称.Text
        rsPatiInfor!合同单位ID = Val(.txt单位名称.Tag)
        rsPatiInfor!区域 = Trim(.txt区域.Text)
        rsPatiInfor!单位电话 = Trim(.txt单位电话.Text)
        rsPatiInfor!单位邮编 = Trim(.txt单位邮编.Text)
        rsPatiInfor!家庭电话 = Trim(.txt家庭电话.Text)
        rsPatiInfor!家庭邮编 = Trim(.txt家庭邮编.Text)
        rsPatiInfor.Update
    End With
    
    Err = 0: On Error Resume Next
    'CommitPatiInfo(byVal 卡号,rsInfo As ADO.RecordSet) As Boolean
    '传入本次发卡卡号，以及病人信息集。病人信息集为动态记录集，具备的字段与QueryPatiInfo所返回的对应。 _
    '因为是在事务外，挂号程序可以不对返回值作判断限制处理。
    If gobjPlugIn.CommitPatiInfo(strCardNo, rsPatiInfor) = False Then
        Exit Function
    End If
    zlCommitPlugInpati = True
    If Err <> 0 Then Exit Function
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Function zlReadPlugInPati(ByVal str卡号 As String, Optional blnHavePati As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取插建病人信息数据
    '入参:
    '出参:blnHavePati-是否接口返回了true,但有病人信息
    '编制:刘兴洪
    '日期:2011-06-10 17:50:09
    '问题:38841
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsPatiInfor As ADODB.Recordset
    On Error GoTo errHandle
    mblnNotQuery = False
    If CreatePlugInOK(mlngModul) = False Then zlReadPlugInPati = True: Exit Function
    If Not zlInitPati(rsPatiInfor) Then Exit Function
    'QueryPatiInfo(ByVal lngSys As Long, ByVal lngModule As Long, _
    ByVal str卡号 As String, ByRef rsInfo As ADO.Recordset)
    Err = 0: On Error Resume Next
    If gobjPlugIn.QueryPatiInfo(glngSys, mlngModul, str卡号, rsPatiInfor) = False Then
        If Err <> 0 Then zlReadPlugInPati = True: Exit Function
        mblnNotQuery = True
        Exit Function
    End If
    If Err <> 0 Then
        Exit Function
    End If
    If rsPatiInfor Is Nothing Then
        mblnNotQuery = True: GoTo ErrMsg:
    End If
    Err = 0: On Error GoTo errHandle
    blnHavePati = False
    If rsPatiInfor.State <> 1 Then
        mblnNotQuery = True
        zlReadPlugInPati = True: Exit Function
    End If
    If rsPatiInfor.RecordCount = 0 Then
        mblnNotQuery = True
        zlReadPlugInPati = True: Exit Function
    End If
    blnHavePati = True
    With mobjfrmPatiInfo
        txtPatient.Text = Nvl(rsPatiInfor!姓名) '会调用Change事件
        cbo性别.ListIndex = cbo.FindIndex(cbo性别, Nvl(rsPatiInfor!性别), True) '年龄在后面根据出生日期算
        cbo家庭地址.Text = Nvl(rsPatiInfor!家庭地址)
        '89242:李南春,2015/12/7,读取病人地址信息
        Call zlReadAddrInfo(padd家庭地址, Val(Nvl(rsPatiInfor!病人ID)), 0, 3, cbo家庭地址.Text)
        Call zlControl.CboSetIndex(cbo费别.Hwnd, cbo.FindIndex(cbo费别, "" & rsPatiInfor!费别, True))
'        txt门诊号.Text = Nvl(rsPatiInfor!门诊号, mstr门诊号)
'        txt门诊号.Enabled = (Val(txt门诊号.Text) = 0)
        
        txtIDCard.Text = Nvl(rsPatiInfor!身份证号, txtIDCard.Text) '身份证号:31182
        txtIDCard.Tag = Nvl(rsPatiInfor!身份证号, txtIDCard.Text)  '以便反过来再查
 
        '医疗付款方式
        If Not IsNull(rsPatiInfor!医疗付款方式) Then
            cbo付款方式.ListIndex = cbo.FindIndex(cbo付款方式, rsPatiInfor!医疗付款方式, True)
        ElseIf mstrYBPati <> "" Then
            cbo付款方式.ListIndex = cbo.FindIndex(cbo付款方式, "1", True)
        End If
        
        If Not IsNull(rsPatiInfor!医保号) And mlngOutModeMC <> 0 Then Call SetCboDefault(cbo医疗类别)
        '详细病人信息设置
        Call CopyInfoTofrmPatiInfo
        .txtPatiMCNO(0).Text = "" & Nvl(rsPatiInfor!医保号)
        .txtPatiMCNO(0).Tag = "" & Nvl(rsPatiInfor!医保号)
        .txtPatiMCNO(1).Text = .txtPatiMCNO(0).Text
        Call LoadOldData("" & rsPatiInfor!年龄, .txt年龄, .cbo年龄单位)
        .mblnChange = False
        .txt出生日期.Text = Format(IIf(IsNull(rsPatiInfor!出生日期), "____-__-__", rsPatiInfor!出生日期), "YYYY-MM-DD")
        .mblnChange = True
        .txt年龄.Text = Nvl(rsPatiInfor!年龄)
        .txt年龄.Tag = Nvl(rsPatiInfor!年龄)
        .cbo国籍.ListIndex = cbo.FindIndex(.cbo国籍, Nvl(rsPatiInfor!国籍), True)
        .cbo民族.ListIndex = cbo.FindIndex(.cbo民族, Nvl(rsPatiInfor!民族), True)
        .cbo婚姻.ListIndex = cbo.FindIndex(.cbo婚姻, Nvl(rsPatiInfor!婚姻状况), True)
        .cbo职业.ListIndex = cbo.FindIndex(.cbo职业, Nvl(rsPatiInfor!职业))
        .txt身份证号.Text = Nvl(rsPatiInfor!身份证号)
        .txt身份证号.Tag = .txt身份证号.Text
        .txt单位名称.Text = Nvl(rsPatiInfor!工作单位)
        .txt单位名称.Tag = Nvl(rsPatiInfor!合同单位ID)
        .txt区域.Text = Trim(Nvl(rsPatiInfor!区域))
        .txt区域.Tag = .txt区域.Text
        .txt单位电话.Text = Nvl(rsPatiInfor!单位电话)
        .txt单位邮编.Text = Nvl(rsPatiInfor!单位邮编)
        .txt家庭电话.Text = Nvl(rsPatiInfor!家庭电话)
        .txt家庭邮编.Text = Nvl(rsPatiInfor!家庭邮编)
        If Trim(.txt门诊号) = "" Then .txt门诊号 = zlGet门诊号
    End With
    zlReadPlugInPati = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Exit Function
ErrMsg:
    MsgBox "接口未转入病人信息,请检查!", vbInformation + vbOKOnly, gstrSysName
End Function
Private Function zlInitPati(ByRef rsPatiInfor As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始病人信息集
    '返回:病人信息集
    '编制:刘兴洪
    '日期:2011-06-10 17:52:18
    '问题:38841
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set rsPatiInfor = New ADODB.Recordset
    With rsPatiInfor
        If .State = adStateOpen Then .Close
        '病人ID,姓名,性别,年龄,出生日期,出生地点,身份证号,其他证件,身份,职业,家庭地址,家庭电话,家庭邮编,
        '工作单位,单位邮编,医保号,医疗付款方式,费别,国籍,民族,婚姻状况,区域
        
        .Fields.Append "病人ID", adDouble, 18, adFldIsNullable
        .Fields.Append "姓名", adLongVarChar, zlGetPatiInforMaxLen.intPatiName, adFldIsNullable
        .Fields.Append "性别", adLongVarChar, 4, adFldIsNullable
        .Fields.Append "年龄", adLongVarChar, 10, adFldIsNullable
        .Fields.Append "出生日期", adDate, , adFldIsNullable
        .Fields.Append "出生地点", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "身份证号", adLongVarChar, 18, adFldIsNullable
        .Fields.Append "其他证件", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "身份", adLongVarChar, 10, adFldIsNullable
        .Fields.Append "职业", adLongVarChar, 80, adFldIsNullable
        .Fields.Append "家庭地址", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "家庭电话", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "家庭邮编", adLongVarChar, 6, adFldIsNullable
        .Fields.Append "合同单位ID", adDouble, 18, adFldIsNullable
        .Fields.Append "工作单位", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "单位电话", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "单位邮编", adLongVarChar, 6, adFldIsNullable
        .Fields.Append "医保号", adLongVarChar, 30, adFldIsNullable
        .Fields.Append "医疗付款方式", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "费别", adLongVarChar, 10, adFldIsNullable
        .Fields.Append "国籍", adLongVarChar, 30, adFldIsNullable
        .Fields.Append "民族", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "婚姻状况", adLongVarChar, 4, adFldIsNullable
        .Fields.Append "区域", adLongVarChar, 30, adFldIsNullable
        .CursorLocation = adUseClient
        .Open , , adOpenStatic, adLockOptimistic
    End With
    zlInitPati = True
End Function

Private Function InitIDKindData() As Boolean
    Dim objCard As Card, rsTmp As ADODB.Recordset
    Dim lngCardID As Long, strSQL As String, IDkindStr As String
    If gobjSquare Is Nothing Then Exit Function
    On Error GoTo Errhand
    '90875:李南春,2016/3/2,医疗卡证件类型
    IDkindStr = "身|身份证号|0"
    strSQL = "Select 名称,缺省标志 from 证件类型  Where  名称 Not Like '其他%' and 名称 Not Like '%身份证'" & vbNewLine & _
            " And Not 名称 in (Select 名称 from  医疗卡类别 Where Nvl(是否证件,0)=0 or Nvl(是否启用,0)=0)"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If rsTmp.RecordCount > 0 Then
        Do While Not rsTmp.EOF
            IDkindStr = IDkindStr & ";" & Left(Nvl(rsTmp!名称), 1) & "|" & Nvl(rsTmp!名称) & "|0"
            rsTmp.MoveNext
        Loop
    End If
    Call IDKind证件.zlInit(Me, glngSys, mlngModul, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, IDkindStr, Me.txtIDCard)
    '强制把身份证号设置为手动输入
    Set objCard = IDKind证件.GetIDKindCard("身份证号")
    If Not objCard Is Nothing Then objCard.是否接触式读卡 = False: IDKind证件.Refrash
    
    Call IDKind.zlInit(Me, glngSys, mlngModul, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, "", Me.txtPatient)
    If mbytInState = 1 Then Exit Function
    lngCardID = Val(zlDatabase.GetPara("缺省医疗卡类别", glngSys, mlngModul, 0))
    mblnAlwaysSend = Val(Nvl(zlDatabase.GetPara("非严格控制时始终发卡", glngSys, mlngModul, 0), 0)) = 1
    If lngCardID <> 0 Then
        strSQL = "Select Nvl(是否严格控制,0) As 控制 From 医疗卡类别 Where ID=[1] And Nvl(是否启用,0)=1"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngCardID)
        If Not rsTmp.EOF Then
            IDKind.DefaultCardType = lngCardID
            mblnSendCard = ((Val(rsTmp!控制) = 0) And mblnAlwaysSend)
        End If
    Else
        strSQL = "Select Nvl(是否严格控制,0) As 控制,ID From 医疗卡类别 Where 缺省标志=1 And Nvl(是否启用,0)=1"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        If Not rsTmp.EOF Then
            IDKind.DefaultCardType = Val(rsTmp!ID)
            mblnSendCard = ((Val(rsTmp!控制) = 0) And mblnAlwaysSend)
        End If
    End If
    Set objCard = IDKind.GetfaultCard
    '76824，李南春，2014/8/19，医疗卡发卡处理
    Call InitSendCardPreperty(mlngModul, Val(IDKind.DefaultCardType))
    IDKind.ShowPropertySet = InStr(";" & mstrPrivs & ";", "参数设置") > 0
    If IDKind.Cards.按缺省卡查找 And Not objCard Is Nothing Then
        gobjSquare.bln缺省卡号密文 = objCard.卡号密文规则 <> ""
        gobjSquare.int缺省卡号长度 = objCard.卡号长度
        Set gobjSquare.objDefaultCard = objCard
       
    Else
        gobjSquare.bln缺省卡号密文 = IDKind.Cards.加密显示
        gobjSquare.int缺省卡号长度 = 100
    End If
    Exit Function
Errhand:
    If ErrCenter() = 1 Then Resume
End Function

Private Function LoadIdKindStr() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载IDKindStr
    '返回:加载成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-09-06 13:36:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strIdKindStr As String, varTemp As Variant, varData As Variant
    Dim i As Long, j As Long, strIDKindTemp As String, strTemp As String
    If gobjSquare.objSquareCard Is Nothing Then Exit Function
    '缺省定为发卡类别
    If mblnStation And mbytMode = 0 And mTy_Para.bln挂号必须刷卡 Then
        '38603
        strIdKindStr = gobjSquare.objSquareCard.zlGetIDKindStr("姓|姓名或就诊卡|0")
    Else
        strIdKindStr = gobjSquare.objSquareCard.zlGetIDKindStr(IDKind.IDkindStr)
    End If
    
    If Not (gCurSendCard.lng卡类别ID = 0 Or gCurSendCard.bln缺省标志) Then
        '短名|完成名|读卡标志|卡类别ID|卡号长度|缺省标志(1-当前缺省;0-非缺省)|
        '是否存在帐户(1-存在帐户;0-不存在帐户)|卡号密文(第几位至第几位加密,空为不加密);...
        varData = Split(strIdKindStr, ";")
        strIDKindTemp = ""
        For i = 0 To UBound(varData)
            varTemp = Split(varData(i), "|")
            If Val(varTemp(3)) = gCurSendCard.lng卡类别ID Then
                strTemp = ""
                For j = 0 To UBound(varTemp)
                    If j = 5 Then
                        strTemp = strTemp & "|" & 1
                    Else
                        strTemp = strTemp & "|" & varTemp(j)
                    End If
                Next
                If strTemp <> "" Then strTemp = Mid(strTemp, 2)
            Else
                '检查是否缺省标志
                If Val(varTemp(5)) = 1 Then
                    strTemp = ""
                    For j = 0 To UBound(varTemp)
                        If j = 5 Then
                            strTemp = strTemp & "|" & 0
                        Else
                            strTemp = strTemp & "|" & varTemp(j)
                        End If
                    Next
                    If strTemp <> "" Then strTemp = Mid(strTemp, 2)
                Else
                    strTemp = varData(i)
                End If
            End If
             strIDKindTemp = strIDKindTemp & ";" & strTemp
        Next
        strIdKindStr = Mid(strIDKindTemp, 2)
    End If
    IDKind.IDkindStr = strIdKindStr
    
    '取缺省的刷卡方式
    '短名|完成名|刷卡标志|卡类别ID|卡号长度|缺省标志(1-当前缺省;0-非缺省)|
    '是否存在帐户(1-存在帐户;0-不存在帐户)|卡号密文(第几位至第几位加密,空为不加密)
    '第7位后,就只能用索引,不然取不到数
    gobjSquare.bln缺省卡号密文 = IDKind.GetfaultCard.卡号密文规则 <> ""
    'gobjSquare.lng缺省卡类别ID = IDKind.GetCurCard.接口序号
    If mintIDKind > 0 And mintIDKind <= IDKind.ListCount Then IDKind.IDKind = mintIDKind
End Function
Private Sub InitCardSquareData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:创建或关闭结算卡对象
    '入参:blnClosed:关闭对象
    '编制:刘兴洪
    '日期:2010-01-05 14:51:23
    '问题:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strExpend As String
    '1-查阅,2-查阅冲销预约单据,3-查询被冲销单据
     
    Call InitIDKindData
End Sub
Private Sub Load支付方式()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载有效的支付方式
    '编制:刘兴洪
    '日期:2011-07-08 11:41:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim j As Long, strPayType As String, varData As Variant, varTemp As Variant, i As Long
    Dim rsTemp As ADODB.Recordset, blnFind As Boolean
    Dim strSQL As String, str性质 As String
    
    If mbln连续挂号 Then Exit Sub
    
    If mblnStation Then
        str性质 = "(3,7,8)"
    Else
        str性质 = "(1,2,3,7,8)"
    End If
    strSQL = _
        "Select B.编码,B.名称,Nvl(B.性质,1) as 性质,Nvl(A.缺省标志,0) as 缺省" & _
        " From 结算方式应用 A,结算方式 B" & _
        " Where A.应用场合=[1] And B.名称=A.结算方式 and B.性质 In " & str性质 & _
        " Order by B.编码"
        
    Err = 0: On Error GoTo errH
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, "挂号")
    
    Set mcolCardPayMode = New Collection
    '短|全名|刷卡标志|卡类别ID(消费卡序号)|长度|是否消费卡|结算方式|是否密文|是否自制卡;…
    If Not gobjSquare.objSquareCard Is Nothing Then
        strPayType = gobjSquare.objSquareCard.zlGetAvailabilityCardType
    End If
    varData = Split(strPayType, ";")
    
    With cbo结算方式
        .Clear: j = 0
        Do While Not rsTemp.EOF
            blnFind = False
            For i = 0 To UBound(varData)
                varTemp = Split(varData(i) & "|||||", "|")
                If varTemp(6) = Nvl(rsTemp!名称) Then
                    blnFind = True
                    Exit For
                End If
            Next
            If rsTemp!性质 = 3 Then mstr个人帐户 = rsTemp!名称: blnFind = True  '问题号:57711
            If rsTemp!性质 = 7 Or rsTemp!性质 = 8 Then blnFind = True
                         
            If Not blnFind Then
                .AddItem Nvl(rsTemp!编码) & "-" & Nvl(rsTemp!名称)
                mcolCardPayMode.Add Array("", Nvl(rsTemp!名称), 0, 0, 0, 0, Nvl(rsTemp!名称), 0, 0), "K" & j
                If Nvl(rsTemp!编码) & "-" & Nvl(rsTemp!名称) = gstr结算方式 Then
                     .ItemData(.NewIndex) = 1
                     .ListIndex = .NewIndex
                End If
                If Val(Nvl(rsTemp!缺省)) = 1 Then
                    If .ListIndex = -1 Then
                         .ItemData(.NewIndex) = 1: .ListIndex = .NewIndex
                    End If
                End If
                j = j + 1
            End If
            rsTemp.MoveNext
        Loop
    
        For i = 0 To UBound(varData)
            blnFind = False
            rsTemp.MoveFirst
            Do While Not rsTemp.EOF
                If Split(varData(i) & "|||||", "|")(6) = Nvl(rsTemp!名称) Then
                    blnFind = True
                    Exit Do
                End If
                rsTemp.MoveNext
            Loop
            If InStr(1, varData(i), "|") <> 0 And blnFind Then
                varTemp = Split(varData(i), "|")
                mcolCardPayMode.Add varTemp, "K" & j
                .AddItem varTemp(1): .ItemData(.NewIndex) = -1
                '74552,冉俊明,2014-7-2,挂号管理中设置默认结算方式时候可以选择结算方式性质为"7-一卡通结算"的结算方式
                If varTemp(1) = gstr结算方式 Then .ListIndex = .NewIndex
                j = j + 1
            End If
        Next
        If .ListCount > 0 And .ListIndex < 0 Then .ListIndex = 0
    End With
        
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function CheckBrushCard(ByVal dblMoney As Double) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查刷卡
    '返回:合法,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-18 14:52:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsMoney As ADODB.Recordset, str年龄 As String
    Dim strXmlIn As String, lng病人ID As Long
    Dim objCard As Card
    
    On Error GoTo errHandle
    '68991
    If mRegistFeeMode = EM_RG_记帐 Then CheckBrushCard = True: Exit Function
    If dblMoney = 0 Then
        '问题:51527
        CheckBrushCard = True: Exit Function
    End If
    
    If mCurCardPay.lng医疗卡类别ID = 0 Then
        MsgBox cbo结算方式.Text & "异常,请检查!", vbOKOnly + vbInformation, gstrSysName
        Exit Function
    End If
    If mstrYBPati <> "" Then
        MsgBox "不支持医保病人使用" & mCurCardPay.str名称 & "支付！", vbInformation, gstrSysName
        Exit Function
    End If
    If gobjSquare.objSquareCard Is Nothing Then
        MsgBox "使用" & mCurCardPay.str名称 & "支付必须先初始化接口部件！", vbInformation, gstrSysName
        Exit Function
    End If
    
    If mCurCardPay.bln消费卡 = False Then '三方卡支付必须有部件支持
        If gobjSquare.objSquareCard.zlGetCard(mCurCardPay.lng医疗卡类别ID, False, objCard) = False Then Exit Function
        If objCard Is Nothing Then
            MsgBox "使用" & mCurCardPay.str名称 & "支付必须先初始化接口部件！", vbInformation, gstrSysName
            Exit Function
        End If
        If objCard.接口程序名 = "" Then
            MsgBox "使用" & mCurCardPay.str名称 & "支付必须先初始化接口部件！", vbInformation, gstrSysName
            Exit Function
        End If
        Set mCurCardPay.objCard = objCard
    End If
    
    Call zlGetClassMoney(rsMoney)
    
     '弹出刷卡界面
    'zlBrushCard(frmMain As Object, _
    'ByVal lngModule As Long, _
    'ByVal rsClassMoney As ADODB.Recordset, _
    'ByVal lngCardTypeID As Long, _
    'ByVal bln消费卡 As Boolean, _
    'ByVal strPatiName As String, ByVal strSex As String, _
    'ByVal strOld As String, ByVal dbl金额 As Double, _
    'Optional ByRef strCardNo As String, _
    'Optional ByRef strPassWord As String, _
    Optional ByRef bln退费 As Boolean = False, _
    Optional ByRef blnShowPatiInfor As Boolean = False, _
    Optional ByRef bln退现 As Boolean = False, _
    Optional ByVal bln余额不足禁止 As Boolean = True, _
    Optional ByRef varSquareBalance As Variant, _
    Optional ByVal bln转预交 As Boolean = False, _
    Optional ByVal blnAllPay As Boolean = False, _
    Optional ByVal strXmlIn As String = "", _
    Optional ByVal str费用来源 As String, _
    Optional ByVal lng病人ID As Long) As Boolean
    str年龄 = Trim(txt年龄.Text)
    If IsNumeric(str年龄) Then str年龄 = str年龄 & cbo年龄单位.Text
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据指定支付类别,弹出刷卡窗口
    '入参:rsClassMoney:收费类别,金额
    '        lngCardTypeID-为零时,为老一卡通刷卡
    '       bln余额不足禁止-目前只针对消费卡,表示余额不足时,禁止继续操作,否则用余额进行支付
    '       str费用来源 - 当前支付费用的费用来源，多种用逗号分隔(使用消费卡支付时传入)
    '       lng病人ID - 病人ID(使用消费卡支付时传入)
   '58322
   strXmlIn = "<IN><CZLX>0</CZLX></IN>"
   If Not mrsInfo Is Nothing Then lng病人ID = Val(Nvl(mrsInfo!病人ID))
   If gobjSquare.objSquareCard.zlBrushCard(Me, mlngModul, rsMoney, mCurCardPay.lng医疗卡类别ID, mCurCardPay.bln消费卡, _
    txtPatient.Text, NeedName(cbo性别.Text), str年龄, dblMoney, mCurCardPay.str刷卡卡号, mCurCardPay.str刷卡密码, _
    False, True, False, True, Nothing, False, True, strXmlIn, "1", lng病人ID) = False Then Exit Function
    '保存前,一些数据检查
    'zlPaymentCheck(frmMain As Object, ByVal lngModule As Long, _
    ByVal strCardTypeID As Long, ByVal strCardNo As String, _
    ByVal dblMoney As Double, ByVal strNOs As String, _
    Optional ByVal strXMLExpend As String
    If gobjSquare.objSquareCard.zlPaymentCheck(Me, mlngModul, mCurCardPay.lng医疗卡类别ID, _
        mCurCardPay.bln消费卡, mCurCardPay.str刷卡卡号, dblMoney, "", "") = False Then Exit Function
        '暂无
''    zlGetAccountMoney(ByVal frmMain As Object, ByVal lngModule As Long, _
''    ByVal strCardTypeID As Long, _
''    ByVal strCardNo As String, strExpand As String, dblMoney As Double
'    '入参:frmMain-调用的主窗体
'    '        lngModule-模块号
'    '        strCardNo-卡号
'    '        strExpand-预留，为空,以后扩展
'    '出参:dblMoney-返回帐户余额
'    Dim strExpand As String, dbl帐户余额 As Double
'    If gobjSquare.objSquareCard.zlGetAccountMoney(Me, mlngModule, mCurCardPay.lng医疗卡类别ID, _
'          mCurCardPay.str刷卡卡号, strExpand, dbl帐户余额, mCurCardPay.bln消费卡) = False Then Exit Function
'    stbThis.Panels(4).Text = Format(dbl帐户余额, "0.00")
'    stbThis.Panels(4).ToolTipText = mCurCardPay.str结算方式 & "的帐户余额:" & Format(dbl帐户余额, "0.00")
'    mCurCardPay.dbl帐户余额 = Round(dbl帐户余额, 2)
    CheckBrushCard = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function zlInterfacePrayMoney(ByVal lngCard结帐ID As Long, ByVal lng挂号结帐ID As Long, ByRef cllPro As Collection, _
    ByRef cllThreeSwap As Collection, dblMoney As Double) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:接口支付金额
    '出参:cllPro-修改三方交易数据
    '        cll三方交易-增加三交方易数据
    '返回:支付成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-17 13:34:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSwapGlideNO As String, strSwapMemo As String, strSwapExtendInfor As String
    Dim strBalanceIDs As String
    
    If mCurCardPay.lng医疗卡类别ID = 0 Or dblMoney = 0 Then zlInterfacePrayMoney = True: Exit Function
    
    If mCurCardPay.Have挂号费 Then strBalanceIDs = lng挂号结帐ID
    If mCurCardPay.Have卡费 Then strBalanceIDs = strBalanceIDs & IIf(strBalanceIDs = "", "", ",") & lngCard结帐ID
    
'    If cbo结算方式.ItemData(cbo结算方式.ListIndex) <> -1 Then zlInterfacePrayMoney = True: Exit Function
    'zlPaymentMoney(ByVal frmMain As Object, _
    ByVal lngModule As Long, ByVal lngCardTypeID As Long, _
    ByVal bln消费卡 As Boolean, _
    ByVal strCardNo As String, ByVal strBalanceIDs As String, _
    byval  strPrepayNos as string , _
    ByVal dblMoney As Double, _
    ByRef strSwapGlideNO As String, _
    ByRef strSwapMemo As String, _
    Optional ByRef strSwapExtendInfor As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:帐户扣款交易
    '入参:frmMain-调用的主窗体
    '        lngModule-调用模块号
    '        strBalanceIDs-结帐ID,多个用逗号分离
    '        strPrepayNos-缴预交时有效. 预交单据号,多个用逗号分离
    '       strCardNo-卡号
    '       dblMoney-支付金额
    '出参:strSwapGlideNO-交易流水号
    '       strSwapMemo-交易说明
    '       strSwapExtendInfor-交易扩展信息: 格式为:项目名称1|项目内容2||…||项目名称n|项目内容n
    '返回:扣款成功,返回true,否则返回Flase
    '说明:
    '   在所有需要扣款的地方调用该接口,目前规划在:收费室；挂号室;自助查询机;医技工作站；药房等。
    '   一般来说，成功扣款后，都应该打印相关的结算票据，可以放在此接口进行处理.
    '   在扣款成功后，返回交易流水号和相关备注说明；如果存在其他交易信息，可以放在交易说明中以便退费.
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If gobjSquare.objSquareCard.zlPaymentMoney(Me, mlngModul, mCurCardPay.lng医疗卡类别ID, mCurCardPay.bln消费卡, mCurCardPay.str刷卡卡号, strBalanceIDs, "", dblMoney, strSwapGlideNO, strSwapMemo, strSwapExtendInfor) = False Then Exit Function
    '更新三交交易数据
     If lng挂号结帐ID <> 0 Then
        '问题:58322
        'mbytMode As Integer '0-挂号,1-预约,2-接收,3-取消预约 ,4-退号 预约有两种模式:0-挂号,此时预约要收费,1-预约,不收费
        If Not ((mbytMode = 0 Or mbytMode = 2) And mCurCardPay.bln消费卡) Then
            '消费卡已经在插入挂号记录时,已经扣款
            Call zlAddUpdateSwapSQL(False, lng挂号结帐ID, mCurCardPay.lng医疗卡类别ID, mCurCardPay.bln消费卡, mCurCardPay.str刷卡卡号, strSwapGlideNO, strSwapMemo, cllPro)
        End If
        
        Call zlAddThreeSwapSQLToCollection(False, lng挂号结帐ID, mCurCardPay.lng医疗卡类别ID, mCurCardPay.bln消费卡, mCurCardPay.str刷卡卡号, strSwapExtendInfor, cllThreeSwap)
    End If
    If lngCard结帐ID <> 0 Then
        If Not ((mbytMode = 0 Or mbytMode = 2) And mCurCardPay.bln消费卡) Then
                '消费卡已经在发卡记录时,已经扣款
                Call zlAddUpdateSwapSQL(False, lngCard结帐ID, mCurCardPay.lng医疗卡类别ID, mCurCardPay.bln消费卡, mCurCardPay.str刷卡卡号, strSwapGlideNO, strSwapMemo, cllPro)
        End If
        Call zlAddThreeSwapSQLToCollection(False, lngCard结帐ID, mCurCardPay.lng医疗卡类别ID, mCurCardPay.bln消费卡, mCurCardPay.str刷卡卡号, strSwapExtendInfor, cllThreeSwap)
    End If
    zlInterfacePrayMoney = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function zlGetClassMoney(ByRef rsMoney As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存时,初始化支付类别(收费类别,实收金额)
    '返回:加载成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-06-10 17:52:18
    '问题:38841
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strSQL  As String
    
    Err = 0: On Error GoTo Errhand:
    
    Set rsMoney = New ADODB.Recordset
    With rsMoney
        If .State = adStateOpen Then .Close
        .Fields.Append "收费类别", adLongVarChar, 20, adFldIsNullable
        '58322
        .Fields.Append "金额", adDouble, 18, adFldIsNullable
        .CursorLocation = adUseClient
        .Open , , adOpenStatic, adLockOptimistic
        If (mbytMode = 2 Or mbytMode = 0 And mbytInState = 0 And mstrNoIn <> "") And Not mTy_Para.bln预约接收确定挂号费 Then      '预约接收
            strSQL = "Select 收费类别,sum(nvl(实收金额,0)) as 实收 from 门诊费用记录 where NO=[1] and 记录性质=4 And 记录状态=0  Group by 收费类别"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrNoIn)
            Do While Not rsTemp.EOF
                 .AddNew
                !收费类别 = Nvl(rsTemp!收费类别, "无")
                !金额 = Val(Nvl(rsTemp!实收))
                .Update
                rsTemp.MoveNext
            Loop
              '处理预约接收时,发卡收费的状况(非接收时确定挂号费) 60171
            If Not mrsItems Is Nothing Then
                mrsItems.Filter = "性质=4"    '卡费
                If mrsItems.RecordCount > 0 Then
                    Do While Not mrsItems.EOF
                        mrsInComes.Filter = "项目ID=" & mrsItems!项目ID
                        rsMoney.Filter = "收费类别='" & Nvl(mrsItems!类别, "无") & "'"
                        If rsMoney.EOF Then
                            .AddNew
                        Else
                            rsMoney.Filter = 0
                        End If
                        !收费类别 = Nvl(mrsItems!类别, "无")
                        Do While Not mrsInComes.EOF
                            !金额 = Val(Nvl(!金额)) + Val(Nvl(mrsInComes!实收))
                            mrsInComes.MoveNext
                        Loop
                        .Update
                        mrsItems.MoveNext
                    Loop
                End If
                mrsItems.Filter = 0
            End If
            rsMoney.Filter = 0
            zlGetClassMoney = True
            Exit Function
        End If
        '58322
        mrsItems.Filter = 0
        If mrsItems.RecordCount <> 0 Then mrsItems.MoveFirst
        Do While Not mrsItems.EOF
            mrsInComes.Filter = "项目ID=" & mrsItems!项目ID
            rsMoney.Filter = "收费类别='" & Nvl(mrsItems!类别, "无") & "'"
            If rsMoney.EOF Then
                .AddNew
            Else
                rsMoney.Filter = 0
            End If
            !收费类别 = Nvl(mrsItems!类别, "无")
            Do While Not mrsInComes.EOF
                !金额 = Val(Nvl(!金额)) + Val(Nvl(mrsInComes!实收))
                mrsInComes.MoveNext
            Loop
            .Update
            mrsItems.MoveNext
        Loop
    End With
    rsMoney.Filter = 0
    zlGetClassMoney = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub AddCardDataSQL(ByVal lng病人ID As Long, ByVal dtCurdate As Date, _
    ByRef cllPro As Collection, ByRef lngCard结帐ID As Long, Optional ByVal bln记帐 As Boolean, _
    Optional ByVal lng项目id As Long = 0)

    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:就诊卡发放处理
    '入参:lng病人ID
    '       int记帐-卡费是否采用记帐方式
    '出参:lngCard结帐ID-卡费的结帐ID
    '编制:刘兴洪
    '日期:2011-07-07 04:36:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim byt操作类型 As Byte, strNO As String, strPassWord As String, strSQL As String
    Dim str原卡号 As String, str年龄 As String, strCard As String, str变动原因 As String
    Dim strICCard As String, lngBrushCardTypeID As Long, str结算方式 As String, strBrushCardNo As String
    Dim bln消费卡 As Boolean, blnInRange As Boolean   '范围内的卡
    Dim lngIndex As Long, byt变动类型 As Byte, lng结帐ID As Long
    Dim str密码  As String, strYLKNo As String
    Dim lngPay卡类别id As Long, blnPay消费卡 As Boolean, strPayCardNo As String
    On Error GoTo errHandle
    str密码 = Trim(mobjfrmPatiInfo.txt密码.Text)
    strCard = UCase(mobjfrmPatiInfo.txt卡号.Text): strICCard = IIf(mblnICCard, strCard, "")
    If Not ((strCard <> "" Or strICCard <> "")) Then Exit Sub
    
    lng结帐ID = 0: blnInRange = True
    '115168:李南春，2017/12/13，保存发卡的医疗卡类型
    If mCurSendCard.lng卡类别ID = 0 Then mCurSendCard = gCurSendCard
    If mCurSendCard.blnOneCard And mCurSendCard.bln严格控制 Then blnInRange = mlng磁卡领用ID > 0
    '77805
    If mrsItems Is Nothing Then
        blnInRange = False
    Else
        If lng项目id = 0 Then
            mrsItems.Filter = "性质=4"
            blnInRange = mrsItems.RecordCount <> 0
            If mrsItems.RecordCount > 0 Then
                mrsInComes.Filter = "项目ID=" & mrsItems!项目ID
            End If
        Else
            blnInRange = True
            mrsInComes.Filter = "项目ID=" & lng项目id
        End If
    End If
    '院外卡且不能发卡的,只能是绑定卡
    If bln发卡(strCard) = False Then
        blnInRange = False
    Else
        blnInRange = True
    End If
    If blnInRange Then
        blnInRange = True
        byt操作类型 = 0: byt变动类型 = 1
    Else
        blnInRange = False
        byt变动类型 = 11: byt操作类型 = 0
    End If
    
    str年龄 = Trim(txt年龄.Text)
    If IsNumeric(str年龄) Then str年龄 = str年龄 & cbo年龄单位.Text
    
    str变动原因 = "病人挂号发卡"
    
    strPassWord = zlCommFun.zlStringEncode(str密码)
    If blnInRange = False Then
          'Zl_医疗卡变动_Insert
           strSQL = "Zl_医疗卡变动_Insert("
          '      变动类型_In   Number,
          '发卡类型=1-发卡(或11绑定卡);2-换卡;3-补卡(13-补卡停用);4-退卡(或14取消绑定); ５-密码调整(只记录);6-挂失(16取消挂失)
          strSQL = strSQL & "" & byt变动类型 & ","
          '      病人id_In     住院费用记录.病人id%Type,
          strSQL = strSQL & "" & lng病人ID & ","
          '      卡类别id_In   病人医疗卡信息.卡类别id%Type,
          strSQL = strSQL & "" & mCurSendCard.lng卡类别ID & ","
          '      原卡号_In     病人医疗卡信息.卡号%Type,
          strSQL = strSQL & "'" & str原卡号 & "',"
          '      医疗卡号_In   病人医疗卡信息.卡号%Type,
          strSQL = strSQL & "'" & strCard & "',"
          '      变动原因_In   病人医疗卡变动.变动原因%Type,
          '      --变动原因_In:如果密码调整，变动原因为密码.加密的
          strSQL = strSQL & "'" & str变动原因 & "',"
          '      密码_In       病人信息.卡验证码%Type,
          strSQL = strSQL & "'" & strPassWord & "',"
          '      操作员姓名_In 住院费用记录.操作员姓名%Type,
          strSQL = strSQL & "'" & UserInfo.姓名 & "',"
          '      变动时间_In   住院费用记录.登记时间%Type,
          strSQL = strSQL & "to_date('" & Format(dtCurdate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
          '      Ic卡号_In     病人信息.Ic卡号%Type := Null,
          strSQL = strSQL & "'" & strICCard & "',"
          '      挂失方式_In   病人医疗卡变动.挂失方式%Type := Null
          strSQL = strSQL & "NULL)"
          lngCard结帐ID = 0
          zlAddArray cllPro, strSQL
    Else
        If gbln卡费仅划价 Or CheckIsPrice Then  '挂号是划价单，哪么卡应该也是划价单
            strNO = zlDatabase.GetNextNo(13)
            strYLKNo = zlDatabase.GetNextNo(16)  '医疗卡
            strSQL = "zl_门诊划价记录_Insert('" & strNO & "',1," & lng病人ID & ",NULL," & IIf(txt门诊号.Text = "", Null, txt门诊号.Text) & "," & _
                      "'" & NeedCode(cbo付款方式) & "','" & txtPatient.Text & "','" & NeedName(cbo性别.Text) & "','" & txt年龄.Text & cbo年龄单位.Text & "'," & _
                      "'" & NeedName(cbo费别.Text) & "',0," & UserInfo.部门ID & "," & _
                      UserInfo.部门ID & ",'" & UserInfo.姓名 & "',NULL," & gCurSendCard.rs卡费!收费细目ID & "," & _
                      "'" & gCurSendCard.rs卡费!收费类别 & "','" & gCurSendCard.rs卡费!计算单位 & "',NULL,1,1,0," & mlng挂号科室ID & ",NULL," & _
                      gCurSendCard.rs卡费!收入项目ID & ",'" & gCurSendCard.rs卡费!收据费目 & "'," & Format(gCurSendCard.rs卡费!现价, "0.000") & "," & _
                      Format(gCurSendCard.rs卡费!现价, "0.00") & "," & Format(gCurSendCard.rs卡费!现价, "0.00") & "," & "to_date('" & Format(dtCurdate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss')" & "," & _
                      "to_date('" & Format(dtCurdate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss')" & ",NULL,'" & UserInfo.姓名 & "','" & strYLKNo & "')"
            zlAddArray cllPro, strSQL
            
            '存在卡费需要生成住院费用记录
            strSQL = zlGetSaveCardFeeSQL(mCurSendCard.lng卡类别ID, 0, strYLKNo, lng病人ID, 0, UserInfo.部门ID, mlng挂号科室ID, 0, _
            zlStr.NeedName(cbo费别.Text), "", Trim(txtPatient.Text), zlStr.NeedName(cbo性别.Text), str年龄, _
            strCard, strPassWord, "挂号发卡", 0, 0, "", dtCurdate, mlng磁卡领用ID, gCurSendCard.rs卡费, _
            strICCard, mCurCardPay.lng医疗卡类别ID, mCurCardPay.bln消费卡, mCurCardPay.str刷卡卡号, , strNO)
            zlAddArray cllPro, strSQL
        Else
            strNO = zlDatabase.GetNextNo(16)  '医疗卡
            lng结帐ID = zlDatabase.GetNextId("病人结帐记录")
            '结算方式为空时为记帐方式
            '68991
            '137473:李南春，2019/1/24，三方卡支付卡费时才填写支付卡号
            If Not bln记帐 Then
                str结算方式 = mstrCard结算方式
                If mCurCardPay.Have卡费 Then
                    lngPay卡类别id = mCurCardPay.lng医疗卡类别ID
                    blnPay消费卡 = mCurCardPay.bln消费卡
                    strPayCardNo = mCurCardPay.str刷卡卡号
                End If
            End If
            strSQL = zlGetSaveCardFeeSQL(mCurSendCard.lng卡类别ID, byt操作类型, strNO, lng病人ID, 0, 0, mlng挂号科室ID, 0, _
             NeedName(cbo费别.Text), "", Trim(txtPatient.Text), NeedName(cbo性别.Text), str年龄, _
            strCard, strPassWord, str变动原因, IIf(mCurSendCard.bln变价 = False, mCurSendCard.dbl应收金额, mCurSendCard.dbl实收金额), mCurSendCard.dbl实收金额, str结算方式, _
            dtCurdate, mlng磁卡领用ID, gCurSendCard.rs卡费, strICCard, lngPay卡类别id, blnPay消费卡, strPayCardNo, lng结帐ID)
            zlAddArray cllPro, strSQL
        End If
    End If
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
 End Sub
 
Private Sub AddCardTypeSQL(ByVal intOper As Integer, ByVal lng卡类别ID As Long, ByVal strCode As String, ByVal str全名 As String, ByVal str短名 As String, _
                           ByVal lng卡号长度 As Long, ByRef colPro As Collection)
    Dim strSQL As String
    ' Zl_医疗卡类别_Update
        strSQL = "Zl_医疗卡类别_Update("
        '  Id_In           In 医疗卡类别.ID%Type,
        strSQL = strSQL & "" & lng卡类别ID & ","
        '  编码_In         In 医疗卡类别.编码%Type,
        strSQL = strSQL & "'" & strCode & "',"
        '  名称_In         In 医疗卡类别.名称%Type,
        strSQL = strSQL & "'" & str全名 & "',"
        '  短名_In         In 医疗卡类别.短名%Type,
        strSQL = strSQL & "'" & str短名 & "',"
        '  前缀文本_In     In 医疗卡类别.前缀文本%Type,
        strSQL = strSQL & "'" & "" & "',"
        '  卡号长度_In     In 医疗卡类别.卡号长度%Type,
        strSQL = strSQL & "" & lng卡号长度 & ","
        '  缺省标志_In     In 医疗卡类别.缺省标志%Type,
        strSQL = strSQL & "" & 0 & ","
        '  是否固定_In     In 医疗卡类别.是否固定%Type,
        strSQL = strSQL & "1,"
        '  是否严格控制_In In 医疗卡类别.是否严格控制%Type,
        strSQL = strSQL & "" & 0 & ","
        '  是否自制_In     In 医疗卡类别.是否自制%Type,
        strSQL = strSQL & "" & 0 & ","
        '  是否存在帐户_In In 医疗卡类别.是否存在帐户%Type,
        strSQL = strSQL & "" & 0 & ","
        '  是否全退_In     In 医疗卡类别.是否全退%Type,
        strSQL = strSQL & "0,"
        '  部件_In         In 医疗卡类别.部件%Type,
        strSQL = strSQL & "'" & "" & "',"
        '  备注_In         In 医疗卡类别.备注%Type,
        strSQL = strSQL & "'" & "" & "',"
        '  特定项目_In     In 医疗卡类别.特定项目%Type,
        strSQL = strSQL & "'" & strCode & "',"
        '    收费细目id_In   In 收费项目目录.ID%Type,
        strSQL = strSQL & "" & "0" & ","
        '  结算方式_In     In 医疗卡类别.结算方式%Type,
        strSQL = strSQL & "'" & "" & "',"
        '  是否启用_In     In 医疗卡类别.是否启用%Type,
        strSQL = strSQL & "1,"
        '  卡号密文_In     In 医疗卡类别.卡号密文%Type,
        strSQL = strSQL & "" & "NULL" & ","
        '  是否重复使用_In In 医疗卡类别.是否重复使用%Type,
        strSQL = strSQL & "" & 1 & ","
        '密码长度_In     In 医疗卡类别.密码长度%Type,
        strSQL = strSQL & "" & 10 & ","
        '密码长度限制_In In 医疗卡类别.密码长度限制%Type,
        strSQL = strSQL & "" & 0 & ","
        '密码规则_In     In 医疗卡类别.密码规则%Type,
        strSQL = strSQL & "" & 0 & ","
        strSQL = strSQL & "" & 1 & ","
        '  操作方式_In     In Integer := 0
        strSQL = strSQL & "" & intOper & ","
        '是否模糊查找_In     In 医疗卡类别.是否模糊查找%Type:=0
        strSQL = strSQL & "" & 0 & ","
        '问题号:51072
        '密码输入限制_In     In 医疗卡类别.密码输入限制%Type:=0
        strSQL = strSQL & "" & 0 & ","
        '是否缺省密码_In     In 医疗卡类别.是否缺省密码%Type:=0
        strSQL = strSQL & "" & 0 & ","
        '问题号:56508
        '是否制卡_In
        strSQL = strSQL & "" & 0 & ","
        '是否发卡_In
        strSQL = strSQL & "" & 0 & ","
        '是否写卡_In
        strSQL = strSQL & "" & 0 & ","
        '问题号:57697
        '险类_In
        strSQL = strSQL & "" & 0 & ","
        '问题号:57326
        strSQL = strSQL & "" & 1 & ","
        '77872,李南春,2014/12/3:是否支持转帐及代扣
        '是否转帐及代扣_In  In 医疗卡类别.是否转帐及代扣%Type:=0
        strSQL = strSQL & "" & 0 & ","
        '读卡性质_In       In 医疗卡类别.读卡性质%Type := '1000',
        strSQL = strSQL & "" & "1000" & ","
        '键盘控制方式_In   In 医疗卡类别.键盘控制方式%Type := 0,
        strSQL = strSQL & "" & 0 & ","
        '90875:李南春,2015/12/16,增加医疗卡证件类型
        '是否证件_In  In 医疗卡类别.是否证件%Type:=0
        strSQL = strSQL & "" & 1 & ")"
        
        zlAddArray colPro, strSQL
End Sub
 
Private Function IsCheckCancelValied(ByVal lng挂号结帐ID As Long, ByVal lng卡费结帐ID As Long, _
    ByVal cllBillBalance As Collection, ByVal dbl金额 As Double, Optional ByVal bln退款验卡 As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查退费时的数据有效性
    '返回:有效,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-12 18:43:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strName As String, bln消费卡 As Boolean, lng卡类别ID As Long
    Dim str验证卡号  As String, strXmlIn As String, str刷卡密码 As String
    Dim str卡号 As String, str交易流水号 As String, str交易说明 As String, str结算信息 As String
    Dim strXMLExpend As String
    Dim cllSquareBalance As Collection
    
    On Error GoTo errHandle
    strName = IIf(glngSys \ 100 = 8, "会员卡", "医疗卡")
    If cllBillBalance Is Nothing Then IsCheckCancelValied = True: Exit Function
    '问题号:58567
    'Array(卡类别ID,卡号,是否消费卡(1-是;0-否),交易流水号,交易说明,strNO,结帐ID,消费卡ID)
    lng卡类别ID = cllBillBalance(1)(0)
    If lng卡类别ID = 0 Then IsCheckCancelValied = True: Exit Function
    
    str卡号 = cllBillBalance(1)(1)
    bln消费卡 = Val(cllBillBalance(1)(2)) = 1
    str交易流水号 = cllBillBalance(1)(3)
    str交易说明 = cllBillBalance(1)(4)
    
    Set cllSquareBalance = New Collection
    'Array(卡类别ID,消费卡ID,刷卡金额, 卡号,密码,限制类别,是否密文,剩余未退金额)
    cllSquareBalance.Add Array(lng卡类别ID, cllBillBalance(1)(7), 0, str卡号, "", "", False, dbl金额)
    
    If gobjSquare Is Nothing Then
        Call InitCardSquareData
    End If
    '4.3.3.2.6   zlReturnCheck:帐户回退交易前的检查
    'zlPaymentCheck帐户扣款交易检查
    '参数名  参数类型    入/出   备注
    'frmMain Object  In  调用的主窗体
    'lngModule   Long    In  模块号
    'lngCardTypeID   Long    In  卡类别ID:医疗卡类别.ID
    'strCardNo   String  IN  卡号
    'strBalanceIDs:格式:收费类型( 1-预交款,2-结帐,3-收费,4-挂号,5-医疗卡收款)|ID1,ID2…IDn||收费类型n|ID1,ID2…IDn
    'dblMoney    Double  IN  退款金额
    'strSwapNo   String  In  交易流水号(退款时检查)
    'strSwapMemo String  In  交易说明(退款时传入)
    '    Boolean 函数返回    True:调用成功,False:调用失败
    '说明:
    '在调用扣款前，由于存在Oracle事务问题，因此，再调用回退交易前，先进行数据的合法性检查,以便控制死锁情况。
    If lng卡费结帐ID <> 0 Then str结算信息 = str结算信息 & "||5|" & lng卡费结帐ID
    If lng挂号结帐ID <> 0 Then str结算信息 = str结算信息 & "||4|" & lng挂号结帐ID
    If str结算信息 <> "" Then str结算信息 = Mid(str结算信息, 3)
    
    If gobjSquare.objSquareCard.zlReturnCheck(Me, mlngModul, lng卡类别ID, bln消费卡, str卡号, str结算信息, dbl金额, str交易流水号, str交易说明, strXMLExpend) = False Then
        Exit Function
    End If
    
    If bln消费卡 And gbln消费卡退费验卡 _
        Or bln消费卡 = False And bln退款验卡 Then
        strXmlIn = "<IN><CZLX>2</CZLX></IN>"
        If gobjSquare.objSquareCard.zlBrushCard(Me, mlngModul, Nothing, lng卡类别ID, bln消费卡, _
            txtPatient.Text, NeedName(cbo性别.Text), txt年龄.Text & (IIf(cbo年龄单位.Visible, cbo年龄单位.Text, "")), dbl金额, str卡号, str刷卡密码, _
            True, True, False, False, cllSquareBalance, False, True, strXmlIn) = False Then Exit Function
    End If
    
    IsCheckCancelValied = True
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Private Function CallBackBalanceInterface(ByVal cllBalance As Collection, _
    ByVal lng挂号结帐ID As Long, ByVal lng卡费结帐ID As Long, _
    ByVal dblMoney As Double, _
    ByRef cllUpdate As Collection, _
    ByRef cllThreeSwap As Collection, ByRef strErrMsg As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:调用回退接口
    '入参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-13 10:33:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str卡号 As String, strSwapGlideNO As String, strSwapMemo As String, str结算信息 As String, strSwapExtendInfor As String
    Dim varData As Variant, varTemp As Variant, i As Long, cllPro As Collection
    Dim bln消费卡 As Boolean, lng卡类别ID As Long, strSQL As String, rsTemp As ADODB.Recordset
    Dim lng挂号冲销ID As Long, lng退卡冲销ID As Long, strTemp As String
    
    Err = 0: On Error GoTo Errhand:
    '卡类别ID,卡号,是否消费卡(1-是;0-否),交易流水号,交易说明,strNO
    'cllBalance.Add Array(Val(Nvl(rsTmp!卡类别ID)), Trim(Nvl(rsTmp!卡号)), IIf(Val(Nvl(rsTmp!结算卡序号)) <> 0, 1, 0), Trim(Nvl(rsTmp!交易流水号)), Trim(Nvl(rsTmp!交易说明))), strNO
    If cllBalance Is Nothing Then CallBackBalanceInterface = True: Exit Function
    '问题号:58567
    bln消费卡 = Val(cllBalance(1)(2)) = 1
    lng卡类别ID = cllBalance(1)(0)
    
    '卡类别ID,卡号,是否消费卡(1-是;0-否),交易流水号,交易说明,strNO
    If lng卡类别ID = 0 Then CallBackBalanceInterface = True: Exit Function
    
    str卡号 = cllBalance(1)(1)
    strSwapGlideNO = cllBalance(1)(3)
    strSwapMemo = cllBalance(1)(4)
    If lng卡费结帐ID <> 0 Then str结算信息 = str结算信息 & "||5|" & lng卡费结帐ID
    If lng挂号结帐ID <> 0 Then str结算信息 = str结算信息 & "||4|" & lng挂号结帐ID
    If str结算信息 <> "" Then str结算信息 = Mid(str结算信息, 3)
    
    
    If lng卡费结帐ID <> 0 Then
        strSQL = " Select 结帐ID,记帐费用 From 住院费用记录  Where 记录性质=5 And NO =(Select Max(NO) From 住院费用记录 where 结帐ID=[1] and  记录性质=5  )  and 记录状态=2"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng卡费结帐ID)
        If rsTemp.EOF Then
            strErrMsg = "未找到退卡信息，不能继续": Exit Function
        End If
        lng退卡冲销ID = Val(Nvl(rsTemp!结帐ID))
    End If
    
    If lng挂号结帐ID <> 0 Then
        strSQL = "Select 结帐ID From 门诊费用记录  Where 记录性质=4 And NO =(Select Max(NO) From 门诊费用记录 where 结帐ID=[1] and  记录性质=4  )  and 记录状态=2"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng挂号结帐ID)
        If rsTemp.EOF Then
            strErrMsg = "未找到退号信息，不能继续": Exit Function
        End If
        lng挂号冲销ID = Val(Nvl(rsTemp!结帐ID))
    End If

    '81489,冉俊明,2015-1-22,退费传入冲销ID
    If lng退卡冲销ID <> 0 Then strSwapExtendInfor = strSwapExtendInfor & "||5|" & lng退卡冲销ID
    If lng挂号冲销ID <> 0 Then strSwapExtendInfor = strSwapExtendInfor & "||4|" & lng挂号冲销ID
    If strSwapExtendInfor <> "" Then strSwapExtendInfor = Mid(strSwapExtendInfor, 3)
    strTemp = strSwapExtendInfor
    'zlReturnMoney(frmMain As Object, ByVal lngModule As Long, _
        ByVal lngCardTypeID As Long, ByVal strCardNo As String, ByVal strBalanceIDs As String, _
        ByVal dblMoney As Double, _
        ByRef strSwapGlideNO As String, ByRef strSwapMemo As String, _
        ByRef strSwapExtendInfor As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:帐户扣款回退交易
    '入参:frmMain-调用的主窗体
    '       lngModule-调用的模块号
    '       lngCardTypeID-卡类别ID:医疗卡类别.ID
    '       strCardNo-卡号
    '       strBalanceIDs-本次支付所涉及的结算ID(这是原结帐ID):
    '                           格式:收费类型(|ID1,ID2…IDn||收费类型n|ID1,ID2…IDn
    '                           收费类型:1-预交款,2-结帐,3-收费,4-挂号,5-医疗卡收款
    '       dblMoney-退款金额
    '       strSwapNo-交易流水号(扣款时的交易流水号)
    '       strSwapMemo-交易说明(扣款时的交易说明)
    '       strSwapExtendInfor-传入，本次退费的冲销ID：
    '                           格式:收费类型1|ID1,ID2…IDn||收费类型n|ID1,ID2…IDn
    '                           收费类型:1-预交款,2-结帐,3-收费,4-挂号,5-医疗卡收款
    '       strSwapExtendInfor-传出，交易的扩展信息
    '           格式为:项目名称1|项目内容2||…||项目名称n|项目内容n 每个项目中不能包含|字符
    If gobjSquare.objSquareCard.zlReturnMoney(Me, mlngModul, lng卡类别ID, bln消费卡, str卡号, str结算信息, dblMoney, strSwapGlideNO, strSwapMemo, strSwapExtendInfor) = False Then Exit Function
    Set cllUpdate = New Collection: Set cllThreeSwap = New Collection
    If lng退卡冲销ID <> 0 Then
        '问题号:58536
        If Not bln消费卡 Then
            Call zlAddUpdateSwapSQL(False, lng退卡冲销ID, lng卡类别ID, bln消费卡, str卡号, strSwapGlideNO, strSwapMemo, cllUpdate)
        End If
        If strTemp <> strSwapExtendInfor Then
            Call zlAddThreeSwapSQLToCollection(False, lng退卡冲销ID, lng卡类别ID, bln消费卡, str卡号, strSwapExtendInfor, cllThreeSwap)
        End If
    End If
    If lng挂号冲销ID <> 0 Then
        Call zlAddUpdateSwapSQL(False, lng挂号冲销ID, lng卡类别ID, bln消费卡, str卡号, strSwapGlideNO, strSwapMemo, cllUpdate)
        If strTemp <> strSwapExtendInfor Then
            Call zlAddThreeSwapSQLToCollection(False, lng挂号冲销ID, lng卡类别ID, bln消费卡, str卡号, strSwapExtendInfor, cllThreeSwap)
        End If
    End If
    CallBackBalanceInterface = True
    Exit Function
Errhand:
    If ErrCenter() = 1 Then Resume
    SaveErrLog
End Function
Private Function IsValiedMzNo(ByVal lng病人ID As Long, ByRef str门诊号 As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查门诊号
    '入参:str门诊号-门诊号
    '出参:str门诊号-返回新的门诊号
    '返回:合法,返回true,否则返回false
    '编制:刘兴洪
    '日期:2011-10-31 10:22:12
    '问题:42616
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str门诊号1 As String, strNew门诊号 As String
    str门诊号1 = str门诊号
    If mTy_Para.bln预约不产生门诊号 And mbytMode = 1 Then IsValiedMzNo = True: Exit Function
    
    If str门诊号 = "" And mbln建病案 Then
        Call MsgBox("未输入门诊号,不能继续!", vbInformation + vbOKOnly, gstrSysName)
        If txt门诊号.Enabled Then txt门诊号.SetFocus
        Exit Function
    End If
    
    If Not Exist门诊号(str门诊号, lng病人ID) Then IsValiedMzNo = True: Exit Function
    '42638
    If Not (gbln自动门诊号 Or mblnStation) Then
        If str门诊号 <> "" Then
            Call MsgBox("当前门诊号:" & str门诊号1 & " 已经被其他病人使用,不能继续!", vbInformation + vbOKOnly, gstrSysName)
            If txt门诊号.Enabled Then txt门诊号.SetFocus
            Exit Function
        End If
    End If
    
    
    '重新获取门诊号
GoGetMzNo:
    strNew门诊号 = zlGet门诊号
    If Len(strNew门诊号) > txt门诊号.MaxLength Then
           MsgBox "当前门诊号已经被其它病人使用,系统自动更换门诊号为:" & strNew门诊号 & _
               vbCrLf & "但超过了允许的最大门诊号长度:" & txt门诊号.MaxLength & "位,请输入一个门诊号!", vbInformation, gstrSysName
           If txt门诊号.Enabled Then txt门诊号.SetFocus
           Exit Function
    End If
    If strNew门诊号 <> "" Then
        If Exist门诊号(strNew门诊号, lng病人ID) Then GoTo GoGetMzNo:
        '问题:42616自动生成门诊号,不提醒,直接保存
        If gbln自动门诊号 Then
            str门诊号 = strNew门诊号: IsValiedMzNo = True: Exit Function
        End If
        '需要提醒
        If MsgBox("当前门诊号:" & str门诊号1 & " 已经被其他病人使用," & IIf(strNew门诊号 <> "", vbCrLf & "  系统自动更换为" & strNew门诊号, "") & " 是否继续?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            txt门诊号.Text = strNew门诊号
            If txt门诊号.Enabled Then txt门诊号.SetFocus
            Exit Function
        End If
        '可能在用户操作时,因并发原因,再次被他人使用,因此还要检查门诊号是否被其他人使用
        If Exist门诊号(strNew门诊号, lng病人ID) Then
            If Not (gbln自动门诊号 Or mblnStation) Then
                Call MsgBox("当前门诊号:" & str门诊号 & " 已经被其他病人使用,不能继续!", vbInformation + vbOKOnly, gstrSysName)
                txt门诊号.Text = strNew门诊号
                If txt门诊号.Enabled Then txt门诊号.SetFocus
                Exit Function
            End If
            GoTo GoGetMzNo:
        End If
    End If
    str门诊号 = strNew门诊号
    txt门诊号.Text = str门诊号
    If str门诊号 = "" And mbln建病案 Then
         Call MsgBox("未输入门诊号,不能继续!", vbInformation + vbOKOnly, gstrSysName)
         If txt门诊号.Enabled Then txt门诊号.SetFocus
         Exit Function
     End If
     IsValiedMzNo = True
End Function

Private Sub ReInitPatiInvoice(Optional blnFact As Boolean = True, _
    Optional ByVal intInsure_IN As Integer = 0, Optional ByVal lng病人ID_In As Long = 0)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重新初始化病人发票信息
    '入参:blnFact-是否重新取发票号
    '编制:刘兴洪
    '日期:2011-04-29 14:17:33
    '问题:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strInvoiceFormat As String, lng病人ID As Long
    Dim intInsure As Integer, strUseType As String
    If mblnStartFactUseType = False Then Exit Sub
    
    lng病人ID = lng病人ID_In
    
    If lng病人ID_In = 0 Then
        If Not mrsInfo Is Nothing Then
            If mrsInfo.State = 1 Then lng病人ID = mrsInfo!病人ID
        End If
    End If
    
    If mblnStationPrice Then
        Exit Sub
    End If
    
    intInsure = IIf(intInsure_IN <> 0, intInsure_IN, mintInsure)
    strUseType = mstrUseType
    mstrUseType = "": mlngShareUseID = 0: mintInvoiceFormat = 0
    mstrUseType = zl_GetInvoiceUserType(lng病人ID, 0, intInsure)
    '切换了票据类型
    If mstrUseType <> strUseType And mblnStartFactUseType Then mlng领用ID = 0
    mlngShareUseID = zl_GetInvoiceShareID(mlngModul, mstrUseType)
    mintInvoiceFormat = zl_GetInvoicePrintFormat(mlngModul, mstrUseType)
    mintInvoicePrint = zl_GetInvoicePrintMode(mlngModul, mstrUseType)
    
    'Call ShowBillFormat
    If blnFact Then Call RefreshFact
End Sub

Private Function GetActiveView()
    '******************************************************************************
    '   得到当前挂号业务  采取那种类型的流程
    '******************************************************************************
        Dim strSQL          As String
        Dim rsTmp           As ADODB.Recordset
        Dim lng记录ID       As Long
        
        On Error GoTo Hd
        lng记录ID = Val(vsfPlan.TextMatrix(vsfPlan.Row, GetCol("记录ID")))
        
        strSQL = "Select 1　From 临床出诊记录 Where ID=[1] And Nvl(是否分时段,0)=1 "
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng记录ID)
         If rsTmp.RecordCount > 0 And vsfPlan.TextMatrix(vsfPlan.Row, GetCol("序号控制")) <> "" Then
            '*********************
            '专家号分时段
            '*********************
            mViewMode = v_专家号分时段
        '78640:李南春,2014/10/16,挂号处预约显示所有可预约的号别
         ElseIf rsTmp.RecordCount > 0 And vsfPlan.TextMatrix(vsfPlan.Row, GetCol("序号控制")) = "" And (mbytMode = 1 Or (mbytMode = 0 And chkBooking.Visible And chkBooking.Value = 1)) Then
            '*********************
            '普通号分时段
            '*********************
            mViewMode = V_普通号分时段
         ElseIf vsfPlan.TextMatrix(vsfPlan.Row, GetCol("序号控制")) <> "" Then
            '*********************
            '专家号不分时段
            '*********************
            mViewMode = v_专家号
            vsfList.Visible = True
            picSplit.Visible = True
          Else
            '*********************
            '普通号
            '*********************
            mViewMode = V_普通号
            vsfList.Visible = False
            picSplit.Visible = False
         End If
        
        Set rsTmp = Nothing
Exit Function
Hd:
    If ErrCenter() = 1 Then
         Resume
    End If
    SaveErrLog
    
End Function

Private Function InitTimePlan() As Boolean
    '**************************************
    '加载时段
    '返回时段是否加载成功或是否有分时段
    '**************************************
    Dim strSQL         As String
    Dim dateCur        As Date
    Dim lng记录ID      As Long
    On Error GoTo Hd
     '挂号分时段
    strSQL = "Select 记录id, 序号, To_Char(开始时间, 'hh24') || ':00' As 时间点, To_Char(开始时间, 'hh24:mi') As 开始时间," & vbNewLine & _
            "       To_Char(终止时间, 'hh24:mi') As 结束时间, 数量 As 限制数量, 是否预约, 预约顺序号, 开始时间 As 详细开始时间, 终止时间 As 详细结束时间 " & vbNewLine & _
            "From 临床出诊序号控制" & vbNewLine & _
            "Where 记录id = [1] And 开始时间 <> 终止时间 " & vbNewLine & _
            "Order By 详细开始时间"
    lng记录ID = Val(vsfPlan.TextMatrix(vsfPlan.Row, GetCol("记录ID")))
    Set mrs时间段 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng记录ID)
    If mrs时间段.EOF Then Exit Function
    
    InitTimePlan = True
    Exit Function
Hd:
    If ErrCenter() = 1 Then Resume
    SaveErrLog
End Function

Private Function Check有效号别(ByVal str号别 As String, ByVal datThis As Date, Optional ByVal blnPlan As Boolean = False) As Boolean
   '***********************************************************
   '对挂号或者预约时
   '输入有效时间的验证
   '***********************************************************
    Dim strSQL          As String
    Dim rsTmp           As ADODB.Recordset
    Dim rs时间段        As ADODB.Recordset
    Dim str安排         As String
    Dim dat开始时间     As Date
    Dim dat结束时间     As Date
    Dim blnOK           As Boolean
    Dim str时间()       As String
    Dim i               As Long
    Dim Datsys          As Date
    
    '******************************
    '挂号检查时 在分时段的情况下
    '只在挂号下检查 因为 预约已限制
    '发生时间不能小于 时段的开始时间
    '******************************
     On Error GoTo Hd
    If blnPlan = False And mbytMode = 0 And mViewMode = v_专家号分时段 Then
        Datsys = zlDatabase.Currentdate
        If datThis <= Datsys Then
            MsgBox "时段的开始时间" & Format(datThis, "HH:MM") & "小于了当前时间" & Format(Datsys, "hh:MM") & "!请检查", vbOKOnly, Me.Caption
            Exit Function
        End If
    End If
    If blnPlan Then
        Datsys = zlDatabase.Currentdate
        If datThis <= Datsys Then
            MsgBox "预约时间" & Format(datThis, "yyyy-mm-DD HH:MM") & "小于了当前时间" & Format(Datsys, "yyyy-mm-DD hh:MM") & "!请检查", vbOKOnly, Me.Caption
            Exit Function
        End If
    End If
 
   Check有效号别 = True
    Exit Function
Hd:
    If ErrCenter() = 1 Then
        Resume
    End If
    SaveErrLog
End Function

Private Sub InitActionType()
    '-------------------------
    '获取 是否采用了分时段的处理方式
    '判断依据为 挂号安排列表是否有数据
    '-------------------------
    Dim strSQL       As String
    Dim rsTmp        As ADODB.Recordset
    strSQL = _
    "    Select 1  dt From  临床出诊记录 Where 是否分时段 = 1 And Rownum< 2"
    
    On Error GoTo Hd
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    mcustomTime = t_普通
    If rsTmp.RecordCount <> 0 Then mcustomTime = t_时段
    Select Case mcustomTime
    Case t_普通:
        Me.dtpAppointmentDate.CustomFormat = "yyyy-MM-dd"
        Me.dtpAppointmentTime.CustomFormat = "HH:mm"
        Form_Resize
    Case t_时段:
        Me.dtpAppointmentDate.CustomFormat = "yyyy-MM-dd"
        Me.dtpAppointmentTime.CustomFormat = "HH:mm"
        Form_Resize
    End Select
    
    Set rsTmp = Nothing
    Exit Sub
Hd:
    If ErrCenter() = 1 Then
        Resume
    End If
    SaveErrLog
End Sub

Private Sub MBox(ByVal strMsg As String, Optional ByVal strTitle As String = "")
    '------------------------------------------------
    '消息框
    '------------------------------------------------
    If strTitle = "" Then strTitle = Me.Caption
    MsgBox strMsg, vbInformation, strTitle
End Sub

Private Function SetBrushCard(ByVal objContrl As Object, KeyAscii As Integer) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:刷卡操作
    '入参:
    '出参:
    '返回:刷卡读取的,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-11-10 10:01:51
    '问题:38603
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Static sngBegin As Single
    Dim sngNow As Single, blnCard As Boolean, lng医疗卡长度 As Long
    If Not (mblnStation And mTy_Para.bln挂号必须刷卡 And mbytMode = 0) Then Exit Function
    lng医疗卡长度 = IDKind.GetCardNoLen
    objContrl.PasswordChar = IIf(IDKind.ShowPassText, "*", "")
    '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
    objContrl.IMEMode = 0
    
    '是否刷卡完成
    blnCard = KeyAscii <> 8 And Len(objContrl.Text) = lng医疗卡长度 - 1 And objContrl.SelLength <> Len(objContrl.Text)
    If blnCard Or KeyAscii = 13 Then
        If KeyAscii <> 13 Then
            objContrl.Text = objContrl.Text & Chr(KeyAscii)
            objContrl.SelStart = Len(objContrl.Text)
        End If
        KeyAscii = 0
        mblnCard = True
        Call txtPatient_Validate(True)
        mblnCard = False
        '刘兴洪:27494  20100117
        If Replace(txtPatient.Text, vbCrLf, "") = "" Then
            DoEvents: txtPatient.SetFocus
        End If
    Else
        If InStr(":：;；?？" & Chr(22), Chr(KeyAscii)) > 0 Then
            KeyAscii = 0 '去除特殊符号，并且不允许粘贴
        Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        End If
        '安全刷卡检测
        If KeyAscii <> 0 And KeyAscii > 32 Then
            sngNow = Timer
            If objContrl.Text = "" Then
                sngBegin = sngNow
            ElseIf Format((sngNow - sngBegin) / (Len(objContrl.Text) + 1), "0.000") >= 0.04 Then '>0.007>=0.01
                objContrl.Text = Chr(KeyAscii)
                objContrl.SelStart = 1
                KeyAscii = 0
                sngBegin = sngNow
            End If
        End If
    End If
    SetBrushCard = True
End Function
Private Sub CreateMobjIDCard()
'创建IDCard
    '弹出小窗口中的mobjIDCard和本页面的mobjIDCard冲突
    '导致 不能重新刷 身份证 原因未找到
    If (mbytMode = 0 Or mbytMode = 1) And mbytInState = 0 Then
        Set mobjIDCard = New clsIDCard
        Call mobjIDCard.SetParent(Me.Hwnd)
        If Me.ActiveControl Is Me.txtPatient And Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (Me.txtPatient.Text = "")
    End If
End Sub

Public Function Get失约号(ByVal lng记录ID As Long, datThis As Date) As Long
   '获取安排在某一天.预约失约数
    Dim strSQL  As String
    Dim rsTmp   As ADODB.Recordset
    Dim strBegin  As String, strEnd As String
    If mTy_Para.bln失约用于挂号 = False Or mTy_Para.lng预约有效时间 = 0 Then Exit Function
    strSQL = "Select Count(1) As 失约号" & vbNewLine & _
            " From 病人挂号记录" & vbNewLine & _
            " Where 出诊记录ID = [1] And 记录性质 = 2 And 记录状态 = 1 And 发生时间 - [2] / 24 / 60 < Sysdate And 发生时间 Between to_Date([3],'YYYY-MM-DD') And to_Date([4],'YYYY-MM-DD') - 1/24/60/60"
    strBegin = Format(datThis, "yyyy-MM-dd")
    strEnd = Format(datThis + 1, "yyyy-MM-dd")
    On Error GoTo Hd
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng记录ID, mTy_Para.lng预约有效时间, strBegin, strEnd)
    If rsTmp.EOF Then
        Get失约号 = 0
        Set rsTmp = Nothing
        Exit Function
    End If
    Get失约号 = Val(Nvl(rsTmp!失约号, 0))
    Set rsTmp = Nothing
   Exit Function
Hd:
    If ErrCenter() = 1 Then
        Resume
    End If
    SaveErrLog
End Function

Public Sub zl_StationInitPatient(ByVal lng病人ID As Long)
    '功能说明:门诊工作站调用时初始化病人信息
    '参数说明:str门诊号
    If mTy_Para.bln挂号必须刷卡 Or mblnStation = False Or lng病人ID = 0 Then Exit Sub
    txtPatient.Text = "-" & lng病人ID
    txtPatient_Validate False
End Sub
Private Sub AddDeposit()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:缴预存款
 
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objFun          As Object
    Dim lng病人ID       As Long
    Err = 0: On Error Resume Next
    Set objFun = CreateObject("zl9Patient.clsPatient")
    If Err <> 0 Then Exit Sub
    
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能： 调用预交款收款窗口
    '参数：
    '   lngModul:需要执行的功能序号
    '   cnMain:主程序的数据库连接
    '   frmMain:主窗体
    '   strDBUser:当前数据库登录用户名
    '  bytCallObject:刘兴洪加入(0-预交款调用(缺省的);1-病人费用查询调用,2-医疗卡调用)
    '  lng病人ID-缺省的病人ID
    '  lng主页ID-缺省的主页ID
    '  dblDefPrePayMoney-缺省的预付金额
    If Not mrsInfo Is Nothing Then lng病人ID = Val(Nvl(mrsInfo!病人ID))
    If objFun.PlusDeposit(glngSys, gcnOracle, Me, gstrDBUser, 2, lng病人ID, 0, 0, 0) = False Then
        Set objFun = Nothing
        Exit Sub
    End If
    Set objFun = Nothing
    If lng病人ID <> 0 Then
        txtPatient.Text = "-" & lng病人ID
        mblnOnVilidate = True
        Call txtPatient_Validate(False)
        mblnOnVilidate = False
    End If
End Sub

Private Sub InitTimeSect(ByVal lng记录ID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化时间段
    '编制:刘兴洪
    '日期:2012-03-12 15:45:57
    '问题:45509
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    On Error GoTo errHandle
    strSQL = "Select 开始时间,终止时间,缺省预约时间 As 缺省时间  From 临床出诊记录 Where ID=[1]"

    Set mrsALL时间段 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng记录ID)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub SetDefaultRegistTime()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置缺省的预约时间
    '编制:刘兴洪
    '日期:2012-03-12 15:49:38
    '问题:45509
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str安排 As String, str时间 As String
    Dim dtValue As Date, str号码 As String
    Dim str缺省时间 As String, Datsys As Date
    Static str上次号码 As String
    On Error GoTo errHandle
    If mblnAppointmentChange Then Exit Sub
    Datsys = zlDatabase.Currentdate
    With vsfPlan
        If .Row < 0 Or .Row > .Rows - 1 Then Exit Sub
        If .ColIndex(mstrCurKey) < 0 Then Exit Sub
       str安排 = .Cell(flexcpData, .Row, .ColIndex(mstrCurKey))
       str号码 = .TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("记录ID"))
    End With
    
    Call InitTimeSect(Val(str号码))
    If mrsALL时间段.EOF Then
        dtpAppointmentTime.Value = Format(Datsys, "HH:MM:SS")
        If dtpAppointmentDate.Visible Then
            txt发生时间.Text = Format(dtpAppointmentDate.Value, "yyyy-MM-dd") & " " & Format(dtpAppointmentTime.Value, "hh:mm:ss")
        Else
            txt发生时间.Text = Format(Datsys, "yyyy-MM-dd") & " " & Format(dtpAppointmentTime.Value, "hh:mm:ss")
        End If
        str上次号码 = str号码
        Exit Sub
    Else
        str上次号码 = str号码
    End If
    
    If (mViewMode = V_普通号分时段 Or mViewMode = v_专家号分时段) Then
        If mbytMode = 1 Or chkBooking.Value = 1 Then
            txt发生时间.Text = Format(Format(dtpAppointmentDate.Value, "yyyy-mm-dd" & " " & Format(dtpAppointmentTime.Value, "hh:mm:ss")), "yyyy-mm-dd hh:mm:ss")
        Else
            txt发生时间.Text = Format(Format(Datsys, "yyyy-mm-dd" & " " & Format(dtpAppointmentTime.Value, "hh:mm:ss")), "yyyy-mm-dd hh:mm:ss")
        End If
        str上次号码 = str号码
        Exit Sub
    Else
        If dtpAppointmentDate.Visible Then
            If Format(Nvl(mrsALL时间段!缺省时间, dtpAppointmentDate.Value), "yyyy-MM-dd") <> Format(dtpAppointmentDate.Value, "yyyy-MM-dd") Then
                If Format(Nvl(mrsALL时间段!缺省时间, dtpAppointmentDate.Value), "yyyy-MM-dd") > Format(dtpAppointmentDate.Value, "yyyy-MM-dd") Then
                    dtpAppointmentTime.Value = Format(Nvl(mrsALL时间段!开始时间, Datsys), "HH:MM:SS")
                Else
                    dtpAppointmentTime.Value = Format(Nvl(mrsALL时间段!终止时间, Datsys), "HH:MM:SS")
                End If
            Else
                dtpAppointmentTime.Value = Format(Nvl(mrsALL时间段!缺省时间, Datsys), "HH:MM:SS")
            End If
            txt发生时间.Text = Format(dtpAppointmentDate.Value, "yyyy-MM-dd") & " " & Format(dtpAppointmentTime.Value, "hh:mm:00")
        Else
            If Format(Nvl(mrsALL时间段!缺省时间, Datsys), "hh:mm:ss") < Format(Datsys, "hh:mm:ss") Then
                dtpAppointmentTime.Value = Format(Datsys, "HH:MM:SS")
            Else
                dtpAppointmentTime.Value = Format(Nvl(mrsALL时间段!缺省时间, Datsys), "hh:mm:ss")
            End If
            txt发生时间.Text = Format(Datsys, "yyyy-MM-dd") & " " & Format(dtpAppointmentTime.Value, "hh:mm:00")
        End If
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Function CancelBill(ByVal frmMain As Object, _
    ByVal strNoIn As String, ByVal lngModul As Long, ByVal strPrivs As String, Optional ByVal blnCenter As Boolean = False, _
    Optional ByVal intCancel As Integer = 0) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:退号操作(刘兴洪给李光福补上frmMain参数及功能说明
    '入参:frmMain-调用的主窗体
    '     intCancel-0=退号;1=退病历费;2=退附加费
    '返回:退费成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-10-23 17:09:50
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mstrNoIn = strNoIn:   mstrPrivs = strPrivs:    mlngModul = lngModul
    mbytMode = 4:    mbytInState = 1: mblnCenter = blnCenter
    mintCancel = intCancel
    mblnOk = False
    Me.Show 1, frmMain
    CancelBill = mblnOk
End Function

Public Function CancelApp(ByVal frmMain As Object, _
    ByVal strNoIn As String, ByVal lngModul As Long, ByVal strPrivs As String, Optional ByVal blnCenter As Boolean = False) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:取消预约操作
    '入参:frmMain-调用的主窗体
    '返回:退费成功返回true,否则返回False
    '编制:刘尔旋
    '日期:2016-04-12
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mstrNoIn = strNoIn:   mstrPrivs = strPrivs:    mlngModul = lngModul
    mbytMode = 3:    mbytInState = 1: mblnCenter = blnCenter
    mblnOk = False
    Me.Show 1, frmMain
    CancelApp = mblnOk
End Function

Private Function GetMaxLapseNO() As Long
    '功能说明:获取采用序号控制最大的无效号码是多少
    '返回值:
    Dim i As Long
    Dim j As Long
    Dim nStart As Long
    Dim lngResult As Long
    Dim lngTmp As Long
    If mViewMode = V_普通号 Or mViewMode = V_普通号分时段 Then Exit Function
    nStart = IIf(mViewMode = v_专家号, 0, 1)
    With vsfList
        For i = 0 To .Rows - 1
            For j = nStart To .Cols - 1
                If Trim(.TextMatrix(i, j)) <> "" Then
                     If .Cell(flexcpForeColor, i, j) <> vbRed And .Cell(flexcpForeColor, i, j) <> vbBlue And .Cell(flexcpForeColor, i, j) <> vbGrayText And .Cell(flexcpForeColor, i, j) <> &HC000C0 Then ' And .Cell(flexcpForeColor, i, j) <> vbGreen then
                         '空出来 暂时不做处理 方便以后添加功能
                        If Not mrsSNState Is Nothing And .Cell(flexcpForeColor, i, j) <> vbGreen Then
                            lngTmp = Val(Get时段(i, j, False))
                            mrsSNState.Filter = "序号=" & lngTmp
                            If mrsSNState.RecordCount > 0 Then
                                GetMaxLapseNO = lngTmp
                            End If
                        End If
                         
                     Else
                        If mViewMode = v_专家号分时段 Then
                            If .Cell(flexcpForeColor, i, j) = &HC000C0 And mTy_Para.bln随机序号选择 = False Then
                                '如果不能随机序号选择,同时是预约接收,暂不处理
                            Else
                                
                                GetMaxLapseNO = Val(Get时段(i, j, False))
                            End If
                        Else
                            GetMaxLapseNO = Val(.TextMatrix(i, j))
                        End If
                     End If
                End If
            Next
        Next
    End With
End Function

'获取idkind的默认kind值
Private Function IDKindDefaultKind() As Long
    Dim lngIndex As Long
    'IDkind的默认Kind
    If IDKind.DefaultCardType = "" Then
        lngIndex = -1
    Else
        If IsNumeric(IDKind.DefaultCardType) Then
           lngIndex = IDKind.GetKindIndex(IDKind.GetfaultCard.名称)
        Else
           lngIndex = IDKind.GetKindIndex(IDKind.DefaultCardType)
        End If
    End If
    IDKindDefaultKind = lngIndex
End Function

'控件名称是否匹配
Private Function IsCardType(ByVal IDKindCtl As IDKindNew, ByVal strCardName As String) As Boolean
    If IDKindCtl Is Nothing Then Exit Function
    If UCase(TypeName(IDKindCtl)) <> "IDKINDNEW" Then Exit Function
    Select Case strCardName
     Case "姓名", "姓名或就诊卡"
          IsCardType = IDKindCtl.GetCurCard.名称 Like "姓名*"
     Case "身份证", "身份证号", "二代身份证"
          IsCardType = IDKindCtl.GetCurCard.名称 Like "*身份证*"
     Case "IC卡号", "IC卡"
          IsCardType = IDKindCtl.GetCurCard.名称 Like "IC卡*"
     Case "医保号"
          IsCardType = IDKindCtl.GetCurCard.名称 = "医保号"
     Case "门诊号"
          IsCardType = IDKindCtl.GetCurCard.名称 = "门诊号"
     Case Else
            If IDKindCtl.GetCurCard Is Nothing Then Exit Function
            If Not IsNumeric(strCardName) Or Val(strCardName) <= 0 Then Exit Function
            If IDKindCtl.GetCurCard.接口序号 <= 0 Then Exit Function
            IsCardType = IDKindCtl.GetCurCard.接口序号 = Val(strCardName)
     End Select
End Function

'Private Function SetCreateCardObject() As Boolean
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    '功能:设置制卡对象
'    '编制:王吉
'    '日期:2012-12-17 11:06:41
'    '问题号:56599
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    On Error GoTo Errhand:
'    If mobjHealthCard Is Nothing Then
'        Set mobjHealthCard = CreateObject("zl9Card_HealthCard.clsHealthCard")
'    End If
'    SetCreateCardObject = True
'    Exit Function
'Errhand:
'    SetCreateCardObject = False
'End Function

Private Function zlExistsTodaysAppointment(ByVal lngPatientID As Long) As Boolean
'检查病人在当日是否有预约单据
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim rsInfo As ADODB.Recordset
    Dim strOutNo As String
    Dim frmNew As frmSelRegist
    Dim blnExit As Boolean
    Dim strMsg As String

    If mbytInState = 1 Then Exit Function
    If InStr(1, mstrPrivs, ";接收预约;") = 0 Then Exit Function
    If Not (chkCancel.Value = 0 And chkPrint.Value = 0 And chkBooking.Value = 0 And Not mblnStation) Then Exit Function
    If mbytMode = 1 Or mbytMode = 2 Then Exit Function

    strSQL = "Select a.NO, a.病人id, a.姓名, a.号别, a.号序, a.发生时间, a.登记时间,b.名称 as 执行科室 " & vbNewLine & _
           "       From 病人挂号记录 a,部门表 b" & vbNewLine & _
           "       Where a.发生时间 Between Trunc(Sysdate) And Trunc(Sysdate) + 1 - 1 / 24 / 60 / 60 And a.记录性质 = 2 And a.记录状态 = 1 And a.病人ID=[1] And A.执行部门ID=B.ID"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngPatientID)
    If rsTmp.EOF Then Exit Function

    If rsTmp.RecordCount = 1 Then
        '只有一条挂号记录,提醒操作员是否接收本条预约单据
        strMsg = "病人[" & Nvl(rsTmp!姓名) & "]在今日在科室[" & Nvl(rsTmp!执行科室) & "]存在预约单据(" & Nvl(rsTmp!NO) & ")是否接收?"
        If MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton1, Me.Caption) = vbYes Then
            Call ReadBooking(rsTmp!NO)
        Else
            Exit Function    '不提取本条预约单据
        End If
    Else
        '只有一条挂号记录,提醒操作员是否接收本条预约单据
        strMsg = "病人[" & Nvl(rsTmp!姓名) & "]在今日预约了多张单据,是否需要接收?"
        If MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton1, Me.Caption) = vbYes Then

            Call CloseIDCard    '47007
            Set frmNew = New frmSelRegist
            If frmNew.ShowRegist(Me, mstrPrivs, mblnOlnyBJYB, mTy_Para.int预约失效次数, strOutNo, rsInfo, Val(Nvl(rsTmp!病人ID))) = False Then
                blnExit = True
            End If
            If Not frmNew Is Nothing Then Unload frmNew
            Set frmNew = Nothing
            Call NewCardObject
            If blnExit Then Exit Function
            Call ReadBooking(strOutNo)
        Else
            Exit Function    '不提取本条预约单据
        End If
    End If
    zlExistsTodaysAppointment = True
End Function
Private Sub SetDelBillCtlEnabled(Optional bln三方结算 As Boolean)
    '设置病人退号时,病历相关退费控件状态
    Dim blnEnabled As Boolean
    Dim blnNotEnabled As Boolean
    If Not (mbytInState = 1 And mbytMode = 4 Or chkCancel.Value = 1) Then Exit Sub
    If bln三方结算 Then chk病历费.Enabled = False: Exit Sub '三方结算.不能部分退,至少暂时不支持

    If mrsBill Is Nothing Then Exit Sub
    If mrsBillAdvance Is Nothing Then Exit Sub
    
    mrsBillAdvance.Filter = 0
    mrsBill.Filter = "附加标志=1"
    If mrsBill.RecordCount = 0 Then
        blnNotEnabled = blnNotEnabled Or True
    End If
    mrsBill.Filter = 0
    chk病历费.Enabled = Not blnNotEnabled And mintCancel = 0
End Sub
Private Sub InitInputMaxLen()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化输入的最大长度
    '编制:刘兴洪
    '日期:2013-11-11 11:28:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    txtPatient.MaxLength = zlGetPatiInforMaxLen.intPatiName
    txtPatientPrint.MaxLength = txtPatient.MaxLength
    txt年龄.MaxLength = zlGetPatiInforMaxLen.intPatiAge
    txt门诊号.MaxLength = zlGetPatiInforMaxLen.intPatiMzNo
End Sub

Private Function zlGetInvoiceGroupUseID(ByRef lng领用ID As Long, _
    Optional intNum As Integer = 1, Optional strInvoiceNO As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取票据的领用ID
    '入参:lng领用ID-领用id
    '       intNum-页数
    '       strInvoiceNO-输入的发票号
    '出参:lng领用ID-领用ID
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-11-19 16:32:57
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    lng领用ID = GetInvoiceGroupID(IIf(gblnSharedInvoice, 1, 4), intNum, lng领用ID, glng挂号ID, strInvoiceNO, IIf(mblnStartFactUseType, mstrUseType, ""))
    If lng领用ID <= 0 Then
        Select Case lng领用ID
            Case 0 '操作失败
            Case -1
                If Trim(mstrUseType) = "" Then
                    MsgBox "你没有自用和共用的收费票据,请先领用一批票据或设置本地共用票据！", vbInformation, gstrSysName
                Else
                    MsgBox "你没有自用和共用的『" & mstrUseType & "』收费票据,请先领用一批票据或设置本地共用票据！", vbInformation, gstrSysName
                End If
                Exit Function
            Case -2
                If Trim(mstrUseType) = "" Then
                    MsgBox "本地的共用票据已经用完,请先领用一批票据或重新设置本地共用票据！", vbInformation, gstrSysName
                Else
                    MsgBox "本地的共用票据的『" & mstrUseType & "』收费票据已经用完,请先领用一批票据或重新设置本地共用票据！", vbInformation, gstrSysName
                End If
                Exit Function
            Case -3
                MsgBox "当前票据号码不在可用领用批次的有效票据号范围内,请重新输入！", vbInformation, gstrSysName
                If txtFact.Enabled And txtFact.Visible Then txtFact.SetFocus
                Exit Function
        End Select
    End If
    zlGetInvoiceGroupUseID = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function zlIsAllowPatiChargeFeeMode(ByVal lng病人ID As Long, ByVal int原结算模式 As Integer) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:是否允许改变病人收费模式
    '入参:lng病人ID-病人ID
    '       int原结算模式-0表示先结算后诊疗;1表示先诊疗后结算
    '返回:允许调整收费模式,返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-12-25 10:06:49
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    Dim dtDate As Date, intDay As Integer
    On Error GoTo errHandle
    
    If mbytMode = 1 Then zlIsAllowPatiChargeFeeMode = True: Exit Function '预约不处理
    '模式未调整，直接返回true
    If int原结算模式 = mPatiChargeMode Then zlIsAllowPatiChargeFeeMode = True: Exit Function
    
      
    If int原结算模式 = 1 Then
        '原为先诊疗后结算且存在未结费用的,则必须采用记帐模式
        strSQL = "" & _
        "   Select 1 " & _
        "   From 病人未结费用 " & _
        "   Where 病人id = [1] And (来源途径 In (1, 4) Or 来源途径 = 3 And Nvl(主页id, 0) = 0) And Nvl(金额, 0) <> 0 And Rownum < 2"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID)
        If rsTemp.EOF = False Then
            MsgBox "注意:" & vbCrLf & "  当前病人的就诊模式为先诊疗后结算且存在未结费用，" & _
                                          vbCrLf & "不允许调整该病人的就诊模式,你可以先对未结费用结帐后" & _
                                          vbCrLf & "再挂号或不调整病人的就诊模式", vbOKOnly + vbInformation, gstrSysName
            Exit Function
        End If
        intDay = IIf(lbl急.Visible, -1 * gSysPara.Sy_Reg.bytNoDayseMergency, -1 * gSysPara.Sy_Reg.bytNODaysGeneral)
        dtDate = DateAdd("d", intDay, zlDatabase.Currentdate)
        ' 上次为"先诊疗后结算",本次为"先结算后诊疗"的,同时满足未发生医嘱业务数据的 ,
        '   则不允许更改就诊模式
        strSQL = "Select 1 " & _
        " From 病人挂号记录 A, 病人医嘱记录 B " & _
        " Where a.病人id + 0 = b.病人id And a.No || '' = b.挂号单  " & _
        "               And a.记录状态 = 1 And a.记录性质 = 1 And a.登记时间 - 0 >= [2] " & _
        "               And  a.病人id = [1] And rownum<2"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, dtDate)
        If rsTemp.EOF Then
            '未发生医嘱数据
            MsgBox "注意:" & vbCrLf & "  当前病人的就诊模式为先诊疗后结算," & vbCrLf & "  不允许调整该病人的就诊模式!", vbOKOnly + vbInformation, gstrSysName
            Exit Function
        End If
    End If
    zlIsAllowPatiChargeFeeMode = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
 End Function
 
 Public Sub SendMsgModule(ByVal strNO As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:消息发送处理
    '入参: strNO-挂号单号
    '编制:刘兴洪
    '日期:2014-03-11 11:59:22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim objXML As New clsXML
    
    '0-挂号,1-预约,2-接收,3-取消预约 ,4-退号 预约有两种模式:0-挂号,此时预约要收费,1-预约,不收费
    If Not (mbytMode = 0 Or mbytMode = 2) Or mbytInState = 1 Then Exit Sub
    If mobjMsgModule Is Nothing Then Exit Sub
    If mobjMsgModule.IsConnect = False Then Exit Sub

    strSQL = "" & _
    " Select A.id ,A.姓名,nvl(A.门诊号,B.门诊号) as 门诊号,A.病人Id,b.身份证号,A.NO,A.执行部门ID,C.名称 as 执行部门名称,A.诊室,A.执行人  " & _
    " From 病人挂号记录 A,病人信息 B,部门表 C  " & _
    " where A.No=[1] and a.记录状态 =1 And a.记录性质=1 and a.病人ID=b.病人id and a.执行部门id=c.id"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
    If rsTemp.EOF Then Exit Sub
    
    'ZLHIS_REGIST_001 门诊病人挂号通知
    '节点名称    属性    含义    重复    类型    缺省值  值域描述
    '<patient_info>
    '    <patient_id>病人ID</patient_id>
    '    <patient_name>病人姓名</patient_name>
    '    <identity_card>身份证号</identity_card>
    '    <out_number>门诊号</out_number>
    '</patient_info>
    '<register_info>
    '    <register_id>挂号id</register_id>
    '    <register_no>挂号单号</register_no>
    '    <register_dept_id>挂号科室id</register_dept_id>
    '    <register_dept_title>挂号科室</register_dept_title>
    '    <register_room>挂号诊室</register_room>
    '    <register_doctor>挂号医生</register_doctor>
    '</register_info>
    objXML.ClearXmlText
    Call objXML.AppendNode("patient_info")
        Call objXML.appendData("patient_id", Val(Nvl(rsTemp!病人ID)))
        Call objXML.appendData("patient_name", Nvl(rsTemp!姓名))
        Call objXML.appendData("identity_card", Nvl(rsTemp!身份证号))
        Call objXML.appendData("out_number", Nvl(rsTemp!门诊号))
    Call objXML.AppendNode("patient_info", True)
    
    Call objXML.AppendNode("register_info")
        Call objXML.appendData("register_id", Val(Nvl(rsTemp!ID)))
        Call objXML.appendData("register_no", strNO)
        Call objXML.appendData("register_dept_id", Val(Nvl(rsTemp!执行部门id)))
        Call objXML.appendData("register_dept_title", Nvl(rsTemp!执行部门名称))
        Call objXML.appendData("register_room", Nvl(rsTemp!诊室))
        Call objXML.appendData("register_doctor", Nvl(rsTemp!执行人))
    Call objXML.AppendNode("register_info", True)
    Call mobjMsgModule.CommitMessage("ZLHIS_REGIST_001", objXML.XmlText)
    objXML.ClearXmlText
 End Sub
 
 Private Function ShowPatiPic() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示病人照片
    '编制:冉俊明
    '日期:2014-7-8
    '---------------------------------------------------------------------------------------------------------------------------------------------
    picPatiPicBack.Visible = True
    Set imgPatiPic.Picture = mobjfrmPatiInfo.imgPatient.Picture
    lblShow.Visible = imgPatiPic.Picture = 0
 End Function

Private Sub LoadIDImage()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载身份证图像
    '编制:刘兴洪
    '日期:2014-06-30 16:20:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Dim objStdPic As StdPicture
    
    If mobjIDCard Is Nothing Then Exit Sub
    Call mobjIDCard.GetPhotoAsStdPicture(objStdPic)
    mobjfrmPatiInfo.imgPatient.Picture = objStdPic
    mobjfrmPatiInfo.mlng图像操作 = 4
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Public Property Get SendCard() As Boolean
    SendCard = mblnSendCard
End Property

Private Sub Update证件(ByVal lng病人ID As Long, ByVal str证件名 As String)
    '功能：更新当前证件类型的卡号
    '问题号:90875
    Dim strSQL As String, rsTmp As ADODB.Recordset
    On Error GoTo Errhand
    If IDKind证件.IDKind = IDKind证件.GetKindIndex("身份证号") Then Exit Sub
    If str证件名 = "身份证号" Then Exit Sub
    txt证件.Text = "": txt证件.Tag = ""
    If mrsInfo Is Nothing Then Exit Sub
    strSQL = "Select A.卡号,B.名称 from 病人医疗卡信息 A,医疗卡类别 B,证件类型 C " & _
            "Where A.卡类别ID=B.ID And B.名称=C.名称 And A.病人ID=[1] And B.名称=[2] Order by C.编码 "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, str证件名)
    If Not rsTmp.EOF Then txt证件.Text = Nvl(rsTmp!卡号): txt证件.Tag = txt证件.Text
    Exit Sub
Errhand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub txt证件_GotFocus()
    zlControl.TxtSelAll txt证件
End Sub

Private Sub txt证件_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txt证件_Validate(Cancel As Boolean)
    If txt证件.Text = txt证件.Tag Then Exit Sub
    '更新病人信息
    Call CopyZJTofrmPatiInfo
    If Trim(txt证件.Text) = "" Then Exit Sub
    If Len(Trim(txt证件.Text)) > 30 Then
         MsgBox "证件输入字符超出最大字符数30,多出的字符将被自动截除！", vbInformation, gstrSysName
         txt证件.Tag = Mid(Trim(txt证件.Text), 1, 30)
         txt证件.Text = Mid(Trim(txt证件.Text), 1, 30)
    End If
    Call GetPatient(IDKind证件.GetCurCard, txt证件.Text, False, False, Cancel, True)
End Sub

Private Function AddCertificate(ByVal lng病人ID As Long, ByRef colPro As Collection, ByVal dtCurdate As Date) As Boolean
    '-------------------------------------------------------------------------------------------------------------------------
    '功能:建立证件卡类信息，如果是第一次建立卡类别
    '编制:李南春
    '时间:2015/12/17 17:37:27
    '问题:90875
    '-------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim lngRow As Integer, lngCol As Integer
    Dim lngID As Long, strCode As String
    
    On Error GoTo Errhand
    If IDKind证件.IDKind = IDKind证件.GetKindIndex("身份证号") Or txt证件.Text = "" Then AddCertificate = True: Exit Function
    '检查卡号是否被他人使用
    strSQL = "Select 1 from 病人医疗卡信息 A,医疗卡类别 B " & _
            "Where A.卡类别ID=B.ID And B.名称=[1] And B.是否证件=1 And A.卡号=[2] And  A.病人ID<>[3]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, IDKind证件.GetCurCard.名称, Trim(txt证件.Text), lng病人ID)
    If rsTemp.RecordCount > 0 Then
        MsgBox IDKind证件.GetCurCard.名称 & ":" & txt证件.Text & "正在被使用,请检查!", vbInformation, gstrSysName
        If txt证件.Visible And txt证件.Enabled Then txt证件.SetFocus
        Exit Function
    End If
    '绑定卡前要判断卡类别是否存在
    strSQL = "Select B.ID,B.编码,B.卡号长度,B.名称,A.卡号,A.病人ID,Decode(A.卡号 ,NULL,1,0) as 标识 from 病人医疗卡信息 A,医疗卡类别 B " & _
            "Where A.卡类别ID(+)=B.ID And B.是否证件=1 And A.状态(+)=0 And B.名称=[1] "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, IDKind证件.GetCurCard.名称)

    If rsTemp.RecordCount = 0 Then
        lngID = zlDatabase.GetNextId("医疗卡类别")
        strCode = zlDatabase.GetMax("医疗卡类别", "编码", 4)
        mobjfrmPatiInfo.mstrFirstCode = strCode
        Call AddCardTypeSQL(0, lngID, strCode, IDKind证件.GetCurCard.名称, IDKind证件.GetCurCard.短名, Len(Trim(txt证件.Text)), colPro)
    ElseIf Len(Trim(txt证件.Text)) > Val(Nvl(rsTemp!卡号长度)) Then
        lngID = rsTemp!ID
        Call AddCardTypeSQL(1, lngID, Nvl(rsTemp!编码), IDKind证件.GetCurCard.名称, IDKind证件.GetCurCard.短名, Len(Trim(txt证件.Text)), colPro)
    Else
        lngID = rsTemp!ID
    End If
    
    '进行证件卡绑定
    rsTemp.Filter = "名称='" & IDKind证件.GetCurCard.名称 & "' And 卡号='" & Trim(txt证件.Text) & "'"
    If rsTemp.RecordCount = 0 Then
        '先将病人原来的卡解绑
        rsTemp.Filter = "名称='" & IDKind证件.GetCurCard.名称 & "' And 病人ID= " & lng病人ID
        If rsTemp.RecordCount > 0 Then
            rsTemp.MoveFirst
            Do While Not rsTemp.EOF
                'Zl_医疗卡变动_Insert
                 strSQL = "Zl_医疗卡变动_Insert("
                '      变动类型_In   Number,
                '发卡类型=1-发卡(或11绑定卡);2-换卡;3-补卡(13-补卡停用);4-退卡(或14取消绑定); ５-密码调整(只记录);6-挂失(16取消挂失)
                strSQL = strSQL & "" & 14 & ","
                '      病人id_In     住院费用记录.病人id%Type,
                strSQL = strSQL & "" & lng病人ID & ","
                '      卡类别id_In   病人医疗卡信息.卡类别id%Type,
                strSQL = strSQL & "" & lngID & ","
                '      原卡号_In     病人医疗卡信息.卡号%Type,
                strSQL = strSQL & "'" & "" & "',"
                '      医疗卡号_In   病人医疗卡信息.卡号%Type,
                strSQL = strSQL & "'" & rsTemp!卡号 & "',"
                '      变动原因_In   病人医疗卡变动.变动原因%Type,
                '      --变动原因_In:如果密码调整，变动原因为密码.加密的
                strSQL = strSQL & "'" & "证件卡取消绑定" & "',"
                '      密码_In       病人信息.卡验证码%Type,
                strSQL = strSQL & "'" & "" & "',"
                '      操作员姓名_In 住院费用记录.操作员姓名%Type,
                strSQL = strSQL & "'" & UserInfo.姓名 & "',"
                '      变动时间_In   住院费用记录.登记时间%Type,
                strSQL = strSQL & "to_date('" & Format(dtCurdate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
                '      Ic卡号_In     病人信息.Ic卡号%Type := Null,
                strSQL = strSQL & "'" & "" & "',"
                '      挂失方式_In   病人医疗卡变动.挂失方式%Type := Null
                strSQL = strSQL & "NULL)"

                zlAddArray colPro, strSQL
                rsTemp.MoveNext
            Loop
        End If
            
        '进行证件卡绑定
        'Zl_医疗卡变动_Insert
         strSQL = "Zl_医疗卡变动_Insert("
        '      变动类型_In   Number,
        '发卡类型=1-发卡(或11绑定卡);2-换卡;3-补卡(13-补卡停用);4-退卡(或14取消绑定); ５-密码调整(只记录);6-挂失(16取消挂失)
        strSQL = strSQL & "" & 11 & ","
        '      病人id_In     住院费用记录.病人id%Type,
        strSQL = strSQL & "" & lng病人ID & ","
        '      卡类别id_In   病人医疗卡信息.卡类别id%Type,
        strSQL = strSQL & "" & lngID & ","
        '      原卡号_In     病人医疗卡信息.卡号%Type,
        strSQL = strSQL & "'" & "" & "',"
        '      医疗卡号_In   病人医疗卡信息.卡号%Type,
        strSQL = strSQL & "'" & Trim(txt证件.Text) & "',"
        '      变动原因_In   病人医疗卡变动.变动原因%Type,
        '      --变动原因_In:如果密码调整，变动原因为密码.加密的
        strSQL = strSQL & "'" & "证件卡绑定" & "',"
        '      密码_In       病人信息.卡验证码%Type,
        strSQL = strSQL & "'" & "" & "',"
        '      操作员姓名_In 住院费用记录.操作员姓名%Type,
        strSQL = strSQL & "'" & UserInfo.姓名 & "',"
        '      变动时间_In   住院费用记录.登记时间%Type,
        strSQL = strSQL & "to_date('" & Format(dtCurdate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
        '      Ic卡号_In     病人信息.Ic卡号%Type := Null,
        strSQL = strSQL & "'" & "" & "',"
        '      挂失方式_In   病人医疗卡变动.挂失方式%Type := Null
        strSQL = strSQL & "NULL)"
    
        zlAddArray colPro, strSQL
    End If
    AddCertificate = True
    Exit Function
Errhand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub CreateCommunity()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:创建社区部件
    '编制:刘兴洪
    '日期:2017-08-09 11:25:57
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnInit As Boolean
    If mbytMode <> 0 Then Exit Sub
    
    '社区接口初始化
    Err = 0: On Error Resume Next
    
    blnInit = False
    If mobjCommunity Is Nothing Then
       Set mobjCommunity = CreateObject("zlCommunity.clsCommunity")
       If Not mobjCommunity Is Nothing Then
           blnInit = mobjCommunity.Initialize(gcnOracle)
           If Not blnInit Then Set mobjCommunity = Nothing
       End If
    End If
    blnInit = Not mobjCommunity Is Nothing
    cmdComminuty.Visible = blnInit
    cmdComminuty.Enabled = blnInit
    Err = 0: On Error GoTo 0
End Sub
Private Function GetRegistMoney(Optional blnOnlyReg As Boolean = False, Optional blnNoBook As Boolean = False) As Double
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取当前挂号单的合计金额
    '入参:blnOnlyReg-是否仅仅读取挂号费用
    '     blnNoBook-读取病历费
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2017-11-03 16:53:41
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dbl合计 As Double, i As Integer
    Dim k As Integer
    If Not blnOnlyReg Then
        dbl合计 = FormatEx(GetTotalFromMshMoney, 5)
    Else
        If mrsItems Is Nothing Then
             GetRegistMoney = FormatEx(GetTotalFromMshMoney, 3): Exit Function
        End If
        mrsItems.Filter = " 性质 <> 4"
        If mrsItems.RecordCount = 0 Then
            mrsItems.Filter = 0
            GetRegistMoney = FormatEx(GetTotalFromMshMoney, 3): Exit Function
        End If
        With mrsItems
            Do While Not .EOF
                dbl合计 = dbl合计 + GetTotalFromMshMoney(Nvl(mrsItems!项目名称, "-"))
                .MoveNext
            Loop
        End With
        mrsItems.Filter = 0
    End If
    
    If blnNoBook Then
        If Not mrsItems Is Nothing Then
            mrsItems.Filter = " 性质 = 3"
            Do While Not mrsItems.EOF
                dbl合计 = dbl合计 + GetTotalFromMshMoney(Nvl(mrsItems!项目名称, "-"))
                mrsItems.MoveNext
            Loop
            mrsItems.Filter = 0
        End If
    End If
    GetRegistMoney = FormatEx(dbl合计, 5)
End Function

Private Function GetTotalFromMshMoney(Optional ByVal str项目名称 As String = "") As Double
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取汇总金额
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2017-11-03 16:57:41
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblMoney As Double, i As Long
    
    On Error GoTo errHandle
    With vsfMoney
        For i = 1 To .Rows - 1
            If str项目名称 = "" Or Trim(.TextMatrix(i, 0)) = str项目名称 Then
                dblMoney = dblMoney + Val(.TextMatrix(i, 2))
            End If
        Next
    End With
    GetTotalFromMshMoney = dblMoney
    Exit Function
errHandle:
    GetTotalFromMshMoney = 0
End Function
Private Function GetCardMoney() As Double
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取卡费金额
    '编制:刘兴洪
    '日期:2017-11-03 17:39:27
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dbl合计 As Double
    If mrsItems Is Nothing Then GetCardMoney = 0: Exit Function
    mrsItems.Filter = " 性质 = 4"
    If mrsItems.RecordCount = 0 Then
        mrsItems.Filter = 0
        GetCardMoney = 0: Exit Function
    End If
    With mrsItems
        Do While Not .EOF
            dbl合计 = dbl合计 + GetTotalFromMshMoney(Nvl(mrsItems!项目名称, "-"))
            .MoveNext
        Loop
    End With
    mrsItems.Filter = 0
End Function

Private Function CheckIsPrice() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:功能检查当前是否为划价单保存
    '返回:保存为划价单成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2017-11-03 14:03:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dbl缴款 As Double
    Dim dbl个帐 As Double, dbl预交 As Double
    Dim blnPrice As Boolean
    
    On Error GoTo errHandle
    If mRegistFeeMode = EM_RG_记帐 Then CheckIsPrice = False: Exit Function
    If InStr(1, "02", mbytMode) = 0 Then CheckIsPrice = False: Exit Function
    
    If Not gblnPrice Or txtPatient.Text = "" Then CheckIsPrice = False: Exit Function
    blnPrice = picBookingDate.Visible = False And vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("病案")) <> ""
    If blnPrice Then blnPrice = GetRegistMoney <> 0
    
    CheckIsPrice = blnPrice
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function
 
Private Sub ReLoadCardFee(Optional ByVal blnFeedName As Boolean, Optional ByVal blnReflashfee As Boolean)
    '离开检查卡费
    Dim lng病人ID As Long, lng收费细目ID As Long
    Dim strSQL As String, str年龄 As String
    Dim rsTmp As ADODB.Recordset
    
    gCurSendCard.lng收费细目ID = 0
    If gCurSendCard.rs卡费 Is Nothing Then Exit Sub
    If gCurSendCard.rs卡费.RecordCount = 0 Then Exit Sub
    If gCurSendCard.lng卡类别ID = 0 Then Exit Sub
    If Trim(txtPatient.Text) = "" Or Trim(mobjfrmPatiInfo.txt卡号.Text) = "" Then Exit Sub
    If mrsInfo Is Nothing Then
        lng病人ID = 0
    ElseIf mrsInfo.State <> 1 Then
        lng病人ID = 0
    Else
        lng病人ID = mrsInfo!病人ID
    End If
    If blnFeedName = False And lng病人ID <> 0 Then Exit Sub
    
    str年龄 = Trim(txt年龄.Text)
    If IsNumeric(str年龄) Then str年龄 = str年龄 & cbo年龄单位.Text
    gCurSendCard.rs卡费.MoveFirst
    
    strSQL = "Select Zl1_Ex_CardFee([1],[2],[3],[4],[5],[6],[7],[8],[9]) as 收费细目ID From Dual "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "卡费", mlngModul, gCurSendCard.lng卡类别ID, Trim(mobjfrmPatiInfo.txt卡号.Text), lng病人ID, _
                Trim(txtPatient.Text), NeedName(cbo性别.Text), str年龄, txtIDCard.Text, Val(Nvl(gCurSendCard.rs卡费!收费细目ID)))
    If rsTmp.EOF Then Exit Sub
    
    lng收费细目ID = Val(Nvl(rsTmp!收费细目ID))
    Set rsTmp = zlGetSpecialItemFee("", mobjfrmPatiInfo.mstrPriceGrade, lng收费细目ID)
    If Not rsTmp Is Nothing Then
        Set gCurSendCard.rs卡费 = rsTmp
        gCurSendCard.lng收费细目ID = lng收费细目ID
        If blnReflashfee Then Call ShowRegistFromInput
    End If
End Sub

Private Sub imgColPlan_Click()
    Dim lngLeft As Long, lngTop As Long
    Dim vRect  As RECT
    vRect = zlControl.GetControlRect(picImgPlan.Hwnd)
    lngLeft = vRect.Left
    lngTop = vRect.Top + picImgPlan.Height
    Call frmVsColSel.ShowColSet(Me, Me.Caption, vsfPlan, lngLeft, lngTop, imgColPlan.Height)
    zl_vsGrid_Para_Save mlngModul, vsfPlan, Me.Caption, "vsfPlan" & mbytMode
End Sub

Private Sub picImgPlan_Click()
    Call imgColPlan_Click
End Sub

Private Sub InitRegist()
    '初始化挂号
    Dim strDept As String
    Set mobjRegist = New clsRegist
    mobjRegist.zlInitCommon glngSys, gcnOracle, gstrDBUser
    mobjRegist.zlCancelRegNo '如果上次是程序以外崩溃，需要进行解锁
End Sub

Private Function ReserveRegNo(ByVal lng记录ID As Long, ByRef lngSN As Long, _
                            ByVal str发生时间 As String, ByVal Datsys As Date) As Boolean
    Dim blnLock As Boolean, bln分时段 As Boolean
    Dim strLockTime As String
    Dim strSQL As String, rsTmp As ADODB.Recordset
    On Error GoTo errH
    mlng锁号记录ID = 0
    If vsfPlan.TextMatrix(vsfPlan.Row, GetCol("序号控制")) <> "" And lng记录ID <> 0 Then
        mlng锁号记录ID = lng记录ID
        bln分时段 = (mViewMode = v_专家号分时段 Or mViewMode = V_普通号分时段)
        If Not (mbytMode = 2 Or mbytMode = 0 And mbytInState = 0 And mstrNoIn <> "") Then
            blnLock = True: strLockTime = str发生时间
        Else
            If mTy_Para.byt接收模式 = 0 And bln分时段 And Format(dtpAppointmentDate.Value, "yyyy-MM-dd") <> Format(Datsys, "yyyy-MM-dd") Then
                MsgBox "分时段的预约挂号单只能当天接收。", vbInformation, gstrSysName
                Exit Function
            End If
            If (mTy_Para.byt接收模式 = 0 And Format(dtpAppointmentDate.Value, "yyyy-MM-dd") <> Format(Datsys, "yyyy-MM-dd")) Then
                blnLock = True: strLockTime = Format(Datsys, "yyyy-MM-dd")
                strSQL = "Select a.Id" & vbNewLine & _
                        "From 临床出诊记录 a, 临床出诊记录 b" & vbNewLine & _
                        "Where a.号源id = b.号源id And a.是否分时段 = b.是否分时段 And a.是否序号控制 = b.是否序号控制 And a.科室id = b.科室id And" & vbNewLine & _
                        "      Nvl(a.医生id, 0) = Nvl(b.医生id, 0) And a.上班时段 = b.上班时段 And Nvl(a.是否发布, 0) = 1 And a.出诊日期 = Trunc(Sysdate) And" & vbNewLine & _
                        "      b.Id = [1] And Rownum < 2"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "更换出诊ID", lng记录ID)
                If rsTmp.RecordCount = 0 Then
                    MsgBox "接收当天没有对应的出诊安排，无法接收。", vbInformation, gstrSysName
                    Exit Function
                End If
                mlng锁号记录ID = rsTmp!ID
            End If
        End If
        If blnLock Then
            If mobjRegist.zlReserveRegNo(txt号别.Text, True, bln分时段, strLockTime, lngSN, "挂号窗口锁号", mlng锁号记录ID) = False Then Exit Function
        End If
    End If
    ReserveRegNo = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

