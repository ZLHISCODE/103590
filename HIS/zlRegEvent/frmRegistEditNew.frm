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
   Caption         =   "����ҺŴ���"
   ClientHeight    =   10920
   ClientLeft      =   60
   ClientTop       =   1350
   ClientWidth     =   15105
   BeginProperty Font 
      Name            =   "����"
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
   StartUpPosition =   1  '����������
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
               Caption         =   "�˸��ӷ�"
               Height          =   240
               Left            =   1320
               TabIndex        =   102
               Top             =   2160
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.TextBox txt����ʱ�� 
               Height          =   360
               Left            =   4725
               Locked          =   -1  'True
               TabIndex        =   101
               TabStop         =   0   'False
               Top             =   2100
               Width           =   2550
            End
            Begin VB.ComboBox cbo��ע 
               Height          =   330
               Left            =   525
               TabIndex        =   100
               Top             =   2520
               Width           =   6765
            End
            Begin VB.ComboBox cboԤԼ��ʽ 
               Height          =   330
               Left            =   2475
               Style           =   2  'Dropdown List
               TabIndex        =   99
               Top             =   2115
               Visible         =   0   'False
               Width           =   1080
            End
            Begin VB.CheckBox chk������ 
               Caption         =   "������"
               Height          =   240
               Left            =   0
               TabIndex        =   98
               Top             =   2160
               Width           =   1275
            End
            Begin VB.ComboBox cbo�ѱ� 
               Height          =   330
               IMEMode         =   3  'DISABLE
               Left            =   4725
               Style           =   2  'Dropdown List
               TabIndex        =   97
               Top             =   0
               Width           =   2550
            End
            Begin VB.ComboBox cbo���ʽ 
               Height          =   330
               Left            =   915
               Style           =   2  'Dropdown List
               TabIndex        =   96
               Top             =   435
               Width           =   2550
            End
            Begin VB.TextBox txt����� 
               Enabled         =   0   'False
               Height          =   360
               IMEMode         =   3  'DISABLE
               Left            =   4725
               TabIndex        =   95
               ToolTipText     =   "���ո�����µ������"
               Top             =   435
               Width           =   2550
            End
            Begin VB.ComboBox cboҽ����� 
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
                  Name            =   "����"
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
            Begin VB.Label lblժҪ 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��ע"
               Height          =   210
               Left            =   0
               TabIndex        =   110
               Top             =   2595
               Width           =   420
            End
            Begin VB.Label lbl����ʱ�� 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����ʱ��"
               Height          =   210
               Left            =   3840
               TabIndex        =   109
               Top             =   2175
               Width           =   840
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "�ѱ�"
               Height          =   210
               Left            =   4260
               TabIndex        =   108
               Top             =   60
               Width           =   420
            End
            Begin VB.Label lbl����� 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "�����"
               Height          =   210
               Left            =   4050
               TabIndex        =   107
               Top             =   495
               Width           =   630
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "���ʽ"
               Height          =   210
               Left            =   45
               TabIndex        =   106
               Top             =   495
               Width           =   840
            End
            Begin VB.Label lblҽ����� 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "ҽ�����"
               Height          =   210
               Left            =   45
               TabIndex        =   105
               Top             =   60
               Width           =   840
            End
            Begin VB.Label lblԤԼ��ʽ 
               AutoSize        =   -1  'True
               Caption         =   "ԤԼ��ʽ"
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
               Caption         =   "�ϼ�"
               BeginProperty Font 
                  Name            =   "����"
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
            Begin VB.Label lbl�ϼ� 
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
            Tag             =   "���֤��"
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
               Caption         =   "ȷ��(&O)"
               Height          =   390
               Left            =   690
               TabIndex        =   23
               ToolTipText     =   "�ȼ�:F2"
               Top             =   1830
               Width           =   1200
            End
            Begin VB.CommandButton cmdCancel 
               Cancel          =   -1  'True
               Caption         =   "ȡ��(&C)"
               Height          =   390
               Left            =   2100
               TabIndex        =   24
               Top             =   1845
               Width           =   1200
            End
            Begin VB.TextBox txt�Ҳ� 
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
            Begin VB.TextBox txt�ɿ� 
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
            Begin VB.TextBox txt����Ӧ�� 
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
               ToolTipText     =   "����Ӧ�ɺϼ�=�ۼ�ʵ�ɽ��-�ۼƸ����ʻ�֧��-�ۼƳ�Ԥ����"
               Top             =   465
               Width           =   2760
            End
            Begin VB.TextBox txt�ϼ� 
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
               ToolTipText     =   "����Ӧ�ɺϼ�=�ۼ�ʵ�ɽ��-�ۼƸ����ʻ�֧��-�ۼƳ�Ԥ����"
               Top             =   0
               Width           =   2760
            End
            Begin VB.ComboBox cbo���㷽ʽ 
               BeginProperty Font 
                  Name            =   "����"
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
            Begin VB.Label lbl�Ҳ� 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "�Ҳ�"
               BeginProperty Font 
                  Name            =   "����"
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
            Begin VB.Label lbl�ɿ� 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "�ɿ�"
               BeginProperty Font 
                  Name            =   "����"
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
            Begin VB.Label lblӦ�� 
               AutoSize        =   -1  'True
               Caption         =   "Ӧ��"
               BeginProperty Font 
                  Name            =   "����"
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
               Caption         =   "�ϼ�"
               BeginProperty Font 
                  Name            =   "����"
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
         Begin VB.TextBox txt��ͥ�绰 
            Height          =   360
            Left            =   4815
            MaxLength       =   20
            TabIndex        =   13
            Top             =   2520
            Width           =   2550
         End
         Begin VB.TextBox txt���� 
            Height          =   360
            IMEMode         =   2  'OFF
            Left            =   5730
            TabIndex        =   9
            Top             =   2055
            Width           =   930
         End
         Begin VB.ComboBox cbo�Ա� 
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
            ToolTipText     =   "��������(Ctrl+M)"
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
            ToolTipText     =   "���Ҳ���(Ctrl+F)"
            Top             =   1635
            Width           =   350
         End
         Begin VB.ComboBox cbo���䵥λ 
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
            ToolTipText     =   "�ȼ�:F11"
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
            ToolTipText     =   "�󶨾��￨:F10"
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
            ToolTipText     =   "�������������֤"
            Top             =   1635
            Width           =   350
         End
         Begin VB.CommandButton cmdYb 
            Caption         =   "ҽ��"
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
         Begin VB.TextBox txt�ű� 
            BackColor       =   &H00EBFFFF&
            Height          =   360
            IMEMode         =   3  'DISABLE
            Left            =   600
            TabIndex        =   1
            ToolTipText     =   "F9��λ��ѯ�ʹҺſ��ң�����""+""��������,����"".""������,����""-""�ű�ʾ��ʾ���кű�"
            Top             =   630
            Width           =   2355
         End
         Begin VB.TextBox txt���� 
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
         Begin VB.ComboBox cboҽ�� 
            ForeColor       =   &H00C00000&
            Height          =   330
            IMEMode         =   2  'OFF
            Left            =   4815
            TabIndex        =   4
            ToolTipText     =   "����ѡ�ѱ�ҽ��Ϊ���ұ��ز���Ҫ����ҽ��ʱ����������"
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
            Caption         =   "��"
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   6585
            Style           =   1  'Graphical
            TabIndex        =   54
            ToolTipText     =   "�ȼ�:F8"
            Top             =   135
            Width           =   360
         End
         Begin VB.CheckBox chkPrint 
            Caption         =   "��"
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   6975
            Style           =   1  'Graphical
            TabIndex        =   53
            ToolTipText     =   "�ȼ�:F7"
            Top             =   135
            Width           =   360
         End
         Begin VB.ComboBox cboNO 
            ForeColor       =   &H00C00000&
            Height          =   330
            IMEMode         =   3  'DISABLE
            Left            =   4815
            TabIndex        =   52
            ToolTipText     =   "�ȼ�:F12"
            Top             =   135
            Width           =   1725
         End
         Begin VB.CheckBox chkBooking 
            Caption         =   "Ԥ"
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   6975
            Style           =   1  'Graphical
            TabIndex        =   51
            ToolTipText     =   "Ԥ�ҽ����ĺ�,�ȼ�:Ctrl+F12"
            Top             =   645
            Width           =   360
         End
         Begin VB.CommandButton cmdPatiPic 
            Height          =   330
            Left            =   6585
            Picture         =   "frmRegistEditNew.frx":139C
            Style           =   1  'Graphical
            TabIndex        =   50
            ToolTipText     =   "��ʾ������Ƭ,�ȼ�:Ctrl+W"
            Top             =   645
            Width           =   360
         End
         Begin VB.TextBox txtPatientPrint 
            Height          =   360
            Left            =   1200
            TabIndex        =   49
            ToolTipText     =   "�ȼ�:F11"
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
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontSize        =   10.5
            FontName        =   "����"
            IDKind          =   -1
            ShowPropertySet =   -1  'True
            DefaultCardType =   "0"
            BackColor       =   -2147483633
         End
         Begin MSMask.MaskEdBox txt����ʱ�� 
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
               Name            =   "����"
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
         Begin MSMask.MaskEdBox txt�������� 
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
               Name            =   "����"
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
               Name            =   "����"
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
         Begin zlIDKind.IDKindNew IDKind֤�� 
            Height          =   375
            Left            =   600
            TabIndex        =   91
            TabStop         =   0   'False
            Top             =   2520
            Width           =   630
            _ExtentX        =   1111
            _ExtentY        =   661
            Appearance      =   2
            IDKindStr       =   "��|�������֤|0|0|0|0|0|0"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontSize        =   12
            FontName        =   "����"
            IDKind          =   -1
            NotAutoAppendKind=   -1  'True
            BackColor       =   -2147483633
         End
         Begin VB.TextBox txt֤�� 
            BeginProperty Font 
               Name            =   "����"
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
            Tag             =   "֤��"
            Top             =   2520
            Width           =   2550
         End
         Begin ZlPatiAddress.PatiAddress padd���ڵ�ַ 
            Height          =   750
            Left            =   1005
            TabIndex        =   17
            Tag             =   "���ڵ�ַ"
            Top             =   3750
            Visible         =   0   'False
            Width           =   6360
            _ExtentX        =   11218
            _ExtentY        =   1323
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
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
         Begin ZlPatiAddress.PatiAddress padd��ͥ��ַ 
            Height          =   750
            Left            =   1005
            TabIndex        =   16
            Tag             =   "��סַ"
            Top             =   2955
            Visible         =   0   'False
            Width           =   6360
            _ExtentX        =   11218
            _ExtentY        =   1323
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
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
         Begin VB.ComboBox cbo���ڵ�ַ 
            Height          =   330
            Left            =   1005
            TabIndex        =   15
            Top             =   3375
            Width           =   6360
         End
         Begin VB.ComboBox cbo��ͥ��ַ 
            Height          =   330
            Left            =   1005
            TabIndex        =   14
            Top             =   2955
            Width           =   6360
         End
         Begin VB.Label lblIDCard 
            AutoSize        =   -1  'True
            Caption         =   "֤��"
            Height          =   210
            Left            =   135
            TabIndex        =   90
            ToolTipText     =   "֤������"
            Top             =   2595
            Width           =   420
         End
         Begin VB.Label lbl���ڵ�ַ 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���ڵ�ַ"
            Height          =   210
            Left            =   135
            TabIndex        =   83
            Top             =   3810
            Width           =   840
         End
         Begin VB.Label lbl��ͥ�绰 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "��ͥ�绰"
            Height          =   210
            Left            =   3930
            TabIndex        =   82
            Top             =   2595
            Width           =   840
         End
         Begin VB.Label lbl��ͥ��ַ 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��ס��ַ"
            Height          =   210
            Left            =   135
            TabIndex        =   81
            Top             =   3015
            Width           =   840
         End
         Begin VB.Label lbl�������� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��������"
            Height          =   210
            Left            =   2340
            TabIndex        =   80
            Top             =   2130
            Width           =   840
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����"
            Height          =   210
            Left            =   5280
            TabIndex        =   79
            Top             =   2130
            Width           =   420
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����"
            Height          =   210
            Left            =   135
            TabIndex        =   78
            Top             =   1695
            Width           =   420
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�Ա�"
            Height          =   210
            Left            =   135
            TabIndex        =   77
            Top             =   2130
            Width           =   420
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�ű�"
            Height          =   210
            Left            =   135
            TabIndex        =   76
            Top             =   705
            Width           =   420
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����"
            Height          =   210
            Left            =   135
            TabIndex        =   75
            Top             =   1140
            Width           =   420
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ҽ��"
            Height          =   210
            Left            =   4350
            TabIndex        =   74
            Top             =   1140
            Width           =   420
         End
         Begin VB.Label lblSN 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���"
            Height          =   210
            Left            =   4350
            TabIndex        =   73
            Top             =   705
            Width           =   420
         End
         Begin VB.Label lblNO 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���ݺ�"
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
            Caption         =   "Ʊ��"
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
            Caption         =   "����"
            BeginProperty Font 
               Name            =   "����"
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
         Tag             =   "���ڵ�ַ"
         Top             =   960
         Visible         =   0   'False
         Width           =   6360
         _ExtentX        =   11218
         _ExtentY        =   582
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
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
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����"
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
      Begin VB.Label lbl�� 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����"
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
      Begin VB.Label lbl���� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "������ҽ��"
         BeginProperty Font 
            Name            =   "����"
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
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
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
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   1376
            MinWidth        =   18
            Picture         =   "frmRegistEditNew.frx":997F
            Object.ToolTipText     =   "���˵��"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
               Name            =   "����"
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
            Caption         =   "ʣ��ɹҺϼ�:"
            BeginProperty Font 
               Name            =   "����"
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
            ToolTipText     =   "�����޺źű��ʣ��ɹ������ϼ�"
            Top             =   60
            Width           =   1875
         End
         Begin VB.Label lblRegTotal 
            AutoSize        =   -1  'True
            Caption         =   "333"
            BeginProperty Font 
               Name            =   "����"
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
         Begin VB.Label lblԤԼʱ�� 
            AutoSize        =   -1  'True
            Caption         =   "ʱ��"
            Height          =   210
            Left            =   3105
            TabIndex        =   88
            Top             =   150
            Width           =   420
         End
      End
      Begin VB.CheckBox chkShowAll 
         BackColor       =   &H00A0A0A0&
         Caption         =   "���кű�"
         ForeColor       =   &H00000005&
         Height          =   240
         Left            =   2565
         TabIndex        =   38
         TabStop         =   0   'False
         ToolTipText     =   "�ȼ���F6(ָ����Ŀ��ҷ�Χ�����кű�)"
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
               Name            =   "����"
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
            Caption         =   "ԤԼ����"
            Height          =   210
            Left            =   90
            TabIndex        =   33
            Top             =   60
            Width           =   840
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "ʱ��"
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
            Name            =   "����"
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
               ToolTipText     =   "ѡ����Ҫ��ʾ����(Ctrl+E)"
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
            Name            =   "����"
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
      Begin XtremeSuiteControls.ShortcutCaption sc���� 
         Height          =   315
         Left            =   15
         TabIndex        =   46
         Top             =   420
         Width           =   4605
         _Version        =   589884
         _ExtentX        =   8123
         _ExtentY        =   556
         _StockProps     =   6
         Caption         =   "�ҺŰ��ű�"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
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
            Caption         =   "����Ƭ"
            BeginProperty Font 
               Name            =   "����"
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
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����"
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
Option Explicit 'Ҫ���������
'��������ڲ���
Public mstrPrivs As String
Public mlngModul As Long
Public mbytMode As Integer '0-�Һ�,1-ԤԼ,2-����,3-ȡ��ԤԼ ,4-�˺� ԤԼ������ģʽ:0-�Һ�,��ʱԤԼҪ�շ�,1-ԤԼ,���շ�
Public mbytInState As Byte '0-ִ��,1-����
Public mintCancel As Integer '0-�˺�,1-�˲�����,2-�˸��ӷ�
Public int��¼״̬ As Integer '2-���ĳ���ԤԼ����,3-���ı�������ԭʼ���� ע��ȡ��ԤԼʱ mbytinstate=1
Public mblnViewCancel As Boolean '�Ƿ�鿴�˺ŵ���
Public mstrNoIn As String 'Ҫ���ջ���ĵĵ��ݺ�
Public mblnCharge As Boolean '�Ƿ��շ��ڵ���
Public mstr����NO As String '�˺�ͬʱҪɾ���Ļ��۵�
Public mblnICCard As Boolean 'IC������
'����ҽ��վʹ�õı���
Public mblnStation As Boolean '�Ƿ�ҽ������վ�ڵ��ùҺ�
Public mstrRoom As String 'ҽ������վ�Ľ�������
Public mstrRegNo As String 'ҽ��վ�Һųɹ�ʱ�ĹҺŵ���
Public mblnNoneCut As Boolean '�Ƿ�����ʹ�ô��۷ѱ�("�Һŷѱ����"Ȩ��)
Public mblnStationPrice As Boolean 'ҽ��վ�Һ�ʱ�Ƿ��������ɻ��۵��չҺŷ�
Public mblnViewOriginal As Boolean
  
'��Ϣģ��ʹ�õı���
Public mobjMsgModule As clsMipModule

'������ر�������������ȱʡ�������ͺ�ȱʡ��������
Private mCurSendCard As Ty_CardProperty   '���Ѻ͹Һŷ�һ����ʱ��Ч���ȷ������������������𷢿����ͱ�������Ҫ��ģ�������¼

'Ʊ����ر���
Private mlngShareUseID As Long '������������ID
Private mstrUseType As String 'ʹ�����
Private mintInvoiceFormat As Integer  '��ӡ�ķ�Ʊ��ʽ,��Ʊ��ʽ���
Private mblnStartFactUseType As Boolean   '�Ƿ�������ʹ�����
Private mintInvoicePrint As Integer  '0-����ӡ;1-�Զ���ӡ;2-��ʾ��ӡ

'״̬���Ʋ���
Private mblnOneCard As Boolean      '�Ƿ�������һ��ͨ�ӿ�,��ģʽ�£�Ʊ���ϸ����Ʊ�ŷ�Χ��ķ�����󶨿����շ�
Private mrsOneCard As ADODB.Recordset
Private mlng�ſ�����ID As Long '��ǰʹ�õľ��￨��������
Private mblnOnVilidate As Boolean
Private mlngĬ�Ͽ����ID As Long
Private mstrCardPrivs As String, mstrForceNote As String
Private mblnPre���� As Boolean
Private mblnUnitReg As Boolean  '��ԤԼʱ�Ƿ���Һź�����λ���ź���
Private mblnOk  As Boolean, mbln�����Һ� As Boolean
Private mblnStateChange As Boolean '�����ڽ��йҺ����״̬�����ʱ��,������vsflex���¼�
Private mstrPre�ű� As String '��һ����Ч�ű�
Private mlngPreRow As Long  '��һ����Ч����
Private mdblԤ����� As Double, mblnCenter As Boolean
Private mcbrToolBar As CommandBar, mbln�˺�ԭ�� As Boolean
Private mdbl������� As Double, mstrԭժҪ As String
Private mdblԭ��� As Double
Private mblnCancel As Boolean
Private mblnActivate As Boolean
Private mblnReadBooking As Boolean
Private mblnAppointmentChange As Boolean
Private mblnUserCancel As Boolean

Private mblnCard As Boolean '��ǰ�Ƿ���￨ˢ��
Private mblnNewCard As Boolean '���¿�
Private mblnUnload As Boolean, mblnChange As Boolean
Private mblnSendCard As Boolean
Private mblnBuyHisBook As Boolean
Private mblnUnChange As Boolean
Private mblnManualInput As Boolean
Private mintSysAppLimit As Integer
Private mblnFirst As Boolean
Private mblnAlwaysSend As Boolean '���ϸ����ʱʼ�շ���
Private mblnCheckNOValidity As Boolean
Private mstr����� As String
Private mdatLast As Date, msngTime As Single, mlngRow As Long
Private mblnChangeByCode As Boolean
Private mcur���� As Currency
Private mblnNoClearPrompt As Boolean
Public mblnNOMoved As Boolean
Public mintNOLength As Integer  '����ų���
Private mDatLastRefresh As Date '�ű��ϴ�ˢ��ʱ��
Private mblnReSetIDKind As Boolean 'ˢ����ŷ�ʽʱ,�����Һź�,�ָ�������Ϊ����ŷ�ʽ
Private mblnIDCardKind  As Boolean 'ԤԼ�Һ�ʱ,�������֤�ź�,�²����ڱ�����Ƿ��Զ��ָ������֤�ű���
Private mblnAddCardItem As Boolean '���Ѻ͹Һŷ�һ����ȡ
Private mblnBoundPati As Boolean '�󶨿�,����ȡ���˿���
Private mblnNotClick As Boolean '�Ƿ�����IDKind
Private mblnNotChange As Boolean '���ڿ����Ƿ���봥����txtsn��validate�¼�
Private mblnFinishReg As Boolean
Private mbln������Ϣ���� As Boolean '�Ƿ�����������˻�����Ϣ
Private mblnStructAdress As Boolean  '���˵�ַ�ṹ��¼��
Private mblnShowTown As Boolean      '�����ַ�ṹ��¼��

'��¼�Һ���ط�����Ϣ
Private mrsItems As ADODB.Recordset '��¼�Һ���Ŀ(����������Ŀ)
Private mrsInComes As ADODB.Recordset '��¼������Ŀ(����������Ϣ)
Private mrsDoctor As ADODB.Recordset '����������ҽ��ʱ(gblnҽ��),�ͻ��˻���ҽ����Ϣ
Private mrs��ͥ��ַ As ADODB.Recordset  '�����ͥ��ַ,��ʼʱ��ȡ������
Private mrsSNState As ADODB.Recordset   '��ǰ�ű�����״̬
Private mrsʱ��� As ADODB.Recordset    ' �ҺŰ���ʱ���
Private mrsUnitReg As ADODB.Recordset  '������λ����
Private mrsBill As ADODB.Recordset     'ԤԼ����ʱ����ԤԼ������Ϣ
Private mrsBillAdvance As ADODB.Recordset '�˺�ʱ,���ݶ�Ӧ��Ԥ����¼��Ϣ

Private mdblReg     As Double           '�Һŷ���
Private mlng�Һſ���ID As Long
Private mstrҽ������ As String
Private mlngҽ��ID As Long
Private mbln������ As Boolean
Private mstr�˷���ĿIDs As String
Private mbln���ӷ� As Boolean, mbln������ As Boolean
Private mstr���ӷ� As String, mstr������ĿID As String
Private mrs�ѱ� As ADODB.Recordset '�ѱ��б�
Private mstr�����Һ�_�Һ�NO As String, mstr�����Һ�_���￨NO As String
Private mblnUnChkClick As Boolean  '������checkbox��Click�¼�
Private mrsALLʱ��� As ADODB.Recordset '����:45509
Private mstrCurKey As String '��ǰ���ڼ�

'����ģ�����
Private mobjCommunity As Object     '�����ӿڲ���
Private mint���� As Integer
Private mstr������ As String

Private mrsPlan As ADODB.Recordset '�����ҺŰ�����Ϣ
 
Private mrsInfo As ADODB.Recordset '�����ҺŲ��������Ϣ
Private mbln������ As Boolean '�Ƿ������ȡ����������
Private mbln���������� As Boolean '�˺ŵĵ������Ƿ����������
Private mlng����ID As Long
Private mblnLEDKey As Boolean
Private mstrSort As String '�ű������ֶ�
Private mintIDKind As Integer '�ϴ�ʹ�õ�������ؼ�
Private mbln�Ӻ�   As Boolean '�Ƿ��ǼӺ��������

Private mstrPrePati As String '�ϴιҺŵĲ���,�򱾴����������֤����ݵĲ���
Private mstrPreNO As String '�ϴκű�
Private mcur�ϼ� As Currency '��ǰ�ۼƵ��ĺϼƽ��
Private mcurӦ�� As Currency '��ǰ�ۼƵ���Ӧ�ɽ��
Private mint�Һ��� As Integer     '�����Һ�ʱ��ͬһ�����ѹҺŶ��ŹҺ���
Private mstrPrepayPrivs As String 'Ԥ��Ȩ��
Private mlng��¼ID As Long 'ԤԼ�ĳ����¼ID������ʱ��ȡ
Private mlng���ż�¼ID As Long '���ŵļ�¼ID����ǰ���ջ��Ӻ����ʱ��ԭ��¼ID��һ��
Private mobjRegist As clsRegist

'ҽ����ر���
Private mintInsure As Integer
Private mlngOutModeMC As Long '����ҽ�����õ����ʽҽ������
Private mblnOlnyBJYB   As Boolean '�����Ǳ���ҽ��:������:����:26982
Private mblnNotQuery As Boolean  'δ�ҵ�����е�����,�ٱ���Һ�ʱ,��������
Private mblnBrushPlugin As Boolean '��ǰ�Ƿ�Ӳ����ȡ�Ĳ�����Ϣ
Private mstrYBPati As String 'ҽ�����������֤��Ϣ
Private mcur������� As Currency '�����ʻ����
Private mcur����͸֧ As Currency '�����ʻ�����͸֧���
Private mstr�����ʻ� As String  '�Һ��Ƿ�����ʹ�ø����ʻ�
Private mlng����ID As Long 'ҽ���˺�ʱ�Ľ���ID
'���˺� ����:26962 ����:2009-12-25 11:25:27
Private Type Ty_ModulePara
    bln�Һ����ɶ���         As Boolean '�Ŷӽк����ɶ���:ʵ�����Ƕ�ȡ���Ƿ������Ĳ���
    intͬ����Լ��           As Integer  'ͬ������Լ
    intͬ���޹���           As Integer
    blnͬ���޹Ҽ���         As Boolean
    int����ԤԼ������       As Integer
    int���˹Һſ�����       As Integer
    lngԤԼ��Чʱ��         As Long
    intԤԼʧЧ����         As Integer
    blnԤԼ����ȷ���Һŷ�   As Boolean
    bln����סԺ���˹Һ�     As Boolean '31724
    blnԤԼ�����������     As Boolean
    bln�����ͷ����         As Boolean '�Ƿ���������ͷ����
    bln������ѡ��         As Boolean ' ��������ŵ������ �Ƿ����� ����Ա���ѡ�����
    blnʧԼ���ڹҺ�         As Boolean '��ʱ��ʱ  ʧԼ���ڹҺ�
    lngN��ȡ��ԤԼ          As Long    'ԤԼN���ڲ���ȡ��ԤԼ
    bln�˺����             As Boolean '��N����ȡ��ԤԼ �Ƿ���Ҫͨ�����
    lngԤԼ����ʱ��         As Long    '����ԤԼ������ʱ�����С��� __����
    lngԤԼȱʡ����         As Long    'ԤԼʱȱʡ�������
    bln�Һű���ˢ��         As Boolean '38603
    byt��ͥ��ַ����         As Byte  '�Һż�ͥ��ַ���뷽ʽ �Ƿ�����
    bln�໤��¼��           As Boolean '�Ƿ���Ƽ໤��¼��
    lngN������¼��໤��    As Long '�໤��¼���������
    bln�ϸ�ʱ�ιҺ�       As Boolean  '�ϸ�ʱ�ιҺ�
    blnReuseCancelNO        As Boolean '�����������Һ�
    intר�ҺŹҺ�����       As Integer
    intר�Һ�ԤԼ����       As Integer
    bln��ֹ��������         As Boolean
    byt�ɿʽ             As Byte
    byt����ģʽ             As Byte
End Type
Private Enum SortType
    by�ű� '���ݺű��������
    by���� '���� ����-->��Ŀ--�ѹ��� ����������
    by����and�ѹ���
End Enum
Private mSortType As SortType '�������ʽ
Private mTy_Para As Ty_ModulePara
Private mstr��ǰ���� As String
Private mstrPre�ѱ� As String, mstrCard���㷽ʽ As String
Private mstr���� As String 'ԭ����
Private mstr�Ա� As String 'ԭ�Ա�
Private mstr���� As String 'ԭ����
Private mstr���䵥λ As String
Private mstr�������� As String

'�����һ��������������
Private Enum CustomTime
    t_��ͨ
    t_ʱ��
End Enum
Private Enum ViewMode
     V_��ͨ��
     v_ר�Һ�
     v_ר�Һŷ�ʱ��
     V_��ͨ�ŷ�ʱ��
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
'���㿨���
Private Type Ty_PayMoney
    lngҽ�ƿ����ID As Long
    bln���ѿ� As Boolean
    str���㷽ʽ As String
    str���� As String
    strˢ������ As String
    strˢ������ As String
    dbl�ʻ���� As Double
    objCard As Card
    Have�Һŷ� As Boolean
    Have���� As Boolean
End Type

Private mCurCardPay As Ty_PayMoney '���ο�֧��
Private mstrPassWord As String
Private mcolCardPayMode As Collection
Private mobjPayCard As Card

'�Һ����״̬��������' 2012-10-29 lgf
'��ʱֻ������ſ���,��ʱ�� ��״̬����
Private Type Ty_RegPlanState
    '״̬��¼
    str�ű�                 As String 'ѡ�еĺű�
    lngLastNO               As Long '����һ�����
    strLastNO_Time          As String '���һ��ʱ�ο�ʼʱ��
    strLastNo_EndTime       As String '����һ��ʱ�ν���ʱ��
    lngLastNO_X             As Long '���һ��������ڵ�λ��
    lngLastNO_Y             As Long '���һ��������ڵ�λ��
    bln��ſ���             As Boolean '��ſ���
    lng�޺���               As Long '�޺���
    lng��Լ��               As Long '��Լ��
    '״̬���Ʊ���
    '���±���,��Ҫ����,��ʱ��,��Ϊ��ʱ�εĺ�,������ź�ʱ��ͬʱ���ڵ����
    blnAdditionalNumber     As Boolean '�Ƿ��Ѿ�׷����� '׷����ŵ��ص�(�ҳ�ȥ�����,��Ŵ������õ�������,����ʱ����ڻ��ߵ���,���һ��ʱ�εĽ���ʱ��)
    lngSelX                 As Long 'ѡ�е���
    lngSelY                 As Long 'ѡ�е���
    lngSelNO                As Long 'ѡ�е����
    strSelTime              As String  'ѡ�е���Ŷ�Ӧʱ�εĿ�ʼʱ��
End Type

Private mtyRegPlanState As Ty_RegPlanState '�Һ�״̬����
Private mbln���� As Boolean '��ʶ��ǰ�����Ƿ��Ƿ���,True - ���� False - �󶨿�  �����:56599
Private mobjHealthCard As Object '�ƿ��ӿڶ���
Private mblnRegReceiveByNo As Boolean '�ж��Ƿ���ͨ���ڹҺŴ������뵥�ݺŽ���ԤԼ���ղ��� �����:57423
'-----------------------------------------------------------------------------------
Private mobjDelCards As Cards '��ǰ�˺����

Private Type TYPE_MedicarePAR
    ҽ���ӿڴ�ӡƱ�� As Boolean
    ʹ�ø����ʻ�   As Boolean  'support�Һ�ʹ�ø����ʻ�
    �����Һ�  As Boolean    'support�����Һ�
    ���ղ����� As Boolean   'support�ҺŲ���ȡ������
    �Һż����Ŀ As Boolean
End Type
Private MCPAR As TYPE_MedicarePAR
'-----------------------------------------------------------------------------------
Private Enum EM_REGISTFEE_MODE  '68991�Һŷ�����ȡ��ʽ
        EM_RG_���� = 0
        EM_RG_���� = 1
        EM_RG_���� = 2
End Enum
Private Enum EM_PATI_CHARGE_MODE    '�����շ�ģʽ
    EM_�Ƚ�������� = 0
    EM_�����ƺ���� = 1
End Enum
Private mRegistFeeMode As EM_REGISTFEE_MODE '�Һŷ�����ȡ��ʽ
Private mPatiChargeMode As EM_PATI_CHARGE_MODE    '�����շ�ģʽ
Private mstr���˼���IDs As String '����ʹ�ü���Ԥ����79868
Private mblnNotEMPIQuery As Boolean '��ֹ�����ĵ��ýӿ�
Private mlngEMPI����ID As Long '�ӿ��еĲ���ID
Private mstrPrePriceGrade As String
Private mblnGetBirth As Boolean '�ж��Ƿ�����ͨ�������������

Private Sub initInsurePara(ByVal lng����ID As Long)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��ҽ������
    '���:lng����ID-����ID
    '����:���˺�
    '����:2013-11-19 15:43:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    MCPAR.ҽ���ӿڴ�ӡƱ�� = gclsInsure.GetCapability(supportҽ���ӿڴ�ӡƱ��, lng����ID, mintInsure)
    MCPAR.ʹ�ø����ʻ� = gclsInsure.GetCapability(support�Һ�ʹ�ø����ʻ�, lng����ID, mintInsure)
    MCPAR.�����Һ� = gclsInsure.GetCapability(support�����Һ�, lng����ID, mintInsure)
    MCPAR.���ղ����� = gclsInsure.GetCapability(support�ҺŲ���ȡ������, lng����ID, mintInsure)
    MCPAR.�Һż����Ŀ = gclsInsure.GetCapability(support�Һż����Ŀ, lng����ID, mintInsure)
End Sub

Private Function SelectMemo(ByVal strInput As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ѡ����ժҪ
    '���:strInput-���봮;Ϊ��ʱ,��ʾȫ��
    '����:ѡ��ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-08-04 16:06:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnCancel As Boolean, strSQL As String, strWhere As String
    Dim rsInfo As ADODB.Recordset
    Dim vRect As RECT, strKey As String
    strKey = GetMatchingSting(strInput, False)
    If strInput <> "" Then
        If zlCommFun.IsCharChinese(cbo��ע.Text) Then
             strWhere = " And  ���� like [1] "
        ElseIf zlCommFun.IsNumOrChar(cbo��ע.Text) Then
             strWhere = " And (���� like upper([1]) or ���� like upper([1]))"
        End If
    End If
    
    strSQL = "" & _
     "   Select RowNum AS ID,����,����,����  " & _
     "   From ���ùҺ�ժҪ " & _
     "   Where 1=1 " & strWhere & _
     "   Order by ȱʡ��־"
     vRect = zlControl.GetControlRect(cbo��ע.Hwnd)
     On Error GoTo Hd
     Set rsInfo = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "���ùҺ�ժҪ", False, _
                    "", "", False, False, True, vRect.Left, vRect.Top, cbo��ע.Height, blnCancel, True, False, strKey)
     If blnCancel Then Exit Function
     If rsInfo Is Nothing Then
        If strInput = "" Then
            MsgBox "û�����ó��ùҺ�ժҪ,�����ֵ����������", vbOKOnly + vbInformation, gstrSysName
        End If
        zlCommFun.PressKey vbKeyTab: Exit Function
     End If
     zlControl.CboSetText Me.cbo��ע, Nvl(rsInfo!����)
     cbo��ע.Tag = Nvl(rsInfo!����)
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
    txt�ű�.Text = ""
    mblnUnChange = False
    Call ShowPlans(, True)
End Sub

Private Sub cbo��ע_Change()
    cbo��ע.Tag = ""
End Sub

Private Sub cbo��ע_Click()
    If mblnNotChange Then Exit Sub
    If chkCancel.Value = 1 Or mbytMode = 4 Then
        Call cbo��ע_KeyDown(13, 0)
    End If
End Sub

Private Sub cbo��ע_KeyDown(KeyCode As Integer, Shift As Integer)
    If chkCancel.Value = 1 Or mbytMode = 4 Then
        If KeyCode <> vbKeyReturn Then Exit Sub
        If Trim(cbo��ע.Text) = "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
        If SetDelMemo(Trim(cbo��ע.Text)) = True Then
            zlCommFun.PressKey vbKeyTab: Exit Sub
        End If
    Else
        If KeyCode <> vbKeyReturn Then Exit Sub
        If cbo��ע.Tag <> "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
        If Trim(cbo��ע.Text) = "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
        If SelectMemo(Trim(cbo��ע.Text)) = False Then
            zlCommFun.PressKey vbKeyTab: Exit Sub
        End If
    End If
End Sub

Private Function SetDelMemo(ByVal strInput As String) As Boolean
    Dim rsMemo As ADODB.Recordset, strSQL As String
    On Error GoTo errH
    If mbln�˺�ԭ�� = False Then SetDelMemo = True: Exit Function
    cbo��ע.Clear
    If strInput = "" Then
        strSQL = "Select ����,ȱʡ��־ From �����˺�ԭ�� Order By ȱʡ��־ Desc,����"
        Set rsMemo = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        If rsMemo.RecordCount <> 0 Then
            Do While Not rsMemo.EOF
                cbo��ע.AddItem rsMemo!����
                If Val(Nvl(rsMemo!ȱʡ��־)) = 1 Then
                    mblnNotChange = True
                    cbo��ע.ListIndex = cbo��ע.NewIndex: cbo��ע.Tag = cbo��ע.Text
                    mblnNotChange = False
                End If
                rsMemo.MoveNext
            Loop
        End If
    Else
        strSQL = "Select ����,ȱʡ��־,����,���� From �����˺�ԭ�� Order By ȱʡ��־ Desc,����"
        Set rsMemo = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        If rsMemo.RecordCount <> 0 Then
            Do While Not rsMemo.EOF
                cbo��ע.AddItem rsMemo!����

                If Nvl(rsMemo!����) Like UCase(strInput) & "*" Or Nvl(rsMemo!����) Like UCase(strInput) & "*" Or Nvl(rsMemo!����) Like strInput & "*" Then
                    mblnNotChange = True
                    cbo��ע.ListIndex = cbo��ע.NewIndex
                    mblnNotChange = False
                    cbo��ע.Tag = cbo��ע.Text
                End If
                rsMemo.MoveNext
            Loop
            If cbo��ע.Text = "" Then
                MsgBox "û���ҵ���Ӧ���˺�ԭ��,����������", vbInformation, gstrSysName
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

Private Sub cbo���ʽ_Click()
    Dim strPriceGrade As String
    
    If mbytInState = 1 Then Exit Sub
    
    If gintPriceGradeStartType < 2 Then Exit Sub
    Call gobjPublicExpense.zlGetPriceGrade(gstrNodeNo, 0, 0, zlStr.NeedName(cbo���ʽ.Text), , , strPriceGrade)
    mobjfrmPatiInfo.mstrPriceGrade = strPriceGrade
    If mstrPrePriceGrade = strPriceGrade Then Exit Sub
    mstrPrePriceGrade = strPriceGrade
    
    '31182:����ԤԼ����
    If (mbytMode = 2 Or mbytMode = 0 And mbytInState = 0 And mstrNoIn <> "") Then
        'ԤԼ����
        If mTy_Para.blnԤԼ����ȷ���Һŷ� = False Then
            If Not mrsInfo Is Nothing Then
                Exit Sub
            End If
        End If
    End If
    
    If txt�ű�.Text <> "" Then
        mblnBuyHisBook = True
        Call ShowRegistFromInput
        mblnBuyHisBook = False
    End If
End Sub

Private Sub cbo���䵥λ_LostFocus()
    Dim strBirth As String
    If cbo���䵥λ.Locked Then Exit Sub
    '������������
    With mobjfrmPatiInfo
        '69026,Ƚ����,2014-8-8,�����������
        If Trim(txt����.Text) <> "" Then
            If .mobjPubPatient Is Nothing Then Exit Sub
            If .mobjPubPatient.CheckPatiAge(Trim(txt����.Text) & cbo���䵥λ.Text) = False Then
                If txt����.Visible And txt����.Enabled And Not txt����.Locked Then
                    txt����.SetFocus: Exit Sub
                End If
            End If
        End If
    
        .txt����.Text = txt����.Text
        .txt����.Tag = txt����.Text
        If .cbo���䵥λ.ListCount = 0 Then CopyCboTofrmPatiInfo
        .cbo���䵥λ.ListIndex = cbo���䵥λ.ListIndex
        .cbo���䵥λ.Visible = cbo���䵥λ.Visible
        
        If cbo���䵥λ.Tag <> cbo���䵥λ.Text Then
            .mblnChange = False
            If mblnGetBirth Then
                If mobjfrmPatiInfo.mobjPubPatient.ReCalcBirthDay(Trim(txt����.Text) & cbo���䵥λ.Text, strBirth) Then
                    .txt��������.Text = Format(strBirth, "yyyy-mm-dd")
                    .txt����ʱ��.Text = Format(strBirth, "hh:mm")
                End If
            End If
            .mblnChange = True
            Call ReLoadCardFee(, True)
        Else
            Exit Sub
        End If
        '89130:���ϴ�,2015/10/13,���³�������
        mblnChange = False
        txt��������.Text = .txt��������.Text
        txt����ʱ��.Text = .txt����ʱ��.Text
        mblnChange = True
        cbo���䵥λ.Tag = cbo���䵥λ.Text
        Call ShowRegistFromInput
    End With
End Sub

Private Sub cbo�Ա�_LostFocus()
    Call ReLoadCardFee(, True)
End Sub


Private Sub cboԤԼ��ʽ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cbo�Ա�_Click()
    If mblnNotChange Then Exit Sub
    If cbo�Ա�.Enabled = False Then Exit Sub
    If cbo�Ա�.Tag <> cbo�Ա�.Text Then
        Call ShowRegistFromInput
    End If
    cbo�Ա�.Tag = cbo�Ա�.Text
End Sub


Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case conMenu_File_Exit
            Unload Me
        Case 2605 'Ԥ��
            Call HoldRegNo
        Case 2604 'ȡ��Ԥ��
            Call HoldRegNo
        Case conMenu_File_Print
            Call PrintSetup
        Case 3816 '��Ԥ��
            Call AddDeposit
        Case conMenu_View_Refresh
            Call RefreshFace
        Case 816 '�����
            frmRegistList.ShowMe Me, Val(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("��ԴID")))
        Case 4006 '��ʷ�Һ�
            Call SelectHistoryRegist
    End Select
End Sub

Private Sub PrintSetup()
    Dim strTmp As String
    
    If gblnPrintCase Then
        strTmp = zlCommFun.ShowMsgbox("��ӡ����", "��ѡ�����һ�ִ�ӡ���ݽ�������", "!�Һ�Ʊ��(&1),�Һ�ƾ��(&2),������ǩ(&3)", Me, vbInformation)
        If strTmp = "�Һ�Ʊ��" Then
            Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1111", Me)
        End If
        If strTmp = "�Һ�ƾ��" Then
            Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1111_3", Me)
        End If
        If strTmp = "������ǩ" Then
            Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1111_2", Me)
        End If
    Else
        Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1111", Me)
    End If
End Sub

Private Sub RefreshFace()
    mstrPreNO = vsfPlan.TextMatrix(vsfPlan.Row, GetCol("�ű�"))
    Call ShowPlans
    If gblnҽ�� And Not mblnStation Then Call GetAllҽ��
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If mbytInState = 0 And (mbytMode = 0 Or mbytMode = 1) Then
        Select Case Control.ID
        Case 816
            If vsfPlan.Row > vsfPlan.Rows - 1 Or vsfPlan.Col > vsfPlan.Cols - 1 Or vsfPlan.ColIndex("��ԴID") = -1 Then
                Control.Enabled = False
            Else
                Control.Enabled = vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("��ԴID")) <> ""
            End If
        Case 4006
            If mrsInfo Is Nothing Then
                Control.Enabled = False
            ElseIf mrsInfo.State <> 1 Then
                Control.Enabled = False
            ElseIf mrsInfo.RecordCount = 0 Then
                Control.Enabled = False
            ElseIf IsNull(mrsInfo!����ID) Then
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
        lblԤԼ��ʽ.Visible = True
        cboԤԼ��ʽ.Visible = True
        picBookingDate.Visible = True
    Else
        lblԤԼ��ʽ.Visible = False
        cboԤԼ��ʽ.Visible = False
        picBookingDate.Visible = False
    End If
    lblժҪ.Visible = True: cbo��ע.Visible = True
    Call SetPlanGrid
    Call SetPicTimeObjectVisible
    
    If chkBooking.Tag = "����" Then Exit Sub
    
    mblnUnChange = True     '����txt�ű�.Text = "" ʱ����ShowPlans
    Call ClearBill(, False)
    mblnUnChange = False
    Curdate = zlDatabase.Currentdate
    If blnBooking And Curdate > dtpAppointmentDate.Value Then  '����֮ǰ��ԤԼʱ��
        If Curdate < gdatRegistTime Then
            dtpAppointmentDate.Value = Format(gdatRegistTime + IIf(gintԤԼ���� >= 7, 7, mTy_Para.lngԤԼȱʡ����), "yyyy-MM-dd " & gstr�ϰ�ʱ��)
            dtpAppointmentDate.MinDate = Format(gdatRegistTime, "yyyy-MM-dd 00:00")
        Else
            dtpAppointmentDate.Value = Format(Curdate + IIf(gintԤԼ���� >= 7, 7, mTy_Para.lngԤԼȱʡ����), "yyyy-MM-dd " & gstr�ϰ�ʱ��)
            dtpAppointmentDate.MinDate = Format(Curdate, "yyyy-MM-dd 00:00")  '27781:��ǰ��һСʱ
        End If
    End If
    Call ShowPlans
    Call Form_Resize
    Call picPlan_Resize
    If txt�ű�.Visible And txt�ű�.Enabled Then txt�ű�.SetFocus
End Sub

Private Function GetPatiIDByComminuty(ByVal int���� As Integer, ByVal strComminuty As String) As Long
    Dim rsTmp As ADODB.Recordset, strSQL As String
        
    strSQL = "Select ����ID From ����������Ϣ Where ���� = [1] And ������ = [2]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, int����, strComminuty)
    If rsTmp.RecordCount > 0 Then GetPatiIDByComminuty = rsTmp!����ID
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
 

Private Sub cmdComminuty_Click()
    Dim lng����ID As Long
    Dim colInfo As Collection, strTmp As String
    
    If mrsInfo Is Nothing Then
        lng����ID = 0
    Else
        lng����ID = mrsInfo!����ID
    End If
    If Not mobjCommunity Is Nothing Then
        If mobjCommunity.Identify(glngSys, mlngModul, mint����, mstr������, colInfo, lng����ID) Then
            strTmp = GetColItem(colInfo, "����")
            If lng����ID = 0 Then
                lng����ID = GetPatiIDByComminuty(mint����, mstr������)
                If lng����ID = 0 Then
                    txtPatient.Text = strTmp
                Else
                    txtPatient.Text = "-" & lng����ID
                    Call txtPatient_Validate(False)
                End If
            Else
                If strTmp <> Trim(txtPatient.Text) Then
                    MsgBox "������֤�ӿڷ��صĲ��������뵱ǰ������������,�����Ƿ���ͬһ����!", vbInformation
                    Exit Sub
                End If
            End If
            strTmp = GetColItem(colInfo, "�Ա�")
            If strTmp <> "" Then cbo�Ա�.ListIndex = cbo.FindIndex(cbo�Ա�, strTmp, True)
                
            strTmp = GetColItem(colInfo, "��ͥ��ַ")
            If strTmp <> "" Then cbo��ͥ��ַ.Text = strTmp
            '89242:���ϴ�,2015/12/7,��ȡ���˵�ַ��Ϣ
            Call zlReadAddrInfo(padd��ͥ��ַ, lng����ID, 0, 3, cbo��ͥ��ַ.Text)
                                       
            '��ϸ������Ϣ����
            
            Call CopyCboTofrmPatiInfo
            Call CopyInfoTofrmPatiInfo
            With mobjfrmPatiInfo
                strTmp = GetColItem(colInfo, "����")
                If strTmp <> "" Then Call LoadOldData(strTmp, .txt����, .cbo���䵥λ)
                
                strTmp = GetColItem(colInfo, "��������")
                If IsDate(strTmp) Then
                    .mblnChange = False
                    .txt��������.Text = Format(strTmp, "YYYY-MM-DD")
                    .mblnChange = True
                    If CDate(.txt��������.Text) - CDate(strTmp) <> 0 Then .txt����ʱ��.Text = Format(strTmp, "HH:MM")
                    
                    .txt����.Text = ReCalcOld(CDate(.txt��������.Text), .cbo���䵥λ, lng����ID) '���ݳ���������������
                    .txt����.Tag = .txt����.Text
                Else
                    .mblnChange = False
                    .txt��������.Text = ReCalcBirth(.txt����.Text, .cbo���䵥λ.Text)
                    .mblnChange = True
                    .txt����ʱ��.Text = "__:__"
                End If
                
                txt����.Text = .txt����.Text
                txt����.Tag = txt����.Text
                cbo���䵥λ.ListIndex = .cbo���䵥λ.ListIndex
                Call txt����_Validate(False)
                
                strTmp = GetColItem(colInfo, "����")
                If strTmp <> "" Then .cbo����.ListIndex = cbo.FindIndex(.cbo����, strTmp, True)
                strTmp = GetColItem(colInfo, "����")
                If strTmp <> "" Then .cbo����.ListIndex = cbo.FindIndex(.cbo����, strTmp, True)
                strTmp = GetColItem(colInfo, "����״��")
                If strTmp <> "" Then .cbo����.ListIndex = cbo.FindIndex(.cbo����, strTmp, True)
                strTmp = GetColItem(colInfo, "ְҵ")
                If strTmp <> "" Then .cboְҵ.ListIndex = cbo.FindIndex(.cboְҵ, strTmp)
                strTmp = GetColItem(colInfo, "���֤��")
                If strTmp <> "" Then .txt���֤��.Text = strTmp: .txt���֤��.Tag = .txt���֤��.Text
                
                strTmp = GetColItem(colInfo, "������λ")
                If strTmp <> "" Then .txt��λ����.Text = strTmp
                strTmp = GetColItem(colInfo, "��λ�绰")
                If strTmp <> "" Then .txt��λ�绰.Text = strTmp
                strTmp = GetColItem(colInfo, "��λ�ʱ�")
                If strTmp <> "" Then .txt��λ�ʱ�.Text = strTmp
                
                strTmp = GetColItem(colInfo, "��ͥ�绰")
                If strTmp <> "" Then .txt��ͥ�绰.Text = strTmp
                strTmp = GetColItem(colInfo, "��ͥ��ַ�ʱ�")
                If strTmp <> "" Then .txt��ͥ�ʱ�.Text = strTmp
                strTmp = GetColItem(colInfo, "����")
                If strTmp <> "" Then .txt����.Text = strTmp: .txt����.Tag = .txt����.Text
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
    '����:ȡ��ԤԼ�Һ�
    '����:ȡ���ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-08-08 17:47:43
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    'ȡ��ԤԼ
    If mstrNoIn = "" Then Exit Function
    If zlCommFun.ActualLen(Me.cbo��ע.Text) > 50 Then
        MsgBox "��ע���ֻ������25�����ֻ�50���ַ�,����!", vbInformation + vbOKOnly, gstrSysName
        If cbo��ע.Enabled And cbo��ע.Visible Then cbo��ע.SetFocus
        Exit Function
    End If
    
    On Error GoTo errHandle
    
    strSQL = "zl_���˹Һż�¼_����_DELETE('" & mstrNoIn & "','" & UserInfo.��� & "','" & UserInfo.���� & "','" & Me.cbo��ע.Text & "')"
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
        'ȡ��ԤԼ
        If CancelBespeakRegist = False Then Exit Sub
        mblnOk = True
        gblnOk = True: Unload Me
        Exit Sub
    End If
    Call SaveData
    If Trim(txtSN.Text) <> "" Then Call mobjRegist.zlCancelRegNo(mlng���ż�¼ID)
End Sub

Private Sub cmdPatiPic_Click()
    '74430,Ƚ����,2014-7-8,�ҺŽ�����ʾ������Ƭ�ĸ�������
    Call ShowPatiPic
End Sub

Private Sub cmdRemark_Click()
    If SelectMemo("") = False Then
        zlCommFun.PressKey vbKeyTab: Exit Sub
    End If
End Sub
Private Sub cmdYb_Click()
     'ҽ�����֤��֤
     Call zlInusreIdentify
End Sub
Private Sub cmd�����Һ�_Click()
    Call SaveData(True)
End Sub
 
Private Sub dtpAppointmentTime_Change()
    Dim str���� As String, i As Integer, lngRow As Long
    If Not dtpAppointmentTime.Visible Then Exit Sub
    If Not dtpAppointmentTime.Enabled Then Exit Sub
    If dtpAppointmentDate.Visible Then
        str���� = Format(dtpAppointmentDate.Value, "yyyy-MM-dd")
    Else
        str���� = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    End If
    If str���� = "" Then str���� = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    
    txt����ʱ��.Text = str���� & " " & Format(dtpAppointmentTime.Value, "hh:mm:00")
    lngRow = 0
    If CDate(txt����ʱ��.Text) > CDate(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("����ʱ��"))) Then
        '����ʱ��İ��ţ�����Ѱ�Ҷ�λ
        For i = 1 To vsfPlan.Rows - 1
            With vsfPlan
                If .TextMatrix(.Row, .ColIndex("�ű�")) = .TextMatrix(i, .ColIndex("�ű�")) And _
                    CDate(vsfPlan.TextMatrix(i, vsfPlan.ColIndex("����ʱ��"))) >= CDate(txt����ʱ��.Text) Then
                    lngRow = i
                    Exit For
                End If
            End With
        Next i
    ElseIf CDate(txt����ʱ��.Text) < CDate(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("�Һ�ʱ��"))) Then
        '����ʱ��İ��ţ�����Ѱ�Ҷ�λ
        For i = 1 To vsfPlan.Rows - 1
            With vsfPlan
                If .TextMatrix(.Row, .ColIndex("�ű�")) = .TextMatrix(i, .ColIndex("�ű�")) And _
                    CDate(vsfPlan.TextMatrix(i, vsfPlan.ColIndex("�Һ�ʱ��"))) <= CDate(txt����ʱ��.Text) Then
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
    Dim lng�����ID As Long, strOutCardNO As String, strExpand
    Dim strOutPatiInforXML As String
    If txtPatient.Locked Then Exit Sub
    
    mbln���� = True '�����:56599
    If objCard.���� Like "IC��*" And objCard.ϵͳ Then
        'ϵͳIC��
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
    
    lng�����ID = objCard.�ӿ����
    
    If lng�����ID <= 0 Then Exit Sub
    '    zlReadCard(frmMain As Object, _
    '    ByVal lngModule As Long, _
    '    ByVal lngCardTypeID As Long, _
    '    ByVal blnOlnyCardNO As Boolean, _
    '    ByVal strExpand As String, _
    '    ByRef strOutCardNO As String, _
    '    ByRef strOutPatiInforXML As String) As Boolean
    '    '---------------------------------------------------------------------------------------------------------------------------------------------
    '    '����:�����ӿ�
    '    '���:frmMain-���õĸ�����
    '    '       lngModule-���õ�ģ���
    '    '       strExpand-��չ����,������
    '    '       blnOlnyCardNO-������ȡ����
    '    '����:strOutCardNO-���صĿ���
    '    '       strOutPatiInforXML-(������Ϣ����.XML��)
    '    '����:��������    True:���óɹ�,False:����ʧ��\
'    If gobjSquare.objSquareCard.zlReadCard(Me, mlngModul, lng�����ID, True, strExpand, strOutCardNO, strOutPatiInforXML) = False Then Exit Sub
'    txtPatient.Text = strOutCardNO
    
'    If txtPatient.Text <> "" Then
'        mblnUnChange = True
'        Call txtPatient_Validate(False)
'        mblnUnChange = False
'    End If
    
End Sub

Private Sub IDKind_ItemClick(index As Integer, objCard As zlIDKind.Card)
    If mbytInState > 0 Then Exit Sub
    '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
    txtPatient.IMEMode = 0
    '��Ҫ�����Ϣ,����ˢ����,���л�,���������ʾʧȥ����
    If txtPatient.Text <> "" And Not mblnNotClick Then txtPatient.Text = ""
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    
    zlControl.TxtSelAll txtPatient
    '83089:���ϴ�,2015/3/17,����ȱʡ�ķ������
    If IDKind.GetCurCard.���� Like "����*" Then
        Call InitSendCardPreperty(mlngModul)
    End If
End Sub

Private Sub IDKind_KeyDown(KeyCode As Integer, Shift As Integer)
    If txtPatient.Locked Or txtPatient.Enabled = False Or txtPatient.Text <> "" Then Exit Sub
    If IDKind.GetCurCard Is Nothing Then Exit Sub
    '�������IDKind
    IDKind.ActiveFastKey
End Sub

Private Sub IDKind_ReadCard(ByVal objCard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
    Dim lngPreIDKind As Long, intIndex As Integer
    Dim dtDate As Date
    Dim blnNew As Boolean
    Dim blnCard As Boolean    '�Ƿ���￨

    If txtPatient.Locked Or txtPatient.Text <> "" Then Exit Sub    'Or Not Me.ActiveControl Is txtPatient
    '״̬������ֵ
    mblnNotClick = True
    intIndex = IDKind.GetKindIndex(objCard.����)
    lngPreIDKind = IDKind.IDKind
    If intIndex > 0 Then IDKind.IDKind = intIndex
    txtPatient.Text = objPatiInfor.����
    Call txtPatient_Validate(False)
    
    If mrsInfo Is Nothing Then
        blnNew = True
    ElseIf mrsInfo.State <> 1 Then
        blnNew = True
    End If
    '�����²���
    If (txtPatient.Text = "" Or blnNew) And objPatiInfor.���� <> "" Then
        txtPatient.Text = objPatiInfor.����
        intIndex = IDKind.GetKindIndex("����")
        If intIndex > 0 Then IDKind.IDKind = IDKind.GetKindIndex("����")
        Call txtPatient_KeyPress(vbKeyReturn)
        If txtPatient.Text <> "" Then
            Call zlControl.CboLocate(cbo�Ա�, objPatiInfor.�Ա�)
            If IsDate(objPatiInfor.��������) = False Then
                txt����.Text = ReCalcOld(CDate(objPatiInfor.��������), cbo���䵥λ)
            End If
        End If
    End If
    IDKind.IDKind = lngPreIDKind
    mblnNotClick = False
End Sub

Private Sub MovePatiPic()
    '----------------------------------------------------------------------------------------------------------------
    '���ܣ��ƶ��������
    '���ƣ�Ƚ����
    '���ڣ�2014-7-8
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

Private Sub IDKind֤��_ItemClick(index As Integer, objCard As zlIDKind.Card)
    Dim blnVisible As Boolean, lngRow As Long, lngCol As Long
    If IDKind֤��.IDKind = IDKind֤��.GetKindIndex("���֤��") Then blnVisible = True
    If blnVisible And txtPatient = "" Then txtIDCard.Tag = "": txtIDCard.Text = ""
    txtIDCard.Visible = blnVisible:  txt֤��.Visible = Not blnVisible
    If txtIDCard.Visible And txtIDCard.Enabled Then txtIDCard.SetFocus
    If txt֤��.Visible And txt֤��.Enabled Then txt֤��.SetFocus
    txt֤��.Text = "": txt֤��.Tag = ""
    If blnVisible Then Exit Sub
    '105357:���ϴ���2017/2/6�������ʼ��ʱ�ᴥ��ItemClick
    If mobjfrmPatiInfo Is Nothing Then Exit Sub
    With mobjfrmPatiInfo.vsCertificate
        For lngRow = 1 To .Rows - 1
            For lngCol = 0 To .Cols - 1 Step 2
                If .TextMatrix(lngRow, lngCol) = IDKind֤��.GetCurCard.���� Then
                    txt֤��.Tag = .TextMatrix(lngRow, lngCol + 1)
                    txt֤��.Text = txt֤��.Tag
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

'72168,Ƚ����,2014/4/22,�Һ�ʱͨ���Һſ���ȷ����ѡ�ѱ�
Private Sub mobjfrmPatiInfo_ReturnVisitClick()
    Dim i As Long
    
    Call Init�ѱ�(mobjfrmPatiInfo.chk����.Value = 0, True)
    With mobjfrmPatiInfo
        .cbo�ѱ�.Clear
        For i = 0 To cbo�ѱ�.ListCount - 1
            .cbo�ѱ�.AddItem cbo�ѱ�.List(i)
            .cbo�ѱ�.ItemData(i) = cbo�ѱ�.ItemData(i)
        Next
        .cbo�ѱ�.ListIndex = cbo�ѱ�.ListIndex
    End With
End Sub

Private Sub mobjfrmPatiInfo_PatiMerged(����ID As Long)
        '�ϲ���Ĳ���
        Call GetPatient(IDKind.GetCurCard, "-" & ����ID, False)
End Sub

Private Sub mobjfrmPatiInfo_���ʽClick(index As Long)
    cbo���ʽ.ListIndex = index
End Sub

Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, _
                            ByVal strNation As String, ByVal datBirthDay As Date, ByVal strAddress As String)
    Dim lngPreIDKind As Long
    Dim blnNewCard   As Boolean
    Dim blnAddCardItem  As Boolean
    
    If txt�ű�.Text <> "" And Not txtPatient.Locked And txtPatient.Text = "" And Me.ActiveControl Is txtPatient Then
        mblnNotClick = True
        lngPreIDKind = IDKind.IDKind
        IDKind.IDKind = IDKind.GetKindIndex("���֤��")
        txtPatient.Text = strID
        mblnNotEMPIQuery = True
        mblnUserCancel = False
        Call txtPatient_Validate(False)
        mblnNotEMPIQuery = False
        '107049:���ϴ�,2017/4/14,���his�м�¼����his��Ϣ�����ӿ�
        If Not mrsInfo Is Nothing Then Call zlQueryEMPIPatiInfo
        
        If txtPatient.Text = "" And mblnUserCancel = True Then mblnNotClick = False: Exit Sub
        
        If txtPatient.Text = "" Then   '�²���
            IDKind.IDKind = IDKind.GetKindIndex("����")
            txtPatient.Text = strName
            '107049:���ϴ�,2017/4/14,Ϊ�˽����֤�ϵ���Ϣ�����ӿ�
            mblnNotEMPIQuery = True
            Call txtPatient_Validate(False)
            If txtPatient.Text <> "" Then
                txtIDCard.Text = strID
                txtIDCard.Tag = strID
                With mobjfrmPatiInfo
                    .txt���֤��.Text = strID
                    Call zlControl.CboLocate(.cbo�Ա�, strSex)
                    Call zlControl.CboLocate(.cbo����, strNation)
                    .txt��������.Text = Format(datBirthDay, "yyyy-MM-dd")
                    .txt����ʱ��.Text = "00:00"
                    txt��������.Text = Format(datBirthDay, "yyyy-MM-dd")
                    txt����ʱ��.Text = "00:00"
                    .cbo��ͥ��ַ.Text = IIf(Trim(cbo��ͥ��ַ.Text) = "", strAddress, cbo��ͥ��ַ.Text)
                    .txtRegLocation.Text = strAddress
                    cbo���ڵ�ַ.Text = .txtRegLocation.Text
                    cbo�Ա�.ListIndex = .cbo�Ա�.ListIndex
                    txt����.Text = .txt����.Text
                    txt����.Tag = .txt����.Text '38564
                    
                    cbo���䵥λ.ListIndex = .cbo���䵥λ.ListIndex
                    Call txt����_Validate(False)
                    cbo��ͥ��ַ.Text = .cbo��ͥ��ַ.Text
                    '89242:���ϴ�,2015/12/7,��ȡ���˵�ַ��Ϣ
                    padd��ͥ��ַ.Value = cbo��ͥ��ַ.Text
                    padd���ڵ�ַ.Value = cbo���ڵ�ַ.Text
                    .padd��ͥ��ַ.Value = cbo��ͥ��ַ.Text
                    .padd���ڵ�ַ.Value = cbo���ڵ�ַ.Text
                    .cbo���䵥λ.Tag = .cbo���䵥λ.Text
                    cbo���䵥λ.Tag = cbo���䵥λ.Text
                End With
            End If
            mblnNotEMPIQuery = False
            Call zlQueryEMPIPatiInfo
        End If
        IDKind.IDKind = lngPreIDKind
        mblnNotClick = False
        mobjfrmPatiInfo.mblnNewPatient = False
        '75717,Ƚ����,2014-7-22,�Һ�ԤԼʱ��ȡ�²������֤��Ƭ
        If mobjfrmPatiInfo.imgPatient.Picture = 0 Then
            Call LoadIDImage
        End If
        If cbo���ڵ�ַ.Text = "" Then
            mobjfrmPatiInfo.txtRegLocation.Text = strAddress
            cbo���ڵ�ַ.Text = strAddress
            padd���ڵ�ַ.Value = cbo���ڵ�ַ.Text
            mobjfrmPatiInfo.padd���ڵ�ַ.Value = cbo���ڵ�ַ.Text
        Else
            If mblnStructAdress Then
                If padd���ڵ�ַ.CheckDefrentValue(padd���ڵ�ַ.Value, strAddress) = False Then
                    If MsgBox("���֤�ϵĵ�ַ" & strAddress & "��ԭ�в��˵Ļ��ڵ�ַ" & padd���ڵ�ַ.Value & "��һ��,�Ƿ񽫲��˵Ļ��ڵ�ַ����Ϊ���֤�ϵĵ�ַ?", vbQuestion + vbYesNo, gstrSysName) = vbYes Then
                        mobjfrmPatiInfo.txtRegLocation.Text = strAddress
                        cbo���ڵ�ַ.Text = strAddress
                        padd���ڵ�ַ.Value = cbo���ڵ�ַ.Text
                        mobjfrmPatiInfo.padd���ڵ�ַ.Value = cbo���ڵ�ַ.Text
                    End If
                End If
            Else
                If cbo���ڵ�ַ.Text <> strAddress Then
                    If MsgBox("���֤�ϵĵ�ַ" & strAddress & "��ԭ�в��˵Ļ��ڵ�ַ" & cbo���ڵ�ַ.Text & "��һ��,�Ƿ񽫲��˵Ļ��ڵ�ַ����Ϊ���֤�ϵĵ�ַ?", vbQuestion + vbYesNo, gstrSysName) = vbYes Then
                        mobjfrmPatiInfo.txtRegLocation.Text = strAddress
                        cbo���ڵ�ַ.Text = strAddress
                        padd���ڵ�ַ.Value = cbo���ڵ�ַ.Text
                        mobjfrmPatiInfo.padd���ڵ�ַ.Value = cbo���ڵ�ַ.Text
                    End If
                End If
            End If
        End If
        'û�м�ͥ��ַ��,���¼�ͥ��ַ
        If cbo��ͥ��ַ.Text = "" Then
            mobjfrmPatiInfo.cbo��ͥ��ַ.Text = strAddress
            cbo��ͥ��ַ.Text = strAddress
            padd��ͥ��ַ.Value = cbo��ͥ��ַ.Text
            mobjfrmPatiInfo.padd��ͥ��ַ.Value = cbo��ͥ��ַ.Text
        End If
    End If
End Sub


Private Sub mobjICCard_ShowICCardInfo(ByVal strNO As String)
    Dim lngPreIDKind As Long
    
    If txt�ű�.Text <> "" And Not txtPatient.Locked And txtPatient.Text = "" And Me.ActiveControl Is txtPatient Then
        mblnNotClick = True
        lngPreIDKind = IDKind.IDKind
        IDKind.IDKind = IDKind.GetKindIndex("IC����")
        txtPatient.Text = strNO
        If txtPatient.Text <> "" Then
            mblnUnChange = True
            Call txtPatient_Validate(False)
            mblnUnChange = False
            Call SetOneCardBalance
        Else
            Call mobjICCard.SetEnabled(False) '��������Ϸ������������ü����Զ���ȡ
        End If
        IDKind.IDKind = lngPreIDKind
        mblnNotClick = False
        If mobjICCard Is Nothing Then Call NewCardObject
        If txt�ű�.Text <> "" And Not txtPatient.Locked And Me.ActiveControl Is txtPatient Then mobjICCard.SetEnabled (txtPatient.Text = "")
    End If
End Sub

Private Sub cbo�ѱ�_Click()
    Dim str�ѱ� As String
    
    If mbytInState = 1 Or Not Visible Then Exit Sub
    '31182:����ԤԼ����
    If (mbytMode = 2 Or mbytMode = 0 And mbytInState = 0 And mstrNoIn <> "") And InStr(1, mstrPrivs, ";�����޸ķѱ�;") <= 0 Then
        'ԤԼ����
        If mTy_Para.blnԤԼ����ȷ���Һŷ� = False Then
            If Not mrsInfo Is Nothing Then
                Exit Sub
            End If
        End If
    End If
   ' If Not (mbytMode = 2 Or mbytMode = 0 And mbytInState = 0 And mstrNoIn <> "") And mTy_Para.blnԤԼ����ȷ���Һŷ� = False And Not (mrsInfo Is Nothing And mbytMode = 2) Then Exit Sub
    
    str�ѱ� = NeedName(cbo�ѱ�)
    If mstrPre�ѱ� = str�ѱ� Then Exit Sub
    mstrPre�ѱ� = str�ѱ�
    
    If txt�ű�.Text <> "" Then
        mblnBuyHisBook = True
        Call ShowRegistFromInput
        mblnBuyHisBook = False
    End If
End Sub



Private Sub cbo���䵥λ_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtPatientPrint.Visible Then
            If cmdOK.Enabled And cmdOK.Visible Then cmdOK.SetFocus
        Else
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    End If
End Sub

Private Sub cboҽ�����_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If cboҽ�����.Locked Then Exit Sub
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        lngIdx = zlControl.CboMatchIndex(cboҽ�����.Hwnd, KeyAscii)
        If lngIdx = -1 And cboҽ�����.ListCount > 0 Then lngIdx = 0
        cboҽ�����.ListIndex = lngIdx
    End If
End Sub

Private Sub cboҽ��_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub cboҽ��_Validate(Cancel As Boolean)
    Dim i As Integer
    Dim strDoctor As String
    Dim blnFinded As Boolean
    
    If cboҽ��.ListCount = 0 Then cboҽ��.Text = "": Exit Sub
    
    strDoctor = cboҽ��.Text
    
    If mrsDoctor.State = 1 Then
        If mrsDoctor.RecordCount = 0 Then cboҽ��.Text = "": Exit Sub
        mrsDoctor.MoveFirst
        For i = 1 To mrsDoctor.RecordCount
            If strDoctor = mrsDoctor!��� Or strDoctor = mrsDoctor!���� Or UCase(strDoctor) = mrsDoctor!���� Or strDoctor = mrsDoctor!���� & "-" & mrsDoctor!���� Then
                strDoctor = mrsDoctor!ID
                blnFinded = True
                Exit For
            End If
            mrsDoctor.MoveNext
        Next
        If Not blnFinded Then Call zlCommFun.PressKey(vbKeyF4)
    End If
        
    If blnFinded Then
        If zlControl.CboLocate(cboҽ��, strDoctor, True) Then
            mstrҽ������ = Mid(cboҽ��.Text, InStr(1, cboҽ��.Text, "-") + 1)
            mlngҽ��ID = cboҽ��.ItemData(cboҽ��.ListIndex)
        Else
            Call zlControl.TxtSelAll(cboҽ��)
            Cancel = True
        End If
    Else
        Call zlControl.TxtSelAll(cboҽ��)
        Cancel = mrsDoctor.State = 1
    End If
End Sub

Private Sub chkShowAll_Click()
    If mblnUnChkClick = True Or mblnReadBooking Then Exit Sub
    Call ShowPlans
End Sub

Private Sub chk������_GotFocus()
    chk������.ForeColor = vbBlue
End Sub

Private Sub chk������_LostFocus()
    chk������.ForeColor = &H80000012
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
    txt�ű�.Enabled = blnEnable
    txt����.Enabled = blnEnable
    txtSN.Enabled = blnEnable
    cboҽ��.Enabled = blnEnable
End Sub

Private Sub SetPatiEnable(ByVal blnEnable As Boolean)
    IDKind.Enabled = blnEnable
    txtPatient.Enabled = blnEnable
    cmdLookup.Enabled = blnEnable
    cmdCard.Enabled = blnEnable
    cmdMore.Enabled = blnEnable
    cmdComminuty.Enabled = blnEnable
    cmdYb.Enabled = blnEnable
    cbo�Ա�.Enabled = blnEnable
    txt����.Enabled = blnEnable
    cbo���䵥λ.Enabled = blnEnable
    txt��������.Enabled = blnEnable
    txt����ʱ��.Enabled = blnEnable
    txtIDCard.Enabled = blnEnable
    txt��ͥ�绰.Enabled = blnEnable
    cbo��ͥ��ַ.Enabled = blnEnable
    cbo���ڵ�ַ.Enabled = blnEnable
    padd���ڵ�ַ.Enabled = blnEnable
    padd��ͥ��ַ.Enabled = blnEnable
    cboҽ�����.Enabled = blnEnable
    cbo�ѱ�.Enabled = blnEnable
    cbo���ʽ.Enabled = blnEnable
    txt�����.Enabled = blnEnable
    IDKind֤��.Enabled = blnEnable
End Sub

Private Sub chkCancel_Click()
    cboNO.Text = ""
    
    SetCodeEnable chkCancel.Value = 0
    SetPatiEnable chkCancel.Value = 0
    vsfPlan.Enabled = chkCancel.Value = 0
    
    Call RemoveShowItem
    Call ClearBill
    
    mcur�ϼ� = 0: mcurӦ�� = 0: txt�ϼ�.Text = "0.00": txt����Ӧ��.Text = "0.00": mint�Һ��� = 0
    txt�ɿ�.Text = "0.00": txt�ɿ�.Enabled = chkCancel.Value = 0
    txt�Ҳ�.Text = "0.00": txt�Ҳ�.Enabled = chkCancel.Value = 0
        
    Call SetCHKState(chkCancel)
    
    If chkCancel.Value = 0 Then
        chkCancel.ForeColor = 0
        lbl��.Visible = False
        txtFact.Locked = False
        txt�ű�.Locked = False
        
        txtPatient.Locked = False
        txt����.Locked = False
        cbo��ͥ��ַ.Locked = False
        cbo���ڵ�ַ.Locked = False
        padd��ͥ��ַ.ControlLock = False
        padd���ڵ�ַ.ControlLock = False
        txt�����.Locked = False
        
        cbo�Ա�.Locked = False
        cbo���ʽ.Locked = False
        cbo�ѱ�.Locked = False
        txtIDCard.Locked = False
        cbo���㷽ʽ.Locked = False
        
        chk������.Enabled = mbln������
        chk������.Caption = "������"
        chkExtra.Visible = False
        lblԤԼ��ʽ.Visible = mbytMode <> 0
        cboԤԼ��ʽ.Visible = mbytMode <> 0
        'ˢ��Ʊ�ݺ�
        If mbytMode <> 1 And gbytInvoice <> 0 Then Call RefreshFact
        If mbytMode <> 1 Then Load֧����ʽ
    Else
        chkCancel.ForeColor = vbRed
        
        lbl��.Visible = False
                
        txtFact.Locked = Not (InStr(1, mstrPrivs, ";�޸�Ʊ�ݺ�;") > 0) And gblnBill�Һ�  ' True:���˺�:20000,�����޸�Ʊ�ݺ�Ȩ��
        txt�ű�.Locked = True
        
        txtPatient.Locked = True
        txt����.Locked = True
        cbo��ͥ��ַ.Locked = True
        cbo���ڵ�ַ.Locked = True
        padd��ͥ��ַ.ControlLock = True
        padd���ڵ�ַ.ControlLock = True
        txt�����.Locked = True
        txtIDCard.Locked = True
        cbo�Ա�.Locked = True
        cbo���ʽ.Locked = True
        cbo�ѱ�.Locked = True
        cbo���㷽ʽ.Visible = False
        
        chk������.Enabled = False
        chk������.Caption = "�˲�����"
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
    
    mcur�ϼ� = 0: mcurӦ�� = 0: txt�ϼ�.Text = "0.00": txt����Ӧ��.Text = "0.00": mint�Һ��� = 0
    txt�ɿ�.Text = "0.00": txt�ɿ�.Enabled = chkPrint.Value = 0
    txt�Ҳ�.Text = "0.00": txt�Ҳ�.Enabled = chkPrint.Value = 0
        
    Call SetCHKState(chkPrint)
    
    If txtPatientPrint.Visible Then
        txtPatientPrint.Text = ""
        txtPatientPrint.Visible = False
        txtPatientPrint.Locked = False
        Call SetRePrintPatiEnabled(True)
    End If
    
    If chkPrint.Value = 0 Then
        chkPrint.ForeColor = 0
                                
        lbl��.Visible = False
        
        txtFact.Locked = False
        txt�ű�.Locked = False
        
        txtPatient.Locked = False
        txt����.Locked = False
        cbo��ͥ��ַ.Locked = False
        cbo���ڵ�ַ.Locked = False
        padd��ͥ��ַ.ControlLock = False
        padd���ڵ�ַ.ControlLock = False
        txt�����.Locked = False
        cbo�Ա�.Locked = False
        cbo���ʽ.Locked = False
        cbo�ѱ�.Locked = False
        cbo���㷽ʽ.Locked = False
        
        chk������.Enabled = mbln������
        '74017:���ϴ���2014-6-17���˳��Һ��ش�ʱ���ָ�cmdCard��״̬
        cmdCard.Enabled = True
        'ˢ��Ʊ�ݺ�
        If mbytMode <> 1 And gbytInvoice <> 0 Then Call RefreshFact
    Else
        chkPrint.ForeColor = vbBlue
                
        lbl��.Visible = False
                
        txtFact.Locked = Not (InStr(1, mstrPrivs, ";�޸�Ʊ�ݺ�;") > 0) And gblnBill�Һ�  'True:���˺�:20000,�����޸�Ʊ�ݺ�Ȩ��
        txt�ű�.Locked = True
        
        If InStr(1, mstrPrivs, ";�޸������ش�;") > 0 Then
            txtPatientPrint.Width = txtPatient.Width
            txtPatientPrint.Visible = True
        End If
        
        txtPatient.Locked = True
        txt����.Locked = True
        cbo��ͥ��ַ.Locked = True
        cbo���ڵ�ַ.Locked = True
        padd��ͥ��ַ.ControlLock = True
        padd���ڵ�ַ.ControlLock = True
        txt�����.Locked = True
        cbo�Ա�.Locked = True
        cbo���ʽ.Locked = True
        cbo�ѱ�.Locked = True
        cbo���㷽ʽ.Locked = True
        
        chk������.Enabled = False
                
        cboNO.Text = "": txtFact.Text = ""
        
        If cboNO.Enabled And cboNO.Visible Then cboNO.SetFocus
    End If
    Call AdjustInfoPosition
End Sub

Private Sub chk������_Click()
    If Not mbln������ And mbytInState = 0 Then
        chk������.Value = 0: Exit Sub
    End If
    
    '�˺�
    If mbytInState = 1 And mbytMode = 4 Or chkCancel.Value = 1 Then
        If mblnNotClick Then Exit Sub
        Call IsCheckBackExtra(True)
        Exit Sub
    End If
    '31182:
    If (mbytMode = 0 And mbytInState = 0 And mstrNoIn <> "") And mTy_Para.blnԤԼ����ȷ���Һŷ� = False Then Exit Sub
    
    If Not chk������.Enabled Then Exit Sub
    If mblnNotClick Then Exit Sub
    mblnBuyHisBook = True
    Call ShowRegistFromInput
    mblnBuyHisBook = False
End Sub

Private Sub chkExtra_Click()
    If Not mbln������ And mbytInState = 0 Then
        chk������.Value = 0: Exit Sub
    End If
    
    '�˺�
    If Not (mbytInState = 1 And mbytMode = 4 Or chkCancel.Value = 1) Then Exit Sub
    If mblnNotClick Then Exit Sub
    Call IsCheckBackExtra
End Sub

Private Function IsCheckBackExtra(Optional ByVal bln������ As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�˺�ʱ��鸽����Ŀ�Ƿ�����ֿ���
    '���:bln������-��鲡����
    '����:�ɹ�����true,���򷵻�False
    '����:���ϴ�
    '����:2018/5/2 11:35:08
    '����:123874
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strFeeItem As String
    Dim curMoney As Currency, curTotal As Currency, curDiff As Currency
    Dim curAdvance As Currency 'Ԥ���Ľɿ�
    Dim curInsure As Currency
    Dim curCash As Currency
    Dim i As Long
    Dim strFilter As String
    Dim strItem() As String
    If Not (mbytInState = 1 And mbytMode = 4 Or chkCancel.Value = 1) Then IsCheckBackExtra = True: Exit Function
    strFeeItem = IIf(bln������, "������", "���ӷ�")
    If Not mrsBillAdvance Is Nothing Then
        mrsBillAdvance.Filter = 0
        If mrsBillAdvance.RecordCount > 0 Then mrsBillAdvance.MoveFirst
        Do While Not mrsBillAdvance.EOF
            If InStr(",7,8,", "," & mrsBillAdvance!���� & ",") > 0 And (mrsBillAdvance!��¼���� <> 1 And mrsBillAdvance!��¼���� <> 11) Then
                MsgBox "ʹ�������ӿڽ���ĹҺŵ���,���ܽ�" & strFeeItem & "��Һŷѷֿ���!", vbInformation, gstrSysName
                mblnNotClick = True
                If bln������ Then
                    chk������.Value = 1
                Else
                    chkExtra.Value = 1
                End If
                mblnNotClick = False
                Exit Function
            End If
            If InStr(",3,", "," & mrsBillAdvance!���� & ",") > 0 And (MCPAR.���ղ����� = False Or Not bln������) Then
                MsgBox "ҽ�������˻���ȡ" & strFeeItem & "ʱ,��֧��" & strFeeItem & "��Һŷѷֱ���!", vbInformation, gstrSysName
                mblnNotClick = True
                If bln������ Then
                    chk������.Value = 1
                Else
                    chkExtra.Value = 1
                End If
                mblnNotClick = False
                Exit Function
            End If
            mrsBillAdvance.MoveNext
        Loop
    End If
    '�����ʱ�������¼�,��ʾ����
    If mrsBill Is Nothing Then IsCheckBackExtra = True: Exit Function
    If mstr������ĿID <> "" Then
        strFilter = ""
        strItem = Split(mstr������ĿID, ",")
        For i = 0 To UBound(strItem)
            If strFilter = "" Then
                strFilter = "�շ�ϸĿID <> " & strItem(i)
            Else
                strFilter = strFilter & " And �շ�ϸĿID <> " & strItem(i)
            End If
        Next i
    End If
    
    '��ȡ���ܽ��
    mrsBill.Filter = 0
    If mrsBill.RecordCount > 0 Then mrsBill.MoveFirst
    For i = 1 To mrsBill.RecordCount
        curTotal = curTotal + mrsBill!ʵ��
        mrsBill.MoveNext
    Next
    
    '��ȡ��ѡ��Ľ�����Ŀ.�п����ǻָ�,����Ӱ��
    If chkExtra.Value = 0 And strFilter <> "" Then
        If chk������.Value = 1 Then
            mrsBill.Filter = strFilter
        Else
            mrsBill.Filter = "���ӱ�־<>1 And " & strFilter
        End If
    Else
        If chk������.Value = 1 Then
            mrsBill.Filter = 0
        Else
            mrsBill.Filter = "���ӱ�־<>1"
        End If
    End If
    If mrsBill.RecordCount > 0 Then mrsBill.MoveFirst
    vsfMoney.Rows = mrsBill.RecordCount + 1
    For i = 1 To mrsBill.RecordCount
        vsfMoney.TextMatrix(i, 0) = mrsBill!��Ŀ
        vsfMoney.TextMatrix(i, 1) = Format(mrsBill!Ӧ��, "0.00")
        vsfMoney.TextMatrix(i, 2) = Format(mrsBill!ʵ��, "0.00")
        curMoney = curMoney + mrsBill!ʵ��
        mrsBill.MoveNext
    Next
    txt�ϼ�.Text = Format(curMoney, "0.00")
    mrsBill.Filter = 0: If mrsBill.RecordCount > 0 Then mrsBill.MoveFirst
    
    'ȡ���,Ȼ���ٵ���
    curDiff = curTotal - curMoney
    Call Load������Ϣ(Val(curMoney), Val(curDiff))
    Set�����Һ�
    IsCheckBackExtra = True
End Function

Private Sub RecalPay()
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim dblTotal As Double, rsTx As ADODB.Recordset
    On Error GoTo errH
    vsfPay.Rows = 1
    vsfPay.Clear 1
    vsfPay.Rows = 2
    dblTotal = Val(txt�ϼ�.Text)
    strSQL = "Select b.���� As ���㷽ʽ, a.��Ԥ��, b.����" & vbNewLine & _
            "From ����Ԥ����¼ A, ���㷽ʽ B" & vbNewLine & _
            "Where a.����id = [1] And a.���㷽ʽ = b.���� And a.��¼���� = 4" & vbNewLine & _
            "Union" & vbNewLine & _
            "Select 'Ԥ����' As ���㷽ʽ, a.��Ԥ��, 0 As ���� From ����Ԥ����¼ A Where a.����id = [1] And Mod(a.��¼����, 10) = 1 Order By ����"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID)
    rsTemp.Filter = "���� <> 1 And ���� <> 2"
    Do While Not rsTemp.EOF
        If dblTotal > 0 Then
            If dblTotal > Val(rsTemp!��Ԥ��) Then
                vsfPay.TextMatrix(vsfPay.Rows - 1, 0) = rsTemp!���㷽ʽ
                vsfPay.TextMatrix(vsfPay.Rows - 1, 1) = Format(Val(rsTemp!��Ԥ��), "0.00")
                vsfPay.RowData(vsfPay.Rows - 1) = Val(rsTemp!����)
                If Val(rsTemp!����) = 7 Or Val(rsTemp!����) = 8 Then
                    strSQL = "Select ID,�Ƿ����� From ҽ�ƿ���� Where ���㷽ʽ=[1]"
                    Set rsTx = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, rsTemp!���㷽ʽ)
                    If rsTx.EOF Then
                        strSQL = "Select ���,�Ƿ����� From ���ѿ����Ŀ¼ Where ���㷽ʽ=[1]"
                        Set rsTx = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, rsTemp!���㷽ʽ)
                        If rsTx.EOF Then
                            vsfPay.TextMatrix(vsfPay.Rows - 1, vsfPay.ColIndex("�޸�")) = "1"
                        Else
                            vsfPay.TextMatrix(vsfPay.Rows - 1, 4) = Nvl(rsTx!���)
                            vsfPay.TextMatrix(vsfPay.Rows - 1, vsfPay.ColIndex("�޸�")) = IIf(Val(rsTx!�Ƿ�����) = 1, "0", "1")
                        End If
                    Else
                        vsfPay.TextMatrix(vsfPay.Rows - 1, 4) = Nvl(rsTx!ID)
                        vsfPay.TextMatrix(vsfPay.Rows - 1, vsfPay.ColIndex("�޸�")) = IIf(Val(rsTx!�Ƿ�����) = 1, "0", "1")
                    End If
                End If
                If Val(rsTemp!����) = 0 Or Val(rsTemp!����) = 3 Then
                    vsfPay.TextMatrix(vsfPay.Rows - 1, vsfPay.ColIndex("�޸�")) = "1"
                End If
                vsfPay.Rows = vsfPay.Rows + 1
                dblTotal = dblTotal - Val(rsTemp!��Ԥ��)
            Else
                vsfPay.TextMatrix(vsfPay.Rows - 1, 0) = rsTemp!���㷽ʽ
                vsfPay.TextMatrix(vsfPay.Rows - 1, 1) = Format(dblTotal, "0.00")
                vsfPay.RowData(vsfPay.Rows - 1) = Val(rsTemp!����)
                If Val(rsTemp!����) = 7 Or Val(rsTemp!����) = 8 Then
                    strSQL = "Select ID,�Ƿ����� From ҽ�ƿ���� Where ���㷽ʽ=[1]"
                    Set rsTx = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, rsTemp!���㷽ʽ)
                    If rsTx.EOF Then
                        strSQL = "Select ���,�Ƿ����� From ���ѿ����Ŀ¼ Where ���㷽ʽ=[1]"
                        Set rsTx = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, rsTemp!���㷽ʽ)
                        If rsTx.EOF Then
                            vsfPay.TextMatrix(vsfPay.Rows - 1, vsfPay.ColIndex("�޸�")) = "1"
                        Else
                            vsfPay.TextMatrix(vsfPay.Rows - 1, 4) = Nvl(rsTx!���)
                            vsfPay.TextMatrix(vsfPay.Rows - 1, vsfPay.ColIndex("�޸�")) = IIf(Val(rsTx!�Ƿ�����) = 1, "0", "1")
                        End If
                    Else
                        vsfPay.TextMatrix(vsfPay.Rows - 1, 4) = Nvl(rsTx!ID)
                        vsfPay.TextMatrix(vsfPay.Rows - 1, vsfPay.ColIndex("�޸�")) = IIf(Val(rsTx!�Ƿ�����) = 1, "0", "1")
                    End If
                End If
                If Val(rsTemp!����) = 0 Or Val(rsTemp!����) = 3 Then
                    vsfPay.TextMatrix(vsfPay.Rows - 1, vsfPay.ColIndex("�޸�")) = "1"
                End If
                vsfPay.Rows = vsfPay.Rows + 1
                dblTotal = 0
            End If
        End If
        rsTemp.MoveNext
    Loop
    rsTemp.Filter = "���� = 1 Or ���� = 2"
    rsTemp.Sort = "���� Desc"
    Do While Not rsTemp.EOF
        If dblTotal > 0 Then
            If dblTotal > Val(rsTemp!��Ԥ��) Then
                vsfPay.TextMatrix(vsfPay.Rows - 1, 0) = rsTemp!���㷽ʽ
                vsfPay.TextMatrix(vsfPay.Rows - 1, 1) = Format(Val(rsTemp!��Ԥ��), "0.00")
                vsfPay.RowData(vsfPay.Rows - 1) = Val(rsTemp!����)
                vsfPay.TextMatrix(vsfPay.Rows - 1, vsfPay.ColIndex("�޸�")) = IIf(Val(rsTemp!����) = 1, "1", "0")
                vsfPay.Rows = vsfPay.Rows + 1
                dblTotal = dblTotal - Val(rsTemp!��Ԥ��)
            Else
                vsfPay.TextMatrix(vsfPay.Rows - 1, 0) = rsTemp!���㷽ʽ
                vsfPay.TextMatrix(vsfPay.Rows - 1, 1) = Format(dblTotal, "0.00")
                vsfPay.RowData(vsfPay.Rows - 1) = Val(rsTemp!����)
                vsfPay.TextMatrix(vsfPay.Rows - 1, vsfPay.ColIndex("�޸�")) = IIf(Val(rsTemp!����) = 1, "1", "0")
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
    
 
Private Sub chk������_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub cbo�ѱ�_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If cbo�ѱ�.Locked Then Exit Sub
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        lngIdx = zlControl.CboMatchIndex(cbo�ѱ�.Hwnd, KeyAscii)
        If lngIdx = -1 And cbo�ѱ�.ListCount > 0 Then lngIdx = 0
        cbo�ѱ�.ListIndex = lngIdx
    End If
End Sub

Private Sub cbo���㷽ʽ_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If cbo���㷽ʽ.Locked Then Exit Sub
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        lngIdx = zlControl.CboMatchIndex(cbo���㷽ʽ.Hwnd, KeyAscii)
        If lngIdx = -1 And cbo���㷽ʽ.ListCount > 0 Then lngIdx = 0
        cbo���㷽ʽ.ListIndex = lngIdx
    End If
End Sub

Private Sub cbo�Ա�_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    
    If cbo�Ա�.Locked Then Exit Sub
    
    If KeyAscii = 13 And cbo�Ա�.ListIndex <> -1 Then
        KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    End If
    
    Call SendMessage(cbo�Ա�.Hwnd, CB_GETDROPPEDSTATE, 0, 0)
    lngIdx = MatchIndex(cbo�Ա�.Hwnd, KeyAscii)
    If lngIdx <> -2 Then cbo�Ա�.ListIndex = lngIdx
    If cbo�Ա�.ListCount > 0 And cbo�Ա�.ListIndex = -1 Then cbo�Ա�.ListIndex = 0
End Sub

Private Sub cbo���ʽ_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        If cbo���ʽ.Locked Then Exit Sub
        
        lngIdx = zlControl.CboMatchIndex(cbo���ʽ.Hwnd, KeyAscii)
        If lngIdx = -1 And cbo���ʽ.ListCount > 0 Then lngIdx = 0
        cbo���ʽ.ListIndex = lngIdx
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
    ElseIf mbytMode = 0 And mbytInState = 0 And mstrNoIn <> "" Then '����ԤԼ
        Call ClearBill
        Call SetReceiveState(False)
        
    ElseIf mbytMode = 2 Or mbytInState = 1 Or (mbytInState = 0 And mrsItems Is Nothing) Then
        Unload Me
    Else
        mbln�����Һ� = False
        Call YBIdentifyCancel 'ȡ��ҽ�����������֤
        Call ClearBill
        
        'ˢ��Ʊ�ݺ�
        If mbytMode <> 1 And gbytInvoice <> 0 Then Call RefreshFact
    End If
    mblnCancel = False
End Sub

Private Sub picDetailFee_Resize()
    Err = 0: On Error Resume Next
    With picDetailFee
         
        cbo��ע.Top = .ScaleHeight - cbo��ע.Height - 50
        lblժҪ.Top = cbo��ע.Top + (cbo��ע.Height - lblժҪ.Height) \ 2
        
        txt����ʱ��.Top = IIf(cbo��ע.Visible, cbo��ע.Top - 20, .ScaleHeight - txt����ʱ��.Height - 50) - txt����ʱ��.Height
        lbl����ʱ��.Top = txt����ʱ��.Top + (txt����ʱ��.Height - lbl����ʱ��.Height) \ 2
        
        cboԤԼ��ʽ.Top = txt����ʱ��.Top
        lblԤԼ��ʽ.Top = lbl����ʱ��.Top
        
        chkExtra.Top = txt����ʱ��.Top + (txt����ʱ��.Height - chkExtra.Height) \ 2
        chk������.Top = chkExtra.Top
        vsfMoney.Height = IIf(txt����ʱ��.Top - vsfMoney.Top - 20 < 0, 0, txt����ʱ��.Top - vsfMoney.Top - 20)
        
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
    '����:���������Ϣ
    '���:blnClearPati-���������Ϣ
    '     blnClearFact-�����Ʊ��Ϣ
    '     blnClearInsure-���ҽ����Ϣ
    '����:
    '����:
    '����:���˺�
    '����:2009-12-02 10:32:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnIDKind As Boolean, strTemp As String, i As Integer
    
    Call SetShowBalance '68991
    blnIDKind = mblnIDCardKind
    txtSN.Text = ""
    mstrNoIn = ""
    mlng��¼ID = 0
    mlng���ż�¼ID = 0
    If mbytMode <> 1 Then
        If chkShowAll.Value = 1 Then chkShowAll.Value = 0
    End If
    lbl��.Visible = False
    If blnClearFact Then txtFact.Text = ""
    mblnNoClearPrompt = False
    txt�ű�.Text = ""                       '����Change�¼����غű��б�
    txt����.Text = ""
    cboҽ��.Clear
    txtIDCard.Text = ""
    mblnAppointmentChange = False
    txt֤��.Text = ""
    mstrForceNote = ""
    txt��ͥ�绰.Text = ""
    If mlngOutModeMC > 0 Then cboҽ�����.ListIndex = 0
    '69338,������,�Һ����ʱδ���������ƺ������Ϣ������
    mRegistFeeMode = EM_RG_����
    mPatiChargeMode = EM_�Ƚ��������
    mlng�Һſ���ID = 0
    mstrҽ������ = ""
    mblnViewOriginal = False
    mlngҽ��ID = 0
    mbln������ = False
'    txtժҪ.Text = ""
    cbo��ע.Text = ""
    mstrPreNO = ""
    mintCancel = 0
    mbln���ӷ� = False
    mstrPrePriceGrade = ""
    
    txt�ű�.Locked = False
    txt�ű�.Enabled = True
    If mbytMode <> 2 Then cbo�ѱ�.Locked = False: cbo�ѱ�.TabStop = gbln�ѱ�
    
    mstr����NO = ""
    If vsfMoney.Rows < 2 Then
        cmdOK.Visible = True
    Else
        If vsfMoney.RowData(1) = 0 Then
            cmdOK.Visible = True
        End If
    End If
    '�����:58843
    If blnClearPati Then Set mrsInfo = Nothing '������Ϣ���
    Set mobjDelCards = Nothing
    mstr���˼���IDs = ""
    
    Call SetPatiInfoEnabled(False, mrsInfo Is Nothing, Not blnClearPati) '���ݲ���,�����Ҫ��������,���ߺű𲻽�����,��������������
    
    mblnIDCardKind = False
    
    If blnClearPati Then
        Call ClearPatientInfo
        Call Init�ѱ�(True, False)
        Call SetCboDefault(cbo�ѱ�)
        Call ClearmobjfrmPatiInfoFace
    Else
        '54537:������,2014-02-27,ҽ�����˷ѱ�δ��յ�����
        If mintInsure <> 0 And mstrYBPati <> "" Then Call SetCboDefault(cbo�ѱ�)
        mblnICCard = False
        mblnAddCardItem = False
    End If
    
    If mblnNewCard Then
        mobjfrmPatiInfo.txt���� = ""
        mobjfrmPatiInfo.mstrCard = ""
        lblPrompt.Caption = ""
        gCurSendCard.lng�շ�ϸĿID = 0
        vsfPay.Height = 2220
        mblnNewCard = False
    End If
    
    'ҽ���Ķ�
    mlng����ID = 0
    
    If blnClearPati = False And blnClearInsure = False Then
        'ҽ������,�������Һ�ʱ��Ч
    Else
        mintInsure = 0
        mstrYBPati = ""
        txtPatient.ForeColor = Me.ForeColor
        mobjfrmPatiInfo.txtPatient.ForeColor = Me.ForeColor
        Call SetIdentifyLocked(False)
    End If
    
    cmdComminuty.Enabled = True
    mint���� = 0
    mstr������ = ""
    
    Call ShowMedicareInfo(blnClearPati = False And blnClearInsure = False)
    
    '�̶����Ԥ��֧����Ϣ
    Call ShowDeposit(False)

    If mblnReSetIDKind And txtPatient.Text = "" Then IDKind.IDKind = IDKind.GetKindIndex("�����")
    If blnIDKind And txtPatient.Text = "" Then IDKind.IDKind = IDKind.GetKindIndex("���֤��")
    mblnReSetIDKind = False
    mstr����� = "": txt�����.TabStop = True
    
    chk������.Enabled = False
    chk������.Value = 0
    chk������.Enabled = mbln������
    If blnClearPati And mbln������ Then
        If mbytMode = 0 Or mbytMode = 1 Then chk������.Value = IIf(zlDatabase.GetPara("Ĭ�Ϲ�����", glngSys, mlngModul, 0) = "1", 1, 0)
    End If
    
'    txtժҪ.Text = ""

    Call ClearCardMoney
    Set mrsItems = Nothing
    Set mrsInComes = Nothing
    Call ClearMoney
'    Call SetCboDefault(cbo���㷽ʽ)
    Call Load֧����ʽ
    
    If cboԤԼ��ʽ.Visible Then
        strTemp = zlDatabase.GetPara("ȱʡԤԼ��ʽ", glngSys, IIf(mblnStation, 1260, mlngModul), "")
        '�����:112838,����,2017/09/05,�����ֵ����δ�����κ�ԤԼ��ʽʱ�ᱨ��
        If cboԤԼ��ʽ.ListCount <> 0 Then
            For i = 0 To cboԤԼ��ʽ.ListCount - 1
                If Mid(cboԤԼ��ʽ.List(i), InStr(cboԤԼ��ʽ.List(i), ".") + 1) = strTemp Then
                    cboԤԼ��ʽ.ListIndex = i
                End If
            Next i
            If cboԤԼ��ʽ.ListIndex < 0 Then cboԤԼ��ʽ.ListIndex = 0
        End If
    End If
    
    If mbytMode = 0 And mbytInState = 0 And gblnLED Then
        zl9LedVoice.DisplayPatient ""
    End If
    If txt�ű�.Enabled And txt�ű�.Visible Then txt�ű�.SetFocus
End Sub

Private Sub ClearCardMoney()
    With mCurCardPay
        .lngҽ�ƿ����ID = 0
        .bln���ѿ� = False
        .str���㷽ʽ = ""
        .str���� = ""
        .strˢ������ = ""
        .strˢ������ = ""
        .dbl�ʻ���� = 0
        .Have�Һŷ� = False
        .Have���� = False
        Set .objCard = Nothing
    End With
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.Hwnd, Me.Name
End Sub

Private Function DefMainCommandBars() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ���˵���������
    '����:���óɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-01-25 15:29:24
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
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
        .ShowExpandButtonAlways = False
    End With
    cbsThis.EnableCustomization False
    cbsThis.ActiveMenuBar.ModifyStyle &H400000, 0 'ȥ���˵���ǰ׺
    '�˵�����
    cbsThis.ActiveMenuBar.Visible = False
    cbsThis.ActiveMenuBar.EnableDocking xtpFlagAlignTop Or xtpFlagHideWrap Or xtpFlagStretched
    
    '����������
    Set mcbrToolBar = cbsThis.Add("������", xtpBarTop)
    mcbrToolBar.ShowTextBelowIcons = False
    mcbrToolBar.ContextMenuPresent = False
    mcbrToolBar.EnableDocking xtpFlagStretched
    mcbrToolBar.ModifyStyle &H400000, 0 'ȥ���˵���ǰ׺
    With mcbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, 2605, "Ԥ��")
        Set cbrControl = .Add(xtpControlButton, 2604, "ȡ��Ԥ��")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ����")
        Set cbrControl = .Add(xtpControlButton, 3816, "��Ԥ��")
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��")
        Set cbrControl = .Add(xtpControlButton, 816, "�����")
        Set cbrControl = .Add(xtpControlButton, 4006, "��ʷ�Һ���Ϣ")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
    
    End With
    For Each cbrControl In mcbrToolBar.Controls
        cbrControl.Style = xtpButtonIconAndCaption
    Next
    
    '�����
    With cbsThis.KeyBindings

    End With
    
    '���ò����ò˵�
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
        .Options.UseSplitterTracker = True 'ʵʱ�϶�
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
    
    If vsfList.Rows = 0 Or mViewMode = V_��ͨ�ŷ�ʱ�� Or vsfList.Visible = False Then Exit Sub
    If mViewMode <> v_ר�Һŷ�ʱ�� Then
        lngSN = Val(vsfList.TextMatrix(vsfList.Row, vsfList.Col))
    Else
        lngSN = Val(Getʱ��(vsfList.Row, vsfList.Col, False))
    End If
    If lngSN > 0 Then
        blnCan = True
        If Not mrsSNState Is Nothing Then
            mrsSNState.Filter = "���=" & lngSN
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
            Case V_��ͨ��:
                datThis = dtpAppointmentDate.Value
            Case Else
                datThis = CDate(Format(dtpAppointmentDate.Value, "yyyy-MM-dd") & " " & Format(Getʱ��(vsfList.Row, vsfList.Col, True), "hh:mm:ss"))
            End Select
        Else
            If mViewMode <> v_ר�Һŷ�ʱ�� Then
                datThis = zlDatabase.Currentdate
            Else
                datThis = CDate(Format(zlDatabase.Currentdate, "yyyy-MM-dd") & " " & Format(Getʱ��(vsfList.Row, vsfList.Col, True), "hh:mm:ss"))
            End If
        End If
        If mViewMode <> v_ר�Һŷ�ʱ�� Then
            strSQL = "Zl_�Һ����״̬_Update('" & vsfPlan.TextMatrix(vsfPlan.Row, GetCol("�ű�")) & _
                  "',To_Date('" & Format(datThis, "yyyy-MM-dd") & "','YYYY-MM-DD')," & lngSN & _
                  ",3,'" & UserInfo.���� & "'," & IIf(mcbrToolBar.Controls.Find(xtpControlButton, 2605).Enabled, "1", "0") & ",Null," & vsfPlan.TextMatrix(vsfPlan.Row, GetCol("��¼ID")) & ")"
        Else
            strSQL = "Zl_�Һ����״̬_Update('" & vsfPlan.TextMatrix(vsfPlan.Row, GetCol("�ű�")) & _
                  "',To_Date('" & Format(datThis, "yyyy-MM-dd hh:mm:ss") & "','YYYY-MM-DD hh24:mi:ss')," & lngSN & _
                  ",3,'" & UserInfo.���� & "'," & IIf(mcbrToolBar.Controls.Find(xtpControlButton, 2605).Enabled, "1", "0") & ",Null," & vsfPlan.TextMatrix(vsfPlan.Row, GetCol("��¼ID")) & ")"
        End If
        
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        
        'ˢ��״̬
        txtSN.Text = ""
        Call vsfPlan_EnterCell
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub ClearPatientInfo()
'����:������������Ϣ
    Dim i As Integer
    If Not (mblnNewCard And gblnNewCardNoPop) Then mblnAddCardItem = False
    mblnICCard = False
    mstrPrePati = ""
    txtPatient.Text = ""
    txtPatient.IMEMode = 0
    Call ShowDeposit(False)
    lbl����.Caption = ""
    lbl����.Visible = False
    If mbytMode = 1 Then
        vsfPay.Visible = False
    Else
        vsfPay.Visible = True
    End If
    If mbytMode = 1 And mblnIDCardKind Then
        '31182
    Else
        txt����.Text = ""
        txt����.Tag = ""
        cbo���䵥λ.Tag = ""
        Call zlControl.CboLocate(cbo���䵥λ, "��")
        Call txt����_Validate(False)
        If gstr�Ա� <> "��" Then SetCboDefault cbo�Ա�
    End If
    mdblԤ����� = 0
    For i = 1 To vsfPay.Rows - 1
        If vsfPay.RowData(i) = 0 And vsfPay.TextMatrix(i, 0) <> "" Then
            vsfPay.TextMatrix(i, 6) = 0
        End If
    Next i
    mdbl������� = 0
    cbo��ͥ��ַ.Text = ""
    cbo���ڵ�ַ.Text = ""
    txtIDCard.Text = ""
    txtIDCard.Tag = ""
    txt֤��.Tag = "": txt֤��.Text = ""
    txt��ͥ�绰.Text = ""
    '89242:���ϴ�,2015/12/7,��ȡ���˵�ַ��Ϣ
    Call zlLoadDefaultAddr(padd��ͥ��ַ)
    Call zlLoadDefaultAddr(padd���ڵ�ַ)
    txt�����.Text = ""
    txt��������.Text = "____-__-__"
    txt����ʱ��.Text = "__:__"
    stbThis.Panels(2).Text = ""
    imgPatiPic.Picture = Nothing
    SetCboDefault cbo���ʽ
End Sub

Private Sub CopyCboTofrmPatiInfo()
    Dim i As Long
    
    With mobjfrmPatiInfo
        .cbo�Ա�.Clear
        For i = 0 To cbo�Ա�.ListCount - 1
            .cbo�Ա�.AddItem cbo�Ա�.List(i)
            .cbo�Ա�.ItemData(i) = cbo�Ա�.ItemData(i)
        Next
        .cbo���䵥λ.Clear
        For i = 0 To cbo���䵥λ.ListCount - 1
            .cbo���䵥λ.AddItem cbo���䵥λ.List(i)
            .cbo���䵥λ.ItemData(i) = cbo���䵥λ.ItemData(i)
        Next
        .cbo���ʽ.Clear
        For i = 0 To cbo���ʽ.ListCount - 1
            .cbo���ʽ.AddItem cbo���ʽ.List(i)
            .cbo���ʽ.ItemData(i) = cbo���ʽ.ItemData(i)
        Next
        .cbo�ѱ�.Clear
        For i = 0 To cbo�ѱ�.ListCount - 1
            .cbo�ѱ�.AddItem cbo�ѱ�.List(i)
            .cbo�ѱ�.ItemData(i) = cbo�ѱ�.ItemData(i)
        Next
    End With
End Sub

Private Sub CopyInfoTofrmPatiInfo()
    With mobjfrmPatiInfo
        .txtPatient.Text = txtPatient.Text: .txtPatient.MaxLength = txtPatient.MaxLength
        '74428�����ϴ���2014-7-8������������ɫ����
        .txtPatient.ForeColor = txtPatient.ForeColor
        If Not mrsInfo Is Nothing And (mbytMode = 2 Or mbytMode = 0 And mbytInState = 0 And mstrNoIn <> "") Then
            '31182:ֻ��ԤԼ�ҺŲŻ����
            .txt����.Tag = Val(Nvl(mrsInfo!����ID))
        Else
            .txt����.Tag = 0
        End If
        If Not mrsInfo Is Nothing Then
            .mlng����ID = Val(Nvl(mrsInfo!����ID))
        Else
            .mlng����ID = 0
        End If
        .cbo�Ա�.ListIndex = cbo�Ա�.ListIndex
        .cbo���䵥λ.ListIndex = cbo���䵥λ.ListIndex
        .cbo���䵥λ.Tag = .cbo���䵥λ.Text
        .txt����.Text = txt����.Text: .txt����.MaxLength = txt����.MaxLength
        .txt����.Tag = txt����.Text
        .cbo��ͥ��ַ.Text = cbo��ͥ��ַ.Text
        .txtRegLocation.Text = cbo���ڵ�ַ.Text
        '89242:���ϴ�,2015/12/7,��ȡ���˵�ַ��Ϣ
        Call .padd��ͥ��ַ.LoadStructAdress(padd��ͥ��ַ.valueʡ, padd��ͥ��ַ.value��, padd��ͥ��ַ.value����, padd��ͥ��ַ.value����, padd��ͥ��ַ.value��ϸ��ַ)
        Call .padd���ڵ�ַ.LoadStructAdress(padd���ڵ�ַ.valueʡ, padd���ڵ�ַ.value��, padd���ڵ�ַ.value����, padd���ڵ�ַ.value����, padd���ڵ�ַ.value��ϸ��ַ)
        .txt�����.Text = txt�����.Text: .txt�����.MaxLength = txt�����.MaxLength
        .cbo���ʽ.ListIndex = cbo���ʽ.ListIndex
        .txt��ͥ�绰.Text = txt��ͥ�绰.Text
        .cbo�ѱ�.ListIndex = cbo�ѱ�.ListIndex
        .cbo�ѱ�.Locked = cbo�ѱ�.Locked
        .cbo�ѱ�.TabStop = cbo�ѱ�.TabStop
        .txt��������.Tag = txt��������.Text
        .txt����ʱ��.Tag = txt����ʱ��.Text
        .txt��������.Text = txt��������.Text
        .txt����ʱ��.Text = txt����ʱ��.Text
        .txt���֤��.Text = txtIDCard.Text
        .txt���֤��.Tag = txtIDCard.Text
        .imgPatient.Picture = imgPatiPic.Picture
    End With
    
    Call CopyZJTofrmPatiInfo
End Sub

Private Sub CopyZJTofrmPatiInfo()
    Dim lngRow As Long, lngCol As Long, blnFind As Boolean
    '��֤����Ϣ��ֵ��֤���б��ж�Ӧ�Ŀ��������棬û�о��Զ�����
     '���֤������
    If IDKind֤��.IDKind = IDKind֤��.GetKindIndex("���֤��") Then Exit Sub
    With mobjfrmPatiInfo.vsCertificate
        For lngRow = 1 To .Rows - 1
            For lngCol = 0 To .Cols - 1 Step 2
                If .TextMatrix(lngRow, lngCol) = IDKind֤��.GetCurCard.���� Then
                    .TextMatrix(lngRow, lngCol + 1) = txt֤��.Text
                    blnFind = True
                    Exit For
                End If
            Next
        Next
        'û�ҵ��Զ����
        If Trim(txt֤��.Text) <> "" And Not blnFind Then
            blnFind = False '�Ƿ��ҵ��˿�λ���
            For lngRow = 1 To .Rows - 1
                For lngCol = 0 To .Cols - 1 Step 2
                    If .TextMatrix(lngRow, lngCol) = "" And .TextMatrix(lngRow, lngCol + 1) = "" Then
                        .TextMatrix(lngRow, lngCol) = IDKind֤��.GetCurCard.����
                        .TextMatrix(lngRow, lngCol + 1) = txt֤��.Text
                        blnFind = True: Exit For
                    End If
                Next
            Next
            
            If Not blnFind Then
                If lngCol = 2 Then
                    .TextMatrix(lngRow, lngCol) = IDKind֤��.GetCurCard.����
                    .TextMatrix(lngRow, lngCol + 1) = txt֤��.Text
                Else
                    .Rows = .Rows + 1
                    .TextMatrix(lngRow, 0) = IDKind֤��.GetCurCard.����
                    .TextMatrix(lngRow, 1) = txt֤��.Text
                End If
            End If
        End If
    End With
End Sub

Private Sub CopyInfoFromobjfrmPatiInfo()
    Dim lngRow As Long, lngCol As Long
    
    With mobjfrmPatiInfo
        txtPatient.Text = .txtPatient.Text  '����Change�¼�
        '74428�����ϴ���2014-7-8������������ɫ����
        txtPatient.ForeColor = .txtPatient.ForeColor
        mstrPrePati = txtPatient.Text
        cbo�Ա�.ListIndex = .cbo�Ա�.ListIndex
        txt����.Text = .txt����.Text
        txt����.Tag = txt����.Text
        txt��ͥ�绰.Text = .txt��ͥ�绰.Text
        cbo���䵥λ.ListIndex = .cbo���䵥λ.ListIndex
        txt��������.Text = .txt��������.Text
        txt����ʱ��.Text = .txt����ʱ��.Text
        Call txt����_Validate(False)
        
        cbo��ͥ��ַ.Text = .cbo��ͥ��ַ.Text
        cbo���ڵ�ַ.Text = .txtRegLocation.Text
        '89242:���ϴ�,2015/12/7,��ȡ���˵�ַ��Ϣ
        Call padd��ͥ��ַ.LoadStructAdress(.padd��ͥ��ַ.valueʡ, .padd��ͥ��ַ.value��, .padd��ͥ��ַ.value����, .padd��ͥ��ַ.value����, .padd��ͥ��ַ.value��ϸ��ַ)
        Call padd���ڵ�ַ.LoadStructAdress(.padd���ڵ�ַ.valueʡ, .padd���ڵ�ַ.value��, .padd���ڵ�ַ.value����, .padd���ڵ�ַ.value����, .padd���ڵ�ַ.value��ϸ��ַ)
        txt�����.Text = .txt�����.Text
        cbo���ʽ.ListIndex = .cbo���ʽ.ListIndex
        cbo�ѱ�.ListIndex = .cbo�ѱ�.ListIndex
        cbo���䵥λ.Tag = cbo���䵥λ.Text
        txtIDCard.Tag = .txt���֤��.Text
        txtIDCard.Text = .txt���֤��.Text
        imgPatiPic.Picture = .imgPatient.Picture
        
        If Trim(.txtPatiMCNO(0).Text) <> "" Then Call SetCboDefault(cboҽ�����)
    End With
    '90875:���ϴ�,2016/11/8,ҽ�ƿ�֤������
    '��֤���б����ҵ���ǰ�����ͺͿ���
    '���֤������
    If IDKind֤��.IDKind = IDKind֤��.GetKindIndex("���֤��") Then Exit Sub
    With mobjfrmPatiInfo.vsCertificate
        For lngRow = 1 To .Rows - 1
            For lngCol = 0 To .Cols - 1 Step 2
                If .TextMatrix(lngRow, lngCol) = IDKind֤��.GetCurCard.���� Then
                    txt֤��.Tag = .TextMatrix(lngRow, lngCol + 1)
                    txt֤��.Text = txt֤��.Tag
                    Exit For
                End If
            Next
        Next
    End With
End Sub


Private Function LoadCard(blnBoundCard As Boolean, Optional blnNotCardFee As Boolean = False) As Boolean
'����:ˢ������
'����:blnBoundCard-�󶨾��￨,��ģʽ��,������Ϣ������ʾ������¼����￨,����Ϊ���¿�ģʽ
'        blnNotCardFee-����ȡ����(ֻ���ڵ�󶨿����Ҳ���������Ϊ��ʱ,��Ϊ�ǰ󶨿�),����:38841
'����:True-δ����,���Ѻ͹Һŷ�һ����,false-�ѽ���,���Ѵ�Ϊ���۵�

    Dim blnInRange As Boolean
    Dim strCardNo As String
    '90875:���ϴ�,2016/11/8,ҽ�ƿ�֤������
    If IDKind.GetCurCard.�Ƿ�֤�� Then Exit Function
    
    mbln���� = False '�����:56599
    '115168:���ϴ���2017/12/13�����淢����ҽ�ƿ�����
    mCurSendCard = gCurSendCard
    If Not blnBoundCard Then
        Call ClearmobjfrmPatiInfoFace
    End If
    
    With mobjfrmPatiInfo
        .mbytFun = 1
        Set .mrs��ͥ��ַ = mrs��ͥ��ַ
        
        If blnBoundCard Then
            .mstrCard = ""
            Call CopyCboTofrmPatiInfo
            Call CopyInfoTofrmPatiInfo
        
            If .txt�����.Text = "" Then .txt�����.Text = zlGet�����
        Else
            '���¿�,��ˢ��ʱ�ͼ����￨�Ƿ��У��Ƿ��ڷ�Χ��
            blnInRange = True
            .mblnInRange = blnInRange
            .mstrCard = UCase(txtPatient.Text)
            .txt����.Text = .mstrCard
            
            mbln���� = bln����(.txt����.Text)
            
            If mbln���� = False And InStr(mstrPrivs, ";�󶨿���;") = 0 Then
                MsgBox "��û�а󶨿��ŵ�Ȩ�ޣ����ܰ󶨸ÿ���", vbInformation, gstrSysName
                Exit Function
            End If
            
            If Not gblnNewCardNoPop Then
                .txt�����.Text = zlGet�����
                txt�����.Text = .txt�����.Text
            End If
        End If
        If Not blnBoundCard And CreatePlugInOK(mlngModul) Then
            If Not zlReadPlugInPati(UCase(txtPatient.Text), mblnBrushPlugin) Then
                .txt����.Text = ""
                .txt����.Text = ""
                .txt��֤.Text = ""
                mblnAddCardItem = False
                Exit Function
            End If
        Else
            mblnBrushPlugin = False
        End If
        
        If blnBoundCard Or Not gblnNewCardNoPop Then
            '�����:53408
            Set mobjfrmPatiInfo.mrsPatiInfo = mrsInfo
            '�����:56599
            mobjfrmPatiInfo.mbln���� = mbln����
            .mlng�໤������ = mTy_Para.lngN������¼��໤��
            .mbln�໤��¼�� = mTy_Para.bln�໤��¼��
            If mrsInfo Is Nothing Then
                .mlng����ID = 0
            Else
                .mlng����ID = mrsInfo!����ID
            End If
            Call CloseIDCard '47007
            
            .ShowMe 1, Me
            
            Call NewCardObject '47007
            If .GetmblnCancel = True Then
                .txt����.Text = ""
                .txt����.Text = ""
                .txt��֤.Text = ""
                Call CopyCboTofrmPatiInfo
                Call CopyInfoTofrmPatiInfo
                Call NewCardObject
                Exit Function
            End If
            
            Set mrsInfo = Nothing
            Set mrsInfo = mobjfrmPatiInfo.mrsPatiInfo
            mstr����� = mobjfrmPatiInfo.txt�����
        Else
            '104238:���ϴ���2017/2/15����鿨���Ƿ����㷢����������
            If .txt����.Text <> "" And Len(.txt����.Text) <> gCurSendCard.lng���ų��� And Not gCurSendCard.bln�ϸ���� Then
                Select Case gCurSendCard.byt��������
                    Case 0
                        MsgBox "����Ŀ���С��" & gCurSendCard.str������ & "�趨�Ŀ��ų��ȣ����������룡", vbExclamation, gstrSysName
                        If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus
                            Exit Function
                    Case 2
                        If MsgBox("����Ŀ���С��" & gCurSendCard.str������ & "�趨�Ŀ��ų��ȣ��Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus
                            Exit Function
                        End If
                End Select
            End If
        End If
        '���˺�:27493 20100117:lnBoundCard = False
        If blnBoundCard Then
            If .mlng����ID <> 0 And gbln���ѽ����� Then
                strCardNo = .mlng����ID
                Call GetPatient(IDKind.GetCurCard, "-" & strCardNo, True)
                LoadCard = True
                cmdCard.Enabled = False
                Exit Function
            End If
            Call CopyInfoFromobjfrmPatiInfo
            blnInRange = IIf(blnNotCardFee, False, True)
            If .txt����.Text <> "" Then
                mbln���� = bln����(.txt����.Text)
            End If
            '31182
            If (mbytMode = 2 Or mbytMode = 0 And mbytInState = 0 And mstrNoIn <> "") And Not mrsInfo Is Nothing Then
                mblnAddCardItem = .txt����.Text <> "" And blnInRange And mbln����
            Else
                mblnAddCardItem = .txt����.Text <> "" And blnInRange And mbln����
            End If
            If .txt����.Text <> "" Then
                lblPrompt.Caption = gCurSendCard.str������ & ":" & .txt����.Text & "(" & IIf(mbln����, "����", "�󶨿�") & ")"
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
                If gbln���ѽ����� And Not gblnNewCardNoPop Then     '���������ɹ�,�󶨾��￨ģʽ�̶�������
                    Call GetPatient(IDKind.GetCurCard, txtPatient.Text, True)
                Else
                    mblnUnChange = True
                    Call CopyInfoFromobjfrmPatiInfo
                    mblnUnChange = False
                    If Me.ActiveControl Is txtPatient Then
                            If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(txtPatient.Text = "")
                            If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(txtPatient.Text = "")
                    End If
                    If gbln���ѽ����� Then
                        mblnAddCardItem = False
                    Else
                        mblnAddCardItem = mbln����
                    End If
                    lblPrompt.Caption = gCurSendCard.str������ & ":" & .mstrCard & "(" & IIf(mbln����, "����", "�󶨿�") & ")"
                    vsfPay.Height = 1755
                    lblPrompt.Top = vsfPay.Top + vsfPay.Height + 60
                End If
                Call ReLoadCardFee
                LoadCard = True
            Else '�ڵ�������ѡ����ȡ�����¿�
                cmdMore.Enabled = False
            End If
            cmdCard.Enabled = False
        End If
    End With
    Call AdjustInfoPosition
    If CheckIsPrice Or mRegistFeeMode = EM_RG_���� Then
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
    Dim i As Long, str���� As String
    
    With mobjfrmPatiInfo
    
        .cbo����.ListIndex = cbo.FindIndex(.cbo����, Nvl(mrsInfo!����), True)
        .cbo����.ListIndex = cbo.FindIndex(.cbo����, Nvl(mrsInfo!����), True)
        .cbo����.ListIndex = cbo.FindIndex(.cbo����, Nvl(mrsInfo!����״��), True)
        '76314,���ϴ���2014-08-06��������Ϣ��ȷ��ȡ
        .cboְҵ.ListIndex = cbo.FindIndex(.cboְҵ, Nvl(mrsInfo!ְҵ))
        .txt���֤��.Text = Nvl(mrsInfo!���֤��)
        .txt���֤��.Tag = .txt���֤��.Text
        .txt��λ����.Text = Nvl(mrsInfo!������λ)
        .txt����.Text = Trim(Nvl(mrsInfo!����))
        .txt����.Tag = .txt����.Text
        .txt��λ����.Tag = Nvl(mrsInfo!��ͬ��λID)
        .txt��λ�绰.Text = Nvl(mrsInfo!��λ�绰)
        .txt��λ�ʱ�.Text = Nvl(mrsInfo!��λ�ʱ�)
        .txt��ͥ�绰.Text = Nvl(mrsInfo!��ͥ�绰)
        .txt��ͥ�ʱ�.Text = Nvl(mrsInfo!��ͥ��ַ�ʱ�)
        .txt��ϵ�����֤.Text = Nvl(mrsInfo!��ϵ�����֤��)
        .txtBirthLocation.Text = Nvl(mrsInfo!�����ص�)
        .txtRegLocation.Text = Nvl(mrsInfo!���ڵ�ַ)
        '89242:���ϴ�,2015/12/7,��ȡ���˵�ַ��Ϣ
        Call zlReadAddrInfo(.padd���ڵ�ַ, Val(Nvl(mrsInfo!����ID)), 0, 4, Nvl(mrsInfo!���ڵ�ַ))
        .txt���ڵ�ַ�ʱ�.Text = Nvl(mrsInfo!���ڵ�ַ�ʱ�)
'        '73609:���ϴ���2014-8-1��������Ϣ����
'        .txtRegLocation.Tag = Nvl(mrsInfo!���ڵ�ַ�ʱ�)
        '�����:40005
        .txt��ϵ�˵绰.Text = Nvl(mrsInfo!��ϵ�˵绰)
        '84313,���ϴ�,2015/4/27,��ϵ�˹�ϵ�Լ�������ϵ
        .txt������ϵ.Text = ""
        .cbo��ϵ�˹�ϵ.ListIndex = cbo.FindIndex(.cbo��ϵ�˹�ϵ, Nvl(mrsInfo!��ϵ�˹�ϵ), True)
        If .cbo��ϵ�˹�ϵ.ListIndex <> 8 Then .txt������ϵ.Text = "": .txt������ϵ.Visible = False
        .txt��ϵ������.Text = Nvl(mrsInfo!��ϵ������)
        .txt�໤��.Text = Nvl(mrsInfo!�໤��)
        .Load�����������Ϣ (mrsInfo!����ID)
        .LoadCertificate (mrsInfo!����ID)
    End With
End Sub

Private Sub ShowPatiInfo()
    Dim i As Integer
    Dim strSimilar As String
    
    If txtPatient.Text = "" Then Exit Sub
    
    With mobjfrmPatiInfo
        .mbytFun = 0
        Set .mrs��ͥ��ַ = mrs��ͥ��ַ
        Call CopyCboTofrmPatiInfo
        Call CopyInfoTofrmPatiInfo
                
        If .txt�����.Text = "" Then .txt�����.Text = zlGet�����
'        .txt�����.Enabled = mrsInfo Is Nothing
                
        If mlngOutModeMC > 0 Then
            .txtPatiMCNO(0).Enabled = (mstrYBPati = "")
            .txtPatiMCNO(1).Enabled = .txtPatiMCNO(0).Enabled
        End If
    End With
    mobjfrmPatiInfo.mlng�໤������ = mTy_Para.lngN������¼��໤��
    mobjfrmPatiInfo.mbln�໤��¼�� = mTy_Para.bln�໤��¼��
    mobjfrmPatiInfo.mstrPrivs = mstrPrivs
    mobjfrmPatiInfo.mlngModul = mlngModul
    Call CloseIDCard
    mobjfrmPatiInfo.ShowMe 1, Me
    Call NewCardObject
    If mobjfrmPatiInfo.GetmblnCancel = False Then
        '�����ˢ���½����˵���,����mobjfrmPatiInfo���ȷ��ʱ���ɲ�����Ϣ֮ǰ����
        If Trim(mobjfrmPatiInfo.txt���֤��.Text) <> "" And cmdMore.Tag = "" And mobjfrmPatiInfo.cmdOK.Caption Like "����*" And mobjfrmPatiInfo.txt���֤��.Tag <> Trim(mobjfrmPatiInfo.txt���֤��.Text) Then
            '������Ʋ�����Ϣ(����֮ǰ���,����������ظ���Ϣ������)
            With mobjfrmPatiInfo
                strSimilar = SimilarIDs(.txt���֤��.Text)
            End With
            cmdMore.Tag = "�Ѽ��"      '��txtPatient_change�����
            
            If strSimilar <> "" Then
                i = UBound(Split(strSimilar, "|")) + 1
                strSimilar = Replace(strSimilar, "|", vbCrLf)
                If i > 20 Then strSimilar = Mid(strSimilar, 1, 200) & "..."
                
                If MsgBox("�����еĲ�����Ϣ�з��� " & i & " ����Ϣ���ƵĲ���(���֤����ͬ): " & vbCrLf & vbCrLf & _
                    strSimilar & vbCrLf & vbCrLf & "�Ǽ�Ϊ�²�����ѡ��[��],��ȡ���еĲ�����Ϣ��ѡ��[��]��", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
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
    
    '74430,Ƚ����,2014-7-8,�ҺŽ�����ʾ������Ƭ�ĸ�������
    If picPatiPicBack.Visible Then Call ShowPatiPic
    
    If cbo���㷽ʽ.Enabled And cbo���㷽ʽ.Visible Then
        cbo���㷽ʽ.SetFocus
    ElseIf chk������.Enabled And chk������.Visible Then
        chk������.SetFocus
    Else
        cmdOK.SetFocus
    End If
End Sub

Private Sub cmdCard_Click()
    Dim blnBound As Boolean
    
    If LoadCard(True, blnBound) Then
        Call ShowRegistFromInput    '�����Ȱ󶨿��ŷ��غ��ٴν����������
         '�����:56039,56355
        If Val(zlDatabase.GetPara("�Һŷ�Ʊ��ӡ��ʽ", glngSys, mlngModul)) <> 0 Then
           Call ReInitPatiInvoice
        End If
        
        If mobjfrmPatiInfo.txt����.Text <> "" Then
            mblnNewCard = True
            Call SetOneCardBalance
        Else
            SetCboDefault cbo���㷽ʽ
        End If
    End If
    If cbo���㷽ʽ.Enabled And cbo���㷽ʽ.Visible Then
        cbo���㷽ʽ.SetFocus
    ElseIf chk������.Enabled And chk������.Visible Then
        chk������.SetFocus
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
    If frmPatiFind.mlng����ID <> 0 Then
        Me.Refresh
        txtPatient.Text = "-" & frmPatiFind.mlng����ID
        Call txtPatient_Validate(False)
    Else
        If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    End If
End Sub

Private Sub dtpAppointmentDate_Change()
    txtSN.Text = ""
    Call ShowPlans
    dtpAppointmentDate.Tag = Format(dtpAppointmentDate.Value, "yyyy-mm-dd HH:MM:SS")
    If txt�ű�.Text <> "" Then
        If zlCheck��Լ���޺���(Val(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("��¼ID")))) = False Then
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
    Dim lng�ű� As Long
    '�����:57491
    Call picInfoFrame_Resize
    If Not mblnFirst Then Exit Sub
    
    mblnFirst = False
    
    If mblnUnload Then mblnUnload = False: Unload Me: Exit Sub
    
    Call zlȨ�޿���
    
    'ҽ��վ�Һ�ʱ�����ֻ��һ���ţ����Զ�����
    With vsfPlan
        If .Rows = 2 Then
            lng�ű� = GetCol("�ű�")
            If .TextMatrix(1, lng�ű�) <> "" And txt�ű�.Visible And txt�ű�.Enabled Then
                txt�ű�.SetFocus
                txt�ű�.Text = .TextMatrix(.Row, lng�ű�)
            End If
        End If
    End With
    If mbytInState = 0 And mbytMode = 0 Then
        txtPatient_Change
    End If
    '˫����ʾ��������ڵ�ǰ������ʾ֮�������ʾ�����ƶ�����
    If mbytMode = 0 And mbytInState = 0 And gblnLED And Trim(txtPatient.Text) = "" Then
        zl9LedVoice.DisplayPatient ""
    End If
    
    If mbytMode = 2 And cmdOK.Visible And cmdOK.Enabled Then
        cmdOK.SetFocus
    End If
    If gCurSendCard.str������ <> "" Then
        cmdCard.ToolTipText = "��" & gCurSendCard.str������ & ": F10"
        If mblnSendCard Then cmdCard.ToolTipText = "��" & gCurSendCard.str������ & ": F10"
    End If
    Call picPlan_Resize
    mblnActivate = True
    If mbytMode = 2 And mbytInState = 0 Then
        '102230,������Ҳ����ӿ�
        If Not mrsInfo Is Nothing Then
            If PatiValiedCheckByPlugIn(mlngModul, Val(Nvl(mrsInfo!����ID)), _
                "<YSXM>" & NeedName(cboҽ��.Text) & "</YSXM>") = False Then Unload Me: Exit Sub
        End If
    Else
        Call vsfPlan_EnterCell: If txt�ű�.Visible And txt�ű�.Enabled Then txt�ű�.SetFocus
    End If
    mblnActivate = False
    Call picInfo_Resize
End Sub
Private Sub zlȨ�޿���()
      '���˺� ����:27438 ����:2010-01-13 17:42:32
    If mbytInState <> 0 Then Exit Sub
    If mbytMode = 0 Then
        cmdCard.Visible = InStr(1, mstrPrivs, ";�󶨿���;") > 0
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
            '����ctrl+M
            If Shift <> vbCtrlMask Then Exit Sub
            If ActiveControl Is txtPatient Then Call txtPatient_Validate(False)
            If Shift = vbCtrlMask And cmdMore.Enabled And cmdMore.Visible Then cmdMore_Click
        Case vbKeyF2
            If ActiveControl Is txtPatient Then
                Call txtPatient_Validate(False)
            End If
            If Not blnCancel And cmdOK.Enabled And cmdOK.Visible Then Call cmdOK_Click  '�������ý���,��Ϊ�����¼����Դ��ж��Ƿ������������
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
            If txt�ű�.Enabled And txt�ű�.Visible Then
                mblnLEDKey = True
                If Not Me.ActiveControl Is txt�ű� Then
                    txt�ű�.SetFocus
                Else
                    Call txt�ű�_GotFocus 'LED��������
                End If
            End If
        Case vbKeyF10
            mbln���� = False '�����:56599
            If ActiveControl Is txtPatient Then Call txtPatient_Validate(False)
            If cmdCard.Visible And cmdCard.Enabled Then Call cmdCard_Click
        Case vbKeyF11
            If txtPatient.Enabled And txtPatient.Visible And Not txtPatient.Locked Then
                If Me.ActiveControl Is txtPatient Then
                    IDKind.IDKind = IIf(IDKind.IDKind = UBound(Split(IDKind.IDkindStr, ";")), IDKind.GetKindIndex("����"), IDKind.IDKind + 1)
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
            If mbytInState = 0 And Not mbln������ Then Exit Sub
            If (mbytMode = 2 Or mbytMode = 0 And mbytInState = 0 And mstrNoIn <> "") Or chkCancel.Value = 1 Or chkPrint.Value = 1 Or txt�ű�.Text = "+" Then Exit Sub
            If ActiveControl.Name <> txt�ű�.Name Then
                chk������.Value = IIf(chk������.Value = 0, 1, 0)
            End If
        Case 192, 229  '����:28604:��
             If Shift <> vbCtrlMask Then
                Exit Sub
             End If
             Call SelectHistoryRegist
    End Select
    
    '74430,Ƚ����,2014-7-8,�ҺŽ�����ʾ������Ƭ�ĸ�������
    If Shift = 2 And KeyCode = vbKeyW Then
         Call ShowPatiPic
    End If
    If Shift = 2 And KeyCode = vbKeyE Then
        Call imgColPlan_Click
    End If
End Sub

Private Sub SelectHistoryRegist()
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ�ѡ�����ιҺźű�
    '���ƣ����˺�
    '���ڣ�2010-08-18 16:14:58
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim lng����ID As Long, lngPre����ID As Long, str�ű� As String
    Dim blnFind As Boolean, i As Long
    If mbytMode = 2 Then Exit Sub 'ԤԼ���ղ�����
    If mbytInState >= 1 Then Exit Sub  '���Ĳ�����
    If mrsInfo Is Nothing Then
        lng����ID = 0
    ElseIf mrsInfo.State <> 1 Then
        lng����ID = 0
    Else
       lng����ID = Val(Nvl(mrsInfo!����ID))
    End If
    lngPre����ID = lng����ID
    str�ű� = ""
    CloseIDCard
    If frmRegistHistory.ShowRegist(Me, mstrPrivs, mTy_Para.bln����סԺ���˹Һ�, mblnOlnyBJYB, lng����ID, str�ű�) = False Then NewCardObject: Exit Sub
    Call CreateMobjIDCard
    If lng����ID <> lngPre����ID Then
       '���˲���ʱ,ֱ�Ӷ�ȡ����
       Call GetPatient(IDKind.GetCurCard, "-" & lng����ID, False)
    End If
    
    '�����д˺ű�û��
    With vsfPlan
       blnFind = False
       For i = 1 To .Rows - 1
           If .TextMatrix(i, .ColIndex("�ű�")) = str�ű� Then
                   .Row = i: .Col = .ColIndex("�ű�")
                   Call .ShowCell(.Row, .Col)
                   Call vsfPlan_KeyDown(13, 0)
                   blnFind = True: Exit For
           End If
       Next
    End With
    If blnFind = False Then
       Call MsgBox("ע��:" & vbCrLf & "    ����Ϊ��" & str�ű� & "���ĺ����ڵ�ǰδ���йҺŰ���,�޷���λ!", vbInformation + vbOKOnly, gstrSysName)
       Exit Sub
    End If
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    
    If InStr("':��;��?��", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0: Exit Sub
    ElseIf KeyAscii = Asc("+") Then
        If (mbytMode = 2 Or mbytMode = 0 And mbytInState = 0 And mstrNoIn <> "") Or chkCancel.Value = 1 Or chkPrint.Value = 1 Then KeyAscii = 0
    End If
    If mbytInState = 1 Then Exit Sub
    If InStr("`��", Chr(KeyAscii)) > 0 Then
        '�����ʾ���￨
         KeyAscii = 0
        If gblnLED Then zl9LedVoice.Speak "#30"  '`Ϊ��������:�е����:����Ӧ����192,����֪��ô���229:32663
    End If
    
End Sub

Private Sub Form_Load()
    Dim lng������ID As Long
    Dim strSQL As String, rsTmp As ADODB.Recordset
'    Call InitTimeSect
    '��ʼ�� ������õ� ��ʽ
    InitActionType
    Call zlInitParaSet  '��ʼ�����ز���
    '����ߴ�����
    '�����彨
    Call InitCardSquareData
    Call DefMainCommandBars
    Call InitRegist
    Call InitPanel
   ' Call zlInitParaSet  '��ʼ�����ز���
    mblnStartFactUseType = False
    If gblnSharedInvoice Then
        '�Һ�������Ʊ��:42703
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
    mstr�����ʻ� = ""
    mlng����ID = 0
    mintInsure = 0
    mstrYBPati = ""
    mlng�ſ�����ID = 0
    
    cmdComminuty.Visible = False
    If (mbytMode = 0 Or mbytMode = 1) And mbytInState = 0 Then
        Set mobjIDCard = New clsIDCard
        Set mobjICCard = New clsICCard
        Call mobjIDCard.SetParent(Me.Hwnd)
        Call mobjICCard.SetParent(Me.Hwnd)
        Set mobjICCard.gcnOracle = gcnOracle

        '�����ӿڳ�ʼ��
        Call CreateCommunity
        
    End If
    
    If mintCancel = 1 Then
        lng������ID = 0
        strSQL = "Select �շ�ϸĿID From �շ��ض���Ŀ Where �ض���Ŀ='������'"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        If Not rsTmp.EOF Then
            lng������ID = Val(Nvl(rsTmp!�շ�ϸĿID))
        End If
        
        If lng������ID = 0 Then
            MsgBox "û�з��ֲ����ѵ��շ��ض���Ŀ�����飡", vbExclamation, gstrSysName
            mblnUnload = True
        Else
            mstr�˷���ĿIDs = lng������ID
        End If
    End If
    
    mstr���ӷ� = ""
    mstr������ĿID = ""
    strSQL = "Select zl_Fun_RegCustomName As ���ӷ� From Dual"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If Not rsTmp.EOF Then
        mstr���ӷ� = Split(Nvl(rsTmp!���ӷ�) & "|", "|")(0)
        mstr������ĿID = Split(Nvl(rsTmp!���ӷ�) & "|", "|")(1)
    End If
    
    If mstr���ӷ� <> "" Then
        chkExtra.Caption = "��" & mstr���ӷ�
    End If

    '��ʼ������
    If mbytInState = 0 Then
        mobjfrmPatiInfo.mstrPriceGrade = gstrPriceGrade
    End If
    Call Load֧����ʽ
    Call InitFace
    If mbytInState <> 1 And mbytInState <> 2 And mbytInState <> 3 Then
        Call RestoreWinState(Me, App.ProductName, mbytMode & mbytInState)
        stbThis.Visible = True
    End If
    
    Call InitData
    '�����:57491
    If mblnUnload Then
        Exit Sub
    End If
    
    Call SetDelBillCtlEnabled
    
    
    If mblnStation And mbytMode = 0 And mTy_Para.bln�Һű���ˢ�� Then LoadIdKindStr  '�����ҽ������վ�ҺŲ��ҹҺű���ˢ��ʱ��Ҫ ���¼��� IDKind����Ӧ��Ϣ
    If mblnUnload Then Exit Sub
    
    If mbytMode = 1 Then
        'ԤԼ ��Ҫ��ʼ��������λ�Һ�
        Call InitUnitRegData
    End If
    
    If Me.Height < glngMinH Then Me.Height = glngMinH
    If Me.Width <= glngMinW Then Me.Width = glngMinW
    
    If mbytInState = 1 Or (mbytInState = 0 And mbytMode = 2) Then '����ʱ,���ܸ��Ĵ����С:25623
        Call zlSetWindowsBroldStyle(Me)
        Call Form_Resize
    End If
    zlControl.PicShowFlat picInfoFrame, -1, , taCenterAlign
    zlControl.PicShowFlat picPlan, -1, , taCenterAlign
    zlControl.PicShowFlat picTotal, -1, , taCenterAlign

    'LED��ʼ��
    If mbytMode <> 1 And mbytInState = 0 And gblnLED Then
        zl9LedVoice.Reset com
        zl9LedVoice.Init UserInfo.��� & " �Һ�ԱΪ������", mlngModul, gcnOracle
    End If
End Sub

Private Sub InitUnitRegData()
    Dim strSQL As String
    Dim rsTmp   As ADODB.Recordset
    
    strSQL = " Select 1 as ����  From �ٴ�����Һſ��Ƽ�¼ Where ����=1 And Rownum < 2 "
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
        If MsgBox("���Ҫ�˳���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
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
            lbl�ϼ�.Left = picTotal.Width - lbl�ϼ�.Width - 150
        End If
    ElseIf mbytMode = 3 Then
        picTotal.Width = picBal.Left - picTotal.Left
        lbl�ϼ�.Left = picTotal.Width - lbl�ϼ�.Width - 150
    ElseIf mbytMode = 2 Then
        If mbytInState = 1 Then
            picTotal.Width = picBal.Left - picTotal.Left
            lbl�ϼ�.Left = picTotal.Width - lbl�ϼ�.Width - 150
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
    lbl��.Left = lblFree.Left - lbl��.Width - 150
    lbl����.Left = lbl��.Left - lbl����.Width - 180
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call YBIdentifyCancel 'ȡ��ҽ�����������֤
    
    Call SaveWinState(Me, App.ProductName, mbytMode & mbytInState)
    
    mblnRegReceiveByNo = False '�����:57423
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
    Set mrsALLʱ��� = Nothing
    Set mrsʱ��� = Nothing
    Set mrsItems = Nothing
    Set mrsInComes = Nothing
    Set mrsPlan = Nothing
    Set mrsInfo = Nothing
    Set mrs�ѱ� = Nothing
    Set mrsDoctor = Nothing
    Set mrsSNState = Nothing
    Set mrsBillAdvance = Nothing
    Set mobjDelCards = Nothing
    Set mobjPayCard = Nothing
    If Not mrs��ͥ��ַ Is Nothing Then
        If mrs��ͥ��ַ.State = 1 Then
            On Error Resume Next
            Kill App.Path & "\ZLAddressForRegEvent.Adtg"
            Err.Clear
            mrs��ͥ��ַ.Filter = ""
            mrs��ͥ��ַ.Save App.Path & "\ZLAddressForRegEvent.Adtg"
        End If
    End If
    Set mrs��ͥ��ַ = Nothing
    
    mbln������ = False
    mbln���������� = False
    mlng����ID = 0
    
    mstrPrePati = ""
    mcur�ϼ� = 0: mint�Һ��� = 0
    mcurӦ�� = 0
    
    If Not mobjfrmPatiInfo Is Nothing Then Unload mobjfrmPatiInfo
    Set mobjfrmPatiInfo = Nothing
    
    If Not OS.IsDesinMode And glngOld > 0 Then
        Call SetWindowLong(Me.Hwnd, GWL_WNDPROC, glngOld)
    End If
    If Not mobjRegist Is Nothing Then Set mobjRegist = Nothing
    
    'LED��ʼ��
    If mbytMode <> 1 And mbytInState = 0 And gblnLED Then
        zl9LedVoice.DisplayPatient ""
        zl9LedVoice.Reset com
    End If
    mintIDKind = IDKind.IDKind
    If mbytInState = 0 Then
        Call SaveRegInFor(g˽��ģ��, Me.Name, "idkind", mintIDKind)
    End If
    If mbytMode = 1 And mbytInState = 0 Then
        Call zlDatabase.SetPara("ԤԼ��ʾ���кű�", IIf(chkShowAll.Value = 1, 1, 0), glngSys, mlngModul, InStr(mstrPrivs, ";��������;") > 0)
    End If
    
    Call CloseIDCard
    mbytMode = 0
    mbytInState = 0
    mbln���ӷ� = False
    mstrPrivs = ""
    '�����:53408
    mstr����� = ""
    '�����:56599
    mbln���� = False
    Set mobjHealthCard = Nothing
    mblnNotEMPIQuery = False
    '127839�����ϴ�,2018/6/27����ձ���
    mcustomTime = t_��ͨ
    mViewMode = V_��ͨ��
    mblnUnload = False
    mbln�����Һ� = False
    mblnGetBirth = False
End Sub

Private Sub lbl�ϼ�_Change()
    Call txt�ɿ�_Change
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
    sc����.Width = picPlan.ScaleWidth
    vsfPlan.Width = picPlan.ScaleWidth
    vsfList.Width = picPlan.ScaleWidth
    picBookingDate.Width = picPlan.ScaleWidth
    picSplit.Width = picPlan.ScaleWidth
    picTime.Width = vsfPlan.Width
    
    sc����.Top = IIf(picBookingDate.Visible, picBookingDate.Height + 30, 0)
    chkShowAll.Top = sc����.Top + 45
    chkShowAll.Left = sc����.Width - chkShowAll.Width - 300
    vsfPlan.Top = sc����.Top + sc����.Height + 45
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
    lblԤԼʱ��.Left = dtpAppointmentTime.Left - lblԤԼʱ��.Width - 20
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
    If vsfPlan.TextMatrix(vsfPlan.Row, GetCol("�ű�")) <> txt�ű�.Text And txt�ű�.Text <> "" Then
        If mlngPreRow <> vsfPlan.Row And mlngPreRow < vsfPlan.Rows And mlngPreRow <> 0 Then
            vsfPlan.Row = mlngPreRow
            lblSN.Tag = "��Ų���ȫѡ" '����ֵʱ����ĸı䲻����ȫѡ
            Call zlControl.ControlSetFocus(txtSN)
        End If
    End If
End Sub

Private Sub txt����ʱ��_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub vsfPay_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim strӦ�� As String
    If Col = 1 Then
        strӦ�� = txt����Ӧ��.Text
        txt����Ӧ��.Text = Format(Val(txt����Ӧ��.Text) + mdblԭ��� - Val(vsfPay.TextMatrix(Row, 1)), "0.00")
        If txt����Ӧ��.Text < mcurӦ�� Or Val(vsfPay.TextMatrix(Row, 1)) > Val(vsfPay.TextMatrix(Row, 6)) Then
            txt����Ӧ��.Text = strӦ��
            vsfPay.TextMatrix(Row, 1) = Format(mdblԭ���, "0.00")
        Else
            vsfPay.TextMatrix(Row, 1) = Format(Val(vsfPay.TextMatrix(Row, 1)), "0.00")
        End If
    End If
    If Col = 2 Then
        If zlCommFun.ActualLen(vsfPay.TextMatrix(Row, 2)) > 30 Then
            MsgBox "����������,��������д!", vbInformation, gstrSysName
            vsfPay.TextMatrix(Row, 2) = ""
        End If
    End If
End Sub

Private Sub vsfPay_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If vsfPay.TextMatrix(Row, 0) = "" Then Cancel = True
    If mblnCenter Then Cancel = True: Exit Sub
    If mbytMode = 4 Or chkCancel.Value = 1 Then
        '�˺�
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
        If Val(vsfPay.TextMatrix(Row, vsfPay.ColIndex("����޸�"))) <> 1 Then
            Cancel = True
        Else
            vsfPay.ComboList = ""
            mdblԭ��� = Val(vsfPay.TextMatrix(Row, 1))
        End If
    End If
    
    If Col = 0 Then
        If mbln�����Һ� Then
            Cancel = True
        Else
            If Val(vsfPay.TextMatrix(Row, vsfPay.ColIndex("�޸�"))) = 0 Then
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
            If Val(vsfPay.RowData(Row)) <> 8 And Val(vsfPay.RowData(Row)) <> 7 And vsfPay.TextMatrix(Row, 0) Like "*��*" Then
                vsfPay.ComboList = ""
            Else
                Cancel = True
            End If
        End If
    End If
End Sub

Private Sub vsfPay_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim dblMoney As Double, i As Integer, blnFind As Boolean, rsTemp As ADODB.Recordset
    Dim strSQL As String, str�˿����Ա As String
    On Error GoTo errH
    If vsfPay.TextMatrix(Row, Col) = "" Then Exit Sub
    If mbytMode = 4 Or chkCancel.Value = 1 Then
    
        If (Val(vsfPay.RowData(Row)) = 7 Or Val(vsfPay.RowData(Row)) = 8) And Val(vsfPay.TextMatrix(Row, vsfPay.ColIndex("�޸�"))) = 1 Then
            If InStr(mstrCardPrivs, ";�����˿�ǿ������;") = 0 Then
                str�˿����Ա = zlDatabase.UserIdentifyByUser(Me, "ǿ��������֤", glngSys, 1151, "�����˿�ǿ������")
                If str�˿����Ա = "" Then
                    MsgBox "¼��Ĳ���Ա��֤ʧ�ܻ���¼��Ĳ���Ա���߱�ǿ������Ȩ�ޣ�����ǿ�����֣�", vbInformation, gstrSysName
                    Exit Sub
                End If
                mstrForceNote = str�˿����Ա & "ǿ������:" & vsfPay.TextMatrix(Row, vsfPay.ColIndex("���㷽ʽ")) & "," & vsfPay.TextMatrix(Row, 1) & "Ԫ"
            Else
                If MsgBox(vsfPay.TextMatrix(Row, 0) & "��֧�����֣��Ƿ�ǿ�����֣�", vbQuestion + vbYesNo, gstrSysName) <> vbYes Then Exit Sub
                mstrForceNote = UserInfo.���� & "ǿ������:" & vsfPay.TextMatrix(Row, vsfPay.ColIndex("���㷽ʽ")) & "," & vsfPay.TextMatrix(Row, 1) & "Ԫ"
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
            strSQL = "Select ���� From ���㷽ʽ Where ����=1 Order By ȱʡ��־ Desc"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
            For i = 1 To vsfPay.Rows - 1
                If vsfPay.TextMatrix(i, 0) = "" Then
                    blnFind = True
                    vsfPay.TextMatrix(i, 0) = Nvl(rsTemp!����)
                    vsfPay.TextMatrix(i, 1) = Format(dblMoney, "0.00")
                    vsfPay.TextMatrix(i, vsfPay.ColIndex("�޸�")) = "1"
                    vsfPay.RowData(i) = 1
                End If
                If blnFind Then Exit For
            Next i
            
            If blnFind = False Then
                vsfPay.Rows = vsfPay.Rows + 1
                vsfPay.TextMatrix(vsfPay.Rows - 1, 0) = Nvl(rsTemp!����)
                vsfPay.TextMatrix(vsfPay.Rows - 1, 1) = Format(dblMoney, "0.00")
                vsfPay.TextMatrix(vsfPay.Rows - 1, vsfPay.ColIndex("�޸�")) = "1"
                vsfPay.RowData(vsfPay.Rows - 1) = 1
            End If
        End If
    Else
        dblMoney = Val(vsfPay.TextMatrix(Row, 1))
        txt����Ӧ��.Text = Format(Val(txt����Ӧ��.Text) + dblMoney, "0.00")
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
    '����:���ùҺźű���ɫ
    '����:���˺�
    '����:2010-02-04 14:13:50
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
    '����:������ع̶��еı���ɫ
    '����:blnCurDate-�Ƿ�ǰ������,�������ԤԼ������
    '����:���˺�
    '����:2010-02-04 14:39:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim PreRedaw As RedrawSettings, i As Long, strSQL As String, strNow As String
    Dim strKey As String, rsTmp As ADODB.Recordset, strColor As String
    Dim j As Long, strPrevKey As String
    Dim DatCur As Date
    With vsfPlan
         .Redraw = flexRDNone
         If blnCurDate Then
             strKey = zlGet��ǰ���ڼ�
             DatCur = zlDatabase.Currentdate
             strPrevKey = zlGet��ǰ���ڼ�(Format(DatCur - 1, "yyyy-mm-dd"))
             For i = 1 To .Rows - 1
                If Format(DatCur, "yyyy-mm-dd") = Format(.TextMatrix(i, .ColIndex("��������")), "yyyy-mm-dd") Then
                    .Cell(flexcpData, 0, .ColIndex(strKey)) = 1 '��ǰ����
                    .Cell(flexcpBackColor, i, .ColIndex(strKey), i, .ColIndex(strKey)) = &HFF8080
                    .Cell(flexcpFontBold, i, .ColIndex(strKey), i, .ColIndex(strKey)) = True
                Else
                    .Cell(flexcpData, 0, .ColIndex(strPrevKey)) = 1 '��������
                    .Cell(flexcpBackColor, i, .ColIndex(strPrevKey), i, .ColIndex(strPrevKey)) = &HFF8080
                    .Cell(flexcpFontBold, i, .ColIndex(strPrevKey), i, .ColIndex(strPrevKey)) = True
                End If
             Next i
             
            strColor = zlDatabase.GetPara("��ǰ�Һ���ɫ", glngSys, mlngModul, "0")
            strNow = Format(DatCur, "YYYY-MM-DD HH:MM:SS")
            For i = 1 To .Rows - 1
                If .Cell(flexcpData, i, .ColIndex("�ű�")) = "1" Then
                    For j = 1 To .Cols - 1
                        If .Cell(flexcpData, 0, j) <> 1 Then
                            .Cell(flexcpForeColor, i, j, i, j) = &H8000000C
                        End If
                    Next j
                Else
                    If .TextMatrix(i, .ColIndex("��ǰʱ��")) <> "" Then
                        If strNow < Format(.TextMatrix(i, .ColIndex("�Һ�ʱ��")), "YYYY-MM-DD HH:MM:SS") Then
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
            strKey = zlGet��ǰ���ڼ�(Format(dtpAppointmentDate.Value, "yyyy-mm-dd"))
            strPrevKey = zlGet��ǰ���ڼ�(Format(dtpAppointmentDate.Value - 1, "yyyy-mm-dd"))
            If .ColIndex(strKey) < 0 Then Exit Sub
            For i = 1 To .Cols - 1
                If i <> .ColIndex(mstr��ǰ����) Then  '��ǰԤԼ������
                    For j = 1 To .Rows - 1
                        If j Mod 2 = 1 Then
                            .Cell(flexcpBackColor, j, i, j, i) = &H80000005
                        Else
                            .Cell(flexcpBackColor, j, i, j, i) = &HF6F6F6
                        End If
                    Next j
                     .Cell(flexcpFontBold, 1, i, .Rows - 1, i) = False
                ElseIf Val(.ColData(.ColIndex(strKey))) = 1 Then    '��ǰ���ڵ����ڼ���
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
                If Format(DatCur, "yyyy-mm-dd") = Format(.TextMatrix(i, .ColIndex("��������")), "yyyy-mm-dd") Then
                    .Cell(flexcpData, 0, .ColIndex(strKey)) = 1 '��ǰ����
                    .Cell(flexcpBackColor, i, .ColIndex(strKey), i, .ColIndex(strKey)) = &HFF8080
                    .Cell(flexcpFontBold, i, .ColIndex(strKey), i, .ColIndex(strKey)) = True
                Else
                    .Cell(flexcpData, 0, .ColIndex(strPrevKey)) = 1 '��������
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

Private Sub SetSnStyle(Optional ByVal bln��ʱ�� As Boolean = False)
'****************************************
'�Ա����ʽ��������
'****************************************
    Dim i           As Long
    Dim lngWidth    As Long
    Dim X           As Long
    Dim Y           As Long
    Dim j           As Long
    Dim lngHeight   As Long
    Select Case bln��ʱ��
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
    '����ʱ���
    '***************************************
    Dim i               As Integer
    Dim j               As Integer
    Dim blnPre          As Boolean
    Dim lngThis         As Long
    Dim lngMax          As Long
    Dim datThis         As Date
    Dim lngCurrSn       As Long
    Dim lngMaxSn        As Long 'ԤԼ�����ʹ�ú�
    Dim strSQL          As String
    Dim rsʱ��ͳ��      As ADODB.Recordset
    Dim strʱ���       As String
    Dim lngԤԼ����     As Long
    Dim lngTatol        As Long '���ڷ�ʱ�� ������¼�������
    Dim strMaxDate      As String  '���ڷ�ʱ�α����ԤԼʱ��
    Dim lngCols         As Long
    Dim lngRows         As Long
    Dim strData         As String
    Dim strDate         As String
    Dim lng��¼ID       As Long
    Dim blnHave         As Boolean
    Dim datMax          As Date
    Dim Datsys          As Date
    Dim blnʧԼ���ڹҺ� As Boolean
    Dim blnInserted     As Boolean
    Dim lng������λ���� As Long
    Dim blnFindSN      As Boolean '�Ƿ���Ҫ���¶�λ���ϴκű�����,����ˢ���б�ʱ,���ݱ���
    Dim lngFindSN      As Long '��Ҫ���ҵ����
    Dim str����         As String
    str���� = "(��)"
    vsfList.Visible = True
    picSplit.Visible = True
    vsfList.Redraw = False
    mblnStateChange = True
    vsfList.Clear
    '***************************************
    '�����Ϣ����
    '***************************************
    If dkpMain.Panes(2).Hidden Then
        dkpMain.Panes(2).Hidden = False
        mcbrToolBar.Controls.Find(xtpControlButton, 2605).Visible = InStr(1, mstrPrivs, ";Ԥ������;") > 0 '36294
        mcbrToolBar.Controls.Find(xtpControlButton, 2604).Visible = InStr(1, mstrPrivs, ";Ԥ������;") > 0 '36294
        Call Form_Resize
    Else
        mcbrToolBar.Controls.Find(xtpControlButton, 2605).Visible = InStr(1, mstrPrivs, ";Ԥ������;") > 0 '36294
        mcbrToolBar.Controls.Find(xtpControlButton, 2604).Visible = InStr(1, mstrPrivs, ";Ԥ������;") > 0 '36294
        Call Form_Resize
    End If
    If mbytMode = 1 Then
        lngMax = Val(vsfPlan.TextMatrix(vsfPlan.Row, GetCol("��Լ")))
    Else
        lngMax = Val(vsfPlan.TextMatrix(vsfPlan.Row, GetCol("�޺�"))) '�ҽ����ĺŲ�����ԤԼ,��Ϊ�ѽ���,Ӧ���ɹҺ�
    End If
    If mbytMode = 1 Then
        lngMax = Val(vsfPlan.TextMatrix(vsfPlan.Row, GetCol("�޺�")))
    End If
    
    '1.����λ��
    If lngMax > 1000 Then
        vsfList.FontWidth = 4
    Else
        vsfList.FontWidth = 0 '�ָ�ȱʡ����
    End If
    '***************************************
    '��ʼ��ʱ���
    '***************************************
     If InitTimePlan() = False Then vsfList.Redraw = True: Exit Sub
     Datsys = zlDatabase.Currentdate
    '***************************************
    '��ʼ�����
    '***************************************
     
     If mrsʱ��� Is Nothing Then vsfList.Redraw = True: Exit Sub
     'If mrsʱ���.RecordCount = 0 Then Exit Sub
 
    '***************************************
    '������
    '***************************************
     With vsfList
        .Rows = 1
        .Cols = 1
        .Clear
     End With
     lngCurrSn = -1
     If mstrPre�ű� <> "" Then
        blnFindSN = mstrPre�ű� = mtyRegPlanState.str�ű�
        blnFindSN = blnFindSN And mViewMode = v_ר�Һŷ�ʱ�� And txtSN.Text <> ""
        If blnFindSN Then lngFindSN = Val(txtSN.Text)
     End If
    Select Case mViewMode
    Case V_��ͨ�ŷ�ʱ��:
       
        strSQL = "Select Count(1) As ԤԼ����, To_Char(��ʼʱ��, 'HH24:MI') As ����" & vbNewLine & _
                "From �ٴ�������ſ���" & vbNewLine & _
                "Where Nvl(�Һ�״̬,0) <> 0 And Nvl(�Ƿ�ԤԼ,0) = 1 And ԤԼ˳��� Is Not Null And ��¼id = [1]" & vbNewLine & _
                "Group By ��ʼʱ��"
        
        lng��¼ID = vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("��¼ID"))
        On Error GoTo Hd
        Set rsʱ��ͳ�� = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng��¼ID)
        blnHave = False
        
        strʱ��� = ""
        With mrsʱ���
          datMax = CDate("00:00:00")
          mdatLast = CDate("00:00:00")
          lngRows = -1: lngCols = 0
           Do While Not .EOF
                If IsNull(!ԤԼ˳���) Then
                    If datMax < CDate(Nvl(!��ʼʱ��, "00:00:00")) Then datMax = CDate(!��ʼʱ��)
                    If mdatLast < CDate(Nvl(!����ʱ��, "00:00:00")) Then mdatLast = CDate(!����ʱ��)
                    'ԤԼ״̬ ֻ�������ԤԼ��ʱ���
                    '�Һ�ʱ�����ֶ����
                     rsʱ��ͳ��.Filter = " ����='" & Nvl(!��ʼʱ��, "_") & "'"
                     If rsʱ��ͳ��.RecordCount = 0 Then
                        lngԤԼ���� = 0
                     Else
                        lngԤԼ���� = rsʱ��ͳ��!ԤԼ����
                     End If
                     
                     lng������λ���� = 0
                     If mbytMode = 1 And mblnUnitReg And Not mrsUnitReg Is Nothing Then
                         mrsUnitReg.Filter = "���=" & Val(Nvl(!���))
                         lng������λ���� = 0
                         If mrsUnitReg.RecordCount > 0 Then
                            lng������λ���� = Val(Nvl(mrsUnitReg!����))
                         End If
                     End If
                      
                     If Nvl(!��������, 0) <> 0 Then
                        If strʱ��� <> Nvl(!ʱ���) Then
                            lngRows = lngRows + 1
                            strʱ��� = Nvl(!ʱ���)
                            If lngRows > vsfList.Rows - 1 Then vsfList.Rows = vsfList.Rows + 1: lngCols = 0
                            If lngCols > vsfList.Cols - 1 Then vsfList.Cols = vsfList.Cols + 1
                            vsfList.TextMatrix(lngRows, 0) = strʱ���
                         End If
                        lngCols = lngCols + 1
                        If lngCols > vsfList.Cols - 1 Then vsfList.Cols = vsfList.Cols + 1
                        lngԤԼ���� = Nvl(!��������, 0) - lngԤԼ���� - lng������λ����
                        If vsfPlan.Cell(flexcpData, vsfPlan.Row, vsfPlan.ColIndex("����ҽ��")) <> "" And _
                            Format(!��ϸ��ʼʱ��, "yyyy-mm-dd hh:mm:ss") >= Format(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("���￪ʼʱ��")), "yyyy-mm-dd hh:mm:ss") And _
                            Format(!��ϸ��ʼʱ��, "yyyy-mm-dd hh:mm:ss") <= Format(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("������ֹʱ��")), "yyyy-mm-dd hh:mm:ss") Then
                          strData = "ԤԼ" & IIf(lngԤԼ���� < 0, 0, lngԤԼ����) & "��" & str���� & vbCrLf & _
                                                !��ʼʱ�� & "-" & !����ʱ��
                        Else
                          strData = "ԤԼ" & IIf(lngԤԼ���� < 0, 0, lngԤԼ����) & "��" & vbCrLf & _
                                                !��ʼʱ�� & "-" & !����ʱ��
                        End If
                        vsfList.TextMatrix(lngRows, lngCols) = strData
                        If lngԤԼ���� <= 0 Then
                             vsfList.Cell(flexcpForeColor, lngRows, lngCols) = vbGreen
                        End If
                        If Format(Datsys, "yyyy-mm-dd") <= Format(dtpAppointmentDate + 1, "yyyy-mm-dd") Then
                              If Format(DateAdd("n", mTy_Para.lngԤԼ����ʱ��, Datsys), "yyyy-mm-dd hh:mm:ss") > Format(!��ϸ����ʱ��, "yyyy-mm-dd hh:mm:ss") Then
                                vsfList.Cell(flexcpForeColor, lngRows, lngCols) = vbGrayText
                              End If
                        End If
                        If Format(!��ϸ����ʱ��, "yyyy-mm-dd") <> Format(dtpAppointmentDate.Value, "yyyy-mm-dd") And dtpAppointmentDate.Visible Then
                            vsfList.Cell(flexcpForeColor, lngRows, lngCols) = vbGrayText
                        End If
                     End If
                 End If
                .MoveNext
          Loop
          .Filter = ""
        End With
        Set rsʱ��ͳ�� = Nothing
    Case v_ר�Һŷ�ʱ��:
     '*******************************
     'ר�Һŷ�ʱ��
     'ÿ����ʱ�������
     '*******************************
     
regHD:
        blnInserted = False
        strʱ��� = ""
        With mrsʱ���
          mtyRegPlanState.lngLastNO = 0
          lngRows = -1: lngCols = 0
           datMax = CDate("00:00:00")
           Do While Not .EOF
                 If datMax < CDate(Nvl(!��ʼʱ��, "00:00:00")) Then datMax = CDate(!��ʼʱ��)
                'ԤԼ״̬ ֻ�������ԤԼ��ʱ���
                '�Һ�ʱ�����ֶ����
                If blnFindSN Then
                    If Val(Nvl(!���)) = lngFindSN And lngFindSN > 0 Then
                          lngCurrSn = lngFindSN
                    End If
                End If
'                If (mbytMode = 1 And Nvl(!�Ƿ�ԤԼ, 0) = 1 Or blnHave) Or mbytMode <> 1 Then
                '78643:���ϴ�,2014/10/16,�ҺŴ�ԤԼ�ĹҺŰ������������ԤԼ�ŶΣ�ֻ��ʾԤԼʱ�β���
                If ((mbytMode = 1 Or (mbytMode = 0 And chkBooking.Visible And chkBooking.Value = 1)) And Nvl(!�Ƿ�ԤԼ, 0) = 1 Or blnHave) Or _
                    Not (mbytMode = 1 Or (mbytMode = 0 And chkBooking.Visible And chkBooking.Value = 1)) Then
                    If strʱ��� <> Nvl(!ʱ���) Then
                        lngRows = lngRows + 1
                        strʱ��� = Nvl(!ʱ���)
                        If lngRows > vsfList.Rows - 1 Then vsfList.Rows = vsfList.Rows + 1: lngCols = 0
                        If lngCols > vsfList.Cols - 1 Then vsfList.Cols = vsfList.Cols + 1
                        vsfList.TextMatrix(lngRows, 0) = strʱ���
                        vsfList.Cell(flexcpForeColor, lngRows, 0, lngRows, 0) = vsfPlan.Cell(flexcpForeColor, vsfPlan.Row, 0, vsfPlan.Row, 0)
                     End If
                    lngCols = lngCols + 1
                      If lngCols > vsfList.Cols - 1 Then vsfList.Cols = vsfList.Cols + 1
                    If vsfPlan.Cell(flexcpData, vsfPlan.Row, vsfPlan.ColIndex("����ҽ��")) <> "" And _
                        Format(!��ϸ��ʼʱ��, "yyyy-mm-dd hh:mm:ss") >= Format(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("���￪ʼʱ��")), "yyyy-mm-dd hh:mm:ss") And _
                        Format(!��ϸ��ʼʱ��, "yyyy-mm-dd hh:mm:ss") <= Format(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("������ֹʱ��")), "yyyy-mm-dd hh:mm:ss") Then
                      strData = !��� & str���� & vbCrLf & !��ʼʱ�� & "-" & !����ʱ��
                    Else
                      strData = !��� & vbCrLf & !��ʼʱ�� & "-" & !����ʱ��
                    End If
                    vsfList.TextMatrix(lngRows, lngCols) = strData
                    
                    Select Case mbytMode
                    Case 0:
                        If chkBooking.Visible And chkBooking.Value = 1 Then
                            If Format(Datsys, "yyyy-mm-dd") <= Format(dtpAppointmentDate + 1, "yyyy-mm-dd") Then
                               If (Format(DateAdd("n", mTy_Para.lngԤԼ����ʱ��, Datsys), "yyyy-mm-dd hh:mm:ss") > Format(!��ϸ��ʼʱ��, "yyyy-mm-dd hh:mm:ss")) Then
                                   vsfList.Cell(flexcpForeColor, lngRows, lngCols) = vbGrayText
                               End If
                             End If
                        ElseIf (Format(Datsys, "yyyy-mm-dd hh:mm:ss") > Format(!��ϸ��ʼʱ��, "yyyy-mm-dd hh:mm:ss") And mbytMode = 0) Then
                             vsfList.Cell(flexcpFontUnderline, lngRows, lngCols) = True
                             vsfList.Cell(flexcpForeColor, lngRows, lngCols) = vbGrayText
                        End If
                    Case 1:
                        If Format(Datsys, "yyyy-mm-dd") <= Format(dtpAppointmentDate + 1, "yyyy-mm-dd") Then
                            If (Format(DateAdd("n", mTy_Para.lngԤԼ����ʱ��, Datsys), "yyyy-mm-dd hh:mm:ss") > Format(!��ϸ��ʼʱ��, "yyyy-mm-dd hh:mm:ss")) Then
                                vsfList.Cell(flexcpForeColor, lngRows, lngCols) = vbGrayText
                            End If
                        End If
                    Case Else:
                    End Select
                    If Format(!��ϸ����ʱ��, "yyyy-mm-dd") <> Format(dtpAppointmentDate.Value, "yyyy-mm-dd") And dtpAppointmentDate.Visible Then
                        vsfList.Cell(flexcpForeColor, lngRows, lngCols) = vbGrayText
                    End If
                End If
                
                '�����õ�������ű��浽mtyRegPlanState�� �������ԱȻ����������� 'lgf
                If mtyRegPlanState.lngLastNO < Val(Nvl(!���)) Then
                    With mtyRegPlanState
                        .lngLastNO = Val(Nvl(mrsʱ���!���))
                        .lngLastNO_X = lngRows
                        .lngLastNO_Y = lngCols
                    End With
                    
                End If
                
                .MoveNext
          Loop
          If blnHave = False And vsfList.Rows = 1 And vsfList.Cols = 1 And mrsʱ���.RecordCount > 0 Then blnHave = True: mrsʱ���.MoveFirst: GoTo regHD
          
          '��ȡ���һ��ʱ�ε����,��ʼʱ��,����ʱ�� 'lgf
          mrsʱ���.Filter = 0
          If mrsʱ���.RecordCount > 0 And mtyRegPlanState.lngLastNO > 0 Then
                mrsʱ���.Filter = "���=" & mtyRegPlanState.lngLastNO
                If mrsʱ���.RecordCount > 0 Then
                    mtyRegPlanState.strLastNO_Time = Nvl(!��ʼʱ��)
                    mtyRegPlanState.strLastNo_EndTime = Nvl(!����ʱ��)
                End If
                mrsʱ���.Filter = 0
          End If
          If InStr(mstrPrivs, ";�Ӻ�;") > 0 And mbytMode = 0 Then
            If (Format(Nvl(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("��������")), "3000-01-01"), "yyyy-mm-dd") = Format(dtpAppointmentDate.Value, "yyyy-MM-dd")) Or dtpAppointmentDate.Visible = False Then
                .MoveLast
                For i = 1 To vsfList.Cols - 1
                    If vsfList.TextMatrix(vsfList.Rows - 1, i) = "" Then
                        If blnInserted = False Then
                            vsfList.TextMatrix(vsfList.Rows - 1, i) = " " & vbCrLf & !����ʱ�� & "�Ժ�"
                            vsfList.Cell(flexcpData, vsfList.Rows - 1, i) = "�Ӻ�"
                            blnInserted = True
                        End If
                    End If
                Next i
                If blnInserted = False Then
                    vsfList.Cols = vsfList.Cols + 1
                    vsfList.TextMatrix(vsfList.Rows - 1, vsfList.Cols - 1) = " " & vbCrLf & !����ʱ�� & "�Ժ�"
                    vsfList.Cell(flexcpData, vsfList.Rows - 1, vsfList.Cols - 1) = "�Ӻ�"
                End If
            End If
          End If
        End With
    End Select
    dtpAppointmentTime.Tag = Format(datMax, "hh:mm:ss")
    '***************************************
    '��ű��״̬����
    '***************************************
    Call SetSnStyle(True)
    '***************************************
    '���״̬ ���
    '���ڹҺ�״̬��Ҫ����ֻ��һ��״̬
    '***************************************
     If mViewMode = v_ר�Һŷ�ʱ�� Then
        If picBookingDate.Visible Or mbytMode = 1 Or mbytMode = 2 Or (mbytMode = 0 And mbytInState = 0 And mstrNoIn <> "") Then             'ԤԼ�����ʱ������
            datThis = CDate(Format(dtpAppointmentDate.Value, "yyyy-MM-dd"))
        Else
            datThis = zlDatabase.Currentdate
        End If
         
         If mTy_Para.blnʧԼ���ڹҺ� Then
            'ר�Һŷ�ʱ��ʱ  ʧԼ��������ڿ��ų����Һ�
            blnʧԼ���ڹҺ� = True
            Datsys = DateAdd("n", -1 * mTy_Para.lngԤԼ��Чʱ��, Datsys)
         End If
        
        Set mrsSNState = GetSNState(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("��¼ID")))

        If mrsSNState.RecordCount > 0 Then
                For i = 0 To vsfList.Rows - 1
                   For j = 1 To vsfList.Cols - 1
                       If vsfList.TextMatrix(i, j) <> "" And Not vsfList.Cell(flexcpData, i, j) Like "��*" Then
                        '**********************************************
                        '
                        '**********************************************
                          vsfList.Row = i: vsfList.Col = j
                          lngFindSN = Val(Getʱ��(i, j, False))
                          mrsSNState.Filter = "���=" & lngFindSN
                          If mrsSNState.RecordCount > 0 Then
                            If lngCurrSn = lngFindSN Then lngCurrSn = -1
                            Select Case mrsSNState!״̬
                            Case 1  '�ѹ�
                                  If Nvl(mrsSNState!ԤԼ, "0") = "0" Then
                                    vsfList.Cell(flexcpForeColor, i, j) = vbRed
                                  Else
                                    vsfList.Cell(flexcpForeColor, i, j) = &HC000C0
                                  End If
                                  vsfList.Cell(flexcpFontStrikethru, i, j) = True
                            Case 2  '��Լ
                                vsfList.Cell(flexcpForeColor, i, j) = vbGreen
                            If lngMaxSn < Val(Nvl(mrsSNState!���)) Then
                                lngMaxSn = Val(Nvl(mrsSNState!���))
                            End If
                            Case 3  '����
                              vsfList.Cell(flexcpForeColor, i, j) = vbBlue
                            Case 4  '�˺�
'                                If mTy_Para.blnReuseCancelNO = False Then
                                    vsfList.Cell(flexcpForeColor, i, j) = vbGrayText
                                    vsfList.Cell(flexcpFontStrikethru, i, j) = True
'                                End If
                            Case 5  '����
                                vsfList.Cell(flexcpForeColor, i, j) = vbRed
                            Case 6  'ͣ��
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
                        mrsUnitReg.Filter = "���=" & Getʱ��(i, j, False)
                        If mrsUnitReg.RecordCount > 0 Then vsfList.Cell(flexcpForeColor, i, j) = &HC000C0
                    End If
                Next
            Next
            mrsUnitReg.Filter = 0
        End If
     End If
     '���п�����ŵ�����£����μӺ���
    If CheckAddAvailable = False Then
        For i = 0 To vsfList.Rows - 1
            For j = 1 To vsfList.Cols - 1
                If vsfList.Cell(flexcpData, i, j) Like "��*" Then
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
    locateSnByʱ�� lngCurrSn, True
    mblnStateChange = False
    Exit Sub
Hd:
    If ErrCenter() = 1 Then
        Resume
    End If
    SaveErrLog
End Sub

Private Sub locateSnByʱ��(Optional ByVal lngSN As Long = -1, _
    Optional blnǿ�ƶ�λ As Boolean)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��λ��ָ����ʱ��
    '���:lngSN:>0��Ҫ��λ�������,-1:��ʾ������ȡ��
    '����:blnǿ�ƶ�λ-ǿ�ƶ�λ��ָ������������
    '����:���˺�
    '����:2013-12-07 13:01:55
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, j As Long, lngRow As Long, lngCol As Long
    Dim blnFind  As Boolean, blnExit As Boolean, blnMaxSn As Boolean
    Dim lngLastRow As Long, lngLastCol As Long
     lngRow = 0: lngCol = 1
     
    vsfList.HighLight = flexHighlightAlways
    Select Case mViewMode
    Case V_��ͨ�ŷ�ʱ��:
         '****************************
         '��ͨ�ŷ�ʱ�� ��Ŷ�λ
         '****************************
         vsfList.Redraw = False
         blnMaxSn = True
          For i = 0 To vsfList.Rows - 1
            For j = 1 To vsfList.Cols - 1
                With vsfList
                    If .TextMatrix(i, j) <> "" Then
                        If .Cell(flexcpForeColor, i, j) <> vbGrayText Then
                            If Val(Getʱ��(i, j, False)) > 0 Then
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
        
        dtpAppointmentTime.Value = IIf(blnFind, CDate(Getʱ��(lngRow, lngCol, True)), CDate(mdatLast))
        vsfList.Redraw = True
    Case v_ר�Һŷ�ʱ��:
        blnMaxSn = True
        With vsfList
            For i = 0 To .Rows - 1
                For j = 1 To .Cols - 1
                    If .TextMatrix(i, j) <> "" Then
                        'Ԥ��
                        If .Cell(flexcpForeColor, i, j) = vbBlue Then
                            If lngSN <> -1 Then
                                 If lngSN = Val(Getʱ��(i, j, False)) Then
                                    .Row = i: .Col = j
                                     blnFind = True
                                     lngRow = i: lngCol = j
                                     blnMaxSn = False
                                     dtpAppointmentTime.Value = CDate(Getʱ��(i, j, True))
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
                                If Not mTy_Para.bln������ѡ�� Or lngSN = -1 Then  '66788
                                    blnFind = True
                                    lngRow = i: lngCol = j
                                    blnMaxSn = False
                                    If mbytMode <> 1 Then
                                        blnExit = True: Exit For  '45768
                                    End If
                                End If
                             End If
                             
                             If lngSN <> -1 Then
                                 If lngSN = Val(Getʱ��(i, j, False)) Then
                                    .Row = i: .Col = j
                                     blnFind = True
                                    lngRow = i: lngCol = j
                                    blnMaxSn = False
                                     dtpAppointmentTime.Value = CDate(Getʱ��(i, j, True))
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
            If blnǿ�ƶ�λ Then mblnNotClick = True
            vsfList.Row = lngRow: vsfList.Col = lngCol
            mblnNotClick = False
        Else
            vsfList.HighLight = flexHighlightAlways
        End If
        If blnFind = False And blnMaxSn And Me.dtpAppointmentTime.Tag <> "" Then
            dtpAppointmentTime.Value = Format(CDate(Me.dtpAppointmentTime.Tag), "hh:mm:ss")
        Else
            dtpAppointmentTime.Value = Format(CDate(Getʱ��(lngRow, lngCol, True)), "hh:mm:ss")
            If dtpAppointmentDate.Visible Then
                txt����ʱ��.Text = Format(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("��������")), "yyyy-MM-dd") & " " & Format(Getʱ��(lngRow, lngCol, True), "hh:mm:ss")
            Else
                txt����ʱ��.Text = Format(zlDatabase.Currentdate, "yyyy-MM-dd") & " " & Format(Getʱ��(lngRow, lngCol, True), "hh:mm:ss")
            End If
        End If
        If blnǿ�ƶ�λ = False Then Call vsfList_DblClick
    Case Else: Exit Sub
    End Select
End Sub
Private Function Getʱ��(ByVal lngRow As Long, ByVal lngCol As Long, Optional ByVal blnTime As Boolean = False, Optional ByVal blnLastTime As Boolean = False) As String
    '*****************************************************************
    '����˵��:�ڹҺ�ר�Һŷ�ʱʱ ��ȡ ���,���� ��ʼʱ��
    '����:  blntime �Ƿ��ȡʱ�� �����ȡʱ��  ���򷵻����
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
            Getʱ�� = Split(Split(Replace(vsfList.TextMatrix(lngRow, lngCol), "(��)", ""), vbCrLf)(1), "-")(i)
        Else
            Getʱ�� = Split(Split(Replace(vsfList.TextMatrix(lngRow, lngCol), "(��)", ""), vbCrLf)(1), "��")(i)
        End If
        Exit Function
    End If
    If mViewMode = v_ר�Һŷ�ʱ�� Then
       strResult = Split(Replace(vsfList.TextMatrix(lngRow, lngCol), "(��)", ""), vbCrLf)(0)
    ElseIf mViewMode = V_��ͨ�ŷ�ʱ�� Then
       strResult = Replace(Replace(Split(Replace(vsfList.TextMatrix(lngRow, lngCol), "(��)", ""), vbCrLf)(0), "ԤԼ", ""), "����", "")
    End If
    Getʱ�� = strResult
End Function

Private Sub ClearRegState()
    Dim strSQL As String, rsTemp As ADODB.Recordset
    '��ʼ��״̬������Ϣ
    'lgf 2012-10-30
   '��ʼ��״̬������Ϣ
    With mtyRegPlanState
        .str�ű� = "" 'ѡ�еĺű�
        .lngLastNO = 0 '����һ�����
        .strLastNO_Time = "" '���һ��ʱ�ο�ʼʱ��
        .strLastNo_EndTime = "" '����һ��ʱ�ν���ʱ��
        .blnAdditionalNumber = False '�Ƿ��Ѿ�׷����� '׷����ŵ��ص�(�ҳ�ȥ�����,��Ŵ������õ�������,����ʱ����ڻ��ߵ���,���һ��ʱ�εĽ���ʱ��)
        .lngSelX = 0 'ѡ�е���
        .lngSelY = 0 'ѡ�е���
        .lngSelNO = 0  'ѡ�е����
        .strSelTime = ""   'ѡ�е���Ŷ�Ӧʱ�εĿ�ʼʱ��
        .bln��ſ��� = False    '��ſ���
        .lng�޺��� = 0             '�޺���
        .lng��Լ�� = 0             '��Լ��
        .lngLastNO_X = 0 '���һ����ŵ�λ��
        .lngLastNO_Y = 0
        '.lngPlanRow = 0 '�ű�������
    End With
    '73767
    If mTy_Para.blnʧԼ���ڹҺ� = True And mTy_Para.lngԤԼ��Чʱ�� <> 0 Then
        '�����:110549,����,2017/07/21,SQL��������
        strSQL = "Select 1" & vbNewLine & _
                " From ���˹Һż�¼ A, �ٴ�������ſ��� B" & vbNewLine & _
                " Where a.ԤԼʱ�� < Sysdate + 1 / 24 / 60 * " & mTy_Para.lngԤԼ��Чʱ�� & " And a.ԤԼʱ�� > Trunc(Sysdate) And a.��¼���� = 2 And" & vbNewLine & _
                "       a.�����¼Id = b.��¼Id And a.�����¼Id = [1] And (a.���� = b.��� Or to_Char(a.����) = b.��ע) And Nvl(b.�Һ�״̬,0) = 2 And rownum < 2"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("��¼ID"))))
        If Not rsTemp.EOF Then
            Call zlDatabase.ExecuteProcedure("zl_�Һ����״̬_����_DELETE(" & Val(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("��¼ID"))) & ")", Me.Caption)
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
    Dim lngMaxSn    As Long 'ԤԼ�����ʹ�ú�
    Dim lngRow      As Long
    Dim lngCol      As Long
    Dim blnChk      As Boolean
    Dim sngTime     As Single
    Dim DatCur      As Date
    If Me.Visible = False Then GoTo regTab
    
    '125595:���ϴ���2018/5/16��ԤԼ���ճ����¼��λ����
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
    '����ʱ��Ҫ����,������ʾ,��Ϊ������Ҫ�޸����
    If mbytInState <> 0 Then
        Exit Sub
    End If
   
    dtpAppointmentTime.MaxDate = CDate("23:59:59")
    dtpAppointmentTime.MinDate = CDate("00:00:00")
    
    DatCur = zlDatabase.Currentdate
    If mbytMode = 1 Or chkBooking.Value = 1 Then
        txt����ʱ��.Text = Format(Format(dtpAppointmentDate.Value, "yyyy-mm-dd" & " " & Format(dtpAppointmentTime.Value, "hh:mm:ss")), "yyyy-mm-dd hh:mm:ss")
    Else
        txt����ʱ��.Text = Format(DatCur, "yyyy-mm-dd hh:mm:ss")
    End If
    
    
    '��ʱֻ�����ʱ���������,��Ҫ����,��ʱ���и���ʱ��,����ʱ�ε���ź�ʱ�ε�ʱ��Բ��ϵ����,
    '��ʼ��������Ϣ
    Call ClearRegState
    
    mtyRegPlanState.str�ű� = vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("�ű�"))
    
    '*****************************
    '��ȡʹ���������̴���Һ�
    '******************************
    If mcustomTime = t_ʱ�� Then
         GetActiveView
         If mcustomTime = t_��ͨ Then
            dtpAppointmentTime.Enabled = False
            dtpAppointmentTime.Visible = False
         Else
           If (mViewMode = V_��ͨ�ŷ�ʱ�� Or mViewMode = v_ר�Һŷ�ʱ��) Then
               dtpAppointmentTime.Enabled = False
              
           ElseIf (mbytMode = 1 Or (chkBooking.Visible And chkBooking.Value = 1)) And (mViewMode = V_��ͨ�� Or mViewMode = v_ר�Һ�) Then
                dtpAppointmentTime.Enabled = True
                Call SetDefaultRegistTime
           ElseIf mbytMode = 0 Then
               dtpAppointmentTime.Enabled = False
           End If
           
         End If
        If mViewMode = V_��ͨ�ŷ�ʱ�� Or mViewMode = v_ר�Һŷ�ʱ�� Then
           If mbytMode = 1 And mblnUnitReg Then
                '�����ԤԼͬʱ�����˹Һź�����λ��Ϣ�Ļ�����ȼ��� ������λ����Ϣ
                LoadUnitReg (Val(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("��¼ID"))))
            End If
           '*************************************************
           '������ڷ�ʱ�ε���� ʹ�÷�ʱ�εĴ�����
           '*************************************************
           LoadTimePlan
           SetDefaultRegistTime
           Call picPlan_Resize
'           vsfList.Height = picPlan.ScaleHeight - vsfList.Top - 350
'           vsfPlan.Height = picPlan.ScaleHeight - IIf(picBookingDate.Visible, picBookingDate.Height + 30, 0) - 360 - IIf(vsfList.Visible, vsfList.Height + picSplit.Height + 30, 0)
'           Call locateSnByʱ��(, True)
           vsfPlan.ShowCell vsfPlan.Row, vsfPlan.Col
           If InStr(vsfList.TextMatrix(vsfList.Row, vsfList.Col), "��") > 0 Then
                vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("����ҽ��")) = vsfPlan.Cell(flexcpData, vsfPlan.Row, vsfPlan.ColIndex("����ҽ��"))
           Else
                vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("����ҽ��")) = ""
           End If
           Exit Sub
        End If
    Else
         If vsfPlan.TextMatrix(vsfPlan.Row, GetCol("��ſ���")) <> "" Then
                mViewMode = v_ר�Һ�
         Else
                mViewMode = V_��ͨ��
         End If
    End If
    
    If vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("���￪ʼʱ��")) <> "" And Not (mViewMode = V_��ͨ�ŷ�ʱ�� Or mViewMode = v_ר�Һŷ�ʱ��) Then
        If CDate(txt����ʱ��.Text) >= CDate(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("���￪ʼʱ��"))) And CDate(txt����ʱ��.Text) <= CDate(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("������ֹʱ��"))) Then
            vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("����ҽ��")) = vsfPlan.Cell(flexcpData, vsfPlan.Row, vsfPlan.ColIndex("����ҽ��"))
        Else
            vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("����ҽ��")) = ""
        End If
    End If
    
    If mbytMode = 1 And mblnUnitReg Then
        '�����ԤԼͬʱ�����˹Һź�����λ��Ϣ�Ļ�����ȼ��� ������λ����Ϣ
        LoadUnitReg (Val(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("��¼ID"))))
    End If
    vsfList.Redraw = False
    vsfList.Clear
    If mbytMode = 1 Then
        lngMax = Val(vsfPlan.TextMatrix(vsfPlan.Row, GetCol("��Լ")))
        If lngMax = 0 Then lngMax = Val(vsfPlan.TextMatrix(vsfPlan.Row, GetCol("�޺�")))
    Else
        lngMax = Val(vsfPlan.TextMatrix(vsfPlan.Row, GetCol("�޺�"))) '�ҽ����ĺŲ�����ԤԼ,��Ϊ�ѽ���,Ӧ���ɹҺ�
    End If
    If lngMax > 0 And vsfPlan.TextMatrix(vsfPlan.Row, GetCol("��ſ���")) <> "" Then
        If mbytMode = 1 Then
              lngMax = Val(vsfPlan.TextMatrix(vsfPlan.Row, GetCol("�޺�")))  'ԤԼ���ų���,�û�ѡ��:����????
        End If
        If lngMax = 0 Then GoTo regTab
        '1.����λ��
        If lngMax > 1000 Then
            vsfList.FontWidth = 4
        Else
            vsfList.FontWidth = 0 '�ָ�ȱʡ����
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
            mcbrToolBar.Controls.Find(xtpControlButton, 2605).Visible = InStr(1, mstrPrivs, ";Ԥ������;") > 0
            mcbrToolBar.Controls.Find(xtpControlButton, 2604).Visible = InStr(1, mstrPrivs, ";Ԥ������;") > 0
            Call Form_Resize
        Else
            mcbrToolBar.Controls.Find(xtpControlButton, 2605).Visible = InStr(1, mstrPrivs, ";Ԥ������;") > 0
            mcbrToolBar.Controls.Find(xtpControlButton, 2604).Visible = InStr(1, mstrPrivs, ";Ԥ������;") > 0
            Call Form_Resize
        End If
                                
        '������
        lngThis = 1
        For i = 0 To vsfList.Rows - 1
            For j = 0 To vsfList.Cols - 1
                vsfList.TextMatrix(i, j) = lngThis
                lngThis = lngThis + 1
                If lngThis > lngMax Then Exit For
            Next
            If lngThis > lngMax Then Exit For
        Next
             
        If picBookingDate.Visible Or mbytMode = 1 Or mbytMode = 2 Or (mbytMode = 0 And mbytInState = 0 And mstrNoIn <> "") Then             'ԤԼ�����ʱ������
            datThis = CDate(Format(dtpAppointmentDate.Value, "yyyy-MM-dd"))
        Else
            datThis = CDate(Format(zlDatabase.Currentdate, "yyyy-MM-dd"))
        End If
        
        
        Set mrsSNState = GetSNState(Val(vsfPlan.TextMatrix(vsfPlan.Row, GetCol("��¼ID"))))
        lngMaxSn = 0
        For i = 0 To mrsSNState.RecordCount - 1
            If mrsSNState!��� <= lngMax Then
                If (mrsSNState!��� \ SNCOLS) * SNCOLS = mrsSNState!��� Then
                   lngRow = (mrsSNState!��� \ SNCOLS) - 1
                   lngRow = IIf(lngRow < 0, 0, lngRow) '�����:51843
                Else
                    lngRow = (mrsSNState!��� \ SNCOLS)
                End If
                    lngCol = (mrsSNState!��� - 1) Mod SNCOLS
                    lngCol = IIf(lngCol < 0, 0, lngCol) '�����:51843
                Select Case mrsSNState!״̬
                    Case 1  '�ѹ�
                       If Nvl(mrsSNState!ԤԼ, "0") = "0" Then
                          vsfList.Cell(flexcpForeColor, lngRow, lngCol) = vbRed
                          '������Ŷ�λ������Ч�ź�
                          If lngMaxSn < Val(Nvl(mrsSNState!���)) Then
                            lngMaxSn = Val(Nvl(mrsSNState!���))
                          End If
                       Else
                          'ԤԼ����
                          vsfList.Cell(flexcpForeColor, lngRow, lngCol) = &HC000C0
                       End If
                    Case 2  '��Լ
                          vsfList.Cell(flexcpForeColor, lngRow, lngCol) = vbGreen
                    Case 3  '����
                      vsfList.Cell(flexcpForeColor, lngRow, lngCol) = vbBlue
                    Case 4  '�˺�
'                        If mTy_Para.blnReuseCancelNO = False Then
                            vsfList.Cell(flexcpForeColor, lngRow, lngCol) = vbGrayText
                            vsfList.Cell(flexcpFontStrikethru, lngRow, lngCol) = True
'                        End If
                    Case 5  '����
                        vsfList.Cell(flexcpForeColor, lngRow, lngCol) = vbRed
                End Select
            End If
            mrsSNState.MoveNext
        Next
        
        If mbytMode = 1 And mblnUnitReg And Not mrsUnitReg Is Nothing Then
            For i = 0 To vsfList.Rows - 1
                For j = 0 To vsfList.Cols - 1
                    If Trim(vsfList.TextMatrix(i, j)) <> "" Then
                        mrsUnitReg.Filter = "���=" & vsfList.TextMatrix(i, j)
                        If mrsUnitReg.RecordCount > 0 Then
                            vsfList.Cell(flexcpForeColor, i, j) = &HC000C0
                            If lngMaxSn < Val(Trim(vsfList.TextMatrix(i, j))) Then lngMaxSn = Val(Trim(vsfList.TextMatrix(i, j)))
                        End If
                    End If
                Next
            Next
            mrsUnitReg.Filter = 0
        End If
        
        If Trim(txtSN.Text) = "" Then  '��ʱˢ��ʱ��������Ĳ���
           lngCurrSn = GetCurrSN(IIf(mbytMode = 0, lngMaxSn, -1))
           txtSN.Text = lngCurrSn
        Else
            lngCurrSn = Val(txtSN.Text)
            '���������ţ�38779
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

Private Sub LoadUnitReg(ByVal lng��¼ID As Long)
 '���عҺź�����λ������Ϣ
    Dim strSQL As String
        
    strSQL = "Select ���� As ������λ, ���Ʒ�ʽ, ���, ���� From �ٴ�����Һſ��Ƽ�¼ Where ��¼id = [1] And ���� = 1"
    
    On Error GoTo Hd
    Set mrsUnitReg = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng��¼ID)
    Exit Sub
Hd:
    If ErrCenter() = 1 Then
        Resume
    End If
    SaveErrLog
End Sub

Private Sub LocateSN(lngCurrSn As Long)
'����:��λ��ָ�������
'     �����������ű�����,����ű��ý���
    Dim lngRow          As Long
    Dim i               As Long
    Dim j               As Long
    Dim blnHave         As Boolean
    If lngCurrSn = 0 Then Exit Sub
   
    If mViewMode = V_��ͨ�� Or mViewMode = v_ר�Һ� Then
        '************************************************
        '����ʱ�� ��Ŷ�λ���ǰ�����ǰ�ķ�ʽ
        '************************************************
        If (lngCurrSn \ SNCOLS) * SNCOLS = lngCurrSn Then
            lngRow = (lngCurrSn - 1) \ SNCOLS
        Else
            lngRow = (lngCurrSn \ SNCOLS)
        End If
        If Not vsfList.RowIsVisible(lngRow) Then
            If lngRow >= 1 Then  '������һ�пɼ�
                vsfList.TopRow = lngRow - 1
            Else
                vsfList.TopRow = lngRow
            End If
        End If
        '�����:52335
        mblnNotClick = True
        vsfList.Row = lngRow
        vsfList.RowSel = vsfList.Row
        vsfList.Col = (lngCurrSn - 1) Mod SNCOLS
        vsfList.ColSel = vsfList.Col
        '�����:52335
        mblnNotClick = False
     
    ElseIf mViewMode = v_ר�Һŷ�ʱ�� Then
        '*******************************************
        'ר�Һŷ�ʱ�� ��Ŷ�λ
        '*******************************************
        For i = 0 To vsfList.Rows - 1
            For j = 1 To vsfList.Cols - 1
               If vsfList.TextMatrix(i, j) <> "" Then
                    If lngCurrSn = Val(Getʱ��(i, j, False)) Then
                     If Not vsfList.RowIsVisible(i) Then
                        If lngRow >= 1 Then  '������һ�пɼ�
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
                     dtpAppointmentTime.Value = CDate(Getʱ��(i, j, True))
                     Exit For
                      
                     
                    End If
                End If
            Next
            If blnHave Then Exit For
        Next
    End If
    Call vsfList_EnterCell
    If vsfList.Visible And vsfList.Enabled _
                And Not Me.ActiveControl Is txt�ű� And Not Me.ActiveControl Is txtSN _
                And Not Me.ActiveControl Is dtpAppointmentDate And Not Me.ActiveControl Is vsfPlan Then Call vsfList.SetFocus     '�����ںű�������������
End Sub

Private Function GetSNState(lng��¼ID As Long, Optional lngSN As Long) As ADODB.Recordset
    Dim strSQL           As String
    On Error GoTo errH

    strSQL = "    " & vbNewLine & " Select A.���,Decode(�Ƿ�ͣ��,1,6,Nvl(A.�Һ�״̬,0)) As ״̬,A.����Ա����,Decode(A.�Һ�״̬,2,1,0) as ԤԼ,To_Char(B.��������,'hh24:mi:ss') as ����  "
    strSQL = strSQL & vbNewLine & " From �ٴ�������ſ��� A, �ٴ������¼ B "
    strSQL = strSQL & vbNewLine & " Where B.ID=[1] And B.ID=A.��¼ID"
    strSQL = strSQL & vbNewLine & IIf(lngSN > 0, " And A.���=[2]", "")
    Set GetSNState = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng��¼ID, lngSN)

    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Sub vsfPlan_LeaveCell()
    Call SetvsfplanColor
End Sub

Private Sub vsfPlan_KeyDown(KeyCode As Integer, Shift As Integer)
    'ѡ��ű���йҺ�
    If KeyCode = 13 Then
        
        If CheckNoValied(vsfPlan.Row) = False Then
             txt�ű�.Text = "": txt�ű�.SetFocus: Exit Sub
        End If
        vsfPlan.Tag = vsfPlan.TextMatrix(vsfPlan.Row, GetCol("�ű�"))
        If txt�ű�.Visible And txt�ű�.Enabled Then txt�ű�.SetFocus
        If txt�ű�.Text = vsfPlan.Tag Then
            Call txt�ű�_Change
        Else
            txt�ű�.Text = vsfPlan.Tag
        End If
    vsfPlan.Tag = ""
    Call locateSnByʱ��(, True)
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
    
    If mTy_Para.bln�����ͷ���� = False Then Exit Sub
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
'����:��鵱ǰѡ��ĺű�Ӻ��Ƿ����
'����:���÷���True,�����÷���False
'����:������
'����:2014-01-15
'��ע:
'-----------------------------------------------------------------------------------------------------------------------
    Dim i As Integer, j As Integer
    Dim intTotal As Integer, intUse As Integer
    If vsfList.Visible = False Then Exit Function
    intTotal = 0
    intUse = 0
    'ֻ�Է�ʱ�ν��д���
    If mViewMode = V_��ͨ�ŷ�ʱ�� Or mViewMode = v_ר�Һŷ�ʱ�� Or mViewMode = V_��ͨ�� Or mViewMode = v_ר�Һ� Then
        With vsfList
            For j = 1 To .Cols - 1
                For i = 0 To .Rows - 1
                    If .TextMatrix(i, j) <> "" And Not .Cell(flexcpData, i, j) Like "��*" Then
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
    '�����:52203
    '�����:52335
   
    If mblnNotClick Then Exit Sub
    If (mViewMode = V_��ͨ�ŷ�ʱ�� Or mViewMode = v_ר�Һŷ�ʱ�� Or mViewMode = v_ר�Һ�) And mTy_Para.bln������ѡ�� = False _
        And Not (mbytMode = 1 Or chkBooking.Value = 1 And chkBooking.Visible) And vsfList.Cell(flexcpForeColor, NewRow, NewCol) <> vbBlue Then
        Cancel = True
        Exit Sub
    End If
    If vsfList.TextMatrix(NewRow, NewCol) = "" Then Cancel = True
    If vsfList.Cell(flexcpForeColor, NewRow, NewCol) <> vbBlack And vsfList.Cell(flexcpForeColor, NewRow, NewCol) <> vbBlue Then Cancel = True
    If Not CheckAddAvailable And mbytMode = 0 Then
        If vsfList.Cell(flexcpData, NewRow, NewCol) Like "��*" Then Cancel = True
    End If
'    'vsflist.Cell(flexcpBackColor, OldRow, OldCol) = vbWhite
'    'vsflist.Cell(flexcpBackColor, NewRow, NewCol) = &HECBAAA
End Sub

Private Sub vsfList_DblClick()
    Dim lngSN       As Long
    Dim datThis     As Date
    Dim strTmp      As String
    If mViewMode = V_��ͨ�� Or mViewMode = v_ר�Һ� Then
        '*************************************************
        '��ͨ�ź�û�з�ʱ�ε�ר�Һ� ������ǰ������
        '*************************************************
        lngSN = Val(vsfList.TextMatrix(vsfList.Row, vsfList.Col))
        If Not mrsSNState Is Nothing And lngSN > 0 Then
            mrsSNState.Filter = "���=" & lngSN & " And ״̬ <> 0"
            If mrsSNState.RecordCount > 0 Then
                If mrsSNState!״̬ = 3 And mrsSNState!����Ա���� = UserInfo.���� Then
                    '����Ԥ���Ŀ���ֱ�������Һ�
                    vsfPlan.Tag = vsfPlan.TextMatrix(vsfPlan.Row, GetCol("�ű�"))
                    txt�ű�.Text = vsfPlan.Tag
                    txtSN.Text = lngSN
                    mstrPre�ű� = txt�ű�.Text
                    mlngPreRow = vsfPlan.Row
                    vsfPlan.Tag = ""
                  If mcustomTime = t_��ͨ Or dtpAppointmentTime.Enabled = False Then
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
                vsfPlan.Tag = vsfPlan.TextMatrix(vsfPlan.Row, GetCol("�ű�"))
                txt�ű�.Text = vsfPlan.Tag
                txtSN.Text = lngSN
                vsfPlan.Tag = ""
                mstrPre�ű� = txt�ű�.Text
                mlngPreRow = vsfPlan.Row
                If mcustomTime = t_��ͨ Or dtpAppointmentTime.Enabled = False Then
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
    
    If mViewMode = V_��ͨ�� Or mViewMode = v_ר�Һ� Then Exit Sub
    
    '*************************************************
    '��ʱ�� �����µķ�ʽ������
    '*************************************************
    
    Select Case mViewMode
    Case V_��ͨ�ŷ�ʱ��:
        If vsfList.CellForeColor = vbGrayText Then Exit Sub
        If vsfList.TextMatrix(vsfList.Row, vsfList.Col) = "" Then Exit Sub
        If Val(Getʱ��(vsfList.Row, vsfList.Col, False)) = 0 Then Exit Sub
        strTmp = Format(dtpAppointmentDate.Value, "yyyy-MM-dd") & " " & Getʱ��(vsfList.Row, vsfList.Col, True)
        txt����ʱ��.Text = Format(strTmp, "yyyy-mm-dd hh:mm:ss")
        datThis = CDate(Format(strTmp, "hh:mm:ss"))
        dtpAppointmentTime.Value = datThis
        dtpAppointmentTime.Tag = strTmp
        vsfPlan.Tag = vsfPlan.TextMatrix(vsfPlan.Row, GetCol("�ű�"))
        txt�ű�.Text = vsfPlan.Tag
        txtSN.Text = ""
        vsfPlan.Tag = ""
        
        '�������
        mtyRegPlanState.lngSelNO = 0
        mtyRegPlanState.lngSelX = vsfList.Row
        mtyRegPlanState.lngSelY = vsfList.Col
        mtyRegPlanState.strSelTime = Getʱ��(vsfList.Row, vsfList.Col, True)
        mstrPre�ű� = txt�ű�.Text
        mlngPreRow = vsfPlan.Row
        If InStr(vsfList.TextMatrix(vsfList.Row, vsfList.Col), "��") > 0 Then
            vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("����ҽ��")) = vsfPlan.Cell(flexcpData, vsfPlan.Row, vsfPlan.ColIndex("����ҽ��"))
        Else
            vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("����ҽ��")) = ""
        End If
        If txtSN.Enabled And txtSN.Visible Then
            txtSN.SetFocus
        Else
            If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
        End If
    Case v_ר�Һŷ�ʱ��:
        '**********************************************
        '������Ϊ�ѹһ�����Լ�Ĳ�����ѡ��
        '
        '**********************************************
        If vsfList.TextMatrix(vsfList.Row, vsfList.Col) = "" Then Exit Sub
        If vsfList.CellForeColor = vbRed Or vsfList.CellForeColor = vbGreen Or vsfList.CellForeColor = vbGrayText Or vsfList.CellForeColor = &HC000C0 Then Exit Sub  '--And .CellForeColor <> vbBlue
        If dtpAppointmentDate.Visible Then
            strTmp = Format(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("��������")), "yyyy-MM-dd") & " " & Getʱ��(vsfList.Row, vsfList.Col, True)
        Else
            strTmp = Format(zlDatabase.Currentdate, "yyyy-MM-dd") & " " & Getʱ��(vsfList.Row, vsfList.Col, True)
        End If
        txt����ʱ��.Text = Format(strTmp, "yyyy-mm-dd hh:mm:ss")
        datThis = CDate(strTmp)
        dtpAppointmentTime.Value = Getʱ��(vsfList.Row, vsfList.Col, True)
        dtpAppointmentTime.Tag = strTmp
        vsfPlan.Tag = vsfPlan.TextMatrix(vsfPlan.Row, GetCol("�ű�"))
        txt�ű�.Text = vsfPlan.Tag
        
        mblnNotChange = True
        txtSN.Text = Getʱ��(vsfList.Row, vsfList.Col, False)
        If txtSN.Text = "�Ӻ�" Then txtSN.Text = ""
        mtyRegPlanState.lngSelNO = Val(txtSN.Text)
        mtyRegPlanState.lngLastNO_X = vsfList.Row
        mtyRegPlanState.lngLastNO_Y = vsfList.Col
        mtyRegPlanState.strSelTime = Getʱ��(vsfList.Row, vsfList.Col, True)
        mblnNotChange = False
        
        mstrPre�ű� = txt�ű�.Text
        mlngPreRow = vsfPlan.Row
        vsfPlan.Tag = ""
        If InStr(vsfList.TextMatrix(vsfList.Row, vsfList.Col), "��") > 0 Then
            vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("����ҽ��")) = vsfPlan.Cell(flexcpData, vsfPlan.Row, vsfPlan.ColIndex("����ҽ��"))
        Else
            vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("����ҽ��")) = ""
        End If
        If txtSN.Enabled And txtSN.Visible Then txtSN.SetFocus
    Case Else
        Exit Sub
    End Select
     
End Sub

Private Sub vsfList_EnterCell()
'�����Ƿ�����Ԥ��
    '***************************************
    '���ﴦ��Ԥ����
    'Ԥ���Ŵ������Ϊ
    'ר�ҺŲ���ʱ�� ��ǰ�Ĵ���ʽ
    'ר�Һ� ��ʱ�� �´���ʽ
    '��ͨ�ŷ�ʱ�� ������Ԥ��
    '***************************************
    If mViewMode = V_��ͨ�ŷ�ʱ�� Then
        mcbrToolBar.Controls.Find(xtpControlButton, 2605).Enabled = False
        mcbrToolBar.Controls.Find(xtpControlButton, 2604).Enabled = False
        Exit Sub
    End If
    If vsfList.Row <> -1 Then
         '�����:52335
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
        '�����:52335
        If vsfList.Cols > vsfList.Col And vsfList.Rows > vsfList.Row Then
            Select Case mViewMode
            Case v_ר�Һ�:
                mrsSNState.Filter = "���=" & Val(vsfList.TextMatrix(vsfList.Row, vsfList.Col))
            Case v_ר�Һŷ�ʱ��:
                mrsSNState.Filter = "���=" & Val(Getʱ��(vsfList.Row, vsfList.Col, False))
            End Select
        End If
        If mrsSNState.RecordCount > 0 Then
            mrsSNState.MoveFirst
            If Val(Nvl(mrsSNState!״̬)) = 3 Then
                If mrsSNState!״̬ = 3 And mrsSNState!����Ա���� = UserInfo.���� Then
                    'ȡ��Ԥ��
                    mcbrToolBar.Controls.Find(xtpControlButton, 2605).Enabled = False
                    mcbrToolBar.Controls.Find(xtpControlButton, 2604).Enabled = True
                Else
                    mcbrToolBar.Controls.Find(xtpControlButton, 2605).Enabled = False
                    mcbrToolBar.Controls.Find(xtpControlButton, 2604).Enabled = False
                    '64184:������,2014-03-20,ѡ��Ԥ������
                    If Me.ActiveControl Is vsfList Then
                        Select Case mViewMode
                            Case v_ר�Һ�:
                                MsgBox Val(vsfList.TextMatrix(vsfList.Row, vsfList.Col)) & "���ѱ�" & mrsSNState!����Ա���� & "Ԥ��!�޷�ѡ��.", vbInformation, gstrSysName
                            Case v_ר�Һŷ�ʱ��:
                                MsgBox Val(Getʱ��(vsfList.Row, vsfList.Col, False)) & "���ѱ�" & mrsSNState!����Ա���� & "Ԥ��!�޷�ѡ��.", vbInformation, gstrSysName
                        End Select
                        txt�ű�_KeyPress (13)
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
     If mTy_Para.bln������ѡ�� Then Exit Sub
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
    
    '81103,Ƚ����,2014-12-26,¼�����֤�ź�,�������ڡ����䡢�Ա��ͬ���������͵���
    If txtIDCard.Visible And txtIDCard.Enabled And Not mobjfrmPatiInfo.mobjPubPatient Is Nothing Then
        'CheckPatiIdcard(ByVal strIdcard As String, Optional strBirthday As String, _
        '    Optional strAge As String, Optional strSex As String, Optional strErrInfo As String) As Boolean
        '���ܣ����֤����Ϸ���У��
        '��Σ�strIdCard ���֤����
        '���Σ�strBirthday  ��������TrueΪ��������
        '         strAge ��������TrueΪ����
        '         strSex ��������TrueΪ�Ա�
        '         strErrInfo ��������FalseΪ������Ϣ
        '���أ�True/False  ���֤�Ϸ�����True(�ɴ�strBirthday��strSex��ȡ�������ں��Ա�)��
        '       ���򷵻�False(�ɴ�strErrInfo��ȡ��ϸ������Ϣ)
        If mobjfrmPatiInfo.mobjPubPatient.CheckPatiIdcard(Trim(txtIDCard.Text), strbirthday, strAge, strSex, strErrInfo) Then
            '�²��˻������ҵ�����ݵ����в�����Ϣʱ��ʾ�Ƿ������һ�µĻ�����Ϣ
            If strSex <> NeedName(cbo�Ա�.Text) Then strInfo = "�Ա�"
            If strAge <> Trim(txt����.Text) & cbo���䵥λ Then strInfo = strInfo & IIf(strInfo = "", "����", "������")
            
            If strInfo <> "" Then
                If Trim(txtPatient.Text) = "" Then '67213,��������ݺ�����������ʱ,��Ӧ������,����ֱ�������֤�����Ա�����
                    Call zlControl.CboLocate(cbo�Ա�, strSex)
                    txt����.Text = ReCalcOld(CDate(strbirthday), cbo���䵥λ)
                    txt��������.Text = Format(strbirthday, "yyyy-mm-dd")
                    Call txt��������_Validate(False)
                Else
                    If MsgBox("�����" & strInfo & "�����֤�ŵ�" & strInfo & "��һ�£�" & _
                            "���������֤���޸�" & strInfo & "���Ƿ������", vbInformation + vbYesNo, gstrSysName) = vbYes Then
                        Call zlControl.CboLocate(cbo�Ա�, strSex)
                        txt����.Text = ReCalcOld(CDate(strbirthday), cbo���䵥λ)
                        txt��������.Text = Format(strbirthday, "yyyy-mm-dd")
                        Call txt��������_Validate(False)
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
    
    '�������,�϶���Ҫȥ����һ��,��������Ϣ���Ƿ���ڸ����֤�ŵĲ���:
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
    If txt�ű�.Text = "" Then KeyAscii = 0: Exit Sub
    If txtPatientPrint.Text <> "" And KeyAscii = vbKeyReturn Then
        If cbo�Ա�.Enabled And cbo�Ա�.Visible Then
            cbo�Ա�.SetFocus
        Else
            If cmdOK.Enabled And cmdOK.Visible Then cmdOK.SetFocus
        End If
    End If
End Sub

Private Sub txtPatientPrint_Validate(Cancel As Boolean)
    txtPatientPrint.Text = Trim(txtPatientPrint.Text)
End Sub

Private Sub txtSN_GotFocus()
    If (Not mTy_Para.bln������ѡ��) And mbytMode <> 1 Then
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
        ElseIf txt�ű�.Text = "" Or mrsSNState Is Nothing Then
            KeyAscii = 0
        End If
    Else
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub txtSN_Validate(Cancel As Boolean)
'����������ŵ���Ч��
    Dim i As Long, j As Long, blnHave As Boolean
    Dim lngSN As Long
    Dim blnʧЧ As Boolean
    Dim bln
    Dim blnLock As Boolean
    Dim blnLocateSn As Boolean
    Dim lngLocateSnX As Long
    Dim lngLocateSnY As Long
    Dim lngRow As Long, lngCol As Long
    If mblnNotChange Then Exit Sub
    If Val(txtSN.Text) = 0 Then txtSN.Text = ""
    If Trim(txtSN.Text) = "" Then Exit Sub
    If txtSN.Tag = txtSN.Text Then Exit Sub '����ԤԼʱû�б����ü��
    If Not IsNumeric(txtSN.Text) Then
        Cancel = True
        Call zlControl.TxtSelAll(txtSN)
        Exit Sub
    End If
    
    If Not vsfList.Visible Then Exit Sub
    If mViewMode = V_��ͨ�� Or mViewMode = v_ר�Һ� Then
        '**********************************************
        '����ʱ�� �ж� ������ǰ�ķ���
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
                MsgBox "�úű���δʹ����ţ��㲻��ʹ�üӺ���ţ�", vbInformation, gstrSysName
                txtSN.Text = ""
                Exit Sub
            End If
            If InStr(mstrPrivs, ";�Ӻ�;") <= 0 Then
                MsgBox lngSN & "�ų�������޺���!��û�����ź�����Һŵ�Ȩ��.", vbInformation, gstrSysName
                Cancel = True
                txtSN.Text = ""
            Else
                If MsgBox(lngSN & "�ų�������޺���!��ȷ��Ҫʹ����?", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Cancel = True
                    Call zlControl.TxtSelAll(txtSN)
                Else
                    If mbytMode = 0 Then
                        With vsfList
                            For i = 0 To .Rows - 1
                                For j = 0 To .Cols - 1
                                    If .Cell(flexcpData, i, j) Like "��*" Then .Select i, j
                                Next j
                            Next i
                        End With
                    End If
                End If
            End If
        ElseIf Not mrsSNState Is Nothing Then
            mrsSNState.Filter = "���=" & lngSN
            If mrsSNState.RecordCount > 0 Then
                If mrsSNState!״̬ = 1 Or mrsSNState!״̬ = 2 Then
                    Cancel = True
                    MsgBox lngSN & "���ѱ�" & IIf(mrsSNState!״̬ = 1, "ʹ��", "ԤԼ") & "!����������һ����.", vbInformation, gstrSysName
                    Call zlControl.TxtSelAll(txtSN)
                ElseIf mrsSNState!״̬ = 3 Then
                    If mrsSNState!����Ա���� = UserInfo.���� Then
                        If MsgBox(lngSN & "����Ԥ����!��ȷ��Ҫʹ����?", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            Cancel = True
                            txtSN.Text = ""
                            Call zlControl.TxtSelAll(txtSN)
                        Else
                            Call LocateSN(lngSN)
                        End If
                    Else
                        Cancel = True
                        MsgBox lngSN & "���ѱ�" & mrsSNState!����Ա���� & "Ԥ��!����������һ����.", vbInformation, gstrSysName
                        txtSN.Text = ""
                        Call zlControl.TxtSelAll(txtSN)
                    End If
                ElseIf mrsSNState!״̬ = 4 Then
                    If mTy_Para.blnReuseCancelNO = False Then
                        Cancel = True
                        MsgBox lngSN & "���ѱ��˺�,�޷��ٴ�ʹ��" & "!����������һ����.", vbInformation, gstrSysName
                        Call zlControl.TxtSelAll(txtSN)
                    End If
                ElseIf mrsSNState!״̬ = 5 Then
                    Cancel = True
                    MsgBox lngSN & "���ѱ�����������,�޷�ʹ��" & "!����������һ����.", vbInformation, gstrSysName
                    Call zlControl.TxtSelAll(txtSN)
                ElseIf mrsSNState!״̬ = 6 Then
                    Cancel = True
                    MsgBox lngSN & "���ѱ�ͣ��,�޷�ʹ��" & "!����������һ����.", vbInformation, gstrSysName
                    Call zlControl.TxtSelAll(txtSN)
                End If
            Else
                If blnHave And vsfList.Cell(flexcpForeColor, lngRow, lngCol) = &HC000C0 Then
                    Cancel = True
                    MsgBox lngSN & "�Ų�����!����������һ����.", vbInformation, gstrSysName
                    Call zlControl.TxtSelAll(txtSN)
                Else
                    Call LocateSN(lngSN)
                End If
            End If
        End If
    Else
        '*****************************************************
        '��ʱ�� ������
        'ֻ��ר�ҺŽ�����֤
        '��ͨ�ŷ�ʱ�� ������Ž�����֤
        '*****************************************************
        If mViewMode <> v_ר�Һŷ�ʱ�� Then Exit Sub
        lngSN = Val(txtSN.Text)
        For i = 0 To vsfList.Rows - 1
            For j = 1 To vsfList.Cols - 1
                If lngSN = Val(Getʱ��(i, j, False)) Then
                    lngLocateSnX = i
                    lngLocateSnY = j
                    blnHave = True
                    blnLock = vsfList.Cell(flexcpForeColor, i, j) = vbRed And vsfList.Cell(flexcpFontStrikethru, i, j) = False
                    blnʧЧ = vsfList.Cell(flexcpForeColor, i, j) = vbGrayText
                    Exit For
                End If
            Next
            If blnHave Then Exit For
        Next
        If blnLock Then
            MsgBox lngSN & "���Ѿ�������!�����������Ž��йҺ�.", vbInformation, gstrSysName
            Cancel = True
            txtSN.Text = ""
        End If
        If blnʧЧ Then
            MsgBox lngSN & "���Ѿ�ʧЧ!��������Ч�Ž��йҺ�.", vbInformation, gstrSysName
            Cancel = True
            txtSN.Text = ""
        End If
        If Not blnHave Then
            If Not CheckAddAvailable Then
                MsgBox "�úű���δʹ����ţ��㲻��ʹ�üӺ���ţ�", vbInformation, gstrSysName
                txtSN.Text = ""
                Call locateSnByʱ��(-1)
                Exit Sub
            End If
            If InStr(mstrPrivs, ";�Ӻ�;") <= 0 Then
                MsgBox lngSN & "�ų�������޺���!��û�����ź�����Һŵ�Ȩ��.", vbInformation, gstrSysName
                Cancel = True
                txtSN.Text = ""
            Else
                If MsgBox(lngSN & "�ų�������޺���!��ȷ��Ҫʹ����?", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Cancel = True
                    Call zlControl.TxtSelAll(txtSN)
                Else
                    If mbytMode = 0 Then
                        With vsfList
                            For i = 0 To .Rows - 1
                                For j = 0 To .Cols - 1
                                    If .Cell(flexcpData, i, j) Like "��*" Then .Select i, j
                                Next j
                            Next i
                        End With
                    End If
                End If
            End If
        ElseIf Not mrsSNState Is Nothing Then
            mrsSNState.Filter = "���=" & lngSN
            If mrsSNState.RecordCount > 0 Then
                If mrsSNState!״̬ = 1 Or mrsSNState!״̬ = 2 Then
                    Cancel = True
                    MsgBox lngSN & "���ѱ�" & IIf(mrsSNState!״̬ = 1, "ʹ��", "ԤԼ") & "!����������һ����.", vbInformation, gstrSysName
                    Call zlControl.TxtSelAll(txtSN)
                ElseIf mrsSNState!״̬ = 3 Then
                    If mrsSNState!����Ա���� = UserInfo.���� Then
                        If MsgBox(lngSN & "����Ԥ����!��ȷ��Ҫʹ����?", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            Cancel = True
                            txtSN.Text = ""
                            Call zlControl.TxtSelAll(txtSN)
                        Else
                            Call locateSnByʱ��(lngSN)
                        End If
                    Else
                        Cancel = True
                        MsgBox lngSN & "���ѱ�" & mrsSNState!����Ա���� & "Ԥ��!����������һ����.", vbInformation, gstrSysName
                        txtSN.Text = ""
                        Call zlControl.TxtSelAll(txtSN)
                    End If
                ElseIf mrsSNState!״̬ = 4 Then
                    If mTy_Para.blnReuseCancelNO = False Then
                        Cancel = True
                        MsgBox lngSN & "���ѱ��˺�,�޷��ٴ�ʹ��" & "!����������һ����.", vbInformation, gstrSysName
                        Call zlControl.TxtSelAll(txtSN)
                    End If
                ElseIf mrsSNState!״̬ = 5 Then
                    Cancel = True
                    MsgBox lngSN & "���ѱ�����," & "!����������һ����.", vbInformation, gstrSysName
                    Call zlControl.TxtSelAll(txtSN)
                ElseIf mrsSNState!״̬ = 6 Then
                    Cancel = True
                    MsgBox lngSN & "���ѱ�ͣ��," & "!����������һ����.", vbInformation, gstrSysName
                    Call zlControl.TxtSelAll(txtSN)
                End If
                If Cancel = False Then Call locateSnByʱ��(lngSN)
            Else
                If blnHave And vsfList.Cell(flexcpForeColor, lngLocateSnX, lngLocateSnY) = &HC000C0 Then
                    Cancel = True
                    MsgBox lngSN & "�Ų�����!����������һ����.", vbInformation, gstrSysName
                    Call zlControl.TxtSelAll(txtSN)
                Else
                    Call locateSnByʱ��(lngSN)
                End If
            End If
        End If
    End If
End Sub

Private Sub txt��������_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txt��������.Text = "____-__-__" Then
           zlCommFun.PressKey (vbKeyTab) '����ʱ��
           zlCommFun.PressKey (vbKeyTab)
       Else
           zlCommFun.PressKey (vbKeyTab)
       End If
    End If
End Sub

Private Sub txt��������_Validate(Cancel As Boolean)
    If txt��������.Tag <> txt��������.Text Then
        With mobjfrmPatiInfo '������������
            .txt��������.Text = txt��������.Text
            txt��������.Tag = txt��������.Text
            .txt����.Text = txt����.Text
            .txt����.Tag = txt����.Text
            txt����.Tag = txt����.Text
            .cbo���䵥λ.Visible = cbo���䵥λ.Visible
            If .cbo���䵥λ.ListCount <> 0 Then .cbo���䵥λ.ListIndex = cbo���䵥λ.ListIndex
        End With
        Call ShowRegistFromInput
    End If
End Sub

Private Sub txt����ʱ��_Change()
    Dim str����ʱ�� As String
    '76669�����ϴ�,2014-8-18,�����������
    If IsDate(txt��������.Text) And mblnChange Then
        str����ʱ�� = txt��������.Text & IIf(IsDate(txt����ʱ��.Text), " " & txt����ʱ��.Text, "")
        txt����.Text = ReCalcOld(CDate(str����ʱ��), cbo���䵥λ)
        txt����.Tag = txt����.Text
    End If
End Sub

Private Sub txt����ʱ��_GotFocus()
    zlControl.TxtSelAll txt����ʱ��
End Sub

Private Sub txt����ʱ��_KeyPress(KeyAscii As Integer)
    If Not IsDate(txt��������.Text) Then
        KeyAscii = 0
        txt����ʱ��.Text = "__:__"
    End If
    If KeyAscii = 13 Then
        KeyAscii = 0
        zlCommFun.PressKey (vbKeyTab)
    End If
End Sub


Private Sub txt����ʱ��_Validate(Cancel As Boolean)
    If txt����ʱ��.Text <> "__:__" And Not IsDate(txt����ʱ��.Text) Then
        txt����ʱ��.SetFocus
        Cancel = True
    End If
End Sub

Private Sub txt��������_Change()
    Dim str����ʱ�� As String
    
    If IsDate(txt��������.Text) And mblnChange Then
        mblnChange = False
        txt��������.Text = Format(CDate(txt��������.Text), "yyyy-mm-dd") '0002-02-02�Զ�ת��Ϊ2002-02-02,����,��������2002,ʵ��ֵȴ��0002
        mblnChange = True
        
        str����ʱ�� = txt��������.Text & IIf(IsDate(txt����ʱ��.Text), " " & txt����ʱ��.Text, "")
        txt����.Text = ReCalcOld(CDate(str����ʱ��), cbo���䵥λ)
        txt����.Tag = txt����.Text
        mblnGetBirth = False
    End If
End Sub
Private Sub txt��������_GotFocus()
    zlControl.TxtSelAll txt��������
End Sub

Private Sub txt��������_LostFocus()
    If txt��������.Text <> "____-__-__" And Not IsDate(txt��������.Text) Then
      If txt��������.Enabled And txt��������.Visible Then txt��������.SetFocus
    End If
End Sub

Private Sub txtPatient_Change()
    If txtPatient.Locked Then Exit Sub
    If mblnUnChange Or mbytInState = 1 Then Exit Sub
    
    '74430,Ƚ����,2014-7-8,�ҺŽ�����ʾ������Ƭ�ĸ�������
    picPatiPicBack.Visible = False: cmdPatiPic.Enabled = txtPatient.Text <> ""
    
    mblnBoundPati = False
    mblnUnChange = True
    txt�����.Enabled = txtPatient.Text <> "" And InStr(mstrPrivs, ";��������;") > 0
    cmdMore.Enabled = txtPatient.Text <> "" And InStr(mstrPrivs, ";��������;") > 0
    cmdMore.Tag = ""    '�����ж��Ƿ���벡����Ϣ�༭����ȡ�����в���
    cmdCard.Enabled = Not mblnNewCard   'txtPatient.Text <> "" And
    cmdCard.Enabled = cmdCard.Enabled And Not (mblnStation And mTy_Para.bln�Һű���ˢ��)
    
    If Me.ActiveControl Is txtPatient Then
        If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(txtPatient.Text = "")
        If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(txtPatient.Text = "")
        IDKind.SetAutoReadCard (txtPatient.Text = "")
    End If
    
    If Trim(txtPatient.Text) = "" Then
        '�������ʱ��������в�����Ϣ
        If mstr����� = "" Then '���Զ�ˢ���������ʱ�����
            Call ClearPatientInfo
            Call Init�ѱ�(True, False) '�ָ�ȱʡ�ѱ�
            Set mrsInfo = Nothing
            Call ClearmobjfrmPatiInfoFace(Not (mblnNewCard And gblnNewCardNoPop))
        End If
    End If
    mblnUnChange = False
    '��ԭ�ı�����ɫ
    txtPatient.ForeColor = Me.ForeColor
End Sub

Private Sub txtPatient_GotFocus()

    Call zlControl.TxtSelAll(txtPatient)
    
    'LED��������
    If gblnLED And mbytMode <> 1 And mbytInState = 0 And txt�ű�.Text <> "" And txtPatient.Text = "" Then
        zl9LedVoice.Speak "#4" '�����������
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
    '���ܣ�ҽ������鿨
    '���ƣ����˺�
    '���ڣ�2010-07-14 11:32:08
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim lng����ID As Long
    Dim str�������� As String
    Dim rsTmp As ADODB.Recordset
    Dim cur��� As Currency
    Dim curMoney As Currency
    Dim i As Integer
    Dim curPayed As Currency
    Dim curTotal As Currency
    If mrsInfo Is Nothing Then
        lng����ID = 0
        str�������� = ""
    Else
        lng����ID = Val(Nvl(mrsInfo!����ID))
        str�������� = Nvl(mrsInfo!��������)
    End If
    '52867
    Call SetShowBalance
    If gblnLED Then zl9LedVoice.Speak "#50"

    If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(False)
    If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(False)
    IDKind.SetAutoReadCard False
    '68991
    Dim strAdvance As String    '����ģʽ(0-�Ƚ�������ƻ�1-�����ƺ����)|�Һŷ���ȡ��ʽ(0-���ջ�1-����)
    Dim varData As Variant
    mstrYBPati = gclsInsure.Identify(3, lng����ID, mintInsure, strAdvance)
    mRegistFeeMode = EM_RG_����: mPatiChargeMode = EM_�Ƚ��������
    If txtPatient.Text = "" And Not txtPatient.Locked Then
        If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(True)
        If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(True)
        IDKind.SetAutoReadCard True
    End If
    
    If mstrYBPati = "" Then
        '�޸����⣺38917 ���ߣ�Ƚ��
        If Not txtPatient.Enabled Then txtPatient.Enabled = True
         mstrYBPati = "": mintInsure = 0: txtPatient.SetFocus
         Exit Sub
    End If
    
    '�ջ�0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤;7��λ����(����);8����ID
    If UBound(Split(mstrYBPati, ";")) >= 8 Then
        If IsNumeric(Split(mstrYBPati, ";")(8)) Then lng����ID = Val(Split(mstrYBPati, ";")(8))
    End If
        
    If lng����ID = 0 Then
        mstrYBPati = "": mintInsure = 0: txtPatient.SetFocus
        Exit Sub
    End If
        
    '����:29283
    '  -- ����:���ó���-1-�Һ�;2-�շ�
    '  --        ����id_In-����ID(δ������,������)
    '  --        ����_In: ˢ������;δˢ��ʱ,Ϊ��
    '  --         ˢ����ʽ_In:  1-����ˢ��;2-ҽ��ˢ��
    If mbytMode <> 1 Then
        If zlPatiCardCheck(1, lng����ID, CStr(Split(mstrYBPati, ";")(0)), 2) = False Then
            Call ClearmobjfrmPatiInfoFace: ClearPatientInfo
            mstrYBPati = "": mintInsure = 0: txtPatient.SetFocus
            Exit Sub
        End If
    End If
    Call initInsurePara(lng����ID)
    txtPatient.Text = "-" & lng����ID
    Call SetIdentifyLocked(False)
    Call txtPatient_Validate(False)    '���е�Setfocus����ʹ���¼�(txtPatient_KeyPress)ִ�����,�����ٴ��Զ�ִ��txtPatient_Validate
    '74428�����ϴ���2014-7-8������������ʾ��ɫ����
    If mblnUnload Then
        mblnUnload = False
        Exit Sub
    End If
    Call SetPatiColor(txtPatient, str��������, vbRed)
    mobjfrmPatiInfo.txtPatient.ForeColor = txtPatient.ForeColor
    Call SetIdentifyLocked(True)
    '68991
    If strAdvance <> "" Then
        varData = Split(strAdvance & "|", "|")
        mPatiChargeMode = IIf(Val(varData(0)) = 1, EM_�����ƺ����, EM_�Ƚ��������)
        mRegistFeeMode = IIf(Val(varData(1)) = 1, EM_RG_����, EM_RG_����)
    End If
    
    Dim dbl������� As Double
    Set rsTmp = GetMoneyInfo(lng����ID, , , 1, , , True)
    cur��� = 0: stbThis.Panels(4).ToolTipText = ""
    Do While Not rsTmp.EOF
        cur��� = cur��� + Val(Nvl(rsTmp!Ԥ�����))
        cur��� = cur��� - Val(Nvl(rsTmp!�������))
        If Val(Nvl(rsTmp!����)) = 1 Then
            dbl������� = Val(Nvl(rsTmp!Ԥ�����)) - Val(Nvl(rsTmp!�������))
        End If
        rsTmp.MoveNext
    Loop
    
    mcur������� = gclsInsure.SelfBalance(lng����ID, CStr(Split(mstrYBPati, ";")(1)), 10, mcur����͸֧, mintInsure)
    mdbl������� = mcur�������
    stbThis.Panels(3).Text = "�����ʻ����:" & Format(mcur�������, "0.00")
    Call CalcYBMoney
    Call initInsurePara(lng����ID)
    
    Call ShowMedicareInfo(Not mRegistFeeMode = EM_RG_����)
    Call ShowDeposit(False)
    If cur��� > 0 Then
        Call ShowDeposit(Not mRegistFeeMode = EM_RG_����)
        mdblԤ����� = cur���
        curTotal = GetRegistMoney(True)
        curPayed = 0
        For i = 1 To vsfPay.Rows - 1
            curPayed = curPayed + Val(vsfPay.TextMatrix(i, 1))
        Next i
        For i = 1 To vsfPay.Rows - 1
            If vsfPay.RowData(i) = 0 And vsfPay.TextMatrix(i, 0) <> "" Then
                vsfPay.TextMatrix(i, 6) = mdblԤ�����
                If gblnPrePayPriority Then
                    If mdblԤ����� > curTotal - curPayed Then
                        vsfPay.TextMatrix(i, 1) = Format(curTotal - curPayed, "0.00")
                    Else
                        vsfPay.TextMatrix(i, 1) = Format(mdblԤ�����, "0.00")
                    End If
                    Call Set�����Һ�
                End If
            End If
        Next i
        stbThis.Panels(4).Text = "����Ԥ�����:" & Format(cur���, "0.00")
        If Round(dbl�������, 6) <> 0 Then stbThis.Panels(4).ToolTipText = "������Ԥ����" & Format(dbl�������, "0.00")
        
        'ҽ��վ�Һ�ȱʡʹ��Ԥ����
        curMoney = GetRegistMoney
    End If
    
    If MCPAR.ʹ�ø����ʻ� Then
        If mstr�����ʻ� = "" Then MsgBox "�Һų���δ���ø����ʻ����㣬�����ʻ�����֧����", vbInformation, gstrSysName
    End If
    
    '68991
    If mRegistFeeMode = EM_RG_���� Or CheckIsPrice Then
        Call SetUndisplayBalance
    End If
    
End Sub

 
Private Sub txtPatient_KeyDown(KeyCode As Integer, Shift As Integer)
    '�����:44114
    If KeyCode = 38 And 1 < IDKind.IDKind And IDKind.IDKind <= IDKind.ListCount Then 'С�����Ϸ����
        IDKind.IDKind = IDKind.IDKind - 1
    ElseIf KeyCode = 40 And IDKind.IDKind < IDKind.ListCount Then 'С�����·����
        IDKind.IDKind = IDKind.IDKind + 1
    End If
End Sub
Private Sub txtPatient_KeyPress(KeyAscii As Integer)
    Dim lng����ID As Long, blnCard As Boolean
    
    '����:51488
    '�ո����
'    If IDKind.GetCurCard.�Ƿ�ˢ�� = False And KeyAscii = Asc(" ") And mbytInState = 0 Then
'        KeyAscii = 0: Call IDKind_Click(IDKind.GetCurCard): Exit Sub
'    End If
    
    If (KeyAscii = Asc("/") Or KeyAscii = Asc("��") Or KeyAscii = Asc("��") Or KeyAscii = Asc("��")) And Trim(txtPatient.Text) = "" Then
        'ԤԼ����ʱ,������ݺ��������"/"��"��"(ȫ�ǺͰ��),���Զ�����С����,��ԤԼ�Һ���"
        KeyAscii = 0:        Call ShowBookSeled
        Call CreateMobjIDCard
        Exit Sub
    End If
    If SetBrushCard(txtPatient, KeyAscii) = True Then Exit Sub
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If mbytMode <> 1 And Not gblnPrice And Trim(txtPatient.Text) = "" And mobjfrmPatiInfo.mstrCard = "" Then
            'ҽ������鿨
            Call zlInusreIdentify
        Else
            zlCommFun.PressKey (vbKeyTab)
        End If
    ElseIf InStr(1, "'[]+", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0 'סԺ���˲������ٹҺţ������ַ�������Form_KeyPress�н���
    Else
        If txtPatient.Text = "" Then gsngStartTime = Timer
        gblnLen = False
        If IDKind.GetCurCard Is Nothing Then Exit Sub
        If IDKind.GetCurCard.���� = "�����" Then
            If KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then
                If Not (IsNumeric(Chr(KeyAscii)) Or Chr(KeyAscii) = "-") Then KeyAscii = 0: Exit Sub
            End If
        ElseIf IDKind.GetCurCard.���� = "����" Or IDKind.GetCurCard.���� = "��������￨" Then
            blnCard = zlCommFun.InputIsCard(txtPatient, KeyAscii, gCurSendCard.str�������� <> "")
            mblnCard = blnCard
            If blnCard And Len(txtPatient.Text) = gCurSendCard.lng���ų��� - 1 And KeyAscii <> 8 Then
                txtPatient.Text = txtPatient.Text & Chr(KeyAscii): txtPatient.SelStart = Len(txtPatient.Text)
                KeyAscii = 0
                gblnLen = True
                gsngStartTime = Timer
                Call txtPatient_Validate(False)
                mblnCard = False
                '���˺�:27494  20100117
                If Replace(txtPatient.Text, vbCrLf, "") = "" Then
                    DoEvents: txtPatient.SetFocus
                End If
            End If
        ElseIf IDKind.GetCurCard.�ӿ���� <> 0 Then
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
                '���˺�:27494  20100117
                If Replace(txtPatient.Text, vbCrLf, "") = "" Then
                    DoEvents: txtPatient.SetFocus
                End If
            End If
        
        Else
            txtPatient.PasswordChar = IIf(IDKind.ShowPassText, "*", "")
            '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
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
    Dim strTemp As String, lng�����ID As Long
    If txtPatient.Locked And mblnOnVilidate = False Then Exit Sub
    If mstrPrePati = txtPatient.Text Then
        '�Զ�����Ŵ���
        If txt�����.Text = "" Then
            If txt�����.Enabled And txt�����.Visible Then
                txt�����.TabStop = True
                If gbln�Զ������ Or mblnStation Then
                    If txt�ű�.Text <> "" And mbln������ And txt�����.Text = "" And txtPatient.Text <> "" Then
                        txt�����.Text = zlGet�����
                        mintNOLength = Len(txt�����.Text)  '�����ж��޸������ʱ���쳣����
                        txt�����.TabStop = False
                    End If
                End If
            End If
        End If
        If mblnOnVilidate = False Then Exit Sub
    End If
        
    '�ϴιҺŵķ������,�º�ʱ���
    txt�ɿ�.Text = "0.00": txt�Ҳ�.Text = "0.00"
    txt�ϼ�.Text = Format(mcur�ϼ� + GetRegistMoney, "0.00"): mint�Һ��� = 0
    
    Call Set�����Һ�
    If mbytMode = 0 And txt�ɿ�.Enabled = False Then txt�ɿ�.Enabled = True
    
    '�������˻����벡�˺�,����Һ��ۼ�,ԤԼʱ����ɿ�,һֱ�����ۼ�
    If Not (mTy_Para.byt�ɿʽ = 1 And mbytMode <> 1) Then mcur�ϼ� = 0: mcurӦ�� = 0
    
    If txtPatient.Text <> "" Then
        txtPatient.Text = Trim(txtPatient.Text)
        strTemp = txtPatient.Text
        If (Left(txtPatient.Text, 1) = "*" Or Left(txtPatient.Text, 1) = "-") And IsNumeric(Mid(txtPatient.Text, 2)) Then blnTmp = True
        
        Call GetPatient(IDKind.GetCurCard, txtPatient.Text, mblnCard)
        
        '69730,������,2014-01-23,��ҽ������վ�����˹Һű���ˢ�������ļ��
        If mblnStation And mbytMode = 0 And mTy_Para.bln�Һű���ˢ�� Then
            If mrsInfo Is Nothing Then
                MsgBox "û���ҵ��ÿ���Ӧ�Ĳ�����Ϣ������ÿ��Ƿ���ȷ��", vbInformation, gstrSysName
                txtPatient.Text = ""
                txtPatient.SetFocus
                Cancel = True
                Exit Sub
            End If
        End If
        
        '�����:58843
        If mblnStation Then
            If Not mrsInfo Is Nothing Then mstrPrePati = txtPatient.Text
            SetPatiInfoEnabled vsfPlan.TextMatrix(vsfPlan.Row, GetCol("����")) <> "", mrsInfo Is Nothing
        End If
        
        
        '����ԤԼ���ݽ�������
        If Not mblnStation And Not mrsInfo Is Nothing And mbytMode = 0 Then
            If mbytInState = 0 And mstrNoIn <> "" Then Exit Sub
            If zlExistsTodaysAppointment(mrsInfo!����ID) Then Exit Sub
        End If
        
        
        If Not IDKind.GetCurCard.���� Like "����*" Then
            lng�����ID = IDKind.GetCurCard.�ӿ����
            If lng�����ID <> IDKind.GetDefaultCardTypeID And lng�����ID > 0 Then
                mblnCard = False
            End If
            '���˺�:65945,������ȱʡ����Ϊ��������,���������ž�������.
          ' If lng�����ID <= 0 Then lng�����ID = IDKind.GetDefaultCardTypeID

        End If
 
        If mblnCard Or (IsCardType(IDKind, "IC����") _
            Or (gCurSendCard.lng�����ID = lng�����ID And lng�����ID > 0)) And Not blnTmp And lblPrompt.Caption = "" Then
            mblnCard = False
            mbln���� = True '�����:56599
            If mrsInfo Is Nothing Then
                If mblnStation Or mbytMode = 1 Then 'ҽ��վ��ԤԼʱ��֧�ַ���,��Ϊ����Ҫ�շ�
                    Cancel = True: txtPatient.Text = "": Exit Sub
                Else
                    If mTy_Para.bln����סԺ���˹Һ� = False Then
                        If PatiExist(UCase(txtPatient.Text)) Then
                            MsgBox "���ָóֿ�������Ժ,��ò�����ϢĿǰ������!�����Դ˿��Һ�!", vbInformation, gstrSysName
                            Cancel = True: txtPatient.Text = "":  Exit Sub
                        End If
                    End If
                    If IsCardType(IDKind, "IC��") Then mblnICCard = True
                    
                    '������Ѻ͹Һŷ�һ����ȡ���ʱû�е���,����Һŵ�ʱ�ٽ���.���򿨷Ѵ�Ϊ���۵�,��ʱ�ѽ���
                    If LoadCard(False) Then
                        mblnNewCard = True
                        '����:29283
                        '  -- ����:���ó���-1-�Һ�;2-�շ�
                        '  --        ����id_In-����ID(δ������,������)
                        '  --        ����_In: ˢ������;δˢ��ʱ,Ϊ��
                        '  --         ˢ����ʽ_In:  1-����ˢ��;2-ҽ��ˢ��
                        '����:And mbytMode <> 1 :40482
                        If mstrYBPati = "" And mbytMode <> 1 Then
                            If zlPatiCardCheck(1, 0, Trim(mobjfrmPatiInfo.txt����.Text), 1) = False Then
                                Call ClearmobjfrmPatiInfoFace: ClearPatientInfo
                                Cancel = True: txtPatient.Text = "":  Exit Sub
                            End If
                        End If
                        
                        Call ShowRegistFromInput    '���¼��ؿ�����Ϣ
                        txtPatient.PasswordChar = ""
                    Else
                        txtPatient.PasswordChar = ""
                        Cancel = True: txtPatient.Text = "": Exit Sub
                    End If
                End If
            Else
                '����:29283
                '  -- ����:���ó���-1-�Һ�;2-�շ�
                '  --        ����id_In-����ID(δ������,������)
                '  --        ����_In: ˢ������;δˢ��ʱ,Ϊ��
                '  --         ˢ����ʽ_In:  1-����ˢ��;2-ҽ��ˢ��
                'And mbytMode <> 1:40482
                If mstrYBPati = "" And mbytMode <> 1 Then
                    If zlPatiCardCheck(1, Val(Nvl(mrsInfo!����ID)), strTemp, 1) = False Then
                        Call ClearmobjfrmPatiInfoFace: ClearPatientInfo
                        Set mrsInfo = Nothing: txt�����.Enabled = True
                        Cancel = True: txtPatient.Text = "":  Exit Sub
                    End If
               End If
                 '���￨������
                If Mid(gstrCardPass, 1, 1) = "1" And mstrPassWord <> "" Then
                    '54501
                    If Not zlCommFun.VerifyPassWord(Me, "" & mstrPassWord, mrsInfo!����, mrsInfo!�Ա�, "" & mrsInfo!����) Then
                        txt�����.Enabled = True: Set mrsInfo = Nothing
                        Cancel = True: txtPatient.Text = "":  Exit Sub
                    End If
                End If
            End If
        Else
                '����:29283
                '  -- ����:���ó���-1-�Һ�;2-�շ�
                '  --        ����id_In-����ID(δ������,������)
                '  --        ����_In: ˢ������;δˢ��ʱ,Ϊ��
                '  --         ˢ����ʽ_In:  1-����ˢ��;2-ҽ��ˢ��
                'And mbytMode <> 1:40482
                If mstrYBPati = "" And mbytMode <> 1 Then
                    If mrsInfo Is Nothing Then
                        If Trim(mobjfrmPatiInfo.txt����.Text) <> "" Then    '��ȡ�п��ŵĲ���ʱû�м��ؿ��ŵ�����
                            strTemp = Trim(mobjfrmPatiInfo.txt����.Text)
                        End If
                    
                        If zlPatiCardCheck(1, 0, strTemp, 1) = False Then
                            Call ClearmobjfrmPatiInfoFace: ClearPatientInfo
                            Set mrsInfo = Nothing: txt�����.Enabled = True
                            Cancel = True: txtPatient.Text = "":  Exit Sub
                        End If
                    Else
                        If zlPatiCardCheck(1, Val(Nvl(mrsInfo!����ID)), "", 1) = False Then
                            Call ClearmobjfrmPatiInfoFace: ClearPatientInfo
                            Set mrsInfo = Nothing: txt�����.Enabled = True
                            Cancel = True: txtPatient.Text = "":  Exit Sub
                        End If
                    End If
               End If
               mblnCard = False
        End If
        
        If Not mrsInfo Is Nothing And gblnPrice And mbytMode = 0 And txt�ɿ�.Enabled Then txt�ɿ�.Enabled = False
        
        
        If mbytMode <> 2 Then
            If Not mrsInfo Is Nothing And InStr(1, mstrPrivs, ";�������˷ѱ�;") = 0 And Not mblnStation Then
                cbo�ѱ�.Locked = True: cbo�ѱ�.TabStop = False
            Else
                cbo�ѱ�.Locked = False: cbo�ѱ�.TabStop = gbln�ѱ�
            End If
        End If
        '����ͨ��cbo�ѱ�_Click�¼������ShowRegistFromInput
        Call Init�ѱ�((mbytMode = 2 Or mbytMode = 0 And mbytInState = 0 And mstrNoIn <> "") Or mrsInfo Is Nothing, Not mrsInfo Is Nothing Or mblnNewCard)

        If txtPatient.Text = "" And mstr����� <> "" Then 'ʹ��������������Ϊȱʡ��
            Cancel = True
            If IDKind.IDKind = IDKind.GetKindIndex("�����") Then
                IDKind.IDKind = IDKind.GetKindIndex("����")
                mblnReSetIDKind = True
            End If
            txt�����.Text = mstr�����
            Call txtPatient_GotFocus 'LED:��������
            Exit Sub
        End If
        
        '������Ĳ���
        If mrsInfo Is Nothing And (Not mblnNewCard Or gblnNewCardNoPop) And Not mblnBrushPlugin Then
            If mblnIDCardKind And mbytMode = 1 Then
                    '���������,��Ϊ�������ʱ��,�Ѿ��������֤�Ŷ�ȡ������:31182
            Else
                txt����.Text = ""
                Call zlControl.CboLocate(cbo���䵥λ, "��")
                If gstr�Ա� <> "��" Then
                    Call SetCboDefault(cbo�Ա�)
                Else
                    cbo�Ա�.ListIndex = -1
                End If
                txtIDCard.Text = "": txtIDCard.Tag = ""
                txt֤��.Text = "": txt֤��.Tag = ""
            End If
            cbo��ͥ��ַ.Text = ""
            cbo���ڵ�ַ.Text = ""
            txt��ͥ�绰.Text = ""
            '89242:���ϴ�,2015/12/7,��ȡ���˵�ַ��Ϣ
            Call zlLoadDefaultAddr(padd��ͥ��ַ)
            Call zlLoadDefaultAddr(padd���ڵ�ַ)
            '�²��˱�������������
            If Not (txt�����.Text <> "" And mstr����� = txt�����.Text) Then txt�����.Text = ""
            Call SetCboDefault(cbo���ʽ)
            If mbytMode <> 2 Then Call SetCboDefault(cbo�ѱ�)
            Call ClearmobjfrmPatiInfoFace(Not (mblnNewCard And gblnNewCardNoPop))
            Call zlQueryEMPIPatiInfo
        End If
        
        '����ҽ��վ�Һţ��򱾵ز��������Զ����������
        If txt�����.Enabled And txt�����.Visible Then
            txt�����.TabStop = True
            If gbln�Զ������ Or mblnStation Then
                If txt�ű�.Text <> "" And mbln������ And txt�����.Text = "" And txtPatient.Text <> "" Then
                    txt�����.Text = zlGet�����
                    mintNOLength = Len(txt�����.Text)  '�����ж��޸������ʱ���쳣����
                    txt�����.TabStop = False
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
            ElseIf txt�����.Text = "" And txt�����.Enabled And txt�����.Visible Then
                txt�����.SetFocus
            ElseIf cbo���㷽ʽ.Enabled And cbo���㷽ʽ.Visible Then
                cbo���㷽ʽ.SetFocus
            ElseIf chk������.Enabled And chk������.Visible Then
                chk������.SetFocus
            ElseIf txt�ɿ�.Enabled And txt�ɿ�.Visible And mTy_Para.byt�ɿʽ = 1 Then
                txt�ɿ�.SetFocus
            Else
                cmdOK.SetFocus
            End If
        ElseIf Not mrsInfo Is Nothing Then
            '89242:���ϴ�,2015/12/7,��ȡ���˵�ַ��Ϣ
            If mblnStructAdress Then
                If padd��ͥ��ַ.CheckNullValue <> "" And padd��ͥ��ַ.Enabled And padd��ͥ��ַ.Visible And padd��ͥ��ַ.TabStop Then
                    padd��ͥ��ַ.SetFocus
                ElseIf padd���ڵ�ַ.CheckNullValue <> "" And padd���ڵ�ַ.Enabled And padd���ڵ�ַ.Visible And padd���ڵ�ַ.TabStop Then
                    padd���ڵ�ַ.SetFocus
                End If
            Else
                If cbo��ͥ��ַ.Text = "" And cbo��ͥ��ַ.Enabled And cbo��ͥ��ַ.Visible And cbo��ͥ��ַ.TabStop Then
                     cbo��ͥ��ַ.SetFocus
                End If
            End If
            If txt�����.Enabled And txt�����.Visible And IsNull(mrsInfo!�����) And txt�����.TabStop Then
                 txt�����.SetFocus
            ElseIf cbo���㷽ʽ.Enabled And cbo���㷽ʽ.Visible Then
                 cbo���㷽ʽ.SetFocus
            ElseIf chk������.Enabled And chk������.Visible Then
                 chk������.SetFocus
            ElseIf txt�ɿ�.Enabled And txt�ɿ�.Visible And mTy_Para.byt�ɿʽ = 1 Then
                txt�ɿ�.SetFocus
            Else
                 If cmdOK.Enabled And cmdOK.Visible Then cmdOK.SetFocus
            End If
        Else
            If txtPatient.Text = "" And txtPatient.Enabled And txtPatient.Visible Then Cancel = True
        End If
        
    Else 'Ϊ�ձ�ʾ�������벡����Ϣ
        Call ClearPatientInfo
        If mbytMode <> 2 Then Call SetCboDefault(cbo�ѱ�)
        Call ShowRegistFromInput
        
        Call ClearmobjfrmPatiInfoFace(Not (mblnNewCard And gblnNewCardNoPop))
        
        If cbo�ѱ�.Enabled And cbo�ѱ�.Visible Then
             cbo�ѱ�.SetFocus
        ElseIf cbo���㷽ʽ.Enabled And cbo���㷽ʽ.Visible Then
             cbo���㷽ʽ.SetFocus
        ElseIf chk������.Enabled And chk������.Visible Then
             chk������.SetFocus
        Else
             cmdOK.SetFocus
        End If
    End If
    Call ReLoadCardFee(True, True)
    Call Led��ӭ��Ϣ
    
    If CheckIsPrice Or mRegistFeeMode = EM_RG_���� Then
        Call SetUndisplayBalance
    Else
        Call SetShowBalance
    End If
    
    mstrPrePati = txtPatient.Text
End Sub

Private Sub Led��ӭ��Ϣ()
    Dim strInfo As String, lngPatient As Long
    'LED��ʼ��
    If mbytMode = 0 And mbytInState = 0 And gblnLED Then
        zl9LedVoice.Speak "#1"
        
        strInfo = Trim(txtPatient.Text)
        If Not mrsInfo Is Nothing Then
            If mrsInfo.State = 1 Then strInfo = strInfo & " " & mrsInfo!�Ա� & " " & mrsInfo!����: lngPatient = Val("" & mrsInfo!����ID)
        End If
        zl9LedVoice.DisplayPatient strInfo, lngPatient
    End If
End Sub

Private Sub txt�ű�_Validate(Cancel As Boolean)
    '�����һ�ŵ��ݺ�
    If mbytInState = 0 And chkCancel.Value = 0 Then
        If cboNO.ListIndex <> -1 Then cboNO.ListIndex = -1
    End If
    mstrPre�ű� = Trim(txt�ű�.Text) '53299
    If Trim(txt�ű�.Text) = "" Then Exit Sub
    mlngPreRow = vsfPlan.Row
    If CheckNoValied(vsfPlan.Row) = False Then
        mstrPre�ű� = "" '53299
        mlngPreRow = 0
        Cancel = True
         txt�ű�.Text = "": txt�ű�.SetFocus: Exit Sub
    End If
End Sub

 
Private Sub txt��ͥ�绰_GotFocus()
    zlCommFun.OpenIme False
End Sub

Private Sub txt��ͥ�绰_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txt��ͥ�绰_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txt��ͥ�绰, KeyAscii, m�ı�ʽ
End Sub

Private Sub txt��ͥ�绰_Validate(Cancel As Boolean)
    If mobjfrmPatiInfo Is Nothing Then Exit Sub
    With mobjfrmPatiInfo
        .txt��ͥ�绰.Text = txt��ͥ�绰.Text
    End With
End Sub

Private Sub txt�ɿ�_Change()
    Dim curӦ�� As Currency
    If Val(txt�ɿ�.Text) = 0 Then
        txt�Ҳ�.Text = "0.00"
    Else
        curӦ�� = mcurӦ�� + GetRegistMoney
        txt�Ҳ�.Text = Format(Val(txt�ɿ�.Text) - Val(txt����Ӧ��.Text), "0.00")
    End If
End Sub

Private Sub txt�ɿ�_GotFocus()
    Dim curӦ�� As Currency
    
    'ֻ�Խɿ���Ϊ�շѽ�������ʱ,��������ɿ��0
    If mTy_Para.byt�ɿʽ = 1 Then
        If Val(txt�ɿ�.Text) = 0 And Me.ActiveControl Is txt�ɿ� Then
            txt�ɿ�.Text = ""
        End If
    End If
    Call zlControl.TxtSelAll(txt�ɿ�)
    
    'LED��������
     If Not (mintInsure <> 0 And mstrYBPati <> "") Then
        curӦ�� = mcurӦ�� + GetRegistMoney
        If gblnLED And mbytMode <> 1 And mbytInState = 0 Then
            zl9LedVoice.Speak "#21 " & Format(curӦ��, "0.00")
        End If
    End If
End Sub

Private Sub txt�ɿ�_KeyPress(KeyAscii As Integer)
    Dim curӦ�� As Currency
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txt�ɿ�.Text = "" Then
            If GetRegistMoney = 0 Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
        End If
        If mTy_Para.byt�ɿʽ = 1 And txt�ɿ�.Text = "" Then Exit Sub
'        If Val(txt�ɿ�.Text) <> 0 Then
'            If Val(txt�Ҳ�.Text) < 0 Then
'                MsgBox "�ɿ���㡣", vbInformation, gstrSysName
'                Call zlControl.TxtSelAll(txt�ɿ�): Exit Sub
'            End If
'        End If
        Call zlCommFun.PressKey(vbKeyTab)
        
        'LED��ʾ
         If Not (mintInsure <> 0 And mstrYBPati <> "") Then
            If gblnLED And mbytMode <> 1 And mbytInState = 0 And Val(txt�Ҳ�.Text) >= 0 Then
                curӦ�� = mcurӦ�� + GetRegistMoney
                zl9LedVoice.DispCharge Format(curӦ��, "0.00"), txt�ɿ�.Text, txt�Ҳ�.Text
                zl9LedVoice.Speak "#22 " & txt�ɿ�.Text
                zl9LedVoice.Speak "#23 " & txt�Ҳ�.Text
                zl9LedVoice.Speak "#3"
                txt�ɿ�.Tag = "1"
            End If
        End If
    Else
        If KeyAscii = Asc(".") And InStr(txt�ɿ�.Text, ".") > 0 Then KeyAscii = 0:  Exit Sub
        If InStr("0123456789." & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txt�ɿ�_LostFocus()
    txt�ɿ�.Text = Format(Val(txt�ɿ�.Text), "0.00")
End Sub

Private Sub txt�ɿ�_Validate(Cancel As Boolean)
    txt�ɿ�.Text = Format(Val(txt�ɿ�.Text), "0.00")
End Sub

Private Sub txt����_Change()
    If Not mrsInfo Is Nothing Then
        If mlng�Һſ���ID > 0 And txt����.Text <> "" Then
            mobjfrmPatiInfo.chk����.Value = IIf(Check����(mrsInfo!����ID, mlng�Һſ���ID), 1, 0)
        End If
    End If
End Sub

Private Sub txt�����_GotFocus()
    If InStr(";" & mstrPrivs & ";", ";�����޸������;") > 0 Then
        '�����޸�������ǲ�ȫ��ѡ��
        Call zlControl.TxtSelAll(txt�����)
    End If
End Sub

Private Sub txt�����_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        
        If txt�����.Enabled And txt�����.Visible And mintNOLength > 0 And mblnCheckNOValidity Then
        '����ֹ��������쳣�����������ʾ
            If Len(txt�����.Text) > mintNOLength + 1 Then
                MsgBox "ע��,���������Ź���,��ȷ���Ƿ���������!", vbInformation, gstrSysName
                txt�����.SetFocus
                txt�����.SelStart = 0: txt�����.SelLength = Len(txt�����.Text)
                Exit Sub
            End If
        End If
        
        Call zlCommFun.PressKey(vbKeyTab)
    ElseIf KeyAscii = 32 Then
        KeyAscii = 0
        If txt�����.Text = "" Then
            txt�����.Text = zlGet�����
            mintNOLength = Len(txt�����.Text)      '�����ж��޸������ʱ���쳣����
        End If
        If ActiveControl Is txt����� Then
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    ElseIf InStr("1234567890" & Chr(8), Chr(KeyAscii)) = 0 Or InStr(";" & mstrPrivs & ";", ";�����޸������;") = 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txt�����_Validate(Cancel As Boolean)
    '��������������,�򲻿����
    If txt�����.Text = "" Then
        If Not mrsInfo Is Nothing Then
            txt�����.Text = Nvl(mrsInfo!�����)
        End If
    End If
End Sub

Private Sub txt����_LostFocus()
    zlCommFun.OpenIme
End Sub

Private Sub txt����_Validate(Cancel As Boolean)
    Dim strBirth As String
    If txt����.Locked Then Exit Sub
    txt����.Text = Trim(txt����.Text)
    If Not IsNumeric(txt����.Text) And Trim(txt����.Text) <> "" Then
        cbo���䵥λ.ListIndex = -1: cbo���䵥λ.Visible = False: txt����.Width = 1320
    ElseIf cbo���䵥λ.Visible = False Then
        cbo���䵥λ.ListIndex = 0: cbo���䵥λ.Visible = True: txt����.Width = 600
    End If
    '69026,Ƚ����,2014-8-8,�����������
    If txt����.Visible And Trim(txt����.Text <> "") Then
        If mobjfrmPatiInfo.mobjPubPatient Is Nothing Then Exit Sub
        If mobjfrmPatiInfo.mobjPubPatient.CheckPatiAge(Trim(txt����.Text) & IIf(cbo���䵥λ.Visible, cbo���䵥λ.Text, "")) = False Then
            Cancel = True: txt����.SetFocus: Exit Sub
        End If
    End If
    
    If txt����.Tag <> txt����.Text Then
        With mobjfrmPatiInfo '������������
            .txt����.Text = txt����.Text
            .txt����.Tag = txt����.Text
            If .cbo���䵥λ.ListCount = 0 Then CopyCboTofrmPatiInfo
            .cbo���䵥λ.ListIndex = cbo���䵥λ.ListIndex
            .cbo���䵥λ.Visible = cbo���䵥λ.Visible
            If Not IsDate(txt��������.Text) Then mblnGetBirth = True
            .mblnChange = False
            '125451�������Ƿ�����ͨ����������������
            If mblnGetBirth Then
    '                .txt��������.Text = ReCalcBirth(.txt����.Text, .cbo���䵥λ.Text)
                If mobjfrmPatiInfo.mobjPubPatient.ReCalcBirthDay(Trim(txt����.Text) & IIf(cbo���䵥λ.Visible, cbo���䵥λ.Text, ""), strBirth) Then
                    .txt��������.Text = Format(strBirth, "yyyy-mm-dd")
                    .txt����ʱ��.Text = Format(strBirth, "hh:mm")
                End If
            End If
            .mblnChange = True
        End With
        
        txt����.Tag = txt����.Text
        '89130:���ϴ�,2015/10/13,���³�������
        mblnChange = False
        txt��������.Text = mobjfrmPatiInfo.txt��������.Text
        txt����ʱ��.Text = mobjfrmPatiInfo.txt����ʱ��.Text
        mblnChange = True
        Call ShowRegistFromInput
        Call ReLoadCardFee(, True)
    End If
End Sub
Private Function CheckNoValied(ByVal lngRow As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ����ָ���еĺű��Ƿ���Ч
    '���أ���Ч,����true,���򷵻�False
    '���ƣ����˺�
    '���ڣ�2010-08-17 16:00:11
    '˵����31922
    '------------------------------------------------------------------------------------------------------------------------
    If InStr(1, mstrPrivs, ";��ʱ�Һ�;") > 0 Or mblnStation Or mbytMode <> 0 Then
        CheckNoValied = True: Exit Function
    End If
    With vsfPlan
        If Val(.Cell(flexcpData, lngRow, .ColIndex("�ű�"))) = 1 Then
            '31922
            '���ܹҴ˺�
            MsgBox "�ű�" & .TextMatrix(lngRow, .ColIndex("�ű�")) & "��������Ч��Χ�ڻ���Ȩ�޲���,���ܹҺ�,����!", vbInformation + vbOKOnly + vbDefaultButton1
            Exit Function
        End If
    End With
    CheckNoValied = True
End Function

Private Sub txt�ű�_Change()
'���ܣ���������ű���ʾ����
    Dim strInfo As String, i As Integer
    Dim blnChkLimit As Boolean
    
    '�����һ�ŵ��ݺ�
    mlng�Һſ���ID = 0
    txt����.Text = ""
    txtSN.Text = ""
    mlngPreRow = 0
        
    If mbytInState = 1 Then Exit Sub
    If chkCancel.Value = 1 Or chkPrint.Value = 1 Then Exit Sub
    If mblnUnChange Then Exit Sub
    
    'ˢ�ºű�ֱ�Ӵӻ����ж�ȡ����
    If vsfPlan.Tag = "" Then
        mblnManualInput = True
        Call ShowPlans(, Len(txt�ű�) > 0 And IsNumeric(Trim(txt�ű�.Text)), False)
        mblnManualInput = False
    End If
    
    If Trim(txt�ű�.Text) = "" Then
        chk������.Enabled = mbln������
        lblFree.Visible = False
        Exit Sub
    End If
    
    '�ϴιҺŵĽɿ����,�º�ʱ���
    txt�ɿ�.Text = "0.00": txt�Ҳ�.Text = "0.00"
    
    If txt�ű�.Text = "+" Then '��������
        txtSN.Text = ""
        txtSN.Enabled = False
        
        mlng�Һſ���ID = UserInfo.����ID
        If Not mrsInfo Is Nothing Then
            Call Init�ѱ�(mobjfrmPatiInfo.chk����.Value = 0, True)
        Else
            Call Init�ѱ�(True, mblnNewCard)
        End If
        Call ShowRegistFromInput
        Call zlCommFun.PressKey(vbKeyTab)
    ElseIf (IsNumeric(Trim(txt�ű�.Text)) And Len(Trim(txt�ű�.Text)) = gint�ų� Or vsfPlan.Rows = 2) Or vsfPlan.Tag <> "" Then
        If vsfPlan.Tag = "" Then
            If vsfPlan.Rows = 2 And Trim(txt�ű�.Text) <> vsfPlan.TextMatrix(vsfPlan.Row, GetCol("�ű�")) Then
                '��ǰ�ű��б�ֻ��һ��ʱ�����û���������ű𣬲��Զ�ƥ�䣬���ǰ��س�
                Exit Sub
            End If
            '��λ����еĺű�
            For i = 1 To vsfPlan.Rows - 1
                If Trim(vsfPlan.TextMatrix(i, GetCol("�ű�"))) = Trim(txt�ű�.Text) Then
                    If CheckNoValied(i) = False Then
                         txt�ű�.Text = "": txt�ű�.SetFocus: Exit Sub
                    End If
                    Call vsfPlan_LeaveCell
                    vsfPlan.Row = i: vsfPlan.RowSel = i
                    vsfPlan.Col = vsfPlan.ColIndex("IDS"): vsfPlan.ColSel = vsfPlan.Cols - 1
                    Call vsfPlan_EnterCell
                    SetGridTop i
                    Exit For
                End If
            Next
            '�ű����ް���ʱҪ������
            If i = vsfPlan.Rows Then
                txt�ű�.Text = "": txt�ű�.SetFocus: Exit Sub
            End If
        End If
        
        '����Ȩ�޿���
        If vsfPlan.TextMatrix(vsfPlan.Row, GetCol("����")) <> "" Then
            If InStr(mstrPrivs, ";��������;") = 0 Then
                MsgBox "�úű�Ҫ������˽������ﲡ��������û�н���������Ȩ�ޡ����ܼ����Һţ�", vbInformation, gstrSysName
                txt�ű� = "": txt�ű�.SetFocus: Exit Sub
            End If
            Call SetPatiInfoEnabled(True, mrsInfo Is Nothing) '�����:58843
            If mrs��ͥ��ַ Is Nothing And Not mblnStructAdress Then Call Load��ͥ��ַ
        Else
            Call SetPatiInfoEnabled(False, mrsInfo Is Nothing) '�����:58843
        End If
        
        If mbytMode = 1 Then
            blnChkLimit = vsfPlan.TextMatrix(vsfPlan.Row, GetCol("��Լ")) <> ""
            If blnChkLimit = False Then
                blnChkLimit = vsfPlan.TextMatrix(vsfPlan.Row, GetCol("�޺�")) <> ""
            End If
        Else
            blnChkLimit = vsfPlan.TextMatrix(vsfPlan.Row, GetCol("�޺�")) <> ""
        End If
        '�޺ſ���
        If chkCancel.Value = 0 And blnChkLimit And Not mblnFinishReg Then
            '����:26962 ����:2009-12-25 11:46:30
            If zlCheck��Լ���޺���(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("��¼ID"))) = False Then Exit Sub
        End If
        
        'ȷ����ǰ���
        txtSN.Enabled = vsfPlan.TextMatrix(vsfPlan.Row, GetCol("��ſ���")) <> ""
        If txtSN.Enabled And vsfList.Tag = "" And vsfList.Visible Then
            txtSN.Text = GetCurrSN(, Not mTy_Para.bln������ѡ��)
            If Val(txtSN.Text) = 0 Then
                txtSN.Text = ""
                If CheckArangement = False Then Exit Sub
            Else
                Call LocateSN(Val(txtSN.Text))
            End If
        End If
        Dim blnCancel As Boolean
        
        'װ��Һ�����
        '�ѱ��¼��е���ShowRegistFromInput
        mstrPre�ѱ� = ""
        
        '72168
        mlng�Һſ���ID = Abs(vsfPlan.RowData(vsfPlan.Row))
        If Not mrsInfo Is Nothing Then
            Call Init�ѱ�(mobjfrmPatiInfo.chk����.Value = 0, True)
        Else
            Call Init�ѱ�(True, mblnNewCard)
        End If
        
        If CheckIsPrice Or mRegistFeeMode = EM_RG_���� Then
            Call SetUndisplayBalance
        Else
            Call SetShowBalance
        End If
        Call zlCommFun.PressKey(vbKeyTab)
    End If
    
End Sub

Private Function GetCurrSN(Optional ByVal lngCurMaxSN As Long = -1, Optional ByVal blnGetLapseNO As Boolean = False) As Long
'����:��ȡ��ǰ�ű�����������
'     ȫ��������ʱ����0
'    blngetlapseNo:�Ƿ����Ч���Ժ�ʼ��
'     lngCurMaxSN-�������ʹ�ú�
    Dim i           As Integer
    Dim j           As Integer
    Dim lngMaxSn    As Long
    Dim lngSN       As Long
    Dim intStart    As Integer
    Dim lngTmp      As Long
    Dim blnUnitReg  As Boolean
    Dim lngMaxLapse As Long '�����Ч����
    If mbytMode = 1 And mblnUnitReg And Not mrsUnitReg Is Nothing Then
        blnUnitReg = True
    End If
    
'    If (mViewMode = v_ר�Һ� Or mViewMode = v_ר�Һŷ�ʱ��) And Not mTy_Para.bln������ѡ�� And blnGetLapseNO Then
'        lngMaxLapse = GetMaxLapseNO
'    End If
    
    mtyRegPlanState.lngSelNO = 0
    mtyRegPlanState.lngSelX = 0
    mtyRegPlanState.lngSelY = 0
    mtyRegPlanState.strSelTime = ""
   
   If Not mrsSNState Is Nothing Or blnUnitReg Then
ReGet:
        If mrsSNState Is Nothing And mbytMode = 1 Then Set mrsSNState = GetSNState(Val(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("��¼ID"))))
        mrsSNState.Filter = ""
        If mrsSNState.RecordCount > 0 Or blnUnitReg Then
        
            If lngCurMaxSN = -1 And mViewMode = v_ר�Һŷ�ʱ�� Then
                With vsfList
                    i = vsfList.Row
                    j = vsfList.Col
                    If .TextMatrix(i, j) <> "" Then
                        If .Cell(flexcpForeColor, i, j) <> vbRed And .Cell(flexcpForeColor, i, j) <> vbBlue And .Cell(flexcpForeColor, i, j) <> vbGreen And .Cell(flexcpForeColor, i, j) <> vbGrayText And .Cell(flexcpForeColor, i, j) <> &HC000C0 Then
                           lngTmp = Val(Getʱ��(i, j, False))
                           mrsSNState.Filter = "���=" & lngTmp & " And ״̬ <> 0"
                            If mrsSNState.RecordCount = 0 And lngTmp > lngMaxLapse Then
                                    GetCurrSN = lngTmp
                                    mtyRegPlanState.lngSelNO = lngTmp
                                    mtyRegPlanState.lngSelX = i
                                    mtyRegPlanState.lngSelY = j
                                    mtyRegPlanState.strSelTime = Getʱ��(i, j, True)
                                    Exit Function
                            End If
                        End If
                    End If
                End With
            End If
            
            
           If lngCurMaxSN = -1 And mViewMode = v_ר�Һ� Then
               lngTmp = 0
               mrsSNState.Filter = "ԤԼ=0 and ״̬=1"
                Do While Not mrsSNState.EOF
                   If lngTmp < Val(mrsSNState!���) Then lngTmp = Val(mrsSNState!���)
                   mrsSNState.MoveNext
                Loop
                mrsSNState.Filter = 0
               If lngTmp <> 0 Then lngCurMaxSN = lngTmp
            End If
            
            
            intStart = IIf(mViewMode = v_ר�Һŷ�ʱ�� Or mViewMode = V_��ͨ�ŷ�ʱ��, 1, 0)
            For i = 0 To vsfList.Rows - 1
                For j = intStart To vsfList.Cols - 1
                    Select Case mViewMode
                    Case V_��ͨ��, v_ר�Һ�:
                        lngSN = Val(vsfList.TextMatrix(i, j))
                        If vsfList.Cell(flexcpForeColor, i, j) = &HC000C0 And mbytMode = 1 Then
                            lngSN = -1
                        End If
                        
                    Case v_ר�Һŷ�ʱ��:
                        With vsfList
                            If .Cell(flexcpForeColor, i, j) = vbGrayText Or .Cell(flexcpForeColor, i, j) = &HC000C0 Then
                                lngSN = -1
                            Else
                               lngSN = IIf(Trim(.TextMatrix(i, j)) = "", -1, Val(Getʱ��(i, j, False)))
                               If lngSN < lngMaxLapse And mTy_Para.bln������ѡ�� = False Then lngSN = -1
                               
                               '�������Ѿ������һ�������,��Ҫ����Ƿ���ڼӺ�,�Լ��Ƿ�������ѡ��,������ѡ��,ʱ ����ѡ���Ѿ��˺ŵ���� 'lgf
                               If lngSN = mtyRegPlanState.lngLastNO And lngSN > 0 And mtyRegPlanState.blnAdditionalNumber And Not mTy_Para.bln������ѡ�� Then lngSN = -1
                            End If
                        End With
                    Case Else
                       Exit Function
                    End Select
                    '73411:Ĭ����ŵ�����
                    If lngSN > -1 Then
                        mrsSNState.Filter = "���=" & lngSN & " And ״̬ <> 0"
                        '�����:52335
                        If mrsSNState.RecordCount = 0 Then
                            lngMaxSn = lngSN
                            mblnStateChange = True
                            vsfList.Select i, j
                            mblnStateChange = False
                            mtyRegPlanState.lngSelNO = lngSN
                            mtyRegPlanState.lngSelX = i
                            mtyRegPlanState.lngSelY = j
                            If mViewMode = V_��ͨ�ŷ�ʱ�� Or mViewMode = v_ר�Һŷ�ʱ�� Then
                                'ֻ�з�ʱ��,�Ŵ���ʱ��
                                mtyRegPlanState.strSelTime = Getʱ��(i, j, True)
                            End If
                            Exit For
                        End If
                    End If
                    
                Next
                
                If lngMaxSn = lngSN Then Exit For
            Next
            If lngCurMaxSN > 0 And lngMaxSn = 0 Then
                '���˺�:???
                '��Ҫ�ǽ��ԤԼ���+1��,����ԤԼ�����,�����ִ�1��ʼ����Ƿ���δѡ���.
                '��:ԤԼ��5��ʼ;����7�Ѿ���������,����ٴ�1��ʼȡ.
               ' lngCurMaxSN = -1: GoTo ReGet:
            End If
            GetCurrSN = lngMaxSn
        Else
            Select Case mViewMode
                Case v_ר�Һŷ�ʱ��:
                     vsfList.Redraw = False
                    For i = 0 To vsfList.Rows - 1
                        For j = 1 To vsfList.Cols - 1
                            If (vsfList.Cell(flexcpForeColor, i, j) = vbBlue Or vsfList.Cell(flexcpForeColor, i, j) = vbBlack) And vsfList.TextMatrix(i, j) <> "" Then
                                GetCurrSN = Val(Getʱ��(i, j, False))
                                mtyRegPlanState.lngSelNO = GetCurrSN
                                mtyRegPlanState.lngSelX = i
                                mtyRegPlanState.lngSelY = j
                                mtyRegPlanState.strSelTime = Getʱ��(i, j, True)
                                vsfList.Redraw = True
                                Exit Function
                            End If
                        Next
                    Next
                    vsfList.Redraw = True
                Case Else:
                  If mbytMode = 1 And mblnUnitReg And Not mrsUnitReg Is Nothing Then
                      mrsUnitReg.Filter = "���=1"
                      If mrsUnitReg.RecordCount = 0 Then GetCurrSN = 1
                      mrsUnitReg.Filter = 0
                  Else
                    GetCurrSN = 1
                  End If
            End Select
        End If
    End If

End Function


Private Sub txt�ű�_GotFocus()
    Call zlControl.TxtSelAll(txt�ű�)
    
    If gblnLED And mbytMode <> 1 And mbytInState = 0 And txt�ű�.Text = "" And mblnLEDKey Then
        zl9LedVoice.Speak "#14" '�������ʲô��
    End If
    mblnLEDKey = False
End Sub

Private Sub txt�ű�_KeyDown(KeyCode As Integer, Shift As Integer)
'�����ƶ��ű�,�Ա����ѡ��
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

Private Sub txt�ű�_KeyPress(KeyAscii As Integer)
    '�ϴιҺŵĽɿ����,�º�ʱ���
    txt�ɿ�.Text = "0.00": txt�Ҳ�.Text = "0.00"
    txt�ϼ�.Text = Format(mcur�ϼ� + GetRegistMoney, "0.00")
    Call Set�����Һ�
    
    If KeyAscii = Asc("/") And Trim(txt�ű�.Text) = "" Then
        'ԤԼ����ʱ,������ݺ��������"/",���Զ�����С����,��ԤԼ�Һ���"
        KeyAscii = 0:        Call ShowBookSeled
        Exit Sub
    End If
    
    If KeyAscii = Asc("+") Then
        If mbytInState = 0 And (Not mbln������ Or picBookingDate.Visible Or mblnStation) Then
            KeyAscii = 0: Exit Sub 'ԤԼʱ������������
        End If
        '����:27493
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
        '����ڰ����˼�
        KeyAscii = 0: zlCommFun.PressKey vbKeyBack
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        If CheckNoValied(vsfPlan.Row) = False Then
             txt�ű�.Text = "": txt�ű�.SetFocus: Exit Sub
        End If
        
        vsfPlan.Tag = vsfPlan.TextMatrix(vsfPlan.Row, GetCol("�ű�"))
        If vsfPlan.Tag <> "" Then
            If txt�ű�.Text <> vsfPlan.Tag Then
                txt�ű�.Text = vsfPlan.Tag  '�Զ�����change�¼�
            Else
                Call txt�ű�_Change
            End If
            vsfPlan.Tag = ""
        End If
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If InStr("1234567890+ABCDEFGHIJKLMNOPQRSTUVWXYZ" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    '�����:110228,����,2017/07/20,�Һ�ʱ���˺ű�ˢ�²�����
    If txt�ű�.SelLength > 0 Then
        Set mrsPlan = Nothing
    End If
End Sub

Private Sub txt����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub txt����_GotFocus()
    'txt����.IMEMode = vbIMEOff
    Call zlCommFun.OpenIme(True)
    Call zlControl.TxtSelAll(txt����)
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    Dim blnTab As Boolean
    
    If txt����.Locked Then Exit Sub
    If KeyAscii = vbKeyReturn Then
        If txtPatient.Text <> "" And txt����.Text = "" And gbln���� Then Exit Sub
        
        If cbo���䵥λ.Visible = False And IsNumeric(txt����.Text) Then
            Call txt����_Validate(False)
            If cbo���䵥λ.Visible And cbo���䵥λ.Enabled Then cbo���䵥λ.SetFocus
        Else
            Call zlCommFun.PressKey(vbKeyTab)
        End If
        If Not IsNumeric(txt����.Text) And cbo���䵥λ.Visible Then Call zlCommFun.PressKey(vbKeyTab)
    Else
        '�������Ƽ��� ָ����������ַ�
        If InStr("~����@#��%����&*��������-+=|����������~`!#$%^&*()-_=+|\/?<>,/<>", UCase(Chr(KeyAscii))) > 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txt����_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        glngTXTProc = GetWindowLong(txt����.Hwnd, GWL_WNDPROC)
        Call SetWindowLong(txt����.Hwnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txt����_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Call SetWindowLong(txt����.Hwnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Function bln����(ByVal strCardNo As String) As Boolean
'---------------------------------------------------------------------------------------------------------------------------------------------
'����:�жϵ�ǰ�Ƿ�Ϊ�������� (���Ƿ����������ǰ󶨿�����)
'���:
'����:56599
'����:2012-12-12 14:55:36
'---------------------------------------------------------------------------------------------------------------------------------------------
    Dim bln�Ƿ񷢿� As Boolean
    '115168:���ϴ���2017/12/13�����淢����ҽ�ƿ�����
    If mCurSendCard.lng�����ID = 0 Then mCurSendCard = gCurSendCard
    '89572:���ϴ�,2015/10/20,�Һŷ�����ȡƱ������ID
    If mCurSendCard.bln�ϸ���� = True Then
        mlng�ſ�����ID = CheckUsedBill(5, IIf(mlng�ſ�����ID > 0, mlng�ſ�����ID, mCurSendCard.lng��������), strCardNo, mCurSendCard.lng�����ID)
        bln�Ƿ񷢿� = IIf(mlng�ſ�����ID <= 0, False, True)
        If mCurSendCard.bln���ƿ� = False Then
            bln�Ƿ񷢿� = (mCurSendCard.bln�Ƿ񷢿� = True)
        End If
    Else
        bln�Ƿ񷢿� = mbln����
        If mblnAlwaysSend Then bln�Ƿ񷢿� = True
        If mCurSendCard.bln���ƿ� = False Then
            bln�Ƿ񷢿� = (mCurSendCard.bln�Ƿ񷢿� = True)
        End If
    End If
    bln���� = bln�Ƿ񷢿�
    mbln���� = bln�Ƿ񷢿�
End Function

Private Sub ClearmobjfrmPatiInfoFace(Optional blnClearCard As Boolean = True)
    Dim i As Integer
            
    With mobjfrmPatiInfo
        Call CopyCboTofrmPatiInfo '�������û��Load,��ʱ��Load����Form_load�¼�
                
        .chk����.Value = 0
        .txt�����.Text = "": .txt�����.MaxLength = txt�����.MaxLength
        SetCboDefault .cbo�ѱ�
        SetCboDefault .cbo�Ա�
            
        .txtPatiMCNO(0).Text = ""
        .txtPatiMCNO(0).Tag = ""
        .txtPatiMCNO(1).Text = ""
        
        If blnClearCard Then
            .mstrCard = ""
            .txt����.Text = ""
            If mblnNoClearPrompt = False Then lblPrompt.Caption = "": gCurSendCard.lng�շ�ϸĿID = 0: vsfPay.Height = 2220
            mblnNewCard = False
            mblnAddCardItem = False
        End If
        .txt����.Text = ""
        .txt��֤.Text = ""
        If mbytMode = 1 And mblnIDCardKind Then
            '31182:��Ϊ�ڶ�ȡ���֤ʱ,�Ѿ���ֵ���������
        Else
            .txt����.Text = "": .txt����.MaxLength = txt����.MaxLength
            .txt����.Tag = ""
            .txt��������.Text = "____-__-__"
            .txt����ʱ��.Text = "__:__"
            Call zlControl.CboLocate(.cbo���䵥λ, "��")
            .cbo���䵥λ.Tag = .cbo���䵥λ.Text
            .txt���֤��.Text = ""
            .txt���֤��.Tag = ""
        End If
        .txtPatient.Text = "": .txtPatient.MaxLength = txtPatient.MaxLength
        
        SetCboDefault .cbo���ʽ
        SetCboDefault .cbo����
        SetCboDefault .cbo����
        SetCboDefault .cbo����
        SetCboDefault .cboְҵ
        
        
        .txt��λ����.Text = ""
        .txt��λ����.Tag = ""
        .txt��λ�绰.Text = ""
        .txt��λ�ʱ�.Text = ""
        .txt����.Text = ""
        .cbo��ͥ��ַ.Text = ""
        .txt��ͥ�ʱ�.Text = ""
        .txt��ͥ�绰.Text = ""
        .txt������Ӧ.Text = ""
        '�����:40005
        .txt��ϵ�˵绰.Text = ""
        .cbo��ϵ�˹�ϵ.ListIndex = -1
        .txtMobile = ""
        .txt��ϵ�����֤.Text = ""
        .txt��ϵ������.Text = ""
        .txtBirthLocation.Text = ""
        .txtRegLocation.Text = ""
        .txt���ڵ�ַ�ʱ�.Text = ""
        '89242:���ϴ�,2015/12/7,��ղ��˵�ַ��Ϣ
        .padd��ͥ��ַ.Value = ""
        .padd���ڵ�ַ.Value = ""
        '82649:���ϴ�,2015/2/13,����໤����Ϣ
        .txt�໤��.Text = ""
        For i = 1 To .msh����.Rows - 1
            .msh����.TextMatrix(i, 0) = ""
            .msh����.TextMatrix(i, 1) = "" '�����:56599
            .msh����.RowData(i) = 0
        Next
        '�����:56599
        .msh����.Rows = 2
        .Clear��������
        If .mblnNewPatient = False Then
            '74430,Ƚ����,2014-7-7,�Һ��еĲ�����Ϣ�༭�������ṩ�ɼ���Ƭ����
            .imgPatient.Picture = Nothing
        End If
    End With
End Sub

Private Function LoadzlIDKindPatiInfor(objPati As zlIDKind.PatiInfor) As Boolean
    'IDKind_Read�¼���,�²��˼�����Ϣ����������
    ClearmobjfrmPatiInfoFace True
Call SetCboDefault(cboҽ�����)
      Call zlControl.CboLocate(cbo�Ա�, objPati.�Ա�)
      
         
    With mobjfrmPatiInfo
        .txtPatient.Text = txtPatient.Text: .txtPatient.MaxLength = txtPatient.MaxLength
        
             
          If 1 = 1 Then
        Else
            .txt����.Tag = 0
        End If
        If Not mrsInfo Is Nothing Then
            .mlng����ID = Val(Nvl(mrsInfo!����ID))
        Else
            .mlng����ID = 0
        End If
        
        
        .cbo�Ա�.ListIndex = cbo�Ա�.ListIndex
        .cbo���䵥λ.ListIndex = cbo���䵥λ.ListIndex
        .txt����.Text = txt����.Text: .txt����.MaxLength = txt����.MaxLength
        .txt����.Tag = txt����.Text
        .cbo��ͥ��ַ.Text = cbo��ͥ��ַ.Text
        .txtRegLocation = cbo���ڵ�ַ.Text
        '89242:���ϴ�,2015/12/7,��ȡ���˵�ַ��Ϣ
        Call .padd��ͥ��ַ.LoadStructAdress(padd��ͥ��ַ.valueʡ, padd��ͥ��ַ.value��, padd��ͥ��ַ.value����, padd��ͥ��ַ.value����, padd��ͥ��ַ.value��ϸ��ַ)
        Call .padd���ڵ�ַ.LoadStructAdress(padd���ڵ�ַ.valueʡ, padd���ڵ�ַ.value��, padd���ڵ�ַ.value����, padd���ڵ�ַ.value����, padd���ڵ�ַ.value��ϸ��ַ)
        .txt�����.Text = txt�����.Text: .txt�����.MaxLength = txt�����.MaxLength
        .cbo���ʽ.ListIndex = cbo���ʽ.ListIndex
        .cbo�ѱ�.ListIndex = cbo�ѱ�.ListIndex
        .cbo�ѱ�.Locked = cbo�ѱ�.Locked
        .cbo�ѱ�.TabStop = cbo�ѱ�.TabStop
        '�����:40005
        If Not mrsInfo Is Nothing Then
            .txt��ϵ�����֤.Text = Nvl(mrsInfo!��ϵ�����֤��)
            .txt��ϵ������.Text = Nvl(mrsInfo!��ϵ������)
            .txt��ϵ�˵绰.Text = Nvl(mrsInfo!��ϵ�˵绰)
            .cbo��ϵ�˹�ϵ.ListIndex = cbo.FindIndex(.cbo��ϵ�˹�ϵ, Nvl(mrsInfo!��ϵ�˹�ϵ), True)
            If .cbo��ϵ�˹�ϵ.ListIndex = -1 And Nvl(mrsInfo!��ϵ�˹�ϵ) <> "" Then
                .cbo��ϵ�˹�ϵ.ListIndex = 8: .txt������ϵ.Text = Nvl(mrsInfo!��ϵ�˹�ϵ)
            End If
        End If
    End With
    
     With mobjfrmPatiInfo
        txtPatient.Text = .txtPatient.Text  '����Change�¼�
        
        cbo�Ա�.ListIndex = .cbo�Ա�.ListIndex
        txt����.Text = .txt����.Text
        txt����.Tag = txt����.Text
        cbo���䵥λ.ListIndex = .cbo���䵥λ.ListIndex
        Call txt����_Validate(False)
        
        cbo��ͥ��ַ.Text = .cbo��ͥ��ַ.Text
        cbo���ڵ�ַ.Text = .txtRegLocation.Text
        '89242:���ϴ�,2015/12/7,��ȡ���˵�ַ��Ϣ
        Call padd��ͥ��ַ.LoadStructAdress(.padd��ͥ��ַ.valueʡ, .padd��ͥ��ַ.value��, .padd��ͥ��ַ.value����, .padd��ͥ��ַ.value����, .padd��ͥ��ַ.value��ϸ��ַ)
        Call padd���ڵ�ַ.LoadStructAdress(.padd���ڵ�ַ.valueʡ, .padd���ڵ�ַ.value��, .padd���ڵ�ַ.value����, .padd���ڵ�ַ.value����, .padd���ڵ�ַ.value��ϸ��ַ)
        txt�����.Text = .txt�����.Text
        cbo���ʽ.ListIndex = .cbo���ʽ.ListIndex
        cbo�ѱ�.ListIndex = .cbo�ѱ�.ListIndex
        
         
    End With
     
End Function
Private Function isCheckInputIDCard(ByVal strInput As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ���鵱ǰ������Ƿ����֤��
    '��Σ�strInput-�����ֵ
    '����:��������֤��,�򷵻�true,���򷵻�False
    '���ƣ����˺�
    '���ڣ�2010-07-14 16:37:51
    '˵����31182
    '      �Զ�ʶ�����֤,��Ҫ������������ȷ��
    '      a.ǰ׺Ϊ".":��û��
    '      b.ǰ׺����ַ�����Ϊ15λ��18λ(��Ϊ���֤Ŀǰֻ��15λ��18λ����)
    '      c.ǰ׺���и������֤ȡ�����������ڣ���ȡ����ֵ�Ƿ�Ϊ���֤.
    '------------------------------------------------------------------------------------------------------------------------
    Dim strTemp As String, strDate As String
    'If Left(strInput, 1) = "." Then Exit Function
    If Len(strTemp) = 15 Or Len(strTemp) = 18 Then Exit Function '���������ʶ�����.�����Ҫ��ԭ���֤ǰ+1λ
    strDate = zlCommFun.GetIDCardDate(strInput)
    If strDate = "" Then Exit Function
    If IsDate(strDate) = False Then Exit Function
    isCheckInputIDCard = True
End Function

Private Sub cbo���ڵ�ַ_GotFocus()
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub cbo���ڵ�ַ_LostFocus()
    Call zlCommFun.OpenIme
End Sub

Private Sub GetPatient(objCard As zlIDKind.Card, ByVal strInput As String, ByVal blnCard As Boolean, Optional blnInputIDCard As Boolean = False, _
                        Optional ByRef Cancel As Boolean, Optional ByRef blnCertificate As Boolean = False)
    '���ܣ���ȡ������Ϣ
    '������blnCard=�Ƿ���￨ˢ��
    '
    '         blnInputIDCard-�Ƿ����֤ˢ��
    '����:Cancel-Ϊtrue��ʾ���صķ�����ȡ������Ϣ
    Dim strSQL As String, lng�����ID As Long, strPassWord As String, strErrMsg As String
    Dim lng����ID As Long, blnHavePassWord As Boolean
    Dim rsTmp As ADODB.Recordset, strTemp As String
    Dim blnSame As Boolean, blnCancel As Boolean
    Dim cur��� As Currency, curMoney As Currency
    Dim strInputInfo As String '���洫��������ı� ������ʹ�����֤�� �Բ��˽��в��Һ� ���滻��"-" ����ID�����
    Dim i As Integer, strPati As String, strRegist As String, rsRegist As ADODB.Recordset
    Dim vRect As RECT, str����Ժ As String
    Dim blnҽ���� As Boolean, dbl������� As Double
    Dim intMsg As VbMsgBoxResult, blnReload As Boolean
    Dim blnOtherType As Boolean '�Ƿ������
    Dim lngRow As Long, lngCol As Long
    
    strInputInfo = strInput
    lbl����.Caption = ""
    lbl����.Visible = False
    
    On Error GoTo errH
    blnҽ���� = False
    
    If objCard Is Nothing Then Set objCard = IDKind.GetCurCard
     
    strSQL = "Select  A.����ID,A.�����,A.סԺ��,A.���￨��,A.�ѱ�,A.ҽ�Ƹ��ʽ,A.����,A.�Ա�,A.����,A.��������,A.�����ص�,A.���֤��,A.����֤��,A.���,A.ְҵ,A.����,A.��������, " & _
             "A.����,A.����,A.����,A.ѧ��,A.����״��,A.��ͥ��ַ,A.��ͥ�绰,A.��ͥ��ַ�ʱ�,A.�໤��,A.��ϵ������,A.��ϵ�˹�ϵ,A.��ϵ�˵�ַ,A.��ϵ�˵绰,A.���ڵ�ַ, " & _
             "A.���ڵ�ַ�ʱ�,A.Email,A.QQ,A.��ͬ��λid,A.������λ,A.��λ�绰,A.��λ�ʱ�,A.��λ������,A.��λ�ʺ�,A.������,A.������,A.��������,A.����ʱ��,A.����״̬, " & _
             "A.��������,A.סԺ����,A.��ǰ����id,A.��ǰ����id,A.��ǰ����,A.��Ժʱ��,A.��Ժʱ��,A.��Ժ,A.IC����,A.������,A.ҽ����,A.����,A.��ѯ����,A.�Ǽ�ʱ��,A.ͣ��ʱ��,A.����,A.��ϵ�����֤��, " & _
             "B.���� ��������,A.��ѯ���� As ����֤��,A.����ģʽ,A.�ֻ��� From ������Ϣ A,������� B  Where A.���� = B.���(+) And A.ͣ��ʱ�� is NULL "
                 
    If mTy_Para.bln����סԺ���˹Һ� = False Then
        str����Ժ = " And Not Exists(Select 1 From ������ҳ Where ����ID=A.����ID   And ��ҳID<>0 And ��ҳID=A.��ҳID And Nvl(��������,0)=0 And ��Ժ���� is Null)"
    End If
   

    If blnCard And objCard.���� Like "����*" And mstrYBPati = "" And InStr("-+*.", Left(strInput, 1)) = 0 Then     'ˢ��
        If IDKind.Cards.��ȱʡ������ And Not IDKind.GetfaultCard Is Nothing Then
            lng�����ID = IDKind.GetfaultCard.�ӿ����
        ElseIf IDKind.GetCurCard.�ӿ���� > 0 Then
            lng�����ID = IDKind.GetCurCard.�ӿ����
        Else
            If lng�����ID = 0 Then lng�����ID = -1
        End If
        
        '����|�����|ˢ����־|�����ID|���ų���|ȱʡ��־(1-��ǰȱʡ;0-��ȱʡ)|�Ƿ�����ʻ�(1-�����ʻ�;0-�������ʻ�);��
        If gobjSquare.objSquareCard.zlGetPatiID(lng�����ID, strInput, False, lng����ID, strPassWord, strErrMsg) = False Then lng����ID = 0

        If IDKind.IsMobileNO(strInput) And lng����ID = 0 Then
            If gobjSquare.objSquareCard.zlGetPatiID("�ֻ���", strInput, False, lng����ID, strPassWord, strErrMsg) = False Then lng����ID = 0
        End If
        If lng����ID <= 0 Then GoTo NewPati:
        strInput = "-" & lng����ID
        blnHavePassWord = True
        '72868,Ƚ����,2014-5-20,������ҺŹ���Ĳ���������δ��ѡ������סԺ���˹Һš��Ĳ�����������Ժ������Ȼ�ܹ�ֱ��ͨ������ҺŹ���ˢ���Һ�
        strSQL = strSQL & " And A.����ID=[2] " & str����Ժ
        mstr����� = "": txt�����.TabStop = True
    ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then
        '�����
        strSQL = strSQL & " And A.�����=[2]" & str����Ժ
        If InStr(mstrPrivs, ";��������;") > 0 Then
            mstr����� = Mid(strInput, 2) '��¼����������
            txt�����.TabStop = False
        End If
    ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then
        '����ID
        strSQL = strSQL & " And A.����ID=[2]" & _
        IIf(mstrYBPati <> "", "", str����Ժ)
        If mstrYBPati = "" Then mstr����� = "": txt�����.TabStop = True
    ElseIf blnInputIDCard Then  '���������֤ʶ��
        strInput = UCase(strInput)
        If gobjSquare.objSquareCard.zlGetPatiID("���֤", strInput, False, lng����ID, strPassWord, strErrMsg, , , , , , , mblnUserCancel) = False Then lng����ID = 0
        strInput = "-" & lng����ID
        strSQL = strSQL & " And A.����ID=[2] " & str����Ժ
        mstr����� = "": txt�����.TabStop = True
        blnHavePassWord = True
    ElseIf blnCertificate Then
        strInput = UCase(strInput)
        If gobjSquare.objSquareCard.zlGetPatiID(objCard.����, strInput, False, lng����ID, strPassWord, strErrMsg, , , , , , blnCertificate) = False Then Exit Sub
        strInput = "-" & lng����ID
        strSQL = strSQL & " And A.����ID=[2] " & str����Ժ
        mstr����� = "": txt�����.TabStop = True
        blnHavePassWord = True
    ElseIf objCard.���� Like "����*" And IDKind.IsMobileNO(strInput) = True Then
        If gobjSquare.objSquareCard.zlGetPatiID("�ֻ���", strInput, False, lng����ID, strPassWord, strErrMsg) = False Then Exit Sub
        strInput = "-" & lng����ID
        strSQL = strSQL & " And A.����ID=[2] " & str����Ժ
        mstr����� = "": txt�����.TabStop = True
    Else
        Select Case objCard.����
            Case "����", "��������￨"
                '����
                blnSame = False
                If Not mrsInfo Is Nothing Then
                    mstr����� = "": txt�����.TabStop = True
                    If txtPatient.Text = mrsInfo!���� Then blnSame = True
                End If
                If Not blnSame Then
                    If Not gblnSeekName Or gblnSeekName And Len(txtPatient.Text) < 2 Or mstr����� <> "" Or mblnNewCard Then
                        Set mrsInfo = Nothing: Exit Sub
                    Else
                         '�����:50485
                        strPati = _
                            " Select /*+Rule */distinct 1 as ����ID,A.����ID as ID,A.����ID,A.����,A.�Ա�,A.����,A.�����,A.��������,A.���֤��,A.��ͥ��ַ,A.������λ,decode(b.����,Null,Null,'��') As �Ƿ���ҽ�ƿ�,A.�ֻ���,A.����ʱ��" & _
                            " From ������Ϣ A, ����ҽ�ƿ���Ϣ B " & _
                            " Where Rownum <101 And a.����ID=b.����ID(+) And b.״̬(+)=0 And B.�����ID(+)=[3]  And A.ͣ��ʱ�� is NULL And A.���� Like [1]" & str����Ժ & _
                            IIf(gintNameDays = 0, "", " And Nvl(A.����ʱ��,A.�Ǽ�ʱ��)>Trunc(Sysdate-[2])")
                            
                        strPati = strPati & " Union ALL " & _
                                "Select 0,0 as ID,-NULL,'[�²���]',NULL,NULL,-NULL,To_Date(NULL),NULL,NULL,NULL,NULL,NULL,To_Date(NULL) From Dual"
                        strPati = strPati & " Order by ����ID,����"
                            
                        vRect = zlControl.GetControlRect(txtPatient.Hwnd)
                        Set rsTmp = zlDatabase.ShowSQLSelect(Me, strPati, 0, "���˲���", 1, "", "��ѡ����", False, False, True, vRect.Left, vRect.Top, txtPatient.Height, blnCancel, False, True, strInput & "%", gintNameDays, Val(zlDatabase.GetPara("ȱʡҽ�ƿ����", glngSys, mlngModul, 0)))
                        If Not rsTmp Is Nothing Then
                            If rsTmp!ID = 0 Then '�����²���
                                Set mrsInfo = Nothing
                                '82859:���ϴ�,2015/4/8,���˻�����Ϣ����
                                If mbytInState = 0 Then SetPatiInfoEnabled vsfPlan.TextMatrix(vsfPlan.Row, GetCol("����")) <> "", mrsInfo Is Nothing
                                Exit Sub
                            Else '�Բ���ID��ȡ
                                strInput = rsTmp!����ID
                                strSQL = strSQL & " And A.����ID=[1]"
                            End If
                        Else 'ȡ��ѡ��
                            txtPatient.Text = ""
                            Set mrsInfo = Nothing: Exit Sub
                        End If
                    End If
                Else
                    'ͬһ������ʱ��Ҫ���¶�ȡԤ������Ϣ
                    If mbytMode <> 1 And mstrYBPati = "" Then
                        Set rsTmp = GetMoneyInfo(mrsInfo!����ID, , , 1, , , True)
                        cur��� = 0: dbl������� = 0: stbThis.Panels(4).ToolTipText = ""
                        Do While Not rsTmp.EOF
                            cur��� = cur��� + Val(Nvl(rsTmp!Ԥ�����))
                            cur��� = cur��� - Val(Nvl(rsTmp!�������))
                            If Val(Nvl(rsTmp!����)) = 1 Then
                                dbl������� = Val(Nvl(rsTmp!Ԥ�����)) - Val(Nvl(rsTmp!�������))
                            End If
                            rsTmp.MoveNext
                        Loop
                        If cur��� > 0 Then
                            Call ShowDeposit(True): Call ShowMedicareInfo(False)
                            mdblԤ����� = cur���
                            For i = 1 To vsfPay.Rows - 1
                                If vsfPay.RowData(i) = 0 And vsfPay.TextMatrix(i, 0) <> "" Then
                                    vsfPay.TextMatrix(i, 6) = mdblԤ�����
                                End If
                            Next i
                            stbThis.Panels(4).Text = "����Ԥ�����:" & mdblԤ�����
                            If Round(dbl�������, 6) <> 0 Then stbThis.Panels(4).ToolTipText = "������Ԥ����" & Format(dbl�������, "0.00")
                            
                            'ҽ��վ�Һ�ȱʡʹ��Ԥ����
                            curMoney = GetRegistMoney
                            '77786,Ƚ����,2014-9-2,��ѡ����ʹ��Ԥ����ɿ�,�Һ�ʱ,û��Ĭ�ϼ��ٳ��
                            '74550,Ƚ����,2014-7-2,�ڲ�����Ժ����,ҽ��������ҽ��վ�Һ�ʱ�ܹ�ѡ����㷽ʽ(��������Ϊ7��һ��ͨ����)
                            If (gblnPrePayPriority Or (mblnStation And Not mblnStationPrice And Not cbo���㷽ʽ.Visible)) And cur��� >= curMoney Then
'                                txtԤ��֧��.Text = Format(curMoney, "0.00")
                            Else
'                                txtԤ��֧��.Text = "0.00"
                            End If
                        End If
                    End If
                    Call zlQueryEMPIPatiInfo
                    Exit Sub
                End If
            Case "ҽ����"
                strInput = UCase(strInput)
                mstr����� = "": txt�����.TabStop = True
                blnҽ���� = True
                If mblnOlnyBJYB And zlCommFun.ActualLen(strInput) >= 9 Then
                    '������ҽ������Ч:������:����:26982
                    strSQL = strSQL & " And A.ҽ���� like [3] " & str����Ժ
                    strTemp = Left(strInput, 9) & "%"
                Else
                    strSQL = strSQL & " And A.ҽ����=[1]" & str����Ժ
                End If
            Case "�ֻ���"
                If IDKind.IsMobileNO(strInput) = False Then Exit Sub
                If gobjSquare.objSquareCard.zlGetPatiID("�ֻ���", strInput, False, lng����ID, strPassWord, strErrMsg) = False Then Exit Sub
                strInput = "-" & lng����ID
                strSQL = strSQL & " And A.����ID=[2] " & str����Ժ
                mstr����� = "": txt�����.TabStop = True
            Case "���֤��", "���֤", "�������֤"
                strInput = UCase(strInput)
                If gobjSquare.objSquareCard.zlGetPatiID("���֤", strInput, False, lng����ID, strPassWord, strErrMsg, , , , , , , mblnUserCancel) = False Then lng����ID = 0
                strInput = "-" & lng����ID
                strSQL = strSQL & " And A.����ID=[2] " & str����Ժ
                mstr����� = "": txt�����.TabStop = True
                blnHavePassWord = True
            Case "IC����", "IC��"
                strInput = UCase(strInput)
                mstr����� = "": txt�����.TabStop = True
                If gobjSquare.objSquareCard.zlGetPatiID("IC��", strInput, False, lng����ID, strPassWord, strErrMsg) = False Then lng����ID = 0
                strInput = "-" & lng����ID
                strSQL = strSQL & " And A.����ID=[2] " & str����Ժ
                blnHavePassWord = True
            Case "�����"
                If Not IsNumeric(strInput) Then strInput = "0"
                '72868,Ƚ����,2014-5-20,������ҺŹ����& str����Ժ����������δ��ѡ������סԺ���˹Һš��Ĳ�����������Ժ������Ȼ�ܹ�ֱ��ͨ������ҺŹ���ˢ���Һ�
                strSQL = strSQL & " And A.�����=[1]" & str����Ժ
                If InStr(mstrPrivs, ";��������;") > 0 Then
                    mstr����� = strInput
                    txt�����.TabStop = False
                End If

             Case Else
                '��������,��ȡ��صĲ���ID
                If objCard.�ӿ���� > 0 Then
                    lng�����ID = objCard.�ӿ����
                    If gobjSquare.objSquareCard.zlGetPatiID(lng�����ID, strInput, False, lng����ID, strPassWord, strErrMsg) = False Then lng����ID = 0
                    blnOtherType = True
                Else
                    If gobjSquare.objSquareCard.zlGetPatiID(objCard.����, strInput, False, lng����ID, _
                        strPassWord, strErrMsg) = False Then lng����ID = 0
                End If
                'If lng����ID <= 0 Then GoTo NotFoundPati:
                '72868,Ƚ����,2014-5-20,������ҺŹ���Ĳ���������δ��ѡ������סԺ���˹Һš��Ĳ�����������Ժ������Ȼ�ܹ�ֱ��ͨ������ҺŹ���ˢ���Һ�
                strSQL = strSQL & " And A.����ID=[2]" & str����Ժ
                strInput = "-" & lng����ID
                blnHavePassWord = True
        End Select
    End If
    If blnInputIDCard And Not mrsInfo Is Nothing Then
        If mrsInfo.State <> 1 Then GoTo ReadPati:
        'ԭ���в���,�ְ����֤��ȡ,����ܴ��ڲ����֤�����:
        '1.���δ�ҵ�,���ǲ���֤
        '2.����ҵ���,�����µĲ���Ϊ׼(ͨ����ʾ��ѡ��)
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strInput, Mid(strInput, 2), strTemp)
        If rsTmp.EOF Then
            mobjfrmPatiInfo.txt���֤�� = txtIDCard.Text
            Call zlQueryEMPIPatiInfo
            Exit Sub
        End If
        If Nvl(rsTmp!����) <> Trim(txtPatient.Text) And Trim(txtPatient.Text) <> "" Then
            If gbln���֤Ψһ Then
                intMsg = MsgBox("ע��:" & vbCrLf & _
                                 "      ¼������֤�ŵ�����Ϊ��" & Nvl(rsTmp!����) & " ����¼��������" & Trim(txtPatient.Text) & " ��" & vbCrLf & _
                                 "      ��һ��,�Ƿ��ԡ�" & Nvl(rsTmp!����) & " ��Ϊ׼���йҺţ�", vbQuestion + vbYesNo, gstrSysName)
                If intMsg = vbNo Then intMsg = vbCancel
            Else
            
                intMsg = MsgBox("ע��:" & vbCrLf & _
                                 "      ¼������֤�ŵ�����Ϊ��" & Nvl(rsTmp!����) & " ����¼��������" & Trim(txtPatient.Text) & " ��" & vbCrLf & _
                                 "      ��һ��,����!   " & vbCrLf & _
                                 "���ǡ���ʾ�����֤���ҵĲ��˽��йҺ�" & vbCrLf & _
                                 "���񡿱�ʾ��������������йҺ�,���֤�Ÿ���Ϊ��ǰ¼������֤��" & vbCrLf & _
                                 "��ȡ������ʾ���֤��¼�����,����¼�����֤��" & vbCrLf & _
                                "", vbQuestion + vbYesNoCancel, gstrSysName)
            End If
            If intMsg = vbCancel Then
              
                Cancel = True: Exit Sub
            End If
            If intMsg = vbYes Then
                Set mrsInfo = rsTmp
                txtPatient.Text = Nvl(rsTmp!����)
                blnReload = True
            End If
            If intMsg = vbNo Then
                mobjfrmPatiInfo.txt���֤�� = txtIDCard.Text
            End If
        End If
    Else
ReadPati:
        Set mrsInfo = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strInput, Mid(strInput, 2), strTemp)
    End If
    
    '82859:���ϴ�,2015/4/8,���˻�����Ϣ����
    If mbytInState = 0 Then SetPatiInfoEnabled vsfPlan.TextMatrix(vsfPlan.Row, GetCol("����")) <> "", True
        
    strInput = strInputInfo
    Call ClearmobjfrmPatiInfoFace(IIf(mblnNewCard, False, True))
    If blnInputIDCard Then mobjfrmPatiInfo.txt���֤��.Text = strInput
    If Not mrsInfo.EOF Then
         '�ڷ���ʱ ������Ա ʹ�ò��˵�ҽ�ƿ���ȡ������Ϣʱ �����·��Ŀ��Ͳ���ӵ�еĿ���ͬ�����͵������
         'ʹ��ԭ���Ŀ� ���ٷ���������
         If mblnNewCard And mbytMode = 0 And blnCard And lng�����ID = gCurSendCard.lng�����ID Then
              mblnNewCard = False
              Call ClearmobjfrmPatiInfoFace(IIf(mblnNewCard, False, True))
         End If
        '31182:��������֤���ҵĲ����Ƿ������������һ��
        If mbytMode = 1 Or mbytMode = 2 Then
            Call zlAutoCalcBackLists(Val(Nvl(mrsInfo!����ID))) '�Զ����������
        End If
        If blnInputIDCard Then
                If Nvl(mrsInfo!����) <> Trim(txtPatient.Text) And Trim(txtPatient.Text) <> "" Then
                    If gbln���֤Ψһ Then
                        intMsg = MsgBox("ע��:" & vbCrLf & _
                                         "      ¼������֤�ŵ�����Ϊ��" & Nvl(mrsInfo!����) & " ����¼��������" & Trim(txtPatient.Text) & " ��" & vbCrLf & _
                                         "      ��һ��,�Ƿ��ԡ�" & Nvl(mrsInfo!����) & " ��Ϊ׼���йҺţ�", vbQuestion + vbYesNo, gstrSysName)
                        If intMsg = vbNo Then intMsg = vbCancel
                    Else
                    
                            intMsg = MsgBox("ע��:" & vbCrLf & _
                                             "      ¼������֤�ŵ�����Ϊ��" & Nvl(mrsInfo!����) & " ����¼��������" & Trim(txtPatient.Text) & " ��" & vbCrLf & _
                                             "      ��һ��,����!   " & vbCrLf & _
                                             "���ǡ���ʾ�����֤���ҵĹҺŶ��� " & vbCrLf & _
                                             "���񡿱�ʾ�����������Ϊ�ҺŶ�����Ҫ���½������˵���" & vbCrLf & _
                                             "��ȡ������ʾ���֤��¼�����,����¼�����֤��" & vbCrLf & _
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
            If Nvl(mrsInfo!����) <> Trim(txtPatient.Text) And Trim(txtPatient.Text) <> "" Then
                intMsg = MsgBox("ע��:" & vbCrLf & _
                                 "      ¼���֤�����������Ϊ��" & Nvl(mrsInfo!����) & " ����¼��������" & Trim(txtPatient.Text) & " ��" & vbCrLf & _
                                 "      ����Ϣ��һ��,�Ƿ���֤�����ҵ�����Ϊ�ҺŶ���   " & vbCrLf & _
                                "", vbQuestion + vbYesNo, gstrSysName)
                If intMsg = vbNo Then
                    Cancel = True: Exit Sub
                End If
            End If
        End If
        
        '102230,������Ҳ����ӿ�
        If (mbytMode = 0 Or mbytMode = 1) And mbytInState = 0 _
            And Not (mbytMode = 0 And mbytInState = 0 And mstrNoIn <> "") Then
            If PatiValiedCheckByPlugIn(mlngModul, Val(Nvl(mrsInfo!����ID)), _
                "<YSXM>" & NeedName(cboҽ��.Text) & "</YSXM>") = False Then
                Set mrsInfo = Nothing: txtPatient.Text = ""
                Cancel = True:  Exit Sub
            End If
        End If
        
        If Not IsNull(mrsInfo!��������) Then
            lbl����.Caption = "" & mrsInfo!��������
            lbl����.Visible = True
        End If
        
        txtPatient.Text = Nvl(mrsInfo!����) '�����Change�¼�
        '�ڵ���txtPatient_Change�¼���������źͲ���������Ϊ�յ������ �޷�ʶ��ò�����Ϣ ���ִ���
        '���������ݿ����ݴ����ٽ��к����Ĵ���
        If mrsInfo Is Nothing Then Cancel = True: Exit Sub
        '74428�����ϴ���2014-7-8������������ʾ��ɫ����
        Call SetPatiColor(txtPatient, Nvl(mrsInfo!��������), IIf(IsNull(mrsInfo!��������), txtPatient.ForeColor, vbRed))

        '113999:���ϴ�,2017/11/14,���ݷ������ʽ��п���
        If Check��������(Val(Nvl(mrsInfo!����ID)), IIf(mCurSendCard.lng�����ID = 0, gCurSendCard.lng�����ID, mCurSendCard.lng�����ID), Trim(mobjfrmPatiInfo.txt����) <> "") = True Then
            cmdCard.Enabled = True
        Else
            cmdCard.Enabled = gCurSendCard.lng�������� <> 1
            mobjfrmPatiInfo.mstrCard = ""
            mobjfrmPatiInfo.txt����.Text = ""
            mobjfrmPatiInfo.txt����.Text = ""
            mobjfrmPatiInfo.txt��֤.Text = ""
            If mblnNoClearPrompt = False Then lblPrompt.Caption = ""
            mblnNewCard = False
            mblnAddCardItem = False
        End If
        cmdCard.Enabled = cmdCard.Enabled And Not (mblnStation And mTy_Para.bln�Һű���ˢ��)
        mstrPassWord = strPassWord
        If Not blnHavePassWord Then mstrPassWord = Nvl(mrsInfo!����֤��)
        cbo�Ա�.ListIndex = cbo.FindIndex(cbo�Ա�, Nvl(mrsInfo!�Ա�), True) '�����ں�����ݳ���������
        cbo��ͥ��ַ.Text = IIf(Nvl(mrsInfo!��ͥ��ַ) = "", Nvl(mrsInfo!���ڵ�ַ), Nvl(mrsInfo!��ͥ��ַ))
        cbo���ڵ�ַ.Text = Nvl(mrsInfo!���ڵ�ַ)
        '89242:���ϴ�,2015/12/7,��ȡ���˵�ַ��Ϣ
        Call zlReadAddrInfo(padd��ͥ��ַ, Val(Nvl(mrsInfo!����ID)), 0, 3, cbo��ͥ��ַ.Text)
        Call zlReadAddrInfo(padd���ڵ�ַ, Val(Nvl(mrsInfo!����ID)), 0, 4, cbo���ڵ�ַ.Text)
        txtPatient.PasswordChar = ""
        
        '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
        txtPatient.IMEMode = 0
        Call zlControl.CboSetIndex(cbo�ѱ�.Hwnd, cbo.FindIndex(cbo�ѱ�, "" & mrsInfo!�ѱ�, True))
        
        If Not blnInputIDCard Or blnReload Or txt�����.Text = "" Then
            txt�����.Text = Nvl(mrsInfo!�����, mstr�����)
'            txt�����.Enabled = (Val(txt�����.Text) = 0)
        End If
        
        If txt�����.Text = "" And txt�����.Enabled And gbln�Զ������ Then
            txt�����.Text = zlGet�����
        End If
        
        If blnReload Then
            txtIDCard.Text = Nvl(mrsInfo!���֤��, txtIDCard.Text) '���֤��:31182
            txtIDCard.Tag = Nvl(mrsInfo!���֤��, txtIDCard.Text)  '�Ա㷴�����ٲ�
        Else
            If Not blnInputIDCard Then
                txtIDCard.Text = Nvl(mrsInfo!���֤��)
                txtIDCard.Tag = Nvl(mrsInfo!���֤��)
            Else
                txtIDCard.Tag = txtIDCard.Text
            End If
        End If
    
        'ҽ�Ƹ��ʽ
        If Not IsNull(mrsInfo!ҽ�Ƹ��ʽ) Then
            cbo���ʽ.ListIndex = cbo.FindIndex(cbo���ʽ, mrsInfo!ҽ�Ƹ��ʽ, True)
        ElseIf mstrYBPati <> "" Then
            cbo���ʽ.ListIndex = cbo.FindIndex(cbo���ʽ, "1", True)
        End If
        If Not IsNull(mrsInfo!ҽ����) And mlngOutModeMC <> 0 Then Call SetCboDefault(cboҽ�����)
        
        If Not blnInputIDCard Or blnReload Then
            txt��������.Text = Format(IIf(IsNull(mrsInfo!��������), "____-__-__", mrsInfo!��������), "YYYY-MM-DD")
            If Not IsNull(mrsInfo!��������) Then
                txt����.Text = ReCalcOld(CDate(mrsInfo!��������), cbo���䵥λ, mrsInfo!����ID) '���ݳ���������������
                
                txt����ʱ��.Text = Format(mrsInfo!��������, "HH:MM")
            Else
                txt����ʱ��.Text = "__:__"
                txt��������.Text = ReCalcBirth(txt����.Text, cbo���䵥λ.Text)
            End If
        End If
        
        '��ϸ������Ϣ����
        txt֤��.Tag = "": txt֤��.Text = ""
        Call CopyInfoTofrmPatiInfo
        With mobjfrmPatiInfo
    
            If mblnOlnyBJYB And blnҽ���� Then
                .txtPatiMCNO(0).Text = strInput
            Else
                .txtPatiMCNO(0).Text = "" & Nvl(mrsInfo!ҽ����)
            End If
            .txtPatiMCNO(0).Tag = "" & Nvl(mrsInfo!ҽ����)
            .txtPatiMCNO(1).Text = .txtPatiMCNO(0).Text
            If Not blnInputIDCard Or blnReload Then
                Call LoadOldData("" & mrsInfo!����, .txt����, .cbo���䵥λ)
                .mblnChange = False
                .txt��������.Text = Format(IIf(IsNull(mrsInfo!��������), "____-__-__", mrsInfo!��������), "YYYY-MM-DD")
                .mblnChange = True
                
                If Not IsNull(mrsInfo!��������) Then
                    .txt����.Text = ReCalcOld(CDate(.txt��������.Text), .cbo���䵥λ, mrsInfo!����ID) '���ݳ���������������
                    .txt����.Tag = .txt����.Text
                    If CDate(.txt��������.Text) - CDate(mrsInfo!��������) <> 0 Then .txt����ʱ��.Text = Format(mrsInfo!��������, "HH:MM")
                Else
                    .txt����ʱ��.Text = "__:__"
                    .mblnChange = False
                    .txt��������.Text = ReCalcBirth(.txt����.Text, .cbo���䵥λ.Text)
                    .mblnChange = True
                End If
            End If
            
            Call SetmobjfrmPatiInfo
            '90875:���ϴ�,2016/8/19,��֤���б��л�ȡ��ǰ֤�����͵ĺ���
            If IDKind֤��.IDKind <> IDKind֤��.GetKindIndex("���֤��") Then
                With mobjfrmPatiInfo.vsCertificate
                    For lngRow = 1 To .Rows - 1
                        For lngCol = 0 To .Cols - 1 Step 2
                            If .TextMatrix(lngRow, lngCol) = IDKind֤��.GetCurCard.���� Then
                                txt֤��.Tag = .TextMatrix(lngRow, lngCol + 1)
                                txt֤��.Text = txt֤��.Tag
                                Exit For
                            End If
                        Next
                    Next
                End With
            End If
                
            txt����.Text = .txt����.Text
            txt����.Tag = txt����.Text
            cbo���䵥λ.ListIndex = .cbo���䵥λ.ListIndex
            cbo���䵥λ.Tag = cbo���䵥λ.Text
            Call txt����_Validate(False)
            
            If mlng�Һſ���ID > 0 Then .chk����.Value = IIf(Check����(mrsInfo!����ID, mlng�Һſ���ID), 1, 0)
            If mbytMode = 1 And Not blnInputIDCard Then
                .txt���֤�� = txtIDCard.Text
            End If
            .mstr���֤�� = Nvl(mrsInfo!���֤��)
            imgPatiPic.Picture = .imgPatient.Picture
            txt��ͥ�绰.Text = .txt��ͥ�绰
            .mstr�������� = .txt��������.Text
            .mstr����ʱ�� = .txt����ʱ��.Text
            .mstr���䵥λ = IIf(IsNumeric(txt����.Text), cbo���䵥λ.Text, "")
            .mstr���� = txt����.Text
            .mstr�Ա� = NeedName(cbo�Ա�.Text)
            .mstr���� = txtPatient.Text
            .mstr���֤�� = txtIDCard.Text
            mstr�������� = .txt��������.Text
            .txtMobile.Text = Nvl(mrsInfo!�ֻ���)
        End With
        mstr���䵥λ = IIf(IsNumeric(txt����.Text), cbo���䵥λ.Text, "")
        mstr���� = txt����.Text
        mstr�Ա� = NeedName(cbo�Ա�.Text)
        mstr���� = txtPatient.Text
        
        '����Ԥ������Ϣ
        If mbytMode <> 1 And mstrYBPati = "" Then
            Set rsTmp = GetMoneyInfo(mrsInfo!����ID, , , 1, , , True)
            cur��� = 0: dbl������� = 0: stbThis.Panels(4).ToolTipText = ""
            Do While Not rsTmp.EOF
                cur��� = cur��� + Val(Nvl(rsTmp!Ԥ�����))
                cur��� = cur��� - Val(Nvl(rsTmp!�������))
                If Val(Nvl(rsTmp!����)) = 1 Then
                    dbl������� = Val(Nvl(rsTmp!Ԥ�����)) - Val(Nvl(rsTmp!�������))
                End If
                rsTmp.MoveNext
            Loop
            If cur��� > 0 Then
                Call ShowMedicareInfo(False): Call ShowDeposit(True)
                stbThis.Panels(4).Text = "����Ԥ�����:" & Format(cur���, "0.00")
                stbThis.Panels(4).AutoSize = sbrContents
                
                mdblԤ����� = cur���
                For i = 1 To vsfPay.Rows - 1
                    If vsfPay.RowData(i) = 0 And vsfPay.TextMatrix(i, 0) <> "" Then
                        vsfPay.TextMatrix(i, 6) = mdblԤ�����
                    End If
                Next i
                If Round(dbl�������, 6) <> 0 Then stbThis.Panels(4).ToolTipText = "������Ԥ����" & Format(dbl�������, "0.00")
                
                'ҽ��վ�Һ�ȱʡʹ��Ԥ����
                curMoney = GetRegistMoney
                '77786,Ƚ����,2014-9-2,��ѡ����ʹ��Ԥ����ɿ�,�Һ�ʱ,û��Ĭ�ϼ��ٳ��
                '74550,Ƚ����,2014-7-2,�ڲ�����Ժ����,ҽ��������ҽ��վ�Һ�ʱ�ܹ�ѡ����㷽ʽ(��������Ϊ7��һ��ͨ����)
                If (gblnPrePayPriority Or (mblnStation And Not mblnStationPrice And Not cbo���㷽ʽ.Visible)) And cur��� >= curMoney Then
'                    txtԤ��֧��.Text = Format(curMoney, "0.00")
                Else
'                    txtԤ��֧��.Text = "0.00"
                End If
            Else
                Call ShowDeposit(False)
            End If
        End If
        mstr����� = "": txt�����.TabStop = True
        mblnIDCardKind = False
        Call zlQueryEMPIPatiInfo
    Else
NewPati:
        txt�����.Enabled = True
        
        '82859:���ϴ�,2015/4/8,���˻�����Ϣ����
        If mbytInState = 0 Then SetPatiInfoEnabled vsfPlan.TextMatrix(vsfPlan.Row, GetCol("����")) <> "", mrsInfo Is Nothing
        
        mblnIDCardKind = False
        If objCard.���� Like "����*" And blnCard = False Then
            lng�����ID = 0
        End If
        If Not (mblnCard Or IsCardType(IDKind, "IC��") _
            Or (gCurSendCard.lng�����ID = lng�����ID And lng�����ID > 0)) And blnInputIDCard = False And lng�����ID <= 0 Then txtPatient.Text = ""    'ˢ��ʱ�������,��Ϊ����Ƿ��¿�Ҫ�Դ˴��뵯������
        
        If lng����ID = 0 And lng�����ID <> gCurSendCard.lng�����ID Then
            If lng�����ID <= 0 And Not IDKind.GetfaultCard Is Nothing Then lng�����ID = IDKind.GetfaultCard.�ӿ����
            If lng�����ID <> 0 And lng�����ID <> gCurSendCard.lng�����ID Then
                Call InitSendCardPreperty(mlngModul, lng�����ID)
                 
                 cmdCard.ToolTipText = "��" & gCurSendCard.str������ & ": F10"
            End If
           If lng�����ID <= 0 And blnOtherType Then Cancel = True: txtPatient.Text = ""
        End If
            
        If isCheckInputIDCard(strInput) Then
            Dim str���䵥λ As String, str���� As String
            txtIDCard.Text = strInput     '���֤��:31182
            txtIDCard.Tag = strInput
            
            strTemp = zlGetIDCardSex(strInput)
            zlControl.CboLocate cbo�Ա�, strTemp
            zlControl.CboLocate mobjfrmPatiInfo.cbo�Ա�, strTemp
            
            mobjfrmPatiInfo.txt���֤�� = strInput
            mobjfrmPatiInfo.txt�������� = zlCommFun.GetIDCardDate(strInputInfo)
            If txt����.Text = "" Then
                str���� = zlGetIDCardAge(mobjfrmPatiInfo.txt��������, str���䵥λ)
                If str���䵥λ <> "" Then
                    zlControl.CboLocate cbo���䵥λ, str���䵥λ
                    txt����.Text = str����
                     zlControl.CboLocate mobjfrmPatiInfo.cbo���䵥λ, str���䵥λ
                      mobjfrmPatiInfo.txt����.Text = str����
                      mobjfrmPatiInfo.txt����.Tag = str����
                End If
            End If
            '67213:���ϴ�,2014/10/23,�������֤�ϵ���Ϣ
            mblnIDCardKind = IDKind.IDKind = IDKind.GetKindIndex("���֤��")
            If mblnIDCardKind Then
                IDKind.IDKind = IDKind.GetKindIndex("����")
            End If
            mblnIDCardKind = blnInputIDCard Or IDKind.IDKind = IDKind.GetKindIndex("���֤��")
        End If
        Set mrsInfo = Nothing
    End If
    
    If Not mrsInfo Is Nothing Then
        If mrsInfo.State = 1 Then
            If mrsInfo.RecordCount <> 0 Then
                If Not IsNull(mrsInfo!����ID) Then
                    strRegist = "Select a.�ű�, b.���� As ����, a.ִ���� As ҽ��, d.���� As ��Ŀ, a.�Ǽ�ʱ��" & vbNewLine & _
                                "From ���˹Һż�¼ A, ���ű� B, ������ü�¼ C, �շ���ĿĿ¼ D " & vbNewLine & _
                                "Where a.����id = [1] And a.��¼״̬ = 1 And a.No = c.No(+) And c.��¼���� = 4 And c.��� = 1 And c.�շ�ϸĿId = d.Id And a.��¼���� = 1 And a.ִ�в���id = b.Id" & vbNewLine & _
                                "Order By a.�Ǽ�ʱ�� Desc"
                    Set rsRegist = zlDatabase.OpenSQLRecord(strRegist, Me.Caption, mrsInfo!����ID)
                    If Not rsRegist.EOF Then
                        stbThis.Panels(2).Text = "�ϴιҺ�:" & "����:" & rsRegist!�ű� & ",����:" & rsRegist!���� & IIf(IsNull(rsRegist!ҽ��), "", ",ҽ��:" & rsRegist!ҽ��) & ",��Ŀ:" & rsRegist!��Ŀ & ",ʱ��:" & Format(rsRegist!�Ǽ�ʱ��, "yyyy-mm-dd hh:mm:ss")
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
    '���ܣ���EMPIƽ̨��ȡ������Ϣ
    '���ڣ�2016/10/9 10:47:13
    '���ƣ����ϴ�
    '˵����101170
    Dim rsTmp As ADODB.Recordset, lng����ID As Long, strDiff As String, strMsgInfo As String
    Dim strSQL As String
    If mblnNotEMPIQuery Then Exit Sub
    If CreatePlugInOK(mlngModul) = False Then Exit Sub
    If Trim(txtPatient.Text) = "" Then Exit Sub
    If mbytMode <> 0 And mbytMode <> 2 Or mbytInState <> 0 Or chkCancel.Value = 1 Then Exit Sub

    On Error GoTo Errhand
    If zlInitMEPIPati(rsTmp) = False Then Exit Sub
    
    If mrsInfo Is Nothing Then
        lng����ID = 0
    ElseIf mrsInfo.State = 0 Then
        lng����ID = 0
    Else
        lng����ID = Val(Nvl(mrsInfo!����ID))
    End If
    
    With rsTmp
        .AddNew
        !����ID = lng����ID
        !����� = txt�����.Text
        !ҽ���� = mobjfrmPatiInfo.txtPatiMCNO(0).Text
        !���֤�� = mobjfrmPatiInfo.txt���֤��.Text
        !���� = txtPatient.Text
        !�Ա� = zlStr.NeedName(cbo�Ա�.Text)
        If IsDate(txt��������.Text) Then
            !�������� = Format(txt��������.Text & " " & IIf(IsDate(txt����ʱ��.Text), txt����ʱ��.Text, "00:00"), "YYYY-MM-DD HH:MM")
        Else
            !�������� = ""
        End If
        !�����ص� = mobjfrmPatiInfo.txtBirthLocation.Text
        !���� = zlStr.NeedName(mobjfrmPatiInfo.cbo����.Text)
        !���� = zlStr.NeedName(mobjfrmPatiInfo.cbo����.Text)
        !ְҵ = zlStr.NeedName(mobjfrmPatiInfo.cboְҵ.Text)
        !������λ = mobjfrmPatiInfo.txt��λ����.Text
        !����״�� = zlStr.NeedName(mobjfrmPatiInfo.cbo����.Text)
        !��ͥ�绰 = mobjfrmPatiInfo.txt��ͥ�绰.Text
        !��ϵ�˵绰 = mobjfrmPatiInfo.txt��ϵ�˵绰.Text
        !��λ�绰 = mobjfrmPatiInfo.txt��λ�绰.Text
        !��ͥ��ַ = cbo��ͥ��ַ.Text
        !��ͥ��ַ�ʱ� = mobjfrmPatiInfo.txt��ͥ�ʱ�.Text
        !���ڵ�ַ = cbo���ڵ�ַ.Text
        !���ڵ�ַ�ʱ� = mobjfrmPatiInfo.txt���ڵ�ַ�ʱ�.Text
        !��λ�ʱ� = mobjfrmPatiInfo.txt��λ�ʱ�.Text
        !��ϵ������ = mobjfrmPatiInfo.txt��ϵ������.Text
        !��ϵ�˹�ϵ = zlStr.NeedName(mobjfrmPatiInfo.cbo��ϵ�˹�ϵ.Text)
        .Update
    End With
    'EMPIû���ҵ�������Ϣ,ֱ�ӷ���
    Dim rsOut As New ADODB.Recordset
    Err = 0: On Error Resume Next
    mlngEMPI����ID = 0
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
        '104905:���ϴ�,2017/1/12,����EMPI���صĲ���ID�����Ҳ���
        '���ղ����˺ſ϶��в���ID
        mlngEMPI����ID = Val(Nvl(!����ID))
        If lng����ID <> mlngEMPI����ID And mlngEMPI����ID <> 0 Then
            mblnNotEMPIQuery = True
            Call GetPatient(IDKind.GetCurCard, "-" & mlngEMPI����ID, False)
            mblnNotEMPIQuery = False
            If mrsInfo.EOF Then
                lng����ID = 0
            Else
                lng����ID = mlngEMPI����ID
            End If
        End If
        
        mobjfrmPatiInfo.mstrPlugChange = ""
        If Nvl(!ҽ����) <> "" Then
            mobjfrmPatiInfo.txtPatiMCNO(0).Text = Nvl(!ҽ����)
            mobjfrmPatiInfo.txtPatiMCNO(1).Text = mobjfrmPatiInfo.txtPatiMCNO(0).Text
        End If
        If mbln������Ϣ���� Or lng����ID = 0 Then
            If Nvl(!���֤��) <> "" Then txtIDCard.Text = Nvl(!���֤��)
            If Nvl(!����) <> "" Then txtPatient.Text = Nvl(!����): mstrPrePati = Nvl(!����)
            If Nvl(!�Ա�) <> "" Then cbo�Ա�.ListIndex = cbo.FindIndex(cbo�Ա�, Nvl(!�Ա�), True)
            If Nvl(!��������) <> "" Then
                txt��������.Text = Format(Nvl(!��������), "YYYY-MM-DD")
                txt����ʱ��.Text = Format(Nvl(!��������), "HH:MM")
            End If
        Else
            If Nvl(!����) <> "" And txtPatient.Text <> Nvl(!����) Then strDiff = ",����"
            If Nvl(!�Ա�) <> "" And cbo�Ա�.ListIndex <> cbo.FindIndex(cbo�Ա�, Nvl(!�Ա�), True) Then strDiff = strDiff & ",�Ա�"
            If Nvl(!��������) <> "" And Format(Nvl(!��������), "YYYY-MM-DD HH:MM:SS") <> Format(txt��������.Text & " " & txt����ʱ��.Text, "YYYY-MM-DD HH:MM:SS") Then strDiff = strDiff & ",��������"
            If Nvl(!���֤��) <> "" And txtIDCard.Text <> Nvl(!���֤��) Then strDiff = strDiff & ",���֤��"
        End If
        If InStr(";" & mstrPrivs & ";", ";�����޸������;") > 0 And Exist�����(Nvl(!�����), lng����ID) = False Then
            If Nvl(!�����) <> "" Then txt�����.Text = Nvl(!�����)
        Else
            If Nvl(!�����) <> "" And txt�����.Text <> Nvl(!�����) Then strDiff = strDiff & ",�����"
        End If
        If Nvl(!�����ص�) <> "" Then mobjfrmPatiInfo.txtBirthLocation.Text = Nvl(!�����ص�)
        If Nvl(!����) <> "" Then mobjfrmPatiInfo.cbo����.ListIndex = cbo.FindIndex(mobjfrmPatiInfo.cbo����, Nvl(!����), True)
        If Nvl(!����) <> "" Then mobjfrmPatiInfo.cbo����.ListIndex = cbo.FindIndex(mobjfrmPatiInfo.cbo����, Nvl(!����), True)
        If Nvl(!ְҵ) <> "" Then mobjfrmPatiInfo.cboְҵ.ListIndex = cbo.FindIndex(mobjfrmPatiInfo.cboְҵ, Nvl(!ְҵ))
        If Nvl(!������λ) <> "" Then mobjfrmPatiInfo.txt��λ����.Text = Nvl(!������λ)
        If Nvl(!����״��) <> "" Then mobjfrmPatiInfo.cbo����.ListIndex = cbo.FindIndex(mobjfrmPatiInfo.cbo����, Nvl(!����״��), True)
        If Nvl(!��ͥ�绰) <> "" Then txt��ͥ�绰.Text = Nvl(!��ͥ�绰)
        If Nvl(!��ϵ�˵绰) <> "" Then mobjfrmPatiInfo.txt��ϵ�˵绰.Text = Nvl(!��ϵ�˵绰)
        If Nvl(!��λ�绰) <> "" Then mobjfrmPatiInfo.txt��λ�绰.Text = Nvl(!��λ�绰)
        If Nvl(!��ͥ��ַ) <> "" Then cbo��ͥ��ַ.Text = Nvl(!��ͥ��ַ): padd��ͥ��ַ.Value = Nvl(!��ͥ��ַ)
        If Nvl(!��ͥ��ַ�ʱ�) <> "" Then mobjfrmPatiInfo.txt��ͥ�ʱ�.Text = Nvl(!��ͥ��ַ�ʱ�)
        If Nvl(!���ڵ�ַ) <> "" Then cbo���ڵ�ַ.Text = Nvl(!���ڵ�ַ): padd���ڵ�ַ.Value = Nvl(!���ڵ�ַ)
        If Nvl(!���ڵ�ַ�ʱ�) <> "" Then mobjfrmPatiInfo.txt���ڵ�ַ�ʱ�.Text = Nvl(!���ڵ�ַ�ʱ�)
        If Nvl(!��λ�ʱ�) <> "" Then mobjfrmPatiInfo.txt��λ�ʱ�.Text = Nvl(!��λ�ʱ�)
        If Nvl(!��ϵ������) <> "" Then mobjfrmPatiInfo.txt��ϵ������.Text = Nvl(!��ϵ������)
        If Nvl(!��ϵ�˹�ϵ) <> "" Then mobjfrmPatiInfo.cbo��ϵ�˹�ϵ.ListIndex = cbo.FindIndex(mobjfrmPatiInfo.cbo��ϵ�˹�ϵ, Nvl(!��ϵ�˹�ϵ), True)
    End With
    Err = 0: On Error GoTo 0
    Call CopyInfoTofrmPatiInfo
    If lng����ID <> 0 Then
        If strDiff <> "" Then strDiff = Mid(strDiff, 2)
        If mobjfrmPatiInfo.mstrPlugChange <> "" Then mobjfrmPatiInfo.mstrPlugChange = Mid(mobjfrmPatiInfo.mstrPlugChange, 2)
        If strDiff <> "" Then
            strMsgInfo = "���˵� " & strDiff & " ��EMPI��Ϣ��һ�£�������������Ӧ��Ȩ�޻�������������Ϣ��ͻ�����β�����и��¡�"
        End If
        If mobjfrmPatiInfo.mstrPlugChange <> "" Then
            If strMsgInfo <> "" Then strMsgInfo = strMsgInfo & vbNewLine
            strMsgInfo = strMsgInfo & "���˵� " & mobjfrmPatiInfo.mstrPlugChange & " ����EMPI��Ϣ�����˵���,��ע���飡"
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
    '����:�ϴ�������Ϣ��EMPIƽ̨,���ƽ̨��Ϣ����ʧ�ܣ���ͬHIS����һ�����
    '����: In-lngPatiID ����ID,lngClinicID �Һ�ID
    '      Out-strErrMsg ������Ϣ����������ʧ����Ч
    '����:True-EMPIƽ̨������Ϣ�ɹ�,False-����ʧ��
    '����:���ϴ�
    '˵��:101170
    Dim blnCharge As Boolean, lngRet As Long
    If CreatePlugInOK(mlngModul) = False Then zlSaveEMPIPatiInfo = True: Exit Function
    If mbytMode <> 0 And mbytMode <> 2 Or mbytInState <> 0 Then zlSaveEMPIPatiInfo = True: Exit Function
    
    On Error GoTo Errhand
    If mobjfrmPatiInfo.mrsEMPIOut Is Nothing Then
        'EMPIû�в�����Ϣ����Ҫ�½�
        On Error Resume Next
        lngRet = gobjPlugIn.EMPI_AddPatiInfo(glngSys, mlngModul, lngPatiID, 0, lngClinicID, strErrMsg)
        Call zlPlugInErrH(Err, "EMPI_AddPatiInfo")
        If lngRet = 0 And Err.Number <> 438 Then Err.Clear: Exit Function
        Err.Clear: On Error GoTo Errhand
    Else
        '�ж�ƽ̨�ش�����Ϣ�Ƿ����ı�
        With mobjfrmPatiInfo.mrsEMPIOut
            If InStr(";" & mstrPrivs & ";", ";�����޸������;") > 0 And Exist�����(Nvl(!�����), lngPatiID) = False Then
                If txt�����.Text <> Nvl(!�����) Then blnCharge = True: GoTo EMPIModify
            End If
            If mobjfrmPatiInfo.txtPatiMCNO(0).Text <> Nvl(!ҽ����) Then blnCharge = True: GoTo EMPIModify
            If mbln������Ϣ���� Or blnNewPati Then
                If txtIDCard.Text <> Nvl(!���֤��) Then blnCharge = True: GoTo EMPIModify
                If txtPatient.Text <> Nvl(!����) Then blnCharge = True: GoTo EMPIModify
                If cbo�Ա�.ListIndex <> cbo.FindIndex(cbo�Ա�, Nvl(!�Ա�), True) Then blnCharge = True: GoTo EMPIModify
                If Format(txt��������.Text, "YYYY-MM-DD") <> Format(Nvl(!��������), "YYYY-MM-DD") Then blnCharge = True: GoTo EMPIModify
                If Format(txt����ʱ��.Text, "HH:MM") <> Format(Nvl(!��������), "HH:MM") Then blnCharge = True: GoTo EMPIModify
            End If
            If mobjfrmPatiInfo.txtBirthLocation.Text <> Nvl(!�����ص�) Then blnCharge = True: GoTo EMPIModify
            If mobjfrmPatiInfo.cbo����.ListIndex <> cbo.FindIndex(mobjfrmPatiInfo.cbo����, Nvl(!����), True) Then blnCharge = True: GoTo EMPIModify
            If mobjfrmPatiInfo.cbo����.ListIndex <> cbo.FindIndex(mobjfrmPatiInfo.cbo����, Nvl(!����), True) Then blnCharge = True: GoTo EMPIModify
            If mobjfrmPatiInfo.cboְҵ.ListIndex <> cbo.FindIndex(mobjfrmPatiInfo.cboְҵ, Nvl(!ְҵ)) Then blnCharge = True: GoTo EMPIModify
            If mobjfrmPatiInfo.txt��λ����.Text <> Nvl(!������λ) Then blnCharge = True: GoTo EMPIModify
            If mobjfrmPatiInfo.cbo����.ListIndex <> cbo.FindIndex(mobjfrmPatiInfo.cbo����, Nvl(!����״��), True) Then blnCharge = True: GoTo EMPIModify
            If txt��ͥ�绰.Text <> Nvl(!��ͥ�绰) Then blnCharge = True: GoTo EMPIModify
            If mobjfrmPatiInfo.txt��ϵ�˵绰.Text <> Nvl(!��ϵ�˵绰) Then blnCharge = True: GoTo EMPIModify
            If mobjfrmPatiInfo.txt��λ�绰.Text <> Nvl(!��λ�绰) Then blnCharge = True: GoTo EMPIModify
            If cbo��ͥ��ַ.Text <> Nvl(!��ͥ��ַ) Then blnCharge = True: GoTo EMPIModify
            If mobjfrmPatiInfo.txt��ͥ�ʱ�.Text <> Nvl(!��ͥ��ַ�ʱ�) Then blnCharge = True: GoTo EMPIModify
            If cbo���ڵ�ַ.Text <> Nvl(!���ڵ�ַ) Then blnCharge = True: GoTo EMPIModify
            If mobjfrmPatiInfo.txt���ڵ�ַ�ʱ�.Text <> Nvl(!���ڵ�ַ�ʱ�) Then blnCharge = True: GoTo EMPIModify
            If mobjfrmPatiInfo.txt��λ�ʱ�.Text <> Nvl(!��λ�ʱ�) Then blnCharge = True: GoTo EMPIModify
            If mobjfrmPatiInfo.txt��ϵ������.Text <> Nvl(!��ϵ������) Then blnCharge = True: GoTo EMPIModify
            If mobjfrmPatiInfo.cbo��ϵ�˹�ϵ.ListIndex <> cbo.FindIndex(mobjfrmPatiInfo.cbo��ϵ�˹�ϵ, Nvl(!��ϵ�˹�ϵ), True) Then blnCharge = True: GoTo EMPIModify
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
'���ܣ���ʾ/����Ԥ��֧����Ϣ
    Dim i As Integer
    Dim intIndex As Integer
    If mbln�����Һ� Then Exit Sub
    If gblnPrice Then blnShow = False
    stbThis.Panels(4).Visible = blnShow

    If Not blnShow Then
        mdblԤ����� = 0
        For i = 1 To vsfPay.Rows - 1
            If vsfPay.RowData(i) = 0 And vsfPay.TextMatrix(i, 0) <> "" Then
                vsfPay.TextMatrix(i, 6) = 0
            End If
        Next i
        stbThis.Panels(4).Text = "����Ԥ�����:0.00"
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
'���ܣ���ʾ/����ҽ�������ʻ�֧����Ϣ
    Dim i As Integer
    If gblnPrice Then blnShow = False
    stbThis.Panels(3).Visible = blnShow
    If Not blnShow Then
        mdbl������� = 0
        stbThis.Panels(3).Text = "0.00"
        For i = 1 To vsfPay.Rows - 1
            If vsfPay.RowData(i) = 3 And vsfPay.TextMatrix(i, 0) <> "" Then vsfPay.RowHidden(i) = True
        Next i
    Else
        If MCPAR.ʹ�ø����ʻ� Then
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
        '�Զ���ʱˢ��,�������ڹҺ�ʱ,��������ѡ�����ʱ
        If mcbrToolBar.Controls.Find(xtpControlButton, conMenu_View_Refresh).Enabled And mcbrToolBar.Controls.Find(xtpControlButton, conMenu_View_Refresh).Visible And txt�ű�.Text = "" And Not Me.ActiveControl Is vsfList Then RefreshFace
        mDatLastRefresh = Now
    End If
End Sub

Private Sub SetGridTop(intRow As Integer)
    Dim intRows As Integer
    intRows = vsfPlan.Height \ vsfPlan.RowHeight(1) - 2
    If vsfPlan.TopRow + intRows > intRow Then Exit Sub
    vsfPlan.TopRow = intRow
End Sub

Private Sub Load��ͥ��ַ()
    Dim strSQL As String, strFile As String
    Dim fld As Field, rsCheck As ADODB.Recordset
    Dim fso As Scripting.FileSystemObject
    Dim rsNew As ADODB.Recordset
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    strFile = App.Path & "\ZLAddressForRegEvent.Adtg"
    
    Set mrs��ͥ��ַ = New ADODB.Recordset
    
    On Error Resume Next
    If fso.FileExists(strFile) Then
        mrs��ͥ��ַ.Open strFile, "Provider=MSPersist", adOpenKeyset, adLockOptimistic, adCmdFile   '��Updateʱ������
    End If
    Err.Clear
    On Error GoTo errH
    
    If mrs��ͥ��ַ.State = 0 Then
        strSQL = "Select 'ϵͳ' As ���, ��ͥ��ַ As ����, Null As ����, 1 As ����" & vbNewLine & _
                "From ������Ϣ" & vbNewLine & _
                "Where 1 = 0" & vbNewLine & _
                "Union" & vbNewLine & _
                "Select 'ϵͳ' As ���, ����, ����, 1 As ���� From ����"

        Call zlDatabase.OpenRecordset(mrs��ͥ��ַ, strSQL, Me.Caption)            '������adUseClient���ܽ�����
        
        If Not mrs��ͥ��ַ.EOF Then
            '��������:����,����
            Set fld = mrs��ͥ��ַ.Fields(1)
            fld.Properties("Optimize") = True
            Set fld = mrs��ͥ��ַ.Fields(2)
            fld.Properties("Optimize") = True
            
            If fso.FileExists(strFile) Then
                Kill strFile
            End If
            mrs��ͥ��ַ.Save strFile, adPersistADTG
        End If
        mrs��ͥ��ַ.Close
        mrs��ͥ��ַ.Open strFile, "Provider=MSPersist", adOpenKeyset, adLockOptimistic, adCmdFile   '��Updateʱ������
    Else
        strSQL = "Select 'ϵͳ' As ���, ��ͥ��ַ As ����, Null As ����, 1 As ����" & vbNewLine & _
                "From ������Ϣ" & vbNewLine & _
                "Where 1 = 0" & vbNewLine & _
                "Union" & vbNewLine & _
                "Select 'ϵͳ' As ���, ����, ����, 1 As ���� From ���� Where 1 = 0"
        Set rsCheck = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        If rsCheck.Fields(1).DefinedSize > mrs��ͥ��ַ.Fields(1).DefinedSize Or rsCheck.Fields(2).DefinedSize > mrs��ͥ��ַ.Fields(2).DefinedSize Then
            If fso.FileExists(strFile) Then
                Kill strFile
            End If
            strSQL = "Select 'ϵͳ' As ���, ��ͥ��ַ As ����, Null As ����, 1 As ����" & vbNewLine & _
                    "From ������Ϣ" & vbNewLine & _
                    "Where 1 = 0" & vbNewLine & _
                    "Union" & vbNewLine & _
                    "Select 'ϵͳ' As ���, ����, ����, 1 As ���� From ����"
            Set rsNew = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
            rsNew.Save strFile, adPersistXML
            mrs��ͥ��ַ.Close
            mrs��ͥ��ַ.Open strFile, "Provider=MSPersist", adOpenKeyset, adLockOptimistic, adCmdFile   '��Updateʱ������
        End If
    End If
    
    lbl��ͥ��ַ.ToolTipText = "�붨�ڱ��ݱ���[��ͥ��ַ]�����ļ�:" & strFile
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub cbo��ͥ��ַ_GotFocus()
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub cbo��ͥ��ַ_LostFocus()
    Call zlCommFun.OpenIme
End Sub

Private Sub cbo��ͥ��ַ_KeyDown(KeyCode As Integer, Shift As Integer)
    '�˹��̴������������ݵ�ɾ��,�Լ���������ʱ���������б�
    '�����б���ʱ,�������ɾ����ʱ,��ɾ�������¼
    
    Dim str��ͥ��ַ As String
    
    If KeyCode = vbKeyDelete Then
        str��ͥ��ַ = cbo��ͥ��ַ.Text
        
        If Not mrs��ͥ��ַ Is Nothing And mTy_Para.byt��ͥ��ַ���� = 1 Then
            If mrs��ͥ��ַ.State = 1 And str��ͥ��ַ <> "" Then
                If cbo��ͥ��ַ.SelText = str��ͥ��ַ And SendMessage(cbo��ͥ��ַ.Hwnd, CB_GETDROPPEDSTATE, 0, 0) = True Then
                    mrs��ͥ��ַ.Filter = "����='" & str��ͥ��ַ & "'"
                    If Not mrs��ͥ��ַ.EOF Then
                        mrs��ͥ��ַ.Delete adAffectCurrent
                        mrs��ͥ��ַ.Update
                    End If
                End If
            End If
        End If
    ElseIf KeyCode = vbKeyDown And cbo��ͥ��ַ.Text <> "" Then
        If SendMessage(cbo��ͥ��ַ.Hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 Then Call SendMessage(cbo��ͥ��ַ.Hwnd, CB_SHOWDROPDOWN, True, 0&)
    End If
End Sub

Private Sub cbo��ͥ��ַ_KeyUp(KeyCode As Integer, Shift As Integer)
    '��ʱtext���ѽ����������Ϣ
    '���¼�����ɾ�����˸��,ɾ������������Ŀ��,�����б�����������Ӧ������ɸѡ
    '���ȫ�����ֶ�ɾ����,����������б�����
        
    Dim str��ͥ��ַ As String, i As Long
    Dim lngλ�� As Long
    
    If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
        If mrs��ͥ��ַ Is Nothing Then Exit Sub
        If mTy_Para.byt��ͥ��ַ���� = 0 Then Exit Sub
        
        str��ͥ��ַ = cbo��ͥ��ַ.Text                      '��ʱ,���ѡ���˲�������,��ѡ��������Ѿ���ɾ��
        lngλ�� = cbo��ͥ��ַ.SelStart
        
        If mrs��ͥ��ַ.State = 1 And Len(str��ͥ��ַ) > 1 Then
            
            If InStr(1, "ABCDEFGHIJKLMNOPQRSTUVWXYZ", UCase(Left(str��ͥ��ַ, 1))) > 0 Then
                mrs��ͥ��ַ.Filter = "���� like '" & gstrLike & UCase(str��ͥ��ַ) & "*'"
            Else
                mrs��ͥ��ַ.Filter = "���� Like '" & gstrLike & str��ͥ��ַ & "*'"
            End If
            
            If Not mrs��ͥ��ַ.EOF Then
                
                If mrs��ͥ��ַ.RecordCount <> cbo��ͥ��ַ.ListCount Then
                    Call SendMessage(cbo��ͥ��ַ.Hwnd, CB_RESETCONTENT, 0, 0)
                    mrs��ͥ��ַ.Sort = "���� Desc,����"
                    For i = 1 To mrs��ͥ��ַ.RecordCount
                        AddComboItem cbo��ͥ��ַ.Hwnd, CB_ADDSTRING, 0, mrs��ͥ��ַ!����
                        mrs��ͥ��ַ.MoveNext
                    Next
                    If SendMessage(cbo��ͥ��ַ.Hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 Then Call SendMessage(cbo��ͥ��ַ.Hwnd, CB_SHOWDROPDOWN, True, 0&)
                                        
                    cbo��ͥ��ַ.Text = str��ͥ��ַ
                    cbo��ͥ��ַ.SelStart = lngλ��
                End If
            Else
                Call SendMessage(cbo��ͥ��ַ.Hwnd, CB_SHOWDROPDOWN, False, 0&)
            End If
        ElseIf str��ͥ��ַ = "" Then
            cbo��ͥ��ַ.Clear
            Call SendMessage(cbo��ͥ��ַ.Hwnd, CB_SHOWDROPDOWN, False, 0&)
        End If
    End If
End Sub

Private Sub cbo��ͥ��ַ_KeyPress(KeyAscii As Integer)
    Dim i As Long
    Dim str���� As String
    Dim str��ͥ��ַ As String
    Dim lng�м������ As Long
    Dim strTemp As String
    
    If mrs��ͥ��ַ Is Nothing Then Exit Sub
    
    If mTy_Para.byt��ͥ��ַ���� = 0 Then
        If KeyAscii = 13 Then Call zlCommFun.PressKey(vbKeyTab)
        Exit Sub
    End If
    '�ñ��ػ���ƥ������
    If KeyAscii <> 13 And KeyAscii <> vbKeyF4 And KeyAscii <> vbKeyEscape And _
        KeyAscii <> vbKeyBack And KeyAscii <> 26 And KeyAscii <> 3 And KeyAscii <> 22 Then   '26��ʾctrl+z,3-ctrl+c,22-ctrl+v
            
        If mrs��ͥ��ַ.State = 0 Or cbo��ͥ��ַ.Text = "" Then  '���һ����ʱ��ƥ��
            Exit Sub
        End If
       
        'ѡ���м䲿���ı�����������
        If cbo��ͥ��ַ.SelText <> "" And (cbo��ͥ��ַ.SelStart + cbo��ͥ��ַ.SelLength) <> Len(cbo��ͥ��ַ.Text) Then
            lng�м������ = cbo��ͥ��ַ.SelStart + 1
            cbo��ͥ��ַ.Text = Mid(cbo��ͥ��ַ.Text, 1, cbo��ͥ��ַ.SelStart) & Chr(KeyAscii) & Mid(cbo��ͥ��ַ.Text, cbo��ͥ��ַ.SelStart + cbo��ͥ��ַ.SelLength + 1)
            cbo��ͥ��ַ.SelText = ""
            str��ͥ��ַ = cbo��ͥ��ַ.Text
        Else
            '�������β��,�����м�ʱ,�������ѡ��
            If cbo��ͥ��ַ.SelStart = Len(cbo��ͥ��ַ.Text) Or (cbo��ͥ��ַ.SelStart + cbo��ͥ��ַ.SelLength) = Len(cbo��ͥ��ַ.Text) Then
                str��ͥ��ַ = Mid(cbo��ͥ��ַ.Text, 1, cbo��ͥ��ַ.SelStart) & Chr(KeyAscii)
            Else
                str��ͥ��ַ = Mid(cbo��ͥ��ַ.Text, 1, cbo��ͥ��ַ.SelStart) & Chr(KeyAscii) & Mid(cbo��ͥ��ַ.Text, cbo��ͥ��ַ.SelStart + 1)
                lng�м������ = cbo��ͥ��ַ.SelStart + 1
            End If
        End If
         
        
        If InStr(1, "ABCDEFGHIJKLMNOPQRSTUVWXYZ", UCase(Left(str��ͥ��ַ, 1))) > 0 Then
            mrs��ͥ��ַ.Filter = "���� like '" & gstrLike & UCase(str��ͥ��ַ) & "*'"
        Else
            mrs��ͥ��ַ.Filter = "���� Like '" & gstrLike & str��ͥ��ַ & "*'"
        End If
        
        If Not mrs��ͥ��ַ.EOF Then
            If mrs��ͥ��ַ.RecordCount <> cbo��ͥ��ַ.ListCount Then
                Call SendMessage(cbo��ͥ��ַ.Hwnd, CB_RESETCONTENT, 0, 0)
                mrs��ͥ��ַ.Sort = "���� Desc,����"
                For i = 1 To mrs��ͥ��ַ.RecordCount
                    AddComboItem cbo��ͥ��ַ.Hwnd, CB_ADDSTRING, 0, mrs��ͥ��ַ!����
                    mrs��ͥ��ַ.MoveNext
                Next
                If SendMessage(cbo��ͥ��ַ.Hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 Then Call SendMessage(cbo��ͥ��ַ.Hwnd, CB_SHOWDROPDOWN, True, 0&)
            End If
            
            i = KeyAscii    '���������ж��Ƿ��ǰ��˸�ɾ����
            KeyAscii = 0
            cbo��ͥ��ַ.Text = str��ͥ��ַ
            cbo��ͥ��ַ.SelStart = Len(cbo��ͥ��ַ.Text)

            mrs��ͥ��ַ.MoveFirst   '�����������ļ���,��ͬ��ȡ��һ�������
            If mrs��ͥ��ַ!���� = str��ͥ��ַ And i <> vbKeyBack Then
                mrs��ͥ��ַ.MoveNext
            End If
            If Not mrs��ͥ��ַ.EOF Then
                If InStr(1, mrs��ͥ��ַ!����, str��ͥ��ַ) > 0 Or mrs��ͥ��ַ!���� = UCase(str��ͥ��ַ) Then    '�������������������ݵ�һ����,��ѡ�л����������
                    i = Len(cbo��ͥ��ַ.Text)
                    strTemp = cbo��ͥ��ַ.Text
                    cbo��ͥ��ַ.Text = mrs��ͥ��ַ!����
                    If InStr(1, mrs��ͥ��ַ!����, str��ͥ��ַ) > 0 Then '����:31570
                        i = InStr(1, cbo��ͥ��ַ.Text, strTemp) + Len(strTemp) - 1
                    End If
                    cbo��ͥ��ַ.SelStart = i
                    cbo��ͥ��ַ.SelLength = Len(cbo��ͥ��ַ.Text) - cbo��ͥ��ַ.SelStart
                    If mrs��ͥ��ַ.RecordCount = 1 Then Exit Sub
                End If
            End If
            
        'û���ҵ�ƥ��Ļ�������ʱ,����������б�����
        Else
            Call SendMessage(cbo��ͥ��ַ.Hwnd, CB_RESETCONTENT, 0, 0)
            If SendMessage(cbo��ͥ��ַ.Hwnd, CB_GETDROPPEDSTATE, 0, 0) = 1 Then Call SendMessage(cbo��ͥ��ַ.Hwnd, CB_SHOWDROPDOWN, False, 0&)
            KeyAscii = 0
            cbo��ͥ��ַ.Text = str��ͥ��ַ
            cbo��ͥ��ַ.SelStart = Len(cbo��ͥ��ַ.Text)
        End If
        
        If lng�м������ > 0 Then cbo��ͥ��ַ.SelStart = lng�м������: cbo��ͥ��ַ.SelText = ""
        
    ElseIf KeyAscii = 13 Then
        'a.��û��ѡ���κ�����,����������Ϊ��,���Ϊ��ĩ��ʱ,ȷ������,��������Ϣ�����ػ���
        Call SendMessage(cbo��ͥ��ַ.Hwnd, CB_SHOWDROPDOWN, False, 0&)
        
        If cbo��ͥ��ַ.Text = "" Then
            If gbln��ͥ��ַ And txtPatient.Text <> "" Then
                Exit Sub
            Else
                Call zlCommFun.PressKey(vbKeyTab): Exit Sub
            End If
        End If
        
        '�����б���ʱ���س�,��λ��ĩβ
        If cbo��ͥ��ַ.SelText = cbo��ͥ��ַ.Text Then
            cbo��ͥ��ַ.SelStart = Len(cbo��ͥ��ַ.Text):
            Exit Sub
       End If
        If mrs��ͥ��ַ.State = 0 Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
        If zlCommFun.ActualLen(cbo��ͥ��ַ.Text) > 100 Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
        'a.������״̬�°��س�,û��ѡ���ı�
        If cbo��ͥ��ַ.SelText = "" Then
            str��ͥ��ַ = cbo��ͥ��ַ.Text
            mrs��ͥ��ַ.Filter = "����='" & str��ͥ��ַ & "'"
            If mrs��ͥ��ַ.EOF Then
                str���� = Mid(zlCommFun.zlGetSymbol(str��ͥ��ַ), 1, 10)
                If str���� <> UCase(str��ͥ��ַ) Then
                    With mrs��ͥ��ַ
                        .AddNew
                        !��� = "�û�"
                        !���� = str��ͥ��ַ
                        !���� = str����
                        !���� = 1
                        .Update                 '�ڴ���Unload��save
                    End With
                End If
            Else
                mrs��ͥ��ַ!���� = mrs��ͥ��ַ!���� + 1
                mrs��ͥ��ַ.Update
                
                If zlCommFun.IsCharAlpha(str��ͥ��ַ) Then
                    If mrs��ͥ��ַ.RecordCount = 1 Then
                        cbo��ͥ��ַ.Text = mrs��ͥ��ַ!����
                    Else
                        Call SendMessage(cbo��ͥ��ַ.Hwnd, CB_SHOWDROPDOWN, True, 0&)
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
 
    strSQL = "Select 1 From ������� Where ���=1 And ���=[1]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strMCCode)

    CheckMCOutMode = rsTmp.RecordCount > 0
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Sub Init���㷽ʽ(ByVal str���� As String, Optional ByVal objCards As Cards)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ�����㷽ʽ
    '���:str����-���㷽ʽ������,����ö��ŷ���
    '                   1-�ֽ���㷽ʽ,2-������ҽ������,
    '                   3-ҽ�������ʻ�,4-ҽ������ͳ��,5-���տ���,6-�����ۿ�,
    '                   7-һ��ͨ����,8-���㿨����)
    '����:objCards-����صĽ��㷽ʽ���ظ�������
    '����:���˺�
    '����:2013-10-24 10:41:51
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, i As Long, objCard As Card
    Dim rsTmp As ADODB.Recordset
    If str���� = "" Then
        str���� = ",1,2,3,4,5,6,7,8,"
    Else
        str���� = "," & str���� & ","
    End If
    
    strSQL = _
        "Select B.����,B.����,Nvl(B.����,1) as ����,Nvl(A.ȱʡ��־,0) as ȱʡ" & _
        " From ���㷽ʽӦ�� A,���㷽ʽ B" & _
        " Where A.Ӧ�ó���=[1] And B.����=A.���㷽ʽ And Instr([2] ,','||B.����||',')>0" & _
        " Order by B.����"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, "�Һ�", str����)
    
    cbo���㷽ʽ.Clear
    Do While Not rsTmp.EOF
        If Not objCards Is Nothing Then
            Set objCard = New Card
            With objCard
                .�ӿ���� = 0
                .���� = Nvl(rsTmp!����)
                .���㷽ʽ = Nvl(rsTmp!����)
                .�ӿڱ��� = Val(Nvl(rsTmp!����))
                .���� = False
            End With
            objCards.Add objCard
        End If
        cbo���㷽ʽ.AddItem rsTmp!���� & "-" & rsTmp!����
        If rsTmp!���� = gstr���㷽ʽ Then
            For i = 0 To cbo���㷽ʽ.ListCount - 1
                cbo���㷽ʽ.ItemData(i) = 0
            Next
            cbo���㷽ʽ.ItemData(cbo���㷽ʽ.NewIndex) = 1
            cbo���㷽ʽ.ListIndex = cbo���㷽ʽ.NewIndex
        End If
        
        If rsTmp!ȱʡ = 1 Then
            If cbo���㷽ʽ.ListIndex = -1 Then
                cbo���㷽ʽ.ItemData(cbo���㷽ʽ.NewIndex) = 1
                cbo���㷽ʽ.ListIndex = cbo���㷽ʽ.NewIndex
            End If
        End If
        rsTmp.MoveNext
    Loop
    
    If cbo���㷽ʽ.ListCount > 0 And cbo���㷽ʽ.ListIndex = -1 Then
        cbo���㷽ʽ.ListIndex = 0
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub InitFace()
    '����:��ʼ������ؼ�
    Dim i As Long, Control As Object
    
    '68991
    mRegistFeeMode = EM_RG_����
    mPatiChargeMode = EM_�Ƚ��������
    
    lblPrompt.Caption = ""
    vsfPay.Height = 2220
    Call ClearMoney
    
    
    If mTy_Para.bln�����ͷ���� Then
       vsfPlan.ExplorerBar = flexExSortShow
    Else
       vsfPlan.ExplorerBar = flexExNone
    End If
    If mbytInState = 0 Then
        Call InitInputMaxLen
        If mbytMode = 0 And Not mblnStation Then
            chkShowAll.Visible = True
        End If
        
        If InStr(mstrPrivs, ";�ش�Ʊ��;") = 0 Then
            chkPrint.Visible = False
        End If
        If InStr(";" & mstrPrepayPrivs & ";", ";����Ԥ��;") = 0 Then
            mcbrToolBar.Controls.Find(xtpControlButton, 3816).Visible = False
            mcbrToolBar.Controls.Find(xtpControlButton, 3816).Enabled = False
        End If
        'Ȩ���޸� ���⣺37798 ���ߣ�Ƚ��
        If InStr(mstrPrivs, ";ԤԼ�Һ�;") = 0 Then chkBooking.Visible = False
        
        lblFree.Left = lblCancel.Left: lblFree.Height = lblCancel.Height
        lblFree.Visible = False
        
        txtFact.Locked = Not (InStr(1, mstrPrivs, ";�޸�Ʊ�ݺ�;") > 0) And gblnBill�Һ�  '���˺�:20000,�����޸�Ʊ�ݺ�Ȩ��
        timPlan.Enabled = glngInterval > 0 And Not mblnStation And (mbytMode = 0 Or mbytMode = 1)
        If timPlan.Enabled Then mDatLastRefresh = Now
    
        Call SetPatiInfoEnabled(False, mrsInfo Is Nothing)  '�����:58843
        
        cboҽ��.Enabled = False
        cbo�ѱ�.Enabled = (gbln�ѱ� Or mblnStation) And mbytMode <> 2
        cbo���㷽ʽ.Enabled = gbln���㷽ʽ And mbytMode <> 1
        txt��ͥ�绰.Enabled = gbln�绰
        lblIDCard.Visible = True
        If IDKind֤��.IDKind = IDKind֤��.GetKindIndex("���֤��") Then
            txtIDCard.Visible = True: txt֤��.Visible = False
        Else
            txtIDCard.Visible = False: txt֤��.Visible = True
        End If
        Call SetPicTimeObjectVisible
        
          If mbytMode = 1 Then
            'ԤԼ�Һ�
            chkPrint.Visible = False: chkCancel.Visible = False: chkBooking.Visible = False
            cboNO.Width = cboҽ��.Width
            cmdPatiPic.Left = chkBooking.Left
            txtSN.Width = txtSN.Width + cmdPatiPic.Width + 60
            '����:26964
            chkShowAll.Visible = Not mblnStation: mblnUnChkClick = True
            If Val(zlDatabase.GetPara("ԤԼ��ʾ���кű�", glngSys, mlngModul, 1, Array(chkShowAll), InStr(mstrPrivs, ";��������;") > 0)) = 1 Then
                chkShowAll.Value = 1
            Else
                chkShowAll.Value = 0
            End If
            mblnUnChkClick = False

            picBookingDate.Visible = True
            lblժҪ.Visible = True: cbo��ע.Visible = True
            lblԤԼ��ʽ.Visible = True: cboԤԼ��ʽ.Visible = True
            '-----------------------------------------------------------------------------------------
            vsfPay.Visible = False
            lblӦ��.Visible = False
            lblIDCard.Visible = True
            If IDKind֤��.IDKind = IDKind֤��.GetKindIndex("���֤��") Then
                txtIDCard.Visible = True: txt֤��.Visible = False
            Else
                txtIDCard.Visible = False: txt֤��.Visible = True
            End If
            txt��ͥ�绰.Visible = True: lbl��ͥ�绰.Visible = True
            cmdCard.Visible = False: cmdYb.Visible = False
            '-----------------------------------------------------------------------------------------
            Call SetUndisplayBalance
        ElseIf mbytMode = 2 Then
            '����ԤԼ
            '���غű��Ų���(��Ҫ������д����)
            vsfPlan.Visible = False: vsfList.Visible = False
            picSplit.Visible = False
            cmdCard.Visible = InStr(1, mstrPrivs, ";�󶨿���;") > 0   '�󶨿���:31182
            cmdYb.Visible = True   'ԤԼ����ʱ,����ˢҽ�� '����:31182
            
            lblժҪ.Visible = True: cbo��ע.Visible = True
            cbo��ע.Enabled = False: cboԤԼ��ʽ.Enabled = False
            lblԤԼ��ʽ.Visible = True: cboԤԼ��ʽ.Visible = True
'            mcbrToolBar.Visible = False
            dkpMain.DestroyPane dkpMain.Panes(1)
            Call SetReceiveState(True)
            Me.Width = glngMinW: Me.Height = glngMinH
            Me.WindowState = 0
        Else
            '�����Һ�
            If InStr(mstrPrivs, ";�˺�;") = 0 Then
                chkCancel.Visible = False
                lblNO.Left = lblNO.Left + chkCancel.Width
                cboNO.Left = cboNO.Left + chkCancel.Width
            End If
            cmdYb.Visible = True
            picBookingDate.Visible = False
        End If
        
        '��ʼ�����״̬���
        vsfList.Cols = SNCOLS
        For i = 0 To vsfList.Cols - 1
            vsfList.ColWidth(i) = 570
            vsfList.ColAlignment(i) = 4
        Next
        vsfList.RowHeightMin = 500
        
        'ȡ���ű�
        Call SetPlanGrid
    
    Else
        If mbytMode = 1 Then '�鿴ԤԼ��ʱ�޽��������Ϣ
            lblժҪ.Visible = True: cbo��ע.Visible = True
            lblԤԼ��ʽ.Visible = True: cboԤԼ��ʽ.Visible = True
            Call SetUndisplayBalance
            lblIDCard.Visible = True:  IDKind֤��.Visible = True
            If IDKind֤��.IDKind = IDKind֤��.GetKindIndex("���֤��") Then
                txtIDCard.Visible = True: txt֤��.Visible = False
            Else
                txtIDCard.Visible = False: txt֤��.Visible = True
            End If
            txt��ͥ�绰.Visible = True: lbl��ͥ�绰.Visible = True
            cmdCard.Visible = False: cmdYb.Visible = False
            If mbytInState = 1 And (mbytMode = 1 Or mbytMode = 3) Then
                lbl����ʱ��.Visible = True: txt����ʱ��.Visible = True
            End If
            vsfPay.Visible = False
            stbThis.Visible = False
        Else
            lbl����ʱ��.Visible = True: txt����ʱ��.Visible = True
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
        chk������.Enabled = False
        cboԤԼ��ʽ.Enabled = False
        cbo��ע.Enabled = False
        txt����ʱ��.Enabled = False
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
        lbl�ɿ�.Visible = False: txt�ɿ�.Visible = False
        lbl�Ҳ�.Visible = False: txt�Ҳ�.Visible = False
        txt����Ӧ��.Visible = False
        lblӦ��.Visible = False
        If mbytMode <> 0 Then
            lblSum.Visible = False: txt�ϼ�.Visible = False
            picTotal.Visible = True
        End If
        Call SetUndisplayBalance
        
        If Not (Me.mbytInState = 1 And (mbytMode = 3 Or mbytMode = 4)) Then
            cmdCancel.Caption = "�˳�(&X)"
            
        Else
            If mbytMode = 3 Then
                cmdOK.Visible = True
                vsfPay.Visible = False
            End If
        End If
        
        If mbytMode = 4 Then
            '�����˺�ʱ , ��ؿؼ�������
'            chk������.Enabled = True
'            Set chk������.Container = Me
'            chk������.Top = chk������.Top + 480
'            chk������.Caption = "�˲�����"
'            chkExtra.Enabled = True
'            Set chkExtra.Container = Me
'            chkExtra.Top = chk������.Top
'            Set cbo��ע.Container = Me
'            cbo��ע.Top = chk������.Top + 450
            lblԤԼ��ʽ.Visible = False
            cboԤԼ��ʽ.Visible = False
            cbo�ѱ�.Enabled = False
            cbo���㷽ʽ.Enabled = False
            cbo���㷽ʽ.Visible = False
            vsfMoney.Enabled = False
            cmdCancel.Visible = True
            cmdOK.Visible = True
            cmdOK.Top = 900
            cmdCancel.Top = 900
        Else
            cbo���㷽ʽ.Visible = False
            If lblӦ��.Visible Then
                cmdCancel.Top = lblӦ��.Top + lblӦ��.Height + 180
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
            chkExtra.Caption = "�˸��ӷ�"
        Else
            chkExtra.Caption = "���ӷ�"
        End If
    End If
      
    Call Set��עEnabled
    lbl����.Visible = False
    '74430,Ƚ����,2014-7-8,�ҺŽ�����ʾ������Ƭ�ĸ�������
    picPatiPicBack.Left = Me.ScaleWidth - picPatiPicBack.Width
    picPatiPicBack.Top = 0
    picPatiPicBack.Visible = False: cmdPatiPic.Enabled = False
    
    If mbytMode <> 0 And mbytMode <> 1 And mbytMode <> 2 Then cmdPatiPic.Visible = False
'    If mbytMode = 1 Or mbytMode = 2 Then cmdPatiPic.Left = picCode.Width - 300
 
    If mblnStructAdress Then
        padd��ͥ��ַ.Visible = True: padd���ڵ�ַ.Visible = True
        padd��ͥ��ַ.ShowTown = mblnShowTown: padd���ڵ�ַ.ShowTown = mblnShowTown
        cbo��ͥ��ַ.Visible = False: padd��ͥ��ַ.MaxLength = glngMax��ͥ��ַ
        
        padd��ͥ��ַ.Top = cbo��ͥ��ַ.Top: padd��ͥ��ַ.Left = cbo��ͥ��ַ.Left
        lbl��ͥ��ַ.Top = padd��ͥ��ַ.Top
        
        cbo���ڵ�ַ.Visible = False: padd���ڵ�ַ.MaxLength = glngMax���ڵ�ַ
        padd���ڵ�ַ.Top = padd��ͥ��ַ.Top + padd��ͥ��ַ.Height + 20: padd���ڵ�ַ.Left = cbo���ڵ�ַ.Left
        lbl���ڵ�ַ.Top = padd���ڵ�ַ.Top
        picDetailFee.Top = padd���ڵ�ַ.Top + padd���ڵ�ַ.Height + 50
    Else
        lbl��ͥ��ַ.Top = cbo��ͥ��ַ.Top + (cbo��ͥ��ַ.Height - lbl��ͥ��ַ.Height) \ 2
        lbl���ڵ�ַ.Top = cbo���ڵ�ַ.Top + (cbo���ڵ�ַ.Height - lbl���ڵ�ַ.Height) \ 2
        picDetailFee.Top = cbo���ڵ�ַ.Top + cbo���ڵ�ַ.Height + 50
    End If
    
End Sub

Private Sub Set��עEnabled()
    '--------------------------
    '��ע�ؼ���λ���Լ������Եĵ���
    '�Һ�,�˺�ʱ ��Ҫ������С�Լ�λ��
    '--------------------------
   Dim Control As Object
   Me.cbo��ע.Visible = mbytInState <= 0
   Me.lblժҪ.Visible = True
   If mbytInState <= 0 Or (mbytInState = 1 And (mbytMode = 3 Or mbytMode = 4)) Then
        'ִ�� ������ԤԼʱ
        Me.cbo��ע.Visible = True
        Me.cboԤԼ��ʽ.Enabled = IIf(mbytInState = 1 And mbytMode = 3 Or mbytMode = 4, False, True)
        Me.cbo��ע.Enabled = IIf(mbytInState = 1, False, True)
        Me.cbo��ע.Visible = True
   Else
        Me.cbo��ע.Visible = True: Me.cbo��ע.Enabled = IIf(mbytInState = 1, False, True)
   End If
 
   If (mbytMode = 4 Or mbytMode = 3) And mbytInState = 1 Then
        Me.cmdOK.Visible = True: Me.cmdOK.Enabled = True
        Me.cboNO.Locked = True: Me.cboNO.TabStop = False
        Me.cmdCancel.TabIndex = Me.cmdOK.TabIndex - 1
  End If
End Sub
Private Sub zlInitParaSet()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����������
    '����:���˺�
    '����:2009-12-25 11:27:09
    '����:26962
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTmp          As String
    Dim lngTmp          As Long
    Err = 0: On Error GoTo Errhand:
    If mblnStation Then zlDatabase.ClearParaCache    'ҽ��վʱ ��ȡ���� ���ӻ����ж�ȡ���������޸Ĳ���������Ч
    strTmp = zlDatabase.GetPara("ԤԼ����ʱ��", glngSys, mlngModul, "1|60")
    With mTy_Para
        .bln�Һ����ɶ��� = Val(zlDatabase.GetPara("�Ŷӽк�ģʽ", glngSys, 1113)) <> 0 And mblnStation = False
        '����:31182
        .intͬ����Լ�� = Val(zlDatabase.GetPara("����ͬ����ԼN����", glngSys, mlngModul, 0))
        .intͬ���޹��� = Val(Split(zlDatabase.GetPara("����ͬ���޹�N����", glngSys, mlngModul, 0) & "|", "|")(0))
        .blnͬ���޹Ҽ��� = Split(zlDatabase.GetPara("����ͬ���޹�N����", glngSys, mlngModul, 0) & "|", "|")(1) = "1"
        .int���˹Һſ����� = Val(zlDatabase.GetPara("���˹Һſ�������", glngSys, mlngModul, 0))
        .int����ԤԼ������ = Val(zlDatabase.GetPara("����ԤԼ������", glngSys, mlngModul, 0))
        .lngԤԼ��Чʱ�� = Val(zlDatabase.GetPara("ԤԼ��Чʱ��", glngSys, mlngModul, 0))
        .intԤԼʧЧ���� = Val(zlDatabase.GetPara("ԤԼʧԼ����", glngSys, mlngModul, 0))
        .blnԤԼ����ȷ���Һŷ� = zlDatabase.GetPara("ԤԼ����ȷ���Һŷ�", glngSys, mlngModul, 0) = "1"
        .bln����סԺ���˹Һ� = zlDatabase.GetPara("����סԺ���˹Һ�", glngSys, mlngModul, 0) = "1"
        .blnԤԼ����������� = Val(zlDatabase.GetPara("ԤԼ�����������", glngSys, mlngModul, 0)) = 1   '36028
        .bln�����ͷ���� = Val(zlDatabase.GetPara("������ͷ����", glngSys, mlngModul, 0)) = 1   '43847
        .bln������ѡ�� = Val(zlDatabase.GetPara("������ѡ��", glngSys, mlngModul, 0)) = 1   '43847
        .blnʧԼ���ڹҺ� = Val(zlDatabase.GetPara("ʧԼ���ڹҺ�", glngSys, mlngModul, 0)) = 1
        .bln�˺���� = Val(zlDatabase.GetPara("�˺����", glngSys, mlngModul, 0)) = 1
        .lngN��ȡ��ԤԼ = Val(zlDatabase.GetPara("N���ڲ���ȡ��ԤԼ��", glngSys, mlngModul, 0))
        .lngԤԼ����ʱ�� = Val(Split(strTmp, "|")(1))
        .lngԤԼȱʡ���� = Val(Split(strTmp, "|")(0))
          '����Ϊ����ҽ������վ�Ĳ�������������
        .bln�Һű���ˢ�� = Val(zlDatabase.GetPara("�Һű���ˢ��", glngSys, 1260, 0)) = 1     '38603
        .byt��ͥ��ַ���� = Val(Nvl(zlDatabase.GetPara("��ͥ��ַ���뷽ʽ", glngSys, mlngModul, 1)))
        lngTmp = Val(zlDatabase.GetPara("N�����±���¼��໤��", glngSys, mlngModul, 0))
        .bln�໤��¼�� = lngTmp > 0
        .lngN������¼��໤�� = lngTmp
        .bln�ϸ�ʱ�ιҺ� = Val(zlDatabase.GetPara("�ϸ�ʱ�ιҺ�", glngSys, mlngModul, 0)) = 1   '62467
        .blnReuseCancelNO = Val(zlDatabase.GetPara("�����������Һ�", glngSys, mlngModul, 1)) = 1
        .intר�ҺŹҺ����� = Val(zlDatabase.GetPara("ר�ҺŹҺ�����", glngSys, , 0))
        .intר�Һ�ԤԼ���� = Val(zlDatabase.GetPara("ר�Һ�ԤԼ����", glngSys, , 0))
        .bln��ֹ�������� = Val(zlDatabase.GetPara("��ֹ��������", glngSys, mlngModul, 0)) = 1
        .byt�ɿʽ = Val(zlDatabase.GetPara("�ҺŽɿ��������", glngSys, mlngModul, 0))
        .byt����ģʽ = Val(zlDatabase.GetPara("ԤԼ����ģʽ", glngSys, mlngModul, 0))
    End With
    If mTy_Para.lngԤԼ����ʱ�� <= 0 Then mTy_Para.lngԤԼ����ʱ�� = 60
    mblnCheckNOValidity = Val(Nvl(zlDatabase.GetPara("�������Ч�Լ��", glngSys, mlngModul, 1), 1)) = 1
    mSortType = Val(zlDatabase.GetPara("ȱʡ����ʽ", glngSys, mlngModul, 0))
    mblnStructAdress = Val(zlDatabase.GetPara(251, glngSys)) <> 0 '���˵�ַ�ṹ��¼��
    mblnShowTown = Val(zlDatabase.GetPara(252, glngSys)) <> 0 '�����ַ�ṹ��¼��
    Exit Sub
Errhand:
    If ErrCenter = 1 Then Resume
End Sub

Private Function zlGet��ǰ���ڼ�(Optional strDate As String = "") As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ���������ڼ�
    '����:���˺�
    '����:2010-02-04 14:42:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset, bln��ǰ���� As Boolean, strTemp As String
    bln��ǰ���� = False
    If strDate = "" Then
        bln��ǰ���� = True
        If mstr��ǰ���� <> "" Then zlGet��ǰ���ڼ� = mstr��ǰ����: Exit Function
        strSQL = "Select Decode(To_Char(Sysdate,'D'),'1','��','2','һ','3','��','4','��','5','��','6','��','7','��',NULL) as ����  From dual"
        strDate = "1990-01-01"
    Else
        strSQL = "Select Decode(To_Char([1],'D'),'1','��','2','һ','3','��','4','��','5','��','6','��','7','��','') As ���� From dual"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CDate(strDate))
    If rsTemp.EOF = True Then
        Exit Function
    End If
    strTemp = Nvl(rsTemp!����)
    If bln��ǰ���� Then mstr��ǰ���� = strTemp
    zlGet��ǰ���ڼ� = strTemp
End Function

Private Sub InitData()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long, strTemp As String
    Dim Curdate As Date, arrTmp As Variant
    
    '��ʼ��������
     On Error GoTo errH
    
    If mbytInState = 0 Then
        Call GetRegInFor(g˽��ģ��, Me.Name, "idkind", strTemp)
        mintIDKind = Val(strTemp)
    End If
    If mintIDKind > 0 And mintIDKind <= IDKind.ListCount Then IDKind.IDKind = mintIDKind
    
    cbo���䵥λ.AddItem "��"
    cbo���䵥λ.AddItem "��"
    cbo���䵥λ.AddItem "��"
    cbo���䵥λ.ListIndex = 0
    
    mblnOlnyBJYB = False: mlngOutModeMC = 0
    If mbytMode = 0 And Not mblnStation Then 'ԤԼ�ͽ��ղ�֧��,����ҽ��վ��֧��
        arrTmp = Split(GetSetting("ZLSOFT", "����ȫ��", "����֧�ֵ�ҽ��", ""), ",")
        strTemp = ""
        For i = 0 To UBound(arrTmp)
            If IsNumeric(arrTmp(i)) Then
                strTemp = strTemp & "," & Val(arrTmp(i))
                If CheckMCOutMode(arrTmp(i)) Then mlngOutModeMC = Val(arrTmp(i)): Exit For
            End If
        Next
        If strTemp <> "" Then strTemp = Mid(strTemp, 2)
        mblnOlnyBJYB = strTemp = "920"  '������:����:26982
    End If
    
      '����ȡ��ԤԼ�Һ������ ����ȡ��ԭ��
     cbo��ע.Clear
    
    'txtIDCard.Width = cbo��ͥ��ַ.Width '31182
    mobjfrmPatiInfo.mlngOutModeMC = mlngOutModeMC
    If mlngOutModeMC = 0 Then
        cboҽ�����.Enabled = False
'        If mbytMode = 1 Or mbytMode = 4 Then
'            cbo��ͥ��ַ.Width = txt��ͥ�绰.Width
'        Else
'            cbo��ͥ��ַ.Width = (cboҽ�����.Left + cboҽ�����.Width - cbo��ͥ��ַ.Left)
'        End If
        'txtIDCard.Width = cbo��ͥ��ַ.Width '31182
    Else
        cboҽ�����.Enabled = True
        strSQL = _
            "Select ����,����,Nvl(ȱʡ��־,0) as ȱʡ From ҽ����� Order by ����"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        cboҽ�����.AddItem ""
        For i = 1 To rsTmp.RecordCount
            cboҽ�����.AddItem rsTmp!���� & "-" & rsTmp!����
            If rsTmp!ȱʡ = 1 Then
                cboҽ�����.ItemData(cboҽ�����.NewIndex) = 1
            End If
            rsTmp.MoveNext
        Next
        cboҽ�����.ListIndex = 0
    End If
    
    '����:26955
    If mbytInState = 0 Then
        zlComboxLoadFromSQL "Select ����,����,����,ȱʡ��־ From ԤԼ��ʽ ", cboԤԼ��ʽ
        strTemp = zlDatabase.GetPara("ȱʡԤԼ��ʽ", glngSys, IIf(mblnStation, 1260, mlngModul), "")
        '�����:112838,����,2017/09/05,�����ֵ����δ�����κ�ԤԼ��ʽʱ�ᱨ��
        If cboԤԼ��ʽ.ListCount <> 0 Then
            For i = 0 To cboԤԼ��ʽ.ListCount - 1
                If Mid(cboԤԼ��ʽ.List(i), InStr(cboԤԼ��ʽ.List(i), ".") + 1) = strTemp Then
                    cboԤԼ��ʽ.ListIndex = i
                End If
            Next i
            If cboԤԼ��ʽ.ListIndex < 0 Then cboԤԼ��ʽ.ListIndex = 0
        End If
    End If
    
    If Not mblnStation Then
        strSQL = "Select Count(1) As ԭ�� From �����˺�ԭ��"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        mbln�˺�ԭ�� = Val(Nvl(rsTmp!ԭ��)) <> 0
    End If
    
    If mbytMode = 4 Then Call SetDelMemo("")
    
    If mbytInState = 0 Then
        If mbytMode = 0 Then
            Set mrsOneCard = GetOneCard
            mblnOneCard = mrsOneCard.RecordCount > 0
        End If
        
        '����������:����ʱ����Ҫ
        mbln������ = True
        If mbytMode <> 2 Then
            mbln������ = Not zlGetSpecialItemFee("������") Is Nothing
            If Not mbln������ Then chk������.Visible = False
        End If
        
        If mbytMode = 0 Or mbytMode = 1 Then chk������.Value = IIf(zlDatabase.GetPara("Ĭ�Ϲ�����", glngSys, mlngModul, 0) = "1", 1, 0)
        
        '���㷽ʽ:ԤԼʱ����Ҫ
        If mbytMode <> 1 Then
            Call Load֧����ʽ
            If cbo���㷽ʽ.ListCount = 0 Then
                '74550,Ƚ����,2014-7-2,�ڲ�����Ժ����,ҽ��������ҽ��վ�Һ�ʱ�ܹ�ѡ����㷽ʽ(��������Ϊ7��һ��ͨ����)
                If mblnStation Or mblnStationPrice Then
                    cbo���㷽ʽ.Visible = False: txt�ɿ�.Left = txt����Ӧ��.Left: txt�ɿ�.Width = txt����Ӧ��.Width '����
                End If
            End If
        End If
            
        '�ѱ�:����ʱ�������ٸ���
        If Not Init�ѱ�(True, False) Then mblnUnload = True: Exit Sub
        If cbo�ѱ�.ListCount = 0 Then
            MsgBox "�ѱ�ȼ�δ���ã����ȵ��ѱ���������÷ѱ�", vbInformation, gstrSysName
            mblnUnload = True: Exit Sub
        End If
    
        '�Ա�
        strSQL = "Select '�Ա�' as ���,����,����,����,Nvl(ȱʡ��־,0) as ȱʡ From �Ա� Union All " & _
                 " Select 'ҽ�Ƹ��ʽ' as ���,����,����,����,Nvl(ȱʡ��־,0) as ȱʡ From ҽ�Ƹ��ʽ " & _
                 " Order by ���,����"
        Set rsTmp = New ADODB.Recordset
        Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
        rsTmp.Filter = "���='�Ա�'"
        
        mblnNotChange = True
        cbo�Ա�.Clear
        Do While Not rsTmp.EOF
            cbo�Ա�.AddItem rsTmp!���� & "-" & rsTmp!����
            If rsTmp!���� = gstr�Ա� Then
                For i = 0 To cbo�Ա�.ListCount - 1
                    cbo�Ա�.ItemData(i) = 0
                Next
                cbo�Ա�.ItemData(cbo�Ա�.NewIndex) = 1
                cbo�Ա�.ListIndex = cbo�Ա�.NewIndex
            End If
            
            If rsTmp!ȱʡ = 1 And cbo�Ա�.ListIndex = -1 Then
                cbo�Ա�.ItemData(cbo�Ա�.NewIndex) = 1
                cbo�Ա�.ListIndex = cbo�Ա�.NewIndex
            End If
            
            rsTmp.MoveNext
        Loop
        If gstr�Ա� = "��" Then cbo�Ա�.ListIndex = -1
        mblnNotChange = False
        
        'ҽ�Ƹ��ʽ
        rsTmp.Filter = "���='ҽ�Ƹ��ʽ'"
        cbo���ʽ.Clear
        Do While Not rsTmp.EOF
            cbo���ʽ.AddItem rsTmp!���� & "-" & rsTmp!����
            If rsTmp!���� = gstr���ʽ Then
                For i = 0 To cbo���ʽ.ListCount - 1
                    cbo���ʽ.ItemData(i) = 0
                Next
                cbo���ʽ.ItemData(cbo���ʽ.NewIndex) = 1
                cbo���ʽ.ListIndex = cbo���ʽ.NewIndex
            End If
            If rsTmp!ȱʡ = 1 Then
                If cbo���ʽ.ListIndex = -1 Then
                    cbo���ʽ.ItemData(cbo���ʽ.NewIndex) = 1
                    cbo���ʽ.ListIndex = cbo���ʽ.NewIndex
                End If
            End If
            rsTmp.MoveNext
        Loop
        If cbo���ʽ.ListIndex = -1 And cbo���ʽ.ListCount > 0 Then cbo���ʽ.ListIndex = 0
        
        If cbo��ͥ��ַ.Enabled And Not mblnStructAdress Then
            Call Load��ͥ��ַ
        End If
        Set mobjfrmPatiInfo.mrsBaseDict = GetBaseDict   '���ڹҺŲ��˴�����ֵ��ʼ
        Set mrsDoctor = New ADODB.Recordset
        If Not mblnStation Then Call GetAllҽ��
         
                
        'A.����
        If mbytMode = 2 Then
            If ReadBooking(mstrNoIn) = False Then
                mblnUnload = True
                Exit Sub
            Else
                If mrsInfo Is Nothing And mbytMode = 2 Then cbo�ѱ�.Enabled = True
            End If
            'ԤԼ����
            If CheckIsPrice Then
                Call SetUndisplayBalance
            Else
                Call SetShowBalance
            End If
            
        'B.�ҺŻ�ԤԼ
        Else
            '�Һ�����,ShowPlans�е�vsfplan_EnterCell���õ�����
            Curdate = zlDatabase.Currentdate
            
            If mbytMode = 1 Then
                If Curdate < gdatRegistTime Then
                    dtpAppointmentDate.Value = Format(gdatRegistTime + mTy_Para.lngԤԼȱʡ����, "yyyy-MM-dd " & gstr�ϰ�ʱ��)
                    dtpAppointmentDate.MinDate = Format(gdatRegistTime, "yyyy-MM-dd 00:00")
                Else
                    dtpAppointmentDate.Value = Format(Curdate + mTy_Para.lngԤԼȱʡ����, "yyyy-MM-dd " & gstr�ϰ�ʱ��)
                    dtpAppointmentDate.MinDate = Format(Curdate, "yyyy-MM-dd 00:00")  '27781:��ǰ��һСʱ
                End If
            End If
        
            Call ShowPlans
        
            '�����жϵ����ű𳤶� GetMaxLen
            gint�ų� = 5
            If Not mrsPlan Is Nothing Then
                If mrsPlan.State = 1 Then
                    gint�ų� = 1
                    mrsPlan.MoveFirst
                    For i = 1 To mrsPlan.RecordCount
                        If Len(mrsPlan!�ű�) > gint�ų� Then gint�ų� = Len(mrsPlan!�ű�)
                        mrsPlan.MoveNext
                    Next
                End If
            Else
                gint�ų� = GetMaxLen
            End If
        End If
        '79619:���ϴ�,2014/11/13,��ʾȱʡ�ĹҺ�ժҪ
        strSQL = "Select ����,����,���� " & _
                 " From ���ùҺ�ժҪ " & _
                 " Where Nvl(ȱʡ��־,0)=1"
        Set rsTmp = New ADODB.Recordset
        Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
        If Not rsTmp.EOF Then
            cbo��ע.Text = rsTmp!����
        End If
        'ˢ��Ʊ�ݺ�
        If mbytMode <> 1 And Not mblnStation And gbytInvoice <> 0 And Not mblnStartFactUseType Then
            If Not RefreshFact Then mblnUnload = True: Exit Sub
        End If
    Else '�鿴
        Call ReadBill(mstrNoIn)
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub Set���㷽ʽEanbled()
    '���ý��㷽ʽ��enabled����
     If mbytInState = 0 Then    '0-ִ��,1-����
        cbo���㷽ʽ.Enabled = gbln���㷽ʽ And mbytMode <> 1
     End If
End Sub
Private Sub SetShowBalance()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾ����ؼ�
    '����:���˺�
    '����:2013-12-24 15:49:21
    '����:68991
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnVisible As Boolean
    '74522:���ϴ�,2014-6-27,ҽ������վ�ҺŲ���ʾ���㷽ʽ����Ϣ
    If mbytInState = 1 Or mblnStation Or mbytInState = 0 And mbytMode = 1 Then Exit Sub
    '��ʾ���㷽ʽ
    blnVisible = True
    lblFact.Visible = blnVisible: txtFact.Visible = blnVisible
    If blnVisible Then
        cbo���㷽ʽ.Visible = True
        txt�ɿ�.Left = cbo���㷽ʽ.Left + cbo���㷽ʽ.Width + 30
        txt�ɿ�.Width = 1305
        vsfPay.Visible = True
    Else
        cbo���㷽ʽ.Visible = False
        txt�ɿ�.Left = txt����Ӧ��.Left
        txt�ɿ�.Width = txt����Ӧ��.Width
        vsfPay.Visible = False
    End If
    lblӦ��.Visible = blnVisible: txt����Ӧ��.Visible = blnVisible
    lbl�ɿ�.Visible = blnVisible: txt�ɿ�.Visible = blnVisible
    lbl�Ҳ�.Visible = blnVisible: txt�Ҳ�.Visible = blnVisible
    
    lblSum.Caption = "�ϼ�"
    lblTotal.Caption = lblSum.Caption
    lblSum.Visible = blnVisible: txt�ϼ�.Visible = blnVisible
    picTotal.Visible = Not blnVisible
End Sub
Private Sub SetUndisplayBalance()
    '���ò���ʾ���������Ϣ
    Dim blnVisible As Boolean
    Dim blnPrice As Boolean
    
    If (mbytInState = 0 Or mbytInState = 1 And mbytMode = 4 Or chkCancel.Value = 1) Then
        blnPrice = CheckIsPrice
        If mRegistFeeMode = EM_RG_���� Or blnPrice Then
            '68991:�ҺŷѲ��ü��ʷ�ʽ,��Ӧ����ʾ����������Ϣ
            cbo���㷽ʽ.Visible = False
            txt�ɿ�.Left = txt����Ӧ��.Left
            txt�ɿ�.Width = txt����Ӧ��.Width
            lblӦ��.Visible = False: txt����Ӧ��.Visible = False
            lbl�ɿ�.Visible = False: txt�ɿ�.Visible = False
            lbl�Ҳ�.Visible = False: txt�Ҳ�.Visible = False
            lblFact.Visible = False: txtFact.Visible = False
            lblSum.Caption = IIf(blnPrice, "����", "����")
            lblTotal.Caption = lblSum.Caption
            vsfPay.Visible = False
            lblSum.Visible = False: txt�ϼ�.Visible = False
            picTotal.Visible = True
            Exit Sub
        End If
    End If
    
    If mbytInState = 1 And mbytMode = 0 And mRegistFeeMode = EM_RG_���� Then
        cbo���㷽ʽ.Visible = False
        txt�ɿ�.Left = txt����Ӧ��.Left
        txt�ɿ�.Width = txt����Ӧ��.Width
        lblӦ��.Visible = False: txt����Ӧ��.Visible = False
        lbl�ɿ�.Visible = False: txt�ɿ�.Visible = False
        lbl�Ҳ�.Visible = False: txt�Ҳ�.Visible = False
        lblFact.Visible = False: txtFact.Visible = False
        lblSum.Caption = "����"
        lblTotal.Caption = lblSum.Caption
        vsfPay.Visible = False
        lblSum.Visible = False: txt�ϼ�.Visible = False
        picTotal.Visible = True
        picTotal.Width = picBal.Left - picTotal.Left - 30
        lbl�ϼ�.Left = picTotal.Width - lbl�ϼ�.Width - 60
        Exit Sub
    End If
    If mbytInState = 1 And mbytMode = 4 Or chkCancel.Value = 1 Then
        '���˺�:�˺�,ֻ��Ҫ��ʾ�˺ŷ�ʽ
        blnVisible = True
        lblFact.Visible = blnVisible: txtFact.Visible = blnVisible
        If blnVisible Then
            cbo���㷽ʽ.Visible = False
            txt�ɿ�.Left = cbo���㷽ʽ.Left + cbo���㷽ʽ.Width + 30
            txt�ɿ�.Width = 1305
        Else
            cbo���㷽ʽ.Visible = False
            txt�ɿ�.Left = txt����Ӧ��.Left
            txt�ɿ�.Width = txt����Ӧ��.Width
        End If
        lblӦ��.Visible = blnVisible: txt����Ӧ��.Visible = blnVisible
        lblӦ��.ForeColor = vbRed: txt����Ӧ��.ForeColor = vbRed
        lbl�ɿ�.Visible = Not blnVisible: txt�ɿ�.Visible = Not blnVisible
        lbl�Ҳ�.Visible = Not blnVisible: txt�Ҳ�.Visible = Not blnVisible
        lblӦ��.Caption = "�˿�": txt����Ӧ��.ToolTipText = "�����˿�=�ۼ�ʵ�ɽ��-�ۼ��˸����ʻ�-�ۼ��˳�Ԥ����"
        lblSum.Visible = blnVisible: txt�ϼ�.Visible = blnVisible
        picTotal.Visible = Not blnVisible
    ElseIf mbytInState = 0 Then
        blnVisible = mbytInState = 0 Or mbytInState = 1 And mbytMode <> 0
        If blnVisible Then
            cbo���㷽ʽ.Visible = True
            txt�ɿ�.Left = cbo���㷽ʽ.Left + cbo���㷽ʽ.Width + 30
            txt�ɿ�.Width = 1305
        Else
            cbo���㷽ʽ.Visible = False
            txt�ɿ�.Left = txt����Ӧ��.Left
            txt�ɿ�.Width = txt����Ӧ��.Width
        End If
        If mbytMode = 1 Then
            cbo���㷽ʽ.Visible = False
            txt�ɿ�.Left = txt����Ӧ��.Left
            txt�ɿ�.Width = txt����Ӧ��.Width
            lblFact.Visible = False: txtFact.Visible = False
            lbl�ɿ�.Visible = False: txt�ɿ�.Visible = False
            lbl�Ҳ�.Visible = False: txt�Ҳ�.Visible = False
            txt����Ӧ��.Visible = False
            lblSum.Visible = False: txt�ϼ�.Visible = False
            picTotal.Visible = True
        Else
            lblFact.Visible = blnVisible: txtFact.Visible = blnVisible
            lbl�ɿ�.Visible = blnVisible: txt�ɿ�.Visible = blnVisible
            lbl�Ҳ�.Visible = blnVisible: txt�Ҳ�.Visible = blnVisible
            txt����Ӧ��.Visible = blnVisible
            lblSum.Visible = blnVisible: txt�ϼ�.Visible = blnVisible
            picTotal.Visible = Not blnVisible
        End If
        lblӦ��.ForeColor = lbl�ɿ�.ForeColor: txt����Ӧ��.ForeColor = &H108000
        lblӦ��.Caption = "Ӧ��"
        txt����Ӧ��.ToolTipText = "����Ӧ�ɺϼ� = �ۼ�ʵ�ɽ�� - �ۼƸ����ʻ�֧�� - �ۼƳ�Ԥ����"
    ElseIf mblnViewCancel Then
        '��ʾ�˵�����
        blnVisible = True
        cbo���㷽ʽ.Visible = True
        txt�ɿ�.Left = cbo���㷽ʽ.Left + cbo���㷽ʽ.Width + 30
        txt�ɿ�.Width = 1590
        lblӦ��.Visible = False: txt����Ӧ��.Visible = False
        lblӦ��.ForeColor = vbRed: txt����Ӧ��.ForeColor = vbRed
        lbl�ɿ�.Visible = Not blnVisible: txt�ɿ�.Visible = Not blnVisible
        lbl�Ҳ�.Visible = Not blnVisible: txt�Ҳ�.Visible = Not blnVisible
        lblӦ��.Caption = "�˿�"
        txt����Ӧ��.ToolTipText = "�����˿�=�ۼ�ʵ�ɽ��-�ۼ��˸����ʻ�-�ۼ��˳�Ԥ����"
        lblSum.Visible = False: txt�ϼ�.Visible = False
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
        lblԤԼʱ��.Visible = chkBooking.Value = 1
        dtpAppointmentTime.Visible = chkBooking.Value = 1
    Else
        lblRegTotal(0).Visible = False
        lblRegTotal(1).Visible = False
        lblԤԼʱ��.Visible = True
        dtpAppointmentTime.Visible = True
    End If
End Sub

Private Sub SetPlanGrid()
    Dim i As Integer, strHead As String
    
    '133363:���ϴ�,2018/11/1,�����С�ʣ��ɹҺϼơ���������ʾ�޺źű��ʣ��ɹ�����
    'ÿ������ColData���ԣ����ڿ������Ƿ�ɼ�
    '��ʼ���ű�
    '����,����,�п�,�Ƿ��ѡ(1-�̶�,-1-����ѡ,0-��ѡ)
    If mbytMode = 1 Then
        strHead = "RowNum,1,285,-1|IDS,1,0,-1|����,1,500,1|�ű�,1,750,1|����,1,1200,1|��Ŀ,1,1500,0|ҽ��,1,700,0|����ҽ��,1,1000,0|ʱ��,1,750,0|ʣ��,1,500,-1|�ѹ�,1,500,-1|�޺�,1,500,-1|��Լ,1,500,0|��Լ,1,500,0" & _
            "|��,4,450,0|һ,4,450,0|��,4,450,0|��,4,450,0|��,4,450,0|��,4,450,0|��,4,450,0" & _
            "|����,4,500,0|����,4,500,0|��ſ���,4,850,0|��������,1,1100,0|��¼ID,1,0,-1|��ǰʱ��,1,0,-1|�Һ�ʱ��,1,0,-1|��ԴID,1,0,-1|�Ƿ��ռ,1,0,-1|���￪ʼʱ��,1,0,-1|������ֹʱ��,1,0,-1|����ҽ������,1,0,-1|����ҽ��ID,1,0,-1|��ʱ��,1,0,-1|����ʱ��,1,0,-1"
    Else
        strHead = "RowNum,1,285,-1|IDS,1,0,-1|����,1,500,1|�ű�,1,750,1|����,1,1200,1|��Ŀ,1,1500,0|ҽ��,1,700,0|����ҽ��,1,1000,0|ʱ��,1,750,0|ʣ��,1,500,0|�ѹ�,1,500,0|�޺�,1,500,0|��Լ,1,500,0|��Լ,1,500,0" & _
            "|��,4,450,0|һ,4,450,0|��,4,450,0|��,4,450,0|��,4,450,0|��,4,450,0|��,4,450,0" & _
            "|����,4,500,0|����,4,500,0|��ſ���,4,850,0|��������,1,1100,0|��¼ID,1,0,-1|��ǰʱ��,1,0,-1|�Һ�ʱ��,1,0,-1|��ԴID,1,0,-1|�Ƿ��ռ,1,0,-1|���￪ʼʱ��,1,0,-1|������ֹʱ��,1,0,-1|����ҽ������,1,0,-1|����ҽ��ID,1,0,-1|��ʱ��,1,0,-1|����ʱ��,1,0,-1"
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
            .ColHidden(.ColIndex("ʣ��")) = True
        End If
        If mbytMode = 1 Then
            .ColHidden(.ColIndex("�ѹ�")) = True: .ColHidden(.ColIndex("�޺�")) = True
        End If
        .ColHidden(.ColIndex("IDS")) = True
        .ColHidden(.ColIndex("��ǰʱ��")) = True: .ColHidden(.ColIndex("�Һ�ʱ��")) = True
        .ColHidden(.ColIndex("��ԴID")) = True: .ColHidden(.ColIndex("�Ƿ��ռ")) = True
        .ColHidden(.ColIndex("���￪ʼʱ��")) = True: .ColHidden(.ColIndex("������ֹʱ��")) = True
        .ColHidden(.ColIndex("����ҽ������")) = True: .ColHidden(.ColIndex("����ҽ��ID")) = True
        .ColHidden(.ColIndex("��ʱ��")) = True
        .RowHeight(0) = 500
        .RowData(0) = 0
        .Redraw = flexRDBuffered
    End With
End Sub

Private Function zlCheck��Լ���޺���(ByVal lng��¼ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����Լ�����޺����Ƿ�Ϸ�
    '���:str�ű�-�ű�
    '����:
    '����:�Ϸ�,����ture,���򷵻�False
    '����:���˺�
    '����:2009-12-30 15:15:57
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, lngTemp As Long, strSQL As String, Curdate As Date
    Dim lng��Լ�� As Long, lng�޺��� As Long, lng�ѹ��� As Long, lng��Լ�� As Long, lngʣ��ԤԼ�� As Long
    Dim lngʧԼ�� As Long
    Dim lng�ѽ��� As Long
    Dim bln��ʱ�� As Boolean
    Dim strMsg As String, int���Ʒ�ʽ As Integer
    Dim lng������λ���� As Long
    Dim blnHaveUnitreg As Boolean
    Dim i As Integer, j As Integer
    Err = 0: On Error GoTo Errhand:
    lng��Լ�� = 0: lng�޺��� = 0: lng�ѹ��� = 0: lng��Լ�� = 0: lngʣ��ԤԼ�� = 0
    mbln�Ӻ� = False
    If picBookingDate.Visible Then
        Curdate = CDate(Format(dtpAppointmentDate.Value, IIf(bln��ʱ��, "yyyy-MM-dd HH:mm:ss", "yyyy-MM-dd")))
    Else
        Curdate = CDate(Format(zlDatabase.Currentdate, "yyyy-MM-dd"))
    End If
    strSQL = "Select Nvl(�޺���, 0) As �޺���, ��Լ��, Nvl(�ѹ���, 0) As �ѹ���, Nvl(��Լ��, 0) As ��Լ��, Nvl(�����ѽ���, 0) As �ѽ���" & vbNewLine & _
            "From �ٴ������¼ Where Id = [1]"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng��¼ID)
    
    If mbytMode = 0 Then
        lngʧԼ�� = GetʧԼ��(lng��¼ID, Curdate)
    End If
    
    If Not rsTmp.EOF Then
        lng��Լ�� = Val(Nvl(rsTmp!��Լ��)): lng�޺��� = Val(Nvl(rsTmp!�޺���))
        lng�ѹ��� = Val(Nvl(rsTmp!�ѹ���)): lng��Լ�� = Val(Nvl(rsTmp!��Լ��)) - Val(Nvl(rsTmp!�ѽ���))
        lng�ѽ��� = Val(Nvl(rsTmp!�ѽ���))
        If lng��Լ�� < 0 Then lng��Լ�� = 0
        lngʣ��ԤԼ�� = IIf(lng�޺��� - lng�ѹ��� - lng��Լ�� <= 0, 0, lng��Լ�� - lng��Լ��): If lngʣ��ԤԼ�� < 0 Then lngʣ��ԤԼ�� = 0
        If lng��Լ�� = 0 And IsNull(rsTmp!��Լ��) Then lng��Լ�� = lng�޺���
        lng��Լ�� = lng��Լ�� - lngʧԼ��
    End If
    If lng�޺��� <= 0 Then
        '��������:����
        zlCheck��Լ���޺��� = True: Exit Function
    End If
    If (mbytMode = 1 Or chkBooking.Value = 1) And mblnUnitReg And Not mrsUnitReg Is Nothing Then
        mrsUnitReg.Filter = 0
        If mrsUnitReg.RecordCount <> 0 Then
            int���Ʒ�ʽ = Val(Nvl(mrsUnitReg!���Ʒ�ʽ))
        End If
        If mViewMode = V_��ͨ�� And mrsUnitReg.RecordCount > 0 Then
           If Val(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("�Ƿ��ռ"))) = 0 Then
               lng������λ���� = 0
           Else
                If int���Ʒ�ʽ = 1 Then
                    Do While Not mrsUnitReg.EOF
                        lng������λ���� = lng������λ���� + Int(Val(Nvl(mrsUnitReg!����)) * lng��Լ�� / 100)
                        mrsUnitReg.MoveNext
                    Loop
                Else
                    Do While Not mrsUnitReg.EOF
                        lng������λ���� = lng������λ���� + Val(Nvl(mrsUnitReg!����))
                        mrsUnitReg.MoveNext
                    Loop
                End If
                mrsUnitReg.MoveFirst
           End If
        End If
        If mViewMode = V_��ͨ�ŷ�ʱ�� And mrsUnitReg.RecordCount > 0 Then
            If Val(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("�Ƿ��ռ"))) = 0 Then
                lng������λ���� = 0
            Else
                If int���Ʒ�ʽ = 1 Then
                    Do While Not mrsUnitReg.EOF
                        lng������λ���� = lng������λ���� + Int(Val(Nvl(mrsUnitReg!����)) * lng��Լ�� / 100)
                        mrsUnitReg.MoveNext
                    Loop
                ElseIf int���Ʒ�ʽ = 2 Then
                    Do While Not mrsUnitReg.EOF
                        lng������λ���� = lng������λ���� + Val(Nvl(mrsUnitReg!����))
                        mrsUnitReg.MoveNext
                    Loop
                End If
                mrsUnitReg.MoveFirst
            End If
        End If
        If (mViewMode = v_ר�Һ� Or mViewMode = v_ר�Һŷ�ʱ��) And mrsUnitReg.RecordCount > 0 Then
            If int���Ʒ�ʽ = 3 Then
                Do While Not mrsUnitReg.EOF
                    lng������λ���� = lng������λ���� + Val(Nvl(mrsUnitReg!����))
                    mrsUnitReg.MoveNext
                Loop
                mrsUnitReg.MoveFirst
            Else
                If Val(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("�Ƿ��ռ"))) = 0 Then
                    lng������λ���� = 0
                Else
                    If int���Ʒ�ʽ = 1 Then
                        Do While Not mrsUnitReg.EOF
                            lng������λ���� = lng������λ���� + Int(Val(Nvl(mrsUnitReg!����)) * lng��Լ�� / 100)
                            mrsUnitReg.MoveNext
                        Loop
                    ElseIf int���Ʒ�ʽ = 2 Then
                        Do While Not mrsUnitReg.EOF
                            lng������λ���� = lng������λ���� + Val(Nvl(mrsUnitReg!����))
                            mrsUnitReg.MoveNext
                        Loop
                    End If
                    mrsUnitReg.MoveFirst
                End If
            End If
        End If
       '�ų��Ѿ��ҳ��ĺ�����λ��
       strSQL = "Select Count(1) As ��Լ�� From ���˹Һż�¼ Where ��¼״̬ = 1 And �����¼ID = [1] And ������λ Is Not Null "
       Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng��¼ID)
       If Not rsTmp.EOF Then
            lng������λ���� = lng������λ���� - Val(rsTmp!��Լ��)
       End If
       If lng������λ���� < 0 Then lng������λ���� = 0
    End If
    
    '************************************************************************
    '�޺���-��Լ��-�ѹ���>0  | �޺���>��Լ�� |�����Լ��=0 ��Լ��=�޺���
    '�ﵽ�޺�������ԤԼʱ�ﵽ��Լ��
    '   ����Աӵ�мӺ�Ȩ�� ��ʾ �ò���Ա�Լ�ѡ���Ƿ�����ҺŻ���ԤԼ
    '   ����Աû�мӺ�Ȩ�� ���ֹ����Ա�����ҺŻ���ԤԼ
    '************************************************************************
    
    
    'mbytMode:0-�Һ�,1-ԤԼ,2-����,ԤԼ������ģʽ:0-�Һ�,��ʱԤԼҪ�շ�,1-ԤԼ,���շ�
    Select Case mbytMode
    Case 1:  'ԤԼ
         If lng�޺��� - lng�ѹ��� - lng��Լ�� - lng������λ���� > 0 Then
            '----------------------------------------------
            '�Һ�+ԤԼ�� û�дﵽ�޺���
            '----------------------------------------------
            
             If lng��Լ�� + lng�ѽ��� + lng������λ���� >= lng��Լ�� Then
                If InStr(mstrPrivs, ";�Ӻ�;") > 0 Then  '��Ҫ��ʾ:
                     If MsgBox("�úű�����Ѵﵽ��Լ��" & lng��Լ�� & IIf(lng������λ���� > 0, "(���а����Һź�����λ��������[" & lng������λ���� & "])", "") & " �����Ƿ����ԤԼ?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        If mbytMode <> 2 Then txt�ű� = "": If txt�ű�.Enabled And txt�ű�.Visible Then txt�ű�.SetFocus
                        Exit Function
                    End If
                    mbln�Ӻ� = True
                Else
                    MsgBox "�úű�����Ѵﵽ��Լ�� " & lng��Լ�� & IIf(lng������λ���� > 0, "(���а����Һź�����λ��������[" & lng������λ���� & "])", "") & "��������ԤԼ��", vbInformation, gstrSysName
                    If mbytMode <> 2 Then txt�ű� = "": If txt�ű�.Enabled And txt�ű�.Visible Then txt�ű�.SetFocus
                    Exit Function
                End If
            End If
        Else
          '------------------------------------------
           '�ѹ���+��Լ�� �ﵽ���޺���
           '����Աӵ�мӺ���Ȩ�� �ò���Աѡ���Ƿ����
           '------------------------------------------
           If InStr(mstrPrivs, ";�Ӻ�;") > 0 Then
                                If MsgBox("�úű�����Ѵﵽ�޺��� " & lng�޺��� & IIf(lng������λ���� > 0, "(���а����Һź�����λ��������[" & lng������λ���� & "])", "") & "�����Ƿ����ԤԼ?", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
                    If mbytMode <> 2 Then txt�ű� = "": If txt�ű�.Enabled And txt�ű�.Visible Then txt�ű�.SetFocus
                    Exit Function
                End If
                mbln�Ӻ� = True
           Else
                                        MsgBox "�úű�����Ѵﵽ�޺��� " & lng�޺��� & IIf(lng������λ���� > 0, "(���а����Һź�����λ��������[" & lng������λ���� & "])", "") & "������ԤԼ��", vbInformation, gstrSysName
                    If mbytMode <> 2 Then txt�ű� = "": If txt�ű�.Enabled And txt�ű�.Visible Then txt�ű�.SetFocus
                    Exit Function
                
           End If
        End If
    Case Else '�Һ�,����
        If mbytMode = 0 And chkBooking.Value = 0 Then
            '�Һ�
            If lng�ѹ��� + lng��Լ�� >= lng�޺��� Then
                If InStr(mstrPrivs, ";�Ӻ�;") > 0 Then
                    If MsgBox("�úű�����Ѵﵽ�޺��� " & lng�޺��� & "�����Ƿ�����Һ�?", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
                         If mbytMode <> 2 Then txt�ű� = "": If txt�ű�.Enabled And txt�ű�.Visible Then txt�ű�.SetFocus
                         Exit Function
                    End If
                    If mbytMode = 0 Then
                        With vsfList
                            For i = 0 To .Rows - 1
                                For j = 0 To .Cols - 1
                                    If .Cell(flexcpData, i, j) Like "��*" Then .Select i, j
                                Next j
                            Next i
                        End With
                    End If
                    mbln�Ӻ� = True
                Else
                    MsgBox "�úű�����Ѵﵽ�޺��� " & lng�޺��� & "�����ٹҺţ�", vbInformation, gstrSysName
                    If mbytMode <> 2 Then txt�ű� = "": If txt�ű�.Enabled And txt�ű�.Visible Then txt�ű�.SetFocus
                    Exit Function
                End If
            End If
        Else
            '����
            If lng�ѹ��� >= lng�޺��� Then
                If InStr(mstrPrivs, ";�Ӻ�;") > 0 Then
                    If MsgBox("�úű�����Ѵﵽ�޺��� " & lng�޺��� & "�����Ƿ�����Һ�?", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
                         If mbytMode <> 2 Then txt�ű� = "": If txt�ű�.Enabled And txt�ű�.Visible Then txt�ű�.SetFocus
                         Exit Function
                    End If
                    If mbytMode = 0 Then
                        With vsfList
                            For i = 0 To .Rows - 1
                                For j = 0 To .Cols - 1
                                    If .Cell(flexcpData, i, j) Like "��*" Then .Select i, j
                                Next j
                            Next i
                        End With
                    End If
                    mbln�Ӻ� = True
                Else
                    MsgBox "�úű�����Ѵﵽ�޺��� " & lng�޺��� & "�����ٹҺţ�", vbInformation, gstrSysName
                    If mbytMode <> 2 Then txt�ű� = "": If txt�ű�.Enabled And txt�ű�.Visible Then txt�ű�.SetFocus
                    Exit Function
                End If
            End If
        End If
    End Select
    zlCheck��Լ���޺��� = True
   
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function

Private Function GetHave(ByVal lng��¼ID As Long) As String
    '����:ȡָ���ű��޺������ѹ���
    '����:"�޺���;�ѹ���;ʣ��ԤԼ��"��"��Լ��;��Լ��;ʣ��ԤԼ��"
    '���˺� ����:26962 ����:2009-12-25 11:46:30 Modify:ʣ��ԤԼ��
    Dim rsTmp As ADODB.Recordset, lngTemp As Long
    Dim strSQL As String, Curdate As Date
    
    GetHave = "0;0;0"
    If picBookingDate.Visible Then
        Curdate = CDate(Format(dtpAppointmentDate.Value, "yyyy-MM-dd"))
    Else
        Curdate = CDate(Format(zlDatabase.Currentdate, "yyyy-MM-dd"))
    End If
    
    strSQL = "Select Nvl(�޺���, 0) As �޺���, Nvl(�ѹ���, 0) - Nvl(�����ѽ���, 0) As �ѹ���, Nvl(��Լ��, 0) As ��Լ��, Nvl(��Լ��, 0) As ��Լ��" & vbNewLine & _
            "From �ٴ������¼" & vbNewLine & _
            "Where ID = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng��¼ID)
    
    If Not rsTmp.EOF Then
        lngTemp = Val(Nvl(rsTmp!��Լ��)) - Val(Nvl(rsTmp!��Լ��))
        If lngTemp < 0 Then lngTemp = 0
        If mbytMode = 1 Then
            GetHave = rsTmp!��Լ�� & ";" & rsTmp!��Լ�� & ";" & lngTemp
        Else
            GetHave = rsTmp!�޺��� & ";" & rsTmp!�ѹ��� & ";" & lngTemp
        End If
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function ShowPlans(Optional strSort As String, Optional blnCache As Boolean, Optional ByVal blnAutoUpdate As Boolean = True, Optional ByVal blnShowStop As Boolean = False) As Boolean
'���ܣ���ȡ���հ�������
'blnCache:��������ű�δ�ﵽ��󳤶�ʱ�Ż���,��Ҫ�ǿ����޺�ʱ���ڱ�
    Dim strTime As String, strState As String
    Dim strSQL As String, strIF As String
    Dim i As Integer, k As Integer, rsDays As ADODB.Recordset
    Dim DateThis As Date, strZero As String, intGap As Integer, intSurplusTotal As Integer
    Dim str�ҺŰ��� As String, strDays As String
    Dim str�ҺŰ��żƻ� As String, str��ԴIDs As String
    Dim str����         As String, strSpan As String
    Dim dat��ʼʱ�� As Date, dat����ʱ�� As Date, datNow As Date
    Dim str��������� As String, rs������ As ADODB.Recordset, str������SQL As String
    Dim strStationSql As String, str������IF As String
    Dim str��ԴIDSQL As String, str���������SQL As String
    
    On Error GoTo errH
    If mblnUnload Then Exit Function
    Select Case mSortType
        Case by�ű�:
                str���� = "�ű�,�������� Desc,�Һ�ʱ��"
        Case by����:
                str���� = "����,��Ŀ,�ѹ�,�������� Desc,�Һ�ʱ��"
        Case by����and�ѹ���:
                str���� = "����,�ѹ�,�������� Desc,�Һ�ʱ��"
        Case Else:
             str���� = "�ű�,�������� Desc,�Һ�ʱ��"
    End Select
    
    If strSort = "" Then strSort = IIf(mstrSort = "", str����, mstrSort)
    If InStr(1, strSort, str����) = 0 Then strSort = strSort & "," & str����
    If blnCache Then blnCache = Not mrsPlan Is Nothing
    
    If Not blnCache Then
        datNow = zlDatabase.Currentdate
        If picBookingDate.Visible Or mbytMode = 1 Or (mbytMode = 2 Or mbytMode = 0 And mbytInState = 0 And mstrNoIn <> "") Then
            DateThis = dtpAppointmentDate.Value
        Else
            DateThis = zlDatabase.Currentdate
        End If
        
        strSQL = "Select a.Id, b.���� As �ű�, a.��������, b.����, b.����id, c.���� As ����, a.��Ŀid, d.���� As ��Ŀ, a.����ҽ��id, a.ҽ��id, a.����ҽ������, a.ҽ������ As ҽ��, Nvl(a.�ѹ���, 0) As �ѹ�," & vbNewLine & _
                "       Nvl(a.��Լ��, 0) As ��Լ, a.�޺��� As �޺�, a.��Լ�� As ��Լ, Nvl(b.�Ƿ񽨲���, 0) As ����, Nvl(d.��Ŀ����, 0) As ����, Decode(a.���﷽ʽ,1,'ָ��',2,'��̬',3,'ƽ��',NULL) As ����," & vbNewLine & _
                "       a.�Ƿ���ſ��� As ��ſ���, a.�ϰ�ʱ�� As �Ű�, a.��Դid, a.��ǰ�Һ�ʱ�� As ��ǰʱ��, a.��ʼʱ�� As �Һ�ʱ��, a.��ֹʱ�� As ����ʱ��, Nvl(a.�Ƿ��ռ,0) As �Ƿ��ռ, a.���￪ʼʱ�� , a.������ֹʱ��, a.�Ƿ��ʱ�� As ��ʱ�� " & vbNewLine & _
                "From �ٴ������¼ A, �ٴ������Դ B, ���ű� C, �շ���ĿĿ¼ D, ��Ա�� E" & vbNewLine & _
                "Where (a.�������� = [6] Or a.�������� = [8]) And a.��Դid = b.Id  And Nvl(C.����ʱ��,To_Date('3000-01-01','YYYY-MM-DD')) > Sysdate And b.����id = c.Id And a.��Ŀid = d.Id And Nvl(a.�Ƿ�����, 0) = 0 " & vbNewLine & _
                "       And a.ҽ��id = e.Id(+) And (d.����ʱ�� is NULL Or d.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD')) And (c.վ��='" & gstrNodeNo & "' Or c.վ�� is Null) " & _
                "       And Nvl(a.�Ƿ񷢲�,0) = 1 "
        
        If mbytMode = 1 Or chkBooking.Value = 1 Then
            If Format(DateThis, "yyyy-mm-dd") = Format(datNow, "yyyy-mm-dd") Then
                strSQL = strSQL & "       And Not Exists (Select 1 From �ٴ������¼ Where Id=a.Id And ��ֹʱ�� < [9])"
            Else
                strSQL = strSQL & "       And Not Exists (Select 1 From �ٴ������¼ Where Id=a.Id And ��ֹʱ�� < [6])"
            End If
        Else
            strSQL = strSQL & "       And Not Exists (Select 1 From �ٴ������¼ Where Id=a.Id And ��ֹʱ�� < [5])"
        End If
        
        If Not blnShowStop And chkShowAll.Value <> 1 Then
            strSQL = strSQL & " And (a.��ʼʱ�� < Nvl(a.ͣ�￪ʼʱ��,a.��ֹʱ��) Or a.��ֹʱ�� > Nvl(a.ͣ����ֹʱ��,a.��ʼʱ��) Or Exists (Select 1 From �ٴ�������ſ��� C,�ٴ������¼ D Where D.ID=A.ID And C.��¼ID=D.ID And Nvl(C.�Ƿ�ͣ��,0) = 0 And D.�Ƿ���ſ��� =1 And D.�Ƿ��ʱ�� = 1 And C.��ʼʱ�� <> C.��ֹʱ��)) "
        End If
        
        If chkShowAll.Value <> 1 Then
            strSQL = strSQL & " And [5] Not Between Nvl(a.ͣ�￪ʼʱ��,a.��ֹʱ��) And Nvl(a.ͣ����ֹʱ��,a.��ʼʱ��) "
        End If
        
        If (mbytMode = 0 Or mbytMode = 1) And mstrNoIn = "" Then
            '�Һź�ԤԼʱ������ʾ��ͣ�û�ɾ���ĺ�Դ
            strSQL = strSQL & " And Nvl(b.�Ƿ�ɾ��, 0) = 0 And (b.����ʱ�� >= To_Date('3000-01-01', 'yyyy-mm-dd') Or b.����ʱ�� Is Null)"
        End If
        
        If gstr�Һſ���ID <> "" Then
            '���ز���ȷ���˵ĹҺſ���
            strIF = " And Instr(','||[4]||',',','||b.����ID||',')>0"
        End If
        
        '������ĺű���ˣ����ű���������вŹ���,��ʱ��ActiveControlһ����txt�ű�
        If Trim(txt�ű�.Text) <> "" And Trim(txt�ű�.Text) <> "+" And ActiveControl Is txt�ű� And mblnReadBooking = False Then
            If IsNumeric(Trim(txt�ű�.Text)) Then
                strIF = strIF & " And b.���� Like [2]"
            Else
                strIF = strIF & " And (zlSpellCode(e.����) Like [2] or c.���� Like [2])"
            End If
        End If
        
        strSQL = strSQL & strIF
   
        If chkShowAll.Value = 1 Then
            '���а���
            If mbytMode = 1 Or chkBooking.Value = 1 Then
                strTime = strTime & " And Nvl(a.ԤԼ����,0) <> 1 "
            End If
        Else
            '�ò������ȡ��������Ӧ��ʱ���
            If mbytMode = 1 Or chkBooking.Value = 1 Then
                strTime = strTime & " And Nvl(a.ԤԼ����,0) <> 1 "
            Else
                strTime = _
                " And [5] Between Nvl(a.��ǰ�Һ�ʱ��, a.��ʼʱ��) And a.��ֹʱ��  "
            End If
        End If
               
        If mbytMode <> 1 Then
            If InStr(mstrPrivs, ";����Ѻ�;") = 0 And mbytMode = 0 Then
                strZero = "" & _
                "   And Exists(Select 1 From �շѼ�Ŀ" & _
                                " Where �շ�ϸĿid = d.Id And [5] Between ִ������ And Nvl(��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD'))" & _
                                " Group By �۸�ȼ� Having Nvl(Sum(�ּ�), 0) <> 0)"
            End If
            
            If InStr(mstrPrivs, ";���շѺ�;") = 0 And mbytMode = 0 Then
                strZero = strZero & _
                "   And Exists(Select 1 From �շѼ�Ŀ" & _
                                " Where �շ�ϸĿid = d.Id And [5] Between ִ������ And Nvl(��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD'))" & _
                                " Group By �۸�ȼ� Having Nvl(Sum(�ּ�), 0) = 0)"
            End If
            
            strSQL = strSQL & strZero
        End If
        
        Dim strWhere As String
        If mbytMode = 1 Or (mbytMode = 0 And chkBooking.Visible And chkBooking.Value = 1) Then
            'ԤԼ�Һ�
            'ԤԼ  �����Ƿ�����˷�ʱ��
            ' �ж��Ƿ����� ��ֻ���ڵ�ǰʱ��� �ų���
            If mcustomTime = t_��ͨ Then
                strSQL = strSQL & strTime
            Else
                strSQL = IIf(chkShowAll.Value = 0, strSQL & strTime, strSQL)
            End If
            strSQL = strSQL & IIf(chkShowAll.Value = 1, "", " And (a.��Լ�� > 0 Or a.��Լ�� Is Null)")
            strSQL = strSQL & IIf(chkShowAll.Value = 1, "", " And Nvl(a.ԤԼ����,0) <> 1 ")
            strSQL = strSQL & " And Sysdate + zl_Fun_GetAppointmentDays + Decode(Nvl(B.ԤԼ����," & gintԤԼ���� & "),0,15,Nvl(B.ԤԼ����," & gintԤԼ���� & ")" & ") > [5] "
        Else
            '�Һ�
            strSQL = strSQL & strTime
            If chkShowAll.Value = 1 Then
                '��ʾ������ű�
                If Trim(txt�ű�.Text) <> "" And Trim(txt�ű�.Text) <> "+" And ActiveControl Is txt�ű� And mblnReadBooking = False Then
                    If IsNumeric(Trim(txt�ű�.Text)) Then
                        str������IF = " And a.���� Like [2]"
                    Else
                        str������IF = " And (zlSpellCode(a.ҽ������) Like [2] or c.���� Like [2])"
                    End If
                End If
                
                strSQL = strSQL & vbNewLine & _
                    "Union All" & vbNewLine & _
                    "Select 0 As ��¼id, a.���� As �ű�, Null As ��������, a.����, a.����id, c.���� As ����, a.��Ŀid, d.���� As ��Ŀ, Null As ����ҽ��id, a.ҽ��id," & vbNewLine & _
                    "       Null As ����ҽ������, a.ҽ������ As ҽ��, 0 As �ѹ�, 0 As ��Լ, Null As �޺�, Null As ��Լ, Nvl(a.�Ƿ񽨲���, 0) As ����," & vbNewLine & _
                    "       Nvl(d.��Ŀ����, 0) As ����, Null As ����, 0 As ��ſ���, Null As �Ű�, a.Id As ��Դid, Null As ��ǰʱ��, Null As �Һ�ʱ��, Null As ����ʱ��," & vbNewLine & _
                    "       Null As �Ƿ��ռ, Null As ���￪ʼʱ��, Null As ������ֹʱ��, 0 As ��ʱ�� " & vbNewLine & _
                    "From �ٴ������Դ A, ���ű� C, �շ���ĿĿ¼ D" & vbNewLine & _
                    "Where a.����id = c.Id And a.��Ŀid = d.Id" & str������IF & vbNewLine & _
                    "      And Nvl(a.����ʱ��, To_Date('3000-01-01', 'YYYY-MM-DD')) > [5]" & vbNewLine & _
                    "      And Nvl(c.����ʱ��, To_Date('3000-01-01', 'YYYY-MM-DD')) > [5]" & vbNewLine & _
                    "      And Nvl(d.����ʱ��, To_Date('3000-01-01', 'YYYY-MM-DD')) > [5]" & vbNewLine & _
                    "      And Exists(Select 1 From �ٴ����ﰲ�� M, �ٴ������ N Where m.��Դid = a.Id And m.����id = n.Id And n.����ʱ�� Is Not Null)" & vbNewLine & _
                    "      And Not Exists(Select 1" & vbNewLine & _
                    "                     From �ٴ������¼" & vbNewLine & _
                    "                     Where ��Դid = a.Id And (�������� = [6] Or �������� = [8])" & vbNewLine & _
                    "                           And [5] Between ��ʼʱ�� And ��ֹʱ��" & vbNewLine & _
                    "                           And (��ʼʱ�� < Nvl(ͣ�￪ʼʱ��, ��ֹʱ��) Or ��ֹʱ�� > Nvl(ͣ����ֹʱ��, ��ʼʱ��))" & vbNewLine & _
                    "                           And Nvl(�Ƿ�����, 0) = 0 And Nvl(�Ƿ񷢲�, 0) = 1)" & vbNewLine
                
                If mbytMode <> 1 Then
                    If InStr(mstrPrivs, ";����Ѻ�;") = 0 Then
                        strZero = "" & _
                        "   And Exists(Select 1 From �շѼ�Ŀ" & _
                                " Where �շ�ϸĿid = d.Id And [5] Between ִ������ And Nvl(��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD'))" & _
                                " Group By �۸�ȼ� Having Nvl(Sum(�ּ�), 0) <> 0)"
                    End If
                    
                    If InStr(mstrPrivs, ";���շѺ�;") = 0 Then
                        strZero = strZero & _
                        "   And Exists(Select 1 From �շѼ�Ŀ" & _
                                " Where �շ�ϸĿid = d.Id And [5] Between ִ������ And Nvl(��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD'))" & _
                                " Group By �۸�ȼ� Having Nvl(Sum(�ּ�), 0) = 0)"
                    End If
                    strSQL = strSQL & strZero
                End If
            End If
        End If
              
        strSQL = strSQL & " Order by " & strSort
        
        Set mrsPlan = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, _
            UserInfo.����, Trim(txt�ű�.Text) & "%", mstrRoom, gstr�Һſ���ID, DateThis, CDate(Format(DateThis, "yyyy-MM-dd")), _
            CDate(Format(DateThis, "yyyy-MM-dd")) + 1 - 1 / 24 / 60 / 60, CDate(Format(DateThis - 1, "yyyy-MM-dd")), datNow, gdatRegistTime)

        mblnNotClick = True
        cboTime.Clear
        cboTime.AddItem "����"
        strSpan = ""
        Do While Not mrsPlan.EOF
            If InStr(strSpan, "," & mrsPlan!�Ű� & ",") = 0 Then
                strSpan = strSpan & "," & mrsPlan!�Ű� & ","
                cboTime.AddItem Nvl(mrsPlan!�Ű�)
            End If
            mrsPlan.MoveNext
        Loop
        cboTime.ListIndex = 0
        If mrsPlan.RecordCount <> 0 Then mrsPlan.MoveFirst
        mblnNotClick = False
    Else
       '�����ɸѡ
        If picBookingDate.Visible Or mbytMode = 1 Or (mbytMode = 2 Or mbytMode = 0 And mbytInState = 0 And mstrNoIn <> "") Then
            DateThis = dtpAppointmentDate.Value
        Else
            DateThis = zlDatabase.Currentdate
        End If
        If txt�ű�.Text = "" Then
            mrsPlan.Filter = IIf(cboTime.ListIndex = 0, "", "�Ű�='" & cboTime.Text & "'")
        Else
            mrsPlan.Filter = "�ű� like '" & txt�ű�.Text & "*'" & IIf(cboTime.ListIndex = 0, "", " And �Ű�='" & cboTime.Text & "'")
        End If
    End If
    
    With vsfPlan
        .Redraw = flexRDNone
        If Not mrsPlan.EOF Then
            '��ȡ���к�ԴID���Լ�������ű�
            Do While Not mrsPlan.EOF
                If zlCommFun.ActualLen(str��ԴIDs & "," & Nvl(mrsPlan!��ԴID)) > 4000 Then
                    str��ԴIDSQL = str��ԴIDSQL & vbNewLine & _
                        " Union All Select Column_Value From Table(f_Num2list('" & Mid(str��ԴIDs, 2) & "'))"
                    str��ԴIDs = ""
                End If
                str��ԴIDs = str��ԴIDs & "," & Nvl(mrsPlan!��ԴID)
                
                If Val(Nvl(mrsPlan!ID)) = 0 Then
                    If zlCommFun.ActualLen(str��������� & "," & Nvl(mrsPlan!�ű�)) > 4000 Then
                        str���������SQL = str���������SQL & vbNewLine & _
                            " Union All Select Column_Value From Table(f_Str2list('" & Mid(str���������, 2) & "'))"
                        str��������� = ""
                    End If
                    str��������� = str��������� & "," & Nvl(mrsPlan!�ű�)
                End If
                mrsPlan.MoveNext
            Loop
            If str��ԴIDs <> "" Then
                str��ԴIDSQL = str��ԴIDSQL & vbNewLine & _
                    " Union All Select Column_Value From Table(f_Num2list('" & Mid(str��ԴIDs, 2) & "'))"
            End If
            If str��������� <> "" Then
                str���������SQL = str���������SQL & vbNewLine & _
                    " Union All Select Column_Value From Table(f_Str2list('" & Mid(str���������, 2) & "'))"
            End If
            If str��ԴIDSQL <> "" Then str��ԴIDSQL = Mid(str��ԴIDSQL, 14)
            If str���������SQL <> "" Then str���������SQL = Mid(str���������SQL, 14)
            
            If str���������SQL <> "" Then
                str������SQL = "Select Count(1) As ����, �ű�" & vbNewLine & _
                                "From ���˹Һż�¼" & vbNewLine & _
                                "Where �����¼ID Is Null And ��¼���� = 1 And ��¼״̬ = 1 And ����ʱ�� >= [1] And ����ʱ�� <= [2]" & vbNewLine & _
                                "      And �ű� In (" & str���������SQL & ")" & vbNewLine & _
                                "Group By �ű�"
                Set rs������ = zlDatabase.OpenSQLRecord(str������SQL, Me.Caption, CDate(Format(DateThis, "yyyy-MM-dd")), CDate(Format(DateThis, "yyyy-MM-dd")) + 1 - 1 / 24 / 60 / 60)
            End If
            mrsPlan.MoveFirst
            If mbytMode = 0 Then
                dat��ʼʱ�� = CDate(Format(DateThis, "yyyy-mm-dd")) - 1
                dat����ʱ�� = CDate(Format(DateThis, "yyyy-mm-dd")) + 5
            Else
                datNow = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd"))
                If CDate(Format(DateThis, "yyyy-mm-dd")) - datNow >= 3 Then
                    dat��ʼʱ�� = CDate(Format(DateThis, "yyyy-mm-dd")) - 3
                    dat����ʱ�� = CDate(Format(DateThis, "yyyy-mm-dd")) + 3
                Else
                    dat��ʼʱ�� = CDate(Format(datNow, "yyyy-mm-dd")) - 1
                    dat����ʱ�� = CDate(Format(datNow, "yyyy-mm-dd")) + 5
                End If
            End If
            strDays = "Select ��Դid, To_Char(��������,'DD') As ����, To_Char(��������, 'D') As ����, �ϰ�ʱ��" & vbNewLine & _
                    "From �ٴ������¼" & vbNewLine & _
                    "Where ��Դid In (" & str��ԴIDSQL & ") And �������� Between [1] And [2]" & vbNewLine & _
                    "Order By ����"
            Set rsDays = zlDatabase.OpenSQLRecord(strDays, Me.Caption, dat��ʼʱ��, dat����ʱ��)
            If Not rsDays.EOF Then
                Do While Not rsDays.EOF
                    Select Case Val(Nvl(rsDays!����))
                        Case 1
                            .TextMatrix(i, .ColIndex("��")) = "��" & vbCrLf & rsDays!���� & ""
                        Case 2
                            .TextMatrix(i, .ColIndex("һ")) = "һ" & vbCrLf & rsDays!���� & ""
                        Case 3
                            .TextMatrix(i, .ColIndex("��")) = "��" & vbCrLf & rsDays!���� & ""
                        Case 4
                            .TextMatrix(i, .ColIndex("��")) = "��" & vbCrLf & rsDays!���� & ""
                        Case 5
                            .TextMatrix(i, .ColIndex("��")) = "��" & vbCrLf & rsDays!���� & ""
                        Case 6
                            .TextMatrix(i, .ColIndex("��")) = "��" & vbCrLf & rsDays!���� & ""
                        Case 7
                            .TextMatrix(i, .ColIndex("��")) = "��" & vbCrLf & rsDays!���� & ""
                    End Select
                    rsDays.MoveNext
                Loop
                rsDays.MoveFirst
            End If
            
            .ToolTipText = "�� " & mrsPlan.RecordCount & " ������"
            .Clear 1
            mblnChangeByCode = True
            .Rows = 2
            If mrsPlan.RecordCount <> 0 Then
                .Rows = mrsPlan.RecordCount + 1
            End If
            mblnChangeByCode = False

            mrsPlan.MoveFirst
            For i = 1 To mrsPlan.RecordCount
                .RowData(i) = IIf(mrsPlan!���� = 1, -1, 1) * mrsPlan!����ID
                .TextMatrix(i, .ColIndex("IDS")) = mrsPlan!ID & "," & mrsPlan!��ĿID & "," & IIf(IsNull(mrsPlan!ҽ��ID), 0, mrsPlan!ҽ��ID)
                .Cell(flexcpData, i, .ColIndex("IDS")) = Nvl(mrsPlan!ID)
                .TextMatrix(i, .ColIndex("����")) = IIf(IsNull(mrsPlan!����), "", mrsPlan!����)
                .TextMatrix(i, .ColIndex("�ű�")) = mrsPlan!�ű�
                .TextMatrix(i, .ColIndex("����")) = mrsPlan!����
                .TextMatrix(i, .ColIndex("��Ŀ")) = mrsPlan!��Ŀ
                .TextMatrix(i, .ColIndex("��������")) = Format(Nvl(mrsPlan!��������), "yyyy-mm-dd")
                .TextMatrix(i, .ColIndex("ҽ��")) = Nvl(mrsPlan!ҽ��)
                .TextMatrix(i, .ColIndex("��Լ")) = Nvl(mrsPlan!��Լ)
                .TextMatrix(i, .ColIndex("��Լ")) = Nvl(mrsPlan!��Լ)
                If Not rs������ Is Nothing And Val(Nvl(mrsPlan!ID)) = 0 Then
                    rs������.Filter = "�ű�=" & "'" & mrsPlan!�ű� & "'"
                    If Not rs������.EOF Then
                        .TextMatrix(i, .ColIndex("�ѹ�")) = Nvl(rs������!����)
                    End If
                Else
                    .TextMatrix(i, .ColIndex("�ѹ�")) = Nvl(mrsPlan!�ѹ�)
                End If
                .TextMatrix(i, .ColIndex("�޺�")) = Nvl(mrsPlan!�޺�)
                If Val(.TextMatrix(i, .ColIndex("�޺�"))) <> 0 And mbytMode = 0 Then
                    .TextMatrix(i, .ColIndex("ʣ��")) = Val(.TextMatrix(i, .ColIndex("�޺�"))) - Val(.TextMatrix(i, .ColIndex("�ѹ�")))
                    intSurplusTotal = intSurplusTotal + Val(.TextMatrix(i, .ColIndex("ʣ��")))
                End If
                .TextMatrix(i, .ColIndex("��ǰʱ��")) = Format(mrsPlan!��ǰʱ��, "yyyy-MM-dd hh:mm:ss")
                .TextMatrix(i, .ColIndex("�Һ�ʱ��")) = Format(mrsPlan!�Һ�ʱ��, "yyyy-MM-dd hh:mm:ss")
                .TextMatrix(i, .ColIndex("����ʱ��")) = Format(mrsPlan!����ʱ��, "yyyy-MM-dd hh:mm:ss")
                .TextMatrix(i, .ColIndex("ʱ��")) = Nvl(mrsPlan!�Ű�)
                .TextMatrix(i, .ColIndex("�Ƿ��ռ")) = Nvl(mrsPlan!�Ƿ��ռ)
                If Nvl(mrsPlan!����ҽ������) <> "" Then
                    .TextMatrix(i, .ColIndex("����ҽ��")) = ""
                    .Cell(flexcpData, i, .ColIndex("����ҽ��")) = Nvl(mrsPlan!����ҽ������) & "(" & Format(Nvl(mrsPlan!���￪ʼʱ��), "hh:mm") & "-" & Format(Nvl(mrsPlan!������ֹʱ��), "hh:mm") & ")"
                    .TextMatrix(i, .ColIndex("���￪ʼʱ��")) = Format(mrsPlan!���￪ʼʱ��, "yyyy-MM-dd hh:mm:ss")
                    .TextMatrix(i, .ColIndex("������ֹʱ��")) = Format(mrsPlan!������ֹʱ��, "yyyy-MM-dd hh:mm:ss")
                    .TextMatrix(i, .ColIndex("����ҽ������")) = Nvl(mrsPlan!����ҽ������)
                    .TextMatrix(i, .ColIndex("����ҽ��ID")) = Nvl(mrsPlan!����ҽ��id)
                End If
                
                rsDays.Filter = "��Դid=" & Val(mrsPlan!��ԴID)
                Do While Not rsDays.EOF
                    Select Case Val(Nvl(rsDays!����))
                    Case 1
                        If Nvl(rsDays!�ϰ�ʱ��) = Nvl(mrsPlan!�Ű�) Then
                            .TextMatrix(i, .ColIndex("��")) = Left(Nvl(mrsPlan!�Ű�), 1)
                            .Cell(flexcpData, i, .ColIndex("��")) = Nvl(mrsPlan!�Ű�)
                        Else
                            If .Cell(flexcpData, i, .ColIndex("��")) <> Nvl(mrsPlan!�Ű�) Then
                                If .TextMatrix(i, .ColIndex("��")) = "" Then
                                    .TextMatrix(i, .ColIndex("��")) = Left(Nvl(rsDays!�ϰ�ʱ��), 1)
                                    .Cell(flexcpData, i, .ColIndex("��")) = Nvl(rsDays!�ϰ�ʱ��)
                                Else
                                    .TextMatrix(i, .ColIndex("��")) = .TextMatrix(i, .ColIndex("��")) & "/" & Left(Nvl(rsDays!�ϰ�ʱ��), 1)
                                    .Cell(flexcpData, i, .ColIndex("��")) = .Cell(flexcpData, i, .ColIndex("��")) & "/" & Nvl(rsDays!�ϰ�ʱ��)
                                End If
                            End If
                        End If
                    Case 2
                        If Nvl(rsDays!�ϰ�ʱ��) = Nvl(mrsPlan!�Ű�) Then
                            .TextMatrix(i, .ColIndex("һ")) = Left(Nvl(mrsPlan!�Ű�), 1)
                            .Cell(flexcpData, i, .ColIndex("һ")) = Nvl(mrsPlan!�Ű�)
                        Else
                            If .Cell(flexcpData, i, .ColIndex("һ")) <> Nvl(mrsPlan!�Ű�) Then
                                If .TextMatrix(i, .ColIndex("һ")) = "" Then
                                    .TextMatrix(i, .ColIndex("һ")) = Left(Nvl(rsDays!�ϰ�ʱ��), 1)
                                    .Cell(flexcpData, i, .ColIndex("һ")) = Nvl(rsDays!�ϰ�ʱ��)
                                Else
                                    .TextMatrix(i, .ColIndex("һ")) = .TextMatrix(i, .ColIndex("һ")) & "/" & Left(Nvl(rsDays!�ϰ�ʱ��), 1)
                                    .Cell(flexcpData, i, .ColIndex("һ")) = .Cell(flexcpData, i, .ColIndex("һ")) & "/" & Nvl(rsDays!�ϰ�ʱ��)
                                End If
                            End If
                        End If
                    Case 3
                        If Nvl(rsDays!�ϰ�ʱ��) = Nvl(mrsPlan!�Ű�) Then
                            .TextMatrix(i, .ColIndex("��")) = Left(Nvl(mrsPlan!�Ű�), 1)
                            .Cell(flexcpData, i, .ColIndex("��")) = Nvl(mrsPlan!�Ű�)
                        Else
                            If .Cell(flexcpData, i, .ColIndex("��")) <> Nvl(mrsPlan!�Ű�) Then
                                If .TextMatrix(i, .ColIndex("��")) = "" Then
                                    .TextMatrix(i, .ColIndex("��")) = Left(Nvl(rsDays!�ϰ�ʱ��), 1)
                                    .Cell(flexcpData, i, .ColIndex("��")) = Nvl(rsDays!�ϰ�ʱ��)
                                Else
                                    .TextMatrix(i, .ColIndex("��")) = .TextMatrix(i, .ColIndex("��")) & "/" & Left(Nvl(rsDays!�ϰ�ʱ��), 1)
                                    .Cell(flexcpData, i, .ColIndex("��")) = .Cell(flexcpData, i, .ColIndex("��")) & "/" & Nvl(rsDays!�ϰ�ʱ��)
                                End If
                            End If
                        End If
                    Case 4
                        If Nvl(rsDays!�ϰ�ʱ��) = Nvl(mrsPlan!�Ű�) Then
                            .TextMatrix(i, .ColIndex("��")) = Left(Nvl(mrsPlan!�Ű�), 1)
                            .Cell(flexcpData, i, .ColIndex("��")) = Nvl(mrsPlan!�Ű�)
                        Else
                            If .Cell(flexcpData, i, .ColIndex("��")) <> Nvl(mrsPlan!�Ű�) Then
                                If .TextMatrix(i, .ColIndex("��")) = "" Then
                                    .TextMatrix(i, .ColIndex("��")) = Left(Nvl(rsDays!�ϰ�ʱ��), 1)
                                    .Cell(flexcpData, i, .ColIndex("��")) = Nvl(rsDays!�ϰ�ʱ��)
                                Else
                                    .TextMatrix(i, .ColIndex("��")) = .TextMatrix(i, .ColIndex("��")) & "/" & Left(Nvl(rsDays!�ϰ�ʱ��), 1)
                                    .Cell(flexcpData, i, .ColIndex("��")) = .Cell(flexcpData, i, .ColIndex("��")) & "/" & Nvl(rsDays!�ϰ�ʱ��)
                                End If
                            End If
                        End If
                    Case 5
                        If Nvl(rsDays!�ϰ�ʱ��) = Nvl(mrsPlan!�Ű�) Then
                            .TextMatrix(i, .ColIndex("��")) = Left(Nvl(mrsPlan!�Ű�), 1)
                            .Cell(flexcpData, i, .ColIndex("��")) = Nvl(mrsPlan!�Ű�)
                        Else
                            If .Cell(flexcpData, i, .ColIndex("��")) <> Nvl(mrsPlan!�Ű�) Then
                                If .TextMatrix(i, .ColIndex("��")) = "" Then
                                    .TextMatrix(i, .ColIndex("��")) = Left(Nvl(rsDays!�ϰ�ʱ��), 1)
                                    .Cell(flexcpData, i, .ColIndex("��")) = Nvl(rsDays!�ϰ�ʱ��)
                                Else
                                    .TextMatrix(i, .ColIndex("��")) = .TextMatrix(i, .ColIndex("��")) & "/" & Left(Nvl(rsDays!�ϰ�ʱ��), 1)
                                    .Cell(flexcpData, i, .ColIndex("��")) = .Cell(flexcpData, i, .ColIndex("��")) & "/" & Nvl(rsDays!�ϰ�ʱ��)
                                End If
                            End If
                        End If
                    Case 6
                        If Nvl(rsDays!�ϰ�ʱ��) = Nvl(mrsPlan!�Ű�) Then
                            .TextMatrix(i, .ColIndex("��")) = Left(Nvl(mrsPlan!�Ű�), 1)
                            .Cell(flexcpData, i, .ColIndex("��")) = Nvl(mrsPlan!�Ű�)
                        Else
                            If .Cell(flexcpData, i, .ColIndex("��")) <> Nvl(mrsPlan!�Ű�) Then
                                If .TextMatrix(i, .ColIndex("��")) = "" Then
                                    .TextMatrix(i, .ColIndex("��")) = Left(Nvl(rsDays!�ϰ�ʱ��), 1)
                                    .Cell(flexcpData, i, .ColIndex("��")) = Nvl(rsDays!�ϰ�ʱ��)
                                Else
                                    .TextMatrix(i, .ColIndex("��")) = .TextMatrix(i, .ColIndex("��")) & "/" & Left(Nvl(rsDays!�ϰ�ʱ��), 1)
                                    .Cell(flexcpData, i, .ColIndex("��")) = .Cell(flexcpData, i, .ColIndex("��")) & "/" & Nvl(rsDays!�ϰ�ʱ��)
                                End If
                            End If
                        End If
                    Case 7
                        If Nvl(rsDays!�ϰ�ʱ��) = Nvl(mrsPlan!�Ű�) Then
                            .TextMatrix(i, .ColIndex("��")) = Left(Nvl(mrsPlan!�Ű�), 1)
                            .Cell(flexcpData, i, .ColIndex("��")) = Nvl(mrsPlan!�Ű�)
                        Else
                            If .Cell(flexcpData, i, .ColIndex("��")) <> Nvl(mrsPlan!�Ű�) Then
                                If .TextMatrix(i, .ColIndex("��")) = "" Then
                                    .TextMatrix(i, .ColIndex("��")) = Left(Nvl(rsDays!�ϰ�ʱ��), 1)
                                    .Cell(flexcpData, i, .ColIndex("��")) = Nvl(rsDays!�ϰ�ʱ��)
                                Else
                                    .TextMatrix(i, .ColIndex("��")) = .TextMatrix(i, .ColIndex("��")) & "/" & Left(Nvl(rsDays!�ϰ�ʱ��), 1)
                                    .Cell(flexcpData, i, .ColIndex("��")) = .Cell(flexcpData, i, .ColIndex("��")) & "/" & Nvl(rsDays!�ϰ�ʱ��)
                                End If
                            End If
                        End If
                    End Select
                    rsDays.MoveNext
                Loop
                .TextMatrix(i, .ColIndex("����")) = IIf(mrsPlan!���� = 1, "��", "")
                .TextMatrix(i, .ColIndex("����")) = Nvl(mrsPlan!����)
                .TextMatrix(i, .ColIndex("��ſ���")) = IIf(mrsPlan!��ſ��� = 1, "��", "")
                .TextMatrix(i, .ColIndex("��¼ID")) = Nvl(mrsPlan!ID)
                .TextMatrix(i, .ColIndex("��ԴID")) = Nvl(mrsPlan!��ԴID)
                .TextMatrix(i, .ColIndex("��ʱ��")) = Val(Nvl(mrsPlan!��ʱ��))
                .Cell(flexcpData, i, .ColIndex("�ű�")) = ""
                If InStr(mstrPrivs, ";��ʱ�Һ�;") = 0 And chkShowAll.Value = 1 Then
                    If Val(Nvl(mrsPlan!ID)) = 0 Or DateThis < CDate(IIf(.TextMatrix(i, .ColIndex("��ǰʱ��")) = "", IIf(.TextMatrix(i, .ColIndex("�Һ�ʱ��")) = "", "3000-01-01", .TextMatrix(i, .ColIndex("�Һ�ʱ��"))), .TextMatrix(i, .ColIndex("��ǰʱ��")))) Then
                        .Cell(flexcpData, i, .ColIndex("�ű�")) = "1"
                        .Cell(flexcpForeColor, i, GetCol("IDS"), i, .Cols - 1) = &H8000000C
                    End If
                End If
                
                If mrsPlan!�ű� = txt�ű�.Text And k = 0 And (Nvl(mrsPlan!ID) = mlng��¼ID Or mlng��¼ID = 0) Then k = i
                '���� 43847
                If k = 0 And mrsPlan!�ű� = mstrPreNO And (mSortType = by�ű� Or txt�ű�.Text = "") Then k = i
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
        '70193:������,2014-2-18,�ű��Զ���λ���������
        If vsfPlan.Row = 1 Then
            vsfPlan.Select 1, 1
            If txt�ű�.Visible And txt�ű�.Enabled Then txt�ű�.SetFocus
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
    '���ܣ��ش�
    '����:�ش�ɹ�,����true,���򷵻�False
    '���ƣ����˺�
    '���ڣ�2010-07-02 10:49:06
    '˵������Ҫ�������������
    '------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, str���� As String, str�Ա� As String, str�������� As String
    Dim lng����ID As Long, lng����ID As Long, intInsure As Integer
    Dim strNO As String, blnVirtualPrint As Boolean
    
    If cboNO.Tag = "" Then
        MsgBox "δ����Һŵ��ݣ������ش�", vbInformation, gstrSysName
        Exit Function
    End If
    strNO = cboNO.Tag
    If CheckBillExistReplenishData(strNO) Then
        MsgBox "ѡ��ĹҺż�¼������ҽ��������㣬�������ش������", vbInformation, gstrSysName
        Exit Function
    End If
    lng����ID = GetBill����ID(strNO, 4, lng����ID)
    intInsure = ExistInsure(strNO)
    If intInsure <> 0 Then
        blnVirtualPrint = gclsInsure.GetCapability(supportҽ���ӿڴ�ӡƱ��, lng����ID, intInsure)
    End If
    
    Dim blnStartFactUseType  As Boolean, strUseType As String
    If gblnSharedInvoice Then
        blnStartFactUseType = zlStartFactUseType("1")
        If blnStartFactUseType Then
            strUseType = zl_GetInvoiceUserType(lng����ID, 0, intInsure)
        End If
    End If
    
    
    If txtPatientPrint.Visible Then
        If txtPatientPrint.Text = "" Then
            MsgBox "����Ϊ��,������������", vbInformation, gstrSysName
            If txtPatientPrint.Enabled Then txtPatientPrint.SetFocus
            Exit Function
        End If
        str���� = Trim(txt����.Text): str�Ա� = NeedName(cbo�Ա�.Text)
        If IsNumeric(str����) Then str���� = str���� & cbo���䵥λ.Text
        If txtPatient.Text <> txtPatientPrint.Text Or mstr���� & mstr���䵥λ <> str���� Or mstr�Ա� <> str�Ա� Then
            If zlExistOperationData(Val(txtPatientPrint.Tag), cboNO.Tag) Then
                MsgBox "ע��:" & vbCrLf & "�ò����Ѿ�����ҽ��ҵ������,���ܵ������˵Ļ�����Ϣ,���ڡ�������Ϣ�����е���!" & vbCrLf & "���ȷ����ָ��޸ĵĲ�����Ϣ��", vbOKOnly + vbDefaultButton1, gstrSysName
                txt����.Text = mstr����
                If mstr���䵥λ <> "" Then cbo���䵥λ.ListIndex = cbo.FindIndex(cbo���䵥λ, mstr���䵥λ, True): cbo���䵥λ.Visible = True: txt����.Width = 600
                str���� = Trim(txt����.Text): str�Ա� = NeedName(cbo�Ա�.Text)
                If IsNumeric(str����) Then str���� = str���� & cbo���䵥λ.Text
                cbo�Ա�.ListIndex = cbo.FindIndex(cbo�Ա�, mstr�Ա�, True)
                txtPatient.Text = mstr����
                Exit Function
            End If
            str�������� = "NULL"
            '35544
            If str���� <> mstr���� Then
                If IsNumeric(CStr(txt����.Text)) Then
                    str�������� = ReCalcBirth(txt����.Text, cbo���䵥λ.Text)
                    If IsDate(str��������) = False Then
                        str�������� = "NULL"
                    Else
                        str�������� = "to_date('" & str�������� & "','yyyy-mm-dd')"
                    End If
                End If
            End If
            'Zl_���˷��ü�¼_Update
            strSQL = "Zl_���˷��ü�¼_Update("
            '  No_In       ������ü�¼.NO%Type,
            strSQL = strSQL & "'" & strNO & "',"
            '  ��¼����_In ������ü�¼.��¼����%Type,
            strSQL = strSQL & "" & 4 & ","
            '  ������_In   ������ü�¼.������%Type,
            strSQL = strSQL & "" & "Null" & ","
            '  ����ʱ��_In ������ü�¼.����ʱ��%Type,
            strSQL = strSQL & "" & "Null" & ","
            '  ����_In     ������ü�¼.����%Type := Null,
            strSQL = strSQL & "'" & txtPatientPrint.Text & "',"
            '  ��Դ_In     Integer := 1,
            strSQL = strSQL & "" & 1 & ","
            '  ����_In     ������ü�¼.����%Type := Null,
            strSQL = strSQL & "" & IIf(str���� = "", "NULL", "'" & str���� & "'") & ","
            '  �Ա�_In     ������ü�¼.�Ա�%Type := Null
            strSQL = strSQL & "" & IIf(str�Ա� = "", "NULL", "'" & str�Ա� & "'") & ","
            '  ��������_In ������Ϣ.��������%Type := Null
            strSQL = strSQL & "" & str�������� & ")"
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        End If
    End If
   '����:53037
    If Not RePrintBill(Me, 3, strNO, lng����ID, intInsure, blnVirtualPrint, strUseType, True) Then Exit Function

    zlRePrintRegistered = True
End Function

Private Function GetTotal(ByVal strNO As String) As Double
    Dim strSQL As String, rsTmp As ADODB.Recordset
    strSQL = "Select Sum(���ʽ��) As �ܽ�� From ������ü�¼ Where No = [1] And ��¼���� = 4"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
    If Not rsTmp.EOF Then GetTotal = Val(Nvl(rsTmp!�ܽ��))
End Function


Private Function zlExcuteDelRegistered() As Boolean
    '------------------------------------ ------------------------------------------------------------------------------------
    '���ܣ��Һ��˺�
    '���أ��˺ųɹ�,����true,���򷵻�False
    '���ƣ����˺�
    '���ڣ�2010-07-02 10:53:29
    '˵���������������ʱ,���ϴ˹���
    '------------------------------------------------------------------------------------------------------------------------
    Dim rsOneCard1 As ADODB.Recordset, rsOneCard2 As ADODB.Recordset, objICCard As Object
    Dim blnPromptClear As Boolean, strSQL As String, strNO As String, lngCard����ID As Long
    Dim strSQLCard As String, intMsgReturn As Integer, bln�˷��ش� As Boolean, blnTrans As Boolean
    Dim bytTogetherDo As Byte, dblTotal As Double                            '0-�޸��Ӳ���,1-ɾ�������
    Dim strAdvance  As String, strCardNo As String, lng����ID As Long
    Dim blnNotCommit As Boolean
    Dim Curdate As Date '�����:56599
    Dim str���� As String '�����:56599
    Dim str���� As String '�����:56599
    Dim rsҽ�ƿ���� As Recordset '�����:56599
    Dim cllPro As Collection, cllBillBalance As Collection, dblThreeMoney As Double
    Dim cllUpdate As Collection, cllThreeIns As Collection, strErrMsg As String
    Dim byt�˷����� As Byte '0-ȫ�� 1-�˹Һŷ� 2-�˲�����
    Dim i As Long, curMoney As Currency, dblԤ��֧�� As Double
    Dim curChkMoney As Currency
    Dim blnCardReprint As Boolean
    Dim objCard As Card, str���㷽ʽ As String, str�ֽ� As String, dbl�ֽ� As Double
    Dim strBalance As String, strDelCardNo As String, strԭ���㷽ʽ As String
    Dim strInvoice As String, lng����ID As Long, lng����ID As Long
    Dim bln���� As Boolean, bln���� As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim blnԭ���� As Boolean, strTemp As String, strBackNote As String
    Dim blnҽ��ԭ���� As Boolean
    Dim dblԤ�� As Double
    Dim rsInvoice As ADODB.Recordset
    Dim strBackInvoice As String, blnReprint As Boolean
    Dim dblCheckThreeMoney As Double
    
    Set cllPro = New Collection
 
    
    strNO = cboNO.Tag
    If strNO = "" Then
        MsgBox "δ����Һŵ��ݣ������˺ţ�", vbInformation, gstrSysName
        Exit Function
    End If
    If CheckBillExistReplenishData(strNO) Then
        MsgBox "ѡ��ĹҺż�¼������ҽ��������㣬����������˺Ų�����", vbInformation, gstrSysName
        Exit Function
    End If
    If cbo��ע.Text <> "" And cbo��ע.Tag = "" And mbln�˺�ԭ�� And cbo��ע.Enabled And cbo��ע.Visible Then
        If cbo��ע.Text <> mstrԭժҪ Then
            MsgBox "����ժҪ��ѡ����ȷ���˺�ԭ��!", vbInformation, gstrSysName
            cbo��ע.SetFocus
            Exit Function
        End If
    End If
    '68991
    lng����ID = GetBill����ID(strNO, 4, lng����ID, bln����)
    If zlCheckIsAllowBackSN(strNO, bln����, bln����) = False Then Exit Function
    
    If Not bln���� Then
        '����:51527
        Call zlReadRegThreeBalance(strNO, cllBillBalance, objCard)
        blnԭ���� = True
        If Not objCard Is Nothing Then
            blnԭ���� = False
            For i = 1 To vsfPay.Rows - 1
                If vsfPay.RowData(i) = 1 Then
                    strTemp = vsfPay.TextMatrix(i, 0)
                End If
                If Val(vsfPay.TextMatrix(i, 4)) = objCard.�ӿ���� Then
                    blnԭ���� = True
                End If
            Next i
        End If
        If blnԭ���� = False Then
            str���㷽ʽ = strTemp
            strԭ���㷽ʽ = objCard.���㷽ʽ
            Set mCurCardPay.objCard = Nothing
            mCurCardPay.lngҽ�ƿ����ID = 0
        End If
    Else
        str���㷽ʽ = ""
    End If
    
    
    blnPromptClear = True
    If vsfMoney.Tag = "����" Then   '����ҺŷѺͿ���û�з�����ǰ��
        If MsgBox("��ǰҪ�˺ŵĵ��ݷ����а������￨��,��һ���˷�!" & vbCrLf & _
            "��ȷʵҪ�����˺���?", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
           cboNO.Text = "": cboNO.SetFocus: Exit Function
        End If
    Else
        strDelCardNo = ExistCardFee(strNO, lngCard����ID, str����)
        If strDelCardNo <> "" Then
            '�����:56599
            If str���� <> "" Then
                '113613�����ϴ���2018/1/18���˿�ʱ��鵱ǰ���Ƿ������˿�
                strSQL = "Select Nvl(�Ƿ�����,0) As �Ƿ�����,zl1_EX_ReFundCard_Check([1],[2],A.�����ID,[3]) as ��֤" & _
                "           From ����ҽ�ƿ���Ϣ A,ҽ�ƿ���� B " & _
                "           Where A.����=[3] And A.�����ID =B.ID "
                Set rsҽ�ƿ���� = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngModul, lng����ID, str����)
                If rsҽ�ƿ����.EOF = False Then
                    If Nvl(rsҽ�ƿ����!��֤) <> "" Then
                        If Not objCard Is Nothing Then
                            If mCurCardPay.lngҽ�ƿ����ID <> 0 And mCurCardPay.bln���ѿ� = False And objCard.�Ƿ�ȫ�� Then
                                MsgBox Nvl(rsҽ�ƿ����!��֤) & "�����ܵ����˹Һŷѣ�", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName
                                cboNO.Text = "": cboNO.SetFocus: Exit Function
                            End If
                        End If
                        If MsgBox(Nvl(rsҽ�ƿ����!��֤) & "���Ƿ񵥶��˹Һŷѣ�", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                            cboNO.Text = "": cboNO.SetFocus: Exit Function
                        End If
                        str���� = "���˺�"
                    ElseIf rsҽ�ƿ����!�Ƿ����� = 0 Then 'Ժ�⿨
                        str���� = zl9ComLib.zlCommFun.ShowMsgbox(gstrSysName, "����:" & str���� & "��ΪԺ�⿨��,��ѡ���˿���ȡ���󶨲���", "�˿�,ȡ����", Me, vbQuestion)
                    End If
                End If
            End If
            
            '�����:56599
            If str���� <> "" Then
                 Select Case str����
                    Case "�˿�"
                        'Zl_ҽ�ƿ���¼_Delete
                        strSQLCard = "Zl_ҽ�ƿ���¼_Delete("
                        '      ���ݺ�_In     סԺ���ü�¼.No%Type,
                        strSQLCard = strSQLCard & "'" & strDelCardNo & "',"
                        '      ����Ա���_In סԺ���ü�¼.����Ա���%Type,
                        strSQLCard = strSQLCard & "'" & UserInfo.��� & "',"
                        '      ����Ա����_In סԺ���ü�¼.����Ա����%Type,
                        strSQLCard = strSQLCard & "'" & UserInfo.���� & "')"
                    Case "ȡ����"
                        Curdate = zlDatabase.Currentdate
                        'Zl_ҽ�ƿ��䶯_Insert
                         strSQLCard = "Zl_ҽ�ƿ��䶯_Insert("
                        '      �䶯����_In   Number,
                        '��������=1-����(��11�󶨿�);2-����;3-����(13-����ͣ��);4-�˿�(��14ȡ����); ��-�������(ֻ��¼);6-��ʧ(16ȡ����ʧ)
                        strSQLCard = strSQLCard & "" & 14 & ","
                        '      ����id_In     סԺ���ü�¼.����id%Type,
                        strSQLCard = strSQLCard & "" & lng����ID & ","
                        '      �����id_In   ����ҽ�ƿ���Ϣ.�����id%Type,
                        strSQLCard = strSQLCard & "" & mCurCardPay.lngҽ�ƿ����ID & ","
                        '      ԭ����_In     ����ҽ�ƿ���Ϣ.����%Type,
                        strSQLCard = strSQLCard & "NULL,"
                        '      ҽ�ƿ���_In   ����ҽ�ƿ���Ϣ.����%Type,
                        strSQLCard = strSQLCard & str���� & ","
                        '      �䶯ԭ��_In   ����ҽ�ƿ��䶯.�䶯ԭ��%Type,
                        strSQLCard = strSQLCard & "'ȡ�����Ű�',"
                        '      ����_In       ������Ϣ.����֤��%Type,
                        strSQLCard = strSQLCard & "NULL,"
                        '      ����Ա����_In סԺ���ü�¼.����Ա����%Type,
                        strSQLCard = strSQLCard & "NULL,"
                        '      �䶯ʱ��_In   סԺ���ü�¼.�Ǽ�ʱ��%Type,
                        strSQLCard = strSQLCard & "to_date('" & Format(Curdate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
                        '      Ic����_In     ������Ϣ.Ic����%Type := Null,
                        strSQLCard = strSQLCard & "NULL,"
                        '      ��ʧ��ʽ_In   ����ҽ�ƿ��䶯.��ʧ��ʽ%Type := Null
                        strSQLCard = strSQLCard & "NULL)"
                 End Select
            Else
                If str���� = "���˺�" Then
                    intMsgReturn = vbNo
                Else
                     '116278:���ϴ�,2017/12/15����֧�ֲ����˵����������˺ű���ͬʱ�˿�
                    If Not objCard Is Nothing Then
                        If mCurCardPay.lngҽ�ƿ����ID <> 0 And mCurCardPay.bln���ѿ� = False And objCard.�Ƿ�ȫ�� Then
                            intMsgReturn = MsgBox("�ò��˹Һ�ʱ������,�˺ű���ͬʱ�˿�,�Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName)
                            If intMsgReturn = vbNo Then Exit Function
                        Else
                            intMsgReturn = MsgBox("�ò��˹Һ�ʱ������,�˺�ͬʱ�˿���", vbQuestion + vbYesNoCancel + vbDefaultButton2, gstrSysName)
                        End If
                    Else
                        intMsgReturn = MsgBox("�ò��˹Һ�ʱ������,�˺�ͬʱ�˿���", vbQuestion + vbYesNoCancel + vbDefaultButton2, gstrSysName)
                    End If
                End If
                If intMsgReturn = vbYes Then
                    strSQLCard = "zl_ҽ�ƿ���¼_DELETE('" & strDelCardNo & "','" & UserInfo.��� & "','" & UserInfo.���� & "')"
                ElseIf intMsgReturn = vbNo Then
                    strSQLCard = ""
                    lngCard����ID = 0
                    bln�˷��ش� = gbln�˷��ش�
                    blnCardReprint = gbln�˷��ش�
                    blnPromptClear = False
                Else
                    Exit Function
                End If
            End If
        End If
    End If
    
    '����:51527
    dblThreeMoney = 0
    If mCurCardPay.lngҽ�ƿ����ID <> 0 Then
        dblThreeMoney = zlGetRegThreeMoney(lng����ID, lngCard����ID, cllBillBalance)
    End If
    dblCheckThreeMoney = zlGetRegThreeMoney(lng����ID, lngCard����ID, cllBillBalance)
    
    bytTogetherDo = 0
    'ȫ��
    If mintCancel = 0 And mbln������ = True Then
        If Not (mbln���������� And chk������.Value = 0) And Not (mbln���ӷ� And chkExtra.Value = 0) Then
            '����Һŵ��ĵǼ�����-������Ϣ�ĵǼ������ڹҺŵ���Ч����֮��,����ʾ�Ƿ�ɾ�������   txt����ʱ��
            If txt�����.Text <> "" And blnPromptClear Then
                If Check�Һ�ʱ����(strNO, txt����ʱ��.Text) Then
                    Select Case gbyt���������Ϣ    '35176
                    Case 0  '�����
                    Case 1  '���
                           bytTogetherDo = 1
                    Case 2  '��ʾ���
                        If MsgBox("�˺ź�Ҫ�����ò�����ص��������Ϣ��!", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                           bytTogetherDo = 1
                        End If
                    End Select
                End If
            End If
        End If
    End If
    
    dblԤ��֧�� = 0
    For i = 1 To vsfPay.Rows - 1
        If vsfPay.RowData(i) = 0 And vsfPay.TextMatrix(i, 0) <> "" Then
            dblԤ��֧�� = dblԤ��֧�� + Val(vsfPay.TextMatrix(i, 1))
        End If
    Next i
    
    '����˷��漰Ԥ����,����Ҫˢ����֤
    If gbytԤ����˷��鿨 <> 0 And dblԤ��֧�� <> 0 Then
        If mrsBill.RecordCount <> 0 Then mrsBill.MoveFirst
        If Not zlDatabase.PatiIdentify(Me, glngSys, Nvl(mrsBill!����ID, 0), dblԤ��֧��, _
                            mlngModul, 1, IDKind.GetCurCard.�ӿ����, , True, , , (gbytԤ����˷��鿨 = 2)) Then Exit Function
    End If
    
    Select Case mintCancel
    Case 0
        If mbln������ Then
            If ((mbln���������� And chk������.Value = 1) Or mbln���������� = False) And ((mbln���ӷ� And chkExtra.Value = 1) Or mbln���ӷ� = False) Then
                '�����˷ѽ�����.
                For i = 1 To vsfMoney.Rows - 1
                    curMoney = Val(vsfMoney.TextMatrix(i, 2)) + curMoney
                Next
                
                byt�˷����� = 0
            ElseIf ((mbln���������� And chk������.Value = 0) Or mbln���������� = False) And ((mbln���ӷ� And chkExtra.Value = 0) Or mbln���ӷ� = False) Then
                If bln���� = False Then
                    If dblCheckThreeMoney <> 0 Then
                        MsgBox "ʹ�������ӿڽ���ĹҺŵ���,���ܽ��Һŷѷֿ���!", vbInformation, gstrSysName
                        Exit Function
                    End If
                    
                    If mintInsure <> 0 And MCPAR.���ղ����� = False Then
                        MsgBox "ʹ��ҽ���ӿڽ���ĹҺŵ���,���ܽ��Һŷѷֿ���!", vbInformation, gstrSysName
                        Exit Function
                    End If
                    
                    If mstr����NO <> "" Then
                        MsgBox "�ҺŲ������۵�ʱ,��֧�ֹҺŷѷֱ���!", vbInformation, gstrSysName
                        Exit Function
                    End If
                End If
                
                '�����˷ѽ�����.
                For i = 1 To vsfMoney.Rows - 1
                    curMoney = Val(vsfMoney.TextMatrix(i, 2)) + curMoney
                Next
                
                byt�˷����� = 1
                bln�˷��ش� = gbln�˷��ش�
            ElseIf mbln���������� And chk������.Value = 1 Then
                If mbln���ӷ� And chkExtra.Value = 0 Then
                    If bln���� = False Then
                        If dblCheckThreeMoney <> 0 Then
                            MsgBox "ʹ�������ӿڽ���ĹҺŵ���,���ܽ��Һŷѷֿ���!", vbInformation, gstrSysName
                            Exit Function
                        End If
                        
                        If mintInsure <> 0 Then
                            MsgBox "ʹ��ҽ���ӿڽ���ĹҺŵ���,���ܽ��Һŷѷֿ���!", vbInformation, gstrSysName
                            Exit Function
                        End If
                        
                        If mstr����NO <> "" Then
                            MsgBox "�ҺŲ������۵�ʱ,��֧�ֹҺŷѷֱ���!", vbInformation, gstrSysName
                            Exit Function
                        End If
                    End If
                End If
                '�����˷ѽ�����.
                For i = 1 To vsfMoney.Rows - 1
                    curMoney = Val(vsfMoney.TextMatrix(i, 2)) + curMoney
                Next
                
                byt�˷����� = 4
                bln�˷��ش� = gbln�˷��ش�
            ElseIf mbln���ӷ� And chkExtra.Value = 1 Then
                If mbln���������� And chk������.Value = 0 Then
                    If bln���� = False Then
                        If dblCheckThreeMoney <> 0 Then
                            MsgBox "ʹ�������ӿڽ���ĹҺŵ���,���ܽ��Һŷѷֿ���!", vbInformation, gstrSysName
                            Exit Function
                        End If
                        
                        If mintInsure <> 0 And MCPAR.���ղ����� = False Then
                            MsgBox "ʹ��ҽ���ӿڽ���ĹҺŵ���,���ܽ��Һŷѷֿ���!", vbInformation, gstrSysName
                            Exit Function
                        End If
                        
                        If mstr����NO <> "" Then
                            MsgBox "�ҺŲ������۵�ʱ,��֧�ֹҺŷѷֱ���!", vbInformation, gstrSysName
                            Exit Function
                        End If
                    End If
                End If
                '�����˷ѽ�����.
                For i = 1 To vsfMoney.Rows - 1
                    curMoney = Val(vsfMoney.TextMatrix(i, 2)) + curMoney
                Next
                
                byt�˷����� = 5
                bln�˷��ش� = gbln�˷��ش�
            End If
        Else
            If (mbln���������� And chk������.Value = 1) And (mbln���ӷ� And chkExtra.Value = 1) Then
                MsgBox "�Ѿ������ĹҺŵ���,���ܽ��������븽�ӷ�һ����!", vbInformation, gstrSysName
                Exit Function
            End If
            
            If (mbln���������� And chk������.Value = 1) Then
                If bln���� = False Then
                    If dblCheckThreeMoney <> 0 Then
                        MsgBox "ʹ�������ӿڽ���ĹҺŵ���,���ܽ���������Һŷѷֿ���!", vbInformation, gstrSysName
                        Exit Function
                    End If
                    
                    If mintInsure <> 0 And MCPAR.���ղ����� = False Then
                        MsgBox "ʹ��ҽ���ӿڽ���ĹҺŵ���,���ܽ��Һŷѷֿ���!", vbInformation, gstrSysName
                        Exit Function
                    End If
                    
                    If mstr����NO <> "" Then
                        MsgBox "�ҺŲ������۵�ʱ,��֧�ֲ�������Һŷѷֱ���!", vbInformation, gstrSysName
                        Exit Function
                    End If
                End If
                
                '�����˷ѽ�����.
                For i = 1 To vsfMoney.Rows - 1
                    curMoney = Val(vsfMoney.TextMatrix(i, 2)) + curMoney
                Next
                
                byt�˷����� = 2
                bln�˷��ش� = gbln�˷��ش�
            End If
            
            If (mbln���ӷ� And chkExtra.Value = 1) Then
                If bln���� = False Then
                    If dblCheckThreeMoney <> 0 Then
                        MsgBox "ʹ�������ӿڽ���ĹҺŵ���,���ܽ��Һŷ���" & mstr���ӷ� & "�ֿ���!", vbInformation, gstrSysName
                        Exit Function
                    End If
                    
                    If mintInsure <> 0 Then
                        MsgBox "ʹ��ҽ���ӿڽ���ĹҺŵ���,���ܽ��Һŷѷֿ���!", vbInformation, gstrSysName
                        Exit Function
                    End If
                    
                    If mstr����NO <> "" Then
                        MsgBox "�ҺŲ������۵�ʱ,��֧�ֹҺŷ���" & mstr���ӷ� & "�ֱ���!", vbInformation, gstrSysName
                        Exit Function
                    End If
                End If
                
                '�����˷ѽ�����.
                For i = 1 To vsfMoney.Rows - 1
                    curMoney = Val(vsfMoney.TextMatrix(i, 2)) + curMoney
                Next
                
                byt�˷����� = 3
                bln�˷��ش� = gbln�˷��ش�
            End If
        End If
    Case 1
        If bln���� = False Then
            If dblCheckThreeMoney <> 0 Then
                MsgBox "ʹ�������ӿڽ���ĹҺŵ���,���ܽ���������Һŷѷֿ���!", vbInformation, gstrSysName
                Exit Function
            End If
            
            If mintInsure <> 0 And MCPAR.���ղ����� = False Then
                MsgBox "ʹ��ҽ���ӿڽ���ĹҺŵ���,���ܽ��Һŷѷֿ���!", vbInformation, gstrSysName
                Exit Function
            End If
            
            If mstr����NO <> "" Then
                MsgBox "�ҺŲ������۵�ʱ,��֧�ֲ�������Һŷѷֱ���!", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        
        '�����˷ѽ�����.
        For i = 1 To vsfMoney.Rows - 1
            curMoney = Val(vsfMoney.TextMatrix(i, 2)) + curMoney
        Next
        
        byt�˷����� = 2
        bln�˷��ش� = gbln�˷��ش�
    Case 2
        If bln���� = False Then
            If dblCheckThreeMoney <> 0 Then
                MsgBox "ʹ�������ӿڽ���ĹҺŵ���,���ܽ��Һŷ���" & mstr���ӷ� & "�ֿ���!", vbInformation, gstrSysName
                Exit Function
            End If
            
            If mintInsure <> 0 Then
                MsgBox "ʹ��ҽ���ӿڽ���ĹҺŵ���,���ܽ��Һŷ���" & mstr���ӷ� & "�ֿ���!", vbInformation, gstrSysName
                Exit Function
            End If
            
            If mstr����NO <> "" Then
                MsgBox "�ҺŲ������۵�ʱ,��֧�ֹҺŷ���" & mstr���ӷ� & "�ֱ���!", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        
        '�����˷ѽ�����.
        For i = 1 To vsfMoney.Rows - 1
            curMoney = Val(vsfMoney.TextMatrix(i, 2)) + curMoney
        Next
        
        byt�˷����� = 3
        bln�˷��ش� = gbln�˷��ش�
    End Select
    
    blnҽ��ԭ���� = True
    If mintInsure <> 0 Then
        Call initInsurePara(lng����ID)
        If bln���� = False Then
            If mstr�����ʻ� = "" Then
                strSQL = "Select ���� From ���㷽ʽ Where ����=3"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
                If Not rsTmp.EOF Then
                    mstr�����ʻ� = Nvl(rsTmp!����)
                End If
            End If
            blnҽ��ԭ���� = False
            For i = 1 To vsfPay.Rows - 1
                If (vsfPay.TextMatrix(i, 0) = mstr�����ʻ�) And vsfPay.TextMatrix(i, 0) <> "" And vsfPay.RowHidden(i) = False Then
                    blnҽ��ԭ���� = True
                End If
            Next i
            strAdvance = IIf(mstr�����ʻ� <> "", mstr�����ʻ�, "�����ʻ�")
            If blnҽ��ԭ���� = True Then
                If gclsInsure.GetCapability(support�����������, , mintInsure, strAdvance) Then
                    strAdvance = ""     '����̴��벻�����˵Ľ��㷽ʽ,�ձ�ʾȫ������
                End If
            End If
            If MCPAR.ҽ���ӿڴ�ӡƱ�� Then
                 If zlGetInvoiceGroupUseID(lng����ID) = False Then Exit Function
                 strInvoice = GetNextBill(lng����ID)
            End If
        End If
    ElseIf bln���� = False Then
        Set rsOneCard1 = GetOneCardBalance(mlng����ID)
        
        If rsOneCard1.RecordCount > 0 Then
            If mbln���������� And chk������.Value = 0 Then
                '����������
                MsgBox "ʹ��һ��ͨ�ӿڽ��пۿ�,���ܽ���������Һŷѷֿ���!", vbInformation, gstrSysName
                Exit Function
            End If
            If mbln���ӷ� And chkExtra.Value = 0 Then
                '����������
                MsgBox "ʹ��һ��ͨ�ӿڽ��пۿ�,���ܽ���������" & mstr���ӷ� & "�ֿ���!", vbInformation, gstrSysName
                Exit Function
            End If
            On Error Resume Next
            Set objICCard = CreateObject("zlICCard.clsICCard")
            On Error GoTo 0
            If objICCard Is Nothing Then
                MsgBox "һ��ͨ�ӿڴ���ʧ��,���ܽ����˷�!����ӿ��ļ�.", vbInformation, gstrSysName
                Exit Function
            End If
            strCardNo = objICCard.Read_Card(Me)
            If strCardNo = "" Then Exit Function
            If strCardNo <> rsOneCard1!��λ�ʺ� Then
                MsgBox "��ǰ������ۿ�Ų�һ��!���ܽ����˷�.", vbInformation, gstrSysName
                Exit Function
            End If
                    
            If lngCard����ID <> 0 Then
                Set rsOneCard2 = GetOneCardBalance(lngCard����ID)
            End If
        End If
        '�����������
        If Not mCurCardPay.objCard Is Nothing Then
            If mCurCardPay.objCard.�ӿ���� <> 0 Then
                If IsCheckCancelValied(lng����ID, lngCard����ID, cllBillBalance, dblThreeMoney, mCurCardPay.objCard.�Ƿ��˿��鿨) = False Then Exit Function
            End If
        End If
    End If
    
    If byt�˷����� = 0 Then
        '��ȡ�ջ�Ʊ��
        strSQL = _
        "   Select A.����" & vbNewLine & _
        "   From Ʊ��ʹ����ϸ A" & vbNewLine & _
        "   Where A.���� = 1 And a.ԭ�� <> 6 " & vbNewLine & _
        "           And A.��ӡid = (Select Max(ID) From Ʊ�ݴ�ӡ���� Where �������� = [2] And NO = [1])" & vbNewLine & _
        "Minus" & vbNewLine & _
        "Select A.����" & vbNewLine & _
        "From Ʊ��ʹ����ϸ A" & vbNewLine & _
        "Where A.���� = 2 And a.ԭ�� <> 6 " & vbNewLine & _
        "   And A.��ӡid = (Select Max(ID) From Ʊ�ݴ�ӡ���� Where �������� = [2] And NO = [1])" & vbNewLine & _
        "Order By ����"
        Set rsInvoice = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�ջ�Ʊ��", strNO, 4)
        Do While Not rsInvoice.EOF
            strBackInvoice = strBackInvoice & "," & rsInvoice!����
            rsInvoice.MoveNext
        Loop
        If strBackInvoice <> "" Then strBackInvoice = Mid(strBackInvoice, 2)
    Else
        If gblnBill�Һ� Then
            If frmReInvoice.ShowMe(Me, strNO, dblTotal, CDbl(curMoney), strBackInvoice, blnReprint) = False Then Exit Function
            If blnReprint = False Then bln�˷��ش� = False
        End If
    End If
    
    strBalance = ""
    str�ֽ� = ""
    dbl�ֽ� = 0
    With vsfPay
        For i = 1 To .Rows - 1
            If .RowData(i) = 1 And .TextMatrix(i, 0) <> "" Then
                str�ֽ� = .TextMatrix(i, 0)
                dbl�ֽ� = Val(.TextMatrix(i, 1))
                Exit For
            End If
        Next i
        dblԤ�� = 0
        For i = 1 To .Rows - 1
            If .RowData(i) = 0 And .TextMatrix(i, 0) <> "" And Val(.TextMatrix(i, 1)) <> 0 Then
                dblԤ�� = dblԤ�� + Val(.TextMatrix(i, 1))
                Exit For
            End If
        Next i
        If str�ֽ� = "" Then
            strSQL = "Select ���� From ���㷽ʽ Where ����=1 Order By ȱʡ��־ Desc"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
            If rsTmp.EOF Then
                str�ֽ� = "�ֽ�"
            Else
                str�ֽ� = Nvl(rsTmp!����)
            End If
        End If
        For i = 1 To .Rows - 1
            If .TextMatrix(i, 0) <> "" Then
                If .TextMatrix(i, 0) = mstr�����ʻ� Then
                    If Val(.TextMatrix(i, 1)) <> 0 Then
                        If InStr(strAdvance, mstr�����ʻ�) <> 0 Then
                            dbl�ֽ� = dbl�ֽ� + Val(.TextMatrix(i, 1))
                        Else
                            strBalance = strBalance & "|" & .TextMatrix(i, 0) & "," & Val(.TextMatrix(i, 1)) & ",0"
                        End If
                    End If
                Else
                    If .RowData(i) = 7 Or .RowData(i) = 8 Then
                        strBalance = strBalance & "|" & mrsBillAdvance!���㷽ʽ & "," & Val(.TextMatrix(i, 1)) & ",1"
                    Else
                        If .RowData(i) <> 0 And .TextMatrix(i, 0) <> "" And .TextMatrix(i, 0) <> str�ֽ� Then
                            strBalance = strBalance & "|" & .TextMatrix(i, 0) & "," & Val(.TextMatrix(i, 1)) & ",0"
                        End If
                    End If
                End If
            End If
        Next i
        If str�ֽ� <> "" And dbl�ֽ� <> 0 Then
            strBalance = strBalance & "|" & str�ֽ� & "," & dbl�ֽ� & ",0"
        End If
        If strBalance <> "" Then strBalance = Mid(strBalance, 2)
    End With
        
    cmdOK.Enabled = False      '��ֹ��ӡ�������ô�ӡ���ķ�ģ̬���弰ҽ�������ӳ�
    On Error GoTo errH
    If mstr����NO <> "" And bln���� = False Then
        strSQL = "zl_���ﻮ�ۼ�¼_Delete('" & mstr����NO & "')"
        zlAddArray cllPro, strSQL
    End If
    If strSQLCard <> "" Then zlAddArray cllPro, strSQLCard   '����ʱ�˿�
    
    If mstrForceNote = "" And strԭ���㷽ʽ <> "" And str���㷽ʽ <> strԭ���㷽ʽ Then
        strBackNote = objCard.���� & "����"
    Else
        strBackNote = mstrForceNote
    End If
    'zl_���˹Һż�¼_Delete
    strSQL = "zl_���˹Һż�¼_����_DELETE("
    '  ���ݺ�_In       ������ü�¼.No%Type,
    strSQL = strSQL & "'" & strNO & "',"
    '  ����Ա���_In   ������ü�¼.����Ա���%Type,
    strSQL = strSQL & "'" & UserInfo.��� & "',"
    '  ����Ա����_In   ������ü�¼.����Ա����%Type,
    strSQL = strSQL & "'" & UserInfo.���� & "',"
    '  ժҪ_In         ������ü�¼.ժҪ%Type := Null, --ԤԼȡ��ʱ ��д ���ԤԼȡ��ԭ��
    strSQL = strSQL & "" & IIf(Me.cbo��ע.Text <> "", "'" & Me.cbo��ע.Text & "'", " NULL ") & ","
    '  ɾ�������_In   Number := 0,
    strSQL = strSQL & "" & bytTogetherDo & ","
    '  ��ԭ���˽���_In Varchar2 := Null,
    If strAdvance <> "" Or str���㷽ʽ <> strԭ���㷽ʽ Then
        If strAdvance <> "" Then strԭ���㷽ʽ = strԭ���㷽ʽ & "," & strAdvance
        If Left(strԭ���㷽ʽ, 1) = "," Then strԭ���㷽ʽ = Mid(strԭ���㷽ʽ, 2)
    End If
    strSQL = strSQL & IIf(strԭ���㷽ʽ = "" Or bln����, "NULL", "'" & strԭ���㷽ʽ & "'") & ","
    '  �˷�����_In     In Number := 0, --0-ȫ�� 1-�˹Һŷ� 2-�˲����� 3-�˸��ӷ� 4-�˹Һ�&������ 5-�˹Һ�&���ӷ�
    strSQL = strSQL & "" & byt�˷����� & ","
    '  ��ָ������_In   ����Ԥ����¼.���㷽ʽ%Type := Null
    strSQL = strSQL & IIf(str���㷽ʽ = "" Or bln����, "NULL", "'" & str���㷽ʽ & "'") & ","
    '  �˺�����_In   Number := 1
    strSQL = strSQL & IIf(mTy_Para.blnReuseCancelNO, 1, 0) & ","
    '  ���㷽ʽ_In   Varchar2 := Null
    strSQL = strSQL & "'" & strBalance & "',"
    '  ��Ԥ��_In     ����Ԥ����¼.��Ԥ��%Type := Null
    strSQL = strSQL & "" & ZVal(dblԤ��) & ","
    strSQL = strSQL & "'" & strBackInvoice & "','"
    '  ����˵��_In   ����Ԥ����¼.����˵��%Type := Null
    strSQL = strSQL & strBackNote & "')"
    zlAddArray cllPro, strSQL
    
    blnNotCommit = False
    '��Ҫ��������ý���
    '�˺�
    Err = 0: On Error GoTo Errhand:
    blnTrans = True
    zlExecuteProcedureArrAy cllPro, Me.Caption, True
    If mintInsure <> 0 And Not (MCPAR.���ղ����� = True And mintCancel = 1) Then
        '68991
        '�Һ���ȡ��ʽ(0��1)|�Һŵ���
        Dim strAdvanceTemp As String
        If bln���� Then strAdvanceTemp = "1|" & strNO
        If Not gclsInsure.RegistDelSwap(mlng����ID, mintInsure, strAdvanceTemp) Then
            gcnOracle.RollbackTrans: cmdOK.Enabled = True: Exit Function
        End If
        
        blnNotCommit = True
    ElseIf Not rsOneCard1 Is Nothing And bln���� = False Then
        If rsOneCard1.RecordCount > 0 Then
            If Not objICCard.ReturnSwap(Nvl(rsOneCard1!��λ�ʺ�), Nvl(rsOneCard1!ҽԺ����), "" & rsOneCard1!�������, Nvl(rsOneCard1!���)) Then
                gcnOracle.RollbackTrans
                MsgBox "һ��ͨ�˷ѽ��׵���ʧ�ܣ��˷Ѳ���ʧ�ܣ�", vbExclamation, gstrSysName
                cmdOK.Enabled = True: Exit Function
            End If
            If Not rsOneCard2 Is Nothing Then
                If rsOneCard2.RecordCount > 0 Then
                    If Not objICCard.ReturnSwap(Nvl(rsOneCard2!��λ�ʺ�), Nvl(rsOneCard2!ҽԺ����), "" & rsOneCard2!�������, Nvl(rsOneCard2!���)) Then
                        gcnOracle.RollbackTrans
                        MsgBox "һ��ͨ�˿��ѽ��׵���ʧ�ܣ��˷Ѳ���ʧ�ܣ�", vbExclamation, gstrSysName
                        cmdOK.Enabled = True: Exit Function
                    End If
                End If
            End If
        End If
    End If
    '��������
    '�˷�
    If mCurCardPay.lngҽ�ƿ����ID <> 0 And bln���� = False And dblThreeMoney <> 0 Then
        If CallBackBalanceInterface(cllBillBalance, lng����ID, lngCard����ID, dblThreeMoney, cllUpdate, cllThreeIns, strErrMsg) = False Then
             gcnOracle.RollbackTrans
             If strErrMsg <> "" Then
                MsgBox strErrMsg, vbExclamation + vbOKOnly, gstrSysName
             Else
                MsgBox "���õ������ӿڽ���ʧ��,�˴��˷Ѳ���ʧ��!", vbExclamation + vbOKOnly, gstrSysName
            End If
             Exit Function
        End If
        If Not cllBillBalance Is Nothing Then
            zlExecuteProcedureArrAy cllUpdate, Me.Caption, True, True
        End If
    End If
    '���������˿����
    If mCurCardPay.lngҽ�ƿ����ID = 0 And bln���� = False And dblThreeMoney = 0 And strSQLCard <> "" Then
        strSQL = " Select A.ʵ�ս�� From סԺ���ü�¼ A,����Ԥ����¼ B,���㷽ʽ C Where A.��¼����=5 And A.NO =(Select Max(NO) From סԺ���ü�¼ where ����ID=[1] and  ��¼����=5  )  And A.��¼״̬=2 " & _
                 "        And a.����ID=b.����ID And b.���㷽ʽ=c.���� And c.���� In (7,8)"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngCard����ID)
        If Not rsTmp.EOF Then
            If CallBackBalanceInterface(cllBillBalance, 0, lngCard����ID, -1 * Val(Nvl(rsTmp!ʵ�ս��)), cllUpdate, cllThreeIns, strErrMsg) = False Then
                gcnOracle.RollbackTrans
                If strErrMsg <> "" Then
                   MsgBox strErrMsg, vbExclamation + vbOKOnly, gstrSysName
                Else
                   MsgBox "���õ������ӿڽ���ʧ��,�˴��˷Ѳ���ʧ��!", vbExclamation + vbOKOnly, gstrSysName
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
    '�����:58567
    If Not cllThreeIns Is Nothing Then
        Err = 0: On Error GoTo ErrOthers:
        zlExecuteProcedureArrAy cllThreeIns, Me.Caption
    End If
    '����ִ��
ResumeExecute:
    '����:31634
    If mintInsure > 0 Then Call gclsInsure.BusinessAffirm(����Enum.Busi_RegistDelSwap, True, mintInsure)
    cmdOK.Enabled = True      '��ֹ��ӡ�������ô�ӡ���ķ�ģ̬���弰ҽ�������ӳ�
    blnTrans = False
    If gblnBillPrint Then
        Err = 0: On Error Resume Next
        Call gobjBillPrint.zlEraseBill_Reg("'" & strNO & "'")
        If Err <> 0 Then
            Err = 0
        End If
        On Error GoTo errH
    End If
    If bln�˷��ش� And Not bln���� And (byt�˷����� <> 0 Or blnCardReprint) Then Call RePrintBill(Me, 2, strNO, lng����ID, mintInsure, MCPAR.ҽ���ӿڴ�ӡƱ��, mstrUseType, bln�˷��ش� And Not bln���� And (byt�˷����� <> 0 Or blnCardReprint))
    
    If blnҽ��ԭ���� = True And strAdvance <> "" And mintInsure <> 0 And Not bln���� Then
        MsgBox "ҽ����֧��[" & strAdvance & "]����,��Ϊ" & str�ֽ� & "." & vbCrLf & vbCrLf & _
            "�˿��:" & Format(GetCashMoney(cboNO.Tag), "0.00") & " Ԫ.", vbInformation, gstrSysName
    End If
    mstr����NO = "": vsfMoney.Tag = ""
    zlExcuteDelRegistered = True
    Exit Function
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    '����:31634
    If mintInsure > 0 And blnNotCommit Then Call gclsInsure.BusinessAffirm(����Enum.Busi_RegistDelSwap, False, mintInsure)
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

Private Function CheckServeRange(intType As Integer, lng�շ�ϸĿID As Long, Optional intRow As Integer = 0) As Boolean
'����:����շ���Ŀ�ķ������,intType:0-�������;1-סԺ����
    Dim strSQL As String, rsTmp As ADODB.Recordset
    strSQL = "Select ����,Nvl(�������,0) As ������� From �շ���ĿĿ¼ Where ID = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "CheckServeRange", lng�շ�ϸĿID)
    If rsTmp.EOF Then
        MsgBox "����ȷ��" & IIf(intRow = 0, "", "��" & intRow & "��") & "�շ���Ŀ�ķ������,������Ŀ�Ƿ���ȷ¼��!"
        Exit Function
    Else
        Select Case intType
        Case 0
            If Val(rsTmp!�������) = 2 Or Val(rsTmp!�������) = 0 Then
                MsgBox IIf(intRow = 0, "", "��" & intRow & "��") & "�շ���Ŀ[" & rsTmp!���� & "]������������,����!"
                Exit Function
            End If
        Case 1
            If Val(rsTmp!�������) = 1 Or Val(rsTmp!�������) = 0 Then
                MsgBox IIf(intRow = 0, "", "��" & intRow & "��") & "�շ���Ŀ[" & rsTmp!���� & "]��������סԺ,����!"
                Exit Function
            End If
        Case Else
            If Val(rsTmp!�������) = 0 Then
                MsgBox IIf(intRow = 0, "", "��" & intRow & "��") & "�շ���Ŀ[" & rsTmp!���� & "]�������ڲ���,����!"
                Exit Function
            End If
        End Select
    End If
    CheckServeRange = True
End Function

Private Function CheckInputValied() As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ�����������Ч��
    '���أ����ݺϷ�,,����True,���򷵻�False
    '���ƣ����˺�
    '���ڣ�2010-07-02 11:15:29
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim dtDate As Date, lngSN As Long, i As Long, j As Long
    Dim blnHave As Boolean, blnPrice As Boolean '�������˴�Ϊ���۵�
    Dim dtԤԼ  As Date, lng��¼ID As Long, lng��Ŀid As Long, rsTemp As ADODB.Recordset
    Dim blnCheckDat   As Boolean, lngTmp As Long
    Dim rsReserve As New ADODB.Recordset, strSQL As String, strErrInfo As String
    Dim bytMode As Byte, rsCheck As ADODB.Recordset, datԤԼʱ�� As Date
    Dim strResult As String, blnר�Һ� As Boolean
    Dim dblԤ��֧�� As Double
    
    blnPrice = gblnPrice And Not mrsInfo Is Nothing And mbytMode = 0 And picBookingDate.Visible = False And mstrNoIn = ""
    dtDate = zlDatabase.Currentdate
    
    If vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("��¼ID")) = "" Then
        MsgBox "��ǰδѡ���κιҺŰ���,����ѡ��һ���ҺŰ��ź��ټ���!", vbInformation, gstrSysName
        Exit Function
    End If
    
    '82859:���ϴ�,2015/4/8,���˻�����Ϣ����
    '87876:���ϴ�,2015/8/31,�ж��ǲ����²��˹Һ�
    With mobjfrmPatiInfo
        If Not mrsInfo Is Nothing And .mlng����ID > 0 And mbln������Ϣ���� And (.mstr���� & .mstr���䵥λ <> IIf(IsNumeric(txt����.Text), txt����.Text & cbo���䵥λ.Text, txt����.Text) Or .mstr�Ա� <> NeedName(cbo�Ա�.Text) Or .mstr���� <> txtPatient.Text Or _
            .mstr���֤�� <> txtIDCard.Text Or .mstr�������� <> txt��������.Text Or .mstr����ʱ�� <> txt����ʱ��.Text) Then
            If MsgBox("���˻�����Ϣ�ѷ����ı䣬�Ƿ������", vbInformation + vbYesNo, gstrSysName) = vbNo Then
                '��¼����ԭʼ��Ϣ
                txtPatient.Text = .mstr����:  cbo�Ա�.ListIndex = cbo.FindIndex(cbo�Ա�, .mstr�Ա�, True)
                txt����.Text = .mstr����: Call txt����_Validate(False)
                If .mstr���䵥λ <> "" Then cbo���䵥λ.ListIndex = cbo.FindIndex(cbo���䵥λ, .mstr���䵥λ, True): cbo���䵥λ.Visible = True: txt����.Width = 600
                txt��������.Text = IIf(.mstr�������� = "", "____-__-__", .mstr��������): txt����ʱ��.Text = IIf(.mstr����ʱ�� = "", "__:__", .mstr����ʱ��)
                txtIDCard.Text = .mstr���֤��
                .txt���֤��.Text = .mstr���֤��
                Exit Function
            Else
                '��¼�����µ���Ϣ
                .mstr���� = txtPatient.Text: .mstr�Ա� = NeedName(cbo�Ա�.Text)
                .mstr���� = txt����.Text: .mstr���䵥λ = NeedName(cbo���䵥λ.Text)
                .mstr�������� = txt��������.Text: .mstr����ʱ�� = txt����ʱ��.Text
                .mstr���֤�� = txtIDCard.Text
            End If
        End If
    End With
    
    If txt�����.Enabled And txt�����.Visible And mintNOLength > 0 And mblnCheckNOValidity Then
    '����ֹ��������쳣�����������ʾ
        If Len(txt�����.Text) > mintNOLength + 1 Then
            MsgBox "ע��,���������Ź���,��ȷ���Ƿ���������!", vbInformation, gstrSysName
            txt�����.SetFocus
            txt�����.SelStart = 0: txt�����.SelLength = Len(txt�����.Text)
            Exit Function
        End If
    End If
    
    '��鵥��������Ч��
    If txtPatient.Text = "" Then
        If picBookingDate.Visible Then        'ԤԼ�Һ�ʱ����Ҫ�в�����Ϣ
            MsgBox "ԤԼ�Һ�ʱ�������벡����Ϣ��", vbInformation, gstrSysName
            If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
            Exit Function
        End If
        
        If txt�����.Text <> "" Then
            MsgBox "�������벡��������", vbInformation, gstrSysName
            txtPatient.SetFocus: Exit Function
        End If
    Else
        
        If CheckTextLength("����", txtPatient) = False Then Exit Function
        If CheckTextLength("����", txt����) = False Then Exit Function
        
        If mblnStructAdress Then
            If Not CheckStructAddr(padd��ͥ��ַ, padd��ͥ��ַ.MaxLength) Then Exit Function
            If Not CheckStructAddr(padd���ڵ�ַ, padd���ڵ�ַ.MaxLength) Then Exit Function
        Else
            If zlCommFun.ActualLen(cbo��ͥ��ַ.Text) > glngMax��ͥ��ַ Then
                MsgBox "��סַ���������ֻ��������" & glngMax��ͥ��ַ & "���ַ���" & glngMax��ͥ��ַ \ 2 & "�����֣�����!", vbInformation, gstrSysName
                cbo��ͥ��ַ.SetFocus: Exit Function
            End If
            
            If zlCommFun.ActualLen(cbo���ڵ�ַ.Text) > glngMax���ڵ�ַ Then
                MsgBox "���ڵ�ַ���������ֻ��������" & glngMax���ڵ�ַ & "���ַ���" & glngMax���ڵ�ַ \ 2 & "�����֣�����!", vbInformation, gstrSysName
                cbo���ڵ�ַ.SetFocus: Exit Function
            End If
        End If
    
        If txt����.Enabled And txt����.Text = "" And Not (gblnAutoAddName And txtPatient.Text = "�²���") Then
            MsgBox "�������벡�����䣡", vbInformation, gstrSysName
            txt����.SetFocus: Exit Function
        End If
        
        If mTy_Para.bln��ֹ�������� Then
            '��ֹ������������,����Ƿ�¼���������
            If txt��������.Enabled And IsDate(txt��������.Text) = False And Not (gblnAutoAddName And txtPatient.Text = "�²���") Then
                MsgBox "�������벡�˳������ڣ�", vbInformation, gstrSysName
                txt��������.SetFocus: Exit Function
            End If
            If mobjfrmPatiInfo.mobjPubPatient Is Nothing Then Exit Function
            If mobjfrmPatiInfo.mobjPubPatient.CheckPatiAge(Trim(txt����.Text) & IIf(cbo���䵥λ.Visible, cbo���䵥λ.Text, ""), _
                IIf(txt��������.Text = "____-__-__", "", txt��������.Text) & _
                IIf(txt����ʱ��.Text = "__:__", "", " " & txt����ʱ��.Text)) = False Then
                If txt��������.Enabled And txt��������.Visible Then txt��������.SetFocus
                Exit Function
            End If
        End If
        
        If cbo�Ա�.Enabled And cbo�Ա�.ListIndex = -1 Then
            MsgBox "�������벡���Ա�", vbInformation, gstrSysName
            cbo�Ա�.SetFocus: Exit Function
            Exit Function
        End If
        '89242:���ϴ�,2015/12/10,��������
        If mblnStructAdress Then
            If padd��ͥ��ַ.Visible And padd��ͥ��ַ.Enabled And padd��ͥ��ַ.Value = "" And gbln��ͥ��ַ And Not mblnStation And Not (gblnAutoAddName And txtPatient.Text = "�²���") Then
                MsgBox "�������벡����סַ��", vbInformation, gstrSysName
                If padd��ͥ��ַ.Enabled And padd��ͥ��ַ.Visible Then
                    padd��ͥ��ַ.SetFocus: Exit Function
                End If
            End If
        Else
            If cbo��ͥ��ַ.Visible And cbo��ͥ��ַ.Enabled And cbo��ͥ��ַ.Text = "" And gbln��ͥ��ַ And Not mblnStation And Not (gblnAutoAddName And txtPatient.Text = "�²���") Then
                MsgBox "�������벡����סַ��", vbInformation, gstrSysName
                If cbo��ͥ��ַ.Enabled And cbo��ͥ��ַ.Visible Then
                    cbo��ͥ��ַ.SetFocus: Exit Function
                End If
            End If
        End If
        If txt��ͥ�绰.Visible And txt��ͥ�绰.Enabled And txt��ͥ�绰.Text = "" And gbln�绰 And Not mblnStation And Not (gblnAutoAddName And txtPatient.Text = "�²���") Then
            MsgBox "�������벡����ϵ�绰��", vbInformation, gstrSysName
            If txt��ͥ�绰.Enabled And txt��ͥ�绰.Visible Then
                txt��ͥ�绰.SetFocus: Exit Function
            End If
        End If
    End If
    
    '69026,Ƚ����,2014-8-11,������Ч�Լ��
    If txt����.Enabled And txt����.Visible And Trim(txt����.Text <> "") Then
        If mobjfrmPatiInfo.mobjPubPatient Is Nothing Then Exit Function
        If mobjfrmPatiInfo.mobjPubPatient.CheckPatiAge(Trim(txt����.Text) & IIf(cbo���䵥λ.Visible, cbo���䵥λ.Text, "")) = False Then
            txt����.SetFocus: Exit Function
        End If
    End If
    '���뽨�������,ԤԼʱ���Բ���
    If mbytMode <> 1 And txt�ű�.Text <> "+" And mbln������ And txt�����.Text = "" Then
        MsgBox "ʹ�õ�ǰ�ű�ʱ��������˽������ﲡ����", vbInformation, gstrSysName
        If txt�����.Enabled Then
            txt�����.SetFocus
        ElseIf txtPatient.Enabled And txtPatient.Text = "" Then
            txtPatient.SetFocus
        End If
        Exit Function
    End If
    
     '��Ҫ����²������ַ�ʽ
    If mintInsure = 0 And mbytInState = 0 And (mbytMode = 0 Or mbytMode = 2) And txtPatient.Text = "" Then
         '��Ҫ����²������ַ�ʽ
         If zlPatiCardCheck(1, 0, "", 1) = False Then
             Call ClearmobjfrmPatiInfoFace: ClearPatientInfo
             Set mrsInfo = Nothing
             If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
             Exit Function
         End If
     End If
    'ҽ�����
    If cboҽ��.ListIndex = -1 And cboҽ��.Enabled Then
        MsgBox "����ȷ�������ҽ��,�����������ѡ����ȷ��ҽ����", vbInformation, gstrSysName
        If cboҽ��.Enabled And cboҽ��.Visible Then cboҽ��.SetFocus
        Exit Function
    End If
    '134429�����ϴ���2019/1/12����鵱ǰ�����¼ID��Ϣ�Ƿ�һ��
    lng��¼ID = Val(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("��¼ID")))
    lng��¼ID = IIf(txt�ű�.Text = "+", 0, lng��¼ID)
    lng��Ŀid = Val(Split(vsfPlan.TextMatrix(vsfPlan.Row, GetCol("IDS")), ",")(1))
    If lng��¼ID <> 0 Then
        strSQL = "Select a.����, b.��Ŀid, b.����id, b.ҽ������" & _
                " From �ٴ������Դ a, �ٴ������¼ b " & _
                " where a.Id = b.��Դid And b.id = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng��¼ID)
        If rsTemp.EOF Then
            MsgBox "�ű���Ϣ����������ѡ��", vbInformation, gstrSysName
            Exit Function
        Else
            strErrInfo = ""
            If txt�ű�.Text <> Nvl(rsTemp!����) Then strErrInfo = "���ű�"
            If mlng�Һſ���ID <> Val(Nvl(rsTemp!����ID)) Then strErrInfo = strErrInfo & "������"
            If lng��Ŀid <> Val(Nvl(rsTemp!��ĿID)) Then strErrInfo = strErrInfo & "���Һ���Ŀ"
            If strErrInfo <> "" Then
                MsgBox "�Һ���Ϣ(" & Mid(strErrInfo, 2) & ")��һ�£�������ѡ��", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End If
    
    If dtpAppointmentDate.Visible And (mbytMode = 1 Or chkBooking.Value = 1) Then '��7781
        dtDate = DateAdd("n", mTy_Para.lngԤԼ����ʱ��, dtDate)
        Select Case mcustomTime
        Case t_��ͨ:
            dtԤԼ = dtpAppointmentDate.Value
        Case t_ʱ��:
            If Format(dtpAppointmentDate.Value, "yyyy-MM-dd") <> Format(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("�Һ�ʱ��")), "yyyy-MM-dd") Then
                If Format(dtpAppointmentTime.Value, "hh:mm:ss") >= Format(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("�Һ�ʱ��")), "hh:mm:ss") Then
                    dtԤԼ = CDate(Format(dtpAppointmentDate.Value - 1, "yyyy-MM-dd") & " " & Format(dtpAppointmentTime.Value, "hh:mm:ss"))
                Else
                    dtԤԼ = CDate(Format(dtpAppointmentDate.Value, "yyyy-MM-dd") & " " & Format(dtpAppointmentTime.Value, "hh:mm:ss"))
                End If
            Else
                dtԤԼ = CDate(Format(dtpAppointmentDate.Value, "yyyy-MM-dd") & " " & Format(dtpAppointmentTime.Value, "hh:mm:ss"))
            End If
        End Select
        Select Case mViewMode
        Case V_��ͨ�ŷ�ʱ��:
            If Format(dtpAppointmentDate.Value, "yyyy-MM-dd") & " " & Trim(Getʱ��(vsfList.Row, vsfList.Col, True, True)) < Format(dtDate, "yyyy-MM-dd hh:mm:ss") Then
                 blnCheckDat = True
            End If
        Case Else:
            If dtԤԼ < dtDate Then     '27781
                  blnCheckDat = True
            End If
        End Select
        If blnCheckDat Then
            MsgBox "��ǰԤԼʱ��,С����" & Format(dtDate, "yyyy-mm-dd HH:MM") & " ,����ԤԼ!"
             If mcustomTime = t_��ͨ Then
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
            '����:51408
            With vsfPlan
                lng��¼ID = Val(.TextMatrix(.Row, .ColIndex("��¼ID")))
            End With
            
            If Check��Чʱ���(lng��¼ID, dtԤԼ) = False Then
                  MsgBox "��ǰԤԼʱ��," & Format(dtԤԼ, "yyyy-mm-dd HH:MM") & " ,�����ڹҺŰ��Ż����Ѿ���ͣ��!", vbOKOnly + vbInformation, gstrSysName
                  If dtpAppointmentDate.Enabled And dtpAppointmentDate.Visible Then dtpAppointmentDate.SetFocus
                  Exit Function
            End If
        End If
    End If
    
    '81103,Ƚ����,2014-12-26,¼�����֤�ź�,�������ڡ����䡢�Ա��ͬ���������͵���
    If Trim(txtIDCard.Text) <> "" Then
        Dim strbirthday As String, strAge As String, strSex As String, strInfo As String
        If txtIDCard.Visible And txtIDCard.Enabled And Not mobjfrmPatiInfo.mobjPubPatient Is Nothing Then
            'CheckPatiIdcard(ByVal strIdcard As String, Optional strBirthday As String, _
            '    Optional strAge As String, Optional strSex As String, Optional strErrInfo As String) As Boolean
            '���ܣ����֤����Ϸ���У��
            '��Σ�strIdCard ���֤����
            '���Σ�strBirthday  ��������TrueΪ��������
            '         strAge ��������TrueΪ����
            '         strSex ��������TrueΪ�Ա�
            '         strErrInfo ��������FalseΪ������Ϣ
            '���أ�True/False  ���֤�Ϸ�����True(�ɴ�strBirthday��strSex��ȡ�������ں��Ա�)��
            '       ���򷵻�False(�ɴ�strErrInfo��ȡ��ϸ������Ϣ)
            If mobjfrmPatiInfo.mobjPubPatient.CheckPatiIdcard(Trim(txtIDCard.Text), strbirthday, strAge, strSex, strErrInfo) Then
                If strSex <> NeedName(cbo�Ա�.Text) Then strInfo = "�Ա�"
                If strAge <> Trim(txt����.Text) & cbo���䵥λ Then strInfo = strInfo & IIf(strInfo = "", "����", "������")
                
                If strInfo <> "" Then
                    If MsgBox("�����" & strInfo & "�����֤�ŵ�" & strInfo & "��һ�£�" & _
                            "���������֤���޸�" & strInfo & "���Ƿ������", vbInformation + vbYesNo, gstrSysName) = vbYes Then
                        Call zlControl.CboLocate(cbo�Ա�, strSex)
                        txt����.Text = ReCalcOld(CDate(strbirthday), cbo���䵥λ)
                        txt��������.Text = Format(strbirthday, "yyyy-mm-dd")
                        Call txt��������_Validate(False)
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
    
    '�ѱ���
    If Not (mbytMode = 2 Or mbytMode = 0 And mbytInState = 0 And mstrNoIn <> "") And cbo�ѱ�.ListIndex = -1 Then
        MsgBox "����ȷ�����˵ķѱ�,���ܹҺţ�", vbInformation, gstrSysName
        If cbo�ѱ�.Enabled And cbo�ѱ�.Visible Then cbo�ѱ�.SetFocus
        Exit Function
    End If
    
    If Not (mbytMode = 2 Or mbytMode = 0 And mbytInState = 0 And mstrNoIn <> "") And cbo�ѱ�.ItemData(cbo�ѱ�.ListIndex) = 2 And Not mrsInfo Is Nothing Then
        MsgBox "�ò��˲����²���,����ʹ�ý��޳���ķѱ�", vbInformation, gstrSysName
        Call SetCboDefault(cbo�ѱ�): Exit Function
    End If
    
    If mbytMode = 1 Or chkBooking.Value = 1 Then
        If vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("��ſ���")) <> "" Then
            If vsfList.Cell(flexcpForeColor, vsfList.Row, vsfList.Col) <> vbBlack Or vsfList.Cell(flexcpFontStrikethru, vsfList.Row, vsfList.Col) = True Then
                MsgBox "��ǰ��ſ��Ƶĺű�ȫ����ž�������,�޷�ԤԼ��", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End If
    
    '74550,Ƚ����,2014-7-2,�ڲ�����Ժ����,ҽ��������ҽ��վ�Һ�ʱ�ܹ�ѡ����㷽ʽ(��������Ϊ7��һ��ͨ����)
    If mbytMode <> 1 And (mblnStation And Not mblnStationPrice And cbo���㷽ʽ.Visible = True) Then
        If cbo���㷽ʽ.ListIndex = -1 And Not blnPrice Then
            MsgBox "����ȷ���Һŷ��õĽ��㷽ʽ,���ܹҺţ�", vbInformation, gstrSysName
            If cbo���㷽ʽ.Enabled And cbo���㷽ʽ.Visible Then cbo���㷽ʽ.SetFocus
            Exit Function
        End If
    End If
    If mlngOutModeMC > 0 And cboҽ�����.Visible Then
        If mobjfrmPatiInfo.txtPatiMCNO(0).Text <> "" Then
            If cboҽ�����.ListIndex <= 0 Then
                MsgBox "��ȷ����ҽ�����˵�ҽ�����", vbInformation, gstrSysName
                If cboҽ�����.Visible And cboҽ�����.Enabled Then cboҽ�����.SetFocus
                Exit Function
            End If
        ElseIf cboҽ�����.ListIndex > 0 Then
            MsgBox "ȷ����ҽ�����,����δ����ҽ���ţ�", vbInformation, gstrSysName
            If cmdMore.Enabled Then Call cmdMore_Click
            Exit Function
        End If
    End If
    If cbo���ʽ.ListIndex = -1 And cbo���ʽ.Enabled And cbo���ʽ.Visible And cbo���ʽ.Locked = False Then
        MsgBox "��ѡ���˵�ҽ�Ƹ��ʽ!", vbInformation, gstrSysName
        cbo���ʽ.SetFocus
        Exit Function
    End If
    If mstr������ <> "" Then
        If Trim(txt�����.Text) = "" Then
            MsgBox "����֤��ݵ���������Ҫ�󽨵�,����Ų���Ϊ�գ�", vbInformation, gstrSysName
            If txt�����.Enabled And txt�����.Visible Then txt�����.SetFocus
            Exit Function
        End If
    End If
    '���Һ���Ŀ�����Ƿ���ȷ
    If Not (mbytMode = 2 Or mbytMode = 0 And mbytInState = 0 And mstrNoIn <> "") Then
        If txt�ű�.Text <> "+" Then
            If Trim(txt����.Text) = "" Or Trim(txt�ű�.Text) = "" Then
                MsgBox "�Һ���Ŀδ��ȷ���룬���飡", vbInformation, gstrSysName
                txt�ű�.SetFocus: Exit Function
            Else
                For i = 1 To vsfPlan.Rows - 1
                    If vsfPlan.TextMatrix(i, GetCol("�ű�")) = txt�ű�.Text Then
                        Exit For
                    End If
                Next
                If i = vsfPlan.Rows Then
                    MsgBox "�Һ���Ŀδ��ȷ���룬���飡", vbInformation, gstrSysName
                    txt�ű�.SetFocus: Exit Function
                End If
            End If
        ElseIf mrsItems Is Nothing Then
            MsgBox "�Һ���Ŀδ��ȷ���룬���飡", vbInformation, gstrSysName
            txt�ű�.SetFocus: Exit Function
        End If
    End If
    If cbo��ע.Visible And cbo��ע.Enabled Then
        If zlCommFun.ActualLen(cbo��ע.Text) > 200 Then
            MsgBox "ժҪ���ݹ��࣬������� " & 100 & " �����ֻ� " & 200 & " ���ַ���", vbInformation, gstrSysName
            cbo��ע.SetFocus: Exit Function
        End If
    End If
    
    '���
    If txtSN.Visible Then
        lngSN = Val(txtSN.Text)
        
        If Trim(txtSN.Text) <> "" And Val(txtSN.Tag) <> Val(txtSN.Text) Then  '����ǽ���ԤԼʱû�б����ü��
            If Not IsNumeric(txtSN.Text) Then
                MsgBox "�Һ����Ҫ�������֣����飡", vbInformation, gstrSysName
               If txtSN.Enabled And txtSN.Visible Then txtSN.SetFocus
               Exit Function
            ElseIf vsfList.Visible Then
                
                For i = 0 To vsfList.Rows - 1
                    For j = 0 To vsfList.Cols - 1
                        If mViewMode = v_ר�Һ� Then
                            If lngSN = Val(vsfList.TextMatrix(i, j)) Then blnHave = True: Exit For
                        ElseIf mViewMode = v_ר�Һŷ�ʱ�� Then
                            If lngSN = Val(Getʱ��(i, j, False)) Then blnHave = True: Exit For
                        End If
                    Next
                    If blnHave Then Exit For
                Next
                If Not blnHave Then
                    If InStr(mstrPrivs, ";�Ӻ�;") <= 0 Then
                        MsgBox lngSN & "�ų�������޺���!��û�����ź�����Һŵ�Ȩ��.", vbInformation, gstrSysName
                        If txtSN.Visible And txtSN.Enabled Then txtSN.SetFocus: Exit Function
                    End If
                End If
            End If
        End If
        '68659,������,2014-01-10,�Һ�ʱ����Ԥ�������޺����Ĺ�ϵ
        If mbytMode = 0 And vsfPlan.TextMatrix(vsfPlan.Row, GetCol("��ſ���")) <> "" Then
            strSQL = "Select Count(1) As Ԥ���� From �ٴ�������ſ��� Where ��¼ID = [1] And �Һ�״̬ = 3 "
            Set rsReserve = zlDatabase.OpenSQLRecord(strSQL, "��ѯ�Һ�Ԥ����", vsfPlan.TextMatrix(vsfPlan.Row, GetCol("��¼ID")))
            If Val(Nvl(rsReserve!Ԥ����)) <> 0 Then
                With vsfPlan
                    If Val(.TextMatrix(.Row, GetCol("�޺�"))) <= Val(Nvl(rsReserve!Ԥ����)) + Val(.TextMatrix(.Row, GetCol("�ѹ�"))) Then
                        If InStr(mstrPrivs, ";�Ӻ�;") = 0 Then
                            MsgBox "�úű��Ѿ�û��ʣ����ú�!(������" & Val(Nvl(rsReserve!Ԥ����)) & "��Ԥ���ű�ʹ��)��û�м����Һŵ�Ȩ��.", vbInformation, gstrSysName
                            CheckInputValied = False
                            Exit Function
                        Else
                            If MsgBox("�úű��Ѿ�û��ʣ����ú�!(������" & Val(Nvl(rsReserve!Ԥ����)) & "��Ԥ���ű�ʹ��)���Ƿ�Ҫ�����Һ�?", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
                                CheckInputValied = False
                                Exit Function
                            End If
                        End If
                    End If
                End With
            End If
        End If
    End If
    'ʹ�ô��۷ѱ�ļ��
    If mblnNoneCut And Not (mbytMode = 2 Or mbytMode = 0 And mbytInState = 0 And mstrNoIn <> "") Then
        For i = 1 To vsfMoney.Rows - 1
            If Val(vsfMoney.TextMatrix(i, 2)) <> Val(vsfMoney.TextMatrix(i, 1)) Then
                MsgBox "��û��Ȩ�޸�����ʹ�õ�ǰ�Ĵ��۷ѱ�""" & NeedName(cbo�ѱ�.Text) & """����ѡ�����������۵ķѱ�", vbInformation, gstrSysName
                If cbo�ѱ�.Enabled And cbo�ѱ�.Visible Then cbo�ѱ�.SetFocus
                Exit Function
            End If
        Next
    End If
    
    strSQL = "Select Zl_�ٴ���������_Check([1],[2],[3]) As �����Լ�� From Dual"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("��¼ID"))), NeedName(cbo�Ա�.Text), txt����.Text & IIf(cbo���䵥λ.Visible, cbo���䵥λ.Text, ""))
    If rsTemp.EOF Then
        MsgBox "��ǰѡ��Ĳ��˲����øúű�!", vbInformation, gstrSysName
        Exit Function
    Else
        If Val(Mid(Nvl(rsTemp!�����Լ��), 1, 1)) <> 0 Then
            MsgBox "��ǰѡ��Ĳ��˲����øúű�!" & vbCrLf & "ԭ��:" & Mid(Nvl(rsTemp!�����Լ��), InStr(Nvl(rsTemp!�����Լ��), "|") + 1), vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    '���������
    If Not mrsItems Is Nothing Then
        mrsItems.Filter = ""
        Do While Not mrsItems.EOF
            If Val(Nvl(mrsItems!��ĿID)) <> 0 Then
                If CheckServeRange(0, Val(Nvl(mrsItems!��ĿID))) = False Then Exit Function
            End If
            mrsItems.MoveNext
        Loop
        mrsItems.MoveFirst
    End If
    
    '********************************************
    ' ��ר�Һźͷ�ʱ�ε��������
    ' ��Ҫ����Чʱ���������
    '********************************************
    If mcustomTime = t_ʱ�� Then
        If (mViewMode <> V_��ͨ�� And mViewMode <> V_��ͨ�ŷ�ʱ�� And mbytMode = 1 And dtpAppointmentTime.Visible) Or (mbytMode = 0 And chkBooking.Value = 1 And chkBooking.Visible) Then
            If Check��Ч�ű�(vsfPlan.TextMatrix(vsfPlan.Row, _
                                            GetCol("�ű�")), CDate(Format(dtpAppointmentDate.Value, "yyyy-MM-dd") & " " & Format(dtpAppointmentTime.Value, "HH:mm:00 ")), True) = False Then
                Exit Function
            End If
        ElseIf mbytMode = 0 And mViewMode = v_ר�Һŷ�ʱ�� Then
            If vsfList.TextMatrix(vsfList.Row, vsfList.Col) <> "" Then
            '-----------------------------------------------
            '�Һ� ��� ʱ���Ƿ��ڹ���ʱ����
            '-----------------------------------------------
                If Format(CDate(Format(dtDate, "hh:mm:ss")), "hh:mm:ss") < Format(CDate(Getʱ��(vsfList.Row, vsfList.Col, True)), "hh:mm:ss") Then
                    If Check��Ч�ű�(vsfPlan.TextMatrix(vsfPlan.Row, _
                                                    GetCol("�ű�")), CDate(Format(dtDate, "yyyy-MM-dd") & " " & Format(dtpAppointmentTime.Value, "HH:mm:00 ")), False) = False Then
                        Exit Function
                    End If
                End If
            End If
        End If
    End If
    
    With vsfPay
        For i = 1 To .Rows - 1
            If .TextMatrix(i, 0) <> "" And .RowData(i) = 0 Then
                dblԤ��֧�� = Val(.TextMatrix(i, 1))
                Exit For
            End If
        Next i
    End With
    
    
    If Val(dblԤ��֧��) <> 0 Then
        mstr���˼���IDs = ""
        If Not zlDatabase.PatiIdentify(Me, glngSys, Nvl(mrsInfo!����ID, 0), Val(dblԤ��֧��), mlngModul, 1, _
                                    IDKind.GetCurCard.�ӿ����, IIf(-1 * gdblԤ��������鿨 >= Val(dblԤ��֧��), False, True), True, mstr���˼���IDs, _
                                    (gdblԤ��������鿨 <> 0), (gdblԤ��������鿨 = 2)) Then Exit Function
    End If
    
    If mbytMode >= 0 And mbytMode <= 2 And Not mrsInfo Is Nothing Then
        strSQL = "Select Zl_Fun_���˹Һż�¼_Check([1],[2],[3],[4],[5],[6]) As ����� From Dual"
        Select Case mbytMode
            Case 0
                If mstrNoIn <> "" Then
                    bytMode = 2
                    datԤԼʱ�� = CDate(Format(dtpAppointmentDate.Value, "yyyy-mm-dd"))
                Else
                    bytMode = mbytMode
                    If chkBooking.Value = 1 Then
                        datԤԼʱ�� = CDate(Format(dtpAppointmentDate.Value, "yyyy-mm-dd"))
                    Else
                        datԤԼʱ�� = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd"))
                    End If
                End If
            Case 1, 2
                bytMode = mbytMode
                datԤԼʱ�� = CDate(Format(dtpAppointmentDate.Value, "yyyy-mm-dd"))
        End Select
        blnר�Һ� = vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("ҽ��")) <> ""
        Set rsCheck = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, bytMode, Val(Nvl(mrsInfo!����ID)), Trim(txt�ű�.Text), _
                                                Val(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("��¼ID"))), datԤԼʱ��, IIf(blnר�Һ�, 1, 0))
        If Not rsCheck.EOF Then
            strResult = Nvl(rsCheck!�����)
            If Val(Mid(strResult, 1, 1)) <> 0 Then
                MsgBox Mid(strResult, 3), vbInformation, gstrSysName
                Exit Function
            End If
        Else
            MsgBox "��Ч�Լ��ʧ��,�޷�������", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    If CheckArangement() = False Then Exit Function
    
    If mbytMode = 2 Then
        If zlCheck��Լ���޺���(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("��¼ID"))) = False Then Exit Function
    End If
    
    If Len(Trim(mobjfrmPatiInfo.txt����.Text)) <= 0 And Len(Trim(mobjfrmPatiInfo.txt����.Text)) > 0 Then
        If mobjfrmPatiInfo.zl_Get����Ĭ�Ϸ������� = False Then
            Call cmdMore_Click
            Exit Function
        End If
    End If
    
    CheckInputValied = True
End Function

Private Function Check��Чʱ���(lng��¼ID As Long, datTime As Date) As Boolean
    Dim strSQL As String, rsTemp As ADODB.Recordset
    If mViewMode = v_ר�Һŷ�ʱ�� Then Check��Чʱ��� = True: Exit Function
    With vsfPlan
        '��ſ��Ʒ�ʱ�κ�,��������ʱ���Ƿ��ڳ����¼ʱ����
        If .TextMatrix(.Row, .ColIndex("��ſ���")) <> "" And Val(.TextMatrix(.Row, .ColIndex("��ʱ��"))) = 1 Then
            strSQL = "Select 1 From �ٴ������¼ Where ID=[1] And [2] Between Nvl(ͣ�￪ʼʱ��,To_Date('3000-01-01', 'yyyy-mm-dd')) And Nvl(ͣ����ֹʱ��,To_Date('3000-01-01', 'yyyy-mm-dd')) "
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng��¼ID, datTime)
            If rsTemp.EOF Then
                Check��Чʱ��� = True
            Else
                Check��Чʱ��� = False
            End If
            Exit Function
        End If
    End With
    strSQL = "Select 1 From �ٴ������¼ Where ID=[1] And [2] Between ��ʼʱ�� And ��ֹʱ�� "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng��¼ID, datTime)
    
    If rsTemp.EOF Then
        Check��Чʱ��� = False
    Else
        strSQL = "Select 1 From �ٴ������¼ Where ID=[1] And [2] Between Nvl(ͣ�￪ʼʱ��,To_Date('3000-01-01', 'yyyy-mm-dd')) And Nvl(ͣ����ֹʱ��,To_Date('3000-01-01', 'yyyy-mm-dd')) "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng��¼ID, datTime)
        If rsTemp.EOF Then
            Check��Чʱ��� = True
        Else
            Check��Чʱ��� = False
        End If
    End If
    
End Function

'��鰲����������Ƿ�Ϸ�
Private Function CheckArangement() As Boolean
    Dim str�ű� As Long, strChkTime As String
    Dim lngSN As Long, i As Long, j As Long
    Dim blnExit As Boolean
    
    If mViewMode = V_��ͨ�� Or mViewMode = v_ר�Һ� Or mbytMode = 2 Then CheckArangement = True: Exit Function
     
    Select Case mViewMode
        Case V_��ͨ�ŷ�ʱ��
        '��ʱ������,�Ժ���������в���
        Case v_ר�Һŷ�ʱ��
            lngSN = Val(txtSN.Text)
            If lngSN = 0 Then
                If mTy_Para.bln�ϸ�ʱ�ιҺ� And InStr(mstrPrivs, ";�Ӻ�;") = 0 Then
                    MsgBox "�úű��ʱ���Ѿ�ʹ�����,�����ٽ��йҺ�!", vbInformation + vbOKOnly, gstrSysName
                    Exit Function
                End If
                CheckArangement = True: Exit Function
            End If
            If vsfList.TextMatrix(vsfList.Row, vsfList.Col) Like "��*" Then CheckArangement = True: Exit Function
            If lngSN = Val(Getʱ��(vsfList.Row, vsfList.Col)) Then CheckArangement = True: Exit Function
            With vsfList
                For i = 0 To .Rows - 1
                    For j = 1 To .Cols - 1
                       If .TextMatrix(i, j) <> "" Then
                            If lngSN = Val(Getʱ��(i, j, False)) Then
                               .Row = i: .Col = j
                                dtpAppointmentTime.Value = CDate(Getʱ��(i, j, True))
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
    Dim int���� As Integer, i As Integer
    If cbo���㷽ʽ.Text = "" Then
        If GetRegistMoney = 0 Then
            CheckPayStyleValied = True
            Exit Function
        Else
            Exit Function
        End If
    End If
    If cbo���㷽ʽ.Visible = False Then Exit Function
    
    If mbln�����Һ� Then
        For i = 1 To vsfPay.Rows - 1
            If vsfPay.TextMatrix(i, 0) = NeedName(cbo���㷽ʽ.Text) Then
                If Val(txt�ɿ�.Text) <> 0 Then
                    If Val(txt�ɿ�.Text) < Val(txt����Ӧ��.Text) Then
                        MsgBox "����Ľɿ����,����������!", vbInformation, gstrSysName
                        If txt�ɿ�.Enabled And txt�ɿ�.Visible Then txt�ɿ�.SetFocus
                        Exit Function
                    End If
                End If
                lngRow = i
                CheckPayStyleValied = True
                Exit Function
            End If
        Next i
    End If
    
    If Val(txt�ɿ�.Text) = 0 And Val(txt����Ӧ��.Text) = 0 Then CheckPayStyleValied = True: Exit Function
    
    For i = 1 To vsfPay.Rows - 1
        If vsfPay.TextMatrix(i, 0) = NeedName(cbo���㷽ʽ.Text) And cbo���㷽ʽ.Enabled And ((Val(txt�ɿ�.Text) <> 0 And Val(txt����Ӧ��.Text) = 0) Or (Val(txt�ɿ�.Text) = 0 And Val(txt����Ӧ��.Text) <> 0) Or (Val(txt�ɿ�.Text) <> 0 And Val(txt����Ӧ��.Text) <> 0)) Then
            If Val(vsfPay.TextMatrix(i, 1)) <> 0 Then
                MsgBox "�Ѿ�����" & NeedName(cbo���㷽ʽ.Text) & "�Ľ��㷽ʽ,������ʹ�øý��㷽ʽ֧��!", vbInformation, gstrSysName
                If cbo���㷽ʽ.Enabled And cbo���㷽ʽ.Visible Then cbo���㷽ʽ.SetFocus
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
    
    If NeedName(cbo���㷽ʽ.Text) = "Ԥ����" Then
        If Val(txt�ɿ�.Text) > Val(txt����Ӧ��.Text) Then
            MsgBox "ʹ��Ԥ����֧�����ܳ������ιҺŽ��!", vbInformation, gstrSysName
            Exit Function
        End If
        CheckPayStyleValied = True: Exit Function
    End If
    
    strSQL = "Select ���� From ���㷽ʽ Where ����=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, NeedName(cbo���㷽ʽ.Text))
    If rsTemp.EOF Then
        strSQL = "Select 8 As ���� From ҽ�ƿ���� Where ����=[1] Union Select 8 As ���� From ���ѿ����Ŀ¼ Where ����=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, NeedName(cbo���㷽ʽ.Text))
        If rsTemp.EOF Then
            MsgBox "����ȷ����ǰѡ��Ľ��㷽ʽ,����!", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    If txt�ɿ�.Visible And txt�ɿ�.Enabled And mTy_Para.byt�ɿʽ = 2 Then
        If Val(txt����Ӧ��.Text) <> 0 And Val(txt�ɿ�.Text) = 0 Then
            MsgBox "������ɿ��", vbInformation, gstrSysName
            txt�ɿ�.SetFocus
            Exit Function
        End If
    End If
    
    If Val(rsTemp!����) <> 1 And Val(txt�ɿ�.Text) > Val(txt����Ӧ��.Text) Then
        MsgBox "ʹ�÷��ֽ���㷽ʽʱ���ܳ������ιҺŽ��!", vbInformation, gstrSysName
        Exit Function
    End If
    
    If Val(rsTemp!����) = 3 Or Val(rsTemp!����) = 8 Or Val(rsTemp!����) = 7 Then
        For i = 1 To vsfPay.Rows - 1
            If (Val(vsfPay.RowData(i)) = 3 Or Val(vsfPay.RowData(i)) = 7 Or Val(vsfPay.RowData(i)) = 8) And Val(vsfPay.TextMatrix(i, 1)) <> 0 Then
                MsgBox "Ŀǰֻ����һ�ֽӿڵ�֧����ʽ,������ʹ��" & NeedName(cbo���㷽ʽ.Text) & "!", vbInformation, gstrSysName
                Exit Function
            End If
        Next i
    End If
    CheckPayStyleValied = True
End Function

Private Function Get����(ByVal str֧����ʽ As String, ByRef str���㷽ʽ As String) As Integer
    On Error GoTo errHandle
    Dim strSQL As String, rsTemp As ADODB.Recordset
    If str֧����ʽ = "Ԥ����" Then Get���� = 0: str���㷽ʽ = str֧����ʽ: Exit Function
    strSQL = "Select ���� From ���㷽ʽ Where ����=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str֧����ʽ)
    If Not rsTemp.EOF Then
        Get���� = Val(rsTemp!����)
        str���㷽ʽ = str֧����ʽ
    Else
        strSQL = "Select ���㷽ʽ From ҽ�ƿ���� Where ����=[1] Union Select ���㷽ʽ From ���ѿ����Ŀ¼ Where ����=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str֧����ʽ)
        If Not rsTemp.EOF Then
            Get���� = 8
            str���㷽ʽ = Nvl(rsTemp!���㷽ʽ)
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
            If (.RowData(i) = 1 Or .RowData(i) = 2) And .TextMatrix(i, 0) = NeedName(cbo���㷽ʽ.Text) Then
                If Val(.TextMatrix(i, 7)) = 0 Then
                    txt����Ӧ��.Text = Format(Val(.TextMatrix(i, 1)), "0.00")
                    For j = 0 To .Cols - 1
                        .TextMatrix(i, j) = ""
                    Next j
                    .RowData(i) = ""
                Else
                    txt����Ӧ��.Text = Format(Val(.TextMatrix(i, 1)), "0.00")
                    .TextMatrix(i, 1) = Format(.TextMatrix(i, 7), "0.00")
                End If
            End If
        Next i
    End With
    mbln�����Һ� = mblnPre����
End Sub

Private Function PrivCheck() As Boolean
    '�Һ�Ȩ�޼��
    '����Ѻ��Լ����շѺŵļ��
    Dim dblMoney As Double
    Dim i As Integer
    
    On Error GoTo Errhand
    If mbytMode <> 0 Then PrivCheck = True: Exit Function
    If zlStr.IsHavePrivs(mstrPrivs, "����Ѻ�") And zlStr.IsHavePrivs(mstrPrivs, "���շѺ�") Then PrivCheck = True: Exit Function
    
    'ͳ�ƹҺ���Ŀ���
    If Not mrsItems Is Nothing Then
        For i = 1 To mrsItems.RecordCount
            dblMoney = 0
            If Not mrsInComes Is Nothing Then
                mrsInComes.Filter = "��ĿID=" & mrsItems!��ĿID
                Do While Not mrsInComes.EOF
                    dblMoney = dblMoney + Val(Nvl(mrsInComes!Ӧ��))
                    mrsInComes.MoveNext
                Loop
            End If
            Exit For
        Next
    End If
        
    If zlStr.IsHavePrivs(mstrPrivs, "����Ѻ�") = False Then
        If RoundEx(dblMoney, 5) = 0 Then
            MsgBox "��û�й���Ѻŵ�Ȩ�ޣ�����Ϊ�ò��˹ҵ�ǰ�ű�", vbInformation, gstrSysName
            Exit Function
        End If
    ElseIf zlStr.IsHavePrivs(mstrPrivs, "���շѺ�") = False Then
        If RoundEx(dblMoney, 5) <> 0 Then
            MsgBox "��û�й��շѺŵ�Ȩ�ޣ�����Ϊ�ò��˹ҵ�ǰ�ű�", vbInformation, gstrSysName
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
        '�Һż��ҺŽ���
        Select Case Val(zlDatabase.GetPara("�Һ�ƾ����ӡ��ʽ", glngSys, mlngModul))
            Case 0    '����ӡ
                blnSlipPrint = False
            Case 1    '�Զ���ӡ
                If InStr(mstrPrivs, ";�Һ�ƾ����ӡ;") > 0 Then
                    blnSlipPrint = True
                Else
                    blnSlipPrint = False
                    MsgBox "��û�йҺ�ƾ����ӡ��Ȩ�ޣ�����ϵ����Ա��", vbInformation, gstrSysName
                End If
            Case 2    'ѡ���ӡ
                If MsgBox("Ҫ��ӡ�Һ�ƾ����", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                    If InStr(mstrPrivs, ";�Һ�ƾ����ӡ;") > 0 Then
                        blnSlipPrint = True
                    Else
                        blnSlipPrint = False
                        MsgBox "��û�йҺ�ƾ����ӡ��Ȩ�ޣ�����ϵ����Ա��", vbInformation, gstrSysName
                    End If
                Else
                    blnSlipPrint = False
                End If
        End Select
    End If
    
    If blnPrice Then
        blnNoPrint = True
        If mbytMode = 1 And mblnStation And InStr(1, gstrPrivsStation, ";ԤԼ�Һŵ�;") > 0 Then    'ҽ��վ����
            Select Case Val(zlDatabase.GetPara("ԤԼ�Һŵ���ӡ��ʽ", glngSys, 1260))    'ʹ��ҽ��վ����ز���
            Case 0    '����ӡ
            Case 1    '��������ӡ
                blnPrintBooking = True
            Case 2    'ѡ���ӡ
                If MsgBox("Ҫ��ӡ�Һ�ԤԼ����", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                    blnPrintBooking = True
                End If
            End Select
        End If
    ElseIf mbytMode <> 1 Then
        If mRegistFeeMode = EM_RG_���� Then
            blnNoPrint = True
        Else
            If Not gblnPrintFree Then blnNoPrint = (GetRegistMoney(False) = 0)
            
            If Not blnNoPrint And txt�ű�.Text = "+" And Not mblnAddCardItem And gbytInvoice <> 0 Then
                If MsgBox("��ǰ����ֻ��������Ҫ��ӡƱ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    blnNoPrint = True
                End If
            End If
            If Not blnNoPrint Then
                If gbytInvoice = 0 Then
                    blnNoPrint = True
                ElseIf gbytInvoice = 2 Then
                    If Not (txt�ű�.Text = "+" And Not mblnAddCardItem) Then    'ǰ������ʾ����,������ʾ
                        If MsgBox("Ҫ��ӡ�Һ�Ʊ����", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                            blnNoPrint = True
                        End If
                    End If
                End If
            End If
        End If
    ElseIf mbytMode = 1 Then
        Select Case Val(zlDatabase.GetPara("ԤԼ�Һŵ���ӡ��ʽ", glngSys, mlngModul))
        Case 0    '����ӡ
        Case 1    '��������ӡ
            blnPrintBooking = True
        Case 2    'ѡ���ӡ
            If MsgBox("Ҫ��ӡ�Һ�ԤԼ����", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                blnPrintBooking = True
            End If
        End Select
        blnNoPrint = True
    End If
    
    If Not mblnStation And mbytMode <> 1 Then
        Select Case gByt��ӡ��������
        Case 0: blnCodePrint = False
        Case 1: blnCodePrint = True
        Case 2:
               If MsgBox("�Ƿ���Ҫ��ӡ�������룿", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    blnCodePrint = True
               Else
                    blnCodePrint = False
               End If
        End Select
    End If
End Sub

Private Sub Getʱ��(ByVal Datsys As Date, ByVal lngSN As Long, ByVal bln׷��ʱ�� As Boolean, ByRef str�Ǽ�ʱ�� As String, _
                    ByRef str����ʱ�� As String, ByRef dat����ʱ�� As Date, ByRef bln�ﵽ�޺��� As Boolean)
    '�ڻ�ȡ�˿�����ź�  �ŶԷ���ʱ����д���
    str�Ǽ�ʱ�� = "To_Date('" & Format(Datsys, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
    If mcustomTime = t_ʱ�� Then    '��ʱ�δ���ֻҪ���а�������һ������������ʱ�θ�mcustomTime��ֵ��������Ϊ��t_ʱ�Σ�Ҳ����˵������������
        If dtpAppointmentTime.Visible = True And mbytMode <> 2 Then
            If picBookingDate.Visible And dtpAppointmentTime.Visible Then
                If lngSN <> 0 And (mViewMode = V_��ͨ�ŷ�ʱ�� Or mViewMode = v_ר�Һŷ�ʱ��) Then
                    mrsʱ���.Filter = "���=" & lngSN
                    If Not mrsʱ���.EOF Then
                        str����ʱ�� = "To_Date('" & Format(mrsʱ���!��ϸ��ʼʱ��, "yyyy-MM-dd hh:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                        dat����ʱ�� = CDate(Format(mrsʱ���!��ϸ��ʼʱ��, "yyyy-MM-dd hh:mm:ss"))
                    Else
                        str����ʱ�� = "To_Date('" & Format(dtpAppointmentDate.Value, "yyyy-MM-dd") & " " & Format(dtpAppointmentTime.Value, "HH:mm:00 ") & "','YYYY-MM-DD HH24:MI:SS')"
                        dat����ʱ�� = CDate(Format(dtpAppointmentDate.Value, "yyyy-MM-dd") & " " & Format(dtpAppointmentTime.Value, "HH:mm:00 "))
                    End If
                    mrsʱ���.Filter = ""
                Else
                    str����ʱ�� = "To_Date('" & Format(dtpAppointmentDate.Value, "yyyy-MM-dd") & " " & Format(dtpAppointmentTime.Value, "HH:mm:00 ") & "','YYYY-MM-DD HH24:MI:SS')"
                    dat����ʱ�� = CDate(Format(dtpAppointmentDate.Value, "yyyy-MM-dd") & " " & Format(dtpAppointmentTime.Value, "HH:mm:00 "))
                End If
            Else
                str����ʱ�� = "To_Date('" & Format(Datsys, "yyyy-MM-dd") & " " & Format(dtpAppointmentTime.Value, "HH:mm:00 ") & "','YYYY-MM-DD HH24:MI:SS')"
                dat����ʱ�� = CDate(Format(Datsys, "yyyy-MM-dd") & " " & Format(dtpAppointmentTime.Value, "HH:mm:00 "))
            End If
        ElseIf picBookingDate.Visible Or mbytMode = 2 Or (mbytMode = 0 And mbytInState = 0 And mstrNoIn <> "") Then
            str����ʱ�� = "To_Date('" & Format(dtpAppointmentDate.Value, "yyyy-MM-dd HH:mm:00") & "','YYYY-MM-DD HH24:MI:SS')"
            dat����ʱ�� = CDate(Format(dtpAppointmentDate.Value, "yyyy-MM-dd HH:mm:00"))
        Else
            str����ʱ�� = str�Ǽ�ʱ��
            dat����ʱ�� = Datsys
            If mbytMode = 0 Then
                If Format(dat����ʱ��, "yyyy-mm-dd hh:mm:ss") < Format(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("�Һ�ʱ��")), "yyyy-MM-dd hh:mm:ss") Then
                    dat����ʱ�� = CDate(Format(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("�Һ�ʱ��")), "yyyy-MM-dd hh:mm:ss"))
                    str����ʱ�� = "To_Date('" & Format(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("�Һ�ʱ��")), "yyyy-MM-dd hh:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                End If
            End If
        End If
        
        If vsfList.Row < vsfList.Rows And vsfList.Col < vsfList.Cols Then
            If mbytMode = 0 And mViewMode = v_ר�Һŷ�ʱ�� And dtpAppointmentTime.Visible = False And mstrNoIn = "" Then
                If vsfList.TextMatrix(vsfList.Row, vsfList.Col) <> "" Then
                    If lngSN <> 0 Then
                        mrsʱ���.Filter = "���=" & lngSN
                        If Not mrsʱ���.EOF Then
                            str����ʱ�� = "To_Date('" & Format(mrsʱ���!��ϸ��ʼʱ��, "yyyy-MM-dd hh:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                            dat����ʱ�� = CDate(Format(mrsʱ���!��ϸ��ʼʱ��, "yyyy-MM-dd hh:mm:ss"))
                        Else
                            If Format(Datsys, "hh:mm:ss") < Format(dtpAppointmentTime.Value, "hh:mm:ss") Then
                                str����ʱ�� = "To_Date('" & Format(Datsys, "yyyy-MM-dd") & " " & Format(dtpAppointmentTime.Value, "hh:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                                dat����ʱ�� = CDate(Format(Datsys, "yyyy-MM-dd") & " " & Format(dtpAppointmentTime.Value, "hh:mm:ss"))
                            End If
                        End If
                        mrsʱ���.Filter = ""
                    Else
                        If Format(Datsys, "hh:mm:ss") < Format(dtpAppointmentTime.Value, "hh:mm:ss") Then
                            str����ʱ�� = "To_Date('" & Format(Datsys, "yyyy-MM-dd") & " " & Format(dtpAppointmentTime.Value, "hh:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                            dat����ʱ�� = CDate(Format(Datsys, "yyyy-MM-dd") & " " & Format(dtpAppointmentTime.Value, "hh:mm:ss"))
                        End If
                    End If
                End If
            End If
        End If
        If Not mrsʱ��� Is Nothing And dtpAppointmentTime.Visible = False And (mViewMode = v_ר�Һŷ�ʱ�� Or mViewMode = V_��ͨ�ŷ�ʱ��) Then
            mrsʱ���.MoveLast
            With vsfPlan
                bln�ﵽ�޺��� = (Val(.TextMatrix(.Row, .ColIndex("�޺�"))) - (Val(.TextMatrix(.Row, .ColIndex("�ѹ�"))) + Val(.TextMatrix(.Row, .ColIndex("��Լ"))) - GetʧԼ��(.TextMatrix(.Row, .ColIndex("�ű�")), Datsys))) <= 0
            End With
            If bln׷��ʱ�� Or mbln�Ӻ� Or _
                (CDate(CStr(DatePart("h", CStr(mrsʱ���!��ʼʱ��))) & ":" & CStr(DatePart("n", CStr(mrsʱ���!��ʼʱ��))) & ":" & CStr(DatePart("s", CStr(mrsʱ���!��ʼʱ��)))) <= CDate(Format(CStr(DatePart("h", CStr(Datsys))) & ":" & CStr(DatePart("n", CStr(Datsys))) & ":" & CStr(DatePart("s", CStr(Datsys))), "hh:mm:ss")) And bln�ﵽ�޺��� = False) Then
                If CDate(CStr(DatePart("h", CStr(mrsʱ���!����ʱ��))) & ":" & CStr(DatePart("n", CStr(mrsʱ���!����ʱ��))) & ":" & CStr(DatePart("s", CStr(mrsʱ���!����ʱ��)))) > CDate(Format(CStr(DatePart("h", CStr(Datsys))) & ":" & CStr(DatePart("n", CStr(Datsys))) & ":" & CStr(DatePart("s", CStr(Datsys))), "hh:mm:ss")) Then
                    str����ʱ�� = "To_Date('" & Format(Datsys, "yyyy-MM-dd") & " " & Format(mrsʱ���!����ʱ��, "hh:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                    dat����ʱ�� = CDate(Format(Datsys, "yyyy-MM-dd") & " " & Format(mrsʱ���!����ʱ��, "hh:mm:ss"))
                Else
                    str����ʱ�� = "To_Date('" & Format(Datsys, "yyyy-MM-dd") & " " & Format(Datsys, "hh:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                    dat����ʱ�� = CDate(Format(Datsys, "yyyy-MM-dd") & " " & Format(Datsys, "hh:mm:ss"))
                End If
            End If
        End If
    Else    '�÷�֧�������а�����û��һ��������ʱ�ε����
        If picBookingDate.Visible Or mbytMode = 2 Or (mbytMode = 0 And mbytInState = 0 And mstrNoIn <> "") Then
            str����ʱ�� = "To_Date('" & Format(dtpAppointmentDate.Value, "yyyy-MM-dd HH:mm:00") & "','YYYY-MM-DD HH24:MI:SS')"
            dat����ʱ�� = CDate(Format(dtpAppointmentDate.Value, "yyyy-MM-dd HH:mm:00"))
        Else
            str����ʱ�� = str�Ǽ�ʱ��
            dat����ʱ�� = Datsys
        End If
    End If
End Sub
Private Function SaveRegister_ԤԼ����(ByVal lng����ID As Long, ByVal dtSysDate As Date, str����� As String, ByVal str���� As String, _
    ByVal rsCardFee As ADODB.Recordset, ByVal str����ʱ�� As String, ByVal str�Ǽ�ʱ�� As String, _
    ByVal blnPrice As Boolean, ByVal blnNoPrint As Boolean, ByVal lngSN As Long, _
    ByVal str���㷽ʽ As String, cur�ֽ� As Currency, cur���� As Currency, curԤ�� As Currency, cur���� As Currency, _
    ByRef lng����ID As Long, _
    ByRef cllPro As Collection, ByRef cllProAfter As Collection, ByRef lngCard����ID As Long, ByRef bln���� As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ԤԼ����,���Խ���ʱ�ļ۸�Ϊ׼
    '���:blnPrice-�Ƿ��Ϊ���۵�
    '     blnNoPrint-��Ʊ��ӡ��־(true-����ӡ,false-��ӡ)
    '     lng����ID-����ID
    '����:cllPro-�������ݱ��漯
    '     cllProAfter-������ִ�е����SQL��
    '     bln����-�Ƿ�ǰ����
    '     lngCard����ID-���ؿ����ݵĽ���ID
    '     lng����ID-����ID
    '����:�ɹ���ȡ���ݱ��漯,����true,���򷵻�False
    '����:���˺�
    '����:2017-11-06 11:01:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim k As Integer, i As Integer, j As Integer, int�۸񸸺� As Integer
    Dim blnHaveBookFee As Boolean '�Ƿ�没����
    Dim dblTotalRegFee As Double
    Dim str����NO As String, str���� As String, strRoom As String
    Dim strNO As String, str�ѱ� As String
    Dim lng��� As Long
    Dim strSQL As String, dblTemp As Double
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    
    '115168:���ϴ���2017/12/13�����淢����ҽ�ƿ�����
    If mCurSendCard.lng�����ID = 0 Then mCurSendCard = gCurSendCard
    If cllPro Is Nothing Then Set cllPro = New Collection
    If cllProAfter Is Nothing Then Set cllProAfter = New Collection
    '��ԤԼ���մ���ֱ�ӷ���True
    If Not (mbytMode = 2 Or mbytMode = 0 And mbytInState = 0 And mstrNoIn <> "") Then SaveRegister_ԤԼ���� = True: Exit Function
    If mlng��¼ID <> 0 And mlng��¼ID <> Val(vsfPlan.TextMatrix(vsfPlan.Row, GetCol("��¼ID"))) Then
        MsgBox "δ�ҵ�ԤԼ���ݵĳ����Ű��¼�����ܽ��յ��ݡ�", vbInformation, gstrSysName
        Exit Function
    End If
    '��ȡ��������
    If mbytMode <> 1 And txt�ű�.Text <> "+" And vsfPlan.TextMatrix(vsfPlan.Row, GetCol("����")) <> "" Then  'ԤԼʱ������
         strRoom = GetRoom(vsfPlan.TextMatrix(vsfPlan.Row, GetCol("��¼ID")))
    End If

    str�ѱ� = NeedName(cbo�ѱ�.Text)
    str���� = Trim(txt����.Text)
    If IsNumeric(str����) Then str���� = str���� & cbo���䵥λ.Text
    strNO = cboNO.Text
    If mTy_Para.blnԤԼ����ȷ���Һŷ� Then  'ԤԼ���ܣ����¼۸�Ϊ׼
        If blnPrice Then
            dblTotalRegFee = GetRegistMoney(True, False)
            '�ҺŷѴ�Ϊ���ұ���Ϊ���۵����Ų�������NO
            If dblTotalRegFee <> 0 Then str����NO = zlDatabase.GetNextNo(13)
        End If
        mrsItems.Filter = ""
        k = 1: mrsItems.MoveFirst
        For i = 1 To mrsItems.RecordCount
            int�۸񸸺� = k
            mrsInComes.Filter = "��ĿID=" & mrsItems!��ĿID
            For j = 1 To mrsInComes.RecordCount
                If mrsItems!���� = 4 Then   '�����ü�ʱ�����ƽ���һ��,��֧�����ö��������Ŀ,Ϊ�˱�������￨������һ��
                    '
                Else
                    'Zl_����ԤԼ�Һż�¼_Update
                    strSQL = "Zl_����ԤԼ�Һż�¼_Update("
                    strSQL = strSQL & "'" & strNO & "',"
                    strSQL = strSQL & "" & k & ","
                    strSQL = strSQL & "" & IIf(int�۸񸸺� = k, "NULL", int�۸񸸺�) & ","
                    strSQL = strSQL & "" & IIf(mrsItems!���� = 2, 1, "NULL") & ","
                    strSQL = strSQL & "'" & mrsItems!��� & "',"
                    strSQL = strSQL & "'" & mrsItems!��ĿID & "',"
                    strSQL = strSQL & "" & Val(Nvl(mrsItems!����)) & ","
                    strSQL = strSQL & "" & Val(Nvl(mrsInComes!����)) & ","
                    strSQL = strSQL & "" & Val(Nvl(mrsInComes!������ĿID)) & ","
                    strSQL = strSQL & "'" & Trim(Nvl(mrsInComes!�վݷ�Ŀ)) & "',"
                    strSQL = strSQL & "" & IIf(blnPrice And dblTotalRegFee <> 0, 0, mrsInComes!Ӧ��) & ","
                    strSQL = strSQL & "" & IIf(blnPrice And dblTotalRegFee <> 0, 0, mrsInComes!ʵ��) & ","
                    strSQL = strSQL & "" & IIf(mrsItems!���� = 3, 1, IIf(mrsItems!���� = 4, 2, 0)) & ","
                    strSQL = strSQL & "" & ZVal(Nvl(mrsItems!���մ���id, 0)) & ","
                    strSQL = strSQL & "" & ZVal(Nvl(mrsItems!������Ŀ��, 0)) & ","
                    strSQL = strSQL & "" & ZVal(Nvl(mrsInComes!ͳ����, 0)) & ","
                    strSQL = strSQL & "'" & Trim(Nvl(mrsItems!���ձ���)) & "',"
                    strSQL = strSQL & "" & mlng�Һſ���ID & ","
                    strSQL = strSQL & "" & IIf(mrsItems!ִ�п���ID = 0, mlng�Һſ���ID, mrsItems!ִ�п���ID) & ","
                    strSQL = strSQL & "" & "NULL" & ","
                    strSQL = strSQL & "" & IIf(mrsItems!���� = 1 Or mrsItems!���� = 2, 1, 0) & ")" '�Һ���Ŀ�Ŵ�True
                    
                    Call zlAddArray(cllPro, strSQL)
                    If blnPrice And dblTotalRegFee <> 0 Then
                        strSQL = _
                        "zl_���ﻮ�ۼ�¼_Insert('" & str����NO & "'," & k & "," & lng����ID & ",NULL," & _
                                 IIf(str����� = "", "NULL", str�����) & ",'" & NeedCode(cbo���ʽ.Text) & "'," & _
                                 "'" & txtPatient.Text & "','" & NeedName(cbo�Ա�.Text) & "','" & str���� & "'," & _
                                 "'" & str�ѱ� & "',NULL," & mlng�Һſ���ID & "," & _
                                 IIf(mblnStation, mlng�Һſ���ID, UserInfo.����ID) & ",'" & UserInfo.���� & "'," & IIf(mrsItems!���� = 2, 1, "NULL") & "," & _
                                 mrsItems!��ĿID & ",'" & mrsItems!��� & "','" & mrsItems!���㵥λ & "'," & _
                                 "NULL,1," & mrsItems!���� & ",NULL," & IIf(mrsItems!ִ�п���ID = 0, mlng�Һſ���ID, mrsItems!ִ�п���ID) & "," & IIf(int�۸񸸺� = k, "NULL", int�۸񸸺�) & "," & _
                                 mrsInComes!������ĿID & ",'" & mrsInComes!�վݷ�Ŀ & "'," & mrsInComes!���� & "," & _
                                 mrsInComes!Ӧ�� & "," & mrsInComes!ʵ�� & "," & str����ʱ�� & "," & str�Ǽ�ʱ�� & ",NULL,'" & UserInfo.���� & "','�Һ�:" & strNO & "')"
                        Call zlAddArray(cllPro, strSQL)
                    End If
                End If
                k = k + 1
                mrsInComes.MoveNext
            Next
            mrsItems.MoveNext
        Next
    Else    '��ԤԼʱ�ļ۸�Ϊ׼
        If blnPrice Then
            dblTotalRegFee = GetRegistMoney(True, False)
            '�ҺŷѴ�Ϊ���ұ���Ϊ���۵����Ų�������NO
            If dblTotalRegFee <> 0 Then str����NO = zlDatabase.GetNextNo(13)
        End If
        blnHaveBookFee = False
        mrsBill.Sort = "��� "
        mrsBill.MoveFirst
        lng��� = 0
        Do While Not mrsBill.EOF
            'Zl_����ԤԼ�Һż�¼_Update
            strSQL = "Zl_����ԤԼ�Һż�¼_Update("
            '  ���ݺ�_In     ������ü�¼.NO%Type,
            strSQL = strSQL & "'" & mrsBill!NO & "',"
            '  ���_In       ������ü�¼.���%Type,
            strSQL = strSQL & "" & mrsBill!��� & ","
            '  �۸񸸺�_In   ������ü�¼.�۸񸸺�%Type,
            strSQL = strSQL & "" & IIf(Val(Nvl(mrsBill!�۸񸸺�)) = 0, "NULL", mrsBill!�۸񸸺�) & ","
            '  ��������_In   ������ü�¼.��������%Type,
            strSQL = strSQL & "" & IIf(Val(Nvl(mrsBill!��������)) = 0, "NULL", mrsBill!��������) & ","
            '  �շ����_In   ������ü�¼.�շ����%Type,
            strSQL = strSQL & "'" & mrsBill!�շ���� & "',"
            '  �շ�ϸĿid_In ������ü�¼.�շ�ϸĿid%Type,
            strSQL = strSQL & "'" & mrsBill!�շ�ϸĿID & "',"
            '  ����_In       ������ü�¼.����%Type,
            strSQL = strSQL & "" & Val(Nvl(mrsBill!����)) & ","
            '  ��׼����_In   ������ü�¼.��׼����%Type,
            strSQL = strSQL & "" & Val(Nvl(mrsBill!��׼����)) & ","
            '  ������Ŀid_In ������ü�¼.������Ŀid%Type,
            strSQL = strSQL & "" & Val(Nvl(mrsBill!������ĿID)) & ","
            '  �վݷ�Ŀ_In   ������ü�¼.�վݷ�Ŀ%Type,
            strSQL = strSQL & "'" & Trim(Nvl(mrsBill!�վݷ�Ŀ)) & "',"
            '  Ӧ�ս��_In   ������ü�¼.Ӧ�ս��%Type,
            strSQL = strSQL & "" & IIf(blnPrice And dblTotalRegFee <> 0, 0, Val(mrsBill!Ӧ��)) & ","
            '  ʵ�ս��_In   ������ü�¼.ʵ�ս��%Type,
            dblTemp = GetActualMoney(str�ѱ�, mrsBill!������ĿID, mrsBill!Ӧ��, mrsBill!�շ�ϸĿID)
            strSQL = strSQL & "" & IIf(blnPrice And dblTotalRegFee <> 0, 0, dblTemp) & ","
            '  ������_In Number, --������¼�Ƿ���������
            If chk������.Value = 0 And Val(Nvl(mrsBill!���ӱ�־)) = 1 Then
                strSQL = strSQL & "3,"
            Else
                strSQL = strSQL & "" & Val(Nvl(mrsBill!���ӱ�־)) & ","
            End If
            If Val(Nvl(mrsBill!���ӱ�־)) = 1 Then blnHaveBookFee = True
            '  ���մ���id_In ������ü�¼.���մ���id%Type,
            strSQL = strSQL & "" & ZVal(Nvl(mrsBill!���մ���id, 0)) & ","
            '  ������Ŀ��_In ������ü�¼.������Ŀ��%Type,
            strSQL = strSQL & "" & ZVal(Nvl(mrsBill!������Ŀ��, 0)) & ","
            '  ͳ����_In   ������ü�¼.ͳ����%Type,
            strSQL = strSQL & "" & ZVal(Nvl(mrsBill!ͳ����, 0)) & ","
            '  ���ձ���_In   ������ü�¼.���ձ���%Type,
            strSQL = strSQL & "'" & Trim(Nvl(mrsBill!���ձ���)) & "',"
            '  ���˿���id_In ������ü�¼.���˿���id%Type,
            strSQL = strSQL & "" & Val(mrsBill!���˿���id) & ","
            '  ִ�в���id_In ������ü�¼.ִ�в���id%Type
            strSQL = strSQL & "" & Val(Nvl(mrsBill!ִ�в���id)) & ","
            'ժҪ_In       ������ü�¼.ժҪ%Type := Null
            strSQL = strSQL & "" & IIf(str����NO <> "", "'����:" & str����NO & "'", "NULL") & ")"
            
            lng��� = Val(Nvl(mrsBill!���))
            Call zlAddArray(cllPro, strSQL)
            If blnPrice And dblTotalRegFee <> 0 Then
                strSQL = _
                "zl_���ﻮ�ۼ�¼_Insert('" & str����NO & "'," & mrsBill!��� & "," & lng����ID & ",NULL," & _
                         IIf(str����� = "", "NULL", str�����) & ",'" & NeedCode(cbo���ʽ.Text) & "'," & _
                         "'" & txtPatient.Text & "','" & NeedName(cbo�Ա�.Text) & "','" & str���� & "'," & _
                         "'" & str�ѱ� & "',NULL," & mlng�Һſ���ID & "," & _
                         IIf(mblnStation, mlng�Һſ���ID, UserInfo.����ID) & ",'" & UserInfo.���� & "'," & "NULL" & "," & _
                         mrsBill!�շ�ϸĿID & ",'" & mrsBill!�շ���� & "',Null," & _
                         "NULL,1," & Val(Nvl(mrsBill!����)) & ",NULL," & IIf(mrsBill!ִ�в���id = 0, mlng�Һſ���ID, mrsBill!ִ�в���id) & "," & IIf(Val(Nvl(mrsBill!�۸񸸺�)) = 0, "NULL", mrsBill!�۸񸸺�) & "," & _
                         Val(Nvl(mrsBill!������ĿID)) & ",'" & Trim(Nvl(mrsBill!�վݷ�Ŀ)) & "'," & Val(Nvl(mrsBill!��׼����)) & "," & _
                         Val(mrsBill!Ӧ��) & "," & dblTemp & "," & str����ʱ�� & "," & str�Ǽ�ʱ�� & ",NULL,'" & UserInfo.���� & "','�Һ�:" & strNO & "')"
                Call zlAddArray(cllPro, strSQL)
            End If
            mrsBill.MoveNext
        Loop
        
        If lng��� = 0 Then lng��� = 1
         
        If blnHaveBookFee = False And Not mrsItems Is Nothing Then
            If blnPrice And dblTotalRegFee <> 0 And str����NO = "" Then str����NO = zlDatabase.GetNextNo(13)
            mrsItems.Filter = "���� = 3"
            Do While Not mrsItems.EOF
                mrsInComes.Filter = "��ĿID=" & mrsItems!��ĿID
                If mrsInComes.RecordCount = 0 Then
                    MsgBox "δ�ҵ�������,�����򲢷�ԭ�����������ȡʧ�ܣ���������ȡ��������!", vbInformation + vbDefaultButton1, gstrSysName
                    Exit Function
                End If
                lng��� = lng��� + 1
                
                strSQL = "Zl_����ԤԼ�Һż�¼_Update("
                strSQL = strSQL & "'" & strNO & "',"
                strSQL = strSQL & "" & lng��� & ","
                strSQL = strSQL & "NULL,"
                strSQL = strSQL & "" & IIf(mrsItems!���� = 2, 1, "NULL") & ","
                strSQL = strSQL & "'" & mrsItems!��� & "',"
                strSQL = strSQL & "'" & mrsItems!��ĿID & "',"
                strSQL = strSQL & "" & Val(Nvl(mrsItems!����)) & ","
                strSQL = strSQL & "" & Val(Nvl(mrsInComes!����)) & ","
                strSQL = strSQL & "" & Val(Nvl(mrsInComes!������ĿID)) & ","
                strSQL = strSQL & "'" & Trim(Nvl(mrsInComes!�վݷ�Ŀ)) & "',"
                strSQL = strSQL & "" & IIf(blnPrice And dblTotalRegFee <> 0, 0, mrsInComes!Ӧ��) & ","
                strSQL = strSQL & "" & IIf(blnPrice And dblTotalRegFee <> 0, 0, mrsInComes!ʵ��) & ","
                
                strSQL = strSQL & "" & IIf(mrsItems!���� = 3, 1, IIf(mrsItems!���� = 4, 2, 0)) & ","
                strSQL = strSQL & "" & ZVal(Nvl(mrsItems!���մ���id, 0)) & ","
                strSQL = strSQL & "" & ZVal(Nvl(mrsItems!������Ŀ��, 0)) & ","
                strSQL = strSQL & "" & ZVal(Nvl(mrsInComes!ͳ����, 0)) & ","
                strSQL = strSQL & "'" & Trim(Nvl(mrsItems!���ձ���)) & "',"
                strSQL = strSQL & "" & mlng�Һſ���ID & ","
                strSQL = strSQL & "" & IIf(mrsItems!ִ�п���ID = 0, mlng�Һſ���ID, mrsItems!ִ�п���ID) & ","
                'ժҪ_In       ������ü�¼.ժҪ%Type := Null
                strSQL = strSQL & "" & IIf(str����NO <> "", "'����:" & str����NO & "'", "NULL") & ")"
                
                Call zlAddArray(cllPro, strSQL)
                If blnPrice And dblTotalRegFee <> 0 Then
                    strSQL = _
                    "zl_���ﻮ�ۼ�¼_Insert('" & str����NO & "'," & mrsBill!��� + i & "," & lng����ID & ",NULL," & _
                             IIf(str����� = "", "NULL", str�����) & ",'" & NeedCode(cbo���ʽ.Text) & "'," & _
                             "'" & txtPatient.Text & "','" & NeedName(cbo�Ա�.Text) & "','" & str���� & "'," & _
                             "'" & str�ѱ� & "',NULL," & mlng�Һſ���ID & "," & _
                             IIf(mblnStation, mlng�Һſ���ID, UserInfo.����ID) & ",'" & UserInfo.���� & "'," & "NULL" & "," & _
                             mrsBill!�շ�ϸĿID & ",'" & mrsBill!�շ���� & "',Null," & _
                             "NULL,1," & Val(Nvl(mrsBill!����)) & ",NULL," & IIf(mrsBill!ִ�в���id = 0, mlng�Һſ���ID, mrsBill!ִ�в���id) & "," & IIf(Val(Nvl(mrsBill!�۸񸸺�)) = 0, "NULL", mrsBill!�۸񸸺�) & "," & _
                             Val(Nvl(mrsBill!������ĿID)) & ",'" & Trim(Nvl(mrsBill!�վݷ�Ŀ)) & "'," & Val(Nvl(mrsBill!��׼����)) & "," & _
                             Val(mrsBill!Ӧ��) & "," & GetActualMoney(str�ѱ�, mrsBill!������ĿID, mrsBill!Ӧ��, mrsBill!�շ�ϸĿID) & "," & str����ʱ�� & "," & str�Ǽ�ʱ�� & ",NULL,'" & UserInfo.���� & "','�Һ�:" & strNO & "')"
                    Call zlAddArray(cllPro, strSQL)
                End If
                mrsItems.MoveNext
            Loop
        End If
        If Not mrsItems Is Nothing Then
            
            If blnPrice And dblTotalRegFee <> 0 And str����NO = "" Then str����NO = zlDatabase.GetNextNo(13)
            
            '����ҩ�·���ѵ�
            mrsItems.Filter = "���� = 5"
            Do While Not mrsItems.EOF
                lng��� = lng��� + 1
                mrsInComes.Filter = "��ĿID=" & mrsItems!��ĿID
                strSQL = "Zl_����ԤԼ�Һż�¼_Update("
                strSQL = strSQL & "'" & strNO & "',"
                strSQL = strSQL & "" & lng��� & ","
                strSQL = strSQL & "NULL,"
                strSQL = strSQL & "" & IIf(mrsItems!���� = 2, 1, "NULL") & ","
                strSQL = strSQL & "'" & mrsItems!��� & "',"
                strSQL = strSQL & "'" & mrsItems!��ĿID & "',"
                strSQL = strSQL & "" & Val(Nvl(mrsItems!����)) & ","
                strSQL = strSQL & "" & Val(Nvl(mrsInComes!����)) & ","
                strSQL = strSQL & "" & Val(Nvl(mrsInComes!������ĿID)) & ","
                strSQL = strSQL & "'" & Trim(Nvl(mrsInComes!�վݷ�Ŀ)) & "',"
                strSQL = strSQL & "" & IIf(blnPrice And dblTotalRegFee <> 0, 0, mrsInComes!Ӧ��) & ","
                strSQL = strSQL & "" & IIf(blnPrice And dblTotalRegFee <> 0, 0, mrsInComes!ʵ��) & ","
                strSQL = strSQL & "" & IIf(mrsItems!���� = 3, 1, IIf(mrsItems!���� = 4, 2, 0)) & ","
                strSQL = strSQL & "" & ZVal(Nvl(mrsItems!���մ���id, 0)) & ","
                strSQL = strSQL & "" & ZVal(Nvl(mrsItems!������Ŀ��, 0)) & ","
                strSQL = strSQL & "" & ZVal(Nvl(mrsInComes!ͳ����, 0)) & ","
                strSQL = strSQL & "'" & Trim(Nvl(mrsItems!���ձ���)) & "',"
                strSQL = strSQL & "" & mlng�Һſ���ID & ","
                strSQL = strSQL & "" & IIf(mrsItems!ִ�п���ID = 0, mlng�Һſ���ID, mrsItems!ִ�п���ID) & ","
                'ժҪ_In       ������ü�¼.ժҪ%Type := Null
                strSQL = strSQL & "" & IIf(str����NO <> "", "'����:" & str����NO & "'", "NULL") & ")"
                Call zlAddArray(cllPro, strSQL)
                If blnPrice And dblTotalRegFee <> 0 Then
                    strSQL = _
                    "zl_���ﻮ�ۼ�¼_Insert('" & str����NO & "'," & mrsBill!��� + i & "," & lng����ID & ",NULL," & _
                     IIf(str����� = "", "NULL", str�����) & ",'" & NeedCode(cbo���ʽ.Text) & "'," & _
                     "'" & txtPatient.Text & "','" & NeedName(cbo�Ա�.Text) & "','" & str���� & "'," & _
                     "'" & str�ѱ� & "',NULL," & mlng�Һſ���ID & "," & _
                     IIf(mblnStation, mlng�Һſ���ID, UserInfo.����ID) & ",'" & UserInfo.���� & "'," & "NULL" & "," & _
                     mrsBill!�շ�ϸĿID & ",'" & mrsBill!�շ���� & "',Null," & _
                     "NULL,1," & Val(Nvl(mrsBill!����)) & ",NULL," & IIf(mrsBill!ִ�в���id = 0, mlng�Һſ���ID, mrsBill!ִ�в���id) & "," & IIf(Val(Nvl(mrsBill!�۸񸸺�)) = 0, "NULL", mrsBill!�۸񸸺�) & "," & _
                     Val(Nvl(mrsBill!������ĿID)) & ",'" & Trim(Nvl(mrsBill!�վݷ�Ŀ)) & "'," & Val(Nvl(mrsBill!��׼����)) & "," & _
                     Val(mrsBill!Ӧ��) & "," & GetActualMoney(str�ѱ�, mrsBill!������ĿID, mrsBill!Ӧ��, mrsBill!�շ�ϸĿID) & "," & str����ʱ�� & "," & str�Ǽ�ʱ�� & ",NULL,'" & UserInfo.���� & "','�Һ�:" & strNO & "')"
                    Call zlAddArray(cllPro, strSQL)
                End If
                mrsItems.MoveNext
            Loop
        End If
    End If
    
    '--ԤԼ����
    strSQL = "" & _
    "Zl_ԤԼ�ҺŽ���_����_insert('" & strNO & "','" & IIf(blnNoPrint, "", txtFact.Text) & "',Null," & _
    lng����ID & ",'" & strRoom & "'," & ZVal(lng����ID) & "," & IIf(str����� = "", "NULL", str�����) & ",'" & txtPatient.Text & "'," & _
    "'" & NeedName(cbo�Ա�.Text) & "','" & str���� & "','" & NeedCode(cbo���ʽ.Text) & "'," & _
    "'" & str�ѱ� & "','" & str���㷽ʽ & "'," & cur�ֽ� - cur���� & "," & curԤ�� & "," & cur���� & "," & _
    str����ʱ�� & "," & ZVal(lngSN) & ",'" & UserInfo.��� & "','" & UserInfo.���� & "'," & IIf(mTy_Para.bln�Һ����ɶ���, 1, 0) & "," & _
    str�Ǽ�ʱ�� & ","  '�����:48350
    '�����id_In   ����Ԥ����¼.�����id%Type := Null,
    strSQL = strSQL & "" & IIf(mCurCardPay.Have�Һŷ� And mCurCardPay.lngҽ�ƿ����ID <> 0 And mCurCardPay.bln���ѿ� = False, mCurCardPay.lngҽ�ƿ����ID, "NULL") & ","
    '���㿨���_In ����Ԥ����¼.���㿨���%Type := Null,
    strSQL = strSQL & "" & IIf(mCurCardPay.Have�Һŷ� And mCurCardPay.lngҽ�ƿ����ID <> 0 And mCurCardPay.bln���ѿ�, mCurCardPay.lngҽ�ƿ����ID, "NULL") & ","
    '����_In       ����Ԥ����¼.����%Type := Null,
    strSQL = strSQL & "" & IIf(mCurCardPay.Have�Һŷ� And mCurCardPay.strˢ������ <> "", "'" & mCurCardPay.strˢ������ & "'", "NULL") & ","
    '������ˮ��_In ����Ԥ����¼.������ˮ��%Type := Null,
    strSQL = strSQL & " NULL,"
    '����˵��_In   ����Ԥ����¼.����˵��%Type := Null
    strSQL = strSQL & " NULL,"
    '����_In       ���˹Һż�¼.����%Type := Null,
    strSQL = strSQL & "" & IIf(mintInsure = 0, "Null", mintInsure) & ","
    '����ģʽ_In   Number := 0,
    strSQL = strSQL & "" & IIf(mPatiChargeMode = EM_�����ƺ����, 1, 0) & ","
    '���ʷ���_In Number:=0
    strSQL = strSQL & "" & IIf(mRegistFeeMode = EM_RG_����, 1, 0) & ","
    '��Ԥ������ids_In Varchar2 := Null
    strSQL = strSQL & "'" & lng����ID & "," & mstr���˼���IDs & "'," '79868,Ƚ����,2015-6-15,ʹ�ü���Ԥ��
    '��������_In      Number := 0,
    strSQL = strSQL & "" & 0 & ","
    '���½������_In  Number := 1,
    strSQL = strSQL & "" & 1 & ","
    'ժҪ_In          ���˹Һż�¼.ժҪ%Type := Null
    strSQL = strSQL & "'" & cbo��ע.Text & "',"
    strSQL = strSQL & IIf(str����NO = "", "Null", "'" & str����NO & "'") & ")"
    Call zlAddArray(cllPro, strSQL)
    
    'ԤԼ�ҺŽ���
    strSQL = "" & _
           " Select D.����id, C.��Ŀid, C.ҽ��id, C.ҽ������,D.���� " & _
           " From ������ü�¼ A, ���˹Һż�¼ B, �ٴ������¼ C, �ٴ������Դ D " & _
           " Where A.��¼���� = 4 And A.��¼״̬ = 0 And A.NO = [1] And A.��� = 1 And A.NO = B.NO And B.�����¼ID = C.ID And C.��ԴID = D.ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
    '����:31187:���ҺŻ��ܵ�������
    If rsTemp.EOF = False Then
        strSQL = "zl_���˹ҺŻ���_Update("
        '  ҽ������_In   �ҺŰ���.ҽ������%Type,
        strSQL = strSQL & "'" & Nvl(rsTemp!ҽ������) & "',"
        '  ҽ��id_In     �ҺŰ���.ҽ��id%Type,
        strSQL = strSQL & "" & ZVal(Val(Nvl(rsTemp!ҽ��ID))) & ","
        '  �շ�ϸĿid_In ������ü�¼.�շ�ϸĿid%Type,
        strSQL = strSQL & "" & Val(Nvl(rsTemp!��ĿID)) & ","
        '  ִ�в���id_In ������ü�¼.ִ�в���id%Type,
        strSQL = strSQL & "" & Val(Nvl(rsTemp!����ID)) & ","
        '  ����ʱ��_In   ������ü�¼.����ʱ��%Type,
        strSQL = strSQL & "" & str����ʱ�� & ","
        '  ԤԼ��־_In   Number := 0  --�Ƿ�ΪԤԼ����:0-��ԤԼ�Һ�; 1-ԤԼ�Һ�,2-ԤԼ����
        strSQL = strSQL & "2" & ","
        '  ����_In       �ҺŰ���.����%Type := Null
        strSQL = strSQL & "'" & IIf(txt�ű�.Text = "+", "", txt�ű�.Text) & "',0,"
        strSQL = strSQL & "" & vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("��¼ID")) & ")"
        Call zlAddArray(cllProAfter, strSQL)
    End If

    SaveRegister_ԤԼ���� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Private Sub SaveData(Optional blnCall�����Һ� As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������
    '����:blnCall�����Һ�-true�����ҺŰ�ť����(����Ϊȷ�ϰ�ť����)
    '����:���˺�
    '����:2009-12-02 16:08:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim lng����ID As Long, str����� As String, lng����ID As Long, lngCard����ID As Long, lngSN As Long
    Dim str�Ǽ�ʱ�� As String, str����ʱ�� As String, strNO As String, strRoom As String, strInfo As String, strTmp As String
    Dim bytType As Byte, str�ѱ� As String, str���� As String, strPatiInforXML As String
    Dim str���� As String, str���� As String, str�������� As String
    Dim strSQL As String, strFact As String, strAdvance As String, strMCAccount As String
    Dim str��ϵ�绰 As String, intԭ����ģʽ As Integer, RegistFeeMode As EM_REGISTFEE_MODE
    Dim blnSlipPrint As Boolean, blnNoDoc As Boolean, blnCodePrint As Boolean
    Dim cur�ֽ� As Currency, cur���� As Currency, curԤ�� As Currency, str���֤�� As String, dblThreeSwap As Double  '����֧����
    Dim curOneCard As Currency, dblOneCardBalance As Double, strFilter As String
    Dim strCardNo As String, intCardType As Integer, strTransFlow As String
    Dim rsOneCard1 As ADODB.Recordset, rsOneCard2 As ADODB.Recordset, bln���� As Boolean   '��ʶ�Һ��У��Ƿ�ͬʱ�����˷�����󿨲���
    Dim objICCard As Object, dblPaySum As Double, str���㷽ʽ, blnPrice As Boolean   '�������˴�Ϊ���۵�
    Dim strStyle As String, rsCheck As ADODB.Recordset, dbl���ý�� As Double, dbl���ʽ�� As Double
    Dim int�۸񸸺� As Integer, intMsgReturn As Integer, blnNotCommit As Boolean, dat����ʱ�� As Date
    Dim blnNoPrint As Boolean, curӦ�� As Currency, cur���� As Currency, bln�ﵽ�޺��� As Boolean, blnEnterPrint As Boolean
    Dim i As Long, j As Long, k As Long, blnAfterRefresh As Boolean, cllProBefor As Collection   '����ǰִ������
    Dim blnCancel As Boolean, str����NO As String, strCardBillNO As String, cllPro As Collection   '����������ִ�е�����
    Dim blnNew As Boolean, blnPati As Boolean, blnTrans As Boolean, cllProAfter As Collection   '�ӿڵ��ú�ִ������
    Dim byt���� As Byte, blnPrintBooking As Boolean, bln���� As Boolean
    Dim rsTmp As ADODB.Recordset, rsSNCheck As ADODB.Recordset
    Dim Datsys As Date, lngRow As Long, blnInsertHisBook As Boolean
    Dim cllCardPro As Collection, cllTheeSwap As Collection, strʱ�� As String, int���� As Integer
    Dim bln׷��ʱ�� As Boolean    '���ڱ�ʶ,�Ƿ�����ʱ���Ѿ�,�ҺŻ��߹���,����û�дﵽ�޺��������,
    Dim dbl�Ҳ� As Double, curCard As Currency, cur���� As Currency
    Dim lng����ID As Long, strҽ������ As String, lngҽ��ID As Long, strErrMsg As String
    Dim lng��¼ID As Long
    Dim dblTotal(0 To 1) As Double, dblTemp As Double
   
    Err = 0: On Error GoTo ErrGo:
    mobjfrmPatiInfo.mstrFirstCode = ""
    If chkPrint.Value = 1 Then    '�ش�
        If zlRePrintRegistered = False Then Exit Sub
    ElseIf chkCancel.Value = 1 Or (mbytInState = 1 And mbytMode = 4) Then    '�˺�
        If zlExcuteDelRegistered = False Then Exit Sub
        If mbytInState = 1 And (mbytMode = 4 Or mbytMode = 3) Then mblnOk = True: Unload Me: Exit Sub
    Else
        '115168:���ϴ���2017/12/13�����淢����ҽ�ƿ�����
        If mCurSendCard.lng�����ID = 0 Then mCurSendCard = gCurSendCard
        If cbo���㷽ʽ.Visible Then
            If CheckPayStyleValied(lngRow) = False Then Exit Sub
            mblnPre���� = mbln�����Һ�
            mbln�����Һ� = False
            If lngRow <> 0 Then
                If Val(txt�ɿ�.Text) = 0 Then
                    int���� = Get����(NeedName(cbo���㷽ʽ.Text), strStyle)
                    If mTy_Para.byt�ɿʽ = 1 And int���� <> 7 And int���� <> 8 Then
                        With vsfPay
                            .TextMatrix(lngRow, 0) = NeedName(cbo���㷽ʽ.Text)
                            If Val(txt����Ӧ��.Text) <> 0 Then
                                .TextMatrix(lngRow, 1) = Format(Val(.TextMatrix(lngRow, 1)) + Val(txt����Ӧ��.Text) - mcurӦ��, "0.00")
                            End If
                            .RowData(lngRow) = int����: .TextMatrix(lngRow, .ColIndex("���㷽ʽ")) = strStyle
                        End With
                        cur���� = Val(txt����Ӧ��.Text) - mcurӦ��: txt����Ӧ��.Text = "0.00": txt�ɿ�.Text = ""
                        mbln�����Һ� = True
                    Else
                        With vsfPay
                            .TextMatrix(lngRow, 0) = NeedName(cbo���㷽ʽ.Text)
                            .TextMatrix(lngRow, 1) = Format(Val(txt����Ӧ��.Text), "0.00")
                            .RowData(lngRow) = int����
                            .TextMatrix(lngRow, .ColIndex("���㷽ʽ")) = strStyle
                        End With
                        txt����Ӧ��.Text = "0.00": txt�ɿ�.Text = ""
                    End If
                Else
                    mcur�ϼ� = 0
                    mcurӦ�� = 0
                    If Get����(NeedName(cbo���㷽ʽ.Text), strStyle) = 1 And Val(txt����Ӧ��.Text) < Val(txt�ɿ�.Text) Then
                        With vsfPay
                            .TextMatrix(lngRow, 0) = NeedName(cbo���㷽ʽ.Text)
                            .TextMatrix(lngRow, 1) = Format(Val(txt����Ӧ��.Text), "0.00")
                            .RowData(lngRow) = Get����(NeedName(cbo���㷽ʽ.Text), strStyle)
                            .TextMatrix(lngRow, .ColIndex("���㷽ʽ")) = strStyle
                        End With
                        dbl�Ҳ� = Val(txt�ɿ�.Text) - Val(txt����Ӧ��.Text)
                        txt����Ӧ��.Text = "0.00": txt�Ҳ�.Text = Format(dbl�Ҳ�, "0.00"): txt�ɿ�.Text = ""
                    Else
                        With vsfPay
                            .TextMatrix(lngRow, 0) = NeedName(cbo���㷽ʽ.Text)
                            .TextMatrix(lngRow, 1) = Format(Val(txt�ɿ�.Text), "0.00")
                            .RowData(lngRow) = Get����(NeedName(cbo���㷽ʽ.Text), strStyle)
                            .TextMatrix(lngRow, .ColIndex("���㷽ʽ")) = strStyle
                        End With
                        If Val(txt����Ӧ��.Text) = Val(txt�ɿ�.Text) Then
                            txt����Ӧ��.Text = "0.00": txt�ɿ�.Text = ""
                        Else
                            txt����Ӧ��.Text = Format(Val(txt����Ӧ��.Text) - Val(txt�ɿ�.Text), "0.00"): txt�ɿ�.Text = ""
                            If cbo���㷽ʽ.ListCount > 0 Then cbo���㷽ʽ.ListIndex = 0
                            If cbo���㷽ʽ.Visible And cbo���㷽ʽ.Enabled Then cbo���㷽ʽ.SetFocus
                            Exit Sub
                        End If
                    End If
                End If
            End If
        End If
        '�Ƿ񱣴�Ϊ���۵�
        blnPrice = CheckIsPrice
        txtPatient.Text = Trim(txtPatient.Text): txt����.Text = Trim(txt����.Text)
        If txtSN.Visible Then
            If Val(txtSN.Text) = 0 Then txtSN.Text = ""
            lngSN = Val(txtSN.Text)
        End If
        '134429:���ϴ���2019/1/12��ֻ����ѡ���кŽ����ж�
        If mlngPreRow <> vsfPlan.Row And mlngPreRow < vsfPlan.Rows And mlngPreRow <> 0 And txt�ű�.Text <> "" And txt�ű�.Text <> "+" Then
            MsgBox "�Һ����ѡ����˱仯��������ѡ��Һźű�", vbInformation, gstrSysName
            Exit Sub
        End If
        lng��¼ID = Val(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("��¼ID")))
        lng��¼ID = IIf(txt�ű�.Text = "+", 0, lng��¼ID)
        '������ݼ��
        If CheckInputValied = False Then RestorePay: Exit Sub
        If CheckNoValied(vsfPlan.Row) = False Then RestorePay: Exit Sub
        If PrivCheck() = False Then RestorePay: Exit Sub
        If Not mrsItems Is Nothing And Not mrsInComes Is Nothing Then
            mrsItems.Filter = "����=4"
            If mrsItems.RecordCount > 0 Then
                If Not mrsItems.EOF Then
                    mrsInComes.Filter = "��ĿID=" & mrsItems!��ĿID
                    If Not mrsInComes.EOF Then
                        '�����:110224,����,2017/06/20
                        If gCurSendCard.rs���� Is Nothing Then
                            MsgBox "���ѵ��շ���Ŀδ��ȷ���ã���������ԣ�", vbInformation, gstrSysName
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
        
        '���洴��������Ϣ
        If Not mobjfrmPatiInfo Is Nothing Then
            If Not mobjfrmPatiInfo.SaveAfterArrList Then Exit Sub
        End If
        
        '102230,������Ҳ����ӿ�
        If mbytMode = 0 And mbytInState = 0 Then
            If mrsInfo Is Nothing Then
                strPatiInforXML = GetPatiInforXML
                If PatiValiedCheckByPlugIn(mlngModul, 0, strPatiInforXML) = False Then
                    Call RestorePay: Exit Sub
                End If
            End If
        End If
        
        'Ʊ�ݴ�ӡ����
        Call SaveInvoiceNotify(blnPrice, blnSlipPrint, blnNoPrint, blnPrintBooking, blnCodePrint)
        
        'Ʊ�ݺ�����
        If mbytMode <> 1 And Not blnNoPrint Then
            If gblnBill�Һ� Then
                If Trim(txtFact.Text) = "" Then
                    MsgBox "��������һ����Ч��Ʊ�ݺ��룡", vbInformation, gstrSysName
                    Call RestorePay
                    txtFact.SetFocus: Exit Sub
                End If

InvoiceHandle:
                mlng����ID = CheckUsedBill(IIf(gblnSharedInvoice, 1, 4), IIf(mlng����ID > 0, mlng����ID, glng�Һ�ID), txtFact.Text, IIf(mblnStartFactUseType, mstrUseType, ""))
                If mlng����ID <= 0 Then
                    Select Case mlng����ID
                    Case 0    '����ʧ��
                    Case -1
                        MsgBox "��û�����ú͹��õĹҺ�Ʊ��,��������һ��Ʊ�ݻ����ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
                    Case -2
                        MsgBox "���صĹ���Ʊ���Ѿ�����,��������һ��Ʊ�ݻ��������ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
                    Case -3
                        MsgBox "Ʊ�ݺ��벻�ڵ�ǰ��Ч���÷�Χ��,���������룡", vbInformation, gstrSysName
                        txtFact.SetFocus
                    End Select
                    Call RestorePay
                    Exit Sub
                End If
            
                '�����������,Ʊ���Ƿ�����
                If CheckBillRepeat(mlng����ID, IIf(gblnSharedInvoice, 1, 4), txtFact.Text) Then
                    If txtFact.Locked = False And txtFact.Tag <> Trim(txtFact.Text) Then
                        MsgBox "Ʊ�ݺ�""" & txtFact.Text & """�Ѿ���ʹ�ã����������롣", vbInformation, gstrSysName
                        Call RestorePay
                        zlControl.ControlSetFocus txtFact: Exit Sub
                    Else
                        Call RefreshFact
                        If txtFact.Text = "" Then
                            Call RestorePay
                            zlControl.ControlSetFocus txtFact: Exit Sub
                        Else
                            MsgBox "��ǰƱ�ݺ��Ѿ���ʹ�ã������»�ȡƱ�ݺ�:" & txtFact.Text, vbInformation, gstrSysName
                            GoTo InvoiceHandle
                        End If
                    End If
                End If
            Else
                If Len(txtFact.Text) <> gbytFactLength And txtFact.Text <> "" Then
                    MsgBox "Ʊ�ݺ��볤��Ӧ��Ϊ " & gbytFactLength & " λ��", vbInformation, gstrSysName
                    Call RestorePay
                    txtFact.SetFocus: Exit Sub
                End If
            End If
        End If
        timPlan.Enabled = False
        
        If mRegistFeeMode <> EM_RG_���� Then
            '���ʲ�����������ʾ
            If Not (mintInsure <> 0 And mstrYBPati <> "") Then
                If gblnLED And mbytMode <> 1 And mbytInState = 0 And txt�ɿ�.Tag = "" Then
                    curӦ�� = mcurӦ�� + GetRegistMoney
                    zl9LedVoice.Speak "#21 " & Format(curӦ��, "0.00")
                End If
            End If
        End If
        txt�ɿ�.Tag = ""
        '----------------
        Set cllPro = New Collection: Set cllProAfter = New Collection: Set cllProBefor = New Collection
        Datsys = zlDatabase.Currentdate
        str�ѱ� = NeedName(cbo�ѱ�.Text)
        str���� = Trim(txt����.Text)
        If IsNumeric(str����) Then str���� = str���� & cbo���䵥λ.Text
        
        '�Һŷ�����Ϣ
        If Not blnPrice Then
            curCard = 0
            mstrCard���㷽ʽ = ""
            If Not mrsItems Is Nothing Then
                mrsItems.Filter = "����=4"
                If mrsItems.RecordCount > 0 Then
                    Do While Not mrsItems.EOF
                        mrsInComes.Filter = "��ĿID=" & mrsItems!��ĿID
                        Do While Not mrsInComes.EOF
                            curCard = curCard + mrsInComes!ʵ��
                            mrsInComes.MoveNext
                        Loop
                        mrsItems.MoveNext
                    Loop
                End If
                mrsItems.Filter = ""
            End If
            
            '137473:���ϴ���2019/1/31��������(���ѿ�)֧��ʱ����֧������
            Call ClearCardMoney
            With vsfPay
                For i = 1 To .Rows - 1
                    If .TextMatrix(i, 0) <> "" Then
                        If .RowData(i) = 0 Then
                            curԤ�� = Val(.TextMatrix(i, 1))
                        Else
                            If .RowData(i) = 3 Then
                                cur���� = Val(.TextMatrix(i, 1))
                            Else
                                If strFilter = "" Then
                                    strFilter = "���㷽ʽ='" & .TextMatrix(i, 4) & "'"
                                Else
                                    strFilter = strFilter & " Or ���㷽ʽ='" & .TextMatrix(i, 4) & "'"
                                End If
                                If curCard <> 0 Then
                                    If Val(.TextMatrix(i, .ColIndex("���"))) = curCard Then
                                        mstrCard���㷽ʽ = .TextMatrix(i, .ColIndex("���㷽ʽ"))
                                        curCard = 0
                                        If .RowData(i) = 7 Or .RowData(i) = 8 Then
                                            mCurCardPay.Have���� = True
                                        End If
                                    ElseIf Val(.TextMatrix(i, 1)) > curCard Then
                                        str���㷽ʽ = str���㷽ʽ & "|" & .TextMatrix(i, 4) & "," & Val(.TextMatrix(i, 1)) - curCard - Val(.TextMatrix(i, 7)) & "," & .TextMatrix(i, 2) & "," & IIf(.RowData(i) = 7 Or .RowData(i) = 8, 1, 0)
                                        mstrCard���㷽ʽ = .TextMatrix(i, 4)
                                        curCard = 0
                                        If .RowData(i) = 7 Or .RowData(i) = 8 Then
                                            mCurCardPay.Have���� = True
                                            mCurCardPay.Have�Һŷ� = True
                                        End If
                                    End If
                                Else
                                    str���㷽ʽ = str���㷽ʽ & "|" & .TextMatrix(i, 4) & "," & .TextMatrix(i, 1) - Val(.TextMatrix(i, 7)) & "," & .TextMatrix(i, 2) & "," & IIf(.RowData(i) = 7 Or .RowData(i) = 8, 1, 0)
                                    If .RowData(i) = 7 Or .RowData(i) = 8 Then
                                        mCurCardPay.Have�Һŷ� = True
                                    End If
                                End If
                                cur�ֽ� = cur�ֽ� + Val(.TextMatrix(i, 1)) - Val(.TextMatrix(i, 7))
                            End If
                        End If
                    End If
                Next i
                If curCard > 0 Then
                    MsgBox "¼��Ľ��㷽ʽ����֧������,������¼��һ�������(" & Format(curCard, "0.00") & ")�Ľ��㷽ʽ!", vbInformation, gstrSysName
                    Call RestorePay: Exit Sub
                End If
                If cur���� > mdbl������� + mcur����͸֧ Then
                    MsgBox "ҽ���ʻ�����,����ʹ��ҽ��֧��!", vbInformation, gstrSysName
                    Call RestorePay: Exit Sub
                End If
                If str���㷽ʽ <> "" Then
                    str���㷽ʽ = Mid(str���㷽ʽ, 2)
                Else
                    Get���� NeedName(cbo���㷽ʽ.Text), strStyle
                    str���㷽ʽ = strStyle & ",0,,0"
                End If
            End With
            
            If mblnOneCard And cur�ֽ� <> 0 And mRegistFeeMode <> EM_RG_���� Then
                mrsOneCard.Filter = strFilter
                If mrsOneCard.RecordCount > 0 Then
                    If mstrYBPati <> "" Then
                        MsgBox "��֧��ҽ������ʹ��һ��֧ͨ����", vbInformation, gstrSysName
                        Call RestorePay: Exit Sub
                    End If
                    If mobjICCard Is Nothing Then
                        MsgBox "ʹ��һ��֧ͨ�������ȶ�����", vbInformation, gstrSysName
                        Call RestorePay: Exit Sub
                    End If
                    For i = 1 To vsfPay.Rows - 1
                        If vsfPay.TextMatrix(i, 0) = Nvl(mrsOneCard!���㷽ʽ) Then
                            curOneCard = mobjICCard.GetSpare
                            If curOneCard < Val(vsfPay.TextMatrix(i, 1)) Then
                                MsgBox "�������" & Format(curOneCard, "0.00") & ",����Ҫ��֧�����" & Format(Val(vsfPay.TextMatrix(i, 1)), "0.00"), vbInformation, gstrSysName
                                Call RestorePay: Exit Sub
                            Else
                                curOneCard = Val(vsfPay.TextMatrix(i, 1))
                            End If
                            Exit For
                        End If
                    Next i
                End If
            End If
            If mRegistFeeMode <> EM_RG_���� Then
                For i = 1 To vsfPay.Rows - 1
                    strSQL = "Select ID,����,���㷽ʽ From ҽ�ƿ���� Where ����= [1] "
                    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, vsfPay.TextMatrix(i, 0))
                    If Not rsTemp.EOF Then
                        mCurCardPay.lngҽ�ƿ����ID = rsTemp!ID
                        mCurCardPay.bln���ѿ� = False
                        mCurCardPay.str���㷽ʽ = rsTemp!���㷽ʽ
                        mCurCardPay.str���� = rsTemp!����
                        If CheckBrushCard(CDbl(vsfPay.TextMatrix(i, 1))) = False Then RestorePay: Exit Sub
                        dblThreeSwap = CDbl(vsfPay.TextMatrix(i, 1))
                        Exit For
                    Else
                        strSQL = "Select ���,����,���㷽ʽ From ���ѿ����Ŀ¼ Where ����= [1] "
                        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, vsfPay.TextMatrix(i, 0))
                        If Not rsTemp.EOF Then
                            mCurCardPay.lngҽ�ƿ����ID = rsTemp!���
                            mCurCardPay.bln���ѿ� = True
                            mCurCardPay.str���㷽ʽ = rsTemp!���㷽ʽ
                            mCurCardPay.str���� = rsTemp!����
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
            mstr�����Һ�_�Һ�NO = mstr�����Һ�_�Һ�NO & "," & strNO
        End If

        If mbytMode <> 1 Then
            lng����ID = zlDatabase.GetNextId("���˽��ʼ�¼")
        End If
        byt���� = Val(mobjfrmPatiInfo.chk����.Value)
        '��ȡ��������
        If mbytMode <> 1 And txt�ű�.Text <> "+" And vsfPlan.TextMatrix(vsfPlan.Row, GetCol("����")) <> "" Then  'ԤԼʱ������
            strRoom = GetRoom(lng��¼ID)
        End If

        '�ҺŲ�����Ϣ����:�·���,�󶨿�,�Լ����������¾ɲ���
        If mblnAddCardItem Or Trim(txt�����.Text) <> "" Or (txtIDCard.Text <> "" And mbytMode = 1) Then
            str����� = txt�����.Text
            If mrsInfo Is Nothing Then
                bytType = 1
                lng����ID = zlDatabase.GetNextNo(1)
                intԭ����ģʽ = 0
            Else
                If IsNull(mrsInfo!�����) Then
                    bytType = 2
                Else
                    bytType = 3
                End If
                lng����ID = mrsInfo!����ID
                intԭ����ģʽ = Val(Nvl(mrsInfo!����ģʽ))
            End If
            blnPati = True
        ElseIf Not mrsInfo Is Nothing Then
            lng����ID = mrsInfo!����ID
            intԭ����ģʽ = Val(Nvl(mrsInfo!����ģʽ))
        End If
        
        If zlIsAllowPatiChargeFeeMode(lng����ID, intԭ����ģʽ) = False Then RestorePay: Exit Sub
        
        If Trim(mobjfrmPatiInfo.txt����.Text) <> "" Then    '��ȡ�п��ŵĲ���ʱû�м��ؿ��ŵ�����
            str���� = Trim(mobjfrmPatiInfo.txt����.Text)
            str���� = zlCommFun.zlStringEncode(Trim(mobjfrmPatiInfo.txt����.Text))
        End If

        '����ż��
        If IsValiedMzNo(lng����ID, str�����) = False Then RestorePay: Exit Sub

        If mViewMode <> V_��ͨ�� Then
            Set mrsSNState = GetSNState(lng��¼ID, CDate(Format(IIf(picBookingDate.Visible, dtpAppointmentDate.Value, Datsys), "yyyy-MM-dd")))
        End If

        '��ż��
        If Trim(txtSN.Text) <> "" And Val(txtSN.Tag) <> Val(txtSN.Text) And Not mrsSNState Is Nothing Then
            strSQL = "Select Nvl(�Һ�״̬,0) As ״̬,����Ա���� From �ٴ�������ſ��� Where ��¼ID = [1] AND ��� = [2] "
            Set rsSNCheck = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng��¼ID, lngSN)
            mrsSNState.Filter = "���=" & lngSN
            If rsSNCheck.RecordCount > 0 Then
                If rsSNCheck!״̬ = 1 Or rsSNCheck!״̬ = 2 Or rsSNCheck!״̬ = 4 Or ((rsSNCheck!״̬ = 5 Or rsSNCheck!״̬ = 3) And rsSNCheck!����Ա���� <> UserInfo.����) Then
                    Call vsfPlan_EnterCell
                    lngSN = GetCurrSN(, True)   '�Զ�ȡ��һ��
                    If lngSN = 0 Then
                        MsgBox "���" & Trim(txtSN.Text) & "�Ѿ����ҳ���ѡ���ĺŽ��йҺš�", vbInformation, gstrSysName
                        Call RestorePay: Exit Sub
                    Else
                        If IsDate(mtyRegPlanState.strSelTime) And mtyRegPlanState.lngSelNO = lngSN And Format(dtpAppointmentTime.Value, "hh:mm:00") <> Format(mtyRegPlanState.strSelTime, "hh:mm:00") Then
                            dtpAppointmentTime.Value = CDate(mtyRegPlanState.strSelTime)
                        End If
                    End If
                End If
            End If
        End If
        '���ڲ���Ա��������ѡ��Ȩ��,������ſ���,û������ʱ�����������
        '�Բ���Աֱ�ӹҳ����һ��������������Ҫ���⴦��
        '��Ϊǰ���Ѿ������������� ���ﲻ�ڽ������������ļ�� ����ֱ�Ӱ�������ſ�����û������ʱ�εİ��Ų������Ϊ�յ���������������⴦��
        If vsfPlan.TextMatrix(vsfPlan.Row, GetCol("��ſ���")) <> "" And mViewMode = v_ר�Һ� And lngSN = 0 Then
            mbln�Ӻ� = True
        ElseIf mViewMode = v_ר�Һŷ�ʱ�� And lngSN = 0 And mbln�Ӻ� = False Then
            '�����Ƕ�ר�Һŷ�ʱ����� ������ڲ���ԭ������������ű�����Ա�����ɾ��������� ���м�� ���� �ָ���Ż��� ��ʾ
            mrsSNState.Filter = 0
            i = vsfList.Row: j = vsfList.Col

            If (mtyRegPlanState.lngSelX <> vsfList.Row Or mtyRegPlanState.lngSelY <> vsfList.Col) And IsDate(mtyRegPlanState.strSelTime) Then
                '���ѡ������ʱ����ȷ,����û����ŵ����
                mblnStateChange = True
                i = mtyRegPlanState.lngSelX
                j = mtyRegPlanState.lngSelY
                If vsfList.Cell(flexcpData, vsfList.Row, vsfList.Col) Like "��*" Then
                    i = vsfList.Row
                    j = vsfList.Col
                End If
                vsfList.Select i, j
                dtpAppointmentTime.Value = CDate(mtyRegPlanState.strSelTime)
                mblnStateChange = False
            End If
            With vsfList
                If Val(vsfPlan.TextMatrix(vsfPlan.Row, GetCol("�޺�"))) <= mrsSNState.RecordCount And InStr(mstrPrivs, ";�Ӻ�;") <= 0 Then
                    '�Ӻ� �Ƿ��мӺ�Ȩ��
                    MsgBox lngSN & "�ų�������޺���!��û�����ź�����Һŵ�Ȩ��.", vbInformation, gstrSysName
                    RestorePay
                    Exit Sub
                End If
                If vsfList.TextMatrix(vsfList.Row, vsfList.Col) <> "" And .Cell(flexcpForeColor, i, j) <> vbRed _
                   And .Cell(flexcpForeColor, i, j) <> vbBlue And .Cell(flexcpForeColor, i, j) <> vbGrayText _
                   And .Cell(flexcpForeColor, i, j) <> &HC000C0 And .Cell(flexcpForeColor, i, j) <> vbGreen _
                   Then
                    If Format(Getʱ��(i, j, True), "hh:mm:00") <> Format(dtpAppointmentTime.Value, "hh:mm:ss") Then
                        dtpAppointmentTime.Value = CDate(Getʱ��(i, j, True))
                    End If
                    lngSN = GetCurrSN(, True)
                    If lngSN = 0 Then mbln�Ӻ� = True
                Else
                    '���ڹ��ڵ�ʱ��,��ʱû�дﵽ�޺���,��ʱû�дﵽ�޺���,���ӵĺ�,����ʱ��,Ϊ���һ��ʱ�εĽ���ʱ��
                    bln׷��ʱ�� = True
                End If
            End With
        End If
        
        Call Getʱ��(Datsys, lngSN, bln׷��ʱ��, str�Ǽ�ʱ��, str����ʱ��, dat����ʱ��, bln�ﵽ�޺���)
        
        If mbytMode <> 2 And mstrNoIn = "" And mViewMode <> v_ר�Һŷ�ʱ�� Then
            If vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("�Һ�ʱ��")) <> "" Then
                If Not (CDate(Format(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("�Һ�ʱ��")), "yyyy-mm-dd hh:mm:ss")) > zlDatabase.Currentdate And CDate(Format(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("��ǰʱ��")), "yyyy-mm-dd hh:mm:ss")) < zlDatabase.Currentdate) Then
                    If Check��Чʱ���(lng��¼ID, dat����ʱ��) = False Then
                        If chkShowAll.Value = 1 Then
                            If MsgBox("��ǰ�Һźű𲻵�������Ѿ���ͣ��,���Ƿ�Ҫ�����Һţ�", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                Call RestorePay: Exit Sub
                            End If
                        Else
                            MsgBox "��ǰ�Һźű𲻵�������Ѿ���ͣ��,���ܼ����Һţ�", vbInformation, gstrSysName
                            Call RestorePay: Exit Sub
                        End If
                    End If
                End If
            Else
                If chkShowAll.Value = 1 Then
                    If MsgBox("��ǰ�Һźű𲻵�������Ѿ���ͣ��,���Ƿ�Ҫ�����Һţ�", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Call RestorePay: Exit Sub
                    End If
                Else
                    MsgBox "��ǰ�Һźű𲻵�������Ѿ���ͣ��,���ܼ����Һţ�", vbInformation, gstrSysName
                    Call RestorePay: Exit Sub
                End If
            End If
        End If
        
        With vsfPlan
            If .TextMatrix(.Row, .ColIndex("���￪ʼʱ��")) <> "" Then
                If .TextMatrix(.Row, .ColIndex("����ҽ��")) <> "" And dat����ʱ�� >= CDate(.TextMatrix(.Row, .ColIndex("���￪ʼʱ��"))) And dat����ʱ�� <= CDate(.TextMatrix(.Row, .ColIndex("������ֹʱ��"))) Then
                    strҽ������ = .TextMatrix(.Row, .ColIndex("����ҽ������"))
                    lngҽ��ID = .TextMatrix(.Row, .ColIndex("����ҽ��ID"))
                Else
                    strҽ������ = mstrҽ������
                    lngҽ��ID = mlngҽ��ID
                End If
            Else
                strҽ������ = mstrҽ������
                lngҽ��ID = mlngҽ��ID
            End If
        End With
        
        If cboԤԼ��ʽ.Visible And (mbytMode = 1 Or chkBooking.Value = 1) Then
            strSQL = "Select Zl_Fun_Get�ٴ�����ԤԼ״̬([1],[2],[3],[4],[5],[6]) As ԤԼ��� From Dual"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng��¼ID, dat����ʱ��, lngSN, NeedName(cboԤԼ��ʽ.Text, , "."), "", IIf(chkBooking.Value = 1, 1, 0))
            If rsTemp.EOF Then
                MsgBox "��ǰѡ��ĺ����޷�ԤԼ,��ѡ����������!", vbInformation, gstrSysName
                If cboԤԼ��ʽ.Enabled And cboԤԼ��ʽ.Visible Then cboԤԼ��ʽ.SetFocus
                Call RestorePay: Exit Sub
            Else
                If Val(Mid(Nvl(rsTemp!ԤԼ���), 1, 1)) <> 0 Then
                    MsgBox "��ǰѡ��ĺ����޷�ԤԼ,��ѡ����������!" & vbCrLf & "ԭ��:" & Mid(Nvl(rsTemp!ԤԼ���), InStr(Nvl(rsTemp!ԤԼ���), "|") + 1), vbInformation, gstrSysName
                    If cboԤԼ��ʽ.Enabled And cboԤԼ��ʽ.Visible Then cboԤԼ��ʽ.SetFocus
                    Call RestorePay: Exit Sub
                End If
            End If
        End If
        
        '137272:���ϴ�,2019/2/20,����Ž�������,�����Ų������򷵻�һ����Ч�����
        If ReserveRegNo(lng��¼ID, lngSN, str����ʱ��, Datsys) = False Then Exit Sub
        
        str���� = Trim(str����)
        If blnPati Then
            With mobjfrmPatiInfo
                If .txt����ʱ�� = "__:__" Then
                    str�������� = IIf(IsDate(.txt��������.Text), "TO_Date('" & .txt��������.Text & "','YYYY-MM-DD')", "NULL")
                Else
                    str�������� = IIf(IsDate(.txt��������.Text), "TO_Date('" & .txt��������.Text & " " & .txt����ʱ��.Text & "','YYYY-MM-DD HH24:MI:SS')", "NULL")
                End If
                str��ϵ�绰 = Trim(txt��ͥ�绰.Text)
                str���֤�� = Trim(txtIDCard.Text)
                strSQL = _
                "zl_�ҺŲ��˲���_INSERT(" & bytType & "," & lng����ID & "," & IIf(str����� = "", "NULL", str�����) & "," & _
                         IIf(str���� = "" Or mCurSendCard.bln���￨ = False, "NULL", "'" & str���� & "'") & ",'" & str���� & "','" & txtPatient.Text & "','" & NeedName(cbo�Ա�.Text) & "'," & _
                         "'" & str���� & "','" & str�ѱ� & "','" & NeedName(cbo���ʽ.Text) & "'," & _
                         "'" & NeedName(.cbo����.Text) & "','" & NeedName(.cbo����.Text) & "','" & NeedName(.cbo����.Text) & "'," & _
                         "'" & NeedName(.cboְҵ.Text, True) & "','" & str���֤�� & "','" & .txt��λ����.Text & "'," & _
                         Val(.txt��λ����.Tag) & ",'" & .txt��λ�绰.Text & "','" & .txt��λ�ʱ�.Text & "','" & IIf(mblnStructAdress, padd��ͥ��ַ.Value, cbo��ͥ��ַ.Text) & "'," & _
                         "'" & str��ϵ�绰 & "','" & .txt��ͥ�ʱ�.Text & "'," & str�Ǽ�ʱ�� & ",''," & str�������� & ",'" & strMCAccount & _
                         "', " & IIf(str���� = "", "NULL", "'" & IIf(mblnICCard, str����, "") & "'") & "," & ZVal(mintInsure) & "," & _
                         IIf(Trim(.txt����.Text) = "", "NULL,", "'" & Trim(.txt����.Text) & "',") & _
                          "'" & IIf(mblnStructAdress, Trim(padd���ڵ�ַ.Value), Trim(cbo���ڵ�ַ.Text)) & "','" & Trim(mobjfrmPatiInfo.txt���ڵ�ַ�ʱ�.Text) & "'," & IIf(Trim(mobjfrmPatiInfo.txt��ϵ�����֤.Text) = "", "NULL,", "'" & Trim(mobjfrmPatiInfo.txt��ϵ�����֤.Text) & "',") & _
                         IIf(Trim(mobjfrmPatiInfo.txt��ϵ������.Text) = "", "NULL,", "'" & Trim(mobjfrmPatiInfo.txt��ϵ������.Text) & "',") & _
                         IIf(Trim(mobjfrmPatiInfo.txt��ϵ�˵绰.Text) = "", "NULL,", "'" & Trim(mobjfrmPatiInfo.txt��ϵ�˵绰.Text) & "',") & _
                         IIf(NeedName(mobjfrmPatiInfo.cbo��ϵ�˹�ϵ.Text) = "", "NULL,", "'" & NeedName(mobjfrmPatiInfo.cbo��ϵ�˹�ϵ.Text) & "',")
                strSQL = strSQL & IIf(Trim(mobjfrmPatiInfo.txt�໤��.Text) = "", "NULL,", "'" & Trim(mobjfrmPatiInfo.txt�໤��.Text) & "',")  'lgf
                strSQL = strSQL & IIf(Trim(mobjfrmPatiInfo.txtBirthLocation.Text) = "", "NULL,", "'" & Trim(mobjfrmPatiInfo.txtBirthLocation.Text) & "',")
                strSQL = strSQL & "'" & mobjfrmPatiInfo.txtMobile.Text & "')"
                Call zlAddArray(cllProBefor, strSQL)
                
                '90875:���ϴ�,2016/11/8,ҽ�ƿ�֤������
                If AddCertificate(lng����ID, cllProBefor, Datsys) = False Then Exit Sub
                
                '89242:���ϴ�,2015/12/7,���²��˵�ַ��Ϣ
                If mblnStructAdress Then
                    If padd��ͥ��ַ.Value <> "" Then
                       strSQL = "zl_���˵�ַ��Ϣ_update(1," & lng����ID & ",NULL,3,'" & padd��ͥ��ַ.valueʡ & "','" & _
                           padd��ͥ��ַ.value�� & "','" & padd��ͥ��ַ.value���� & "','" & padd��ͥ��ַ.value���� & "','" & _
                           padd��ͥ��ַ.value��ϸ��ַ & "','" & padd��ͥ��ַ.Code & "')"
                    Else
                       strSQL = "zl_���˵�ַ��Ϣ_update(2," & lng����ID & ",NULL,3)"
                    End If
                    Call zlAddArray(cllProBefor, strSQL)
                    If padd���ڵ�ַ.Value <> "" Then
                       strSQL = "zl_���˵�ַ��Ϣ_update(1," & lng����ID & ",NULL,4,'" & padd���ڵ�ַ.valueʡ & "','" & _
                           padd���ڵ�ַ.value�� & "','" & padd���ڵ�ַ.value���� & "','" & padd���ڵ�ַ.value���� & "','" & _
                           padd���ڵ�ַ.value��ϸ��ַ & "','" & padd���ڵ�ַ.Code & "')"
                    Else
                       strSQL = "zl_���˵�ַ��Ϣ_update(2," & lng����ID & ",NULL,4)"
                    End If
                    Call zlAddArray(cllProBefor, strSQL)
                End If
                If mobjfrmPatiInfo.txt��ϵ������.Text <> "" And NeedName(mobjfrmPatiInfo.cbo��ϵ�˹�ϵ.Text) = "����" Then
                    strSQL = "Zl_������Ϣ�ӱ�_Update("
                    '����ID_In ������Ϣ�ӱ�.����Id%Type
                    strSQL = strSQL & "" & lng����ID & ","
                    '��Ϣ��_In ������Ϣ�ӱ�.��Ϣ��%Type0
                    strSQL = strSQL & "'��ϵ�˸�����Ϣ',"
                    '��Ϣֵ_In ������Ϣ�ӱ�.��Ϣֵ%Type
                    strSQL = strSQL & "'" & mobjfrmPatiInfo.txt������ϵ.Text & "',"
                    '����Id_In ������Ϣ�ӱ�.����Id%Type
                    strSQL = strSQL & "'')"
                    Call zlAddArray(cllProBefor, strSQL)
                End If
        
                If mlngOutModeMC > 0 And cboҽ�����.ListIndex > 0 Then
                    strInfo = cboҽ�����.Text: strInfo = Mid(strInfo, 1, InStr(1, strInfo, "-") - 1)
                    strSQL = "zl_����ǼǼ�¼_UPDATE(" & mlngOutModeMC & "," & lng����ID & ",0," & str�Ǽ�ʱ�� & ",0,'" & strInfo & "')"
                    Call zlAddArray(cllProBefor, strSQL)
                End If

                If mstr������ <> "" And mint���� <> 0 Then
                    strSQL = "Zl_����������Ϣ_Insert(" & lng����ID & "," & mint���� & ",'" & mstr������ & "',1," & str�Ǽ�ʱ�� & ")"
                    Call zlAddArray(cllProBefor, strSQL)
                End If
            End With
        End If

        strSQL = "Select ID as ����ID From ���˹Һż�¼ Where ��¼״̬ = 1 And NO=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
        If rsTemp.RecordCount > 0 Then lng����ID = Val(Nvl(rsTemp!����ID))
        Err = 0: On Error GoTo ErrFirt:
        '�ȱ��没����Ϣ,Ȼ���ٴ�������,������ɲ�������(��Ҫ�ǲ���IDΪ�ظ�
        zlExecuteProcedureArrAy cllProBefor, Me.Caption, True

        '101170:���ϴ�,2016/10/13,����HIS����Ҫ�ύEMPI���ݣ�ʧ�ܺ��������ݶ�Ҫ����
        If zlSaveEMPIPatiInfo(bytType = 1, lng����ID, lng����ID, strErrMsg) = False Then
            gcnOracle.RollbackTrans
            If strErrMsg = "" Then strErrMsg = "��EMPIƽ̨�ϴ�������Ϣʧ�ܣ�"
            MsgBox strErrMsg, vbInformation, gstrSysName
            Exit Sub
        End If
        gcnOracle.CommitTrans

        Err = 0: On Error GoTo ErrGo:
        If mobjfrmPatiInfo.mblnSavePati = False Then
            Call mobjfrmPatiInfo.SavePatiPic(lng����ID)
            If CreatePlugInOK(mlngModul) And mobjfrmPatiInfo.mlngPlugInHwnd <> 0 Then  '������������Ϣ
                On Error Resume Next
                Call gobjPlugIn.PatiInfoSaveAfter(lng����ID)
                Call zlPlugInErrH(Err, "PatiInfoSaveAfter")
                Err.Clear: On Error GoTo 0
            End If
        End If
        mobjfrmPatiInfo.mblnSavePati = False
        
        RegistFeeMode = mRegistFeeMode
        If mRegistFeeMode <> EM_RG_���� Then
            RegistFeeMode = EM_RG_����
            If str���㷽ʽ = "" Then RegistFeeMode = EM_RG_����
        End If
        
        '������
        cur���� = 0                 '�Һ�ͬʱ�������ض�ֻ���ֽ���㣬���漰ҽ����Ԥ����
        mCurSendCard.dblӦ�ս�� = 0
        mCurSendCard.dblʵ�ս�� = 0
        If Not mrsItems Is Nothing Then
            mrsItems.Filter = "����=4"
            If mrsItems.RecordCount > 0 Then
                bln���� = True
                mrsInComes.Filter = "��ĿID=" & mrsItems!��ĿID
                Do While Not mrsInComes.EOF
                    cur���� = cur���� + mrsInComes!ʵ��
                    mCurSendCard.dblӦ�ս�� = mrsInComes!Ӧ�� + mCurSendCard.dblӦ�ս��
                    mrsInComes.MoveNext
                Loop
                mCurSendCard.dblʵ�ս�� = cur����
                Call AddCardDataSQL(lng����ID, Datsys, cllPro, lngCard����ID, (mRegistFeeMode = EM_RG_����), mrsItems!��ĿID)
            ElseIf str���� <> "" Then
                '����: 42947 �󶨿�,Ҳ��Ҫ��������¼
                bln���� = True    '�����:56599
                Call AddCardDataSQL(lng����ID, Datsys, cllPro, lngCard����ID)
            End If
        ElseIf str���� <> "" Then
            '����: 42947 �󶨿�,Ҳ��Ҫ��������¼
            bln���� = True    '�����:56599
            Call AddCardDataSQL(lng����ID, Datsys, cllPro, lngCard����ID)
        End If
        
        '�������ü�¼SQL���
        '------------------------------------------------------------------------------
        If (mbytMode = 2 Or mbytMode = 0 And mbytInState = 0 And mstrNoIn <> "") Then
            'ԤԼ����
            If SaveRegister_ԤԼ����(lng����ID, Datsys, str�����, str����, gCurSendCard.rs����, str����ʱ��, str�Ǽ�ʱ��, blnPrice, blnNoPrint, lngSN, _
                str���㷽ʽ, cur�ֽ�, cur����, curԤ��, cur����, lng����ID, cllPro, cllProAfter, lngCard����ID, bln����) = False Then Exit Sub
        Else
            If mobjfrmPatiInfo.txt֧������ <> "" And mobjfrmPatiInfo.txt���֤�� <> "" And mbytMode <> 1 Then    'ר����ԡ��������֤������������а�
                bln���� = True    '�����:56999
                Call AddSQL�󶨿�(lng����ID, Val(mobjfrmPatiInfo.txt֧������.Tag), mobjfrmPatiInfo.txt���֤��, zlCommFun.zlStringEncode(mobjfrmPatiInfo.txt֧������), Datsys, mblnICCard, cllPro)
            End If
            If txt�ű�.Text = "+" Then lngSN = 0
            
            dblTotal(0) = GetRegistMoney(True, False)
            dblTotal(1) = GetCardMoney  '����
            If dblTotal(0) <> 0 And blnPrice Then
                '�ҺŷѴ�Ϊ���ұ���Ϊ���۵����Ų�������NO
                   str����NO = zlDatabase.GetNextNo(13)
            End If
            
            mrsItems.Filter = ""
            k = 1: mrsItems.MoveFirst
            For i = 1 To mrsItems.RecordCount
                int�۸񸸺� = k
                mrsInComes.Filter = "��ĿID=" & mrsItems!��ĿID
                For j = 1 To mrsInComes.RecordCount
                    '����
                    If mrsItems!���� = 4 Then   '�����ü�ʱ�����ƽ���һ��,��֧�����ö��������Ŀ,Ϊ�˱�������￨������һ��
                        '
                    Else
                        '�Һ��շ�����
                        strSQL = _
                        "zl_���˹Һż�¼_����_INSERT(" & ZVal(lng��¼ID) & "," & ZVal(lng����ID) & "," & IIf(str����� = "", "NULL", str�����) & ",'" & txtPatient.Text & "','" & NeedName(cbo�Ա�.Text) & "'," & _
                                 "'" & str���� & "','" & NeedCode(cbo���ʽ.Text) & "','" & str�ѱ� & "','" & strNO & "'," & _
                                 "'" & IIf(blnNoPrint, "", txtFact.Text) & "'," & k & "," & IIf(int�۸񸸺� = k, "NULL", int�۸񸸺�) & "," & IIf(mrsItems!���� = 2, 1, "NULL") & "," & _
                                 "'" & mrsItems!��� & "'," & mrsItems!��ĿID & "," & mrsItems!���� & "," & mrsInComes!���� & "," & _
                                 mrsInComes!������ĿID & ",'" & mrsInComes!�վݷ�Ŀ & "','" & str���㷽ʽ & "'," & _
                                 IIf(blnPrice And dblTotal(0) <> 0, 0, mrsInComes!Ӧ��) & "," & IIf(blnPrice And dblTotal(0) <> 0, 0, mrsInComes!ʵ��) & "," & _
                                 mlng�Һſ���ID & "," & IIf(mblnStation, mlng�Һſ���ID, UserInfo.����ID) & "," & IIf(mrsItems!ִ�п���ID = 0, mlng�Һſ���ID, mrsItems!ִ�п���ID) & ",'" & UserInfo.��� & "','" & UserInfo.���� & "'," & _
                                 str����ʱ�� & "," & str�Ǽ�ʱ�� & "," & _
                                 "'" & strҽ������ & "'," & ZVal(lngҽ��ID) & "," & IIf(mrsItems!���� = 3, 1, IIf(mrsItems!���� = 4, 2, 0)) & "," & IIf(lbl��.Visible, 1, 0) & "," & _
                                 "'" & IIf(txt�ű�.Text = "+", "", txt�ű�.Text) & "','" & strRoom & "'," & ZVal(lng����ID) & "," & IIf(blnNoPrint, "NULL", ZVal(mlng����ID)) & "," & _
                                 ZVal(IIf(mbytMode <> 1 And k = 1, curԤ��, 0)) & "," & ZVal(IIf(mbytMode <> 1 And k = 1 And Not blnPrice, cur�ֽ� - cur����, 0)) & "," & _
                                 ZVal(IIf(mbytMode <> 1 And k = 1, cur����, 0)) & "," & ZVal(Nvl(mrsItems!���մ���id, 0)) & "," & _
                                 ZVal(Nvl(mrsItems!������Ŀ��, 0)) & "," & ZVal(Nvl(mrsInComes!ͳ����, 0)) & "," & _
                                 "'" & IIf(str����NO <> "", "����:" & str����NO, Me.cbo��ע.Text) & "'," & IIf(mbytMode = 1, 1, 0) & "," & IIf(gblnSharedInvoice, 1, 0) & ",'" & mrsItems!���ձ��� & "'," & byt���� & "," & ZVal(lngSN) & "," & ZVal(mint����) & "," & _
                                 IIf(mbytMode = 2 Or chkBooking.Value = 1 Or mbytMode = 1, 1, 0) & "," & IIf(mbytMode = 1 Or chkBooking.Value = 1, "'" & Mid(cboԤԼ��ʽ.Text, InStr(cboԤԼ��ʽ.Text, ".") + 1) & "'", "NULL") & "," & _
                                 IIf(mTy_Para.bln�Һ����ɶ���, 1, 0) & ","
                        
                        '�����id_In   ����Ԥ����¼.�����id%Type := Null,
                        strSQL = strSQL & "" & IIf(mCurCardPay.Have�Һŷ� And mCurCardPay.lngҽ�ƿ����ID <> 0 And mCurCardPay.bln���ѿ� = False, mCurCardPay.lngҽ�ƿ����ID, "NULL") & ","
                        '���㿨���_In ����Ԥ����¼.���㿨���%Type := Null,
                        strSQL = strSQL & "" & IIf(mCurCardPay.Have�Һŷ� And mCurCardPay.lngҽ�ƿ����ID <> 0 And mCurCardPay.bln���ѿ�, mCurCardPay.lngҽ�ƿ����ID, "NULL") & ","
                        '����_In       ����Ԥ����¼.����%Type := Null,
                        strSQL = strSQL & "" & IIf(mCurCardPay.Have�Һŷ� And mCurCardPay.strˢ������ <> "", "'" & mCurCardPay.strˢ������ & "'", "NULL") & ","
                        '������ˮ��_In ����Ԥ����¼.������ˮ��%Type := Null,
                        strSQL = strSQL & " NULL,"
                        '����˵��_In   ����Ԥ����¼.����˵��%Type := Null,
                        strSQL = strSQL & " NULL,"
                        '������λ_In   ����Ԥ����¼.������λ%Type := Null
                        strSQL = strSQL & " NULL,"
                        '  ��������_In   Number:=0
                        strSQL = strSQL & IIf(mbln�Ӻ�, "1", "0") & ","
                        '  ����_IN       ���˹Һż�¼.����%type:=null,
                        strSQL = strSQL & IIf(mintInsure = 0, "NULL", mintInsure) & ","
                        '  ����ģʽ_IN   NUMBER :=0,
                        strSQL = strSQL & IIf(mPatiChargeMode = EM_�����ƺ����, 1, 0) & ","
                        '  ���ʷ���_IN Number:=0,
                        strSQL = strSQL & IIf(mRegistFeeMode = EM_RG_����, 1, 0) & ","
                        '  �˺�����_IN Number:=1,
                        strSQL = strSQL & IIf(mTy_Para.blnReuseCancelNO, 1, 0) & ","
                        '  ��Ԥ������ids_In Varchar2 := Null
                        strSQL = strSQL & "'" & lng����ID & "," & mstr���˼���IDs & "'," '79868,Ƚ����,2015-6-15,ʹ�ü���Ԥ��
                        '  �������˷ѱ�_In  Number := 0,
                        strSQL = strSQL & 0 & ","
                        '  ԤԼ˳���_In    �ٴ�������ſ���.ԤԼ˳���%Type := Null,
                        strSQL = strSQL & "Null,"
                        '  ������������_In  Number := 0,
                        strSQL = strSQL & "0,'"
                        '  �շѵ�_In        ���˹Һż�¼.�շѵ�%Type := Null
                        strSQL = strSQL & str����NO & "')"
                        
                        Call zlAddArray(cllPro, strSQL)
                        
                        If Trim(IIf(txt�ű�.Text = "+", "", txt�ű�.Text)) <> "" And k = 1 Then
                            If Nvl(vsfPlan.TextMatrix(vsfPlan.Row, GetCol("ҽ��"))) = "" Then blnNoDoc = True
                            strSQL = "zl_���˹ҺŻ���_Update("
                            '  ҽ������_In   �ҺŰ���.ҽ������%Type,
                            strSQL = strSQL & IIf(blnNoDoc, "Null,", "'" & strҽ������ & "',")
                            '  ҽ��id_In     �ҺŰ���.ҽ��id%Type,
                            strSQL = strSQL & "" & IIf(blnNoDoc, "0,", ZVal(lngҽ��ID) & ",")
                            '  �շ�ϸĿid_In ������ü�¼.�շ�ϸĿid%Type,
                            strSQL = strSQL & "" & Val(Nvl(mrsItems!��ĿID)) & ","
                            '  ִ�в���id_In ������ü�¼.ִ�в���id%Type,
                            strSQL = strSQL & "" & IIf(Val(Nvl(mrsItems!ִ�п���ID)) = 0, mlng�Һſ���ID, Val(Nvl(mrsItems!ִ�п���ID))) & ","
                            '  ����ʱ��_In   ������ü�¼.����ʱ��%Type,
                            strSQL = strSQL & "" & str����ʱ�� & ","
                            '  ԤԼ��־_In   Number := 0  --�Ƿ�ΪԤԼ����:0-��ԤԼ�Һ�; 1-ԤԼ�Һ�,2-ԤԼ����,3-�շ�ԤԼ
                            strSQL = strSQL & Decode(mbytMode, 1, 1, 2, 2, IIf(chkBooking.Value = 1, 3, 0)) & ","
                            '  ����_In       �ҺŰ���.����%Type := Null
                            strSQL = strSQL & "'" & IIf(txt�ű�.Text = "+", "", txt�ű�.Text) & "',0,"
                            strSQL = strSQL & "" & ZVal(lng��¼ID) & ")"
                            Call zlAddArray(cllProAfter, strSQL)
                        End If
                        '����ҽ��վ�Һ�ʱ,������ֽ�֧�������ɻ��۵�,��ʱӦ��/ʵ����дΪ0,ժҪ��дΪ�Һŵ��ݺ�
                        If blnPrice And dblTotal(0) <> 0 Then
                            strSQL = _
                            "zl_���ﻮ�ۼ�¼_Insert('" & str����NO & "'," & k & "," & lng����ID & ",NULL," & _
                                     IIf(str����� = "", "NULL", str�����) & ",'" & NeedCode(cbo���ʽ.Text) & "'," & _
                                     "'" & txtPatient.Text & "','" & NeedName(cbo�Ա�.Text) & "','" & str���� & "'," & _
                                     "'" & str�ѱ� & "',NULL," & mlng�Һſ���ID & "," & _
                                     IIf(mblnStation, mlng�Һſ���ID, UserInfo.����ID) & ",'" & UserInfo.���� & "'," & IIf(mrsItems!���� = 2, 1, "NULL") & "," & _
                                     mrsItems!��ĿID & ",'" & mrsItems!��� & "','" & mrsItems!���㵥λ & "'," & _
                                     "NULL,1," & mrsItems!���� & ",NULL," & IIf(mrsItems!ִ�п���ID = 0, mlng�Һſ���ID, mrsItems!ִ�п���ID) & "," & IIf(int�۸񸸺� = k, "NULL", int�۸񸸺�) & "," & _
                                     mrsInComes!������ĿID & ",'" & mrsInComes!�վݷ�Ŀ & "'," & mrsInComes!���� & "," & _
                                     mrsInComes!Ӧ�� & "," & mrsInComes!ʵ�� & "," & str����ʱ�� & "," & str�Ǽ�ʱ�� & ",NULL,'" & UserInfo.���� & "','�Һ�:" & strNO & "')"
                            Call zlAddArray(cllPro, strSQL)
                        End If
                    End If
                    k = k + 1
                    mrsInComes.MoveNext
                Next
                mrsItems.MoveNext
            Next
        End If
 
        cmdOK.Enabled = False      '��ֹ��ӡ�������ô�ӡ���ķ�ģ̬���弰ҽ�������ӳ�
        'ִ�д���
        Err = 0: On Error GoTo ErrFirt:
        If cllPro.Count > 0 Then
            '����:31187 �������д����������
            Err = 0: On Error GoTo ErrFirt:
            zlExecuteProcedureArrAy cllPro, Me.Caption, True
            
            '�����
            If lng����ID <> 0 Then
                strSQL = "Select Sum(���ʽ��) As ���ý�� From ������ü�¼ Where ��¼����=4 And ����ID=[1]"
                Set rsCheck = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID)
                If Not rsCheck.EOF Then
                    dbl���ý�� = Val(Nvl(rsCheck!���ý��))
                    strSQL = "Select Sum(��Ԥ��) As ���ʽ�� From ����Ԥ����¼ Where ����ID=[1]"
                    Set rsCheck = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID)
                    If Not rsCheck.EOF Then
                        If dbl���ý�� <> Val(Nvl(rsCheck!���ʽ��)) Then
                            gcnOracle.RollbackTrans
                            MsgBox "������Ϣ�������Ϣ���治һ�£���������ȡ��������!", vbInformation, gstrSysName
                            cmdOK.Enabled = True: RestorePay: Exit Sub
                        End If
                    Else
                        If dbl���ý�� <> 0 Then
                            gcnOracle.RollbackTrans
                            MsgBox "������Ϣ�������Ϣ���治һ�£���������ȡ��������!", vbInformation, gstrSysName
                            cmdOK.Enabled = True: RestorePay: Exit Sub
                        End If
                    End If
                End If
            End If

            Err = 0: On Error GoTo errH:
            blnTrans = True
            If curOneCard <> 0 And mRegistFeeMode <> EM_RG_���� Then
                If Not (curOneCard = cur���� And cur���� <> 0) Then    '��ֻ�ǿ���ʱ
                    If Not mobjICCard.PaymentSwap(curOneCard - cur����, dblOneCardBalance, intCardType, Val("" & mrsOneCard!ҽԺ����), strCardNo, strTransFlow, lng����ID, lng����ID) Then
                        gcnOracle.RollbackTrans
                        MsgBox "һ��ͨ����Һŷ�ʧ��", vbInformation, gstrSysName
                        RestorePay
                        cmdOK.Enabled = True: Exit Sub
                    Else
                        strSQL = "zl_һ��ͨ����_Update(" & lng����ID & ",'" & mrsOneCard!���㷽ʽ & "','" & strCardNo & "','" & intCardType & "','" & strTransFlow & "'," & dblOneCardBalance & ")"
                        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
                    End If
                End If

                If cur���� <> 0 Then
                    dblOneCardBalance = 0
                    strTransFlow = ""
                    If Not mobjICCard.PaymentSwap(cur����, dblOneCardBalance, intCardType, Val("" & mrsOneCard!ҽԺ����), strCardNo, strTransFlow, lngCard����ID, lng����ID) Then
                        gcnOracle.RollbackTrans
                        MsgBox "һ��ͨ���㿨��ʧ��", vbInformation, gstrSysName
                        RestorePay
                        cmdOK.Enabled = True: Exit Sub
                    Else
                        strSQL = "zl_һ��ͨ����_Update(" & lngCard����ID & ",'" & mrsOneCard!���㷽ʽ & "','" & strCardNo & "','" & intCardType & "','" & strTransFlow & "'," & dblOneCardBalance & ")"
                        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
                    End If
                End If
            End If

            'ҽ���Ķ�
            blnNotCommit = False
            If mintInsure <> 0 And mstrYBPati <> "" Then
                strAdvance = ""
                If mRegistFeeMode = EM_RG_���� Or mPatiChargeMode = EM_�����ƺ���� Then
                    strAdvance = IIf(mPatiChargeMode = EM_�����ƺ����, "1", "0") & "|" & IIf(mRegistFeeMode = EM_RG_����, "1", "0") & "|" & strNO
                End If
                If Not gclsInsure.RegistSwap(lng����ID, cur����, mintInsure, strAdvance) Then
                    gcnOracle.RollbackTrans: cmdOK.Enabled = True: RestorePay: Exit Sub
                End If
                blnNotCommit = True
            End If
            zlExecuteProcedureArrAy cllProAfter, Me.Caption, True, True
            Set cllCardPro = New Collection: Set cllTheeSwap = New Collection
            If Not mPatiChargeMode = EM_�����ƺ���� Then
                If zlInterfacePrayMoney(lngCard����ID, lng����ID, cllCardPro, cllTheeSwap, dblThreeSwap) = False Then
                    gcnOracle.RollbackTrans: If cmdOK.Enabled = False Then cmdOK.Enabled = True: RestorePay: Exit Sub
                End If
                '������������
                zlExecuteProcedureArrAy cllCardPro, Me.Caption, True, True
            End If
            gcnOracle.CommitTrans
            
            Call zlExcPatiInfo(lng����ID, lng����ID, strNO)
            
            Err = 0: On Error GoTo OthersCommit:
            zlExecuteProcedureArrAy cllTheeSwap, Me.Caption, True, False
OthersCommit:
            gcnOracle.CommitTrans
            'д������
            If bln���� And mCurSendCard.bln�Ƿ�д�� Then Call WriteCard(lng����ID)
            If mintInsure > 0 Then Call gclsInsure.BusinessAffirm(����Enum.Busi_RegistSwap, True, mintInsure)
            Dim strOutPut As String
            Call zlExcuteUploadSwap(lng����ID, strOutPut, mobjICCard)
            
            blnTrans = False
            On Error GoTo 0
            'ҽ����������������
            If mintInsure <> 0 And mstrYBPati <> "" And Not blnPrice And mRegistFeeMode <> EM_RG_���� Then
                '�����ҽ������,��Ҫ���»�ȡ���ν�������ս��
                curӦ�� = GetActualCash(lng����ID)
                If gblnLED And mbytMode <> 1 And mbytInState = 0 Then
                    zl9LedVoice.Speak "#21 " & Format(curӦ��, "0.00")
                    txt�Ҳ�.Text = Format(Val(txt�ɿ�.Text) - curӦ��, "0.00")
                End If
            End If
        End If
        If str���� <> "" Then
            Call zlCommitPlugInpati(str����)
        End If
        '��Ϣ����:
        Call SendMsgModule(strNO)
        '��ӡ����
        If mbytMode <> 1 And Not blnNoPrint Then
RePrint:
            Dim strNotValiedNos As String
            If Not gobjTax Is Nothing And gblnTax Then
                Call TaxInterface(1, "'" & strNO & "'", "")
            Else
                If mRegistFeeMode <> EM_RG_���� Then
                    blnEnterPrint = True
                    Call frmPrint.ReportPrint(1, strNO, "", mlng����ID, mlngShareUseID, txtFact.Text, Datsys, txt�ɿ�.Text, txt�Ҳ�.Text, , mintInsure <> 0 And MCPAR.ҽ���ӿڴ�ӡƱ��, False, mstrUseType)
                    If gblnBill�Һ� Then
                        If zlIsNotSucceedPrintBill(4, strNO, strNotValiedNos) = True Then
                            If MsgBox("�Һŵ���Ϊ[" & strNotValiedNos & "]Ʊ�ݴ�ӡδ�ɹ�,�Ƿ����½���Ʊ�ݴ�ӡ!", vbYesNo + vbDefaultButton1 + vbQuestion, gstrSysName) = vbYes Then GoTo RePrint:
                        End If
                    End If
                End If
            End If
        ElseIf blnPrintBooking And mbytMode = 1 Then
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1111_1", Me, "NO=" & strNO, 2)
        End If
        
        If mbytMode <> 1 And gblnPrintCase Then
            '�������˵���� ����ţ�42452 �޸���:����
            If chk������.Value = 1 And blnPati = True And bytType = 1 Then
                Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1111_2", Me, "����ID=" & lng����ID, 2)
            ElseIf chk������.Value = 1 Or Trim(txt�ű�.Text) = "+" Then
                Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1111_2", Me, "����ID=" & lng����ID, 2)
            End If
        End If
        
        If blnSlipPrint Then
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1111_3", Me, "NO=" & strNO, 2)
            If Not blnEnterPrint Then
                strSQL = "Zl_ƾ����ӡ��¼_Update(4,'" & strNO & "',1,'" & UserInfo.���� & "','��Ʊ��:" & txtFact.Text & "')"
                zlDatabase.ExecuteProcedure strSQL, "ƾ����ӡ��¼"
            End If
        End If
        
        If blnCodePrint Then
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1111_2", Me, "����ID=" & lng����ID, 2)
        End If
        
        If CreatePlugInOK(mlngModul) Then
            On Error Resume Next
            strSQL = "Select ID From ���˹Һż�¼ Where no=[1] And Rownum<2"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
            If Not rsTemp.EOF Then Call gobjPlugIn.OutPatiRegisterAfter(lng����ID, Nvl(rsTemp!ID))
            Err.Clear
        End If
        
        cmdOK.Enabled = True
        'ԤԼ���պ��˳�
        If mbytMode = 2 Then
            If Not gblnBill�Һ� And Not blnNoPrint And mRegistFeeMode <> EM_RG_���� Then
                If gblnSharedInvoice Then
                    zlDatabase.SetPara "��ǰ�շ�Ʊ�ݺ�", txtFact.Text, glngSys, 1121
                Else
                    zlDatabase.SetPara "��ǰ�Һ�Ʊ�ݺ�", txtFact.Text, glngSys, mlngModul
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

        '���뵥����ʷ��¼(�������͵���)
        If strNO <> "" Then
            For i = 0 To cboNO.ListCount - 1
                strNO = strNO & "," & cboNO.List(i)
            Next
            cboNO.Clear
            For i = 0 To UBound(Split(strNO, ","))
                cboNO.AddItem Split(strNO, ",")(i)
                If i = 9 Then Exit For    'ֻ��ʾ10��
            Next
            If cboNO.ListCount > 0 Then cboNO.ListIndex = 0
        End If
        blnNew = True: strFact = txtFact.Text
        If blnNoPrint Then blnNew = False    '����ӡʱ,���ϸ���Ƶ�Ʊ�ݲ����Ӻ�
    End If
    gblnOk = True
    Call SetControlChk
    '���没�˼��ۼ���Ϣ������:����Ҫ����ɿ��Ž���,��ǰδ��ɿ�,�����Ƿ�ҽ������,����������,
    '���ұ��ز���Ҫ��������(����ClearBill�е���SetPatiInfoEnabledʱ���������)
    '���˺�:26602
    ' �����Ӷ�ҽ�����˽��������Һ�,ҽ����������Ϊ:
    '   1.����Ҫ������ɿ������ֹ�����շ�
    '   2.��Ҫ����:support�����Һ�
    Dim blnClearInsure As Boolean
    blnClearInsure = True
    If mintInsure <> 0 And mstrYBPati <> "" Then
        bln���� = gclsInsure.GetCapability(support�����Һ�, lng����ID, mintInsure)
        bln���� = mTy_Para.byt�ɿʽ = 1 And mbytMode <> 1 And Val(txt�ɿ�.Text) = 0 And txtPatient.Text <> "" And bln����
        blnClearInsure = Not bln����
        Dim cur�Ҳ� As Currency, cur�ɿ� As Currency

        If blnCall�����Һ� Then
            If mstr�����Һ�_�Һ�NO <> "" Then mstr�����Һ�_�Һ�NO = Mid(mstr�����Һ�_�Һ�NO, 2)
            If mstr�����Һ�_���￨NO <> "" Then mstr�����Һ�_���￨NO = Mid(mstr�����Һ�_���￨NO, 2)
            txt����Ӧ��.Visible = False: lblӦ��.Visible = False: lbl�ɿ�.Visible = False: txt�ɿ�.Visible = False: lbl�Ҳ�.Visible = False: txt�Ҳ�.Visible = False
            lblSum.Visible = False: txt�ϼ�.Visible = False
            picTotal.Visible = True
            If frmYbPayFeeShow.zlShowPayWindows(Me, gclsInsure, gblnLED, txtPatient.Text, cbo�Ա�.Text, txt����.Text & cbo���䵥λ.Text, lng����ID, mintInsure, mstr�����Һ�_�Һ�NO, mstr�����Һ�_���￨NO, mcur�ϼ� + GetRegistMoney, mcurӦ�� + curӦ��, cur�ɿ�, cur�Ҳ�) Then
                txt����Ӧ��.Text = Format(mcurӦ�� + curӦ��, "0.00")
                txt�ɿ�.Text = Format(cur�ɿ�, "0.00")
                txt�Ҳ�.Text = Format(cur�Ҳ�, "0.00")
                bln���� = False
            End If
            txt����Ӧ��.Visible = True: lblӦ��.Visible = True: lbl�ɿ�.Visible = True: txt�ɿ�.Visible = True: lbl�Ҳ�.Visible = True: txt�Ҳ�.Visible = True
            lblSum.Visible = True: txt�ϼ�.Visible = True
            picTotal.Visible = False
        End If
    Else
        bln���� = mTy_Para.byt�ɿʽ = 1 And mbytMode <> 1 And Val(txt�ɿ�.Text) = 0 And mstrYBPati = "" And txtPatient.Text <> ""
    End If
    
    If Not mbln�����Һ� Then
        mcur�ϼ� = 0: mcurӦ�� = 0: mint�Һ��� = 0
        mstrPrePati = "": mstr�����Һ�_�Һ�NO = "": mstr�����Һ�_���￨NO = ""
        lng����ID = 0
        mblnFinishReg = True
        Call ClearBill(, Not blnNoPrint)
        mblnFinishReg = False
    Else
        If Not blnPrice Then
            mcur�ϼ� = mcur�ϼ� + GetRegistMoney
            mcurӦ�� = mcurӦ�� + cur����
        End If
        mstrPrePati = txtPatient.Text
        '
        Call ClearBill(False, Not blnNoPrint, False)  '���ݲ���,�����Ҫ��������,���ߺű𲻽�����,��������������
        mint�Һ��� = mint�Һ��� + 1
        '���˺�:�����ҽ������,��Ҫ���»�ȡ���
        If mintInsure <> 0 And mstrYBPati <> "" Then
            mcur������� = gclsInsure.SelfBalance(lng����ID, CStr(Split(mstrYBPati, ";")(1)), 10, mcur����͸֧, mintInsure)
            stbThis.Panels(3).Text = "�����ʻ����:" & Format(mcur�������, "0.00")
            mdbl������� = mcur�������
        End If
    End If

    'ˢ��Ʊ�ݺ�
    If mbytMode <> 1 And Not mblnStation And Not blnPrice Then
        If blnNoPrint = False Then Call RefreshFact
    End If

    '�����������Ϣ���˻�ս���Ϣ�Ĳ�����һ�ŵ���ʱ����������Ϣ(���ز���Ҫ��������ʱ)
    If lng����ID > 0 And chkCancel.Value = 0 And txtPatient.Enabled Then
        Call GetPatient(IDKind.GetCurCard, "-" & lng����ID, False)
    End If

    timPlan.Enabled = glngInterval > 0 And mbytInState = 0 And (mbytMode = 0 Or mbytMode = 1)

    'ˢ�µ�ǰ���,ClearBill���ѵ���txt�ű�_change
    If txt�ű�.Enabled And txt�ű�.Visible Then txt�ű�.SetFocus
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
    '����:31634
    If mintInsure > 0 And blnNotCommit Then Call gclsInsure.BusinessAffirm(����Enum.Busi_RegistSwap, False, mintInsure)
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    timPlan.Enabled = glngInterval > 0 And mbytInState = 0 And (mbytMode = 0 Or mbytMode = 1)
    cmdOK.Enabled = True
    mbln�Ӻ� = False
    Exit Sub
ErrGo:
    If ErrCenter() = 1 Then
        Resume
    End If
    timPlan.Enabled = glngInterval > 0 And mbytInState = 0 And (mbytMode = 0 Or mbytMode = 1)
    cmdOK.Enabled = True
End Sub

Private Function GetPatiInforXML() As String
    Dim strPatiInforXML As String, str���� As String, str�������� As String, str���֤�� As String
    
    strPatiInforXML = strPatiInforXML & "<XM>" & Trim(txtPatient.Text) & "</XM>" & vbCrLf
    strPatiInforXML = strPatiInforXML & "<XB>" & NeedName(cbo�Ա�.Text) & "</XB>" & vbCrLf
    str���� = Trim(txt����.Text)
    If IsNumeric(str����) Then str���� = str���� & cbo���䵥λ.Text
    strPatiInforXML = strPatiInforXML & "<NL>" & str���� & "</NL>" & vbCrLf
    If IsDate(txt��������.Text) Then
        str�������� = Format(txt��������.Text & IIf(txt����ʱ�� = "__:__", "", " " & txt����ʱ��.Text), "yyyy-mm-dd HH:mm:ss")
    End If
    strPatiInforXML = strPatiInforXML & "<CSRQ>" & str�������� & "</CSRQ>" & vbCrLf
    strPatiInforXML = strPatiInforXML & "<YBH>" & mobjfrmPatiInfo.txtPatiMCNO(0).Text & "</YBH>" & vbCrLf
    If txtIDCard.Text <> "" And txtIDCard.Visible Then str���֤�� = Trim(txtIDCard.Text)
    strPatiInforXML = strPatiInforXML & "<SFZH>" & str���֤�� & "</SFZH>"
    strPatiInforXML = strPatiInforXML & "<YSXM>" & NeedName(cboҽ��.Text) & "</YSXM>"
    
    GetPatiInforXML = strPatiInforXML
End Function

Private Sub SetControlChk()
    mstrPreNO = txt�ű�.Text
    cboNO.Tag = ""
    If chkCancel.Value = 1 Then chkCancel.Value = 0
    If chkPrint.Value = 1 Then chkPrint.Value = 0
    If chkBooking.Value = 1 Then
        chkBooking.Tag = "����"
        chkBooking.Value = 0
        chkBooking.Tag = ""
    End If
End Sub

Private Sub zlExcPatiInfo(ByVal lng����ID As Long, ByVal lng����ID As Long, ByVal strNO As String)
    Dim cllPro As Collection, Datsys As Date
    Dim strSQL As String, rsTemp As ADODB.Recordset
    '82072:���ϴ�,2015/1/23,Ѫ�ͺ�RH����һ���о���ID�ļ�¼
    '.,���Խ�������Ϣ�ӱ�ת�Ƶ�����
    
    On Error GoTo Errhand
    If lng����ID > 0 And Not ((mbytMode = 2 Or mbytMode = 0 And mbytInState = 0 And mstrNoIn <> "") And mTy_Para.blnԤԼ����ȷ���Һŷ� = False) Then
        Set cllPro = New Collection
        Datsys = zlDatabase.Currentdate
        If lng����ID = 0 Then
            strSQL = "Select ID as ����ID From ���˹Һż�¼ Where ��¼״̬ = 1 And NO=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
            If rsTemp.RecordCount > 0 Then lng����ID = Nvl(rsTemp!����ID, 0)
        End If
        Call mobjfrmPatiInfo.Add�����������Ϣ(lng����ID, cllPro, lng����ID)
        '���没����Ϣ�е�֤��
        Call mobjfrmPatiInfo.AddCertificate(lng����ID, cllPro, Datsys)
        zlExecuteProcedureArrAy cllPro, Me.Caption
    End If
    Exit Sub
Errhand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function WriteCard(lng����ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:д��
    '���:lng����ID - ����ID
    '����:����
    '����:56599
    '����:2012-12-17 15:07:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strExpend As String
    On Error GoTo ErrHandl:
    '115168:���ϴ���2017/12/13�����淢����ҽ�ƿ�����
    If mCurSendCard.lng�����ID = 0 Then mCurSendCard = gCurSendCard
    If mCurSendCard.bln�Ƿ�д�� = False Then Exit Function
    If Not gobjSquare.objSquareCard Is Nothing Then
        WriteCard = gobjSquare.objSquareCard.zlBandCardArfter(Me, mlngModul, mCurSendCard.lng�����ID, lng����ID, strExpend)
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
           mrsOneCard.Filter = "����='" & strName & "'"
           If mrsOneCard.RecordCount > 0 Then
                strName = mrsOneCard!���㷽ʽ
                If NeedName(cbo���㷽ʽ) <> strName Then zlControl.CboLocate cbo���㷽ʽ, strName
           End If
        End If
    End If
End Sub

Private Function RefreshFact() As Boolean
    'ˢ�·�Ʊ��
    '˵����
    '   24363:��Ҫ�ǽ���Զ����ɵĺ��Ƿ��û����ģ�
    '    1.���ĵ�Ʊ�ݺ���Ҫ����Ƿ��ظ����ظ���ֱ�ӷ��ز����ķ�Ʊ��
    '    2.���������������ĵ�����£�����Ƿ��ظ�������ظ����Զ�ȡ��һ�����룡
    Dim strFact As String
    
    If mblnStationPrice Then Exit Function
    'lblFact.tag��Ҫ�Ǽ�鷢Ʊ���Ƿ��ֹ������.�ֹ������,��Ʊ��Ϊ��,�������Զ������ķ�Ʊ��
    If (lblFact.Tag <> "" And txtFact.Text <> "") Or Trim(txtFact.Text) = "" Then
        If gblnBill�Һ� Then
            mlng����ID = CheckUsedBill(IIf(gblnSharedInvoice, 1, 4), IIf(mlng����ID > 0, mlng����ID, glng�Һ�ID), , IIf(mblnStartFactUseType, mstrUseType, ""))
            If mlng����ID <= 0 Then
                Select Case mlng����ID
                    Case 0 '����ʧ��
                    Case -1
                        MsgBox "��û�����ú͹��õĹҺ�Ʊ��,��������һ��Ʊ�ݻ����ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
                    Case -2
                        MsgBox "���صĹ���Ʊ���Ѿ�����,��������һ��Ʊ�ݻ��������ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
                End Select
                txtFact.Text = "": txtFact.Tag = "":  Exit Function
            End If
            
            '�ϸ�ȡ��һ������
            txtFact.Text = GetNextBill(mlng����ID)
        Else
            '��ɢ��ȡ��һ������
            If gblnSharedInvoice Then
                strFact = zlDatabase.GetPara("��ǰ�շ�Ʊ�ݺ�", glngSys, 1121)
            Else
                strFact = zlDatabase.GetPara("��ǰ�Һ�Ʊ�ݺ�", glngSys, mlngModul)
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
        strSQL = " And  (B.���￨�� = [1] Or B.Ic���� = [1] Or B.���֤�� = [1]" & IIf(IsNumeric(strInput), " Or B.����� = [1]", "") & ")"
    End If
    
    strSQL = "" & _
    "Select Min(A.NO) NO" & vbNewLine & _
    "From ������ü�¼ A, ������Ϣ B" & vbNewLine & _
    "Where A.��¼���� = 4 And A.��¼״̬ = 0 And A.����id = B.����id(+)  " & _
                IIf(mTy_Para.intԤԼʧЧ���� > 0, "  And A.����ʱ�� between trunc(sysdate) and  trunc(sysdate)+1-1/24/60/60 ", _
            "  And ((nvl(A.�Ӱ��־,0) =0 And A.����ʱ�� > Trunc(Sysdate) - [2]) or  (nvl(A.�Ӱ��־,0) =1 And A.����ʱ�� > Trunc(Sysdate) - [3])  ) ") & strSQL
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
    '���ܣ����ý���ԤԼʱ��״̬,�Լ�״̬�ָ�
    '���ƣ����˺�
    '���ڣ�2010-07-14 10:27:10
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    If dkpMain.Panes(1).Hidden = False Then '���ںű��б�,����ѡ��ű�
        picPlan.Enabled = Not blnReceive
        mcbrToolBar.Controls.Find(xtpControlButton, conMenu_View_Refresh).Enabled = Not blnReceive   'ˢ��
        mcbrToolBar.Controls.Find(xtpControlButton, 2605).Enabled = Not blnReceive   'Ԥ������
        mcbrToolBar.Controls.Find(xtpControlButton, 2604).Enabled = Not blnReceive   'Ԥ������
    End If
    
    cboNO.Locked = blnReceive       '���ݺ�
        
    chkPrint.Visible = Not blnReceive   '�ش�
    chkCancel.Visible = Not blnReceive    '�˺�
    chkBooking.Visible = Not blnReceive And InStr(1, mstrPrivs, ";ԤԼ�Һ�;") > 0 'ԤԼ
    If mobjCommunity Is Nothing Then
        cmdComminuty.Visible = False
    Else
        cmdComminuty.Visible = Not blnReceive  '��������
    End If
    cmdLookup.Visible = Not blnReceive          '���Ҳ���
    cmdMore.Visible = True            '�������Ĳ�����Ϣ
    lblҽ�����.Visible = True
    cboҽ�����.Visible = True
    
    cmdCard.Visible = InStr(1, mstrPrivs, ";�󶨿���;") > 0   '�󶨿���:31182:Not blnReceive And
    
    If mbytMode = 0 And mbytInState = 0 Then
        cmdYb.Visible = True
    Else
        cmdYb.Visible = blnReceive   'ԤԼ����ʱ,����ˢҽ�� '����:31182
    End If
    
    lblIDCard.Visible = True
    If IDKind֤��.IDKind = IDKind֤��.GetKindIndex("���֤��") Then
        txtIDCard.Visible = True: txt֤��.Visible = False
    Else
        txtIDCard.Visible = False: txt֤��.Visible = True
    End If
    stbThis.Visible = True
    
    txt�ű�.Enabled = Not blnReceive '����ʱ�����ٸ��ĺű�,����������
    cbo���㷽ʽ.Enabled = blnReceive Or gbln���㷽ʽ
    
    '55985:������,2014-02-17,ԤԼ����ʱ�����޸ķѱ�͹�����
    If InStr(1, mstrPrivs, ";�����޸ķѱ�;") > 0 And mTy_Para.blnԤԼ����ȷ���Һŷ� = True Then
        cbo�ѱ�.Enabled = True
        chk������.Enabled = True
    Else
        cbo�ѱ�.Enabled = Not blnReceive '����ѡ����㷽ʽ
        chk������.Enabled = Not blnReceive '����ʱ�����ټ��ղ�����
    End If
    
    txtSN.Locked = blnReceive
    
    If blnReceive Then
         'ȷ����ſ���
         If GetCol("��ſ���") >= 0 Then
            txtSN.Enabled = vsfPlan.TextMatrix(vsfPlan.Row, GetCol("��ſ���")) <> ""
        End If
        If Not txtSN.Enabled And txtSN.Text <> "" Then txtSN.Text = ""
    End If
    Call zlPatiMoveCmdCtrl
    
End Sub

Private Function ReadBooking(ByVal strNO As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ���ȡԤԼ�Һŵ�����
    '��Σ�strNO-ԤԼ�Һŵ��ݺ�
    '����:��ȡ�ɹ�,����True,���򷵻�False
    '���ƣ����˺�
    '���ڣ�2010-07-16 16:21:45
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    On Error GoTo errH
    '��ԤԼ��,������
    If Not (chkCancel.Value = 0 And chkPrint.Value = 0 And chkBooking.Value = 0 And Not mblnStation) Then Exit Function
    mstrNoIn = strNO
    If mstrNoIn = "" Then
        MsgBox "û���ҵ������յ�ԤԼ�Һŵ���", vbInformation, gstrSysName
       ' mblnUnload = True
        cboNO.SetFocus: Exit Function
    End If
    
    mblnReadBooking = True
    If ReadBill(mstrNoIn, True) = False Then mblnReadBooking = False: Exit Function
    If mblnUnload Then mstrNoIn = "": Exit Function
    strSQL = "Select �����¼ID From ���˹Һż�¼ Where NO=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrNoIn)
    If Not rsTemp.EOF Then
        mlng��¼ID = Val(Nvl(rsTemp!�����¼ID))
    Else
        mlng��¼ID = 0
    End If
    
    If Not txt����ʱ��.Text Like "____*" Then
        dtpAppointmentDate.Value = CDate(txt����ʱ��.Text) '��ʱû���Զ�����change�¼�
    End If
    If txt�����.Text = "" And gbln�Զ������ Then
        txt�����.Text = zlGet�����
    End If
    
    chkShowAll.Value = 1
    Call ShowPlans
    mblnReadBooking = False
    
    
    '��λ�ű�,���û�����������
    For i = 1 To vsfPlan.Rows - 1
        If Val(vsfPlan.TextMatrix(i, GetCol("��¼ID"))) = mlng��¼ID Then
            mblnChangeByCode = True
            vsfPlan.Row = i
            mblnChangeByCode = False
            vsfPlan_EnterCell
            Exit For
        End If
    Next
    If mbytMode = 2 Or mbytMode = 0 And mbytInState = 0 And mstrNoIn <> "" Then
        'ԤԼ����
        If CheckIsPrice Then
            Call SetUndisplayBalance
        Else
            Call SetShowBalance
        End If
    End If
    
    If mbln������ And InStr(mstrPrivs, ";��������;") = 0 And txt�����.Text = "" Then
        MsgBox "�úű�Ҫ������˽������ﲡ��������û�н���������Ȩ�ޡ����ܽ��ա�", vbInformation, gstrSysName
        mblnUnload = True: Exit Function
    End If
    cboNO.Text = mstrNoIn
    Call SetReceiveState(True)
    
    
    If gbytInvoice <> 0 Then Call RefreshFact
    If cbo���㷽ʽ.Enabled And cbo���㷽ʽ.Visible Then cbo���㷽ʽ.SetFocus
    If txt�ű�.Text <> "" Then
         Call ShowRegistFromInput
    End If
    '68216
    If Val(txtSN.Tag) <> 0 Then '
        txtSN.Text = txtSN.Tag
        locateSnByʱ�� Val(txtSN.Tag), True
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
    '���ܣ����ݿ��,����ԤԼ�ҺŽ���С����,��ѡ������ԤԼ�Һŵ�
    '���ƣ����˺�
    '���ڣ�2010-07-16 16:34:39
    '˵����31182
    '------------------------------------------------------------------------------------------------------------------------
    Dim rsInfor As ADODB.Recordset
    Dim strOutNo As String
    Dim frmNew As frmSelRegist
    Dim blnExit As Boolean
    If mbytInState = 1 Then Exit Sub
    If InStr(1, mstrPrivs, ";����ԤԼ;") = 0 Then Exit Sub
    If Not (chkCancel.Value = 0 And chkPrint.Value = 0 And chkBooking.Value = 0 And Not mblnStation) Then Exit Sub
    If mbytMode = 1 Or mbytMode = 2 Then Exit Sub
    Call CloseIDCard    '47007
    Set frmNew = New frmSelRegist
    If frmNew.ShowRegist(Me, mstrPrivs, mblnOlnyBJYB, mTy_Para.intԤԼʧЧ����, strOutNo, rsInfor) = False Then
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
    '����:�ر�������������
    '����:���˺�
    '����:2012-03-09 16:26:40
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
    '����:��ʼ���µĿ�����
    '����:���˺�
    '����:2012-03-09 16:28:23
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
    Dim str����NO As String, strNO As String
    Dim blnEnableDel As Boolean, i As Long
    If KeyAscii = Asc("/") And Trim(cboNO.Text) = "" Then
        'ԤԼ����ʱ,������ݺ��������"/",���Զ�����С����,��ԤԼ�Һ���"
        KeyAscii = 0:
        Call ShowBookSeled
        Exit Sub
    End If
    
      If KeyAscii = 13 And Trim(cboNO.Text) <> "" Then
        KeyAscii = 0
        cboNO.Text = Trim(cboNO.Text)
        
        If chkCancel.Value = 0 And chkPrint.Value = 0 And chkBooking.Value = 0 And Not mblnStation Then
            'A.����ԤԼ�Һŵ�
            'cboNO.Text = GetFullNO(cboNO.Text, 12) '�����Զ���ȫ���ݺ�,��Ϊ����Ŀ����������,���֤��
            mblnRegReceiveByNo = True '�����:57423
            strNO = cboNO.Text
            Call ClearBill
            '����:38503
            If InStr(1, mstrPrivs, ";����ԤԼ;") = 0 Then Exit Sub
            mstrNoIn = GetBookingNO(strNO)
            Call ReadBooking(mstrNoIn)        '����ҪmstrNoInֵ
        ElseIf chkCancel.Value = 1 Or chkPrint.Value = 1 Then
            'B.�˺Ż��ش�
            cboNO.Text = GetFullNO(cboNO.Text, 12)
            strNO = cboNO.Text
            '�Ƿ���ת������ݱ���,ע��˴����ܼ�frmRegistFilter.mblnNOMoved�����ж�,��Ϊ�շѴ��ں�ҽ������վ���ڻ�����������.
            If zlDatabase.NOMoved("������ü�¼", strNO, , "4") Then
                If Not ReturnMovedExes(strNO, 4, Me.Caption) Then Exit Sub
                mblnNOMoved = False
            End If
            If InStr(1, mstrPrivs, ";ǿ���˺�;") = 0 Then
                    '���ݲ���Ȩ�޼��,ʱ������,���ü��Һŵ���Ч����
                    If Not ReadBillInfo(1, strNO, 4, strOper, vDate) Then
                        cboNO.Text = "": cboNO.SetFocus: Exit Sub
                    End If
                    If Not BillOperCheck(1, strOper, vDate, IIf(chkCancel.Value = 1, "�˺�", "�ش�")) Then
                        cboNO.Text = "": cboNO.SetFocus: Exit Sub
                    End If
            End If
            
            '�����˺�Ȩ��
            If chkCancel.Value = 1 Then
                If mblnStation Then '����ҽ��վ�˺ż��
                    If Not StationDelete(strNO, str����NO) Then
                        cboNO.Text = "": cboNO.SetFocus: Exit Sub
                    End If
                ElseIf InStr(1, mstrPrivs, ";ǿ���˺�;") = 0 Then
                    If CheckPriceHaveFee(strNO, str����NO) Then Exit Sub
                    '���Һŵ��Ƿ���ִ��
                    blnEnableDel = (InStr(mstrPrivs, ";��ҽ�����˺�;") > 0)
                    If CheckExecuted(strNO, blnEnableDel) Then
                        MsgBox "�Һŵ�" & strNO & "�Ѿ���ҽ��������¹�ҽ��,�����˺ţ�", vbInformation, gstrSysName
                        cboNO.Text = "": cboNO.SetFocus: Exit Sub
                    End If
                    
                    '�Ƿ���������,��δ�˷�
                    If InStr(1, mstrPrivs, ";�շѺ��˺�;") = 0 Then
                        If ExistFee(strNO) Then
                            MsgBox strNO & "�Һŵ��Ĳ����Ѿ������˷���,�����˷Ѳ����˺�.", vbInformation, gstrSysName
                            cboNO.Text = "": cboNO.SetFocus: Exit Sub
                        End If
                    End If
                End If
                mintInsure = ExistInsure(strNO)
                mlng����ID = GetBill����ID(strNO, 4)
            End If
            
            If Not ReadBill(strNO) Then
                MsgBox "û�з���������ĹҺŵ��ݣ�", vbInformation, gstrSysName
                cboNO.Text = "": cboNO.SetFocus
            Else
                mstr����NO = str����NO
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

Private Sub cbo���ڵ�ַ_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Function ReadBill(strNO As String, Optional blnGetBooking As Boolean = False) As Boolean
    '���ܣ����ݵ��ݺŶ�ȡ�Һŵ��ݲ���ʾ�ڽ�����
    '����: �鿴,�˺�,����ԤԼ
    'blnGetBooking-�Ƿ���ԤԼ���� ��Ϊ������Һ�ʹ�á�/�� ��ȡԤԼ����ʱ ȱ�ٶ�����ʱ��ļ�� �������ӿ�ѡ���� ��ͨ��"/"��ȡ��ԤԼ����ʱ ����
       ' Dim rsBill As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Integer
    Dim curMoney As Currency
    Dim Datsys      As Date
    Dim datTmp      As Date
    Dim blnChk      As Boolean
    Dim bytState    As Byte, strTable As String
    Dim blnNotClick As Boolean
    Dim bln���ѿ�   As Boolean
    Dim dblTotal    As Double, dblBalance As Double
    Dim cllBillBalance As Collection
    Dim objCard As Card, rsTx As ADODB.Recordset
    Dim strWhere As String, str����IDs As String
    
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
        strWhere = IIf(mblnStation, " And A.ִ����=[3]", "") 'ҽ��վ��������ִ����
        strWhere = strWhere & IIf(mintCancel = 1 Or mintCancel = 2, "And A.�շ�ϸĿID = M.Column_Value", "") '��ȡָ����Ŀ
   
        If mbytMode = 0 And chkCancel.Value = 0 Then
            strSQL = " " & _
            "   Select  a.No, a.ʵ��Ʊ��, Nvl(a.�۸񸸺�, a.���) As ���, a.��������, a.��ʶ��,D.��������, a.����id, a.���ʽ, d.ҽ�Ƹ��ʽ, f.ҽ�����,a.����, a.�Ա�, a.����, " & _
            "            d.���֤��, d.��ͥ�绰, d.��ͥ��ַ, d.��������, d.���ڵ�ַ, a.�ѱ�,  a.�Ӱ��־, Nvl(a.���ӱ�־, 0) As ���ӱ�־, a.���㵥λ As �ű�, b.���� As ��Ŀ,a.ִ�в���id, " & _
            "           c.���� As ����, nvl(a.Ӧ�ս��,0)+nvl(J.Ӧ��,0) As Ӧ��,nvl(a.ʵ�ս��,0)+nvl(J.ʵ��,0) As ʵ��, G.�˺������, g.�˺����ʱ��, a.ִ����,a.ִ����, a.����ʱ��, a.����Ա����, a.����id, a.ժҪ, a.����, " & _
            "           Decode(g.����, Null, a.��ҩ����, To_Char(g.����)) as ����,a.�շ�ϸĿid,a.������Ŀid, a.�۸񸸺�, a.�շ����, a.����, a.��׼����, a.�վݷ�Ŀ, a.���մ���id, " & _
            "           a.������Ŀ��, a.ͳ����, a.���ձ���, a.���˿���id,Nvl(a.���ʷ���, 0)  As ���ʷ���,Nvl(G.����, 0)  As ���� " & _
            "   From ���˹Һż�¼ G,������ü�¼  A, �շ���ĿĿ¼ B, ���ű� C, ������Ϣ D, ����ǼǼ�¼ F, " & _
            "          (  Select B1.NO,A1.���,sum(A1.Ӧ�ս��) as Ӧ��,sum(A1.ʵ�ս��) as ʵ�� " & _
            "             From ������ü�¼ A1,���˹Һż�¼ B1 " & _
            "             Where b1.�շѵ�=A1.No and a1.��¼����=1 and A1.��¼״̬ in (0,1,3) and b1.NO=[2] And b1.��¼״̬(+)=Decode([1],0,1,[1]) " & _
            "             group by B1.NO,A1.���  ) J " & strTable & _
            "   Where  G.No=A.No and a.��¼����=4 And a.��¼״̬ = [1] and G.NO=[2] And g.��¼״̬(+)=Decode([1],0,1,[1]) " & strWhere & _
            "          And a.�շ�ϸĿid = b.Id And a.ִ�в���id = c.Id And a.����id = d.����id(+)  " & _
            "          And G.�Ǽ�ʱ�� = f.����ʱ��(+) And G.����id = F.����id(+)  And (C.վ��='" & gstrNodeNo & "' Or B.վ�� is Null)" & _
            "          And A.NO=J.No(+) and A.���=J.���(+)"
        Else
              strSQL = "" & _
            " Select A.NO,A.ʵ��Ʊ��,Nvl(A.�۸񸸺�,A.���) as ���,A.��������,A.��ʶ��,D.��������," & _
            "           A.����ID,A.���ʽ ,D.ҽ�Ƹ��ʽ,F.ҽ�����,A.����,A.�Ա�,A.����,D.���֤��,D.��ͥ�绰 ,D.��ͥ��ַ, D.��������,D.���ڵ�ַ,A.�ѱ�,A.�Ӱ��־," & _
            "           Nvl(A.���ӱ�־,0) as ���ӱ�־,A.���㵥λ as �ű�,B.���� as ��Ŀ,A.ִ�в���ID,C.���� as ����," & _
            "           " & IIf(bytState = 2, "-1*", "") & "Sum(Ӧ�ս��) as Ӧ��," & IIf(bytState = 2, "-1*", "") & "Sum(ʵ�ս��) as ʵ��,e.�˺������,e.�˺����ʱ��," & _
            "           A.ִ����,A.����ʱ��,A.����Ա����,A.����ID,A.ժҪ,A.����,Decode(E.����, Null, A.��ҩ����, To_Char(E.����)) ����,A.�շ�ϸĿID,A.������ĿID,  A.�۸񸸺�, A.�շ����," & _
            "           A.����, A.��׼����, A.�վݷ�Ŀ, A.���մ���id, A.������Ŀ��, A.ͳ����, A.���ձ���, A.���˿���id, " & _
            "           max(nvl(A.���ʷ���,0)) as ���ʷ���,Max(nvl(E.����,0)) as  ����" & _
            " From ������ü�¼ A,���˹Һż�¼ E,����ǼǼ�¼ F,�շ���ĿĿ¼ B,���ű� C,������Ϣ D" & strTable & _
            " Where A.NO=E.NO(+) And A.����ID=D.����ID(+) And A.��¼����=4 And A.��¼״̬=[1] And E.��¼״̬(+)=Decode([1],0,1,[1])  " & _
            "       And E.�Ǽ�ʱ��=F.����ʱ��(+) And E.����ID=F.����ID(+)" & strWhere & _
            "       And A.NO=[2] And A.�շ�ϸĿID=B.ID And A.ִ�в���ID=C.ID" & _
            "       And (C.վ��='" & gstrNodeNo & "' Or B.վ�� is Null)" & IIf(mbytMode = 0 And chkCancel.Value = 0, " And e.�շѵ� Is Null ", "") & _
            " Group by A.NO,A.ʵ��Ʊ��,Nvl(A.�۸񸸺�,A.���),A.��������,A.��ʶ��,D.��������,A.����ID,A.���ʽ,D.ҽ�Ƹ��ʽ,F.ҽ�����,A.����,A.�Ա�,D.���֤��,D.��ͥ�绰," & _
            "           A.����,D.��ͥ��ַ,D.���ڵ�ַ,A.�ѱ�,A.�Ӱ��־,A.���ӱ�־,A.���㵥λ,B.����,C.����,A.ִ�в���ID,A.ִ����,A.����ʱ��,A.����Ա����,A.����ID,A.ժҪ,A.����,Decode(E.����, Null, A.��ҩ����, To_Char(E.����)),E.�˺������, E.�˺����ʱ��,A.�շ�ϸĿID,A.������ĿID, A.�۸񸸺�, A.�շ����," & _
            "           A.����, A.��׼����, A.�վݷ�Ŀ, A.���մ���id, A.������Ŀ��, A.ͳ����, A.���ձ���, A.���˿���id, D.��������" & _
            " "
        End If
        strSQL = strSQL & " Order By ��� "
        If mblnNOMoved Then
            strSQL = Replace(strSQL, "����Һż�¼", "H����Һż�¼")
            strSQL = Replace(strSQL, "������ü�¼", "H������ü�¼")
        End If
        Set mrsBill = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, bytState, strNO, UserInfo.����, mstr������ĿID, mstr�˷���ĿIDs)
   Else
        strSQL = "" & _
        "   Select a.No, Null As ʵ��Ʊ��, 0 As ���, Null As ��������, a.����� as ��ʶ��, a.����id, Null As ���ʽ, Null ҽ�Ƹ��ʽ, f.ҽ�����, a.����, a.�Ա�, a.����," & _
        "          d.���֤��, d.��ͥ�绰, d.��ͥ��ַ, d.�ѱ�, a.���� As �Ӱ��־, Nvl(A.���ӱ�־,0) as ���ӱ�־, a.�ű�, b.���� As ��Ŀ, a.ִ�в���id, c.���� As ����, 0  As Ӧ��, 0 As ʵ��, a.ִ����," & _
        "          a.����ʱ��, a.����Ա����, Null As ����ID, a.ժҪ, a.ԤԼ��ʽ As ����, a.����,a.�˺������,a.�˺����ʱ��, 0 as �շ�ϸĿID,0 as ������ĿID,D.��������," & _
        "          0 as ���ʷ���,Nvl(A.����,0) as  ����,D.��������,D.���ڵ�ַ" & _
        "   From ���˹Һż�¼ A, �շ���ĿĿ¼ B,�ҺŰ��� E, ���ű� C, ������Ϣ D, ����ǼǼ�¼ F  " & _
        "   Where E.��Ŀid = b.Id And a.�ű�=e.���� And a.ִ�в���id = c.Id And a.��¼���� = 2 And a.��¼״̬ = [1] And a.����id = d.����id(+) And " & _
        "       A.No=[2] and  a.�Ǽ�ʱ�� = f.����ʱ��(+) And a.����ID=f.����ID(+)  " & _
        "       And (c.վ�� ='" & gstrNodeNo & "' Or b.վ�� Is Null)" & IIf(mblnStation, " And A.ִ����=[3]", "") & vbNewLine & _
        "   Union All " & vbNewLine & _
        "   Select a.No, Null As ʵ��Ʊ��, 0 As ���, Null As ��������, a.����� as ��ʶ��, a.����id, Null As ���ʽ, Null ҽ�Ƹ��ʽ, f.ҽ�����, a.����, a.�Ա�, a.����," & _
        "          d.���֤��, d.��ͥ�绰, d.��ͥ��ַ, d.�ѱ�, a.���� As �Ӱ��־, Nvl(A.���ӱ�־,0) as ���ӱ�־, a.�ű�, b.���� As ��Ŀ, a.ִ�в���id, c.���� As ����, 0  As Ӧ��, 0 As ʵ��, a.ִ����," & _
        "          a.����ʱ��, a.����Ա����, Null As ����ID, a.ժҪ, a.ԤԼ��ʽ As ����, a.����,a.�˺������,a.�˺����ʱ��, 0 as �շ�ϸĿID,0 as ������ĿID,D.��������," & _
        "          0 as ���ʷ���,Nvl(A.����,0) as  ����,D.��������,D.���ڵ�ַ" & _
        "   From ���˹Һż�¼ A, �շ���ĿĿ¼ B,�ҺŰ��� E, ���ű� C, ������Ϣ D, ����ǼǼ�¼ F ,�շѴ�����Ŀ G " & _
        " Where E.��Ŀid = G.����Id And a.�ű�=e.���� And a.ִ�в���id = c.Id And a.��¼���� = 2 And a.��¼״̬ = [1] And a.����id = d.����id(+) And " & _
        "        G.����ID=b.Id And A.No=[2] and  a.�Ǽ�ʱ�� = f.����ʱ��(+) And a.����ID=f.����ID(+)  " & _
        "        And (c.վ�� ='" & gstrNodeNo & "' Or b.վ�� Is Null)" & IIf(mblnStation, " And A.ִ����=[3]", "")
        
        Set mrsBill = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(mbytInState), strNO, UserInfo.����)
    End If
    
    If mrsBill.EOF Then
        If mbytMode = 4 And mbytInState = 1 Then
            MsgBox "û���ҵ����ݺ�Ϊ[" & mstrNoIn & "]�ĵ���!", vbOKOnly, Me.Caption
        End If
        Exit Function
    End If
    mlng����ID = Val(Nvl(mrsBill!����ID))
    '------------------------------------
     ' �Խ��� ����ȡ��ԤԼ �ļ��
     '------------------------------------
    Select Case mbytMode
    Case 2:
     '--����
        If mbytMode = 2 And mTy_Para.lngԤԼ��Чʱ�� <> 0 Then
chkBooking:
            blnChk = True
            Datsys = DateAdd("n", 1 * mTy_Para.lngԤԼ��Чʱ��, zlDatabase.Currentdate)
            If Format(Datsys, "yyyy-MM-dd hh:mm:ss") > Format(mrsBill!����ʱ��, "yyyy-MM-dd hh:mm:ss") Then
                datTmp = DateAdd("n", -1 * mTy_Para.lngԤԼ��Чʱ��, CDate(Format(mrsBill!����ʱ��, "yyyy-MM-dd hh:mm:ss")))
                MsgBox "��ԤԼ���ѹ�ԤԼ������ʱ�� " & Format(datTmp, "yyyy-MM-dd hh:mm:00") & ",���ܽ���", vbInformation, Me.Caption
                mblnUnload = True
                Exit Function
            End If
        End If
    Case 3:
         '--ȡ��ԤԼ
         '----------------------
         'ȡ��ԤԼ
         '���Ʋ���:1. N���ڲ���ȡ��ԤԼ��
         '        2.�˺����
         '   ����1.������ȡ��ԤԼ������ԤԼʱ���N����
         '   ���ȡ��ԤԼ��N����
         '    <1> �˺����Ϊ�� ʱ ��˵�ԤԼ�� �ܹ�ȡ�� ������
         '    <2> �˺����Ϊ�� ʱ ����ȡ��ԤԼ
         '----------------------
         If mTy_Para.lngN��ȡ��ԤԼ > 0 Then
            Datsys = zlDatabase.Currentdate
            datTmp = DateAdd("d", -1 * mTy_Para.lngN��ȡ��ԤԼ, CDate(Format(mrsBill!����ʱ��, "yyyy-MM-dd hh:mm:ss")))
            'ԤԼʱ��-K >datSys
            If Format(Datsys, "yyyy-MM-dd hh:mm:ss") > Format(datTmp, "yyyy-MM-dd hh:mm:ss") Then
                Select Case mTy_Para.bln�˺����
                Case False:
                ' �ϸ���Ʋ���ȡ��ԤԼ
                 MBox "��ԤԼ���Ѿ��������ȡ��ԤԼʱ��" & Format(datTmp, "yyyy-MM-dd hh:mm:ss") & ",����ȡ��ԤԼ!"
                 mblnUnload = True
                 Exit Function
                Case True:
                  If Nvl(mrsBill!�˺������, "") = "" Then
                    MBox "�õ��ݺ�Ϊ" & Nvl(mrsBill!NO) & "��ԤԼ��û�о����˺����!����ȡ��ԤԼ!"
                    mblnUnload = True
                    Exit Function
                  End If
                End Select
            End If
         End If
    Case Else:
    End Select
    
    If mbytMode = 0 And mbytInState = 0 And mstrNoIn <> "" Then
        '102230,������Ҳ����ӿ�
        If PatiValiedCheckByPlugIn(mlngModul, Val(Nvl(mrsBill!����ID)), _
            "<YSXM>" & NeedName(cboҽ��.Text) & "</YSXM>") = False Then Exit Function
    End If
    
    If blnGetBooking And mbytMode <> 2 And mTy_Para.lngԤԼ��Чʱ�� <> 0 And blnChk = False Then GoTo chkBooking
    Call RemoveShowItem
    Call ClearMoney
    cboNO.Text = mrsBill!NO
    cboNO.Tag = mrsBill!NO
    txtFact.Text = Nvl(mrsBill!ʵ��Ʊ��)
    cbo��ע.Text = Nvl(mrsBill!ժҪ)
    
    mbln���������� = False
    mbln���ӷ� = False
    mbln������ = False
    If mrsBill.RecordCount = 1 And Nvl(mrsBill!���ӱ�־, 0) = 1 Then
        '������ȡ������
        mblnUnChange = True
        txt�ű�.Text = "+"
        txtSN.Text = ""
        mblnUnChange = False
        chk������.Enabled = False
        mbln���������� = True
        If mintCancel = 0 And chkCancel.Value = 1 Then
            chk������.Value = 1
        End If
    Else
        '�����Һ�,����������
        vsfMoney.Tag = ""
        mrsBill.MoveFirst
        For i = 1 To mrsBill.RecordCount
            If Nvl(mrsBill!��������, 0) = 0 And Nvl(mrsBill!���ӱ�־, 0) = 0 Then
                'ֻ������һ��
                mblnUnChange = True
                txt�ű�.Text = Nvl(mrsBill!�ű�)
                If Not IsNull(mrsBill!����) Then txtSN.Text = IIf(IsNumeric(mrsBill!����), mrsBill!����, "")
                txtSN.Tag = txtSN.Text
                mblnUnChange = False
                If InStr("," & mstr������ĿID & ",", "," & Nvl(mrsBill!�շ�ϸĿID) & ",") > 0 Then
                    mbln���ӷ� = True
                Else
                    mbln������ = True
                End If
                
                txt����.Text = Nvl(mrsBill!����)
                If (mbytMode = 2 Or mbytMode = 0 And mbytInState = 0 And mstrNoIn <> "") And mlng�Һſ���ID = 0 Then mlng�Һſ���ID = Nvl(mrsBill!ִ�в���id)
                cboҽ��.Clear
                If Not IsNull(mrsBill!ִ����) Then
                    cboҽ��.AddItem mrsBill!ִ����
                    cboҽ��.ListIndex = 0
                End If
           
                lbl��.Visible = Nvl(mrsBill!�Ӱ��־, 0) = 1
            ElseIf Nvl(mrsBill!���ӱ�־, 0) = 1 Then
                blnNotClick = mblnNotClick
                mblnNotClick = True
                'ֻ������һ��
                chk������.Value = 1
                mbln���������� = True
                mblnNotClick = blnNotClick
                
            ElseIf Nvl(mrsBill!���ӱ�־, 0) = 2 Then
                '��־�������￨��
                vsfMoney.Tag = "����"
            End If
            mrsBill.MoveNext
         Next
        mrsBill.MoveFirst
    End If
    Call AdjustInfoPosition
    If chkPrint.Value <> 1 Then
        If mbln���������� = True Then
            chk������.Enabled = mintCancel = 0
        End If
        If mbln���ӷ� = True Then
            mblnNotClick = True
            chkExtra.Value = 1
            mblnNotClick = False
            chkExtra.Enabled = mintCancel = 0
            chkExtra.Visible = mintCancel = 0
            chkExtra.Top = chk������.Top
            lblԤԼ��ʽ.Visible = Not mbln���ӷ�
            cboԤԼ��ʽ.Visible = Not mbln���ӷ�
        Else
            chkExtra.Visible = False
        End If
    End If
    If mbln���������� Then chk������.Enabled = True
    
    mrsBill.MoveFirst
    Do While Not mrsBill.EOF
        dblTotal = dblTotal + Val(Nvl(mrsBill!ʵ��))
        mrsBill.MoveNext
    Loop
    mrsBill.MoveFirst
    
    If (mbytMode = 2 Or mbytMode = 0 And mbytInState = 0 And mstrNoIn <> "") And Not IsNull(mrsBill!����ID) Then
        mblnNotEMPIQuery = True
        Call GetPatient(IDKind.GetCurCard, "-" & mrsBill!����ID, False)
    End If
    If mrsBill.RecordCount <> 0 And mrsBill.EOF Then mrsBill.MoveFirst
    txtPatient.Text = Nvl(mrsBill!����)
    '74428�����ϴ���2014-7-8������������ʾ��ɫ����
    Call SetPatiColor(txtPatient, Nvl(mrsBill!��������), IIf(Val(mrsBill!����) = 0, txtPatient.ForeColor, vbRed))
    If txtPatientPrint.Visible Then
        txtPatientPrint.Text = txtPatient.Text
        txtPatientPrint.Tag = Val(Nvl(mrsBill!����ID))
        txtPatientPrint.ForeColor = txtPatient.ForeColor
        If Val(Nvl(mrsBill!����ID)) <> 0 Then
            '����ǽ�������,�����¹����������:
            '  1.ֻ�йҺ�ʱ�������Ҳ������޸�
            If Not CheckCanModifyName(cboNO.Text) And zlExistOperationData(Val(Nvl(mrsBill!����ID)), cboNO.Text) Then
                txtPatientPrint.Locked = True
                Call SetRePrintPatiEnabled(False)
            Else
                txtPatientPrint.Locked = False
                Call SetRePrintPatiEnabled(True)
            End If
        End If
        '����:53037
        ReInitPatiInvoice True
    End If
    
    If (mbytMode = 2 Or mbytMode = 0 And mbytInState = 0 And mstrNoIn <> "") Then mstrPrePati = txtPatient.Text
    
    
    Call LoadOldData("" & mrsBill!����, txt����, cbo���䵥λ)
    mstr���� = txt����.Text
    mstr���䵥λ = IIf(IsNumeric(txt����.Text), cbo���䵥λ.Text, "")
    cbo��ͥ��ַ.Text = Nvl(mrsBill!��ͥ��ַ)
    cbo���ڵ�ַ.Text = Nvl(mrsBill!���ڵ�ַ)
    '89242:���ϴ�,2015/12/7,��ȡ���˵�ַ��Ϣ
    Call zlReadAddrInfo(padd��ͥ��ַ, Val(Nvl(mrsBill!����ID)), 0, 3, cbo��ͥ��ַ.Text)
    Call zlReadAddrInfo(padd���ڵ�ַ, Val(Nvl(mrsBill!����ID)), 0, 4, cbo���ڵ�ַ.Text)
    txtIDCard.Text = Nvl(mrsBill!���֤��): txt��ͥ�绰.Text = Nvl(mrsBill!��ͥ�绰)
    mblnNotChange = True
    cbo�Ա�.ListIndex = cbo.FindIndex(cbo�Ա�, Nvl(mrsBill!�Ա�), True)
    If cbo�Ա�.ListIndex = -1 Then
        cbo�Ա�.AddItem Nvl(mrsBill!�Ա�), 0
        cbo�Ա�.ListIndex = cbo�Ա�.NewIndex
    End If
    mblnNotChange = False
    mstr�Ա� = NeedName(cbo�Ա�.Text)
    mstr���� = txtPatient.Text
    If mrsBill.RecordCount <> 0 And mrsBill.EOF Then mrsBill.MoveFirst
    txt�����.Text = Nvl(mrsBill!��ʶ��)
    mRegistFeeMode = IIf(Val(Nvl(mrsBill!���ʷ���)) = 1, EM_RG_����, EM_RG_����)
    '103974:���ϴ�,2016/12/16���鿴�����ա��˺ŵȲ���ʱ������������
    'Ҳ�������������
    mblnChange = False
    txt��������.Text = Format(IIf(IsNull(mrsBill!��������), "____-__-__", mrsBill!��������), "YYYY-MM-DD")
    If Not IsNull(mrsBill!��������) Then
        txt����ʱ��.Text = Format(mrsBill!��������, "HH:MM")
    Else
        txt����ʱ��.Text = "__:__"
    End If
    mblnChange = True
    
    '90875:���ϴ�,2016/11/8,ҽ�ƿ�֤������
    If txtIDCard.Text = "" Then
        strSQL = "Select B.����,A.���� from ����ҽ�ƿ���Ϣ A,ҽ�ƿ���� B,֤������ C " & _
                "Where A.�����ID=B.ID And B.����=C.���� And A.����ID=[1]  Order by C.���� "
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "ȱʡ��֤������", Val(Nvl(mrsBill!����ID)))
        If Not rsTmp.EOF Then
            IDKind֤��.IDKind = IDKind֤��.GetKindIndex(Nvl(rsTmp!����))
            txt֤��.Text = Nvl(rsTmp!����): txt֤��.Tag = txt֤��.Text
        End If
    End If
    
    'ҽ�Ƹ��ʽ
    If Not IsNull(mrsBill!ҽ�Ƹ��ʽ) Then
        cbo���ʽ.ListIndex = cbo.FindIndex(cbo���ʽ, mrsBill!ҽ�Ƹ��ʽ, True)
        If cbo���ʽ.ListIndex = -1 Then
            cbo���ʽ.AddItem mrsBill!ҽ�Ƹ��ʽ, 0
            cbo���ʽ.ListIndex = cbo���ʽ.NewIndex
        End If
    ElseIf Not IsNull(mrsBill!���ʽ) Then
        cbo���ʽ.AddItem Getҽ�Ƹ��ʽ(Val(mrsBill!���ʽ)), 0
        cbo���ʽ.ListIndex = cbo���ʽ.NewIndex
    Else
        cbo���ʽ.ListIndex = -1
    End If
    
    cbo�ѱ�.ListIndex = cbo.FindIndex(cbo�ѱ�, Nvl(mrsBill!�ѱ�), True)
    If cbo�ѱ�.ListIndex = -1 Then
        cbo�ѱ�.AddItem Nvl(mrsBill!�ѱ�), 0
        cbo�ѱ�.ListIndex = cbo�ѱ�.NewIndex
    End If
    
    If mlngOutModeMC > 0 Then
        cboҽ�����.ListIndex = cbo.FindIndex(cboҽ�����, "" & mrsBill!ҽ�����, True)
        If cboҽ�����.ListIndex = -1 And Not IsNull(mrsBill!ҽ�����) Then
            cboҽ�����.AddItem "" & mrsBill!ҽ�����, 0
            cboҽ�����.ListIndex = cboҽ�����.NewIndex
        Else
            cboҽ�����.ListIndex = 0
        End If
    End If
    Set mobjDelCards = New Cards
    '134708:���ϴ�,2018/12/14,���һ��ͨ����
    Set mobjPayCard = Nothing
    Dim bln�˺Ŵ��� As Boolean
    
    If mbytInState = 1 And mbytMode = 4 Or chkCancel.Value = 1 Then
        bln�˺Ŵ��� = True
        '�˺�ʱ,��ȡ����ʱ��Ӧ����Ϣ
         If Not zlReadRegThreeBalance(strNO, cllBillBalance, mobjPayCard) Then
         '�����ID,����,�Ƿ����ѿ�(1-��;0-��),������ˮ��,����˵��,strNO,����ID
             SetDelBillCtlEnabled (False)
         Else
            If Not cllBillBalance Is Nothing Then
                bln���ѿ� = Val(cllBillBalance(1)(2)) = 1
                Call SetDelBillCtlEnabled(True)
            End If
         End If
    End If
    '���Ĳ��˹Һ���Ϣʱ,���㷽ʽҲ����Ϊҽ�ƿ�����
    If mbytInState = 1 And mbytMode = 0 Then
        Call zlReadRegThreeBalance(strNO, cllBillBalance, mobjPayCard)
    End If
    '68991
    If Val(Nvl(mrsBill!���ʷ���)) <> 0 Then
        '�Ƿ�ҽ��ˢ��
        mRegistFeeMode = EM_RG_����
        If mintInsure = 0 Then mintInsure = Val(Nvl(mrsBill!����))
        Call SetUndisplayBalance
    Else
        mRegistFeeMode = EM_RG_����
        If mintInsure = 0 Then mintInsure = ExistInsure(strNO)
    End If
    
    If mintInsure <> 0 Then Call initInsurePara(mrsBill!����ID)
    
    If chkCancel.Value = 1 Or (mbytInState = 1 And mbytMode = 4) Then
        strSQL = "Select ����ID From ������ü�¼ where NO = [1] and ��¼���� = 4 "
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
        Do While Not rsTmp.EOF
            If InStr("," & str����IDs & ",", "," & Val(Nvl(rsTmp!����ID)) & ",") = 0 Then
                str����IDs = str����IDs & "," & Val(Nvl(rsTmp!����ID))
            End If
            rsTmp.MoveNext
        Loop
        If str����IDs <> "" Then str����IDs = Mid(str����IDs, 2)
    Else
        str����IDs = mlng����ID
    End If
    
'    txtԤ��֧��.Tag = ""
    '���㷽ʽ:���ܰ���ҽ��֧������
    strSQL = "Select Mod(A.��¼����,10) as ��¼����,B.����,A.���㷽ʽ," & _
        IIf(bytState = 2, "-1*", "") & "Sum(A.��Ԥ��) as ���, A.������� ,Nvl(Nvl(C.����, D.����), A.���㷽ʽ) As ���� " & _
        " From " & IIf(mblnNOMoved, "H", "") & "����Ԥ����¼ A,���㷽ʽ B,ҽ�ƿ���� C,���ѿ����Ŀ¼ D" & _
        " Where A.���㷽ʽ=B.����(+) And A.�����ID=C.ID(+) And a.���㿨��� = D.���(+) " & _
        "   And a.����id in (Select /* +cardinality(M,10) */ M.Column_Value From Table(f_Str2list([1])) M)" & _
        " Group by Mod(A.��¼����,10),B.����,A.���㷽ʽ,A.�������,C.����,D.����" & _
        " Having Sum(A.��Ԥ��) <> 0" & _
        " Order by Mod(A.��¼����,10),B.����,A.���㷽ʽ"
    Set mrsBillAdvance = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str����IDs)
'    vsfPay.Clear 1
'    vsfPay.Rows = 2
'    vsfPay.RowHidden(1) = False
    Call Load������Ϣ(dblTotal, 0)
    
    If mrsBillAdvance.RecordCount <> 0 Then mrsBillAdvance.MoveFirst
    For i = 1 To mrsBillAdvance.RecordCount
        If mrsBillAdvance!��¼���� = 1 Or mrsBillAdvance!��¼���� = 11 Then
        Else
            Select Case Val(Nvl(mrsBillAdvance!����))
            Case 3 'ҽ�������˻�
                '74428�����ϴ���2014-7-8������������ʾ��ɫ����
                Call SetPatiColor(txtPatient, Nvl(mrsBill!��������), vbRed)
            Case 7, 8    'һ��ͨ���
                If mobjPayCard Is Nothing Then
                    If bln�˺Ŵ��� Then
                        Set objCard = New Card
                        With objCard
                            .�ӿ���� = 0
                            .���� = Nvl(mrsBillAdvance!���㷽ʽ)
                            .���㷽ʽ = Nvl(mrsBillAdvance!���㷽ʽ)
                            .�ӿڱ��� = Val(Nvl(mrsBillAdvance!����))   ' ��¼����
                            .���� = False
                        End With
                        mobjDelCards.Add objCard
                        cbo���㷽ʽ.ListIndex = -1
                    Else
                        cbo���㷽ʽ.ListIndex = cbo.FindIndex(cbo���㷽ʽ, mrsBillAdvance!���㷽ʽ, True)
                    End If
                    If cbo���㷽ʽ.ListIndex = -1 Then
                        cbo���㷽ʽ.AddItem mrsBillAdvance!���㷽ʽ, 0
                        cbo���㷽ʽ.ListIndex = cbo���㷽ʽ.NewIndex
                    End If
                    txt����Ӧ��.Text = Format(mrsBillAdvance!���, "0.00")
                Else
                  cbo���㷽ʽ.Clear
                   If mobjPayCard.�Ƿ����� Then
                        '֧�����֣���Ҫ��������ֽ�ͷ�ҽ����Ľ��㷽ʽ
                        Call Init���㷽ʽ("1,2", mobjDelCards)
                   End If
                   cbo���㷽ʽ.AddItem IIf(Nvl(mobjPayCard.����) = "", mrsBillAdvance!���㷽ʽ, Nvl(mobjPayCard.����))
                   mobjDelCards.Add mobjPayCard
                   Set mCurCardPay.objCard = mobjPayCard
                   mCurCardPay.lngҽ�ƿ����ID = mobjPayCard.�ӿ����
                   If (mobjPayCard.���� Or cbo���㷽ʽ.ListIndex < 0 Or mobjPayCard.�Ƿ����� = False) Then
                        cbo���㷽ʽ.ListIndex = cbo���㷽ʽ.NewIndex
                    End If
                End If
            Case Else '1,2������
                If mobjPayCard Is Nothing Then
                    If bln�˺Ŵ��� Then
                        Set objCard = New Card
                        With objCard
                            .�ӿ���� = 0
                            .���� = Nvl(mrsBillAdvance!���㷽ʽ)
                            .���㷽ʽ = Nvl(mrsBillAdvance!���㷽ʽ)
                            .�ӿڱ��� = Val(Nvl(mrsBillAdvance!����))   ' ��¼����
                            .���� = False
                        End With
                        mobjDelCards.Add objCard
                        cbo���㷽ʽ.ListIndex = -1
                    Else
                        cbo���㷽ʽ.ListIndex = cbo.FindIndex(cbo���㷽ʽ, mrsBillAdvance!���㷽ʽ, True)
                    End If
                    If cbo���㷽ʽ.ListIndex = -1 Then
                        cbo���㷽ʽ.AddItem mrsBillAdvance!���㷽ʽ, 0
                        cbo���㷽ʽ.ListIndex = cbo���㷽ʽ.NewIndex
                    End If
                Else
                  cbo���㷽ʽ.Clear
                   If mobjPayCard.�Ƿ����� Then
                        '֧�����֣���Ҫ��������ֽ�ͷ�ҽ����Ľ��㷽ʽ
                        Call Init���㷽ʽ("1,2", mobjDelCards)
                   End If
                   mobjDelCards.Add mobjPayCard
                    cbo���㷽ʽ.AddItem IIf(Nvl(mobjPayCard.���㷽ʽ) = "", mrsBillAdvance!���㷽ʽ, Nvl(mobjPayCard.���㷽ʽ))
                    If (mobjPayCard.���� Or cbo���㷽ʽ.ListIndex < 0 Or mobjPayCard.�Ƿ����� = False) Then
                        cbo���㷽ʽ.ListIndex = cbo���㷽ʽ.NewIndex
                    End If
                End If
                txt����Ӧ��.Text = Format(mrsBillAdvance!���, "0.00")
            End Select
        End If
        mrsBillAdvance.MoveNext
    Next
    
    If bln�˺Ŵ��� And Not mobjPayCard Is Nothing Then
        '�˺�:��������,������Ľ��㷽ʽ
        If mobjPayCard.�Ƿ����� Then cbo���㷽ʽ.Enabled = True
    End If
    
    
    txt����ʱ��.Text = Format(mrsBill!����ʱ��, "yyyy-MM-dd HH:mm:ss")
    cbo��ע.Text = Nvl(mrsBill!ժҪ)
    lbl����.Visible = False
    mblnNotChange = True
    zlControl.CboSetText cbo��ע, Nvl(mrsBill!ժҪ)
    mblnNotChange = False
    mstrԭժҪ = Nvl(mrsBill!ժҪ)
    '����:26955
    zlAddComboItem cboԤԼ��ʽ, Nvl(mrsBill!����)
        
    mrsBill.MoveFirst
    vsfMoney.Rows = mrsBill.RecordCount + 1
    For i = 1 To mrsBill.RecordCount
        vsfMoney.TextMatrix(i, 0) = mrsBill!��Ŀ
        vsfMoney.TextMatrix(i, 1) = Format(mrsBill!Ӧ��, "0.00")
        vsfMoney.TextMatrix(i, 2) = Format(mrsBill!ʵ��, "0.00")
        curMoney = curMoney + mrsBill!ʵ��
        mrsBill.MoveNext
    Next
    mrsBill.MoveFirst
    txt�ϼ�.Text = Format(curMoney, "0.00")
    lbl�ϼ�.Caption = Format(curMoney, "0.00")
    Call Set�����Һ�
    If txt�����.Text = "" And mbytMode = 2 And gbln�Զ������ Then
        txt�����.Text = zlGet�����
    End If
    mbln������ = zlIsCreatePatiArchives(txt�ű�.Text)   '36131
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

Private Function Load������Ϣ(ByVal dblTotal As Double, Optional ByVal dblDiff As Double) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����֧����Ϣ�����˷�ʱ���˷���Ϣ
    '���:dblTotal-�����˷��ܽ�dblDiff-ȡ�������ѻ򸽼ӷѵĲ��
    '����:�ɹ�����true,���򷵻�False
    '����:���ϴ�
    '����:2018/5/2 11:35:08
    '����:123874
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim dblBalance As Double
    Dim strSQL As String, rsTx As ADODB.Recordset
    On Error GoTo errH
    Call InitVsfPay(mbln�����Һ�)
    If mrsBillAdvance Is Nothing Then Exit Function
    If mrsBillAdvance.RecordCount > 0 Then mrsBillAdvance.MoveFirst
    For i = 1 To mrsBillAdvance.RecordCount
        If dblTotal <> 0 Then
            If dblTotal < Val(mrsBillAdvance!���) Then
                dblBalance = dblTotal
                dblTotal = 0
            Else
                If FormatEx(Val(mrsBillAdvance!���), 6) >= FormatEx(dblDiff, 6) And dblDiff <> 0 And Val(Nvl(mrsBillAdvance!��¼����)) <> 5 Then
                    dblBalance = Val(mrsBillAdvance!���) - dblDiff: dblDiff = 0
                Else
                    dblBalance = Val(mrsBillAdvance!���)
                End If
                dblTotal = dblTotal - dblBalance
            End If
            If mrsBillAdvance!��¼���� = 1 Or mrsBillAdvance!��¼���� = 11 Then
                With vsfPay
                    .TextMatrix(.Rows - 1, 0) = "��Ԥ��"
                    .TextMatrix(.Rows - 1, 1) = Format(dblBalance, "0.00")
                    .TextMatrix(.Rows - 1, .ColIndex("�޸�")) = "1"
                    .TextMatrix(.Rows - 1, 2) = Nvl(mrsBillAdvance!�������)
                    .RowData(.Rows - 1) = 0
                End With
            Else
                With vsfPay
                    '�����:116146,����,2017/11/24,�˺�ʱ,���㷽ʽ��ʾ����ҽ�ƿ��Ľ��㷽ʽ��ͳһ����Ϊҽ�ƿ�����
                    If mobjPayCard Is Nothing Then
                        .TextMatrix(.Rows - 1, 0) = mrsBillAdvance!���㷽ʽ
                    Else
                        .TextMatrix(.Rows - 1, 0) = IIf(Nvl(mobjPayCard.����) <> "" And (Val(Nvl(mrsBillAdvance!����, -1)) = 7 Or Val(Nvl(mrsBillAdvance!����, -1)) = 8), Nvl(mobjPayCard.����), mrsBillAdvance!���㷽ʽ)
                    End If
                    .TextMatrix(.Rows - 1, 1) = Format(dblBalance, "0.00")
                    .TextMatrix(.Rows - 1, 2) = Nvl(mrsBillAdvance!�������)
                    .RowData(.Rows - 1) = Val(Nvl(mrsBillAdvance!����, -1))
                    If Val(Nvl(mrsBillAdvance!����, -1)) = 7 Or Val(Nvl(mrsBillAdvance!����, -1)) = 8 Then
                        strSQL = "Select ID,�Ƿ����� From ҽ�ƿ���� Where ���㷽ʽ=[1]"
                        Set rsTx = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mrsBillAdvance!���㷽ʽ)
                        If rsTx.EOF Then
                            strSQL = "Select ���,�Ƿ����� From ���ѿ����Ŀ¼ Where ���㷽ʽ=[1]"
                            Set rsTx = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mrsBillAdvance!���㷽ʽ)
                            If rsTx.EOF Then
                                vsfPay.TextMatrix(.Rows - 1, .ColIndex("�޸�")) = "1"
                            Else
                                vsfPay.TextMatrix(.Rows - 1, 4) = Nvl(rsTx!���)
                                vsfPay.TextMatrix(.Rows - 1, .ColIndex("�޸�")) = IIf(Val(rsTx!�Ƿ�����) = 1, "0", "1")
                            End If
                        Else
                            vsfPay.TextMatrix(.Rows - 1, 4) = Nvl(rsTx!ID)
                            vsfPay.TextMatrix(.Rows - 1, .ColIndex("�޸�")) = IIf(Val(rsTx!�Ƿ�����) = 1, "0", "1")
                        End If
                    End If
                    If Val(Nvl(mrsBillAdvance!����, -1)) = 1 Or Val(Nvl(mrsBillAdvance!����, -1)) = 3 Then
                        .TextMatrix(.Rows - 1, .ColIndex("�޸�")) = "1"
                    End If
                End With
            End If
        End If
        mrsBillAdvance.MoveNext
        vsfPay.Rows = vsfPay.Rows + 1
    Next i
    
    Load������Ϣ = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function zlIsCreatePatiArchives(ByVal str���� As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��鵱ǰ�ű��Ƿ񽨵�
    '���:str����-���ź���
    '����:�轨��,����true,���򷵻�False
    '����:���˺�
    '����:2011-03-03 11:15:42
    '����:36131
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    On Error GoTo errHandle
    strSQL = " Select max(��������) as ���� From �ҺŰ��� where ����=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str����)
    zlIsCreatePatiArchives = Val(Nvl(rsTemp!����)) = 1
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function CheckCanModifyName(ByVal strNO As String) As Boolean
'����:���Һŵ��Ƿ�����޸�����,������ǹҺ�ʱ���ĵ�,�Ͳ����޸�.
    Dim rsTmp As ADODB.Recordset, strSQL As String
 
    strSQL = "Select 1" & vbNewLine & _
            "From ������ü�¼ A, ������Ϣ B" & vbNewLine & _
            "Where A.NO = [1] And A.��¼���� = 4 And A.�Ǽ�ʱ�� = B.�Ǽ�ʱ�� And A.����id = B.����id"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
    CheckCanModifyName = rsTmp.RecordCount > 0
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Sub RemoveShowItem()
    '�Ա�
    If cbo�Ա�.ListCount > 0 Then
        If Not cbo�Ա�.List(0) Like "*-*" Then
            cbo�Ա�.RemoveItem 0
            SetCboDefault cbo�Ա�
        End If
    End If
    '���ʽ
    If cbo���ʽ.ListCount > 0 Then
        If Not cbo���ʽ.List(0) Like "*-*" Then
            cbo���ʽ.RemoveItem 0
            SetCboDefault cbo���ʽ
        End If
    End If
    '�ѱ�
    If cbo�ѱ�.ListCount > 0 Then
        If Not cbo�ѱ�.List(0) Like "*-*" Then
            cbo�ѱ�.RemoveItem 0
            SetCboDefault cbo�ѱ�
        End If
    End If
    
    '���㷽ʽ
    If cbo���㷽ʽ.ListCount > 0 Then
        If Not cbo���㷽ʽ.List(0) Like "*-*" Then
            cbo���㷽ʽ.RemoveItem 0
            SetCboDefault cbo���㷽ʽ
        End If
    End If
End Sub
Private Function GetCol(strName As String) As Long
   GetCol = vsfPlan.ColIndex(strName)
End Function

Private Sub SetPatiInfoEnabled(Optional ByVal blnUse As Boolean, Optional ByVal blnNewPati As Boolean, Optional ByVal blnReservePati As Boolean)
'���ܣ����ò�������ʹ��״̬
    Dim blnEnabled As Boolean, lng����ID As Long
    '82859:���ϴ�,2015/4/8,���˻�����Ϣ����
    If Not blnNewPati Then
        If mrsInfo.RecordCount > 0 Then lng����ID = Val(Nvl(mrsInfo!����ID))
    End If
    mbln������Ϣ���� = Not (lng����ID <> 0 And InStr(1, ";" & GetPrivFunc(glngSys, 9003) & ";", ";������Ϣ����;") = 0)
    txtPatient.Enabled = gbln���� Or blnUse
    If mblnStation Then
        blnEnabled = (gbln���ʽ Or blnUse) And blnNewPati
        cbo�Ա�.Enabled = blnEnabled And mbln������Ϣ���� '�����:58843
        txt����.Enabled = blnEnabled And mbln������Ϣ���� And Not mTy_Para.bln��ֹ�������� '�����:58843
        cbo���䵥λ.Enabled = blnEnabled And mbln������Ϣ���� And Not mTy_Para.bln��ֹ�������� '�����:58843
        cbo��ͥ��ַ.Enabled = gbln��ͥ��ַ Or blnUse '�����:58843
        cbo���ڵ�ַ.Enabled = blnUse
        '89242:���ϴ�,2015/12/7,��ȡ���˵�ַ��Ϣ
        padd��ͥ��ַ.Enabled = gbln��ͥ��ַ Or blnUse: padd��ͥ��ַ.ControlLock = Not (gbln��ͥ��ַ Or blnUse)
        padd���ڵ�ַ.Enabled = blnUse: padd���ڵ�ַ.ControlLock = Not blnUse
        cbo���ʽ.Enabled = blnEnabled '�����:58843
        txt��ͥ�绰.Enabled = blnEnabled
    Else
        '���˺�:66032(��������������58843)
        cbo�Ա�.Enabled = mbln������Ϣ���� And (gbln�Ա� Or blnUse)
        txt����.Enabled = mbln������Ϣ���� And (gbln���� Or blnUse) And Not mTy_Para.bln��ֹ��������
        cbo���䵥λ.Enabled = mbln������Ϣ���� And (gbln���� Or blnUse) And Not mTy_Para.bln��ֹ��������
        txtIDCard.Enabled = mbln������Ϣ����
        cbo��ͥ��ַ.Enabled = gbln��ͥ��ַ Or blnUse
        cbo���ڵ�ַ.Enabled = blnUse
        padd��ͥ��ַ.Enabled = gbln��ͥ��ַ Or blnUse: padd��ͥ��ַ.ControlLock = Not (gbln��ͥ��ַ Or blnUse)
        padd���ڵ�ַ.Enabled = blnUse: padd���ڵ�ַ.ControlLock = Not blnUse
        cbo���ʽ.Enabled = gbln���ʽ Or blnUse
        If cbo���ʽ.Enabled Then
            If mbytMode = 2 And gintPriceGradeStartType >= 2 Then
                cbo���ʽ.Enabled = mTy_Para.blnԤԼ����ȷ���Һŷ�
            End If
        End If
        txt����ʱ��.Enabled = mbln������Ϣ���� And blnUse
        txt��������.Enabled = mbln������Ϣ���� And blnUse
        txt��ͥ�绰.Enabled = mbln������Ϣ���� And (gbln�绰 Or blnUse)
    End If
    
    cboҽ�����.Enabled = blnUse
    cmdLookup.Enabled = txtPatient.Enabled And Not txtPatient.Locked
    cmdLookup.Enabled = cmdLookup.Enabled And Not (mblnStation And mTy_Para.bln�Һű���ˢ��)
    If Not txtPatient.Enabled And Not blnReservePati Then
        mstrPrePati = ""
        txtPatient.Text = ""
        txt�����.Text = ""
    End If
    
    'If Not txt����.Enabled  Then txt����.Text = ""
    'If Not cbo��ͥ��ַ.Enabled Then cbo��ͥ��ַ.Text = ""
    
    If Not cbo�Ա�.Enabled And gstr�Ա� <> "��" And txtPatient.Text <> mstrPrePati And mrsInfo Is Nothing Then
        Call SetCboDefault(cbo�Ա�)
    ElseIf gstr�Ա� = "��" And txtPatient.Text <> mstrPrePati Then
        cbo�Ա�.ListIndex = -1
    End If
    If cbo���ʽ.ListIndex = -1 Then Call SetCboDefault(cbo���ʽ)
End Sub

Private Sub Fillҽ��(ByVal lng����ID As Long)
'���ܣ����ݿ��Ҷ�ȡ����ҽ�������б�
    Dim strSQL As String
        
    On Error GoTo errH
    If mrsDoctor.State = 1 Then
        mrsDoctor.Filter = "����id=" & lng����ID
        
        Do While Not mrsDoctor.EOF
            cboҽ��.AddItem IIf(IsNull(mrsDoctor!����), "", mrsDoctor!���� & "-") & mrsDoctor!����
            cboҽ��.ItemData(cboҽ��.NewIndex) = mrsDoctor!ID
            mrsDoctor.MoveNext
        Loop
        If cboҽ��.ListCount > 0 Then
            cboҽ��.ListIndex = 0
            cboҽ��.TabStop = gblnҽ�� And Not mblnStation
            
            mstrҽ������ = Mid(cboҽ��.Text, InStr(1, cboҽ��.Text, "-") + 1)
            mlngҽ��ID = cboҽ��.ItemData(cboҽ��.ListIndex)
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub GetAllҽ��()
    Dim strSQL As String
    On Error GoTo errH
    
    strSQL = "Select a.Id, a.����, Upper(a.����) As ����,b.����id,a.���" & _
            " From ��Ա�� a, ������Ա b, ��Ա����˵�� c" & _
            " Where a.Id = b.��Աid And a.Id = c.��Աid And c.��Ա���� = [1] And (a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null) " & _
            " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
            " Order By a.���� Desc"
    Set mrsDoctor = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, "ҽ��")
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function GetRoom(lng��¼ID As Long) As String
'���ܣ����ݺű�ķ��﷽ʽ��ȡ�ű������
    Dim strSQL As String, strRoomIDs As String
    Dim rsTmp As ADODB.Recordset, rsRoom As ADODB.Recordset
    
    On Error GoTo errH
            
    strSQL = "Select ID,Nvl(���﷽ʽ,0) as ���� From �ٴ������¼ Where ID = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng��¼ID)
    
    If rsTmp.EOF Then Exit Function
    If rsTmp!���� = 0 Then Exit Function '������
    
    '�������
    If rsTmp!���� = 1 Then
        'ָ������
        strSQL = "Select B.���� As �������� From �ٴ��������Ҽ�¼ A,�������� B Where A.����ID=B.ID And A.��¼ID = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CLng(rsTmp!ID))
        If Not rsTmp.EOF Then GetRoom = rsTmp!��������
    ElseIf rsTmp!���� = 2 Then
        '��̬����ø��ű���Һ�δ�������ٵ�����   //todoδ����ԤԼ�Һ�
        strSQL = _
            " Select ��������,Sum(NUM) as NUM From (" & _
                " Select B.���� As ��������,0 as NUM From �ٴ��������Ҽ�¼ A,�������� B Where A.����ID = B.ID And ��¼ID=[1]" & _
                " Union ALL" & _
                " Select ����,Count(����) as NUM From ���˹Һż�¼" & _
                " Where Nvl(ִ��״̬,0)=0 And ��¼����=1 and ��¼״̬=1 and  ����ʱ�� Between Trunc(Sysdate) And Sysdate And �����¼ID = [2]" & _
                " And ���� IN (Select D.���� As �������� From �ٴ��������Ҽ�¼ C,�������� D Where C.��¼ID=[1] And C.����ID = D.ID )" & _
                " Group by ����)" & _
            " Group by �������� Order by Num"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CLng(rsTmp!ID), lng��¼ID)
        If Not rsTmp.EOF Then GetRoom = rsTmp!��������
    ElseIf rsTmp!���� = 3 Then
        'ƽ�������ǰ����=1��ʾ�´�Ӧȡ�ĵ�ǰ����
        strSQL = "Select * From �ٴ��������Ҽ�¼ Where ��¼ID=" & rsTmp!ID
'        strSQL = "Select A.��¼ID,B.���� As ��������,A.��ǰ���� From �ٴ��������Ҽ�¼ A,�������� B Where A.����ID=B.ID And A.��¼ID=" & rsTmp!ID
        Set rsTmp = New ADODB.Recordset
        Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption, adOpenDynamic, adLockOptimistic)
        If Not rsTmp.EOF Then
            Do While Not rsTmp.EOF
                If IIf(IsNull(rsTmp!��ǰ����), 0, rsTmp!��ǰ����) = 1 Then
                    strRoomIDs = rsTmp!����ID
                    rsTmp!��ǰ���� = 0
                    rsTmp.MoveNext
                    If rsTmp.EOF Then rsTmp.MoveFirst
                    rsTmp!��ǰ���� = 1
                    rsTmp.Update
                    Exit Do
                End If
                rsTmp.MoveNext
            Loop
            '�����һ��ƽ������
            If strRoomIDs = "" Then
                rsTmp.MoveFirst
                strRoomIDs = rsTmp!����ID
                rsTmp.MoveNext
                If rsTmp.EOF Then rsTmp.MoveFirst
                rsTmp!��ǰ���� = 1
                rsTmp.Update
            End If
        End If
        If strRoomIDs <> "" Then
            strSQL = "Select ���� From �������� Where ID = [1]"
            Set rsRoom = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strRoomIDs)
            If Not rsRoom.EOF Then
                GetRoom = rsRoom!����
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

Private Function GetActualCash(ByVal lng����ID As Long) As Currency
'���ܣ���ȡ���ιҺ�ҽ��������ֽ�֧�����ݽ��
'200510byZT
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    '���˺�:26242
    '   ԭ����û�м��Ͼ��￨��(���￨����������һ������ID,��Ҫ���շ���ʱ��������

    strSQL = "" & _
    "   Select Sum(��Ԥ��) As ��� " & _
    "   From ����Ԥ����¼ A, ���㷽ʽ B " & _
    "   Where A.���㷽ʽ = B.���� And B.���� = 1 And " & _
    "         (A.�տ�ʱ��, A.����id) In (Select �տ�ʱ��, ����id From ����Ԥ����¼ Where ��¼���� = 4 And ����id = [1])"
    
    
    'strSQL = "" & _
    "   Select A.��Ԥ�� as ���" & _
    "   From ����Ԥ����¼ A,���㷽ʽ B" & _
    "   Where A.���㷽ʽ=B.���� And B.����=1 And A.��¼����=4 And A.����ID=[1] " & _
    "   "
    
    '���Ͽ��Ѵ���
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID)
    If Not rsTmp.EOF Then
        GetActualCash = Nvl(rsTmp!���, 0)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Init�ѱ�(bln���� As Boolean, Optional blnKeepIndex As Boolean) As Boolean
'������bln����=�Ƿ�������޳������Ŀ
'      blnKeepIndex=�Ƿ񱣳�ԭ�еķѱ�ѡ��
    Dim strSQL As String, i As Integer
    Dim strKeep As String
    Dim strȱʡ�ѱ� As String
    
    On Error GoTo errH
    
    strKeep = cbo�ѱ�.Text      '������ǰ�ķѱ�,�п������ڵ�ϵͳ����û�и÷ѱ���
    If strKeep <> "" Then strKeep = Mid(strKeep, InStr(1, strKeep, "-") + 1)
    strȱʡ�ѱ� = gstr�ѱ�      '����ȱʡ�ѱ�,���Ϊ��,������ȡϵͳȱʡ
    
    '72168,Ƚ����,2014/4/22,�Һ�ʱͨ���Һſ���ȷ����ѡ�ѱ�
    If mrs�ѱ� Is Nothing Then '�״ε��øú���ʱ[bln����]Ϊtrue
        Set mrs�ѱ� = New ADODB.Recordset
        '�ѱ�:���Ψһ����Ŀ(������ȱʡ�ѱ�),�����ǳ���,������Ч�ڼ估����
        strSQL = "Select a.����, a.����, a.����, Nvl(a.���޳���, 0) As ����," & _
                "       Nvl(a.ȱʡ��־, 0) As ȱʡ, Nvl(b.����id, 0) As ����id" & _
                " From �ѱ� A, �ѱ����ÿ��� B" & _
                " Where a.���� = b.�ѱ�(+) And a.���� = 1" & _
                "      And Trunc(Sysdate) Between Nvl(a.��Ч��ʼ, To_Date('1900-01-01', 'YYYY-MM-DD'))" & _
                "                         And Nvl(a.��Ч����, To_Date('3000-01-01', 'YYYY-MM-DD'))" & _
                "      And Nvl(a.�������, 3) In (1, 3)" & _
                " Order By a.����"
        Call zlDatabase.OpenRecordset(mrs�ѱ�, strSQL, Me.Caption)
    End If
    
    If mrs�ѱ� Is Nothing Then Exit Function
    If bln���� Then
        mrs�ѱ�.Filter = "����id=" & mlng�Һſ���ID & " or ����id=0"   'adFilterNone
    Else                        '��������޳������Ŀ
        mrs�ѱ�.Filter = "(����=0 and ����id=" & mlng�Һſ���ID & ") or (����=0 and ����id=0)"
    End If
    If mrs�ѱ�.RecordCount > 0 Then mrs�ѱ�.MoveFirst
    
    cbo�ѱ�.Clear: mstrPre�ѱ� = ""
    Do While Not mrs�ѱ�.EOF
        cbo�ѱ�.AddItem mrs�ѱ�!���� & "-" & mrs�ѱ�!����
        '��¼������Ŀ:�����Ǳ���ȱʡ��ϵͳȱʡ
        cbo�ѱ�.ItemData(cbo�ѱ�.NewIndex) = IIf(mrs�ѱ�!���� = 1, 2, 0)
        
        If strȱʡ�ѱ� = "" Then    'û�б���ȱʡʱȡϵͳȱʡ
            If mrs�ѱ�!ȱʡ = 1 Then strȱʡ�ѱ� = mrs�ѱ�!����
        End If
        mrs�ѱ�.MoveNext
    Loop
    
    If blnKeepIndex And Not mrsInfo Is Nothing Then
        If Not mrsInfo.EOF Then Call zlControl.CboLocate(cbo�ѱ�, Nvl(mrsInfo!�ѱ�))
    End If
    If blnKeepIndex And strKeep <> "" Then Call zlControl.CboLocate(cbo�ѱ�, strKeep)

    If cbo�ѱ�.ListIndex = -1 Then Call zlControl.CboLocate(cbo�ѱ�, strȱʡ�ѱ�)
    
    If cbo�ѱ�.ListIndex = -1 Then If cbo�ѱ�.ListCount > 0 Then cbo�ѱ�.ListIndex = 0
    If cbo�ѱ�.ListIndex <> -1 Then cbo�ѱ�.ItemData(cbo�ѱ�.ListIndex) = 1
            
    Init�ѱ� = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function PatiExist(strCard As String) As Boolean
'���ܣ��ж��Ƿ�ȷʵ���ڸÿ��ŵĳֿ�����,��ΪסԺ���˲����ڴ�ˢ��
    Dim rsTmp As ADODB.Recordset
    Dim strSQL  As String
    
    On Error GoTo errH
    
    strSQL = "Select a.���￨�� " & vbNewLine & _
             "From ������Ϣ A, ����ҽ�ƿ���Ϣ B, ҽ�ƿ���� C " & vbNewLine & _
             "Where a.���￨�� = b.���� And c.�ض���Ŀ = '���￨' And b.�����id = c.Id And a.��Ժ = 1 And b.���� = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strCard)
    PatiExist = Not rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function SetIdentifyLocked(blnLocked As Boolean)
'���ܣ�����ҽ�������֤�������޸ĵ���Ϣ��Ŀ
    txtPatient.Locked = blnLocked
    cbo�Ա�.Locked = blnLocked
    cbo�Ա�.TabStop = Not blnLocked
    txt����.Locked = blnLocked
    txt����.TabStop = Not blnLocked
    cbo���䵥λ.Locked = blnLocked
    cbo���䵥λ.TabStop = Not blnLocked
    cbo���ʽ.Locked = blnLocked
    cbo���ʽ.TabStop = Not blnLocked
    cmdLookup.Enabled = IIf(Not blnLocked, txtPatient.Enabled, Not blnLocked)
    cmdLookup.Enabled = cmdLookup.Enabled And Not (mblnStation And mTy_Para.bln�Һű���ˢ��)
    
    If blnLocked Then
        txtPatient.BackColor = &HE0E0E0
    Else
        txtPatient.BackColor = &HFFFFFF
    End If
    txt����.BackColor = txtPatient.BackColor
    cbo�Ա�.BackColor = txtPatient.BackColor
    cbo���䵥λ.BackColor = txtPatient.BackColor
    cbo���ʽ.BackColor = txtPatient.BackColor
    
    With mobjfrmPatiInfo
        .txtPatient.Locked = blnLocked
        .cbo�Ա�.Locked = blnLocked
        .txt����.Locked = blnLocked
        .cbo���䵥λ.Locked = blnLocked
        .cbo���ʽ.Locked = blnLocked
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
    If mbln�����Һ� Then
        cbo���㷽ʽ.Enabled = False
        Call InitVsfPay(True)
        txt����Ӧ��.Text = Format(mcurӦ��, "0.00")
    Else
        cbo���㷽ʽ.Enabled = gbln���㷽ʽ
        Call InitVsfPay
        If mblnCancel Then
            mcur�ϼ� = 0
            mcurӦ�� = 0
            txt�ϼ�.Text = "0.00"
            txt����Ӧ��.Text = "0.00"
        End If
    End If
End Sub

Private Sub InitVsfPay(Optional ByVal bln�����Һ� As Boolean)
    Dim i As Long, j As Long
    Dim blnFinish As Boolean, blnDraw As Boolean
    
    If bln�����Һ� Then
        With vsfPay
            For i = 1 To .Rows - 1
                If .RowData(i) = 0 And .TextMatrix(i, 0) = "Ԥ����" Then
                    .TextMatrix(i, 0) = "Ԥ����"
                    .TextMatrix(i, 1) = "0.00"
                    .TextMatrix(i, 2) = ""
                    .TextMatrix(i, 3) = 1
                    .TextMatrix(i, 4) = "Ԥ����"
                    .TextMatrix(i, 5) = 1
                    .TextMatrix(i, 6) = 0
                    If mdblԤ����� = 0 Then .RowHidden(i) = True
                End If
                If .RowData(i) = 3 And .TextMatrix(i, 0) <> "" And mstr�����ʻ� <> "" Then
                    .TextMatrix(i, 0) = mstr�����ʻ�
                    .TextMatrix(i, 1) = "0.00"
                    .TextMatrix(i, 2) = ""
                    .TextMatrix(i, 3) = 1
                    .TextMatrix(i, 4) = mstr�����ʻ�
                    .TextMatrix(i, 5) = 1
                    .TextMatrix(i, 6) = 0
                    If mdbl������� = 0 Then .RowHidden(i) = True
                End If
                If (.RowData(i) = 1 Or .RowData(i) = 2) And .TextMatrix(i, 0) = NeedName(cbo���㷽ʽ.Text) Then
                    .TextMatrix(i, 7) = Val(.TextMatrix(i, 1))
                Else
                    If .RowData(i) <> 0 And .RowData(i) <> 3 And .TextMatrix(i, 0) <> "" Then
                        .TextMatrix(i, 0) = "ɾ��"
                        blnFinish = False
                    End If
                End If
            Next i
            Do While blnFinish = False
                For i = 1 To .Rows - 1
                    If .TextMatrix(i, 0) = "ɾ��" Then
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
            '����Ԥ����
            .TextMatrix(.Rows - 1, 0) = "Ԥ����"
            .TextMatrix(.Rows - 1, 1) = "0.00"
            .TextMatrix(.Rows - 1, 2) = ""
            .TextMatrix(.Rows - 1, 3) = 1
            .TextMatrix(.Rows - 1, 4) = "Ԥ����"
            .TextMatrix(.Rows - 1, 5) = 1
            .TextMatrix(.Rows - 1, 6) = 0
            .RowData(.Rows - 1) = 0
            If mdblԤ����� = 0 Then .RowHidden(.Rows - 1) = True
            .Rows = .Rows + 1
            '���ظ����ʻ�
            If mstr�����ʻ� <> "" Then
                .TextMatrix(.Rows - 1, 0) = mstr�����ʻ�
                .TextMatrix(.Rows - 1, 1) = "0.00"
                .TextMatrix(.Rows - 1, 2) = ""
                .TextMatrix(.Rows - 1, 3) = 1
                .TextMatrix(.Rows - 1, 4) = mstr�����ʻ�
                .TextMatrix(.Rows - 1, 5) = 1
                .TextMatrix(.Rows - 1, 6) = 0
                .RowData(.Rows - 1) = 3
                If mdbl������� = 0 Then .RowHidden(.Rows - 1) = True
                .Rows = .Rows + 1
            End If
            .Redraw = blnDraw
        End With
    End If
End Sub

Private Sub CalcYBMoney()
'���ܣ����㲢��ʾ��ǰҽ�����˸����ʻ�����֧�ֵĽ��
    Dim cur�ϼ� As Currency
    Dim strInfo As String, i As Long, j As Long, lng����ID As Long
    Dim curTotal As Currency
    Dim lngYBRow As Long, lngYJRow As Long
    
    If mRegistFeeMode = EM_RG_���� Then Exit Sub
    
    For i = 1 To vsfPay.Rows - 1
        If vsfPay.RowData(i) = 3 And vsfPay.TextMatrix(i, 0) <> "" Then lngYBRow = i
        If vsfPay.RowData(i) = 0 And vsfPay.TextMatrix(i, 0) <> "" Then lngYJRow = i
    Next i
    
    cur�ϼ� = GetRegistMoney(True)
    curTotal = cur�ϼ�
    If MCPAR.���ղ����� = True Then
        cur�ϼ� = cur�ϼ� - mcur����
    End If
    If mstrYBPati <> "" Then lng����ID = Val(Split(mstrYBPati, ";")(8))
    
    '���㲢��ʾ�����ʻ�֧�����
    'Ҫ��ҽ��֧�ָ����ʻ�֧����ZLHIS����ʹ�ø����ʻ�
    If mintInsure <> 0 And mstr�����ʻ� <> "" Then
        If gclsInsure.GetCapability(support�Һ�ʹ�ø����ʻ�, lng����ID, mintInsure) Then
            If mdbl������� - cur�ϼ� >= -1 * mcur����͸֧ Then
                vsfPay.TextMatrix(lngYBRow, 1) = Format(cur�ϼ�, "0.00")
                vsfPay.TextMatrix(lngYBRow, 6) = cur�ϼ�
            Else
                If mblnStation Then
                    vsfPay.TextMatrix(lngYBRow, 1) = "0.00"
                ElseIf mcur����͸֧ = 0 And mdbl������� > 0 Then
                    vsfPay.TextMatrix(lngYBRow, 1) = Format(mdbl�������, "0.00")
                    vsfPay.TextMatrix(lngYBRow, 6) = mdbl�������
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
    
    If gblnPrePayPriority And mdblԤ����� >= Val(curTotal - Val(vsfPay.TextMatrix(lngYBRow, 1))) Then
        vsfPay.TextMatrix(lngYJRow, 1) = Format(curTotal - Val(vsfPay.TextMatrix(lngYBRow, 1)), "0.00")
    End If
    If lngYJRow <> 0 Then vsfPay.TextMatrix(lngYJRow, 6) = mdblԤ�����
    
    '��ȡҽ��ͳ���������
    If mintInsure <> 0 And mstrYBPati <> "" And Not mrsItems Is Nothing Then
        mrsItems.MoveFirst
        For i = 1 To mrsItems.RecordCount
            mrsInComes.Filter = "��ĿID=" & mrsItems!��ĿID
            For j = 1 To mrsInComes.RecordCount
                strInfo = gclsInsure.GetItemInsure(lng����ID, mrsItems!��ĿID, mrsInComes!ʵ��, True, mintInsure)
                If strInfo <> "" Then
                    mrsItems!������Ŀ�� = Val(Split(strInfo, ";")(0))
                    mrsItems!���մ���id = Val(Split(strInfo, ";")(1))
                    mrsItems!���ձ��� = CStr(Split(strInfo, ";")(3))
                    mrsInComes!ͳ���� = Format(Val(Split(strInfo, ";")(2)), "0.00")
                End If
                mrsInComes.MoveNext
            Next
            mrsItems.MoveNext
        Next
    End If
    Call Set�����Һ�
End Sub

Private Sub ReCalcԤԼ���շ���()
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ����¼���ԤԼ���շ����Ŀ�������Ϣ
    '���ƣ����˺�
    '���ڣ�2010-07-16 09:38:54
    '˵����31182
    '------------------------------------------------------------------------------------------------------------------------
    Dim blnExitLoop As Boolean, i As Long, j As Long, lngRow As Long, lng����ID As Long
    Dim str�ѱ� As String, curӦ�� As Currency, curʵ��  As Currency, cur�ϼ� As Currency
    Dim cur���� As Currency, rsItems As ADODB.Recordset, rsIncomes As ADODB.Recordset
    
     '31182:ԤԼ����ʱ,ҲҪ��ȡ��Ӧ�Ŀ���
    'ɾ�����ѵ�
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
    str�ѱ� = NeedName(cbo�ѱ�.Text)
    If mrsBill Is Nothing Then Exit Sub
    
    mrsBill.MoveFirst
    Call ReadRegistPrice(mrsBill!�շ�ϸĿID, mbln����������, mblnAddCardItem, str�ѱ�, rsItems, rsIncomes, 0, mintInsure, _
        txt�ű�.Text, 10, mlng�Һſ���ID, mobjfrmPatiInfo.mstrPriceGrade, _
        IIf(dtpAppointmentDate.Visible Or mbytMode = 2, Format(dtpAppointmentDate.Value, "yyyy-mm-dd") & " 23:59:59", ""), gCurSendCard.lng�շ�ϸĿID)
    
    If mintInsure <> 0 Then
        If MCPAR.�Һż����Ŀ = True Then
            If gclsInsure.CheckItem(mintInsure, 2, 0, rsItems) = False Then
                MsgBox "ҽ�������շ���Ŀ���ʧ�ܣ����ܼ����Һţ�", vbInformation, gstrSysName
                Call ClearBill: Exit Sub
            End If
        End If
    End If
    If mrsInfo Is Nothing Then
        lng����ID = 0
    Else
        If mrsInfo.RecordCount = 0 Then
            lng����ID = 0
        Else
            lng����ID = Val(Nvl(mrsInfo!����ID))
        End If
    End If
    Call ReadRegistPrice(0, False, mblnAddCardItem, str�ѱ�, mrsItems, mrsInComes, lng����ID, mintInsure, _
            txt�ű�.Text, mbytMode, , mobjfrmPatiInfo.mstrPriceGrade, _
    IIf(dtpAppointmentDate.Visible Or mbytMode = 2, Format(dtpAppointmentDate.Value, "yyyy-mm-dd") & " 23:59:59", ""), gCurSendCard.lng�շ�ϸĿID)
    
    '��ʾ��������
     If Not mrsItems Is Nothing Then
         vsfMoney.Redraw = False
         curӦ�� = 0: curʵ�� = 0
         For j = 1 To vsfMoney.Rows - 1
             If vsfMoney.RowData(j) = 0 Then    '��Ϊ��ȡ���ݵ�ʱ��,û�м���RowData����
                 curʵ�� = Val(vsfMoney.TextMatrix(j, 2))
                cur�ϼ� = cur�ϼ� + curʵ��
             End If
         Next
         lngRow = vsfMoney.Rows - 1
         vsfMoney.Rows = vsfMoney.Rows + mrsItems.RecordCount
         mrsItems.MoveFirst
        
         For i = 1 To mrsItems.RecordCount
             vsfMoney.RowData(lngRow + i) = mrsItems!��ĿID
             vsfMoney.TextMatrix(lngRow + i, 0) = mrsItems!��Ŀ����
             mrsInComes.Filter = "��ĿID=" & mrsItems!��ĿID
            curӦ�� = 0: curʵ�� = 0
             For j = 1 To mrsInComes.RecordCount
                 curӦ�� = curӦ�� + mrsInComes!Ӧ��
                 curʵ�� = curʵ�� + mrsInComes!ʵ��
                 If mrsItems!���� = 3 Then cur���� = cur���� + mrsInComes!ʵ��
                 mrsInComes.MoveNext
             Next
             vsfMoney.TextMatrix(lngRow + i, 1) = Format(curӦ��, "0.00")
             vsfMoney.TextMatrix(lngRow + i, 2) = Format(curʵ��, "0.00")
             cur�ϼ� = cur�ϼ� + curʵ��
             mrsItems.MoveNext
         Next
         vsfMoney.Redraw = True
     End If
End Sub

Private Sub ShowAcceptFromInput()
    Dim lng��Ŀid As Long, bln���� As Boolean, str�ѱ� As String
    Dim curӦ�� As Currency, curʵ�� As Currency, cur�ϼ� As Currency, cur���� As Currency
    Dim lngRow As Long, i As Long, j As Long
    Dim dblMoney As Double
    
    If mbytMode = 2 And Not mrsBill Is Nothing Then
            mrsBill.MoveFirst
            '���ԤԼʱ,û�н�������,����ʱ���Ը��ķѱ�,
            If Nvl(mrsBill!�ѱ�) <> NeedName(cbo�ѱ�.Text) Then
                '�ѱ�һ�� ��Ҫ���¼���
                str�ѱ� = NeedName(cbo�ѱ�.Text)
                mrsBill.MoveFirst
                vsfMoney.Rows = mrsBill.RecordCount + 1
                For i = 1 To mrsBill.RecordCount
                    vsfMoney.TextMatrix(i, 0) = mrsBill!��Ŀ
                    vsfMoney.TextMatrix(i, 1) = Format(mrsBill!Ӧ��, "0.00")
'                    dblMoney = Val(Nvl(mrsBill!ʵ��))
                    curʵ�� = GetActualMoney(str�ѱ�, mrsBill!������ĿID, mrsBill!Ӧ��, mrsBill!�շ�ϸĿID)
                    vsfMoney.TextMatrix(i, 2) = Format(curʵ��, "0.00")
                    cur�ϼ� = cur�ϼ� + curʵ��
                    mrsBill.MoveNext
                Next
                txt�ϼ�.Text = Format(cur�ϼ�, "0.00")
                lbl�ϼ�.Caption = txt�ϼ�.Text
            Else
                mrsBill.MoveFirst
                vsfMoney.Rows = mrsBill.RecordCount + 1
                For i = 1 To mrsBill.RecordCount
                    vsfMoney.TextMatrix(i, 0) = mrsBill!��Ŀ
                    vsfMoney.TextMatrix(i, 1) = Format(mrsBill!Ӧ��, "0.00")

                    vsfMoney.TextMatrix(i, 2) = Format(mrsBill!ʵ��, "0.00")
                    cur�ϼ� = cur�ϼ� + mrsBill!ʵ��
                    mrsBill.MoveNext
                Next
                txt�ϼ�.Text = Format(cur�ϼ�, "0.00")
                lbl�ϼ�.Caption = txt�ϼ�.Text
            End If
        End If
        '����:31182
        cur�ϼ� = Val(txt�ϼ�.Text)
        Call ReCalcԤԼ���շ���
          '60171 ԤԼ����ʱ,��Ҫ���¼��㿨�Ѻ͹Һŷ�,��ʱ�����������Һ�
        If Not mrsItems Is Nothing Then
            cur�ϼ� = GetRegistMoney
        End If
End Sub

Private Sub ShowRegistFromInput()
    '���ܣ����ݵ�ǰ��������ĺű�,��ȡ�Һŷ��ü�,��ʾ�ڱ����
    Dim lng��Ŀid As Long, bln���� As Boolean, str�ѱ� As String
    Dim curӦ�� As Currency, curʵ�� As Currency, cur�ϼ� As Currency, cur���� As Currency
    Dim lngRow As Long, i As Long, j As Long
    Dim dblMoney As Double, rsItems As ADODB.Recordset, rsIncomes As ADODB.Recordset
    Dim str��¼ID As String, strTemp As String
    Dim strReadSQL As String, rsRead As ADODB.Recordset
    Dim strҽ������ As String, lng����ID As Long
    If mblnReadBooking Then Exit Sub
    If mblnBuyHisBook = False Then
        If (mbytMode = 2 Or mbytMode = 0 And mbytInState = 0 And mstrNoIn <> "") And mTy_Para.blnԤԼ����ȷ���Һŷ� = False Then
            If mbytMode = 2 And Not mrsBill Is Nothing Then
                mrsBill.MoveFirst
                '���ԤԼʱ,û�н�������,����ʱ���Ը��ķѱ�,
                If Nvl(mrsBill!�ѱ�) <> NeedName(cbo�ѱ�.Text) Then
                    '�ѱ�һ�� ��Ҫ���¼���
                    str�ѱ� = NeedName(cbo�ѱ�.Text)
                    mrsBill.MoveFirst
                    vsfMoney.Rows = mrsBill.RecordCount + 1
                    For i = 1 To mrsBill.RecordCount
                        vsfMoney.TextMatrix(i, 0) = mrsBill!��Ŀ
                        vsfMoney.TextMatrix(i, 1) = Format(mrsBill!Ӧ��, "0.00")
    '                    dblMoney = Val(Nvl(mrsBill!ʵ��))
                        curʵ�� = GetActualMoney(str�ѱ�, mrsBill!������ĿID, mrsBill!Ӧ��, mrsBill!�շ�ϸĿID)
                        vsfMoney.TextMatrix(i, 2) = Format(curʵ��, "0.00")
                        cur�ϼ� = cur�ϼ� + curʵ��
                        mrsBill.MoveNext
                    Next
                    txt�ϼ�.Text = Format(cur�ϼ�, "0.00")
                    lbl�ϼ�.Caption = txt�ϼ�.Text
                Else
                    mrsBill.MoveFirst
                    vsfMoney.Rows = mrsBill.RecordCount + 1
                    For i = 1 To mrsBill.RecordCount
                        vsfMoney.TextMatrix(i, 0) = mrsBill!��Ŀ
                        vsfMoney.TextMatrix(i, 1) = Format(mrsBill!Ӧ��, "0.00")
    
                        vsfMoney.TextMatrix(i, 2) = Format(mrsBill!ʵ��, "0.00")
                        cur�ϼ� = cur�ϼ� + mrsBill!ʵ��
                        mrsBill.MoveNext
                    Next
                    txt�ϼ�.Text = Format(cur�ϼ�, "0.00")
                    lbl�ϼ�.Caption = txt�ϼ�.Text
                End If
            End If
            '����:31182
            cur�ϼ� = Val(txt�ϼ�.Text)
            Call ReCalcԤԼ���շ���
              '60171 ԤԼ����ʱ,��Ҫ���¼��㿨�Ѻ͹Һŷ�,��ʱ�����������Һ�
            If Not mrsItems Is Nothing Then
                cur�ϼ� = GetRegistMoney
            End If
            GoTo CalcOther:
            Exit Sub
        End If
    End If
    If chkCancel.Value = 1 Then Exit Sub
    If chkPrint.Value = 1 Then Exit Sub

    Call ClearMoney

    '��ȡ�Һŷ���
    If txt�ű�.Text = "+" Then    '��������
        lng��Ŀid = 0
        bln���� = True

        chk������.Enabled = False
        chk������.Value = 0

        mbln������ = False
        mlng�Һſ���ID = UserInfo.����ID
        mstrҽ������ = "": mlngҽ��ID = 0
        txt����.Text = ""
        cboҽ��.Clear
        cboҽ��.Enabled = False
        lbl��.Visible = False
    ElseIf txt�ű�.Text <> "" Then
        '134441:���ϴ���2019/1/12�������˵��ȷ���˺ű���б�ѡ���ַ����˸ı�
        If mlngPreRow <> 0 Then vsfPlan.Row = mlngPreRow
        If vsfPlan.Row > vsfPlan.Rows - 1 Then
            lngRow = 0
        Else
            lngRow = vsfPlan.Row
        End If
        If lngRow = 0 Then
            mbln������ = False
            mlng�Һſ���ID = 0
            mstrҽ������ = ""
            mlngҽ��ID = 0
            
            If mbytMode <> 2 Then
                chk������.Enabled = False
                chk������.Value = 0
            End If
            txt����.Text = ""
            cboҽ��.Clear
            lbl��.Visible = False
            Exit Sub
        End If

        lng����ID = 0
        If Not mrsInfo Is Nothing Then
            If Not mrsInfo.EOF Then lng����ID = Val(Nvl(mrsInfo!����ID))
        End If
        str��¼ID = ""
        strTemp = vsfPlan.Cell(flexcpData, lngRow, vsfPlan.ColIndex("IDS"))
        If Val(strTemp) <> 0 Then
            str��¼ID = "2|" & Val(strTemp)
        End If
        If str��¼ID = "" Then str��¼ID = "3|" & vsfPlan.TextMatrix(lngRow, vsfPlan.ColIndex("�ű�"))
        
        lng��Ŀid = Val(Split(vsfPlan.TextMatrix(lngRow, GetCol("IDS")), ",")(1))
        strReadSQL = "Select Zl_Custom_Getregeventitem([1],[2],[3],[4],[5],[6],[7]) As ��ĿID From Dual"
        Set rsRead = zlDatabase.OpenSQLRecord(strReadSQL, Me.Caption, lng����ID, txtPatient.Text, txtIDCard.Text, _
                                            CDate(IIf(IsDate(txt��������.Text) = False, "3000-01-01", txt��������.Text)), NeedName(cbo�Ա�.Text), txt����.Text & IIf(cbo���䵥λ.Visible, cbo���䵥λ.Text, ""), str��¼ID)
        If Not rsRead.EOF Then
            If Val(Nvl(rsRead!��ĿID)) <> 0 Then lng��Ŀid = Val(Nvl(rsRead!��ĿID))
        End If
        bln���� = chk������.Value = 1

        If mbytMode <> 2 Then chk������.Enabled = True
        mbln������ = vsfPlan.TextMatrix(lngRow, GetCol("����")) <> ""
        lbl��.Visible = vsfPlan.RowData(lngRow) < 0
        cboҽ��.Enabled = False
       
        mlng�Һſ���ID = Abs(vsfPlan.RowData(lngRow))
        strҽ������ = NeedName(cboҽ��.Text)
        mstrҽ������ = vsfPlan.TextMatrix(lngRow, GetCol("ҽ��"))
        mlngҽ��ID = CLng(Split(vsfPlan.TextMatrix(lngRow, GetCol("IDS")), ",")(2))

        txt����.Text = vsfPlan.TextMatrix(lngRow, GetCol("����"))
        cboҽ��.Clear
        cboҽ��.TabStop = False
        If mstrҽ������ <> "" Then
            cboҽ��.AddItem mstrҽ������
            cboҽ��.ItemData(cboҽ��.NewIndex) = mlngҽ��ID
            cboҽ��.ListIndex = 0
        ElseIf Not mblnStation Then     '���Ҫ����ҽ��,�ű�û��ȷ��ҽ��,�г����ҿ�ѡҽ��
            cboҽ��.Enabled = gblnҽ��
            If gblnҽ�� Then
                Call Fillҽ��(mlng�Һſ���ID)
                zlControl.CboLocate cboҽ��, strҽ������
                mstrҽ������ = NeedName(cboҽ��.Text)
                If mstrҽ������ = "" Then
                    mlngҽ��ID = 0
                Else
                    mlngҽ��ID = cboҽ��.ItemData(cboҽ��.ListIndex)
                End If
            End If
        End If
        
    End If
    Call AdjustInfoPosition
    str�ѱ� = NeedName(cbo�ѱ�.Text)
    Set mrsItems = Nothing
    Set mrsInComes = Nothing
    Call ReadRegistPrice(lng��Ŀid, bln����, mblnAddCardItem, str�ѱ�, rsItems, rsIncomes, 0, mintInsure, _
        txt�ű�.Text, 10, mlng�Һſ���ID, mobjfrmPatiInfo.mstrPriceGrade, _
        IIf(dtpAppointmentDate.Visible Or mbytMode = 2, Format(dtpAppointmentDate.Value, "yyyy-mm-dd") & " 23:59:59", ""), gCurSendCard.lng�շ�ϸĿID)
    
    If mintInsure <> 0 Then
        If MCPAR.�Һż����Ŀ = True Then
            If gclsInsure.CheckItem(mintInsure, 2, 0, rsItems) = False Then
                MsgBox "ҽ�������շ���Ŀ���ʧ�ܣ����ܼ����Һţ�", vbInformation, gstrSysName
                mblnUnload = True
                Call ClearBill: Exit Sub
            End If
        End If
    End If
    
    Call ReadRegistPrice(lng��Ŀid, bln����, mblnAddCardItem, str�ѱ�, mrsItems, mrsInComes, lng����ID, _
        mintInsure, txt�ű�.Text, mbytMode, , mobjfrmPatiInfo.mstrPriceGrade, _
    IIf(dtpAppointmentDate.Visible Or mbytMode = 2, Format(dtpAppointmentDate.Value, "yyyy-mm-dd") & " 23:59:59", ""), gCurSendCard.lng�շ�ϸĿID)

    '��ʾ�Һŷ���
    If Not mrsItems Is Nothing Then
        vsfMoney.Redraw = False
        vsfMoney.Rows = mrsItems.RecordCount + 1
        mrsItems.MoveFirst
        For i = 1 To mrsItems.RecordCount
            If mrsItems!���� = 4 Then
                vsfMoney.RowData(i) = mrsItems!��ĿID
            End If
            vsfMoney.TextMatrix(i, 0) = mrsItems!��Ŀ����

            curӦ�� = 0: curʵ�� = 0
            mrsInComes.Filter = "��ĿID=" & mrsItems!��ĿID
            For j = 1 To mrsInComes.RecordCount
                curӦ�� = curӦ�� + mrsInComes!Ӧ��
                curʵ�� = curʵ�� + mrsInComes!ʵ��
                If mrsItems!���� = 3 Then cur���� = cur���� + mrsInComes!ʵ��
                mrsInComes.MoveNext
            Next

            vsfMoney.TextMatrix(i, 1) = Format(curӦ��, "0.00")
            vsfMoney.TextMatrix(i, 2) = Format(curʵ��, "0.00")
            cur�ϼ� = cur�ϼ� + curʵ��
            mcur���� = cur����
            mrsItems.MoveNext
        Next
        vsfMoney.Redraw = True

    End If

CalcOther:
    'Ԥ����֧����������
    '77786,Ƚ����,2014-9-2,��ѡ����ʹ��Ԥ����ɿ�,�Һ�ʱ,û��Ĭ�ϼ��ٳ��
    '74550,Ƚ����,2014-7-2,�ڲ�����Ժ����,ҽ��������ҽ��վ�Һ�ʱ�ܹ�ѡ����㷽ʽ(��������Ϊ7��һ��ͨ����)
    If (gblnPrePayPriority Or (mblnStation And Not mblnStationPrice And Not cbo���㷽ʽ.Visible)) And mdblԤ����� >= cur�ϼ� And mblnAddCardItem = False Then
        For i = 1 To vsfPay.Rows - 1
            If vsfPay.RowData(i) = 0 And vsfPay.TextMatrix(i, 0) <> "" Then
                vsfPay.TextMatrix(i, 1) = Format(cur�ϼ�, "0.00")
                vsfPay.TextMatrix(i, 6) = mdblԤ�����
            End If
        Next i
    Else
        For i = 1 To vsfPay.Rows - 1
            If vsfPay.RowData(i) = 0 And vsfPay.TextMatrix(i, 0) <> "" Then
                vsfPay.TextMatrix(i, 1) = Format(0, "0.00")
                vsfPay.TextMatrix(i, 6) = mdblԤ�����
            End If
        Next i
    End If
    
    '���Ѻ͹Һŷ���һ����ʱ,����Ԥ����
    If mblnAddCardItem Then ShowDeposit (False)
    
    
    '���㲢��ʾ�����ʻ�֧����
    Call CalcYBMoney
     
    '��ʾ�ۼӷ���
    txt�ϼ�.Text = Format(cur�ϼ� + mcur�ϼ�, "0.00")
    lbl�ϼ�.Caption = txt�ϼ�.Text
    
    Call Set�����Һ�
    '��ʾ����Ѻ�,���㲡����
    If Me.Visible Then
        lblFree.Visible = (cur�ϼ� - cur����) = 0
    Else
        lblFree.Visible = False
    End If
End Sub

Private Sub txt�Ҳ�_GotFocus()
    Call zlControl.TxtSelAll(txt�Ҳ�)
End Sub

Private Sub YBIdentifyCancel()
'���ܣ�ȡ��ҽ�����������֤
    Dim lng����ID As Long
    
    If mbytInState = 0 And mintInsure <> 0 And mstrYBPati <> "" And txtPatient.Text <> "" Then
        If UBound(Split(mstrYBPati, ";")) >= 8 Then
            If IsNumeric(Split(mstrYBPati, ";")(8)) And Val(Split(mstrYBPati, ";")(8)) <> 0 Then
                lng����ID = Val(CLng(Split(mstrYBPati, ";")(8)))
            End If
        End If
        If lng����ID <> 0 Then
            Call gclsInsure.IdentifyCancel(3, lng����ID, mintInsure)
        End If
    End If
End Sub



Private Function StationDelete(ByVal strNO As String, Optional str����NO As String) As Boolean
'���ܣ����ָ���ĹҺŵ��Ƿ������˺�(δ�շ�,������)
'���أ�str����NO=ͬʱҪɾ���Ļ��۵�
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, lng����ID As Long
    
    On Error GoTo errH
    
    '1-ִ���˼�����״̬�ж�
    strSQL = "Select ����ID,ִ����,ִ��״̬ From ���˹Һż�¼ Where NO=[1] and ��¼����=1 and ��¼״̬=1"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
    If rsTmp.EOF Then
        MsgBox "ָ���ĹҺŵ������ڣ��õ��ݿ����Ѿ��˺š�", vbInformation, gstrSysName
        Exit Function
    End If
    If Nvl(rsTmp!ִ��״̬, 0) <> 2 Then
        MsgBox "�ò���" & Decode(Nvl(rsTmp!ִ��״̬, 0), 0, "������ֱ�ӹҺž���״̬", 1, "�Ѿ���ɾ���") & "�������˺š�", vbInformation, gstrSysName
        Exit Function
    End If
    If Nvl(rsTmp!ִ����) <> UserInfo.���� Then
        MsgBox "�ò��˲�����������ҵĺţ������˺š�", vbInformation, gstrSysName
        Exit Function
    End If
    lng����ID = rsTmp!����ID
    
    '2-�ҺŽ���ж�:���ֽ��������Ԥ������Ĳ���ҽ��վ�Һ�
    strSQL = "Select Sum(��Ԥ��) as ��� From ����Ԥ����¼ A,���㷽ʽ B " & _
            " Where A.���㷽ʽ=B.���� And A.��¼����=4 And A.��¼״̬=1 And A.NO=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
    If Not rsTmp.EOF Then
        If Nvl(rsTmp!���, 0) > 0 Then
            MsgBox "�ùҺŲ������������㷽ʽ�������������˺š�", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    '3-�շ��ж�
    strSQL = "Select NO,��¼״̬ From ������ü�¼ " & _
            " Where ��¼����=1 And ����ID=[1] And ��¼״̬ IN(0,1,3) And ���=1 And ժҪ Like [2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, "%" & strNO & "%")
    If Not rsTmp.EOF Then
        If Nvl(rsTmp!��¼״̬, 0) = 1 Then
            MsgBox "�ùҺŵ���Ӧ�ķ����Ѿ��������շѣ������˺š�", vbInformation, gstrSysName
            Exit Function
        ElseIf Nvl(rsTmp!��¼״̬, 0) = 0 Then
            str����NO = rsTmp!NO
        End If
    End If
    
    '4-ҽ���ж�
    strSQL = "Select Count(*) as Num From ����ҽ����¼ Where �Һŵ�=[1] And ҽ��״̬<>4"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
    If Nvl(rsTmp!Num, 0) > 0 Then
        MsgBox "�����Ѿ��´�ҽ���������˺š�", vbInformation, gstrSysName
        Exit Function
    End If
    
    StationDelete = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Function Check����(ByVal lng����ID As Long, ByVal lngִ�в���ID As Long) As Boolean
'����:�жϲ����Ƿ��ٴε�����ͬ�ٴ����ʵ��ٴ����ҡ��Һ�
'     �����ҹ��ŵ�,��ס��Ժ��,���ﲻ��ȷ��ʱ��
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    strSQL = "Select Zl1_Fun_Getreturnvisit([1],[2]) As �����־ From Dual"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, lngִ�в���ID)
    Check���� = Val(Nvl(rsTmp!�����־)) = 1
End Function

Private Sub Set�����Һ�()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���¼���Ӧ�ɿ�ϼ���
    '����:���˺�
    '����:2009-12-02 12:02:51
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����ɿ�ϼƸ��ı���
    Dim strSQL As String, rsTemp As ADODB.Recordset
'    strSQL = "Select ����" & vbNewLine & _
'                        "From ���㷽ʽ" & vbNewLine & _
'                        "Where ���� = [1] And Rownum < 2" & vbNewLine & _
'                        "Union" & vbNewLine & _
'                        "Select a.����" & vbNewLine & _
'                        "From ���㷽ʽ A, ҽ�ƿ���� B" & vbNewLine & _
'                        "Where b.���� = [1] And a.���� = b.���㷽ʽ And Rownum < 2" & vbNewLine & _
'                        "Union" & vbNewLine & _
'                        "Select a.���� From ���㷽ʽ A, ���ѿ����Ŀ¼ B Where b.���� = [1] And a.���� = b.���㷽ʽ And Rownum < 2"
'    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, cbo���㷽ʽ.Text)
'    If rsTemp.RecordCount <> 0 Then
'        If Val(Nvl(rsTemp!����)) <> 7 And Val(Nvl(rsTemp!����)) <> 8 Then
'            txt����Ӧ��.Text = Format(mcurӦ�� + GetRegistMoney, "0.00")
'        Else
'            txt����Ӧ��.Text = Format(GetRegistMoney(False, True), "0.00")
'        End If
'    Else
    txt����Ӧ��.Text = Format(mcurӦ�� + GetRegistMoney - GetPayedMoney, "0.00")
'    End If
'    cmd�����Һ�.Visible = mint�Һ��� > 0 And mintInsure <> 0     'ҽ�����˲Ż����ӽ���ҺŰ�ť
'    txt�ɿ�.Enabled = Not cmd�����Һ�.Visible
'    txt�Ҳ�.Enabled = Not cmd�����Һ�.Visible
End Sub

Private Function GetPayedMoney() As Double
    '��ȡ�ѽ�����
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
    '����:���ݰ�ť״̬,�ƶ�������Ϣ����ذ�ť
    '����:���˺�
    '����:2010-01-15 10:02:29
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

Private Function IsCheckReservationSameDept(ByVal lng����ID As Long, ByVal strConditions As String, ByVal strԤԼʱ�� As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ����ԤԼ�Һ��Ƿ���ͬһ�������Ѿ�����ԤԼ
    '��Σ�strConditions: ����:����ID=...�����֤��=
    '���Σ�
    '���أ����ڷ���true,���򷵻�False
    '���ƣ����˺�
    '���ڣ�2010-03-17 09:44:11
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim varData As Variant, strWhere As String
    On Error GoTo Hd
    varData = Split(strConditions, "=")
    Select Case varData(0)
    Case "����ID"
            strWhere = " And A.����ID=[2]"
    Case "���֤��"
            strWhere = " And B.���֤��=[3]"
     Case "���￨��"
            strWhere = " And B.���￨��=[3]"
    Case Else
            strWhere = strConditions
    End Select
    strSQL = "" & _
    "   Select  1 " & _
    "   From ������ü�¼  A,������Ϣ B " & _
    "   Where A.����ID=B.����ID And A.��¼����=4 and ��¼״̬=0  " & _
    "               and A.����ʱ�� between [4]  and [5]  and A.���˿���ID+0=[1] " & strWhere
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "���ԤԼ�Һ��Ƿ��Ѿ��ҹ�", lng����ID, Val(varData(1)), CStr(varData(1)), CDate(strԤԼʱ��), CDate(strԤԼʱ��) + 1 - 1 / 24 / 60 / 60)
    IsCheckReservationSameDept = (rsTemp.RecordCount <> 0)
    Exit Function
Hd:
    If ErrCenter() = 1 Then Resume
    SaveErrLog
End Function
Private Sub SetRePrintPatiEnabled(ByVal blnEdit As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�Ƿ������޸Ĳ�����Ϣֵ
    '����:���˺�
    '����:2011-01-31 10:33:04
    '����:35544
    '---------------------------------------------------------------------------------------------------------------------------------------------
    txt����.Locked = Not blnEdit
    cbo���䵥λ.Locked = Not blnEdit
    cbo�Ա�.Locked = Not blnEdit
    SetPatiEnable blnEdit
    txt����.Enabled = blnEdit And Not mTy_Para.bln��ֹ��������
    cbo���䵥λ.Enabled = blnEdit And Not mTy_Para.bln��ֹ��������
    cbo�Ա�.Enabled = blnEdit
    cbo���ʽ.Enabled = Not blnEdit And Not mblnStation    '56263
    cbo��ͥ��ַ.Enabled = Not blnEdit
    cbo���ڵ�ַ.Enabled = Not blnEdit
    padd��ͥ��ַ.Enabled = Not blnEdit: padd��ͥ��ַ.ControlLock = blnEdit
    padd���ڵ�ַ.Enabled = Not blnEdit: padd���ڵ�ַ.ControlLock = blnEdit
    cbo�ѱ�.Enabled = Not blnEdit
    cbo���㷽ʽ.Enabled = Not blnEdit
    txt�����.Enabled = Not blnEdit
    txt��ͥ�绰.Enabled = Not blnEdit
    txtIDCard.Enabled = Not blnEdit
    '74017:���ϴ���2014-6-17���Һ��ش�����༭���ಡ����Ϣ���������
    cmdCard.Enabled = False
End Sub
Public Function zlGet�����() As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�Ƿ����������
    '����:�����,���δ����,�򷵻ؿ�
    '����:���˺�
    '����:2011-02-28 15:27:22
    '����:36028
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mTy_Para.blnԤԼ����������� And mbytMode = 1 Then Exit Function
    If gbln�Զ������ Or mblnStation Or mbln������ Then     'Ҫ����ݲ��������� �ñ�Ҫ���������� ������������ �Ա㽨������
        zlGet����� = zlDatabase.GetNextNo(3)
    Else
        zlGet����� = ""
    End If
End Function

Private Function zlCommitPlugInpati(ByVal strCardNo As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ύ�������
    '����:�ɹ�,����true,����ķ���False
    '����:���˺�
    '����:2011-07-22 14:13:11
    '����:40012
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsPatiInfor As ADODB.Recordset, str���� As String, str�������� As String
    Err = 0: On Error GoTo errHandle
    If CreatePlugInOK(mlngModul) = False Then zlCommitPlugInpati = True: Exit Function
    If mblnNotQuery = False Then zlCommitPlugInpati = True:  Exit Function
    If Not zlInitPati(rsPatiInfor) Then Exit Function
    str���� = Trim(txt����.Text)
    If IsNumeric(str����) Then str���� = str���� & cbo���䵥λ.Text
    With mobjfrmPatiInfo
        If .txt����ʱ�� = "__:__" Then
            str�������� = IIf(IsDate(.txt��������.Text), .txt��������.Text, "")
        Else
            str�������� = IIf(IsDate(.txt��������.Text), "" & .txt��������.Text & " " & .txt����ʱ��.Text & "", "")
        End If
        rsPatiInfor.AddNew
        rsPatiInfor!���� = .txtPatient.Text
        rsPatiInfor!�Ա� = NeedName(cbo�Ա�.Text)
        '89242:���ϴ�,2015/12/7,��ȡ���˵�ַ��Ϣ
        If mblnStructAdress Then
            rsPatiInfor!��ͥ��ַ = IIf(Trim(.padd��ͥ��ַ.Value) = "", padd��ͥ��ַ.Value, .padd��ͥ��ַ.Value)
        Else
            rsPatiInfor!��ͥ��ַ = IIf(Trim(.cbo��ͥ��ַ.Text) = "", cbo��ͥ��ַ.Text, .cbo��ͥ��ַ.Text)
        End If
        rsPatiInfor!�ѱ� = NeedName(cbo�ѱ�.Text)
        rsPatiInfor!���֤�� = Trim(.txt���֤��.Text)
        rsPatiInfor!ҽ�Ƹ��ʽ = NeedName(cbo���ʽ.Text)
        rsPatiInfor!ҽ���� = .txtPatiMCNO(0).Text
        rsPatiInfor!���� = str����
        rsPatiInfor!���� = NeedName(.cbo����.Text)
        rsPatiInfor!���� = NeedName(.cbo����.Text)
        rsPatiInfor!����״�� = NeedName(.cbo����.Text)
        rsPatiInfor!ְҵ = NeedName(.cboְҵ.Text, True)
        rsPatiInfor!�������� = IIf(str�������� <> "", CDate(str��������), Null)
        rsPatiInfor!������λ = .txt��λ����.Text
        rsPatiInfor!��ͬ��λID = Val(.txt��λ����.Tag)
        rsPatiInfor!���� = Trim(.txt����.Text)
        rsPatiInfor!��λ�绰 = Trim(.txt��λ�绰.Text)
        rsPatiInfor!��λ�ʱ� = Trim(.txt��λ�ʱ�.Text)
        rsPatiInfor!��ͥ�绰 = Trim(.txt��ͥ�绰.Text)
        rsPatiInfor!��ͥ�ʱ� = Trim(.txt��ͥ�ʱ�.Text)
        rsPatiInfor.Update
    End With
    
    Err = 0: On Error Resume Next
    'CommitPatiInfo(byVal ����,rsInfo As ADO.RecordSet) As Boolean
    '���뱾�η������ţ��Լ�������Ϣ����������Ϣ��Ϊ��̬��¼�����߱����ֶ���QueryPatiInfo�����صĶ�Ӧ�� _
    '��Ϊ���������⣬�Һų�����Բ��Է���ֵ���ж����ƴ���
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


Public Function zlReadPlugInPati(ByVal str���� As String, Optional blnHavePati As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�彨������Ϣ����
    '���:
    '����:blnHavePati-�Ƿ�ӿڷ�����true,���в�����Ϣ
    '����:���˺�
    '����:2011-06-10 17:50:09
    '����:38841
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsPatiInfor As ADODB.Recordset
    On Error GoTo errHandle
    mblnNotQuery = False
    If CreatePlugInOK(mlngModul) = False Then zlReadPlugInPati = True: Exit Function
    If Not zlInitPati(rsPatiInfor) Then Exit Function
    'QueryPatiInfo(ByVal lngSys As Long, ByVal lngModule As Long, _
    ByVal str���� As String, ByRef rsInfo As ADO.Recordset)
    Err = 0: On Error Resume Next
    If gobjPlugIn.QueryPatiInfo(glngSys, mlngModul, str����, rsPatiInfor) = False Then
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
        txtPatient.Text = Nvl(rsPatiInfor!����) '�����Change�¼�
        cbo�Ա�.ListIndex = cbo.FindIndex(cbo�Ա�, Nvl(rsPatiInfor!�Ա�), True) '�����ں�����ݳ���������
        cbo��ͥ��ַ.Text = Nvl(rsPatiInfor!��ͥ��ַ)
        '89242:���ϴ�,2015/12/7,��ȡ���˵�ַ��Ϣ
        Call zlReadAddrInfo(padd��ͥ��ַ, Val(Nvl(rsPatiInfor!����ID)), 0, 3, cbo��ͥ��ַ.Text)
        Call zlControl.CboSetIndex(cbo�ѱ�.Hwnd, cbo.FindIndex(cbo�ѱ�, "" & rsPatiInfor!�ѱ�, True))
'        txt�����.Text = Nvl(rsPatiInfor!�����, mstr�����)
'        txt�����.Enabled = (Val(txt�����.Text) = 0)
        
        txtIDCard.Text = Nvl(rsPatiInfor!���֤��, txtIDCard.Text) '���֤��:31182
        txtIDCard.Tag = Nvl(rsPatiInfor!���֤��, txtIDCard.Text)  '�Ա㷴�����ٲ�
 
        'ҽ�Ƹ��ʽ
        If Not IsNull(rsPatiInfor!ҽ�Ƹ��ʽ) Then
            cbo���ʽ.ListIndex = cbo.FindIndex(cbo���ʽ, rsPatiInfor!ҽ�Ƹ��ʽ, True)
        ElseIf mstrYBPati <> "" Then
            cbo���ʽ.ListIndex = cbo.FindIndex(cbo���ʽ, "1", True)
        End If
        
        If Not IsNull(rsPatiInfor!ҽ����) And mlngOutModeMC <> 0 Then Call SetCboDefault(cboҽ�����)
        '��ϸ������Ϣ����
        Call CopyInfoTofrmPatiInfo
        .txtPatiMCNO(0).Text = "" & Nvl(rsPatiInfor!ҽ����)
        .txtPatiMCNO(0).Tag = "" & Nvl(rsPatiInfor!ҽ����)
        .txtPatiMCNO(1).Text = .txtPatiMCNO(0).Text
        Call LoadOldData("" & rsPatiInfor!����, .txt����, .cbo���䵥λ)
        .mblnChange = False
        .txt��������.Text = Format(IIf(IsNull(rsPatiInfor!��������), "____-__-__", rsPatiInfor!��������), "YYYY-MM-DD")
        .mblnChange = True
        .txt����.Text = Nvl(rsPatiInfor!����)
        .txt����.Tag = Nvl(rsPatiInfor!����)
        .cbo����.ListIndex = cbo.FindIndex(.cbo����, Nvl(rsPatiInfor!����), True)
        .cbo����.ListIndex = cbo.FindIndex(.cbo����, Nvl(rsPatiInfor!����), True)
        .cbo����.ListIndex = cbo.FindIndex(.cbo����, Nvl(rsPatiInfor!����״��), True)
        .cboְҵ.ListIndex = cbo.FindIndex(.cboְҵ, Nvl(rsPatiInfor!ְҵ))
        .txt���֤��.Text = Nvl(rsPatiInfor!���֤��)
        .txt���֤��.Tag = .txt���֤��.Text
        .txt��λ����.Text = Nvl(rsPatiInfor!������λ)
        .txt��λ����.Tag = Nvl(rsPatiInfor!��ͬ��λID)
        .txt����.Text = Trim(Nvl(rsPatiInfor!����))
        .txt����.Tag = .txt����.Text
        .txt��λ�绰.Text = Nvl(rsPatiInfor!��λ�绰)
        .txt��λ�ʱ�.Text = Nvl(rsPatiInfor!��λ�ʱ�)
        .txt��ͥ�绰.Text = Nvl(rsPatiInfor!��ͥ�绰)
        .txt��ͥ�ʱ�.Text = Nvl(rsPatiInfor!��ͥ�ʱ�)
        If Trim(.txt�����) = "" Then .txt����� = zlGet�����
    End With
    zlReadPlugInPati = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Exit Function
ErrMsg:
    MsgBox "�ӿ�δת�벡����Ϣ,����!", vbInformation + vbOKOnly, gstrSysName
End Function
Private Function zlInitPati(ByRef rsPatiInfor As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ������Ϣ��
    '����:������Ϣ��
    '����:���˺�
    '����:2011-06-10 17:52:18
    '����:38841
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set rsPatiInfor = New ADODB.Recordset
    With rsPatiInfor
        If .State = adStateOpen Then .Close
        '����ID,����,�Ա�,����,��������,�����ص�,���֤��,����֤��,���,ְҵ,��ͥ��ַ,��ͥ�绰,��ͥ�ʱ�,
        '������λ,��λ�ʱ�,ҽ����,ҽ�Ƹ��ʽ,�ѱ�,����,����,����״��,����
        
        .Fields.Append "����ID", adDouble, 18, adFldIsNullable
        .Fields.Append "����", adLongVarChar, zlGetPatiInforMaxLen.intPatiName, adFldIsNullable
        .Fields.Append "�Ա�", adLongVarChar, 4, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 10, adFldIsNullable
        .Fields.Append "��������", adDate, , adFldIsNullable
        .Fields.Append "�����ص�", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "���֤��", adLongVarChar, 18, adFldIsNullable
        .Fields.Append "����֤��", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "���", adLongVarChar, 10, adFldIsNullable
        .Fields.Append "ְҵ", adLongVarChar, 80, adFldIsNullable
        .Fields.Append "��ͥ��ַ", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "��ͥ�绰", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "��ͥ�ʱ�", adLongVarChar, 6, adFldIsNullable
        .Fields.Append "��ͬ��λID", adDouble, 18, adFldIsNullable
        .Fields.Append "������λ", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "��λ�绰", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "��λ�ʱ�", adLongVarChar, 6, adFldIsNullable
        .Fields.Append "ҽ����", adLongVarChar, 30, adFldIsNullable
        .Fields.Append "ҽ�Ƹ��ʽ", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "�ѱ�", adLongVarChar, 10, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 30, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "����״��", adLongVarChar, 4, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 30, adFldIsNullable
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
    '90875:���ϴ�,2016/3/2,ҽ�ƿ�֤������
    IDkindStr = "��|���֤��|0"
    strSQL = "Select ����,ȱʡ��־ from ֤������  Where  ���� Not Like '����%' and ���� Not Like '%���֤'" & vbNewLine & _
            " And Not ���� in (Select ���� from  ҽ�ƿ���� Where Nvl(�Ƿ�֤��,0)=0 or Nvl(�Ƿ�����,0)=0)"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If rsTmp.RecordCount > 0 Then
        Do While Not rsTmp.EOF
            IDkindStr = IDkindStr & ";" & Left(Nvl(rsTmp!����), 1) & "|" & Nvl(rsTmp!����) & "|0"
            rsTmp.MoveNext
        Loop
    End If
    Call IDKind֤��.zlInit(Me, glngSys, mlngModul, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, IDkindStr, Me.txtIDCard)
    'ǿ�ư����֤������Ϊ�ֶ�����
    Set objCard = IDKind֤��.GetIDKindCard("���֤��")
    If Not objCard Is Nothing Then objCard.�Ƿ�Ӵ�ʽ���� = False: IDKind֤��.Refrash
    
    Call IDKind.zlInit(Me, glngSys, mlngModul, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, "", Me.txtPatient)
    If mbytInState = 1 Then Exit Function
    lngCardID = Val(zlDatabase.GetPara("ȱʡҽ�ƿ����", glngSys, mlngModul, 0))
    mblnAlwaysSend = Val(Nvl(zlDatabase.GetPara("���ϸ����ʱʼ�շ���", glngSys, mlngModul, 0), 0)) = 1
    If lngCardID <> 0 Then
        strSQL = "Select Nvl(�Ƿ��ϸ����,0) As ���� From ҽ�ƿ���� Where ID=[1] And Nvl(�Ƿ�����,0)=1"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngCardID)
        If Not rsTmp.EOF Then
            IDKind.DefaultCardType = lngCardID
            mblnSendCard = ((Val(rsTmp!����) = 0) And mblnAlwaysSend)
        End If
    Else
        strSQL = "Select Nvl(�Ƿ��ϸ����,0) As ����,ID From ҽ�ƿ���� Where ȱʡ��־=1 And Nvl(�Ƿ�����,0)=1"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        If Not rsTmp.EOF Then
            IDKind.DefaultCardType = Val(rsTmp!ID)
            mblnSendCard = ((Val(rsTmp!����) = 0) And mblnAlwaysSend)
        End If
    End If
    Set objCard = IDKind.GetfaultCard
    '76824�����ϴ���2014/8/19��ҽ�ƿ���������
    Call InitSendCardPreperty(mlngModul, Val(IDKind.DefaultCardType))
    IDKind.ShowPropertySet = InStr(";" & mstrPrivs & ";", "��������") > 0
    If IDKind.Cards.��ȱʡ������ And Not objCard Is Nothing Then
        gobjSquare.blnȱʡ�������� = objCard.�������Ĺ��� <> ""
        gobjSquare.intȱʡ���ų��� = objCard.���ų���
        Set gobjSquare.objDefaultCard = objCard
       
    Else
        gobjSquare.blnȱʡ�������� = IDKind.Cards.������ʾ
        gobjSquare.intȱʡ���ų��� = 100
    End If
    Exit Function
Errhand:
    If ErrCenter() = 1 Then Resume
End Function

Private Function LoadIdKindStr() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����IDKindStr
    '����:���سɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-09-06 13:36:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strIdKindStr As String, varTemp As Variant, varData As Variant
    Dim i As Long, j As Long, strIDKindTemp As String, strTemp As String
    If gobjSquare.objSquareCard Is Nothing Then Exit Function
    'ȱʡ��Ϊ�������
    If mblnStation And mbytMode = 0 And mTy_Para.bln�Һű���ˢ�� Then
        '38603
        strIdKindStr = gobjSquare.objSquareCard.zlGetIDKindStr("��|��������￨|0")
    Else
        strIdKindStr = gobjSquare.objSquareCard.zlGetIDKindStr(IDKind.IDkindStr)
    End If
    
    If Not (gCurSendCard.lng�����ID = 0 Or gCurSendCard.blnȱʡ��־) Then
        '����|�����|������־|�����ID|���ų���|ȱʡ��־(1-��ǰȱʡ;0-��ȱʡ)|
        '�Ƿ�����ʻ�(1-�����ʻ�;0-�������ʻ�)|��������(�ڼ�λ���ڼ�λ����,��Ϊ������);...
        varData = Split(strIdKindStr, ";")
        strIDKindTemp = ""
        For i = 0 To UBound(varData)
            varTemp = Split(varData(i), "|")
            If Val(varTemp(3)) = gCurSendCard.lng�����ID Then
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
                '����Ƿ�ȱʡ��־
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
    
    'ȡȱʡ��ˢ����ʽ
    '����|�����|ˢ����־|�����ID|���ų���|ȱʡ��־(1-��ǰȱʡ;0-��ȱʡ)|
    '�Ƿ�����ʻ�(1-�����ʻ�;0-�������ʻ�)|��������(�ڼ�λ���ڼ�λ����,��Ϊ������)
    '��7λ��,��ֻ��������,��Ȼȡ������
    gobjSquare.blnȱʡ�������� = IDKind.GetfaultCard.�������Ĺ��� <> ""
    'gobjSquare.lngȱʡ�����ID = IDKind.GetCurCard.�ӿ����
    If mintIDKind > 0 And mintIDKind <= IDKind.ListCount Then IDKind.IDKind = mintIDKind
End Function
Private Sub InitCardSquareData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������رս��㿨����
    '���:blnClosed:�رն���
    '����:���˺�
    '����:2010-01-05 14:51:23
    '����:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strExpend As String
    '1-����,2-���ĳ���ԤԼ����,3-��ѯ����������
     
    Call InitIDKindData
End Sub
Private Sub Load֧����ʽ()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������Ч��֧����ʽ
    '����:���˺�
    '����:2011-07-08 11:41:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim j As Long, strPayType As String, varData As Variant, varTemp As Variant, i As Long
    Dim rsTemp As ADODB.Recordset, blnFind As Boolean
    Dim strSQL As String, str���� As String
    
    If mbln�����Һ� Then Exit Sub
    
    If mblnStation Then
        str���� = "(3,7,8)"
    Else
        str���� = "(1,2,3,7,8)"
    End If
    strSQL = _
        "Select B.����,B.����,Nvl(B.����,1) as ����,Nvl(A.ȱʡ��־,0) as ȱʡ" & _
        " From ���㷽ʽӦ�� A,���㷽ʽ B" & _
        " Where A.Ӧ�ó���=[1] And B.����=A.���㷽ʽ and B.���� In " & str���� & _
        " Order by B.����"
        
    Err = 0: On Error GoTo errH
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, "�Һ�")
    
    Set mcolCardPayMode = New Collection
    '��|ȫ��|ˢ����־|�����ID(���ѿ����)|����|�Ƿ����ѿ�|���㷽ʽ|�Ƿ�����|�Ƿ����ƿ�;��
    If Not gobjSquare.objSquareCard Is Nothing Then
        strPayType = gobjSquare.objSquareCard.zlGetAvailabilityCardType
    End If
    varData = Split(strPayType, ";")
    
    With cbo���㷽ʽ
        .Clear: j = 0
        Do While Not rsTemp.EOF
            blnFind = False
            For i = 0 To UBound(varData)
                varTemp = Split(varData(i) & "|||||", "|")
                If varTemp(6) = Nvl(rsTemp!����) Then
                    blnFind = True
                    Exit For
                End If
            Next
            If rsTemp!���� = 3 Then mstr�����ʻ� = rsTemp!����: blnFind = True  '�����:57711
            If rsTemp!���� = 7 Or rsTemp!���� = 8 Then blnFind = True
                         
            If Not blnFind Then
                .AddItem Nvl(rsTemp!����) & "-" & Nvl(rsTemp!����)
                mcolCardPayMode.Add Array("", Nvl(rsTemp!����), 0, 0, 0, 0, Nvl(rsTemp!����), 0, 0), "K" & j
                If Nvl(rsTemp!����) & "-" & Nvl(rsTemp!����) = gstr���㷽ʽ Then
                     .ItemData(.NewIndex) = 1
                     .ListIndex = .NewIndex
                End If
                If Val(Nvl(rsTemp!ȱʡ)) = 1 Then
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
                If Split(varData(i) & "|||||", "|")(6) = Nvl(rsTemp!����) Then
                    blnFind = True
                    Exit Do
                End If
                rsTemp.MoveNext
            Loop
            If InStr(1, varData(i), "|") <> 0 And blnFind Then
                varTemp = Split(varData(i), "|")
                mcolCardPayMode.Add varTemp, "K" & j
                .AddItem varTemp(1): .ItemData(.NewIndex) = -1
                '74552,Ƚ����,2014-7-2,�ҺŹ���������Ĭ�Ͻ��㷽ʽʱ�����ѡ����㷽ʽ����Ϊ"7-һ��ͨ����"�Ľ��㷽ʽ
                If varTemp(1) = gstr���㷽ʽ Then .ListIndex = .NewIndex
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
    '����:���ˢ��
    '����:�Ϸ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-18 14:52:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsMoney As ADODB.Recordset, str���� As String
    Dim strXmlIn As String, lng����ID As Long
    Dim objCard As Card
    
    On Error GoTo errHandle
    '68991
    If mRegistFeeMode = EM_RG_���� Then CheckBrushCard = True: Exit Function
    If dblMoney = 0 Then
        '����:51527
        CheckBrushCard = True: Exit Function
    End If
    
    If mCurCardPay.lngҽ�ƿ����ID = 0 Then
        MsgBox cbo���㷽ʽ.Text & "�쳣,����!", vbOKOnly + vbInformation, gstrSysName
        Exit Function
    End If
    If mstrYBPati <> "" Then
        MsgBox "��֧��ҽ������ʹ��" & mCurCardPay.str���� & "֧����", vbInformation, gstrSysName
        Exit Function
    End If
    If gobjSquare.objSquareCard Is Nothing Then
        MsgBox "ʹ��" & mCurCardPay.str���� & "֧�������ȳ�ʼ���ӿڲ�����", vbInformation, gstrSysName
        Exit Function
    End If
    
    If mCurCardPay.bln���ѿ� = False Then '������֧�������в���֧��
        If gobjSquare.objSquareCard.zlGetCard(mCurCardPay.lngҽ�ƿ����ID, False, objCard) = False Then Exit Function
        If objCard Is Nothing Then
            MsgBox "ʹ��" & mCurCardPay.str���� & "֧�������ȳ�ʼ���ӿڲ�����", vbInformation, gstrSysName
            Exit Function
        End If
        If objCard.�ӿڳ����� = "" Then
            MsgBox "ʹ��" & mCurCardPay.str���� & "֧�������ȳ�ʼ���ӿڲ�����", vbInformation, gstrSysName
            Exit Function
        End If
        Set mCurCardPay.objCard = objCard
    End If
    
    Call zlGetClassMoney(rsMoney)
    
     '����ˢ������
    'zlBrushCard(frmMain As Object, _
    'ByVal lngModule As Long, _
    'ByVal rsClassMoney As ADODB.Recordset, _
    'ByVal lngCardTypeID As Long, _
    'ByVal bln���ѿ� As Boolean, _
    'ByVal strPatiName As String, ByVal strSex As String, _
    'ByVal strOld As String, ByVal dbl��� As Double, _
    'Optional ByRef strCardNo As String, _
    'Optional ByRef strPassWord As String, _
    Optional ByRef bln�˷� As Boolean = False, _
    Optional ByRef blnShowPatiInfor As Boolean = False, _
    Optional ByRef bln���� As Boolean = False, _
    Optional ByVal bln�����ֹ As Boolean = True, _
    Optional ByRef varSquareBalance As Variant, _
    Optional ByVal blnתԤ�� As Boolean = False, _
    Optional ByVal blnAllPay As Boolean = False, _
    Optional ByVal strXmlIn As String = "", _
    Optional ByVal str������Դ As String, _
    Optional ByVal lng����ID As Long) As Boolean
    str���� = Trim(txt����.Text)
    If IsNumeric(str����) Then str���� = str���� & cbo���䵥λ.Text
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ָ��֧�����,����ˢ������
    '���:rsClassMoney:�շ����,���
    '        lngCardTypeID-Ϊ��ʱ,Ϊ��һ��ͨˢ��
    '       bln�����ֹ-Ŀǰֻ������ѿ�,��ʾ����ʱ,��ֹ��������,������������֧��
    '       str������Դ - ��ǰ֧�����õķ�����Դ�������ö��ŷָ�(ʹ�����ѿ�֧��ʱ����)
    '       lng����ID - ����ID(ʹ�����ѿ�֧��ʱ����)
   '58322
   strXmlIn = "<IN><CZLX>0</CZLX></IN>"
   If Not mrsInfo Is Nothing Then lng����ID = Val(Nvl(mrsInfo!����ID))
   If gobjSquare.objSquareCard.zlBrushCard(Me, mlngModul, rsMoney, mCurCardPay.lngҽ�ƿ����ID, mCurCardPay.bln���ѿ�, _
    txtPatient.Text, NeedName(cbo�Ա�.Text), str����, dblMoney, mCurCardPay.strˢ������, mCurCardPay.strˢ������, _
    False, True, False, True, Nothing, False, True, strXmlIn, "1", lng����ID) = False Then Exit Function
    '����ǰ,һЩ���ݼ��
    'zlPaymentCheck(frmMain As Object, ByVal lngModule As Long, _
    ByVal strCardTypeID As Long, ByVal strCardNo As String, _
    ByVal dblMoney As Double, ByVal strNOs As String, _
    Optional ByVal strXMLExpend As String
    If gobjSquare.objSquareCard.zlPaymentCheck(Me, mlngModul, mCurCardPay.lngҽ�ƿ����ID, _
        mCurCardPay.bln���ѿ�, mCurCardPay.strˢ������, dblMoney, "", "") = False Then Exit Function
        '����
''    zlGetAccountMoney(ByVal frmMain As Object, ByVal lngModule As Long, _
''    ByVal strCardTypeID As Long, _
''    ByVal strCardNo As String, strExpand As String, dblMoney As Double
'    '���:frmMain-���õ�������
'    '        lngModule-ģ���
'    '        strCardNo-����
'    '        strExpand-Ԥ����Ϊ��,�Ժ���չ
'    '����:dblMoney-�����ʻ����
'    Dim strExpand As String, dbl�ʻ���� As Double
'    If gobjSquare.objSquareCard.zlGetAccountMoney(Me, mlngModule, mCurCardPay.lngҽ�ƿ����ID, _
'          mCurCardPay.strˢ������, strExpand, dbl�ʻ����, mCurCardPay.bln���ѿ�) = False Then Exit Function
'    stbThis.Panels(4).Text = Format(dbl�ʻ����, "0.00")
'    stbThis.Panels(4).ToolTipText = mCurCardPay.str���㷽ʽ & "���ʻ����:" & Format(dbl�ʻ����, "0.00")
'    mCurCardPay.dbl�ʻ���� = Round(dbl�ʻ����, 2)
    CheckBrushCard = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function zlInterfacePrayMoney(ByVal lngCard����ID As Long, ByVal lng�ҺŽ���ID As Long, ByRef cllPro As Collection, _
    ByRef cllThreeSwap As Collection, dblMoney As Double) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ӿ�֧�����
    '����:cllPro-�޸�������������
    '        cll��������-����������������
    '����:֧���ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-17 13:34:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSwapGlideNO As String, strSwapMemo As String, strSwapExtendInfor As String
    Dim strBalanceIDs As String
    
    If mCurCardPay.lngҽ�ƿ����ID = 0 Or dblMoney = 0 Then zlInterfacePrayMoney = True: Exit Function
    
    If mCurCardPay.Have�Һŷ� Then strBalanceIDs = lng�ҺŽ���ID
    If mCurCardPay.Have���� Then strBalanceIDs = strBalanceIDs & IIf(strBalanceIDs = "", "", ",") & lngCard����ID
    
'    If cbo���㷽ʽ.ItemData(cbo���㷽ʽ.ListIndex) <> -1 Then zlInterfacePrayMoney = True: Exit Function
    'zlPaymentMoney(ByVal frmMain As Object, _
    ByVal lngModule As Long, ByVal lngCardTypeID As Long, _
    ByVal bln���ѿ� As Boolean, _
    ByVal strCardNo As String, ByVal strBalanceIDs As String, _
    byval  strPrepayNos as string , _
    ByVal dblMoney As Double, _
    ByRef strSwapGlideNO As String, _
    ByRef strSwapMemo As String, _
    Optional ByRef strSwapExtendInfor As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ʻ��ۿ��
    '���:frmMain-���õ�������
    '        lngModule-����ģ���
    '        strBalanceIDs-����ID,����ö��ŷ���
    '        strPrepayNos-��Ԥ��ʱ��Ч. Ԥ�����ݺ�,����ö��ŷ���
    '       strCardNo-����
    '       dblMoney-֧�����
    '����:strSwapGlideNO-������ˮ��
    '       strSwapMemo-����˵��
    '       strSwapExtendInfor-������չ��Ϣ: ��ʽΪ:��Ŀ����1|��Ŀ����2||��||��Ŀ����n|��Ŀ����n
    '����:�ۿ�ɹ�,����true,���򷵻�Flase
    '˵��:
    '   ��������Ҫ�ۿ�ĵط����øýӿ�,Ŀǰ�滮��:�շ��ң��Һ���;������ѯ��;ҽ������վ��ҩ���ȡ�
    '   һ����˵���ɹ��ۿ�󣬶�Ӧ�ô�ӡ��صĽ���Ʊ�ݣ����Է��ڴ˽ӿڽ��д���.
    '   �ڿۿ�ɹ��󣬷��ؽ�����ˮ�ź���ر�ע˵���������������������Ϣ�����Է��ڽ���˵�����Ա��˷�.
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If gobjSquare.objSquareCard.zlPaymentMoney(Me, mlngModul, mCurCardPay.lngҽ�ƿ����ID, mCurCardPay.bln���ѿ�, mCurCardPay.strˢ������, strBalanceIDs, "", dblMoney, strSwapGlideNO, strSwapMemo, strSwapExtendInfor) = False Then Exit Function
    '����������������
     If lng�ҺŽ���ID <> 0 Then
        '����:58322
        'mbytMode As Integer '0-�Һ�,1-ԤԼ,2-����,3-ȡ��ԤԼ ,4-�˺� ԤԼ������ģʽ:0-�Һ�,��ʱԤԼҪ�շ�,1-ԤԼ,���շ�
        If Not ((mbytMode = 0 Or mbytMode = 2) And mCurCardPay.bln���ѿ�) Then
            '���ѿ��Ѿ��ڲ���Һż�¼ʱ,�Ѿ��ۿ�
            Call zlAddUpdateSwapSQL(False, lng�ҺŽ���ID, mCurCardPay.lngҽ�ƿ����ID, mCurCardPay.bln���ѿ�, mCurCardPay.strˢ������, strSwapGlideNO, strSwapMemo, cllPro)
        End If
        
        Call zlAddThreeSwapSQLToCollection(False, lng�ҺŽ���ID, mCurCardPay.lngҽ�ƿ����ID, mCurCardPay.bln���ѿ�, mCurCardPay.strˢ������, strSwapExtendInfor, cllThreeSwap)
    End If
    If lngCard����ID <> 0 Then
        If Not ((mbytMode = 0 Or mbytMode = 2) And mCurCardPay.bln���ѿ�) Then
                '���ѿ��Ѿ��ڷ�����¼ʱ,�Ѿ��ۿ�
                Call zlAddUpdateSwapSQL(False, lngCard����ID, mCurCardPay.lngҽ�ƿ����ID, mCurCardPay.bln���ѿ�, mCurCardPay.strˢ������, strSwapGlideNO, strSwapMemo, cllPro)
        End If
        Call zlAddThreeSwapSQLToCollection(False, lngCard����ID, mCurCardPay.lngҽ�ƿ����ID, mCurCardPay.bln���ѿ�, mCurCardPay.strˢ������, strSwapExtendInfor, cllThreeSwap)
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
    '����:����ʱ,��ʼ��֧�����(�շ����,ʵ�ս��)
    '����:���سɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-06-10 17:52:18
    '����:38841
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strSQL  As String
    
    Err = 0: On Error GoTo Errhand:
    
    Set rsMoney = New ADODB.Recordset
    With rsMoney
        If .State = adStateOpen Then .Close
        .Fields.Append "�շ����", adLongVarChar, 20, adFldIsNullable
        '58322
        .Fields.Append "���", adDouble, 18, adFldIsNullable
        .CursorLocation = adUseClient
        .Open , , adOpenStatic, adLockOptimistic
        If (mbytMode = 2 Or mbytMode = 0 And mbytInState = 0 And mstrNoIn <> "") And Not mTy_Para.blnԤԼ����ȷ���Һŷ� Then      'ԤԼ����
            strSQL = "Select �շ����,sum(nvl(ʵ�ս��,0)) as ʵ�� from ������ü�¼ where NO=[1] and ��¼����=4 And ��¼״̬=0  Group by �շ����"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrNoIn)
            Do While Not rsTemp.EOF
                 .AddNew
                !�շ���� = Nvl(rsTemp!�շ����, "��")
                !��� = Val(Nvl(rsTemp!ʵ��))
                .Update
                rsTemp.MoveNext
            Loop
              '����ԤԼ����ʱ,�����շѵ�״��(�ǽ���ʱȷ���Һŷ�) 60171
            If Not mrsItems Is Nothing Then
                mrsItems.Filter = "����=4"    '����
                If mrsItems.RecordCount > 0 Then
                    Do While Not mrsItems.EOF
                        mrsInComes.Filter = "��ĿID=" & mrsItems!��ĿID
                        rsMoney.Filter = "�շ����='" & Nvl(mrsItems!���, "��") & "'"
                        If rsMoney.EOF Then
                            .AddNew
                        Else
                            rsMoney.Filter = 0
                        End If
                        !�շ���� = Nvl(mrsItems!���, "��")
                        Do While Not mrsInComes.EOF
                            !��� = Val(Nvl(!���)) + Val(Nvl(mrsInComes!ʵ��))
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
            mrsInComes.Filter = "��ĿID=" & mrsItems!��ĿID
            rsMoney.Filter = "�շ����='" & Nvl(mrsItems!���, "��") & "'"
            If rsMoney.EOF Then
                .AddNew
            Else
                rsMoney.Filter = 0
            End If
            !�շ���� = Nvl(mrsItems!���, "��")
            Do While Not mrsInComes.EOF
                !��� = Val(Nvl(!���)) + Val(Nvl(mrsInComes!ʵ��))
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

Private Sub AddCardDataSQL(ByVal lng����ID As Long, ByVal dtCurdate As Date, _
    ByRef cllPro As Collection, ByRef lngCard����ID As Long, Optional ByVal bln���� As Boolean, _
    Optional ByVal lng��Ŀid As Long = 0)

    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���￨���Ŵ���
    '���:lng����ID
    '       int����-�����Ƿ���ü��ʷ�ʽ
    '����:lngCard����ID-���ѵĽ���ID
    '����:���˺�
    '����:2011-07-07 04:36:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim byt�������� As Byte, strNO As String, strPassWord As String, strSQL As String
    Dim strԭ���� As String, str���� As String, strCard As String, str�䶯ԭ�� As String
    Dim strICCard As String, lngBrushCardTypeID As Long, str���㷽ʽ As String, strBrushCardNo As String
    Dim bln���ѿ� As Boolean, blnInRange As Boolean   '��Χ�ڵĿ�
    Dim lngIndex As Long, byt�䶯���� As Byte, lng����ID As Long
    Dim str����  As String, strYLKNo As String
    Dim lngPay�����id As Long, blnPay���ѿ� As Boolean, strPayCardNo As String
    On Error GoTo errHandle
    str���� = Trim(mobjfrmPatiInfo.txt����.Text)
    strCard = UCase(mobjfrmPatiInfo.txt����.Text): strICCard = IIf(mblnICCard, strCard, "")
    If Not ((strCard <> "" Or strICCard <> "")) Then Exit Sub
    
    lng����ID = 0: blnInRange = True
    '115168:���ϴ���2017/12/13�����淢����ҽ�ƿ�����
    If mCurSendCard.lng�����ID = 0 Then mCurSendCard = gCurSendCard
    If mCurSendCard.blnOneCard And mCurSendCard.bln�ϸ���� Then blnInRange = mlng�ſ�����ID > 0
    '77805
    If mrsItems Is Nothing Then
        blnInRange = False
    Else
        If lng��Ŀid = 0 Then
            mrsItems.Filter = "����=4"
            blnInRange = mrsItems.RecordCount <> 0
            If mrsItems.RecordCount > 0 Then
                mrsInComes.Filter = "��ĿID=" & mrsItems!��ĿID
            End If
        Else
            blnInRange = True
            mrsInComes.Filter = "��ĿID=" & lng��Ŀid
        End If
    End If
    'Ժ�⿨�Ҳ��ܷ�����,ֻ���ǰ󶨿�
    If bln����(strCard) = False Then
        blnInRange = False
    Else
        blnInRange = True
    End If
    If blnInRange Then
        blnInRange = True
        byt�������� = 0: byt�䶯���� = 1
    Else
        blnInRange = False
        byt�䶯���� = 11: byt�������� = 0
    End If
    
    str���� = Trim(txt����.Text)
    If IsNumeric(str����) Then str���� = str���� & cbo���䵥λ.Text
    
    str�䶯ԭ�� = "���˹Һŷ���"
    
    strPassWord = zlCommFun.zlStringEncode(str����)
    If blnInRange = False Then
          'Zl_ҽ�ƿ��䶯_Insert
           strSQL = "Zl_ҽ�ƿ��䶯_Insert("
          '      �䶯����_In   Number,
          '��������=1-����(��11�󶨿�);2-����;3-����(13-����ͣ��);4-�˿�(��14ȡ����); ��-�������(ֻ��¼);6-��ʧ(16ȡ����ʧ)
          strSQL = strSQL & "" & byt�䶯���� & ","
          '      ����id_In     סԺ���ü�¼.����id%Type,
          strSQL = strSQL & "" & lng����ID & ","
          '      �����id_In   ����ҽ�ƿ���Ϣ.�����id%Type,
          strSQL = strSQL & "" & mCurSendCard.lng�����ID & ","
          '      ԭ����_In     ����ҽ�ƿ���Ϣ.����%Type,
          strSQL = strSQL & "'" & strԭ���� & "',"
          '      ҽ�ƿ���_In   ����ҽ�ƿ���Ϣ.����%Type,
          strSQL = strSQL & "'" & strCard & "',"
          '      �䶯ԭ��_In   ����ҽ�ƿ��䶯.�䶯ԭ��%Type,
          '      --�䶯ԭ��_In:�������������䶯ԭ��Ϊ����.���ܵ�
          strSQL = strSQL & "'" & str�䶯ԭ�� & "',"
          '      ����_In       ������Ϣ.����֤��%Type,
          strSQL = strSQL & "'" & strPassWord & "',"
          '      ����Ա����_In סԺ���ü�¼.����Ա����%Type,
          strSQL = strSQL & "'" & UserInfo.���� & "',"
          '      �䶯ʱ��_In   סԺ���ü�¼.�Ǽ�ʱ��%Type,
          strSQL = strSQL & "to_date('" & Format(dtCurdate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
          '      Ic����_In     ������Ϣ.Ic����%Type := Null,
          strSQL = strSQL & "'" & strICCard & "',"
          '      ��ʧ��ʽ_In   ����ҽ�ƿ��䶯.��ʧ��ʽ%Type := Null
          strSQL = strSQL & "NULL)"
          lngCard����ID = 0
          zlAddArray cllPro, strSQL
    Else
        If gbln���ѽ����� Or CheckIsPrice Then  '�Һ��ǻ��۵�����ô��Ӧ��Ҳ�ǻ��۵�
            strNO = zlDatabase.GetNextNo(13)
            strYLKNo = zlDatabase.GetNextNo(16)  'ҽ�ƿ�
            strSQL = "zl_���ﻮ�ۼ�¼_Insert('" & strNO & "',1," & lng����ID & ",NULL," & IIf(txt�����.Text = "", Null, txt�����.Text) & "," & _
                      "'" & NeedCode(cbo���ʽ) & "','" & txtPatient.Text & "','" & NeedName(cbo�Ա�.Text) & "','" & txt����.Text & cbo���䵥λ.Text & "'," & _
                      "'" & NeedName(cbo�ѱ�.Text) & "',0," & UserInfo.����ID & "," & _
                      UserInfo.����ID & ",'" & UserInfo.���� & "',NULL," & gCurSendCard.rs����!�շ�ϸĿID & "," & _
                      "'" & gCurSendCard.rs����!�շ���� & "','" & gCurSendCard.rs����!���㵥λ & "',NULL,1,1,0," & mlng�Һſ���ID & ",NULL," & _
                      gCurSendCard.rs����!������ĿID & ",'" & gCurSendCard.rs����!�վݷ�Ŀ & "'," & Format(gCurSendCard.rs����!�ּ�, "0.000") & "," & _
                      Format(gCurSendCard.rs����!�ּ�, "0.00") & "," & Format(gCurSendCard.rs����!�ּ�, "0.00") & "," & "to_date('" & Format(dtCurdate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss')" & "," & _
                      "to_date('" & Format(dtCurdate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss')" & ",NULL,'" & UserInfo.���� & "','" & strYLKNo & "')"
            zlAddArray cllPro, strSQL
            
            '���ڿ�����Ҫ����סԺ���ü�¼
            strSQL = zlGetSaveCardFeeSQL(mCurSendCard.lng�����ID, 0, strYLKNo, lng����ID, 0, UserInfo.����ID, mlng�Һſ���ID, 0, _
            zlStr.NeedName(cbo�ѱ�.Text), "", Trim(txtPatient.Text), zlStr.NeedName(cbo�Ա�.Text), str����, _
            strCard, strPassWord, "�Һŷ���", 0, 0, "", dtCurdate, mlng�ſ�����ID, gCurSendCard.rs����, _
            strICCard, mCurCardPay.lngҽ�ƿ����ID, mCurCardPay.bln���ѿ�, mCurCardPay.strˢ������, , strNO)
            zlAddArray cllPro, strSQL
        Else
            strNO = zlDatabase.GetNextNo(16)  'ҽ�ƿ�
            lng����ID = zlDatabase.GetNextId("���˽��ʼ�¼")
            '���㷽ʽΪ��ʱΪ���ʷ�ʽ
            '68991
            '137473:���ϴ���2019/1/24��������֧������ʱ����д֧������
            If Not bln���� Then
                str���㷽ʽ = mstrCard���㷽ʽ
                If mCurCardPay.Have���� Then
                    lngPay�����id = mCurCardPay.lngҽ�ƿ����ID
                    blnPay���ѿ� = mCurCardPay.bln���ѿ�
                    strPayCardNo = mCurCardPay.strˢ������
                End If
            End If
            strSQL = zlGetSaveCardFeeSQL(mCurSendCard.lng�����ID, byt��������, strNO, lng����ID, 0, 0, mlng�Һſ���ID, 0, _
             NeedName(cbo�ѱ�.Text), "", Trim(txtPatient.Text), NeedName(cbo�Ա�.Text), str����, _
            strCard, strPassWord, str�䶯ԭ��, IIf(mCurSendCard.bln��� = False, mCurSendCard.dblӦ�ս��, mCurSendCard.dblʵ�ս��), mCurSendCard.dblʵ�ս��, str���㷽ʽ, _
            dtCurdate, mlng�ſ�����ID, gCurSendCard.rs����, strICCard, lngPay�����id, blnPay���ѿ�, strPayCardNo, lng����ID)
            zlAddArray cllPro, strSQL
        End If
    End If
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
 End Sub
 
Private Sub AddCardTypeSQL(ByVal intOper As Integer, ByVal lng�����ID As Long, ByVal strCode As String, ByVal strȫ�� As String, ByVal str���� As String, _
                           ByVal lng���ų��� As Long, ByRef colPro As Collection)
    Dim strSQL As String
    ' Zl_ҽ�ƿ����_Update
        strSQL = "Zl_ҽ�ƿ����_Update("
        '  Id_In           In ҽ�ƿ����.ID%Type,
        strSQL = strSQL & "" & lng�����ID & ","
        '  ����_In         In ҽ�ƿ����.����%Type,
        strSQL = strSQL & "'" & strCode & "',"
        '  ����_In         In ҽ�ƿ����.����%Type,
        strSQL = strSQL & "'" & strȫ�� & "',"
        '  ����_In         In ҽ�ƿ����.����%Type,
        strSQL = strSQL & "'" & str���� & "',"
        '  ǰ׺�ı�_In     In ҽ�ƿ����.ǰ׺�ı�%Type,
        strSQL = strSQL & "'" & "" & "',"
        '  ���ų���_In     In ҽ�ƿ����.���ų���%Type,
        strSQL = strSQL & "" & lng���ų��� & ","
        '  ȱʡ��־_In     In ҽ�ƿ����.ȱʡ��־%Type,
        strSQL = strSQL & "" & 0 & ","
        '  �Ƿ�̶�_In     In ҽ�ƿ����.�Ƿ�̶�%Type,
        strSQL = strSQL & "1,"
        '  �Ƿ��ϸ����_In In ҽ�ƿ����.�Ƿ��ϸ����%Type,
        strSQL = strSQL & "" & 0 & ","
        '  �Ƿ�����_In     In ҽ�ƿ����.�Ƿ�����%Type,
        strSQL = strSQL & "" & 0 & ","
        '  �Ƿ�����ʻ�_In In ҽ�ƿ����.�Ƿ�����ʻ�%Type,
        strSQL = strSQL & "" & 0 & ","
        '  �Ƿ�ȫ��_In     In ҽ�ƿ����.�Ƿ�ȫ��%Type,
        strSQL = strSQL & "0,"
        '  ����_In         In ҽ�ƿ����.����%Type,
        strSQL = strSQL & "'" & "" & "',"
        '  ��ע_In         In ҽ�ƿ����.��ע%Type,
        strSQL = strSQL & "'" & "" & "',"
        '  �ض���Ŀ_In     In ҽ�ƿ����.�ض���Ŀ%Type,
        strSQL = strSQL & "'" & strCode & "',"
        '    �շ�ϸĿid_In   In �շ���ĿĿ¼.ID%Type,
        strSQL = strSQL & "" & "0" & ","
        '  ���㷽ʽ_In     In ҽ�ƿ����.���㷽ʽ%Type,
        strSQL = strSQL & "'" & "" & "',"
        '  �Ƿ�����_In     In ҽ�ƿ����.�Ƿ�����%Type,
        strSQL = strSQL & "1,"
        '  ��������_In     In ҽ�ƿ����.��������%Type,
        strSQL = strSQL & "" & "NULL" & ","
        '  �Ƿ��ظ�ʹ��_In In ҽ�ƿ����.�Ƿ��ظ�ʹ��%Type,
        strSQL = strSQL & "" & 1 & ","
        '���볤��_In     In ҽ�ƿ����.���볤��%Type,
        strSQL = strSQL & "" & 10 & ","
        '���볤������_In In ҽ�ƿ����.���볤������%Type,
        strSQL = strSQL & "" & 0 & ","
        '�������_In     In ҽ�ƿ����.�������%Type,
        strSQL = strSQL & "" & 0 & ","
        strSQL = strSQL & "" & 1 & ","
        '  ������ʽ_In     In Integer := 0
        strSQL = strSQL & "" & intOper & ","
        '�Ƿ�ģ������_In     In ҽ�ƿ����.�Ƿ�ģ������%Type:=0
        strSQL = strSQL & "" & 0 & ","
        '�����:51072
        '������������_In     In ҽ�ƿ����.������������%Type:=0
        strSQL = strSQL & "" & 0 & ","
        '�Ƿ�ȱʡ����_In     In ҽ�ƿ����.�Ƿ�ȱʡ����%Type:=0
        strSQL = strSQL & "" & 0 & ","
        '�����:56508
        '�Ƿ��ƿ�_In
        strSQL = strSQL & "" & 0 & ","
        '�Ƿ񷢿�_In
        strSQL = strSQL & "" & 0 & ","
        '�Ƿ�д��_In
        strSQL = strSQL & "" & 0 & ","
        '�����:57697
        '����_In
        strSQL = strSQL & "" & 0 & ","
        '�����:57326
        strSQL = strSQL & "" & 1 & ","
        '77872,���ϴ�,2014/12/3:�Ƿ�֧��ת�ʼ�����
        '�Ƿ�ת�ʼ�����_In  In ҽ�ƿ����.�Ƿ�ת�ʼ�����%Type:=0
        strSQL = strSQL & "" & 0 & ","
        '��������_In       In ҽ�ƿ����.��������%Type := '1000',
        strSQL = strSQL & "" & "1000" & ","
        '���̿��Ʒ�ʽ_In   In ҽ�ƿ����.���̿��Ʒ�ʽ%Type := 0,
        strSQL = strSQL & "" & 0 & ","
        '90875:���ϴ�,2015/12/16,����ҽ�ƿ�֤������
        '�Ƿ�֤��_In  In ҽ�ƿ����.�Ƿ�֤��%Type:=0
        strSQL = strSQL & "" & 1 & ")"
        
        zlAddArray colPro, strSQL
End Sub
 
Private Function IsCheckCancelValied(ByVal lng�ҺŽ���ID As Long, ByVal lng���ѽ���ID As Long, _
    ByVal cllBillBalance As Collection, ByVal dbl��� As Double, Optional ByVal bln�˿��鿨 As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����˷�ʱ��������Ч��
    '����:��Ч,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-12 18:43:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strName As String, bln���ѿ� As Boolean, lng�����ID As Long
    Dim str��֤����  As String, strXmlIn As String, strˢ������ As String
    Dim str���� As String, str������ˮ�� As String, str����˵�� As String, str������Ϣ As String
    Dim strXMLExpend As String
    Dim cllSquareBalance As Collection
    
    On Error GoTo errHandle
    strName = IIf(glngSys \ 100 = 8, "��Ա��", "ҽ�ƿ�")
    If cllBillBalance Is Nothing Then IsCheckCancelValied = True: Exit Function
    '�����:58567
    'Array(�����ID,����,�Ƿ����ѿ�(1-��;0-��),������ˮ��,����˵��,strNO,����ID,���ѿ�ID)
    lng�����ID = cllBillBalance(1)(0)
    If lng�����ID = 0 Then IsCheckCancelValied = True: Exit Function
    
    str���� = cllBillBalance(1)(1)
    bln���ѿ� = Val(cllBillBalance(1)(2)) = 1
    str������ˮ�� = cllBillBalance(1)(3)
    str����˵�� = cllBillBalance(1)(4)
    
    Set cllSquareBalance = New Collection
    'Array(�����ID,���ѿ�ID,ˢ�����, ����,����,�������,�Ƿ�����,ʣ��δ�˽��)
    cllSquareBalance.Add Array(lng�����ID, cllBillBalance(1)(7), 0, str����, "", "", False, dbl���)
    
    If gobjSquare Is Nothing Then
        Call InitCardSquareData
    End If
    '4.3.3.2.6   zlReturnCheck:�ʻ����˽���ǰ�ļ��
    'zlPaymentCheck�ʻ��ۿ�׼��
    '������  ��������    ��/��   ��ע
    'frmMain Object  In  ���õ�������
    'lngModule   Long    In  ģ���
    'lngCardTypeID   Long    In  �����ID:ҽ�ƿ����.ID
    'strCardNo   String  IN  ����
    'strBalanceIDs:��ʽ:�շ�����( 1-Ԥ����,2-����,3-�շ�,4-�Һ�,5-ҽ�ƿ��տ�)|ID1,ID2��IDn||�շ�����n|ID1,ID2��IDn
    'dblMoney    Double  IN  �˿���
    'strSwapNo   String  In  ������ˮ��(�˿�ʱ���)
    'strSwapMemo String  In  ����˵��(�˿�ʱ����)
    '    Boolean ��������    True:���óɹ�,False:����ʧ��
    '˵��:
    '�ڵ��ÿۿ�ǰ�����ڴ���Oracle�������⣬��ˣ��ٵ��û��˽���ǰ���Ƚ������ݵĺϷ��Լ��,�Ա�������������
    If lng���ѽ���ID <> 0 Then str������Ϣ = str������Ϣ & "||5|" & lng���ѽ���ID
    If lng�ҺŽ���ID <> 0 Then str������Ϣ = str������Ϣ & "||4|" & lng�ҺŽ���ID
    If str������Ϣ <> "" Then str������Ϣ = Mid(str������Ϣ, 3)
    
    If gobjSquare.objSquareCard.zlReturnCheck(Me, mlngModul, lng�����ID, bln���ѿ�, str����, str������Ϣ, dbl���, str������ˮ��, str����˵��, strXMLExpend) = False Then
        Exit Function
    End If
    
    If bln���ѿ� And gbln���ѿ��˷��鿨 _
        Or bln���ѿ� = False And bln�˿��鿨 Then
        strXmlIn = "<IN><CZLX>2</CZLX></IN>"
        If gobjSquare.objSquareCard.zlBrushCard(Me, mlngModul, Nothing, lng�����ID, bln���ѿ�, _
            txtPatient.Text, NeedName(cbo�Ա�.Text), txt����.Text & (IIf(cbo���䵥λ.Visible, cbo���䵥λ.Text, "")), dbl���, str����, strˢ������, _
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
    ByVal lng�ҺŽ���ID As Long, ByVal lng���ѽ���ID As Long, _
    ByVal dblMoney As Double, _
    ByRef cllUpdate As Collection, _
    ByRef cllThreeSwap As Collection, ByRef strErrMsg As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���û��˽ӿ�
    '���:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-13 10:33:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str���� As String, strSwapGlideNO As String, strSwapMemo As String, str������Ϣ As String, strSwapExtendInfor As String
    Dim varData As Variant, varTemp As Variant, i As Long, cllPro As Collection
    Dim bln���ѿ� As Boolean, lng�����ID As Long, strSQL As String, rsTemp As ADODB.Recordset
    Dim lng�Һų���ID As Long, lng�˿�����ID As Long, strTemp As String
    
    Err = 0: On Error GoTo Errhand:
    '�����ID,����,�Ƿ����ѿ�(1-��;0-��),������ˮ��,����˵��,strNO
    'cllBalance.Add Array(Val(Nvl(rsTmp!�����ID)), Trim(Nvl(rsTmp!����)), IIf(Val(Nvl(rsTmp!���㿨���)) <> 0, 1, 0), Trim(Nvl(rsTmp!������ˮ��)), Trim(Nvl(rsTmp!����˵��))), strNO
    If cllBalance Is Nothing Then CallBackBalanceInterface = True: Exit Function
    '�����:58567
    bln���ѿ� = Val(cllBalance(1)(2)) = 1
    lng�����ID = cllBalance(1)(0)
    
    '�����ID,����,�Ƿ����ѿ�(1-��;0-��),������ˮ��,����˵��,strNO
    If lng�����ID = 0 Then CallBackBalanceInterface = True: Exit Function
    
    str���� = cllBalance(1)(1)
    strSwapGlideNO = cllBalance(1)(3)
    strSwapMemo = cllBalance(1)(4)
    If lng���ѽ���ID <> 0 Then str������Ϣ = str������Ϣ & "||5|" & lng���ѽ���ID
    If lng�ҺŽ���ID <> 0 Then str������Ϣ = str������Ϣ & "||4|" & lng�ҺŽ���ID
    If str������Ϣ <> "" Then str������Ϣ = Mid(str������Ϣ, 3)
    
    
    If lng���ѽ���ID <> 0 Then
        strSQL = " Select ����ID,���ʷ��� From סԺ���ü�¼  Where ��¼����=5 And NO =(Select Max(NO) From סԺ���ü�¼ where ����ID=[1] and  ��¼����=5  )  and ��¼״̬=2"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng���ѽ���ID)
        If rsTemp.EOF Then
            strErrMsg = "δ�ҵ��˿���Ϣ�����ܼ���": Exit Function
        End If
        lng�˿�����ID = Val(Nvl(rsTemp!����ID))
    End If
    
    If lng�ҺŽ���ID <> 0 Then
        strSQL = "Select ����ID From ������ü�¼  Where ��¼����=4 And NO =(Select Max(NO) From ������ü�¼ where ����ID=[1] and  ��¼����=4  )  and ��¼״̬=2"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng�ҺŽ���ID)
        If rsTemp.EOF Then
            strErrMsg = "δ�ҵ��˺���Ϣ�����ܼ���": Exit Function
        End If
        lng�Һų���ID = Val(Nvl(rsTemp!����ID))
    End If

    '81489,Ƚ����,2015-1-22,�˷Ѵ������ID
    If lng�˿�����ID <> 0 Then strSwapExtendInfor = strSwapExtendInfor & "||5|" & lng�˿�����ID
    If lng�Һų���ID <> 0 Then strSwapExtendInfor = strSwapExtendInfor & "||4|" & lng�Һų���ID
    If strSwapExtendInfor <> "" Then strSwapExtendInfor = Mid(strSwapExtendInfor, 3)
    strTemp = strSwapExtendInfor
    'zlReturnMoney(frmMain As Object, ByVal lngModule As Long, _
        ByVal lngCardTypeID As Long, ByVal strCardNo As String, ByVal strBalanceIDs As String, _
        ByVal dblMoney As Double, _
        ByRef strSwapGlideNO As String, ByRef strSwapMemo As String, _
        ByRef strSwapExtendInfor As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ʻ��ۿ���˽���
    '���:frmMain-���õ�������
    '       lngModule-���õ�ģ���
    '       lngCardTypeID-�����ID:ҽ�ƿ����.ID
    '       strCardNo-����
    '       strBalanceIDs-����֧�����漰�Ľ���ID(����ԭ����ID):
    '                           ��ʽ:�շ�����(|ID1,ID2��IDn||�շ�����n|ID1,ID2��IDn
    '                           �շ�����:1-Ԥ����,2-����,3-�շ�,4-�Һ�,5-ҽ�ƿ��տ�
    '       dblMoney-�˿���
    '       strSwapNo-������ˮ��(�ۿ�ʱ�Ľ�����ˮ��)
    '       strSwapMemo-����˵��(�ۿ�ʱ�Ľ���˵��)
    '       strSwapExtendInfor-���룬�����˷ѵĳ���ID��
    '                           ��ʽ:�շ�����1|ID1,ID2��IDn||�շ�����n|ID1,ID2��IDn
    '                           �շ�����:1-Ԥ����,2-����,3-�շ�,4-�Һ�,5-ҽ�ƿ��տ�
    '       strSwapExtendInfor-���������׵���չ��Ϣ
    '           ��ʽΪ:��Ŀ����1|��Ŀ����2||��||��Ŀ����n|��Ŀ����n ÿ����Ŀ�в��ܰ���|�ַ�
    If gobjSquare.objSquareCard.zlReturnMoney(Me, mlngModul, lng�����ID, bln���ѿ�, str����, str������Ϣ, dblMoney, strSwapGlideNO, strSwapMemo, strSwapExtendInfor) = False Then Exit Function
    Set cllUpdate = New Collection: Set cllThreeSwap = New Collection
    If lng�˿�����ID <> 0 Then
        '�����:58536
        If Not bln���ѿ� Then
            Call zlAddUpdateSwapSQL(False, lng�˿�����ID, lng�����ID, bln���ѿ�, str����, strSwapGlideNO, strSwapMemo, cllUpdate)
        End If
        If strTemp <> strSwapExtendInfor Then
            Call zlAddThreeSwapSQLToCollection(False, lng�˿�����ID, lng�����ID, bln���ѿ�, str����, strSwapExtendInfor, cllThreeSwap)
        End If
    End If
    If lng�Һų���ID <> 0 Then
        Call zlAddUpdateSwapSQL(False, lng�Һų���ID, lng�����ID, bln���ѿ�, str����, strSwapGlideNO, strSwapMemo, cllUpdate)
        If strTemp <> strSwapExtendInfor Then
            Call zlAddThreeSwapSQLToCollection(False, lng�Һų���ID, lng�����ID, bln���ѿ�, str����, strSwapExtendInfor, cllThreeSwap)
        End If
    End If
    CallBackBalanceInterface = True
    Exit Function
Errhand:
    If ErrCenter() = 1 Then Resume
    SaveErrLog
End Function
Private Function IsValiedMzNo(ByVal lng����ID As Long, ByRef str����� As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������
    '���:str�����-�����
    '����:str�����-�����µ������
    '����:�Ϸ�,����true,���򷵻�false
    '����:���˺�
    '����:2011-10-31 10:22:12
    '����:42616
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str�����1 As String, strNew����� As String
    str�����1 = str�����
    If mTy_Para.blnԤԼ����������� And mbytMode = 1 Then IsValiedMzNo = True: Exit Function
    
    If str����� = "" And mbln������ Then
        Call MsgBox("δ���������,���ܼ���!", vbInformation + vbOKOnly, gstrSysName)
        If txt�����.Enabled Then txt�����.SetFocus
        Exit Function
    End If
    
    If Not Exist�����(str�����, lng����ID) Then IsValiedMzNo = True: Exit Function
    '42638
    If Not (gbln�Զ������ Or mblnStation) Then
        If str����� <> "" Then
            Call MsgBox("��ǰ�����:" & str�����1 & " �Ѿ�����������ʹ��,���ܼ���!", vbInformation + vbOKOnly, gstrSysName)
            If txt�����.Enabled Then txt�����.SetFocus
            Exit Function
        End If
    End If
    
    
    '���»�ȡ�����
GoGetMzNo:
    strNew����� = zlGet�����
    If Len(strNew�����) > txt�����.MaxLength Then
           MsgBox "��ǰ������Ѿ�����������ʹ��,ϵͳ�Զ����������Ϊ:" & strNew����� & _
               vbCrLf & "��������������������ų���:" & txt�����.MaxLength & "λ,������һ�������!", vbInformation, gstrSysName
           If txt�����.Enabled Then txt�����.SetFocus
           Exit Function
    End If
    If strNew����� <> "" Then
        If Exist�����(strNew�����, lng����ID) Then GoTo GoGetMzNo:
        '����:42616�Զ����������,������,ֱ�ӱ���
        If gbln�Զ������ Then
            str����� = strNew�����: IsValiedMzNo = True: Exit Function
        End If
        '��Ҫ����
        If MsgBox("��ǰ�����:" & str�����1 & " �Ѿ�����������ʹ��," & IIf(strNew����� <> "", vbCrLf & "  ϵͳ�Զ�����Ϊ" & strNew�����, "") & " �Ƿ����?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            txt�����.Text = strNew�����
            If txt�����.Enabled Then txt�����.SetFocus
            Exit Function
        End If
        '�������û�����ʱ,�򲢷�ԭ��,�ٴα�����ʹ��,��˻�Ҫ���������Ƿ�������ʹ��
        If Exist�����(strNew�����, lng����ID) Then
            If Not (gbln�Զ������ Or mblnStation) Then
                Call MsgBox("��ǰ�����:" & str����� & " �Ѿ�����������ʹ��,���ܼ���!", vbInformation + vbOKOnly, gstrSysName)
                txt�����.Text = strNew�����
                If txt�����.Enabled Then txt�����.SetFocus
                Exit Function
            End If
            GoTo GoGetMzNo:
        End If
    End If
    str����� = strNew�����
    txt�����.Text = str�����
    If str����� = "" And mbln������ Then
         Call MsgBox("δ���������,���ܼ���!", vbInformation + vbOKOnly, gstrSysName)
         If txt�����.Enabled Then txt�����.SetFocus
         Exit Function
     End If
     IsValiedMzNo = True
End Function

Private Sub ReInitPatiInvoice(Optional blnFact As Boolean = True, _
    Optional ByVal intInsure_IN As Integer = 0, Optional ByVal lng����ID_In As Long = 0)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���³�ʼ�����˷�Ʊ��Ϣ
    '���:blnFact-�Ƿ�����ȡ��Ʊ��
    '����:���˺�
    '����:2011-04-29 14:17:33
    '����:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strInvoiceFormat As String, lng����ID As Long
    Dim intInsure As Integer, strUseType As String
    If mblnStartFactUseType = False Then Exit Sub
    
    lng����ID = lng����ID_In
    
    If lng����ID_In = 0 Then
        If Not mrsInfo Is Nothing Then
            If mrsInfo.State = 1 Then lng����ID = mrsInfo!����ID
        End If
    End If
    
    If mblnStationPrice Then
        Exit Sub
    End If
    
    intInsure = IIf(intInsure_IN <> 0, intInsure_IN, mintInsure)
    strUseType = mstrUseType
    mstrUseType = "": mlngShareUseID = 0: mintInvoiceFormat = 0
    mstrUseType = zl_GetInvoiceUserType(lng����ID, 0, intInsure)
    '�л���Ʊ������
    If mstrUseType <> strUseType And mblnStartFactUseType Then mlng����ID = 0
    mlngShareUseID = zl_GetInvoiceShareID(mlngModul, mstrUseType)
    mintInvoiceFormat = zl_GetInvoicePrintFormat(mlngModul, mstrUseType)
    mintInvoicePrint = zl_GetInvoicePrintMode(mlngModul, mstrUseType)
    
    'Call ShowBillFormat
    If blnFact Then Call RefreshFact
End Sub

Private Function GetActiveView()
    '******************************************************************************
    '   �õ���ǰ�Һ�ҵ��  ��ȡ�������͵�����
    '******************************************************************************
        Dim strSQL          As String
        Dim rsTmp           As ADODB.Recordset
        Dim lng��¼ID       As Long
        
        On Error GoTo Hd
        lng��¼ID = Val(vsfPlan.TextMatrix(vsfPlan.Row, GetCol("��¼ID")))
        
        strSQL = "Select 1��From �ٴ������¼ Where ID=[1] And Nvl(�Ƿ��ʱ��,0)=1 "
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng��¼ID)
         If rsTmp.RecordCount > 0 And vsfPlan.TextMatrix(vsfPlan.Row, GetCol("��ſ���")) <> "" Then
            '*********************
            'ר�Һŷ�ʱ��
            '*********************
            mViewMode = v_ר�Һŷ�ʱ��
        '78640:���ϴ�,2014/10/16,�ҺŴ�ԤԼ��ʾ���п�ԤԼ�ĺű�
         ElseIf rsTmp.RecordCount > 0 And vsfPlan.TextMatrix(vsfPlan.Row, GetCol("��ſ���")) = "" And (mbytMode = 1 Or (mbytMode = 0 And chkBooking.Visible And chkBooking.Value = 1)) Then
            '*********************
            '��ͨ�ŷ�ʱ��
            '*********************
            mViewMode = V_��ͨ�ŷ�ʱ��
         ElseIf vsfPlan.TextMatrix(vsfPlan.Row, GetCol("��ſ���")) <> "" Then
            '*********************
            'ר�ҺŲ���ʱ��
            '*********************
            mViewMode = v_ר�Һ�
            vsfList.Visible = True
            picSplit.Visible = True
          Else
            '*********************
            '��ͨ��
            '*********************
            mViewMode = V_��ͨ��
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
    '����ʱ��
    '����ʱ���Ƿ���سɹ����Ƿ��з�ʱ��
    '**************************************
    Dim strSQL         As String
    Dim dateCur        As Date
    Dim lng��¼ID      As Long
    On Error GoTo Hd
     '�Һŷ�ʱ��
    strSQL = "Select ��¼id, ���, To_Char(��ʼʱ��, 'hh24') || ':00' As ʱ���, To_Char(��ʼʱ��, 'hh24:mi') As ��ʼʱ��," & vbNewLine & _
            "       To_Char(��ֹʱ��, 'hh24:mi') As ����ʱ��, ���� As ��������, �Ƿ�ԤԼ, ԤԼ˳���, ��ʼʱ�� As ��ϸ��ʼʱ��, ��ֹʱ�� As ��ϸ����ʱ�� " & vbNewLine & _
            "From �ٴ�������ſ���" & vbNewLine & _
            "Where ��¼id = [1] And ��ʼʱ�� <> ��ֹʱ�� " & vbNewLine & _
            "Order By ��ϸ��ʼʱ��"
    lng��¼ID = Val(vsfPlan.TextMatrix(vsfPlan.Row, GetCol("��¼ID")))
    Set mrsʱ��� = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng��¼ID)
    If mrsʱ���.EOF Then Exit Function
    
    InitTimePlan = True
    Exit Function
Hd:
    If ErrCenter() = 1 Then Resume
    SaveErrLog
End Function

Private Function Check��Ч�ű�(ByVal str�ű� As String, ByVal datThis As Date, Optional ByVal blnPlan As Boolean = False) As Boolean
   '***********************************************************
   '�ԹҺŻ���ԤԼʱ
   '������Чʱ�����֤
   '***********************************************************
    Dim strSQL          As String
    Dim rsTmp           As ADODB.Recordset
    Dim rsʱ���        As ADODB.Recordset
    Dim str����         As String
    Dim dat��ʼʱ��     As Date
    Dim dat����ʱ��     As Date
    Dim blnOK           As Boolean
    Dim strʱ��()       As String
    Dim i               As Long
    Dim Datsys          As Date
    
    '******************************
    '�Һż��ʱ �ڷ�ʱ�ε������
    'ֻ�ڹҺ��¼�� ��Ϊ ԤԼ������
    '����ʱ�䲻��С�� ʱ�εĿ�ʼʱ��
    '******************************
     On Error GoTo Hd
    If blnPlan = False And mbytMode = 0 And mViewMode = v_ר�Һŷ�ʱ�� Then
        Datsys = zlDatabase.Currentdate
        If datThis <= Datsys Then
            MsgBox "ʱ�εĿ�ʼʱ��" & Format(datThis, "HH:MM") & "С���˵�ǰʱ��" & Format(Datsys, "hh:MM") & "!����", vbOKOnly, Me.Caption
            Exit Function
        End If
    End If
    If blnPlan Then
        Datsys = zlDatabase.Currentdate
        If datThis <= Datsys Then
            MsgBox "ԤԼʱ��" & Format(datThis, "yyyy-mm-DD HH:MM") & "С���˵�ǰʱ��" & Format(Datsys, "yyyy-mm-DD hh:MM") & "!����", vbOKOnly, Me.Caption
            Exit Function
        End If
    End If
 
   Check��Ч�ű� = True
    Exit Function
Hd:
    If ErrCenter() = 1 Then
        Resume
    End If
    SaveErrLog
End Function

Private Sub InitActionType()
    '-------------------------
    '��ȡ �Ƿ�����˷�ʱ�εĴ���ʽ
    '�ж�����Ϊ �ҺŰ����б��Ƿ�������
    '-------------------------
    Dim strSQL       As String
    Dim rsTmp        As ADODB.Recordset
    strSQL = _
    "    Select 1  dt From  �ٴ������¼ Where �Ƿ��ʱ�� = 1 And Rownum< 2"
    
    On Error GoTo Hd
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    mcustomTime = t_��ͨ
    If rsTmp.RecordCount <> 0 Then mcustomTime = t_ʱ��
    Select Case mcustomTime
    Case t_��ͨ:
        Me.dtpAppointmentDate.CustomFormat = "yyyy-MM-dd"
        Me.dtpAppointmentTime.CustomFormat = "HH:mm"
        Form_Resize
    Case t_ʱ��:
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
    '��Ϣ��
    '------------------------------------------------
    If strTitle = "" Then strTitle = Me.Caption
    MsgBox strMsg, vbInformation, strTitle
End Sub

Private Function SetBrushCard(ByVal objContrl As Object, KeyAscii As Integer) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ˢ������
    '���:
    '����:
    '����:ˢ����ȡ��,����true,���򷵻�False
    '����:���˺�
    '����:2011-11-10 10:01:51
    '����:38603
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Static sngBegin As Single
    Dim sngNow As Single, blnCard As Boolean, lngҽ�ƿ����� As Long
    If Not (mblnStation And mTy_Para.bln�Һű���ˢ�� And mbytMode = 0) Then Exit Function
    lngҽ�ƿ����� = IDKind.GetCardNoLen
    objContrl.PasswordChar = IIf(IDKind.ShowPassText, "*", "")
    '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
    objContrl.IMEMode = 0
    
    '�Ƿ�ˢ�����
    blnCard = KeyAscii <> 8 And Len(objContrl.Text) = lngҽ�ƿ����� - 1 And objContrl.SelLength <> Len(objContrl.Text)
    If blnCard Or KeyAscii = 13 Then
        If KeyAscii <> 13 Then
            objContrl.Text = objContrl.Text & Chr(KeyAscii)
            objContrl.SelStart = Len(objContrl.Text)
        End If
        KeyAscii = 0
        mblnCard = True
        Call txtPatient_Validate(True)
        mblnCard = False
        '���˺�:27494  20100117
        If Replace(txtPatient.Text, vbCrLf, "") = "" Then
            DoEvents: txtPatient.SetFocus
        End If
    Else
        If InStr(":��;��?��" & Chr(22), Chr(KeyAscii)) > 0 Then
            KeyAscii = 0 'ȥ��������ţ����Ҳ�����ճ��
        Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        End If
        '��ȫˢ�����
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
'����IDCard
    '����С�����е�mobjIDCard�ͱ�ҳ���mobjIDCard��ͻ
    '���� ��������ˢ ���֤ ԭ��δ�ҵ�
    If (mbytMode = 0 Or mbytMode = 1) And mbytInState = 0 Then
        Set mobjIDCard = New clsIDCard
        Call mobjIDCard.SetParent(Me.Hwnd)
        If Me.ActiveControl Is Me.txtPatient And Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (Me.txtPatient.Text = "")
    End If
End Sub

Public Function GetʧԼ��(ByVal lng��¼ID As Long, datThis As Date) As Long
   '��ȡ������ĳһ��.ԤԼʧԼ��
    Dim strSQL  As String
    Dim rsTmp   As ADODB.Recordset
    Dim strBegin  As String, strEnd As String
    If mTy_Para.blnʧԼ���ڹҺ� = False Or mTy_Para.lngԤԼ��Чʱ�� = 0 Then Exit Function
    strSQL = "Select Count(1) As ʧԼ��" & vbNewLine & _
            " From ���˹Һż�¼" & vbNewLine & _
            " Where �����¼ID = [1] And ��¼���� = 2 And ��¼״̬ = 1 And ����ʱ�� - [2] / 24 / 60 < Sysdate And ����ʱ�� Between to_Date([3],'YYYY-MM-DD') And to_Date([4],'YYYY-MM-DD') - 1/24/60/60"
    strBegin = Format(datThis, "yyyy-MM-dd")
    strEnd = Format(datThis + 1, "yyyy-MM-dd")
    On Error GoTo Hd
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng��¼ID, mTy_Para.lngԤԼ��Чʱ��, strBegin, strEnd)
    If rsTmp.EOF Then
        GetʧԼ�� = 0
        Set rsTmp = Nothing
        Exit Function
    End If
    GetʧԼ�� = Val(Nvl(rsTmp!ʧԼ��, 0))
    Set rsTmp = Nothing
   Exit Function
Hd:
    If ErrCenter() = 1 Then
        Resume
    End If
    SaveErrLog
End Function

Public Sub zl_StationInitPatient(ByVal lng����ID As Long)
    '����˵��:���﹤��վ����ʱ��ʼ��������Ϣ
    '����˵��:str�����
    If mTy_Para.bln�Һű���ˢ�� Or mblnStation = False Or lng����ID = 0 Then Exit Sub
    txtPatient.Text = "-" & lng����ID
    txtPatient_Validate False
End Sub
Private Sub AddDeposit()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��Ԥ���
 
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objFun          As Object
    Dim lng����ID       As Long
    Err = 0: On Error Resume Next
    Set objFun = CreateObject("zl9Patient.clsPatient")
    If Err <> 0 Then Exit Sub
    
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '���ܣ� ����Ԥ�����տ��
    '������
    '   lngModul:��Ҫִ�еĹ������
    '   cnMain:����������ݿ�����
    '   frmMain:������
    '   strDBUser:��ǰ���ݿ��¼�û���
    '  bytCallObject:���˺����(0-Ԥ�������(ȱʡ��);1-���˷��ò�ѯ����,2-ҽ�ƿ�����)
    '  lng����ID-ȱʡ�Ĳ���ID
    '  lng��ҳID-ȱʡ����ҳID
    '  dblDefPrePayMoney-ȱʡ��Ԥ�����
    If Not mrsInfo Is Nothing Then lng����ID = Val(Nvl(mrsInfo!����ID))
    If objFun.PlusDeposit(glngSys, gcnOracle, Me, gstrDBUser, 2, lng����ID, 0, 0, 0) = False Then
        Set objFun = Nothing
        Exit Sub
    End If
    Set objFun = Nothing
    If lng����ID <> 0 Then
        txtPatient.Text = "-" & lng����ID
        mblnOnVilidate = True
        Call txtPatient_Validate(False)
        mblnOnVilidate = False
    End If
End Sub

Private Sub InitTimeSect(ByVal lng��¼ID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��ʱ���
    '����:���˺�
    '����:2012-03-12 15:45:57
    '����:45509
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    On Error GoTo errHandle
    strSQL = "Select ��ʼʱ��,��ֹʱ��,ȱʡԤԼʱ�� As ȱʡʱ��  From �ٴ������¼ Where ID=[1]"

    Set mrsALLʱ��� = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng��¼ID)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub SetDefaultRegistTime()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ȱʡ��ԤԼʱ��
    '����:���˺�
    '����:2012-03-12 15:49:38
    '����:45509
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str���� As String, strʱ�� As String
    Dim dtValue As Date, str���� As String
    Dim strȱʡʱ�� As String, Datsys As Date
    Static str�ϴκ��� As String
    On Error GoTo errHandle
    If mblnAppointmentChange Then Exit Sub
    Datsys = zlDatabase.Currentdate
    With vsfPlan
        If .Row < 0 Or .Row > .Rows - 1 Then Exit Sub
        If .ColIndex(mstrCurKey) < 0 Then Exit Sub
       str���� = .Cell(flexcpData, .Row, .ColIndex(mstrCurKey))
       str���� = .TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("��¼ID"))
    End With
    
    Call InitTimeSect(Val(str����))
    If mrsALLʱ���.EOF Then
        dtpAppointmentTime.Value = Format(Datsys, "HH:MM:SS")
        If dtpAppointmentDate.Visible Then
            txt����ʱ��.Text = Format(dtpAppointmentDate.Value, "yyyy-MM-dd") & " " & Format(dtpAppointmentTime.Value, "hh:mm:ss")
        Else
            txt����ʱ��.Text = Format(Datsys, "yyyy-MM-dd") & " " & Format(dtpAppointmentTime.Value, "hh:mm:ss")
        End If
        str�ϴκ��� = str����
        Exit Sub
    Else
        str�ϴκ��� = str����
    End If
    
    If (mViewMode = V_��ͨ�ŷ�ʱ�� Or mViewMode = v_ר�Һŷ�ʱ��) Then
        If mbytMode = 1 Or chkBooking.Value = 1 Then
            txt����ʱ��.Text = Format(Format(dtpAppointmentDate.Value, "yyyy-mm-dd" & " " & Format(dtpAppointmentTime.Value, "hh:mm:ss")), "yyyy-mm-dd hh:mm:ss")
        Else
            txt����ʱ��.Text = Format(Format(Datsys, "yyyy-mm-dd" & " " & Format(dtpAppointmentTime.Value, "hh:mm:ss")), "yyyy-mm-dd hh:mm:ss")
        End If
        str�ϴκ��� = str����
        Exit Sub
    Else
        If dtpAppointmentDate.Visible Then
            If Format(Nvl(mrsALLʱ���!ȱʡʱ��, dtpAppointmentDate.Value), "yyyy-MM-dd") <> Format(dtpAppointmentDate.Value, "yyyy-MM-dd") Then
                If Format(Nvl(mrsALLʱ���!ȱʡʱ��, dtpAppointmentDate.Value), "yyyy-MM-dd") > Format(dtpAppointmentDate.Value, "yyyy-MM-dd") Then
                    dtpAppointmentTime.Value = Format(Nvl(mrsALLʱ���!��ʼʱ��, Datsys), "HH:MM:SS")
                Else
                    dtpAppointmentTime.Value = Format(Nvl(mrsALLʱ���!��ֹʱ��, Datsys), "HH:MM:SS")
                End If
            Else
                dtpAppointmentTime.Value = Format(Nvl(mrsALLʱ���!ȱʡʱ��, Datsys), "HH:MM:SS")
            End If
            txt����ʱ��.Text = Format(dtpAppointmentDate.Value, "yyyy-MM-dd") & " " & Format(dtpAppointmentTime.Value, "hh:mm:00")
        Else
            If Format(Nvl(mrsALLʱ���!ȱʡʱ��, Datsys), "hh:mm:ss") < Format(Datsys, "hh:mm:ss") Then
                dtpAppointmentTime.Value = Format(Datsys, "HH:MM:SS")
            Else
                dtpAppointmentTime.Value = Format(Nvl(mrsALLʱ���!ȱʡʱ��, Datsys), "hh:mm:ss")
            End If
            txt����ʱ��.Text = Format(Datsys, "yyyy-MM-dd") & " " & Format(dtpAppointmentTime.Value, "hh:mm:00")
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
    '����:�˺Ų���(���˺����⸣����frmMain����������˵��
    '���:frmMain-���õ�������
    '     intCancel-0=�˺�;1=�˲�����;2=�˸��ӷ�
    '����:�˷ѳɹ�����true,���򷵻�False
    '����:���˺�
    '����:2013-10-23 17:09:50
    '˵��:
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
    '����:ȡ��ԤԼ����
    '���:frmMain-���õ�������
    '����:�˷ѳɹ�����true,���򷵻�False
    '����:������
    '����:2016-04-12
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mstrNoIn = strNoIn:   mstrPrivs = strPrivs:    mlngModul = lngModul
    mbytMode = 3:    mbytInState = 1: mblnCenter = blnCenter
    mblnOk = False
    Me.Show 1, frmMain
    CancelApp = mblnOk
End Function

Private Function GetMaxLapseNO() As Long
    '����˵��:��ȡ������ſ���������Ч�����Ƕ���
    '����ֵ:
    Dim i As Long
    Dim j As Long
    Dim nStart As Long
    Dim lngResult As Long
    Dim lngTmp As Long
    If mViewMode = V_��ͨ�� Or mViewMode = V_��ͨ�ŷ�ʱ�� Then Exit Function
    nStart = IIf(mViewMode = v_ר�Һ�, 0, 1)
    With vsfList
        For i = 0 To .Rows - 1
            For j = nStart To .Cols - 1
                If Trim(.TextMatrix(i, j)) <> "" Then
                     If .Cell(flexcpForeColor, i, j) <> vbRed And .Cell(flexcpForeColor, i, j) <> vbBlue And .Cell(flexcpForeColor, i, j) <> vbGrayText And .Cell(flexcpForeColor, i, j) <> &HC000C0 Then ' And .Cell(flexcpForeColor, i, j) <> vbGreen then
                         '�ճ��� ��ʱ�������� �����Ժ���ӹ���
                        If Not mrsSNState Is Nothing And .Cell(flexcpForeColor, i, j) <> vbGreen Then
                            lngTmp = Val(Getʱ��(i, j, False))
                            mrsSNState.Filter = "���=" & lngTmp
                            If mrsSNState.RecordCount > 0 Then
                                GetMaxLapseNO = lngTmp
                            End If
                        End If
                         
                     Else
                        If mViewMode = v_ר�Һŷ�ʱ�� Then
                            If .Cell(flexcpForeColor, i, j) = &HC000C0 And mTy_Para.bln������ѡ�� = False Then
                                '�������������ѡ��,ͬʱ��ԤԼ����,�ݲ�����
                            Else
                                
                                GetMaxLapseNO = Val(Getʱ��(i, j, False))
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

'��ȡidkind��Ĭ��kindֵ
Private Function IDKindDefaultKind() As Long
    Dim lngIndex As Long
    'IDkind��Ĭ��Kind
    If IDKind.DefaultCardType = "" Then
        lngIndex = -1
    Else
        If IsNumeric(IDKind.DefaultCardType) Then
           lngIndex = IDKind.GetKindIndex(IDKind.GetfaultCard.����)
        Else
           lngIndex = IDKind.GetKindIndex(IDKind.DefaultCardType)
        End If
    End If
    IDKindDefaultKind = lngIndex
End Function

'�ؼ������Ƿ�ƥ��
Private Function IsCardType(ByVal IDKindCtl As IDKindNew, ByVal strCardName As String) As Boolean
    If IDKindCtl Is Nothing Then Exit Function
    If UCase(TypeName(IDKindCtl)) <> "IDKINDNEW" Then Exit Function
    Select Case strCardName
     Case "����", "��������￨"
          IsCardType = IDKindCtl.GetCurCard.���� Like "����*"
     Case "���֤", "���֤��", "�������֤"
          IsCardType = IDKindCtl.GetCurCard.���� Like "*���֤*"
     Case "IC����", "IC��"
          IsCardType = IDKindCtl.GetCurCard.���� Like "IC��*"
     Case "ҽ����"
          IsCardType = IDKindCtl.GetCurCard.���� = "ҽ����"
     Case "�����"
          IsCardType = IDKindCtl.GetCurCard.���� = "�����"
     Case Else
            If IDKindCtl.GetCurCard Is Nothing Then Exit Function
            If Not IsNumeric(strCardName) Or Val(strCardName) <= 0 Then Exit Function
            If IDKindCtl.GetCurCard.�ӿ���� <= 0 Then Exit Function
            IsCardType = IDKindCtl.GetCurCard.�ӿ���� = Val(strCardName)
     End Select
End Function

'Private Function SetCreateCardObject() As Boolean
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    '����:�����ƿ�����
'    '����:����
'    '����:2012-12-17 11:06:41
'    '�����:56599
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
'��鲡���ڵ����Ƿ���ԤԼ����
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim rsInfo As ADODB.Recordset
    Dim strOutNo As String
    Dim frmNew As frmSelRegist
    Dim blnExit As Boolean
    Dim strMsg As String

    If mbytInState = 1 Then Exit Function
    If InStr(1, mstrPrivs, ";����ԤԼ;") = 0 Then Exit Function
    If Not (chkCancel.Value = 0 And chkPrint.Value = 0 And chkBooking.Value = 0 And Not mblnStation) Then Exit Function
    If mbytMode = 1 Or mbytMode = 2 Then Exit Function

    strSQL = "Select a.NO, a.����id, a.����, a.�ű�, a.����, a.����ʱ��, a.�Ǽ�ʱ��,b.���� as ִ�п��� " & vbNewLine & _
           "       From ���˹Һż�¼ a,���ű� b" & vbNewLine & _
           "       Where a.����ʱ�� Between Trunc(Sysdate) And Trunc(Sysdate) + 1 - 1 / 24 / 60 / 60 And a.��¼���� = 2 And a.��¼״̬ = 1 And a.����ID=[1] And A.ִ�в���ID=B.ID"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngPatientID)
    If rsTmp.EOF Then Exit Function

    If rsTmp.RecordCount = 1 Then
        'ֻ��һ���Һż�¼,���Ѳ���Ա�Ƿ���ձ���ԤԼ����
        strMsg = "����[" & Nvl(rsTmp!����) & "]�ڽ����ڿ���[" & Nvl(rsTmp!ִ�п���) & "]����ԤԼ����(" & Nvl(rsTmp!NO) & ")�Ƿ����?"
        If MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton1, Me.Caption) = vbYes Then
            Call ReadBooking(rsTmp!NO)
        Else
            Exit Function    '����ȡ����ԤԼ����
        End If
    Else
        'ֻ��һ���Һż�¼,���Ѳ���Ա�Ƿ���ձ���ԤԼ����
        strMsg = "����[" & Nvl(rsTmp!����) & "]�ڽ���ԤԼ�˶��ŵ���,�Ƿ���Ҫ����?"
        If MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton1, Me.Caption) = vbYes Then

            Call CloseIDCard    '47007
            Set frmNew = New frmSelRegist
            If frmNew.ShowRegist(Me, mstrPrivs, mblnOlnyBJYB, mTy_Para.intԤԼʧЧ����, strOutNo, rsInfo, Val(Nvl(rsTmp!����ID))) = False Then
                blnExit = True
            End If
            If Not frmNew Is Nothing Then Unload frmNew
            Set frmNew = Nothing
            Call NewCardObject
            If blnExit Then Exit Function
            Call ReadBooking(strOutNo)
        Else
            Exit Function    '����ȡ����ԤԼ����
        End If
    End If
    zlExistsTodaysAppointment = True
End Function
Private Sub SetDelBillCtlEnabled(Optional bln�������� As Boolean)
    '���ò����˺�ʱ,��������˷ѿؼ�״̬
    Dim blnEnabled As Boolean
    Dim blnNotEnabled As Boolean
    If Not (mbytInState = 1 And mbytMode = 4 Or chkCancel.Value = 1) Then Exit Sub
    If bln�������� Then chk������.Enabled = False: Exit Sub '��������.���ܲ�����,������ʱ��֧��

    If mrsBill Is Nothing Then Exit Sub
    If mrsBillAdvance Is Nothing Then Exit Sub
    
    mrsBillAdvance.Filter = 0
    mrsBill.Filter = "���ӱ�־=1"
    If mrsBill.RecordCount = 0 Then
        blnNotEnabled = blnNotEnabled Or True
    End If
    mrsBill.Filter = 0
    chk������.Enabled = Not blnNotEnabled And mintCancel = 0
End Sub
Private Sub InitInputMaxLen()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ���������󳤶�
    '����:���˺�
    '����:2013-11-11 11:28:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    txtPatient.MaxLength = zlGetPatiInforMaxLen.intPatiName
    txtPatientPrint.MaxLength = txtPatient.MaxLength
    txt����.MaxLength = zlGetPatiInforMaxLen.intPatiAge
    txt�����.MaxLength = zlGetPatiInforMaxLen.intPatiMzNo
End Sub

Private Function zlGetInvoiceGroupUseID(ByRef lng����ID As Long, _
    Optional intNum As Integer = 1, Optional strInvoiceNO As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡƱ�ݵ�����ID
    '���:lng����ID-����id
    '       intNum-ҳ��
    '       strInvoiceNO-����ķ�Ʊ��
    '����:lng����ID-����ID
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2013-11-19 16:32:57
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    lng����ID = GetInvoiceGroupID(IIf(gblnSharedInvoice, 1, 4), intNum, lng����ID, glng�Һ�ID, strInvoiceNO, IIf(mblnStartFactUseType, mstrUseType, ""))
    If lng����ID <= 0 Then
        Select Case lng����ID
            Case 0 '����ʧ��
            Case -1
                If Trim(mstrUseType) = "" Then
                    MsgBox "��û�����ú͹��õ��շ�Ʊ��,��������һ��Ʊ�ݻ����ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
                Else
                    MsgBox "��û�����ú͹��õġ�" & mstrUseType & "���շ�Ʊ��,��������һ��Ʊ�ݻ����ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
                End If
                Exit Function
            Case -2
                If Trim(mstrUseType) = "" Then
                    MsgBox "���صĹ���Ʊ���Ѿ�����,��������һ��Ʊ�ݻ��������ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
                Else
                    MsgBox "���صĹ���Ʊ�ݵġ�" & mstrUseType & "���շ�Ʊ���Ѿ�����,��������һ��Ʊ�ݻ��������ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
                End If
                Exit Function
            Case -3
                MsgBox "��ǰƱ�ݺ��벻�ڿ����������ε���ЧƱ�ݺŷ�Χ��,���������룡", vbInformation, gstrSysName
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

Private Function zlIsAllowPatiChargeFeeMode(ByVal lng����ID As Long, ByVal intԭ����ģʽ As Integer) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�Ƿ�����ı䲡���շ�ģʽ
    '���:lng����ID-����ID
    '       intԭ����ģʽ-0��ʾ�Ƚ��������;1��ʾ�����ƺ����
    '����:��������շ�ģʽ,����true,���򷵻�False
    '����:���˺�
    '����:2013-12-25 10:06:49
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    Dim dtDate As Date, intDay As Integer
    On Error GoTo errHandle
    
    If mbytMode = 1 Then zlIsAllowPatiChargeFeeMode = True: Exit Function 'ԤԼ������
    'ģʽδ������ֱ�ӷ���true
    If intԭ����ģʽ = mPatiChargeMode Then zlIsAllowPatiChargeFeeMode = True: Exit Function
    
      
    If intԭ����ģʽ = 1 Then
        'ԭΪ�����ƺ�����Ҵ���δ����õ�,�������ü���ģʽ
        strSQL = "" & _
        "   Select 1 " & _
        "   From ����δ����� " & _
        "   Where ����id = [1] And (��Դ;�� In (1, 4) Or ��Դ;�� = 3 And Nvl(��ҳid, 0) = 0) And Nvl(���, 0) <> 0 And Rownum < 2"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID)
        If rsTemp.EOF = False Then
            MsgBox "ע��:" & vbCrLf & "  ��ǰ���˵ľ���ģʽΪ�����ƺ�����Ҵ���δ����ã�" & _
                                          vbCrLf & "����������ò��˵ľ���ģʽ,������ȶ�δ����ý��ʺ�" & _
                                          vbCrLf & "�ٹҺŻ򲻵������˵ľ���ģʽ", vbOKOnly + vbInformation, gstrSysName
            Exit Function
        End If
        intDay = IIf(lbl��.Visible, -1 * gSysPara.Sy_Reg.bytNoDayseMergency, -1 * gSysPara.Sy_Reg.bytNODaysGeneral)
        dtDate = DateAdd("d", intDay, zlDatabase.Currentdate)
        ' �ϴ�Ϊ"�����ƺ����",����Ϊ"�Ƚ��������"��,ͬʱ����δ����ҽ��ҵ�����ݵ� ,
        '   ��������ľ���ģʽ
        strSQL = "Select 1 " & _
        " From ���˹Һż�¼ A, ����ҽ����¼ B " & _
        " Where a.����id + 0 = b.����id And a.No || '' = b.�Һŵ�  " & _
        "               And a.��¼״̬ = 1 And a.��¼���� = 1 And a.�Ǽ�ʱ�� - 0 >= [2] " & _
        "               And  a.����id = [1] And rownum<2"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, dtDate)
        If rsTemp.EOF Then
            'δ����ҽ������
            MsgBox "ע��:" & vbCrLf & "  ��ǰ���˵ľ���ģʽΪ�����ƺ����," & vbCrLf & "  ����������ò��˵ľ���ģʽ!", vbOKOnly + vbInformation, gstrSysName
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
    '����:��Ϣ���ʹ���
    '���: strNO-�Һŵ���
    '����:���˺�
    '����:2014-03-11 11:59:22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim objXML As New clsXML
    
    '0-�Һ�,1-ԤԼ,2-����,3-ȡ��ԤԼ ,4-�˺� ԤԼ������ģʽ:0-�Һ�,��ʱԤԼҪ�շ�,1-ԤԼ,���շ�
    If Not (mbytMode = 0 Or mbytMode = 2) Or mbytInState = 1 Then Exit Sub
    If mobjMsgModule Is Nothing Then Exit Sub
    If mobjMsgModule.IsConnect = False Then Exit Sub

    strSQL = "" & _
    " Select A.id ,A.����,nvl(A.�����,B.�����) as �����,A.����Id,b.���֤��,A.NO,A.ִ�в���ID,C.���� as ִ�в�������,A.����,A.ִ����  " & _
    " From ���˹Һż�¼ A,������Ϣ B,���ű� C  " & _
    " where A.No=[1] and a.��¼״̬ =1 And a.��¼����=1 and a.����ID=b.����id and a.ִ�в���id=c.id"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
    If rsTemp.EOF Then Exit Sub
    
    'ZLHIS_REGIST_001 ���ﲡ�˹Һ�֪ͨ
    '�ڵ�����    ����    ����    �ظ�    ����    ȱʡֵ  ֵ������
    '<patient_info>
    '    <patient_id>����ID</patient_id>
    '    <patient_name>��������</patient_name>
    '    <identity_card>���֤��</identity_card>
    '    <out_number>�����</out_number>
    '</patient_info>
    '<register_info>
    '    <register_id>�Һ�id</register_id>
    '    <register_no>�Һŵ���</register_no>
    '    <register_dept_id>�Һſ���id</register_dept_id>
    '    <register_dept_title>�Һſ���</register_dept_title>
    '    <register_room>�Һ�����</register_room>
    '    <register_doctor>�Һ�ҽ��</register_doctor>
    '</register_info>
    objXML.ClearXmlText
    Call objXML.AppendNode("patient_info")
        Call objXML.appendData("patient_id", Val(Nvl(rsTemp!����ID)))
        Call objXML.appendData("patient_name", Nvl(rsTemp!����))
        Call objXML.appendData("identity_card", Nvl(rsTemp!���֤��))
        Call objXML.appendData("out_number", Nvl(rsTemp!�����))
    Call objXML.AppendNode("patient_info", True)
    
    Call objXML.AppendNode("register_info")
        Call objXML.appendData("register_id", Val(Nvl(rsTemp!ID)))
        Call objXML.appendData("register_no", strNO)
        Call objXML.appendData("register_dept_id", Val(Nvl(rsTemp!ִ�в���id)))
        Call objXML.appendData("register_dept_title", Nvl(rsTemp!ִ�в�������))
        Call objXML.appendData("register_room", Nvl(rsTemp!����))
        Call objXML.appendData("register_doctor", Nvl(rsTemp!ִ����))
    Call objXML.AppendNode("register_info", True)
    Call mobjMsgModule.CommitMessage("ZLHIS_REGIST_001", objXML.XmlText)
    objXML.ClearXmlText
 End Sub
 
 Private Function ShowPatiPic() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾ������Ƭ
    '����:Ƚ����
    '����:2014-7-8
    '---------------------------------------------------------------------------------------------------------------------------------------------
    picPatiPicBack.Visible = True
    Set imgPatiPic.Picture = mobjfrmPatiInfo.imgPatient.Picture
    lblShow.Visible = imgPatiPic.Picture = 0
 End Function

Private Sub LoadIDImage()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������֤ͼ��
    '����:���˺�
    '����:2014-06-30 16:20:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Dim objStdPic As StdPicture
    
    If mobjIDCard Is Nothing Then Exit Sub
    Call mobjIDCard.GetPhotoAsStdPicture(objStdPic)
    mobjfrmPatiInfo.imgPatient.Picture = objStdPic
    mobjfrmPatiInfo.mlngͼ����� = 4
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Public Property Get SendCard() As Boolean
    SendCard = mblnSendCard
End Property

Private Sub Update֤��(ByVal lng����ID As Long, ByVal str֤���� As String)
    '���ܣ����µ�ǰ֤�����͵Ŀ���
    '�����:90875
    Dim strSQL As String, rsTmp As ADODB.Recordset
    On Error GoTo Errhand
    If IDKind֤��.IDKind = IDKind֤��.GetKindIndex("���֤��") Then Exit Sub
    If str֤���� = "���֤��" Then Exit Sub
    txt֤��.Text = "": txt֤��.Tag = ""
    If mrsInfo Is Nothing Then Exit Sub
    strSQL = "Select A.����,B.���� from ����ҽ�ƿ���Ϣ A,ҽ�ƿ���� B,֤������ C " & _
            "Where A.�����ID=B.ID And B.����=C.���� And A.����ID=[1] And B.����=[2] Order by C.���� "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, str֤����)
    If Not rsTmp.EOF Then txt֤��.Text = Nvl(rsTmp!����): txt֤��.Tag = txt֤��.Text
    Exit Sub
Errhand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub txt֤��_GotFocus()
    zlControl.TxtSelAll txt֤��
End Sub

Private Sub txt֤��_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txt֤��_Validate(Cancel As Boolean)
    If txt֤��.Text = txt֤��.Tag Then Exit Sub
    '���²�����Ϣ
    Call CopyZJTofrmPatiInfo
    If Trim(txt֤��.Text) = "" Then Exit Sub
    If Len(Trim(txt֤��.Text)) > 30 Then
         MsgBox "֤�������ַ���������ַ���30,������ַ������Զ��س���", vbInformation, gstrSysName
         txt֤��.Tag = Mid(Trim(txt֤��.Text), 1, 30)
         txt֤��.Text = Mid(Trim(txt֤��.Text), 1, 30)
    End If
    Call GetPatient(IDKind֤��.GetCurCard, txt֤��.Text, False, False, Cancel, True)
End Sub

Private Function AddCertificate(ByVal lng����ID As Long, ByRef colPro As Collection, ByVal dtCurdate As Date) As Boolean
    '-------------------------------------------------------------------------------------------------------------------------
    '����:����֤��������Ϣ������ǵ�һ�ν��������
    '����:���ϴ�
    'ʱ��:2015/12/17 17:37:27
    '����:90875
    '-------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim lngRow As Integer, lngCol As Integer
    Dim lngID As Long, strCode As String
    
    On Error GoTo Errhand
    If IDKind֤��.IDKind = IDKind֤��.GetKindIndex("���֤��") Or txt֤��.Text = "" Then AddCertificate = True: Exit Function
    '��鿨���Ƿ�����ʹ��
    strSQL = "Select 1 from ����ҽ�ƿ���Ϣ A,ҽ�ƿ���� B " & _
            "Where A.�����ID=B.ID And B.����=[1] And B.�Ƿ�֤��=1 And A.����=[2] And  A.����ID<>[3]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, IDKind֤��.GetCurCard.����, Trim(txt֤��.Text), lng����ID)
    If rsTemp.RecordCount > 0 Then
        MsgBox IDKind֤��.GetCurCard.���� & ":" & txt֤��.Text & "���ڱ�ʹ��,����!", vbInformation, gstrSysName
        If txt֤��.Visible And txt֤��.Enabled Then txt֤��.SetFocus
        Exit Function
    End If
    '�󶨿�ǰҪ�жϿ�����Ƿ����
    strSQL = "Select B.ID,B.����,B.���ų���,B.����,A.����,A.����ID,Decode(A.���� ,NULL,1,0) as ��ʶ from ����ҽ�ƿ���Ϣ A,ҽ�ƿ���� B " & _
            "Where A.�����ID(+)=B.ID And B.�Ƿ�֤��=1 And A.״̬(+)=0 And B.����=[1] "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, IDKind֤��.GetCurCard.����)

    If rsTemp.RecordCount = 0 Then
        lngID = zlDatabase.GetNextId("ҽ�ƿ����")
        strCode = zlDatabase.GetMax("ҽ�ƿ����", "����", 4)
        mobjfrmPatiInfo.mstrFirstCode = strCode
        Call AddCardTypeSQL(0, lngID, strCode, IDKind֤��.GetCurCard.����, IDKind֤��.GetCurCard.����, Len(Trim(txt֤��.Text)), colPro)
    ElseIf Len(Trim(txt֤��.Text)) > Val(Nvl(rsTemp!���ų���)) Then
        lngID = rsTemp!ID
        Call AddCardTypeSQL(1, lngID, Nvl(rsTemp!����), IDKind֤��.GetCurCard.����, IDKind֤��.GetCurCard.����, Len(Trim(txt֤��.Text)), colPro)
    Else
        lngID = rsTemp!ID
    End If
    
    '����֤������
    rsTemp.Filter = "����='" & IDKind֤��.GetCurCard.���� & "' And ����='" & Trim(txt֤��.Text) & "'"
    If rsTemp.RecordCount = 0 Then
        '�Ƚ�����ԭ���Ŀ����
        rsTemp.Filter = "����='" & IDKind֤��.GetCurCard.���� & "' And ����ID= " & lng����ID
        If rsTemp.RecordCount > 0 Then
            rsTemp.MoveFirst
            Do While Not rsTemp.EOF
                'Zl_ҽ�ƿ��䶯_Insert
                 strSQL = "Zl_ҽ�ƿ��䶯_Insert("
                '      �䶯����_In   Number,
                '��������=1-����(��11�󶨿�);2-����;3-����(13-����ͣ��);4-�˿�(��14ȡ����); ��-�������(ֻ��¼);6-��ʧ(16ȡ����ʧ)
                strSQL = strSQL & "" & 14 & ","
                '      ����id_In     סԺ���ü�¼.����id%Type,
                strSQL = strSQL & "" & lng����ID & ","
                '      �����id_In   ����ҽ�ƿ���Ϣ.�����id%Type,
                strSQL = strSQL & "" & lngID & ","
                '      ԭ����_In     ����ҽ�ƿ���Ϣ.����%Type,
                strSQL = strSQL & "'" & "" & "',"
                '      ҽ�ƿ���_In   ����ҽ�ƿ���Ϣ.����%Type,
                strSQL = strSQL & "'" & rsTemp!���� & "',"
                '      �䶯ԭ��_In   ����ҽ�ƿ��䶯.�䶯ԭ��%Type,
                '      --�䶯ԭ��_In:�������������䶯ԭ��Ϊ����.���ܵ�
                strSQL = strSQL & "'" & "֤����ȡ����" & "',"
                '      ����_In       ������Ϣ.����֤��%Type,
                strSQL = strSQL & "'" & "" & "',"
                '      ����Ա����_In סԺ���ü�¼.����Ա����%Type,
                strSQL = strSQL & "'" & UserInfo.���� & "',"
                '      �䶯ʱ��_In   סԺ���ü�¼.�Ǽ�ʱ��%Type,
                strSQL = strSQL & "to_date('" & Format(dtCurdate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
                '      Ic����_In     ������Ϣ.Ic����%Type := Null,
                strSQL = strSQL & "'" & "" & "',"
                '      ��ʧ��ʽ_In   ����ҽ�ƿ��䶯.��ʧ��ʽ%Type := Null
                strSQL = strSQL & "NULL)"

                zlAddArray colPro, strSQL
                rsTemp.MoveNext
            Loop
        End If
            
        '����֤������
        'Zl_ҽ�ƿ��䶯_Insert
         strSQL = "Zl_ҽ�ƿ��䶯_Insert("
        '      �䶯����_In   Number,
        '��������=1-����(��11�󶨿�);2-����;3-����(13-����ͣ��);4-�˿�(��14ȡ����); ��-�������(ֻ��¼);6-��ʧ(16ȡ����ʧ)
        strSQL = strSQL & "" & 11 & ","
        '      ����id_In     סԺ���ü�¼.����id%Type,
        strSQL = strSQL & "" & lng����ID & ","
        '      �����id_In   ����ҽ�ƿ���Ϣ.�����id%Type,
        strSQL = strSQL & "" & lngID & ","
        '      ԭ����_In     ����ҽ�ƿ���Ϣ.����%Type,
        strSQL = strSQL & "'" & "" & "',"
        '      ҽ�ƿ���_In   ����ҽ�ƿ���Ϣ.����%Type,
        strSQL = strSQL & "'" & Trim(txt֤��.Text) & "',"
        '      �䶯ԭ��_In   ����ҽ�ƿ��䶯.�䶯ԭ��%Type,
        '      --�䶯ԭ��_In:�������������䶯ԭ��Ϊ����.���ܵ�
        strSQL = strSQL & "'" & "֤������" & "',"
        '      ����_In       ������Ϣ.����֤��%Type,
        strSQL = strSQL & "'" & "" & "',"
        '      ����Ա����_In סԺ���ü�¼.����Ա����%Type,
        strSQL = strSQL & "'" & UserInfo.���� & "',"
        '      �䶯ʱ��_In   סԺ���ü�¼.�Ǽ�ʱ��%Type,
        strSQL = strSQL & "to_date('" & Format(dtCurdate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
        '      Ic����_In     ������Ϣ.Ic����%Type := Null,
        strSQL = strSQL & "'" & "" & "',"
        '      ��ʧ��ʽ_In   ����ҽ�ƿ��䶯.��ʧ��ʽ%Type := Null
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
    '����:������������
    '����:���˺�
    '����:2017-08-09 11:25:57
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnInit As Boolean
    If mbytMode <> 0 Then Exit Sub
    
    '�����ӿڳ�ʼ��
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
    '����:��ȡ��ǰ�Һŵ��ĺϼƽ��
    '���:blnOnlyReg-�Ƿ������ȡ�Һŷ���
    '     blnNoBook-��ȡ������
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2017-11-03 16:53:41
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dbl�ϼ� As Double, i As Integer
    Dim k As Integer
    If Not blnOnlyReg Then
        dbl�ϼ� = FormatEx(GetTotalFromMshMoney, 5)
    Else
        If mrsItems Is Nothing Then
             GetRegistMoney = FormatEx(GetTotalFromMshMoney, 3): Exit Function
        End If
        mrsItems.Filter = " ���� <> 4"
        If mrsItems.RecordCount = 0 Then
            mrsItems.Filter = 0
            GetRegistMoney = FormatEx(GetTotalFromMshMoney, 3): Exit Function
        End If
        With mrsItems
            Do While Not .EOF
                dbl�ϼ� = dbl�ϼ� + GetTotalFromMshMoney(Nvl(mrsItems!��Ŀ����, "-"))
                .MoveNext
            Loop
        End With
        mrsItems.Filter = 0
    End If
    
    If blnNoBook Then
        If Not mrsItems Is Nothing Then
            mrsItems.Filter = " ���� = 3"
            Do While Not mrsItems.EOF
                dbl�ϼ� = dbl�ϼ� + GetTotalFromMshMoney(Nvl(mrsItems!��Ŀ����, "-"))
                mrsItems.MoveNext
            Loop
            mrsItems.Filter = 0
        End If
    End If
    GetRegistMoney = FormatEx(dbl�ϼ�, 5)
End Function

Private Function GetTotalFromMshMoney(Optional ByVal str��Ŀ���� As String = "") As Double
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ���ܽ��
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2017-11-03 16:57:41
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblMoney As Double, i As Long
    
    On Error GoTo errHandle
    With vsfMoney
        For i = 1 To .Rows - 1
            If str��Ŀ���� = "" Or Trim(.TextMatrix(i, 0)) = str��Ŀ���� Then
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
    '����:��ȡ���ѽ��
    '����:���˺�
    '����:2017-11-03 17:39:27
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dbl�ϼ� As Double
    If mrsItems Is Nothing Then GetCardMoney = 0: Exit Function
    mrsItems.Filter = " ���� = 4"
    If mrsItems.RecordCount = 0 Then
        mrsItems.Filter = 0
        GetCardMoney = 0: Exit Function
    End If
    With mrsItems
        Do While Not .EOF
            dbl�ϼ� = dbl�ϼ� + GetTotalFromMshMoney(Nvl(mrsItems!��Ŀ����, "-"))
            .MoveNext
        Loop
    End With
    mrsItems.Filter = 0
End Function

Private Function CheckIsPrice() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ܼ�鵱ǰ�Ƿ�Ϊ���۵�����
    '����:����Ϊ���۵��ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2017-11-03 14:03:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dbl�ɿ� As Double
    Dim dbl���� As Double, dblԤ�� As Double
    Dim blnPrice As Boolean
    
    On Error GoTo errHandle
    If mRegistFeeMode = EM_RG_���� Then CheckIsPrice = False: Exit Function
    If InStr(1, "02", mbytMode) = 0 Then CheckIsPrice = False: Exit Function
    
    If Not gblnPrice Or txtPatient.Text = "" Then CheckIsPrice = False: Exit Function
    blnPrice = picBookingDate.Visible = False And vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("����")) <> ""
    If blnPrice Then blnPrice = GetRegistMoney <> 0
    
    CheckIsPrice = blnPrice
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function
 
Private Sub ReLoadCardFee(Optional ByVal blnFeedName As Boolean, Optional ByVal blnReflashfee As Boolean)
    '�뿪��鿨��
    Dim lng����ID As Long, lng�շ�ϸĿID As Long
    Dim strSQL As String, str���� As String
    Dim rsTmp As ADODB.Recordset
    
    gCurSendCard.lng�շ�ϸĿID = 0
    If gCurSendCard.rs���� Is Nothing Then Exit Sub
    If gCurSendCard.rs����.RecordCount = 0 Then Exit Sub
    If gCurSendCard.lng�����ID = 0 Then Exit Sub
    If Trim(txtPatient.Text) = "" Or Trim(mobjfrmPatiInfo.txt����.Text) = "" Then Exit Sub
    If mrsInfo Is Nothing Then
        lng����ID = 0
    ElseIf mrsInfo.State <> 1 Then
        lng����ID = 0
    Else
        lng����ID = mrsInfo!����ID
    End If
    If blnFeedName = False And lng����ID <> 0 Then Exit Sub
    
    str���� = Trim(txt����.Text)
    If IsNumeric(str����) Then str���� = str���� & cbo���䵥λ.Text
    gCurSendCard.rs����.MoveFirst
    
    strSQL = "Select Zl1_Ex_CardFee([1],[2],[3],[4],[5],[6],[7],[8],[9]) as �շ�ϸĿID From Dual "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "����", mlngModul, gCurSendCard.lng�����ID, Trim(mobjfrmPatiInfo.txt����.Text), lng����ID, _
                Trim(txtPatient.Text), NeedName(cbo�Ա�.Text), str����, txtIDCard.Text, Val(Nvl(gCurSendCard.rs����!�շ�ϸĿID)))
    If rsTmp.EOF Then Exit Sub
    
    lng�շ�ϸĿID = Val(Nvl(rsTmp!�շ�ϸĿID))
    Set rsTmp = zlGetSpecialItemFee("", mobjfrmPatiInfo.mstrPriceGrade, lng�շ�ϸĿID)
    If Not rsTmp Is Nothing Then
        Set gCurSendCard.rs���� = rsTmp
        gCurSendCard.lng�շ�ϸĿID = lng�շ�ϸĿID
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
    '��ʼ���Һ�
    Dim strDept As String
    Set mobjRegist = New clsRegist
    mobjRegist.zlInitCommon glngSys, gcnOracle, gstrDBUser
    mobjRegist.zlCancelRegNo '����ϴ��ǳ��������������Ҫ���н���
End Sub

Private Function ReserveRegNo(ByVal lng��¼ID As Long, ByRef lngSN As Long, _
                            ByVal str����ʱ�� As String, ByVal Datsys As Date) As Boolean
    Dim blnLock As Boolean, bln��ʱ�� As Boolean
    Dim strLockTime As String
    Dim strSQL As String, rsTmp As ADODB.Recordset
    On Error GoTo errH
    mlng���ż�¼ID = 0
    If vsfPlan.TextMatrix(vsfPlan.Row, GetCol("��ſ���")) <> "" And lng��¼ID <> 0 Then
        mlng���ż�¼ID = lng��¼ID
        bln��ʱ�� = (mViewMode = v_ר�Һŷ�ʱ�� Or mViewMode = V_��ͨ�ŷ�ʱ��)
        If Not (mbytMode = 2 Or mbytMode = 0 And mbytInState = 0 And mstrNoIn <> "") Then
            blnLock = True: strLockTime = str����ʱ��
        Else
            If mTy_Para.byt����ģʽ = 0 And bln��ʱ�� And Format(dtpAppointmentDate.Value, "yyyy-MM-dd") <> Format(Datsys, "yyyy-MM-dd") Then
                MsgBox "��ʱ�ε�ԤԼ�Һŵ�ֻ�ܵ�����ա�", vbInformation, gstrSysName
                Exit Function
            End If
            If (mTy_Para.byt����ģʽ = 0 And Format(dtpAppointmentDate.Value, "yyyy-MM-dd") <> Format(Datsys, "yyyy-MM-dd")) Then
                blnLock = True: strLockTime = Format(Datsys, "yyyy-MM-dd")
                strSQL = "Select a.Id" & vbNewLine & _
                        "From �ٴ������¼ a, �ٴ������¼ b" & vbNewLine & _
                        "Where a.��Դid = b.��Դid And a.�Ƿ��ʱ�� = b.�Ƿ��ʱ�� And a.�Ƿ���ſ��� = b.�Ƿ���ſ��� And a.����id = b.����id And" & vbNewLine & _
                        "      Nvl(a.ҽ��id, 0) = Nvl(b.ҽ��id, 0) And a.�ϰ�ʱ�� = b.�ϰ�ʱ�� And Nvl(a.�Ƿ񷢲�, 0) = 1 And a.�������� = Trunc(Sysdate) And" & vbNewLine & _
                        "      b.Id = [1] And Rownum < 2"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��������ID", lng��¼ID)
                If rsTmp.RecordCount = 0 Then
                    MsgBox "���յ���û�ж�Ӧ�ĳ��ﰲ�ţ��޷����ա�", vbInformation, gstrSysName
                    Exit Function
                End If
                mlng���ż�¼ID = rsTmp!ID
            End If
        End If
        If blnLock Then
            If mobjRegist.zlReserveRegNo(txt�ű�.Text, True, bln��ʱ��, strLockTime, lngSN, "�ҺŴ�������", mlng���ż�¼ID) = False Then Exit Function
        End If
    End If
    ReserveRegNo = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

