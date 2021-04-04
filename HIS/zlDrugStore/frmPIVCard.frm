VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPIVCard 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   8730
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14760
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   8730
   ScaleWidth      =   14760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame fraDetailCtr 
      BackColor       =   &H00FFEDDD&
      Height          =   480
      Left            =   0
      TabIndex        =   5
      Top             =   360
      Width           =   7575
      Begin VB.TextBox txtFind 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   3360
         MaxLength       =   12
         TabIndex        =   27
         Top             =   142
         Width           =   1815
      End
      Begin VB.CheckBox chkAll 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFEDDD&
         Caption         =   "全选"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   150
         Width           =   735
      End
      Begin VB.CheckBox chkType 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFEDDD&
         Caption         =   "配药"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   960
         TabIndex        =   7
         Top             =   150
         Value           =   1  'Checked
         Width           =   735
      End
      Begin VB.CheckBox chkType 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFEDDD&
         Caption         =   "打包"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   1800
         TabIndex        =   6
         Top             =   150
         Value           =   1  'Checked
         Width           =   735
      End
      Begin VB.Label lbl瓶签号 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFEDDD&
         BackStyle       =   0  'Transparent
         Caption         =   "瓶签号"
         Height          =   180
         Left            =   2760
         TabIndex        =   26
         Top             =   187
         Width           =   540
      End
   End
   Begin VB.PictureBox picHelp 
      BackColor       =   &H00FFEDDD&
      BorderStyle     =   0  'None
      Height          =   345
      Left            =   0
      ScaleHeight     =   345
      ScaleWidth      =   7695
      TabIndex        =   2
      Top             =   120
      Width           =   7695
      Begin VB.PictureBox picHelpIcon 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFEDDD&
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   20
         Picture         =   "frmPIVCard.frx":0000
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   3
         Top             =   30
         Width           =   240
      End
      Begin VB.Label lblHelp 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFEDDD&
         Caption         =   "提示："
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   260
         TabIndex        =   4
         Top             =   50
         Width           =   540
      End
   End
   Begin VB.PictureBox picMain 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   10335
      Left            =   120
      ScaleHeight     =   10335
      ScaleWidth      =   14055
      TabIndex        =   0
      Top             =   1080
      Width           =   14055
      Begin VB.PictureBox picPage 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   4320
         ScaleHeight     =   495
         ScaleWidth      =   4095
         TabIndex        =   29
         Top             =   6600
         Width           =   4095
         Begin VB.CommandButton cmdNext 
            BackColor       =   &H8000000E&
            Caption         =   "下一页(&N)"
            Height          =   350
            Left            =   3120
            TabIndex        =   31
            Top             =   0
            Width           =   975
         End
         Begin VB.CommandButton cmdPrevious 
            Caption         =   "上一页(&P)"
            Height          =   350
            Left            =   120
            TabIndex        =   30
            Top             =   0
            Width           =   975
         End
         Begin VB.Label lblCode 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Caption         =   "第1页/共10页"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1260
            TabIndex        =   32
            Top             =   75
            Width           =   1695
         End
      End
      Begin VB.HScrollBar HSc 
         Height          =   255
         LargeChange     =   50
         Left            =   360
         Max             =   100
         SmallChange     =   50
         TabIndex        =   25
         Top             =   7440
         Width           =   12615
      End
      Begin VB.PictureBox picLableMain 
         BackColor       =   &H80000010&
         BorderStyle     =   0  'None
         Height          =   2880
         Index           =   0
         Left            =   0
         ScaleHeight     =   2880
         ScaleWidth      =   3870
         TabIndex        =   9
         Top             =   0
         Visible         =   0   'False
         Width           =   3870
         Begin VB.PictureBox picLable 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H8000000E&
            ForeColor       =   &H80000008&
            Height          =   2790
            Index           =   0
            Left            =   45
            ScaleHeight     =   2760
            ScaleWidth      =   3750
            TabIndex        =   10
            Top             =   45
            Width           =   3780
            Begin VB.PictureBox picDrug 
               BackColor       =   &H8000000E&
               BorderStyle     =   0  'None
               Height          =   1125
               Index           =   0
               Left            =   0
               ScaleHeight     =   1125
               ScaleWidth      =   4095
               TabIndex        =   11
               Top             =   1200
               Width           =   4095
               Begin VB.Label lblDrug1 
                  AutoSize        =   -1  'True
                  BackColor       =   &H8000000E&
                  Caption         =   "10%葡萄糖注射液"
                  Height          =   180
                  Index           =   0
                  Left            =   0
                  TabIndex        =   14
                  Top             =   0
                  Width           =   1695
               End
               Begin VB.Label lblSpec1 
                  AutoSize        =   -1  'True
                  BackColor       =   &H8000000E&
                  Caption         =   "250ml/袋"
                  Height          =   180
                  Index           =   0
                  Left            =   1920
                  TabIndex        =   13
                  Top             =   0
                  Width           =   1080
               End
               Begin VB.Label lblAmount1 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackColor       =   &H8000000E&
                  Caption         =   "1.00"
                  Height          =   180
                  Index           =   0
                  Left            =   3240
                  TabIndex        =   12
                  Top             =   0
                  Width           =   480
               End
            End
            Begin VB.Image imgOtherDrug 
               Enabled         =   0   'False
               Height          =   240
               Index           =   0
               Left            =   60
               Picture         =   "frmPIVCard.frx":6852
               Top             =   2475
               Width           =   240
            End
            Begin VB.Label lblType 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H8000000E&
               ForeColor       =   &H00000000&
               Height          =   180
               Index           =   0
               Left            =   1200
               TabIndex        =   24
               Top             =   2520
               Width           =   2490
            End
            Begin VB.Label lblDept 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000E&
               Caption         =   "肿瘤介入病房"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   0
               Left            =   360
               TabIndex        =   23
               Top             =   120
               Width           =   1350
            End
            Begin VB.Image imgSex 
               Height          =   360
               Index           =   0
               Left            =   0
               Picture         =   "frmPIVCard.frx":6DDC
               Top             =   525
               Width           =   360
            End
            Begin VB.Label lblPatiName 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000E&
               Caption         =   "张三兄弟"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   0
               Left            =   480
               TabIndex        =   22
               Top             =   570
               Width           =   900
            End
            Begin VB.Label lblBatch 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H8000000E&
               Caption         =   "【1#】"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   12
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   240
               Index           =   0
               Left            =   3000
               TabIndex        =   21
               Top             =   120
               Width           =   780
            End
            Begin VB.Image imgPacker 
               Height          =   360
               Index           =   0
               Left            =   2640
               Picture         =   "frmPIVCard.frx":7546
               Top             =   60
               Width           =   360
            End
            Begin VB.Image imgCheck 
               Height          =   360
               Index           =   0
               Left            =   15
               Picture         =   "frmPIVCard.frx":7CB0
               Top             =   60
               Width           =   360
            End
            Begin VB.Label lblLine1 
               BackColor       =   &H00000000&
               Height          =   30
               Index           =   0
               Left            =   0
               TabIndex        =   20
               Top             =   460
               Width           =   4500
            End
            Begin VB.Image imgPrint 
               Height          =   360
               Index           =   0
               Left            =   2265
               Picture         =   "frmPIVCard.frx":841A
               Top             =   60
               Width           =   360
            End
            Begin VB.Label lblNo 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000E&
               Caption         =   "201110200001"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   0
               Left            =   0
               TabIndex        =   19
               Top             =   870
               Width           =   1440
            End
            Begin VB.Label lblBed 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000E&
               Caption         =   "32床"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   0
               Left            =   2040
               TabIndex        =   18
               Top             =   570
               Width           =   465
            End
            Begin VB.Label lblAge 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000E&
               Caption         =   "40岁"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   0
               Left            =   2880
               TabIndex        =   17
               Top             =   600
               Width           =   705
            End
            Begin VB.Label lblLine2 
               BackColor       =   &H00000000&
               Height          =   30
               Index           =   0
               Left            =   0
               TabIndex        =   16
               Top             =   1095
               Width           =   4500
            End
            Begin VB.Label lblTime 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000E&
               Caption         =   "10-20 15:30"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   0
               Left            =   1800
               TabIndex        =   15
               Top             =   870
               Width           =   1920
            End
            Begin VB.Image imgDrug 
               Height          =   240
               Index           =   0
               Left            =   300
               Picture         =   "frmPIVCard.frx":8B84
               Top             =   2475
               Width           =   240
            End
         End
      End
      Begin VB.VScrollBar VSLBar 
         Height          =   8295
         LargeChange     =   600
         Left            =   12960
         Max             =   300
         SmallChange     =   100
         TabIndex        =   1
         Top             =   360
         Width           =   255
      End
      Begin MSComctlLib.ListView lst批次 
         Height          =   1455
         Left            =   4440
         TabIndex        =   28
         Top             =   360
         Visible         =   0   'False
         Width           =   4600
         _ExtentX        =   8123
         _ExtentY        =   2566
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "批次"
            Object.Width           =   8114
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "颜色"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.Image imgStandardPacker 
      Height          =   360
      Index           =   1
      Left            =   14280
      Picture         =   "frmPIVCard.frx":F3D6
      Top             =   2280
      Width           =   360
   End
   Begin VB.Image imgUnCheck 
      Height          =   360
      Left            =   14280
      Picture         =   "frmPIVCard.frx":F8A3
      Top             =   1080
      Width           =   360
   End
   Begin VB.Image imgStandardUnCheck 
      Height          =   360
      Left            =   14280
      Picture         =   "frmPIVCard.frx":1000D
      Top             =   840
      Width           =   360
   End
   Begin VB.Image imgDown 
      Height          =   240
      Left            =   14280
      Picture         =   "frmPIVCard.frx":10777
      Top             =   5880
      Width           =   240
   End
   Begin VB.Image imgUp 
      Height          =   240
      Left            =   14280
      Picture         =   "frmPIVCard.frx":16FC9
      Top             =   5520
      Width           =   240
   End
   Begin VB.Image imgUnUp 
      Height          =   240
      Left            =   14280
      Picture         =   "frmPIVCard.frx":1D81B
      Top             =   5040
      Width           =   240
   End
   Begin VB.Image imgUnDown 
      Height          =   240
      Left            =   14280
      Picture         =   "frmPIVCard.frx":1DDA5
      Top             =   4680
      Width           =   240
   End
   Begin VB.Image imgBoy 
      Height          =   360
      Left            =   14280
      Picture         =   "frmPIVCard.frx":1E32F
      Top             =   4080
      Width           =   360
   End
   Begin VB.Image imgGirl 
      Height          =   360
      Left            =   14280
      Picture         =   "frmPIVCard.frx":1EA99
      Top             =   3600
      Width           =   360
   End
   Begin VB.Image imgStandardPrint 
      Height          =   360
      Left            =   14280
      Picture         =   "frmPIVCard.frx":1F203
      Top             =   3000
      Width           =   360
   End
   Begin VB.Image imgStandardUnPrint 
      Height          =   360
      Left            =   14280
      Picture         =   "frmPIVCard.frx":1F96D
      Top             =   2640
      Width           =   360
   End
   Begin VB.Image imgStandardUnPacker 
      Height          =   360
      Left            =   14280
      Picture         =   "frmPIVCard.frx":200D7
      Top             =   1920
      Width           =   360
   End
   Begin VB.Image imgStandardPacker 
      Height          =   360
      Index           =   0
      Left            =   14280
      Picture         =   "frmPIVCard.frx":20841
      Top             =   1560
      Width           =   360
   End
   Begin VB.Image imgStandardCheck 
      Height          =   360
      Left            =   14280
      Picture         =   "frmPIVCard.frx":20FAB
      Top             =   360
      Width           =   360
   End
End
Attribute VB_Name = "frmPIVCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

Private mIntOld As Long
Private mIntOldW As Long
Private mstr批次 As String
Private mInt选择 As Integer
Private mint标志 As Integer
Private mblnEdit As Boolean
Private mIntCount As Integer
Private mbln批次设置 As Boolean
Private mbln打包设置 As Boolean
Private mintIndex As Integer
Private mlng序号 As Long
Private mrsData As Recordset
Private mintPage As Integer
Private mbln审核 As Boolean
Private mlngSum As Long
Private mlngPre As Long
Private mstrStep As String

Private Const M_CST_SELECTED_COLOR = &HC5FEC9

'动态生成输液标签
Private Declare Function DestroyWindow Lib "user32 " (ByVal hWnd As Long) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long


Private Sub LoadTransLable(ByVal intIndex As Integer)
    
    Dim blnNewRow As Boolean
    
    If intIndex = 0 Then Exit Sub
    
    blnNewRow = ((intIndex Mod mIntCount) = 0)
    
    picLableMain(0).BackColor = M_CST_SELECTED_COLOR
    imgCheck(0).Tag = 0
    imgCheck(0).Picture = imgStandardUnCheck.Picture
    
    '底框
    Load picLableMain(intIndex)
    With picLableMain(intIndex)
        If blnNewRow = True Then
            .Left = picLableMain(0).Left
            .Top = picLableMain(intIndex - 1).Top + picLableMain(intIndex - 1).Height + 60
        Else
            .Left = picLableMain(intIndex - 1).Left + picLableMain(intIndex - 1).Width + 60
            .Top = picLableMain(intIndex - 1).Top
        End If
        
        .Visible = True
    End With
    
    '标签
    Load picLable(intIndex)
    Set picLable(intIndex).Container = picLableMain(intIndex)
    With picLable(intIndex)
        .Left = picLable(0).Left
        .Top = picLable(0).Top
        .Width = picLable(0).Width
        .Height = picLable(0).Height
        .Visible = True
    End With
    
    '标签中的内容
    Load imgCheck(intIndex)
    Set imgCheck(intIndex).Container = picLable(intIndex)
    With imgCheck(intIndex)
        .Left = imgCheck(0).Left
        .Top = imgCheck(0).Top
        .Width = imgCheck(0).Width
        .Height = imgCheck(0).Height
        .Visible = True
    End With
    
    Load lbldept(intIndex)
    Set lbldept(intIndex).Container = picLable(intIndex)
    With lbldept(intIndex)
        .Left = lbldept(0).Left
        .Top = lbldept(0).Top
        .Width = lbldept(0).Width
        .Height = lbldept(0).Height
        .Visible = True
    End With
    
    Load imgPrint(intIndex)
    Set imgPrint(intIndex).Container = picLable(intIndex)
    With imgPrint(intIndex)
        .Left = imgPrint(0).Left
        .Top = imgPrint(0).Top
        .Width = imgPrint(0).Width
        .Height = imgPrint(0).Height
        .Visible = True
    End With
    
    Load imgPacker(intIndex)
    Set imgPacker(intIndex).Container = picLable(intIndex)
    With imgPacker(intIndex)
        .Left = imgPacker(0).Left
        .Top = imgPacker(0).Top
        .Width = imgPacker(0).Width
        .Height = imgPacker(0).Height
        .Visible = True
    End With
    
    Load lblBatch(intIndex)
    Set lblBatch(intIndex).Container = picLable(intIndex)
    With lblBatch(intIndex)
        .Left = lblBatch(0).Left
        .Top = lblBatch(0).Top
        .Width = lblBatch(0).Width
        .Height = lblBatch(0).Height
        .Visible = True
    End With
    
    Load lblLine1(intIndex)
    Set lblLine1(intIndex).Container = picLable(intIndex)
    With lblLine1(intIndex)
        .Left = lblLine1(0).Left
        .Top = lblLine1(0).Top
        .Width = lblLine1(0).Width
        .Height = lblLine1(0).Height
        .Visible = True
    End With
    
    Load imgSex(intIndex)
    Set imgSex(intIndex).Container = picLable(intIndex)
    With imgSex(intIndex)
        .Left = imgSex(0).Left
        .Top = imgSex(0).Top
        .Width = imgSex(0).Width
        .Height = imgSex(0).Height
        .Visible = True
    End With
    
    Load lblPatiName(intIndex)
    Set lblPatiName(intIndex).Container = picLable(intIndex)
    With lblPatiName(intIndex)
        .Left = lblPatiName(0).Left
        .Top = lblPatiName(0).Top
        .Width = lblPatiName(0).Width
        .Height = lblPatiName(0).Height
        .Visible = True
    End With
    
    Load lblBed(intIndex)
    Set lblBed(intIndex).Container = picLable(intIndex)
    With lblBed(intIndex)
        .Left = lblBed(0).Left
        .Top = lblBed(0).Top
        .Width = lblBed(0).Width
        .Height = lblBed(0).Height
        .Visible = True
    End With
    
    Load lblAge(intIndex)
    Set lblAge(intIndex).Container = picLable(intIndex)
    With lblAge(intIndex)
        .Left = lblAge(0).Left
        .Top = lblAge(0).Top
        .Width = lblAge(0).Width
        .Height = lblAge(0).Height
        .Visible = True
    End With
    
    Load lblNo(intIndex)
    Set lblNo(intIndex).Container = picLable(intIndex)
    With lblNo(intIndex)
        .Left = lblNo(0).Left
        .Top = lblNo(0).Top
        .Width = lblNo(0).Width
        .Height = lblNo(0).Height
        .Visible = True
    End With
    
    Load lblTime(intIndex)
    Set lblTime(intIndex).Container = picLable(intIndex)
    With lblTime(intIndex)
        .Left = lblTime(0).Left
        .Top = lblTime(0).Top
        .Width = lblTime(0).Width
        .Height = lblTime(0).Height
        .Visible = True
    End With
    
    Load lblLine2(intIndex)
    Set lblLine2(intIndex).Container = picLable(intIndex)
    With lblLine2(intIndex)
        .Left = lblLine2(0).Left
        .Top = lblLine2(0).Top
        .Width = lblLine2(0).Width
        .Height = lblLine2(0).Height
        .Visible = True
    End With
    
    Load picDrug(intIndex)
    Set picDrug(intIndex).Container = picLable(intIndex)
    With picDrug(intIndex)
        .Left = picDrug(0).Left
        .Top = picDrug(0).Top
        .Width = picDrug(0).Width
        .Height = picDrug(0).Height
        .Visible = True
    End With
    
    
    Load imgOtherDrug(intIndex)
    Set imgOtherDrug(intIndex).Container = picLable(intIndex)
    With imgOtherDrug(intIndex)
        .Left = imgOtherDrug(0).Left
        .Top = imgOtherDrug(0).Top
        .Width = imgOtherDrug(0).Width
        .Height = imgOtherDrug(0).Height
        .Visible = True
    End With
    
    Load imgDrug(intIndex)
    Set imgDrug(intIndex).Container = picLable(intIndex)
    With imgDrug(intIndex)
        .Left = imgDrug(0).Left
        .Top = imgDrug(0).Top
        .Width = imgDrug(0).Width
        .Height = imgDrug(0).Height
        .Visible = False
    End With
    
    Load lblType(intIndex)
    Set lblType(intIndex).Container = picLable(intIndex)
    With lblType(intIndex)
        .Left = lblType(0).Left
        .Top = lblType(0).Top
        .Width = lblType(0).Width
        .Height = lblType(0).Height
        .Tag = ""
        .Caption = ""
        .Visible = True
    End With
End Sub

Private Sub LoadDrug(ByVal intIndex As Integer, ByVal intCon As Integer, ByVal blnNew As Boolean)
    '药品
    If blnNew Then
        Load lblDrug1(intIndex)
        Set lblDrug1(intIndex).Container = picDrug(intCon)
        With lblDrug1(intIndex)
            .Left = lblDrug1(intIndex - 1).Left
            .Top = lblDrug1(0).Top
            .Width = lblDrug1(intIndex - 1).Width
            .Height = lblDrug1(intIndex - 1).Height
            .Caption = ""
            .Visible = True
        End With
        
        Load lblSpec1(intIndex)
        Set lblSpec1(intIndex).Container = picDrug(intCon)
        With lblSpec1(intIndex)
            .Left = lblSpec1(intIndex - 1).Left
            .Top = lblSpec1(0).Top
            .Width = lblSpec1(intIndex - 1).Width
            .Height = lblSpec1(intIndex - 1).Height
            .Caption = ""
            .Visible = True
        End With
        
        Load lblAmount1(intIndex)
        Set lblAmount1(intIndex).Container = picDrug(intCon)
        With lblAmount1(intIndex)
            .Left = lblAmount1(intIndex - 1).Left
            .Top = lblAmount1(0).Top
            .Width = lblAmount1(intIndex - 1).Width
            .Height = lblAmount1(intIndex - 1).Height
            .Caption = ""
            .Visible = True
        End With
    Else
        Load lblDrug1(intIndex)
        Set lblDrug1(intIndex).Container = picDrug(intCon)
        With lblDrug1(intIndex)
            .Left = lblDrug1(intIndex - 1).Left
            .Top = lblDrug1(intIndex - 1).Top + lblDrug1(intIndex - 1).Height + 50
            .Width = lblDrug1(intIndex - 1).Width
            .Height = lblDrug1(intIndex - 1).Height
            .Caption = ""
            .Visible = True
        End With
        
        Load lblSpec1(intIndex)
        Set lblSpec1(intIndex).Container = picDrug(intCon)
        With lblSpec1(intIndex)
            .Left = lblSpec1(intIndex - 1).Left
            .Top = lblSpec1(intIndex - 1).Top + lblSpec1(intIndex - 1).Height + 50
            .Width = lblSpec1(intIndex - 1).Width
            .Height = lblSpec1(intIndex - 1).Height
            .Caption = ""
            .Visible = True
        End With
        
        Load lblAmount1(intIndex)
        Set lblAmount1(intIndex).Container = picDrug(intCon)
        With lblAmount1(intIndex)
            .Left = lblAmount1(intIndex - 1).Left
            .Top = lblAmount1(intIndex - 1).Top + lblAmount1(intIndex - 1).Height + 50
            .Width = lblAmount1(intIndex - 1).Width
            .Height = lblAmount1(intIndex - 1).Height
            .Caption = ""
            .Visible = True
        End With
    End If
End Sub

Private Sub chkAll_Click()
    mint标志 = 0
    
    Chk_all
End Sub
Private Sub Chk_all()
    Dim i As Integer
    Dim j As Integer
    
    For i = 0 To Me.picLableMain.UBound
        If Me.chkAll.Value = 1 Then
            imgCheck(i).Tag = 1
            imgCheck(i).Picture = imgStandardCheck.Picture
            
            imgCheck(i).Tag = 1
            imgCheck(i).Picture = imgStandardCheck.Picture
            Me.picLable(i).BackColor = M_CST_SELECTED_COLOR
            Me.picDrug(i).BackColor = M_CST_SELECTED_COLOR
            lbldept(i).BackColor = M_CST_SELECTED_COLOR
            lblBatch(i).BackColor = M_CST_SELECTED_COLOR
            lblPatiName(i).BackColor = M_CST_SELECTED_COLOR
            lblBed(i).BackColor = M_CST_SELECTED_COLOR
            lblAge(i).BackColor = M_CST_SELECTED_COLOR
            lblNo(i).BackColor = M_CST_SELECTED_COLOR
            lblTime(i).BackColor = M_CST_SELECTED_COLOR
            lblType(i).BackColor = M_CST_SELECTED_COLOR
            
            For j = CInt(IIf(Me.picDrug(i).Tag = "", "0", Me.picDrug(i).Tag)) - CInt(IIf(imgOtherDrug(i).Tag = "", "0", imgOtherDrug(i).Tag)) + 1 To CInt(IIf(Me.picDrug(i).Tag = "", "0", Me.picDrug(i).Tag))
                lblDrug1(j).BackColor = M_CST_SELECTED_COLOR
                lblSpec1(j).BackColor = M_CST_SELECTED_COLOR
                lblAmount1(j).BackColor = M_CST_SELECTED_COLOR
            Next
            
            Me.picLableMain(i).BackColor = vbRed
        Else
            imgCheck(i).Tag = 0
            imgCheck(i).Picture = imgStandardUnCheck.Picture
            
            Me.picLable(i).BackColor = vbWhite
            Me.picDrug(i).BackColor = vbWhite
            lbldept(i).BackColor = vbWhite
            lblBatch(i).BackColor = vbWhite
            lblPatiName(i).BackColor = vbWhite
            lblBed(i).BackColor = vbWhite
            lblAge(i).BackColor = vbWhite
            lblNo(i).BackColor = vbWhite
            lblTime(i).BackColor = vbWhite
            lblType(i).BackColor = vbWhite
            
            For j = CInt(IIf(Me.picDrug(i).Tag = "", "0", Me.picDrug(i).Tag)) - CInt(IIf(imgOtherDrug(i).Tag = "", "0", imgOtherDrug(i).Tag)) + 1 To CInt(IIf(Me.picDrug(i).Tag = "", "0", Me.picDrug(i).Tag))
                lblDrug1(j).BackColor = vbWhite
                lblSpec1(j).BackColor = vbWhite
                lblAmount1(j).BackColor = vbWhite
            Next
            Me.picLableMain(i).BackColor = M_CST_SELECTED_COLOR
        End If
    Next
    
    If mint标志 <> 1 Then
        frmPIVAMain.chkAllClick Me.chkAll.Value
        mint标志 = 0
    End If
End Sub

Private Sub chkType_Click(index As Integer)
    If chkType(0).Value = 0 And chkType(1).Value = 0 Then
        chkType(index).Value = 1
    End If
    
    frmPIVAMain.ChooseType index, Me.chkType(index).Value
End Sub

Private Sub cmdNext_Click()
    If (mstrStep = "00" Or mstrStep = "10" Or mstrStep = "11") Then
        Call Show医嘱(True)
    Else
        Call ShowCard(True)
    End If
    mlngPre = mlngPre + 1
    lblCode.Caption = "第" & mlngPre & "页/共" & mlngSum & "页"
    If mlngPre = mlngSum Then
        Me.cmdNext.Enabled = False
    End If
    Me.cmdPrevious.Enabled = True
End Sub

Private Sub cmdPrevious_Click()
    If (mstrStep = "00" Or mstrStep = "10" Or mstrStep = "11") Then
        Call Show医嘱(False)
    Else
        Call ShowCard(False)
    End If
    
    mlngPre = mlngPre - 1
    lblCode.Caption = "第" & mlngPre & "页/共" & mlngSum & "页"
    If mlngPre = 1 Then
        Me.cmdPrevious.Enabled = False
    End If
    Me.cmdNext.Enabled = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        Me.lst批次.Visible = False
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    picHelp.Move 0, 5, Me.ScaleWidth, picHelp.Height
    fraDetailCtr.Move 0, picHelp.Height + picHelp.Top - 120, Me.ScaleWidth, fraDetailCtr.Height
    picMain.Move 0, fraDetailCtr.Height + fraDetailCtr.Top + 10, Me.ScaleWidth, Me.ScaleHeight
    VSLBar.Move Me.ScaleWidth - Me.VSLBar.Width, 0, Me.VSLBar.Width, Me.ScaleHeight - (fraDetailCtr.Height + fraDetailCtr.Top + 10)
    HSc.Move 0, Me.VSLBar.Height - Me.HSc.Height, Me.picMain.Width - Me.VSLBar.Width
    picPage.Move picMain.Width / 2 - Me.picPage.Width / 2, HSc.Top - Me.picPage.Height
End Sub

Private Sub HSc_Change()
    Dim lngMove As Long
    Dim i As Integer
    Dim intRow As Integer
    
    lngMove = CLng(HSc.Value)
    lngMove = lngMove / Me.HSc.SmallChange * (Me.picLableMain(i).Width + 400)
    
    For i = 0 To Me.picLableMain.UBound
        Me.picLableMain(i).Left = Me.picLableMain(i).Left - (lngMove - mIntOldW)
    Next
    
    mIntOldW = lngMove
End Sub

Private Sub imgCheck_Click(index As Integer)
    Dim i As Integer
    
    If Val(imgCheck(index).Tag) = 0 Or (mbln审核 And Val(imgCheck(index).Tag) = 1) Then
        If Val(imgCheck(index).Tag) = 0 Then
            imgCheck(index).Tag = 1
            imgCheck(index).Picture = imgStandardCheck.Picture
        Else
            imgCheck(index).Tag = 2
            imgCheck(index).Picture = imgUnCheck.Picture
        End If
        '改变背景颜色
        Me.picLable(index).BackColor = M_CST_SELECTED_COLOR
        Me.picDrug(index).BackColor = M_CST_SELECTED_COLOR
        lbldept(index).BackColor = M_CST_SELECTED_COLOR
        lblBatch(index).BackColor = M_CST_SELECTED_COLOR
        lblPatiName(index).BackColor = M_CST_SELECTED_COLOR
        lblBed(index).BackColor = M_CST_SELECTED_COLOR
        lblAge(index).BackColor = M_CST_SELECTED_COLOR
        lblNo(index).BackColor = M_CST_SELECTED_COLOR
        lblTime(index).BackColor = M_CST_SELECTED_COLOR
        lblType(index).BackColor = M_CST_SELECTED_COLOR
        
        For i = CInt(Me.picDrug(index).Tag) - CInt(imgOtherDrug(index).Tag) + 1 To CInt(Me.picDrug(index).Tag)
            lblDrug1(i).BackColor = M_CST_SELECTED_COLOR
            lblSpec1(i).BackColor = M_CST_SELECTED_COLOR
            lblAmount1(i).BackColor = M_CST_SELECTED_COLOR
        Next
    Else
        imgCheck(index).Tag = 0
        imgCheck(index).Picture = imgStandardUnCheck.Picture
        '改变背景颜色
        Me.picLable(index).BackColor = vbWhite
        Me.picDrug(index).BackColor = vbWhite
        lbldept(index).BackColor = vbWhite
        lblBatch(index).BackColor = vbWhite
        lblPatiName(index).BackColor = vbWhite
        lblBed(index).BackColor = vbWhite
        lblAge(index).BackColor = vbWhite
        lblNo(index).BackColor = vbWhite
        lblTime(index).BackColor = vbWhite
        lblType(index).BackColor = vbWhite
        
        For i = CInt(Me.picDrug(index).Tag) - CInt(imgOtherDrug(index).Tag) + 1 To CInt(Me.picDrug(index).Tag)
            lblDrug1(i).BackColor = vbWhite
            lblSpec1(i).BackColor = vbWhite
            lblAmount1(i).BackColor = vbWhite
        Next
    End If
    picLable_Click index
    Call frmPIVAMain.CheckOne(Me.picLableMain(index).Tag, imgCheck(index).Tag)
    
End Sub

Private Sub imgDrug_Click(index As Integer)
    Dim intSum As Integer
    Dim intStart As Integer
    Dim intEnd As Integer
    Dim i As Integer
    
    intSum = CInt(Me.picDrug(index).Tag)
    intStart = CInt(imgDrug(index).Tag)
    intEnd = CInt(imgOtherDrug(index).Tag)
    
    If intStart + 5 = intSum + 1 Then Exit Sub
    
    '向上一条的按钮可用，并切换图标
    Me.imgOtherDrug(index).Picture = Me.imgUp.Picture
    Me.imgOtherDrug(index).Enabled = True
    
    For i = intSum - intEnd + 1 To intSum
        Me.lblDrug1(i).Top = Me.lblDrug1(i).Top - 225
        Me.lblSpec1(i).Top = Me.lblSpec1(i).Top - 225
        Me.lblAmount1(i).Top = Me.lblAmount1(i).Top - 225
    Next
    
    imgDrug(index).Tag = intStart + 1
    If intStart + 6 = intSum + 1 Then
        Me.imgDrug(index).Picture = Me.imgUnDown.Picture
        Me.imgDrug(index).Enabled = False
    Else
        Me.imgDrug(index).Picture = Me.imgDown.Picture
        Me.imgDrug(index).Enabled = True
    End If
End Sub

Private Sub imgOtherDrug_Click(index As Integer)
    Dim intSum As Integer
    Dim intStart As Integer
    Dim intEnd As Integer
    Dim i As Integer
    
    intSum = CInt(Me.picDrug(index).Tag)
    intStart = CInt(imgDrug(index).Tag)
    intEnd = CInt(imgOtherDrug(index).Tag)
    
    If intStart = intSum - intEnd + 1 Then Exit Sub
    
    '向下的按钮可用，并切换图标
    Me.imgDrug(index).Picture = Me.imgDown.Picture
    Me.imgDrug(index).Enabled = True
    
    For i = intSum - intEnd + 1 To intSum
        Me.lblDrug1(i).Top = Me.lblDrug1(i).Top + 225
        Me.lblSpec1(i).Top = Me.lblSpec1(i).Top + 225
        Me.lblAmount1(i).Top = Me.lblAmount1(i).Top + 225
    Next
    
'    For i = intStart - 1 To intStart + 4
'        Me.lblDrug1(i).Top = Me.lblDrug1(i).Top + 230
'        Me.lblSpec1(i).Top = Me.lblSpec1(i).Top + 230
'        Me.lblAmount1(i).Top = Me.lblAmount1(i).Top + 230
'    Next
     
    imgDrug(index).Tag = intStart - 1
    If intStart - 1 = intSum - intEnd + 1 Then
        Me.imgOtherDrug(index).Picture = Me.imgUnUp.Picture
        Me.imgOtherDrug(index).Enabled = False
    Else
        Me.imgOtherDrug(index).Picture = Me.imgUp.Picture
        Me.imgOtherDrug(index).Enabled = True
    End If

End Sub


Public Function ShowDetailCard(ByVal rsDetail As Recordset, ByVal str批次 As String, ByVal blnEdit As Boolean, ByVal intNum As Integer, ByVal bln批次设置 As Boolean, ByVal bln打包设置 As Boolean, ByVal strStep As String, ByVal bln审核 As Boolean) As Boolean
'*********************************************************
'加载数据
'*********************************************************
    Dim i As Integer
    Dim j As Integer
    Dim lng配药id As Long
    Dim intCount As Integer
    Dim intSum As Integer
    Dim IntSumMain As Integer
    
    chkAll.Enabled = False
    mstr批次 = str批次
    mblnEdit = blnEdit
    mIntCount = intNum
    mbln批次设置 = bln批次设置
    mbln打包设置 = bln打包设置
    Set mrsData = rsDetail
    mlng序号 = 0
    mintPage = 0
    mbln审核 = (strStep = "00" And bln审核)
    
    mstrStep = strStep
    If rsDetail.RecordCount > 0 Then
        rsDetail.MoveLast
        mlngSum = rsDetail!组号 \ (mIntCount * 3) + 1
    End If
    mlngPre = 1
    
    cmdPrevious.Enabled = False
    Me.cmdNext.Enabled = True
    
    If (strStep = "00" Or strStep = "10" Or strStep = "11") And bln审核 Then
        Call Show医嘱(True)
    Else
        Load批次
        Call ShowCard(True)
    End If
    
    lblCode.Caption = "第" & mlngPre & "页/共" & mlngSum & "页"
    chkAll.Enabled = True
    ShowDetailCard = True
End Function

Private Sub ShowCard(Optional ByVal blnNext As Boolean)
    Dim i As Integer
    Dim j As Integer
    Dim lng配药id As Long
    Dim intCount As Integer
    Dim intSum As Integer
    Dim IntSumMain As Integer
    Dim strFilter As String
    Dim str药品id As String
    Dim blnChange As Boolean
    strFilter = mrsData.Filter
    
    If InStr(1, strFilter, "And") > 0 Then
        strFilter = Mid(strFilter, 1, InStr(1, strFilter, "And") - 1)
    ElseIf InStr(1, strFilter, "组号") > 0 And InStr(1, strFilter, "And") = -1 Then
        strFilter = "0"
    End If
    
    For j = 1 To Me.lblDrug1.UBound
        Unload lblDrug1(j)
        Unload Me.lblAmount1(j)
        Unload Me.lblSpec1(j)
    Next
    
    For j = 1 To Me.picLableMain.UBound
        Unload lblTime(j)
        Unload lblNo(j)
        Unload lblBed(j)
        Unload lblAge(j)
        Unload lblPatiName(j)
        Unload imgSex(j)
        Unload lblBatch(j)
        Unload imgPacker(j)
        Unload imgPrint(j)
        Unload lbldept(j)
        Unload imgCheck(j)
        Unload lblLine1(j)
        Unload lblLine2(j)
        Unload picDrug(j)
        Unload Me.imgDrug(j)
        Unload Me.imgOtherDrug(j)
        Unload lblType(j)
        Unload Me.picLable(j)
        Unload picLableMain(j)
    Next
    If blnNext Then
        mrsData.Filter = IIf(strFilter <> "0", strFilter & " And ", "") & "组号>" & mlng序号 - 1
    Else
        mrsData.Filter = IIf(strFilter <> "0", strFilter & " And ", "") & "组号>" & mlng序号 - mintPage - (mIntCount * 3)
    End If
    mintPage = 0
    With mrsData
        .Sort = "组号,病区,姓名,性别,年龄,配药ID,费用序号"
        
        
            Me.picLableMain(0).Visible = True
            Me.imgCheck(0).Picture = imgStandardUnCheck.Picture
            Me.picLableMain(0).BackColor = M_CST_SELECTED_COLOR
            '改变背景颜色
            Me.picLable(0).BackColor = vbWhite
            Me.picDrug(0).BackColor = vbWhite
            lbldept(0).BackColor = vbWhite
            lblBatch(0).BackColor = vbWhite
            lblPatiName(0).BackColor = vbWhite
            lblBed(0).BackColor = vbWhite
            lblAge(0).BackColor = vbWhite
            lblNo(0).BackColor = vbWhite
            lblTime(0).BackColor = vbWhite
            lblType(0).BackColor = vbWhite
            lblDrug1(0).BackColor = vbWhite
            lblSpec1(0).BackColor = vbWhite
            lblAmount1(0).BackColor = vbWhite
            lblBatch(0).Visible = True
            imgPacker(0).Visible = True
            Me.imgPrint(0).Visible = True
        
        Do While Not .EOF
            If lng配药id <> !配药id Then
                
                mintPage = mintPage + 1
                If i = mIntCount * 3 Then
                    mlng序号 = !组号
                    chkAll.Enabled = True
                    If mlng序号 - mintPage - (mIntCount * 3) + 1 < 1 Then
                        Me.cmdPrevious.Enabled = False
                    End If
                    Exit Sub
                End If
                
                intCount = 0
                If i > Me.picLableMain.UBound Then
                    LoadTransLable (i)
                    Call LoadDrug(Me.lblDrug1.UBound + 1, i, True)
                    DoEvents
                End If
                
                If !执行标志 = 1 Then
                    imgCheck(i).Tag = 1
                    imgCheck(i).Picture = imgStandardCheck.Picture
                    Me.picLable(i).BackColor = M_CST_SELECTED_COLOR
                    Me.picDrug(i).BackColor = M_CST_SELECTED_COLOR
                    lbldept(i).BackColor = M_CST_SELECTED_COLOR
                    lblBatch(i).BackColor = M_CST_SELECTED_COLOR
                    lblPatiName(i).BackColor = M_CST_SELECTED_COLOR
                    lblBed(i).BackColor = M_CST_SELECTED_COLOR
                    lblAge(i).BackColor = M_CST_SELECTED_COLOR
                    lblNo(i).BackColor = M_CST_SELECTED_COLOR
                    lblTime(i).BackColor = M_CST_SELECTED_COLOR
                    lblType(i).BackColor = M_CST_SELECTED_COLOR
                End If
                
                lbldept(i).Caption = !科室
                lblBatch(i).Caption = "【" & !配药批次 & "】"
                lblBatch(i).Tag = "【" & !配药批次 & "】"
                lblBatch(i).ForeColor = zlStr.nvl(!颜色, 0)
                lblPatiName(i).Caption = !姓名
                lblBed(i).Caption = IIf(IsNull(!床号), "", !床号)
                lblAge(i).Caption = !年龄
                lblNo(i).Caption = !瓶签号
                lblTime(i).Caption = !执行时间
                Me.picLableMain(i).Tag = !配药id
                If !性别 = "女" Then
                    imgSex(i).Picture = Me.imgGirl.Picture
                Else
                    imgSex(i).Picture = imgBoy.Picture
                End If
                
                
                If !是否打包 > 0 Then
                    imgPacker(i).Tag = 2
                    imgPacker(i).Picture = imgStandardPacker(!是否打包 - 1).Picture
                Else
                    imgPacker(i).Tag = 0
                    imgPacker(i).Picture = imgStandardUnPacker.Picture
                End If
                
                If !打印标志 = 1 Then
                    imgPrint(i).Tag = 1
                    imgPrint(i).Picture = imgStandardPrint.Picture
                End If
                
                lng配药id = !配药id
                lblType(i).Tag = ""
                lblType(i).Caption = ""
                imgDrug(i).Tag = intSum
                IntSumMain = IntSumMain + 1
                i = i + 1
                blnChange = True
                
            Else
                blnChange = False
                If Me.lblDrug1.UBound + 1 < .RecordCount Then
                    Call LoadDrug(Me.lblDrug1.UBound + 1, i - 1, False)
                End If
            End If
            
            If Not (str药品id = !药品名称 And Not blnChange) Then
                str药品id = !药品名称
                intSum = intSum + 1
                intCount = intCount + 1
                Me.picDrug(i - 1).Tag = intSum - 1
                lblDrug1(Me.lblDrug1.UBound).Caption = !通用名
                lblSpec1(Me.lblSpec1.UBound).Caption = zlStr.nvl(!规格)
                lblAmount1(Me.lblAmount1.UBound).Caption = IIf(Val(!单量) = 0, "", Val(!单量))
                imgOtherDrug(i - 1).Visible = (intCount > 5)
                imgDrug(i - 1).Visible = (intCount > 5)
                imgOtherDrug(i - 1).Tag = intCount
            Else
                Unload lblDrug1(Me.lblDrug1.UBound)
                Unload Me.lblAmount1(Me.lblAmount1.UBound)
                Unload Me.lblSpec1(Me.lblSpec1.UBound)
            End If
            
            If InStr(1, lblType(i - 1).Tag, !配药类型) = 0 And !配药类型 <> "" Then
                lblType(i - 1).Tag = lblType(i - 1).Tag & !配药类型
                lblType(i - 1).Caption = lblType(i - 1).Caption & " 【" & !配药类型 & "】"
            End If
            
            mlng序号 = !组号
            
            If !执行标志 = 1 Then
                lblDrug1(i - 1).BackColor = M_CST_SELECTED_COLOR
                lblSpec1(i - 1).BackColor = M_CST_SELECTED_COLOR
                lblAmount1(i - 1).BackColor = M_CST_SELECTED_COLOR
            End If
            
            .MoveNext
        Loop
        
        If .EOF Then
            chkAll.Enabled = True
            Me.cmdNext.Enabled = False
            Exit Sub
        End If
        
        If mlng序号 - mintPage - (mIntCount * 3) + 1 < 1 Then
            Me.cmdPrevious.Enabled = False
        End If
        
'        If ((Me.picLableMain.UBound \ mIntCount) - 2) > 0 Then
''            Me.VSLBar.Visible = True
'            Me.VSLBar.LargeChange = Me.VSLBar.Max / ((Me.picLableMain.UBound \ mIntCount) - 2)
'            Me.VSLBar.SmallChange = Me.VSLBar.Max / ((Me.picLableMain.UBound \ mIntCount) - 2)
'        End If
        
    End With
    
    mrsData.Filter = strFilter
End Sub



Private Sub Show医嘱(Optional ByVal blnNext As Boolean)
    Dim i As Integer
    Dim j As Integer
    Dim lng医嘱id As Long
    Dim intCount As Integer
    Dim intSum As Integer
    Dim IntSumMain As Integer
    
    For j = 1 To Me.lblDrug1.UBound
        Unload lblDrug1(j)
        Unload Me.lblAmount1(j)
        Unload Me.lblSpec1(j)
    Next
    
    For j = 1 To Me.picLableMain.UBound
        Unload lblTime(j)
        Unload lblNo(j)
        Unload lblBed(j)
        Unload lblAge(j)
        Unload lblPatiName(j)
        Unload imgSex(j)
        Unload lblBatch(j)
        Unload imgPacker(j)
        Unload imgPrint(j)
        Unload lbldept(j)
        Unload imgCheck(j)
        Unload lblLine1(j)
        Unload lblLine2(j)
        Unload picDrug(j)
        Unload Me.imgDrug(j)
        Unload Me.imgOtherDrug(j)
        Unload lblType(j)
        Unload Me.picLable(j)
        Unload picLableMain(j)
    Next
    If blnNext Then
        mrsData.Filter = "组号>" & mlng序号 - 1
    Else
        mrsData.Filter = "组号>" & mlng序号 - mintPage - (mIntCount * 3)
    End If
    mintPage = 0
    With mrsData
        .Sort = "组号"
        
        Me.picLableMain(0).Visible = True
        Me.imgCheck(0).Picture = imgStandardUnCheck.Picture
        Me.picLableMain(0).BackColor = M_CST_SELECTED_COLOR
        '改变背景颜色
        Me.picLable(0).BackColor = vbWhite
        Me.picDrug(0).BackColor = vbWhite
        lbldept(0).BackColor = vbWhite
        lblBatch(0).BackColor = vbWhite
        lblPatiName(0).BackColor = vbWhite
        lblBed(0).BackColor = vbWhite
        lblAge(0).BackColor = vbWhite
        lblNo(0).BackColor = vbWhite
        lblTime(0).BackColor = vbWhite
        lblType(0).BackColor = vbWhite
        lblDrug1(0).BackColor = vbWhite
        lblSpec1(0).BackColor = vbWhite
        lblAmount1(0).BackColor = vbWhite
        
        Do While Not .EOF
            If lng医嘱id <> !相关ID Then
                mintPage = mintPage + 1
                If i = mIntCount * 3 Then
                    mlng序号 = !组号
                    chkAll.Enabled = True
                    If mlng序号 - mintPage - (mIntCount * 3) < 1 Then
                        Me.cmdPrevious.Enabled = False
                    End If
                    Exit Sub
                End If
                
                intCount = 0
                If i > Me.picLableMain.UBound Then
                    LoadTransLable (i)
                    Call LoadDrug(Me.lblDrug1.UBound + 1, i, True)
                    DoEvents
                End If
                
                lbldept(i).Caption = !科室
                lblBatch(i).Visible = False
                imgPacker(i).Visible = False
                Me.imgPrint(i).Visible = False
                lblPatiName(i).Caption = !姓名
                lblBed(i).Caption = IIf(IsNull(!床号), "", !床号)
                lblAge(i).Caption = !年龄
                lblNo(i).Caption = zlStr.nvl(!住院号)
                lblTime(i).Caption = !频次
                Me.picLableMain(i).Tag = !相关ID
                If !性别 = "女" Then
                    imgSex(i).Picture = Me.imgGirl.Picture
                Else
                    imgSex(i).Picture = imgBoy.Picture
                End If
                
                lng医嘱id = !相关ID
                lblType(i).Tag = ""
                lblType(i).Caption = ""
                imgDrug(i).Tag = intSum
                IntSumMain = IntSumMain + 1
                i = i + 1
                
            Else
                If Me.lblDrug1.UBound + 1 < .RecordCount Then
                    Call LoadDrug(Me.lblDrug1.UBound + 1, i - 1, False)
                End If
            End If
            
            intSum = intSum + 1
            intCount = intCount + 1
            Me.picDrug(i - 1).Tag = intSum - 1
            lblDrug1(Me.lblDrug1.UBound).Caption = !药品名称
            lblSpec1(Me.lblSpec1.UBound).Caption = !规格
            lblAmount1(Me.lblAmount1.UBound).Caption = IIf(Val(!单量) = 0, "", Val(!单量) & zlStr.nvl(!单位))
            imgOtherDrug(i - 1).Visible = (intCount > 5)
            imgDrug(i - 1).Visible = (intCount > 5)
            imgOtherDrug(i - 1).Tag = intCount
            
            mlng序号 = !组号
            .MoveNext
        Loop
        
        If .EOF Then
            chkAll.Enabled = True
            Me.cmdNext.Enabled = False
            Exit Sub
        End If
        
        If mlng序号 - mintPage - (mIntCount * 3) < 1 Then
            Me.cmdPrevious.Enabled = False
        End If
        
'        If ((Me.picLableMain.UBound \ mIntCount) - 2) > 0 Then
''            Me.VSLBar.Visible = True
'            Me.VSLBar.LargeChange = Me.VSLBar.Max / ((Me.picLableMain.UBound \ mIntCount) - 2)
'            Me.VSLBar.SmallChange = Me.VSLBar.Max / ((Me.picLableMain.UBound \ mIntCount) - 2)
'        End If
        
    End With
End Sub
Private Sub imgPacker_Click(index As Integer)
    picLable_Click index
End Sub

Private Sub imgPacker_DblClick(index As Integer)
    Dim strInput As String
    
    If mblnEdit = False Or mbln打包设置 = False Then Exit Sub
    
    If MsgBox("是否调整为" & IIf(Val(imgPacker(index).Tag) = 0, """打包""", """不打包""") & "状态？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    If Val(imgPacker(index).Tag) <> 2 Then
        imgPacker(index).Tag = 2
        imgPacker(index).Picture = imgStandardPacker(1).Picture
    Else
        imgPacker(index).Tag = 0
        imgPacker(index).Picture = imgStandardUnPacker.Picture
    End If
    
    strInput = Me.picLableMain(index).Tag & "," & imgPacker(index).Tag
    
    frmPIVAMain.PackMain Val(Me.picLableMain(index).Tag), imgPacker(index).Tag
                
    On Error GoTo errHandle
    If strInput <> "" Then
        gstrSQL = "Zl_输液配药记录_打包("
        '配药ID,打包
        gstrSQL = gstrSQL & "'" & strInput & "'"
        gstrSQL = gstrSQL & ")"
        Call zldatabase.ExecuteProcedure(gstrSQL, "打包设置")
    End If
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub imgPrint_Click(index As Integer)
    picLable_Click index
End Sub

Private Sub imgSex_Click(index As Integer)
    picLable_Click index
End Sub

Private Sub lblAge_Click(index As Integer)
    picLable_Click index
End Sub

Private Sub lblAmount1_Click(index As Integer)
    lblDrug1_Click index
End Sub

Private Sub lblBatch_Click(index As Integer)
    mintIndex = index
    picLable_Click index
    
    If mblnEdit = False Or mbln批次设置 = False Then Exit Sub
    Me.lst批次.Visible = True
    Me.lst批次.ZOrder
    Me.lst批次.Left = Me.lblBatch(index).Left + Me.picLable(index).Left + Me.picLableMain(index).Left + 200
    Me.lst批次.Top = Me.lblBatch(index).Top + Me.picLable(index).Top + Me.picLableMain(index).Top + Me.lblBatch(index).Height
    Me.lst批次.SetFocus
End Sub

'Private Sub lblBatch_DblClick(Index As Integer)
'    Dim strText As String
'    Dim strArr() As String
'    Dim i As Integer
'    Dim strInput As String
'
'    If mblnEdit = False Or mbln批次设置 = False Then Exit Sub
'
'    strText = Me.lblBatch(Index).Caption
'
'    strArr = Split(mstr批次, ";")
'    For i = 0 To UBound(strArr) - 1
'        If InStr(1, strText, "【" & strArr(i) & "#】") > 0 Then
'            If i + 1 > UBound(strArr) - 1 Then
'                If Me.lblBatch(Index).Tag = "【" & strArr(0) & "#】" Then
'                    Me.lblBatch(Index).Caption = "【" & strArr(0) & "#】"
'                    frmPIVAMain.ChangeBatchMain Val(Me.picLableMain(Index).Tag), strArr(0) & "#"
'                    strInput = Me.picLableMain(Index).Tag & "," & strArr(0)
'                Else
'                    If MsgBox("是否确认把批次由" & strText & "调整为" & "【" & strArr(0) & "#】" & "？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
'                        Me.lblBatch(Index).Caption = "【" & strArr(0) & "#】"
'                        frmPIVAMain.ChangeBatchMain Val(Me.picLableMain(Index).Tag), strArr(0) & "#"
'                        strInput = Me.picLableMain(Index).Tag & "," & strArr(0)
'                    End If
'                End If
'            Else
'                If Me.lblBatch(Index).Tag = "【" & strArr(i + 1) & "#】" Then
'                    Me.lblBatch(Index).Caption = "【" & strArr(i + 1) & "#】"
'                    frmPIVAMain.ChangeBatchMain Val(Me.picLableMain(Index).Tag), strArr(i + 1) & "#"
'                    strInput = Me.picLableMain(Index).Tag & "," & strArr(i + 1)
'                Else
'                    If MsgBox("是否确认把批次由" & strText & "调整为" & "【" & strArr(i + 1) & "#】" & "？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
'                        Me.lblBatch(Index).Caption = "【" & strArr(i + 1) & "#】"
'                        frmPIVAMain.ChangeBatchMain Val(Me.picLableMain(Index).Tag), strArr(i + 1) & "#"
'                        strInput = Me.picLableMain(Index).Tag & "," & strArr(i + 1)
'                    End If
'                End If
'            End If
'        End If
'    Next
'
'    If strInput <> "" Then
'        gstrSQL = "Zl_输液配药记录_分批("
'        '配药ID,批次
'        gstrSQL = gstrSQL & "'" & strInput & "'"
'        gstrSQL = gstrSQL & ")"
'        Call zlDatabase.ExecuteProcedure(gstrSQL, "设置批次")
'    End If
'End Sub

Private Sub lblBed_Click(index As Integer)
    picLable_Click index
End Sub

Private Sub lblDept_Click(index As Integer)
    picLable_Click index
End Sub

Private Sub lblDrug1_Click(index As Integer)
    Dim strTemp As String
    strTemp = Me.lblDrug1(index).Container.index
    picLable_Click Val(strTemp)
End Sub

Private Sub lblLine1_Click(index As Integer)
    picLable_Click index
End Sub

Private Sub lblLine2_Click(index As Integer)
    picLable_Click index
End Sub

Private Sub lblNo_Click(index As Integer)
    picLable_Click index
End Sub

Private Sub lblPatiName_Click(index As Integer)
    picLable_Click index
End Sub

Private Sub lblSpec1_Click(index As Integer)
    lblDrug1_Click index
End Sub

Private Sub lblTime_Click(index As Integer)
    picLable_Click index
End Sub

Private Sub lblType_Click(index As Integer)
     picLable_Click index
End Sub

Private Sub lst批次_DblClick()
    Dim strTemp As String
    Dim strInput As String
    Dim strLevel As String
    
    On Error GoTo errHandle
    strTemp = "【" & Mid(Me.lst批次.SelectedItem.Tag, 1, InStr(1, Me.lst批次.SelectedItem.Tag, "#")) & "】"
    Me.lst批次.Visible = False
    If strTemp <> Me.lblBatch(mintIndex).Tag Then
        If MsgBox("是否确认把批次由" & Me.lblBatch(mintIndex).Caption & "调整为" & strTemp & "？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Sub
        End If
    End If
    
    If (mstrStep = "00" Or mstrStep = "10" Or mstrStep = "11") And mbln审核 Then
    Else
        mrsData.Filter = "配药id=" & Val(Me.picLableMain(mintIndex).Tag)
        If mrsData.RecordCount > 0 Then
            strLevel = mrsData!优先级
        End If
    End If
    
    Me.lblBatch(mintIndex).Caption = strTemp
    lblBatch(mintIndex).ForeColor = Me.lst批次.SelectedItem.SubItems(1)
    frmPIVAMain.ChangeBatchMain Val(Me.picLableMain(mintIndex).Tag), Mid(Me.lst批次.SelectedItem.Tag, 1, InStr(1, Me.lst批次.SelectedItem.Tag, "#"))
    
    strInput = Me.picLableMain(mintIndex).Tag & "," & Mid(Me.lst批次.SelectedItem.Tag, 1, InStr(1, Me.lst批次.SelectedItem.Tag, "#") - 1) & ":" & strLevel
    If strInput <> "" Then
        gstrSQL = "Zl_输液配药记录_分批("
        '配药ID,批次
        gstrSQL = gstrSQL & "'" & strInput & "'"
        gstrSQL = gstrSQL & ")"
        Call zldatabase.ExecuteProcedure(gstrSQL, "设置批次")
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub lst批次_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        Me.lst批次.Visible = False
    ElseIf KeyCode = 13 Then
        lst批次_DblClick
    End If
End Sub

Private Sub lst批次_LostFocus()
    Me.lst批次.Visible = False
End Sub

Private Sub picDrug_Click(index As Integer)
    picLable_Click index
End Sub

'Private Sub picDrug_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
'    Dim blnMouse As Boolean
'
'    If Me.picLableMain(Index).BackColor = vbRed Then Exit Sub
'    blnMouse = (0 <= x) And (x <= picDrug(Index).Width) And (0 <= y) And (y <= picDrug(Index).Height)
'    If blnMouse Then
'        picLableMain(Index).BackColor = &HFF8080
'        SetCapture picDrug(Index).hWnd
'    Else
'        picLableMain(Index).BackColor = M_CST_SELECTED_COLOR&
'        ReleaseCapture
'    End If
'End Sub

Private Sub picLable_Click(index As Integer)
    Dim i As Integer
    
    Me.picLableMain(index).BackColor = vbRed
    
    For i = 0 To Me.picLableMain.UBound
        If i <> index And Val(imgCheck(i).Tag) <> 1 Then
            Me.picLableMain(i).BackColor = M_CST_SELECTED_COLOR
        End If
    Next
End Sub

'Private Sub picLable_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
'    Dim blnMouse As Boolean
'
'    If Me.picLableMain(Index).BackColor = vbRed Then Exit Sub
'    blnMouse = (0 <= x) And (x <= picLable(Index).Width) And (0 <= y) And (y <= picLable(Index).Height)
'    If blnMouse Then
'        picLableMain(Index).BackColor = &HFF8080
'        SetCapture picLable(Index).hWnd
'    Else
'        picLableMain(Index).BackColor = &HE0E0E0
'        ReleaseCapture
'    End If
'
'End Sub

Private Sub picLable_MouseUp(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim objPopup As CommandBarPopup
    
'    If mblnEdit = False Then Exit Sub
    If Button = 1 Then Exit Sub
    
    mInt选择 = index
    
    frmPIVAMain.Get配药id Val(Me.picLableMain(index).Tag)
    
    Set objPopup = frmPIVAMain.cbsMain.ActiveMenuBar.Controls.Find(xtpControlPopup, 300)
    If Not objPopup Is Nothing Then
        objPopup.CommandBar.ShowPopup
    End If
End Sub

Private Sub txtFind_GotFocus()
    Call zlControl.TxtSelAll(txtFind)
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    frmPIVAMain.SetTxtFind Me.txtFind.Text, KeyAscii
End Sub

Private Sub VSLBar_Change()
    Dim lngMove As Long
    Dim i As Integer
    Dim intRow As Integer
    
'    数据处理前调:
    Call LockWindowUpdate(Me.hWnd)

    lngMove = CLng(VSLBar.Value)
    lngMove = lngMove / Me.VSLBar.SmallChange * (Me.picLableMain(i).Height + 400)
    
    For i = 0 To Me.picLableMain.UBound
        Me.picLableMain(i).Top = Me.picLableMain(i).Top - (lngMove - mIntOld)
    Next
    
    mIntOld = lngMove
    '处理完后:
    Call LockWindowUpdate(0)
End Sub

Public Sub ClearCard()
    Dim i As Integer
    
    For i = 0 To Me.picLableMain.UBound
        Me.picLableMain(i).Visible = False
    Next
End Sub

Public Sub LoadHelp(ByVal strMsg As String)
    Me.lblHelp.Caption = strMsg
End Sub

Public Sub CheckType(ByVal intIndex As Integer, ByVal intValue As Integer)
    Me.chkType(intIndex).Value = intValue
End Sub

Public Function ChooseOne() As String
    imgCheck(mInt选择).Tag = 1
    imgCheck(mInt选择).Picture = imgStandardCheck.Picture
    
    ChooseOne = Val(Me.picLableMain(mInt选择).Tag)
End Function

Public Sub chkClick(ByVal intValue As Integer)
    mint标志 = 1
    Me.chkAll.Value = intValue
    Chk_all
End Sub

Public Sub ChooseOneRec(ByVal str配药id As String, ByVal intValue As Integer)
    Dim i As Integer
    Dim j As Integer
    
    For i = 0 To Me.lblNo.UBound
        If Me.picLableMain(i).Tag = str配药id Then
            If intValue <> 0 Then
                imgCheck(i).Tag = intValue
                
                If intValue = 1 Then
                    imgCheck(i).Picture = imgStandardCheck.Picture
                Else
                    imgCheck(i).Picture = imgUnCheck.Picture
                End If
                
                Me.picLable(i).BackColor = M_CST_SELECTED_COLOR
                Me.picDrug(i).BackColor = M_CST_SELECTED_COLOR
                lbldept(i).BackColor = M_CST_SELECTED_COLOR
                lblBatch(i).BackColor = M_CST_SELECTED_COLOR
                lblPatiName(i).BackColor = M_CST_SELECTED_COLOR
                lblBed(i).BackColor = M_CST_SELECTED_COLOR
                lblAge(i).BackColor = M_CST_SELECTED_COLOR
                lblNo(i).BackColor = M_CST_SELECTED_COLOR
                lblTime(i).BackColor = M_CST_SELECTED_COLOR
                lblType(i).BackColor = M_CST_SELECTED_COLOR
                
                For j = CInt(IIf(Me.picDrug(i).Tag = "", "0", Me.picDrug(i).Tag)) - CInt(IIf(imgOtherDrug(i).Tag = "", "0", imgOtherDrug(i).Tag)) + 1 To CInt(IIf(Me.picDrug(i).Tag = "", "0", Me.picDrug(i).Tag))
                    lblDrug1(j).BackColor = M_CST_SELECTED_COLOR
                    lblSpec1(j).BackColor = M_CST_SELECTED_COLOR
                    lblAmount1(j).BackColor = M_CST_SELECTED_COLOR
                Next
                Me.picLableMain(i).BackColor = vbRed
            Else
                imgCheck(i).Tag = 0
                imgCheck(i).Picture = imgStandardUnCheck.Picture
                
                Me.picLable(i).BackColor = vbWhite
                Me.picDrug(i).BackColor = vbWhite
                lbldept(i).BackColor = vbWhite
                lblBatch(i).BackColor = vbWhite
                lblPatiName(i).BackColor = vbWhite
                lblBed(i).BackColor = vbWhite
                lblAge(i).BackColor = vbWhite
                lblNo(i).BackColor = vbWhite
                lblTime(i).BackColor = vbWhite
                lblType(i).BackColor = vbWhite
                
                For j = CInt(Me.picDrug(i).Tag) - CInt(imgOtherDrug(i).Tag) + 1 To CInt(Me.picDrug(i).Tag)
                    lblDrug1(j).BackColor = vbWhite
                    lblSpec1(j).BackColor = vbWhite
                    lblAmount1(j).BackColor = vbWhite
                Next
                Me.picLableMain(i).BackColor = M_CST_SELECTED_COLOR
            End If
        End If
    Next
End Sub

Public Sub BatchPrint(ByVal str配药id As String)
    Dim i As Integer
    
    For i = 0 To Me.lblNo.UBound
        If InStr(1, str配药id, ";" & Me.picLableMain(i).Tag & ";") <> 0 Then
            imgPrint(i).Tag = 1
            imgPrint(i).Picture = imgStandardPrint.Picture
        End If
    Next
End Sub

Public Sub BatchChoose(ByVal str配药id As String)
'*************************************************
'控制两种模式的选择同步
'*************************************************
    Dim i As Integer
    Dim j As Integer
    
    For i = 0 To Me.lblNo.UBound
        If InStr(1, str配药id, ";" & Me.picLableMain(i).Tag & ";") > 0 Then
            imgCheck(i).Tag = 1
            imgCheck(i).Picture = imgStandardCheck.Picture
            Me.picLable(i).BackColor = M_CST_SELECTED_COLOR
            Me.picDrug(i).BackColor = M_CST_SELECTED_COLOR
            lbldept(i).BackColor = M_CST_SELECTED_COLOR
            lblBatch(i).BackColor = M_CST_SELECTED_COLOR
            lblPatiName(i).BackColor = M_CST_SELECTED_COLOR
            lblBed(i).BackColor = M_CST_SELECTED_COLOR
            lblAge(i).BackColor = M_CST_SELECTED_COLOR
            lblNo(i).BackColor = M_CST_SELECTED_COLOR
            lblTime(i).BackColor = M_CST_SELECTED_COLOR
            lblType(i).BackColor = M_CST_SELECTED_COLOR
            
            For j = CInt(Me.picDrug(i).Tag) - CInt(imgOtherDrug(i).Tag) + 1 To CInt(Me.picDrug(i).Tag)
                lblDrug1(j).BackColor = M_CST_SELECTED_COLOR
                lblSpec1(j).BackColor = M_CST_SELECTED_COLOR
                lblAmount1(j).BackColor = M_CST_SELECTED_COLOR
            Next
        End If
    Next
End Sub

Public Sub PackCard(ByVal lng配药id As Long, ByVal intValue As Integer)
'同步打包数据
    Dim i As Integer

    For i = 0 To Me.lblNo.UBound
        If Val(Me.picLableMain(i).Tag) = lng配药id Then
            If intValue = 2 Then
                imgPacker(i).Tag = 2
                imgPacker(i).Picture = imgStandardPacker(1).Picture
            Else
                imgPacker(i).Tag = 0
                imgPacker(i).Picture = imgStandardUnPacker.Picture
            End If
            Exit For
        End If
    Next
End Sub

Public Sub Changebatch(ByVal lng配药id As Long, ByVal strValue As String)
'同步批次数据
    Dim i As Integer
    Dim lngColor As Long
    
    For i = 1 To Me.lst批次.ListItems.count
        If strValue = Mid(Me.lst批次.ListItems(i).Tag, 1, InStr(1, Me.lst批次.ListItems(i).Tag, "#")) Then
            lngColor = Me.lst批次.ListItems(i).SubItems(1)
            Exit For
        End If
    Next

    For i = 0 To Me.lblNo.UBound
        If Val(Me.picLableMain(i).Tag) = lng配药id Then
            lblBatch(i).Caption = "【" & strValue & "】"
            lblBatch(i).ForeColor = lngColor
            Exit For
        End If
    Next
End Sub

Public Sub GetForce(ByVal lng配药id As Long)
'******************************************************
'查找瓶签信息，调整滚动条的位置
'******************************************************
    Dim i As Integer
    Dim lngMove As Long
    
    For i = 0 To Me.picLableMain.UBound
        If Me.picLableMain(i).Tag = lng配药id Then
            '控制纵向滚动条的移动
            Do While Me.picLableMain(i).Top > Me.HSc.Top - 1200
                lngMove = (Me.picLableMain(i).Top - (Me.HSc.Top - 5000) + mIntOld) / (Me.picLableMain(i).Height + 400) * Me.VSLBar.SmallChange
                VSLBar.Value = IIf(lngMove > VSLBar.Max, VSLBar.Max, lngMove)
            Loop
            
            Do While Me.picLableMain(i).Top < 0
                lngMove = (Me.picLableMain(i).Top - 2500 + mIntOld) / (Me.picLableMain(i).Height + 400) * Me.VSLBar.SmallChange
                VSLBar.Value = IIf(lngMove < 0, 0, lngMove)
            Loop
            
            picLable_Click i
        End If
    Next
End Sub
Private Sub Load批次()
'**********************************************
'加载批次信息
'**********************************************
    Dim i As Integer
    Dim objItem  As listItem
    
    Me.lst批次.ListItems.Clear
    For i = 0 To UBound(Split(mstr批次, "/"))
        Set objItem = Me.lst批次.ListItems.Add(, "_" & i, "批次")
        objItem.Text = Mid(Split(mstr批次, "/")(i), 1, InStr(1, Split(mstr批次, "/")(i), ",") - 1)
        objItem.Tag = Split(Split(mstr批次, "/")(i), " ")(0)
        objItem.SubItems(1) = Mid(Split(mstr批次, "/")(i), InStr(1, Split(mstr批次, "/")(i), ",") + 1)
    Next
End Sub




