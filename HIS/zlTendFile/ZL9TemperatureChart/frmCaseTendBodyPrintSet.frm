VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCaseTendBodyPrintSet 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "体温单打印"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6840
   Icon            =   "frmCaseTendBodyPrintSet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   6840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin TabDlg.SSTab SSTab1 
      Height          =   5130
      Left            =   30
      TabIndex        =   3
      Top             =   90
      Width           =   6765
      _ExtentX        =   11933
      _ExtentY        =   9049
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "常规打印"
      TabPicture(0)   =   "frmCaseTendBodyPrintSet.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblIn"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "picInfo"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "picCHKH"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "打印选项"
      TabPicture(1)   =   "frmCaseTendBodyPrintSet.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fra"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "fra打印"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.PictureBox picCHKH 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4365
         Left            =   15
         ScaleHeight     =   4365
         ScaleWidth      =   6690
         TabIndex        =   30
         Top             =   705
         Visible         =   0   'False
         Width           =   6690
         Begin VB.PictureBox picVsh 
            Appearance      =   0  'Flat
            BackColor       =   &H80000003&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1080
            Left            =   360
            ScaleHeight     =   1080
            ScaleWidth      =   1635
            TabIndex        =   32
            Top             =   1080
            Width           =   1635
            Begin VB.PictureBox picPrint 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   555
               Index           =   0
               Left            =   0
               ScaleHeight     =   555
               ScaleWidth      =   480
               TabIndex        =   33
               Top             =   -15
               Visible         =   0   'False
               Width           =   480
               Begin VB.Label lblNum 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  Caption         =   "9999"
                  BeginProperty Font 
                     Name            =   "宋体"
                     Size            =   9
                     Charset         =   134
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   180
                  Index           =   0
                  Left            =   30
                  TabIndex        =   34
                  Top             =   315
                  Visible         =   0   'False
                  Width           =   435
               End
               Begin VB.Image imgIco 
                  Height          =   240
                  Index           =   0
                  Left            =   120
                  Picture         =   "frmCaseTendBodyPrintSet.frx":0044
                  Top             =   45
                  Width           =   240
               End
            End
         End
         Begin VB.VScrollBar vsc 
            Height          =   1815
            Left            =   4725
            SmallChange     =   50
            TabIndex        =   31
            Top             =   60
            Visible         =   0   'False
            Width           =   200
         End
      End
      Begin VB.PictureBox picInfo 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   75
         ScaleHeight     =   270
         ScaleWidth      =   6525
         TabIndex        =   24
         Top             =   375
         Width           =   6525
         Begin VB.OptionButton optSelect 
            Caption         =   "待打"
            Height          =   180
            Index           =   0
            Left            =   0
            TabIndex        =   29
            Top             =   15
            Width           =   660
         End
         Begin VB.OptionButton optSelect 
            Caption         =   "已打"
            Height          =   180
            Index           =   1
            Left            =   930
            TabIndex        =   28
            Top             =   15
            Width           =   660
         End
         Begin VB.OptionButton optSelect 
            Caption         =   "重打"
            Height          =   180
            Index           =   2
            Left            =   1860
            TabIndex        =   27
            Top             =   15
            Width           =   660
         End
         Begin VB.OptionButton optSelect 
            Caption         =   "所有"
            Height          =   180
            Index           =   3
            Left            =   2805
            TabIndex        =   26
            Top             =   15
            Value           =   -1  'True
            Width           =   660
         End
         Begin VB.CheckBox chkSelect 
            Caption         =   "全选"
            Height          =   180
            Left            =   5910
            TabIndex        =   25
            Top             =   15
            Width           =   705
         End
         Begin VB.Image imgIcon 
            DragIcon        =   "frmCaseTendBodyPrintSet.frx":05CE
            Height          =   240
            Index           =   0
            Left            =   660
            MouseIcon       =   "frmCaseTendBodyPrintSet.frx":0CB8
            Picture         =   "frmCaseTendBodyPrintSet.frx":1242
            Top             =   -15
            Width           =   240
         End
         Begin VB.Image imgIcon 
            DragIcon        =   "frmCaseTendBodyPrintSet.frx":17CC
            Height          =   240
            Index           =   1
            Left            =   1575
            MouseIcon       =   "frmCaseTendBodyPrintSet.frx":1EB6
            Picture         =   "frmCaseTendBodyPrintSet.frx":2440
            Top             =   -15
            Width           =   240
         End
         Begin VB.Image imgIcon 
            DragIcon        =   "frmCaseTendBodyPrintSet.frx":29CA
            Height          =   240
            Index           =   2
            Left            =   2505
            MouseIcon       =   "frmCaseTendBodyPrintSet.frx":30B4
            Picture         =   "frmCaseTendBodyPrintSet.frx":363E
            Top             =   -15
            Width           =   240
         End
      End
      Begin VB.Frame fra打印 
         Caption         =   "打印页脚"
         Height          =   1080
         Left            =   -74760
         TabIndex        =   8
         Top             =   420
         Width           =   4380
         Begin VB.PictureBox picBack 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   2130
            Left            =   2850
            ScaleHeight     =   491.128
            ScaleMode       =   0  'User
            ScaleWidth      =   491.128
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   2175
            Visible         =   0   'False
            Width           =   2130
            Begin VB.PictureBox picShadow 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H00808080&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   1485
               Left            =   450
               ScaleHeight     =   1485
               ScaleWidth      =   1170
               TabIndex        =   19
               TabStop         =   0   'False
               Top             =   315
               Width           =   1170
            End
            Begin VB.PictureBox picPaper 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H00FFFFFF&
               FillColor       =   &H00C0C0C0&
               FillStyle       =   0  'Solid
               ForeColor       =   &H80000008&
               Height          =   1485
               Left            =   405
               ScaleHeight     =   1455
               ScaleMode       =   0  'User
               ScaleWidth      =   1140
               TabIndex        =   17
               TabStop         =   0   'False
               ToolTipText     =   "拖动蓝色线条改变起始位置"
               Top             =   270
               Width           =   1170
               Begin VB.PictureBox pic起始 
                  Appearance      =   0  'Flat
                  AutoRedraw      =   -1  'True
                  BackColor       =   &H00FF0000&
                  BorderStyle     =   0  'None
                  ForeColor       =   &H80000008&
                  Height          =   15
                  Left            =   0
                  MousePointer    =   7  'Size N S
                  ScaleHeight     =   15
                  ScaleMode       =   0  'User
                  ScaleWidth      =   1140
                  TabIndex        =   18
                  TabStop         =   0   'False
                  Top             =   135
                  Width           =   1140
               End
            End
         End
         Begin VB.TextBox txt页号 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   3450
            MaxLength       =   3
            TabIndex        =   15
            Text            =   "1"
            Top             =   360
            Width           =   285
         End
         Begin VB.TextBox txt起始 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   1050
            MaxLength       =   3
            TabIndex        =   12
            Text            =   "25"
            Top             =   1680
            Visible         =   0   'False
            Width           =   600
         End
         Begin VB.CheckBox chk页号 
            Caption         =   "打印页号，第一页页号表示为(&4)"
            Height          =   195
            Left            =   525
            TabIndex        =   11
            Top             =   405
            Value           =   1  'Checked
            Width           =   2910
         End
         Begin VB.CheckBox chk周数 
            Caption         =   "打印住院周数(&5)"
            Height          =   195
            Left            =   525
            TabIndex        =   10
            Top             =   765
            Value           =   1  'Checked
            Width           =   1650
         End
         Begin VB.CheckBox chkOper 
            Caption         =   "打印打印人(&6)"
            Height          =   195
            Left            =   2625
            TabIndex        =   9
            Top             =   765
            Value           =   1  'Checked
            Width           =   1650
         End
         Begin MSComCtl2.UpDown UD页号 
            Height          =   300
            Left            =   3735
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   345
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   529
            _Version        =   393216
            Value           =   1
            BuddyControl    =   "txt页号"
            BuddyDispid     =   196624
            OrigLeft        =   1590
            OrigTop         =   1365
            OrigRight       =   1830
            OrigBottom      =   1665
            Max             =   999
            Min             =   1
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown UD起始 
            Height          =   300
            Left            =   1665
            TabIndex        =   14
            Top             =   1680
            Visible         =   0   'False
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Value           =   25
            BuddyControl    =   "txt起始"
            BuddyDispid     =   196625
            OrigLeft        =   1590
            OrigTop         =   705
            OrigRight       =   1830
            OrigBottom      =   1005
            Max             =   460
            Min             =   1
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "mm"
            Height          =   180
            Left            =   1965
            TabIndex        =   21
            Top             =   1710
            Visible         =   0   'False
            Width           =   180
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "起始位置"
            Height          =   180
            Left            =   255
            TabIndex        =   20
            Top             =   1740
            Visible         =   0   'False
            Width           =   720
         End
      End
      Begin VB.Frame fra 
         Caption         =   "其他"
         Height          =   1440
         Left            =   -74775
         TabIndex        =   4
         Top             =   1560
         Width           =   4380
         Begin VB.CheckBox chk 
            Caption         =   "不打印体温单下方的曲线说明信息(&9)"
            Height          =   195
            Index           =   1
            Left            =   900
            TabIndex        =   22
            Top             =   1080
            Value           =   1  'Checked
            Width           =   3360
         End
         Begin VB.TextBox txt 
            Height          =   300
            Left            =   960
            TabIndex        =   6
            Top             =   255
            Width           =   3210
         End
         Begin VB.CheckBox chk 
            Caption         =   "不打印心率和脉搏间的连线和阴影(&8)"
            Height          =   195
            Index           =   0
            Left            =   915
            TabIndex        =   5
            Top             =   705
            Value           =   1  'Checked
            Width           =   3360
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "质控号(&7)"
            Height          =   180
            Left            =   120
            TabIndex        =   7
            Top             =   300
            Width           =   810
         End
      End
      Begin VB.Label lblIn 
         Caption         =   "共11页"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   2025
         TabIndex        =   36
         Top             =   60
         Width           =   855
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "双击选中或取消(按SHIFT点击卡片范围选择)"
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   3180
         TabIndex        =   35
         Top             =   30
         Width           =   3510
      End
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "打印(&P)"
      Height          =   350
      Left            =   90
      TabIndex        =   0
      Top             =   5310
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5655
      TabIndex        =   1
      Top             =   5310
      Width           =   1100
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "预览(&V)"
      Height          =   350
      Left            =   90
      TabIndex        =   2
      Top             =   5310
      Width           =   1100
   End
   Begin VB.Label lblSelect 
      Caption         =   "已选页码"
      Height          =   180
      Left            =   1290
      TabIndex        =   23
      Top             =   5400
      Width           =   4125
   End
End
Attribute VB_Name = "frmCaseTendBodyPrintSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mbytOpt As Byte

Private mlng文件ID As Long
Private mlngAllPage As Long
Private mintPrintRange As Integer
Private mstrPage As String '选择连续打印时记录开始也和结束页号
Private mlngBeginY As Long
Private mintBeginPage As Integer
Private mlngWidth As Long '自定义纸张宽度,Twip
Private mlngHeight As Long '自定义纸张高度'Twip
Private mlngLeft As Long '左边距'mm
Private mlngRight As Long '右边距'mm
Private mlngTop As Long '上边距'mm
Private mlngBottom As Long '下边距'mm
Private mblnInit As Boolean
Private mblnShift As Boolean

Private mstrPrivs As String

Private Sub chkSelect_Click()
    Dim i As Long
    For i = 0 To picPrint.Count - 1
        If picPrint(i).Visible = True Then
            picPrint(i).BackColor = IIf(chkSelect.Value = 0, &HE0E0E0, vbRed)
            picPrint(i).Tag = IIf(chkSelect.Value = 0, "0", "1")
        End If
    Next
    Call SetSelectInfo
End Sub

Private Sub chk页号_Click()
    txt页号.Enabled = chk页号.Value = 1
    UD页号.Enabled = chk页号.Value = 1
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPreview_Click()
    If Not GetValue Then Exit Sub
    mbytOpt = 1
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    Call zlDatabase.SetPara("质控号", txt.Text, glngSys, 1255, InStr(mstrPrivs, ";护理选项设置;") > 0)
    If Not GetValue Then Exit Sub
    mbytOpt = 2
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 39 Then KeyCode = 0
    If Shift = 1 And KeyCode = vbKeyShift Then
        mblnShift = True
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyShift Then
        mblnShift = False
    End If
End Sub

Private Sub Form_Load()
    mbytOpt = 0
    mblnShift = False
    
    '显示纸张打印位置调整图
    
    mlngWidth = Val(zlDatabase.GetPara("体温单宽度", glngSys, 1255, Printer.Width))
    mlngHeight = Val(zlDatabase.GetPara("体温单高度", glngSys, 1255, Printer.Height))
    mlngLeft = Val(zlDatabase.GetPara("体温单左边距", glngSys, 1255, OFFSET_LEFT))
    mlngRight = Val(zlDatabase.GetPara("体温单右边距", glngSys, 1255, OFFSET_RIGHT))
    mlngTop = Val(zlDatabase.GetPara("体温单上边距", glngSys, 1255, OFFSET_TOP))
    mlngBottom = Val(zlDatabase.GetPara("体温单下边距", glngSys, 1255, OFFSET_BOTTOM))
    
    txt.Text = zlDatabase.GetPara("质控号", glngSys, 1255, "", Array(txt), InStr(mstrPrivs, "护理选项设置") > 0)
    
    If mlngWidth > mlngHeight Then
        picBack.ScaleWidth = mlngWidth / conRatemmToTwip * 1.1
        picBack.ScaleHeight = mlngWidth / conRatemmToTwip * 1.1
    Else
        picBack.ScaleWidth = mlngHeight / conRatemmToTwip * 1.1
        picBack.ScaleHeight = mlngHeight / conRatemmToTwip * 1.1
    End If
    picPaper.Width = mlngWidth / conRatemmToTwip
    picPaper.Height = mlngHeight / conRatemmToTwip
    picPaper.Left = (picBack.ScaleWidth - picPaper.Width) / 2
    picPaper.Top = (picBack.ScaleHeight - picPaper.Height) / 2
    picShadow.Width = picPaper.Width
    picShadow.Height = picPaper.Height
    picShadow.Left = picPaper.Left + 5
    picShadow.Top = picPaper.Top + 5
    
    picPaper.ScaleWidth = mlngWidth / conRatemmToTwip
    picPaper.ScaleHeight = mlngHeight / conRatemmToTwip
    
    '显初始位置
    If Not (mlngBeginY >= mlngTop And mlngBeginY <= picPaper.ScaleHeight - mlngBottom * 2) Then
        mlngBeginY = mlngTop
    End If
    pic起始.Left = 0
    pic起始.Width = picPaper.ScaleWidth
    pic起始.Top = mlngBeginY
    
    UD起始.Min = mlngTop
    UD起始.Max = picPaper.ScaleHeight - 2 * mlngBottom
    UD起始.Value = mlngBeginY
    
    pic起始.ScaleHeight = 1 '不然不能拖动
    
    Call DrawPage
    
    mintPrintRange = Val(zlDatabase.GetPara("连续打印", glngSys, 1255, "1"))
    
    chk页号.Value = Val(zlDatabase.GetPara("打印页号", glngSys, 1255, "1", Array(chk页号)))
    txt页号.Text = Val(zlDatabase.GetPara("起始页号", glngSys, 1255, "1", Array(txt页号, UD页号)))
    chk周数.Value = Val(zlDatabase.GetPara("打印周数", glngSys, 1255, "0", Array(chk周数)))
    '67405:刘鹏飞,2013-11-25
    chkOper.Value = Val(zlDatabase.GetPara("打印打印人", glngSys, 1255, "0", Array(chkOper)))
    chk(0).Value = Val(zlDatabase.GetPara("不打印脉搏短绌图形", glngSys, 1255, "0", Array(chk(0))))
    
    mintBeginPage = Val(txt页号.Text)
    
    UD页号.Value = IIf(mintBeginPage = 0, 1, mintBeginPage)

End Sub

Public Function PrintSet(objParent As Object, ByVal strParam As String, ByRef intPrintRange As Integer, ByRef lngBeginY As Long, ByRef intBeginPage As Integer, strPage As String, ByVal strPrivs As String, ByVal bytMode As Byte) As Byte
'功能：调用打印选项
'参数：blnFirst=是否第一次调用,否则只有"确定","取消",且不允许修改病历打印份数
'      strParam 由当前页连续打印是 需要提取 文件ID;病人体温单总页数
'      blnCurCase=T=只打印当前病历,F=从当前病历开始连续打印病历
'      lngBeginY=本次病历开始打印位置'mm
'      intBeginPage=起始页号,为0表示不打印页号
'      strPage
'返回：0-取消,1-预览,2-打印
    
    mstrPrivs = strPrivs
    
    If strParam <> "" Then
        If InStr(1, strParam, ";") = 0 Then
            mlng文件ID = Val(strParam)
        Else
            mlng文件ID = Val(Split(strParam, ";")(0))
            mlngAllPage = Val(Split(strParam, ";")(1))
        End If
    End If
    mintPrintRange = intPrintRange
    mlngBeginY = lngBeginY
    mintBeginPage = intBeginPage
    mblnInit = True
    cmdPrint.Visible = (bytMode = 1)
    cmdPreview.Visible = (bytMode = 2)
    lblIn.Caption = "共" & mlngAllPage & "页"
    
    Call GetPageNum(mlng文件ID)
    mblnInit = False
    Me.Show 1, objParent
    
    intPrintRange = mintPrintRange
    lngBeginY = mlngBeginY
    intBeginPage = mintBeginPage
    strPage = mstrPage
    PrintSet = mbytOpt
End Function

Public Function GetPageNum(ByVal lng文件ID As Long) As Boolean
'------------------------------------------------
'提取打印页号
'------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    Dim lngPage As Long
    
    On Error GoTo Errhand
    lngPage = 1
    Call LoadPages
    strSQL = "select 打印页号,打印人 From 体温单打印 where 文件ID=[1] Order by 打印页号"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "提取打印数据", lng文件ID)
    
    Do While Not rsTemp.EOF
        lngPage = Val("" & rsTemp!打印页号)
        If lngPage > 0 And lngPage <= picPrint.Count Then
            If "" & rsTemp!打印人 = "" Then
                imgIco(lngPage - 1).Picture = imgIcon(2).Picture
                imgIco(lngPage - 1).Tag = "重打"
            Else
                imgIco(lngPage - 1).Picture = imgIcon(1).Picture
                imgIco(lngPage - 1).Tag = "已打"
            End If
        End If
        rsTemp.MoveNext
    Loop
    
    GetPageNum = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub LoadPages()
    Dim i As Long
    
    picCHKH.Visible = mlngAllPage > 0
   
    For i = 0 To mlngAllPage - 1
        If i > 0 Then
            Load picPrint(i)
            Load imgIco(i)
            imgIco(i).Left = imgIco(0).Left
            imgIco(i).Top = imgIco(0).Top
            Set imgIco(i).Container = picPrint(i)
            Load lblNum(i)
            lblNum(i).Left = lblNum(0).Left
            lblNum(i).Top = lblNum(0).Top
            Set lblNum(i).Container = picPrint(i)
        End If
        imgIco(i).Tag = "待打"
        imgIco(i).Visible = True
        imgIco(i).ZOrder 0
        lblNum(i).Visible = True
        lblNum(i).ZOrder 0
        lblNum(i).Caption = i + 1
        picPrint(i).Tag = "0"
        picPrint(i).BackColor = &HE0E0E0
    Next
    Call ShowCard
End Sub

Private Sub ShowCard(Optional ByVal strTag As String = "ALL")
    '显示并排列卡片位置
    Dim i As Long, j As Long
    Dim lngLeft As Long, lngTop As Long
    Dim lngwNum As Long, lngHnum As Long
    Dim lngPageCount As Long, lngHeight As Long
    Dim lngCHeight As Long
    
    Call LockWindowUpdate(Me.hWnd)
    
    With picVsh
        .Left = 0
        .Top = 0
        .Width = picCHKH.Width
        .Height = picCHKH.Height
        .BackColor = picCHKH.BackColor
    End With
    lngHeight = picVsh.Height
    vsc.Value = 0
    vsc.Visible = False
    For i = 0 To picPrint.Count - 1
        If imgIco(i).Tag = strTag Or strTag = "ALL" Then
            lngPageCount = lngPageCount + 1
        End If
    Next
     '-----计算是否显示滚动条(卡片区域固定)
    '计算每页可显示的卡片数目
    lngwNum = picVsh.Width \ (picPrint(0).Width + 60)
    '计算高度是否超出区域
    lngHnum = (lngPageCount \ lngwNum)
    If lngwNum * lngHnum < lngPageCount Then lngHnum = lngHnum + 1
    If lngHnum * (picPrint(0).Height + 60) > picVsh.Height Then
        lngHeight = lngHnum * (picPrint(0).Height + 60)
    End If
    '说明显示滚动条
    If lngHeight > picCHKH.Height Then
        lngwNum = (picVsh.Width - vsc.Width) \ (picPrint(0).Width + 60)
        lngHnum = (lngPageCount \ lngwNum)
        If lngwNum * lngHnum < lngPageCount Then lngHnum = lngHnum + 1
        If lngHnum * (picPrint(0).Height + 60) > picVsh.Height Then
            lngHeight = lngHnum * (picPrint(0).Height + 60)
        End If
    End If
    picVsh.Height = lngHeight
    lngCHeight = picVsh.Height - picCHKH.Height
    If lngCHeight > 0 Then
        If lngCHeight < picPrint(0).Height Then
            vsc.Max = 1
            vsc.SmallChange = 1
        Else
            vsc.Max = lngCHeight
            vsc.SmallChange = picPrint(0).Height + 60
        End If
        vsc.Top = 0
        vsc.Left = picCHKH.Width - vsc.Width
        vsc.Height = picCHKH.Height
        vsc.Visible = True
        picVsh.Width = vsc.Left
    End If
    
    j = -1
    For i = 0 To picPrint.Count - 1
        If imgIco(i).Tag = strTag Or strTag = "ALL" Then
            If j > -1 Then
                lngLeft = picPrint(j).Left + picPrint(j).Width + 60
                If lngLeft + picPrint(i).Width > picVsh.Width Then
                    lngTop = picPrint(j).Top + picPrint(j).Height + 60
                    lngLeft = 60
                Else
                    lngTop = picPrint(j).Top
                End If
            Else
                lngLeft = 60
                lngTop = 60
            End If
            j = i
            picPrint(i).Left = lngLeft
            picPrint(i).Top = lngTop
            picPrint(i).Visible = True
        End If
    Next
    Call LockWindowUpdate(0)
End Sub

Private Sub SetSelectInfo()
    Dim i As Long
    Dim strInfo As String
    Dim lngStartPage As Long, lngEndPage As Long, lngPrePage As Long
    
    For i = 0 To picPrint.Count - 1
        If Val(picPrint(i).Tag) = "1" Then
            If lngPrePage = 0 Then
                lngStartPage = i + 1
                lngEndPage = lngStartPage
            ElseIf lngPrePage = i Then
                lngEndPage = i + 1
            Else
                strInfo = strInfo & "、" & lngStartPage & IIf(lngEndPage = lngStartPage, "", "-" & lngEndPage)
                lngStartPage = i + 1
                lngEndPage = lngStartPage
            End If
            lngPrePage = i + 1
        End If
    Next
    If lngStartPage > 0 Then
        strInfo = strInfo & "、" & lngStartPage & IIf(lngEndPage = lngStartPage, "", "-" & lngEndPage)
    End If
    If Left(strInfo, 1) = "、" Then strInfo = Mid(strInfo, 2)
    lblSelect.Caption = "已选页码：" & strInfo
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call zlDatabase.SetPara("连续打印", mintPrintRange, glngSys, 1255)
    Call zlDatabase.SetPara("打印页号", chk页号.Value, glngSys, 1255)
    Call zlDatabase.SetPara("起始页号", Val(txt页号.Text), glngSys, 1255)
    Call zlDatabase.SetPara("打印周数", chk周数.Value, glngSys, 1255)
    '67405:刘鹏飞,2013-11-25,添加"打印打印人"
    Call zlDatabase.SetPara("打印打印人", chkOper.Value, glngSys, 1255)
    Call zlDatabase.SetPara("不打印脉搏短绌图形", chk(0).Value, glngSys, 1255)
    Call zlDatabase.SetPara("质控号", txt.Text, glngSys, 1255, InStr(mstrPrivs, ";护理选项设置;") > 0)
End Sub

Private Sub imgIco_Click(Index As Integer)
    Call picPrint_Click(Index)
End Sub

Private Sub imgIco_DblClick(Index As Integer)
    Call picPrint_DblClick(Index)
End Sub

Private Sub lblNum_Click(Index As Integer)
    Call picPrint_Click(Index)
End Sub

Private Sub lblNum_DblClick(Index As Integer)
    Call picPrint_DblClick(Index)
End Sub

Private Sub optSelect_Click(Index As Integer)
    Dim i As Long
    Dim strTag As String
    
    If Me.Visible = False Then Exit Sub
    If optSelect(0).Value = True Then
        strTag = "待打"
    ElseIf optSelect(1).Value = True Then
        strTag = "已打"
    ElseIf optSelect(2).Value = True Then
        strTag = "重打"
    Else
        strTag = "ALL"
    End If
    For i = 0 To picPrint.Count - 1
        picPrint(i).Tag = "0"
        picPrint(i).Visible = False
        picPrint(i).BackColor = &HE0E0E0
    Next
    Call ShowCard(strTag)
    If chkSelect.Value = 1 Then Call chkSelect_Click
End Sub

Private Sub picPrint_Click(Index As Integer)
    '检查是否按了shift键盘
    Dim blnShift As Boolean
    Dim lngIndex As Long, lngStartIndex As Long, lngEndIndex As Long
    Dim i As Long
    
    lngIndex = -1
    If mblnShift = False Then Exit Sub
    For i = Index - 1 To 0 Step -1
        If Val(picPrint(i).Tag) = 1 Then
            lngIndex = i
            Exit For
        End If
    Next
    If lngIndex = -1 Then
        For i = Index + 1 To picPrint.Count - 1
            If Val(picPrint(i).Tag) = 1 Then
                lngIndex = i
                Exit For
            End If
        Next
    End If
    If lngIndex <> -1 Then
        If lngIndex > Index Then
            lngStartIndex = Index
            lngEndIndex = lngIndex
        Else
            lngStartIndex = lngIndex
            lngEndIndex = Index
        End If
        For i = lngStartIndex To lngEndIndex
            picPrint(i).Tag = 1
            picPrint(i).BackColor = IIf(Val(picPrint(i).Tag) = 1, vbRed, &HE0E0E0)
        Next
        Call SetSelectInfo
    End If
End Sub

Private Sub picPrint_DblClick(Index As Integer)
    picPrint(Index).Tag = IIf(picPrint(Index).BackColor = &HE0E0E0, 1, 0)
    picPrint(Index).BackColor = IIf(Val(picPrint(Index).Tag) = 1, vbRed, &HE0E0E0)
    Call SetSelectInfo
End Sub

Private Sub pic起始_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If pic起始.Top + Y > UD起始.Max Or pic起始.Top + Y < UD起始.Min Then Exit Sub
        pic起始.Top = pic起始.Top + Y
        UD起始.Value = pic起始.Top
        Call DrawPage
        Me.Refresh
    End If
End Sub


Private Sub txt起始_Change()
    If Val(txt起始.Text) >= UD起始.Min And Val(txt起始.Text) <= UD起始.Max Then
        UD起始.Value = Val(txt起始.Text)
    End If
End Sub

Private Sub txt起始_GotFocus()
    zlControl.TxtSelAll txt起始
End Sub

Private Sub txt起始_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt页号_GotFocus()
    zlControl.TxtSelAll txt页号
End Sub

Private Sub txt页号_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Function GetValue() As Boolean
    Dim bln连续 As Boolean
    Dim i As Long
    Dim arrPage
    
    bln连续 = False
    If Not (Val(txt起始.Text) >= UD起始.Min And Val(txt起始.Text) <= UD起始.Max) Then
        MsgBox "起始位置应该在 " & UD起始.Min & " 至 " & UD起始.Max & " 之间！", vbInformation, gstrSysName
        txt起始.SetFocus: Exit Function
    End If
    
    arrPage = Array()
    For i = 0 To picPrint.Count - 1
        If picPrint(i).Visible = True And Val(picPrint(i).Tag) = 1 Then
            ReDim Preserve arrPage(UBound(arrPage) + 1)
            arrPage(UBound(arrPage)) = i
        End If
    Next i
    If UBound(arrPage) = -1 Then
        MsgBox "请选择要打印的页！", vbInformation, gstrSysName
        Exit Function
    End If
    
    mstrPage = ""
    If optSelect(3).Value = True And chkSelect.Value = 1 Then
        mintPrintRange = 2 '全部打印
        mstrPage = 0 & ";" & UBound(arrPage)
    ElseIf UBound(arrPage) = 0 Then
        mintPrintRange = 0 '当前页
        mstrPage = arrPage(0)
    Else
        mintPrintRange = 1 '多页打印
        mstrPage = Join(arrPage, ";")
    End If
    
    mlngBeginY = Val(txt起始.Text)
    If chk页号.Value = 1 Then
        mintBeginPage = Val(txt页号.Text)
    Else
        mintBeginPage = 0
    End If
    
    GetValue = True
End Function

Private Sub UD起始_Change()
    pic起始.Top = UD起始.Value
    Call DrawPage
End Sub

Private Sub DrawPage()
    picPaper.Cls
    picPaper.Line (0, mlngTop)-(picPaper.ScaleWidth, mlngTop), &H808080
    picPaper.Line (0, picPaper.ScaleHeight - mlngBottom)-(picPaper.ScaleWidth, picPaper.ScaleHeight - mlngBottom), &H808080
    picPaper.Line (mlngLeft, 0)-(mlngLeft, picPaper.ScaleHeight), &H808080
    picPaper.Line (picPaper.ScaleWidth - mlngRight, 0)-(picPaper.ScaleWidth - mlngRight, picPaper.ScaleHeight), &H808080
    
    picPaper.Line (mlngLeft, UD起始.Value)-(picPaper.ScaleWidth - mlngRight, picPaper.ScaleHeight - mlngBottom), &H808080, B
End Sub

Private Sub vsc_Change()
    picVsh.Top = (-1) * vsc.Value
End Sub

