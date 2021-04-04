VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "Msmask32.ocx"
Begin VB.Form Frm退药设置 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "设置"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6300
   Icon            =   "Frm退药设置.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   6300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MsfMutiSelect 
      Height          =   2085
      Left            =   1920
      TabIndex        =   19
      Top             =   60
      Visible         =   0   'False
      Width           =   3825
      _ExtentX        =   6747
      _ExtentY        =   3678
      _Version        =   393216
      FixedCols       =   0
      GridColor       =   -2147483631
      GridColorFixed  =   8421504
      AllowBigSelection=   0   'False
      FocusRect       =   0
      FillStyle       =   1
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3600
      TabIndex        =   16
      Top             =   2970
      Width           =   1100
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4890
      TabIndex        =   17
      Top             =   2970
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Caption         =   "批次信息"
      Height          =   1545
      Left            =   870
      TabIndex        =   0
      Top             =   1050
      Width           =   5085
      Begin MSMask.MaskEdBox Txt效期 
         Height          =   300
         Left            =   3390
         TabIndex        =   12
         Top             =   690
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Format          =   "yyyy-MM-dd"
         Mask            =   "####-##-##"
         PromptChar      =   "_"
      End
      Begin VB.CommandButton Cmd产地 
         Caption         =   "…"
         Height          =   285
         Left            =   4590
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   1080
         Width           =   285
      End
      Begin VB.TextBox Txt产地 
         Height          =   300
         Left            =   1050
         TabIndex        =   14
         Top             =   1080
         Width           =   3555
      End
      Begin VB.TextBox Txt批号 
         Height          =   300
         Left            =   1050
         MaxLength       =   8
         TabIndex        =   10
         Top             =   690
         Width           =   1485
      End
      Begin VB.TextBox Txt药品 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1050
         TabIndex        =   8
         Tag             =   "3"
         Top             =   300
         Width           =   3825
      End
      Begin VB.TextBox Txt床号 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1050
         TabIndex        =   6
         Tag             =   "1"
         Top             =   690
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox Txt姓名 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1050
         TabIndex        =   4
         Tag             =   "2"
         Top             =   1080
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox Txt科室 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1050
         TabIndex        =   2
         Tag             =   "0"
         Top             =   300
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Lbl效期 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "效期(&X)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   2700
         TabIndex        =   11
         Top             =   750
         Width           =   630
      End
      Begin VB.Label Lbl产地 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "产地(&D)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   360
         TabIndex        =   13
         Top             =   1140
         Width           =   630
      End
      Begin VB.Label Lbl批号 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "批号(&A)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   360
         TabIndex        =   9
         Top             =   750
         Width           =   630
      End
      Begin VB.Label Lbl药品 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "药品(&P)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   360
         TabIndex        =   7
         Top             =   360
         Width           =   630
      End
      Begin VB.Label Lbl床号 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "床号(&B)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   360
         TabIndex        =   5
         Top             =   750
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.Label Lbl姓名 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "姓名(&N)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   360
         TabIndex        =   3
         Top             =   1140
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.Label Lbl科室 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "科室(&R)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Visible         =   0   'False
         Width           =   630
      End
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   210
      Picture         =   "Frm退药设置.frx":000C
      Top             =   180
      Width           =   240
   End
   Begin VB.Label LblNote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "    该药品原来不分批管理，而现在分批管理，因此，请输入该药品的批次信息："
      ForeColor       =   &H8000000D&
      Height          =   420
      Left            =   870
      TabIndex        =   18
      Top             =   240
      Width           =   5040
   End
End
Attribute VB_Name = "Frm退药设置"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private arrData
Private strPar As String
Private strReturn As String
Private StrFindStyle As String
Private rsTmp As New ADODB.Recordset

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
'    If Trim(Txt批号) = "" Then
'        MsgBox "请输入批号！", vbInformation, gstrSysName
'        Txt批号.SetFocus
'        Exit Sub
'    End If
    If Txt效期 <> "____-__-__" Then
        If Not IsDate(Txt效期) Then
            MsgBox "请输入合法的效期！", vbInformation, gstrSysName
            Txt效期.SetFocus
            Exit Sub
        End If
    End If
    If Trim(Txt产地) <> "" Then Call Txt产地_KeyDown(vbKeyReturn, 0)
    Do While True
        If Not MsfMutiSelect.Visible Then Exit Do
    Loop
    If Txt产地 <> Txt产地.Tag Then Exit Sub
    strReturn = Txt批号.Text & "|" & IIf(Txt效期 = "____-__-__", "", Txt效期.Text) & "|" & Txt产地.Tag
    
    Unload Me
End Sub

Private Sub Cmd产地_Click()
    Dim Rec产地 As New ADODB.Recordset
    
    On Error GoTo errHandle
    With Rec产地
        If .State = 1 Then .Close
        gstrSQL = "Select 编码,名称,简码 From 药品生产商 Where Order By 编码 "
        
        Call SQLTest(App.Title, Me.Caption, gstrSQL)
        Set Rec产地 = zldatabase.OpenSQLRecord(gstrSQL, "cmd产地_Click")
        Call SQLTest
        
        If .EOF Then
            MsgBox "请初始化药品生产商（字典管理）！", vbInformation, gstrSysName
            Me.Txt产地.SetFocus
            Txt产地.Tag = ""
            Exit Sub
        End If
        
        With MsfMutiSelect
            .Clear
            Set .DataSource = Rec产地
            .ColWidth(0) = 800
            .ColWidth(1) = 1500
            .ColWidth(2) = 800
            .Visible = True
            .ZOrder 0
            
            .Row = 1
            .ColSel = .Cols - 1
            .SetFocus
        End With
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    strReturn = ""
    arrData = Split(strPar, "|")
    Txt科室 = arrData(Val(Txt科室.Tag))
    Txt床号 = arrData(Val(Txt床号.Tag))
    Txt姓名 = arrData(Val(Txt姓名.Tag))
    Txt药品 = arrData(Val(Txt药品.Tag))
    Txt药品.Tag = arrData(4)
    StrFindStyle = IIf(GetSetting("ZLSOFT", "公共模块\操作", "输入匹配", 0) = "0", "%", "")
    
    On Error GoTo errHandle
    gstrSQL = "Select Nvl(最大效期,0) 效期 From 药品目录 Where 药品ID=[1]"
    Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(Txt药品.Tag))

    With rsTmp
        Txt批号.Tag = !效期
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function ShowME(ByVal frmParent As Object, strShow As String) As String
    'strShow="科室|床号|姓名|药品|药品ID"
    'strReturn="批号|效期|产地"
    strPar = strShow
    Me.Show 1, frmParent
    ShowME = strReturn
End Function

Private Sub MsfMutiSelect_DblClick()
    With MsfMutiSelect
        Txt产地 = .TextMatrix(.Row, 1)
        Txt产地.Tag = Txt产地
    End With
    
    MsfMutiSelect.Visible = False
End Sub

Private Sub MsfMutiSelect_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then MsfMutiSelect_DblClick
End Sub

Private Sub MsfMutiSelect_LostFocus()
    MsfMutiSelect.Visible = False
End Sub

Private Sub Txt产地_GotFocus()
    Call GetFocus(Txt产地)
End Sub

Private Sub Txt产地_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    On Error GoTo errHandle
    Dim StrInput As String
    Dim Rec产地 As New ADODB.Recordset
    StrInput = UCase(Trim(Txt产地))
    If StrInput = "" Then
        Txt产地.Tag = ""
        Exit Sub
    End If

    gstrSQL = "Select 编码,名称,简码 From 药品生产商 Where " & _
             " (Upper(编码) Like [1] Or Upper(名称) Like [1] Or Upper(简码) Like [1]) Order By 编码 "
    Set Rec产地 = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, StrFindStyle & StrInput & "%")
    
    With Rec产地
        If .EOF Then
            If Txt产地.Tag <> UCase(Txt产地.Text) Then
                MsgBox "没有找到匹配的药品生产商，请重新输入！", vbInformation, gstrSysName
                Txt产地.SelStart = 0
                Txt产地.SelLength = LenB(StrConv(Txt产地, vbFromUnicode))
                Txt产地.Tag = ""
            End If
            Exit Sub
        End If
        
        If .RecordCount = 1 Then
            With Txt产地
                .Text = Rec产地!名称
                .Tag = .Text
            End With
        Else
            With MsfMutiSelect
                .Clear
                Set .DataSource = Rec产地
                .ColWidth(0) = 800
                .ColWidth(1) = 1500
                .ColWidth(2) = 800
                .Visible = True
                .ZOrder 0
                
                .Row = 1
                .ColSel = .Cols - 1
                .SetFocus
            End With
        End If
        CmdOK.SetFocus
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Txt批号_Change()
    Dim str批号 As String
    If Trim(Txt批号) = "" Then Exit Sub
    If Len(Trim(Txt批号)) <> 8 Then Exit Sub
    If Val(Txt批号.Tag) = 0 Then Exit Sub
    str批号 = Mid(Txt批号, 1, 4) & "-" & Mid(Txt批号, 5, 2) & "-" & Mid(Txt批号, 7, 2)
    
    If IsDate(str批号) Then
        Txt效期 = Format(DateAdd("m", Val(Txt批号.Tag), str批号), "yyyy-MM-dd")
    End If
    Txt效期.SetFocus
End Sub

Private Sub Txt批号_GotFocus()
    Call GetFocus(Txt批号)
End Sub

Public Sub GetFocus(ByVal TxtBox As TextBox)
    With TxtBox
        .SelStart = 0
        .SelLength = LenB(StrConv(.Text, vbFromUnicode))
    End With
End Sub

Private Sub Txt批号_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Txt效期.SetFocus
End Sub

Private Sub Txt效期_GotFocus()
    With Txt效期
        .SelStart = 0
        .SelLength = Len(Txt效期)
    End With
End Sub

Private Sub Txt效期_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Txt产地.SetFocus
End Sub
