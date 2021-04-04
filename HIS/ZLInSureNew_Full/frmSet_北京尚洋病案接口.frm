VERSION 5.00
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.4#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmSet_北京尚洋病案接口 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "参数设置"
   ClientHeight    =   4005
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6585
   Icon            =   "frmSet_北京尚洋病案接口.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   6585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CMD放弃 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5370
      TabIndex        =   2
      Top             =   1080
      Width           =   1100
   End
   Begin VB.CommandButton cmd确定 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   5370
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   630
      Width           =   1100
   End
   Begin ZL9BillEdit.BillEdit BillEdit1 
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   510
      Width           =   5115
      _ExtentX        =   9022
      _ExtentY        =   5953
      CellAlignment   =   9
      Text            =   ""
      TextMatrix0     =   ""
      MaxDate         =   2958465
      MinDate         =   -53688
      Value           =   36395
      Cols            =   2
      RowHeight0      =   315
      RowHeightMin    =   315
      ColWidth0       =   1005
      BackColor       =   -2147483643
      BackColorBkg    =   -2147483643
      BackColorSel    =   10249818
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      ForeColorSel    =   -2147483634
      GridColor       =   -2147483630
      ColAlignment0   =   9
      ListIndex       =   -1
      CellBackColor   =   -2147483643
   End
   Begin VB.Label lblNote 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "请进行收据项目对照："
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   150
      TabIndex        =   3
      Top             =   210
      Width           =   1800
   End
End
Attribute VB_Name = "frmSet_北京尚洋病案接口"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstr病案费目 As String

Private Sub BillEdit1_cboClick(ListIndex As Long)
    BillEdit1.TextMatrix(BillEdit1.Row, 1) = BillEdit1.CboText
End Sub

Private Sub CMD放弃_Click()
    Unload Me
End Sub

Private Sub cmd确定_Click()
    Dim strSave As String
    Dim intDO As Integer, intCOUNT As Integer
    
    '检查每个项目是否都已设置
    intCOUNT = BillEdit1.Rows - 1
    For intDO = 1 To intCOUNT
        If Trim(BillEdit1.TextMatrix(intDO, 1)) = "" Then
            MsgBox "有项目未设置与医保的对照关系，请检查！", vbInformation, gstrSysName
            Exit Sub
        End If
    Next
    
    '产生串：'中草药,A-中草药|西成药...
    For intDO = 1 To intCOUNT
        If Trim(BillEdit1.TextMatrix(intDO, 1)) <> "" Then
            strSave = strSave & "|" & BillEdit1.TextMatrix(intDO, 0) & "," & BillEdit1.TextMatrix(intDO, 1)
        End If
    Next
    If strSave <> "" Then mstr病案费目 = Mid(strSave, 2)
    
    '保存到注册表中
    Call SaveSetting("ZLSOFT", "私有模块\" & Me.Name, "病案费目", mstr病案费目)
    
    Unload Me
    Exit Sub
End Sub

Private Sub Form_Load()
    Dim arr病案费目
    Dim strSQL As String
    Dim intDO As Integer, intCOUNT As Integer
    Dim rsTemp As New ADODB.Recordset
    
    mstr病案费目 = ""
    With BillEdit1
        .Rows = 2: .Cols = 2
        .Active = True
        .PrimaryCol = 0
        .LocateCol = 1
        
        .TextMatrix(0, 0) = "HIS名称"
        .TextMatrix(0, 1) = "中心名称"
        .ColData(0) = 5
        .ColData(1) = 3
        .ColWidth(0) = 1800
        .ColWidth(1) = 1800
        
        .AddItem "A-床位费"
        .AddItem "B-西药费"
        .AddItem "C-中成药"
        .AddItem "D-中草药"
        .AddItem "E-手术费"
        .AddItem "F-检验病理费"
        .AddItem "G-放射费"
        .AddItem "H-检查费"
        .AddItem "I-治疗费"
        .AddItem "J-诊疗费"
        .AddItem "K-护理费"
        .AddItem "M-氧气费"
        .AddItem "N-接生费"
        .AddItem "O-婴儿费"
        .AddItem "P-陪侍费"
        .AddItem "R-血  费"
        .AddItem "Y-麻醉费"
        .AddItem "Z-其  他"
    End With
    
    '提取病案费目
    strSQL = " Select 编码,名称 From 病案费目"
    Call OpenRecordset(rsTemp, "提取病案费目", strSQL)
    With rsTemp
        Do While Not .EOF
            BillEdit1.TextMatrix(.AbsolutePosition, 0) = !名称
            .MoveNext
            
            If Not .EOF Then BillEdit1.Rows = BillEdit1.Rows + 1
        Loop
    End With
    
    '中草药,A-中草药|西成药...
    mstr病案费目 = GetSetting("ZLSOFT", "私有模块\" & Me.Name, "病案费目", "")
    arr病案费目 = Split(mstr病案费目, "|")
    intCOUNT = UBound(arr病案费目)
    For intDO = 0 To intCOUNT
        rsTemp.MoveFirst
        rsTemp.Find "名称='" & Split(arr病案费目(intDO), ",")(0) & "'"
        If rsTemp.EOF = False Then
            BillEdit1.TextMatrix(rsTemp.AbsolutePosition, 1) = Split(arr病案费目(intDO), ",")(1)
        End If
    Next
End Sub
