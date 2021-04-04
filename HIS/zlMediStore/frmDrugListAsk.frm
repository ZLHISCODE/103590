VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDrugListAsk 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "查询条件设置"
   ClientHeight    =   2100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5415
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2100
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fraRangeSelect 
      Caption         =   "范围选择"
      Height          =   1875
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   3945
      Begin VB.ComboBox cob库房 
         Height          =   300
         Left            =   1005
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   300
         Width           =   2460
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "…"
         Height          =   270
         Left            =   3615
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1395
         Width           =   255
      End
      Begin VB.TextBox txt名称 
         Height          =   300
         Left            =   1005
         TabIndex        =   3
         Top             =   1380
         Width           =   2880
      End
      Begin MSComCtl2.DTPicker dtpEndDate 
         Height          =   300
         Left            =   1005
         TabIndex        =   2
         Top             =   1020
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy年MM月dd日 HH:mm:ss"
         Format          =   101122051
         CurrentDate     =   36257.9583333333
      End
      Begin MSComCtl2.DTPicker dtpStartDate 
         Height          =   300
         Left            =   1005
         TabIndex        =   1
         Top             =   660
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy年MM月dd日 HH:mm:ss"
         Format          =   101122051
         CurrentDate     =   36257
      End
      Begin VB.Label lbl库房 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "库房"
         Height          =   180
         Left            =   615
         TabIndex        =   11
         Top             =   390
         Width           =   360
      End
      Begin VB.Label lbl名称 
         AutoSize        =   -1  'True
         Caption         =   "药品名称"
         Height          =   180
         Left            =   255
         TabIndex        =   10
         Top             =   1455
         Width           =   720
      End
      Begin VB.Label lblStartDate 
         BackStyle       =   0  'Transparent
         Caption         =   "开始日期"
         Height          =   180
         Left            =   240
         TabIndex        =   9
         Top             =   720
         Width           =   735
      End
      Begin VB.Label lblEndDate 
         BackStyle       =   0  'Transparent
         Caption         =   "终止日期"
         Height          =   180
         Left            =   255
         TabIndex        =   8
         Top             =   1080
         Width           =   720
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4185
      TabIndex        =   6
      Top             =   555
      Width           =   1100
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   4185
      TabIndex        =   5
      Top             =   150
      Width           =   1100
   End
End
Attribute VB_Name = "frmDrugListAsk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public blnAskOk As Boolean

Public InDrugId As Long            '药品id
Public InDrugName  As String       '药品名称
Public InDrugStAndard As String      '药品规格
Public inDeptId As Long
Public InDrugUnit As String          '药品单位
Public intUnitLevel As Integer       '单位级数

Dim blnFirst As Boolean
Dim rsDrug As ADODB.Recordset
Dim rsTemp As ADODB.Recordset
Dim PrvPara   As Byte       '参数 0：表示简码名称双向匹配,1: 表示简码名称从左匹配
Dim StrFh As String         '前匹配符号%"
Dim strsql As String
Dim Lng部门 As Long       '保存当前选择器的部门
Dim sngLeft, sngTop As Single
Dim Bln西成药 As Boolean '表示是否具有查询西成药的权限
Dim Bln中成药 As Boolean '表示是否具有查询中成药的权限
Dim Bln中草药 As Boolean '表示是否具有查询中草药的权限
Dim Str材质 As String



Private Sub cmdCancel_Click()
    blnAskOk = False
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    On Error GoTo errHandle
    If Me.txt名称.Tag = "C" Then
        MsgBox "请选择药品!", vbInformation, gstrSysName
        Me.txt名称.SetFocus
        Exit Sub
    End If
    
    If InStr(gstrStockSearchPrivs, "西成药") <> 0 Then
        Bln西成药 = True
    Else
        Bln西成药 = False
    End If
    
    If InStr(gstrStockSearchPrivs, "中成药") <> 0 Then
        Bln中成药 = True
    Else
        Bln中成药 = False
    End If
    
    If InStr(gstrStockSearchPrivs, "中草药") <> 0 Then
        Bln中草药 = True
    Else
        Bln中草药 = False
    End If
    
    Str材质 = "''"
    If Bln西成药 Then Str材质 = "'西成药'"
    If Bln中成药 Then
        If Bln西成药 Then
            Str材质 = Str材质 & ",'中成药'"
        Else
            Str材质 = "'中成药'"
        End If
    End If
    If Bln中草药 Then
        If Bln中成药 Or Bln西成药 Then
            Str材质 = Str材质 & ",'中草药'"
        Else
            Str材质 = "'中草药'"
        End If
    End If

    Set rsTemp = New ADODB.Recordset
    strsql = "Select A.药品id From 药品目录 A,药品信息 B Where A.药名id=B.药名id And A.药品id=[1] And B.材质分类 In (" & Str材质 & ")"
    Call SQLTest(App.Title, Me.Caption, strsql)
    Set rsTemp = zldatabase.OpenSQLRecord(strsql, "cmdOK_Click", InDrugId)
    Call SQLTest

    If rsTemp.RecordCount = 0 Then
        MsgBox "你没有查询该药品的权限！", vbInformation, gstrSysName
        Me.txt名称.SetFocus
        Exit Sub
    End If
    rsTemp.Close
    
    blnAskOk = True
    Me.Hide
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdSelect_Click()
    Dim RecReturn As Recordset
    Dim intLevel As Integer
    
    Set RecReturn = Frm药品选择器.ShowME(Me, 1, cob库房.ItemData(cob库房.ListIndex))
 
      If RecReturn.RecordCount = 0 Then
         Unload Frm药品选择器
         Exit Sub
      End If
    
'InDrugId = RecReturn!药品ID
'    InDrugName = RecReturn!商品名
'    InDrugStAndard = IIf(IsNull(RecReturn!规格), " ", RecReturn!规格)
'    InDrugUnit = IIf(IsNull(RecReturn!剂量单位), " ", RecReturn!剂量单位)
'    Me.txt名称.Text = InDrugName
'    Me.txt名称.Tag = InDrugId
'
    'Unload Frm药品选择器
    With RecReturn
        InDrugId = !药品ID
        InDrugName = "[" & !药品编码 & "]" & !商品名
        InDrugStAndard = IIf(IsNull(!规格), " ", !规格)
        Me.txt名称.Text = InDrugName
        Me.txt名称.Tag = InDrugId
        intLevel = frmDrugQuery.intChoose级数
        
        Select Case intLevel
            Case 1
                InDrugUnit = !售价单位
                frmDrugList.Tag = "1"
            Case 2
                InDrugUnit = !门诊单位
                frmDrugList.Tag = !门诊包装
            Case 3
                InDrugUnit = !药库单位
                frmDrugList.Tag = !药库包装
            Case 4
                InDrugUnit = !住院单位
                frmDrugList.Tag = !住院包装
        End Select
         
    End With
    
    
End Sub

Private Sub cob库房_Validate(Cancel As Boolean)
    Me.txt名称.Tag = "C"
    txt名称_Validate False
End Sub

Private Sub dtpStartDate_Change()
    If Me.dtpStartDate.Value > Me.dtpEndDate.Value Then
        Me.dtpEndDate.Value = Me.dtpStartDate.Value
    End If
End Sub

Private Sub dtpEndDate_Change()
    If Me.dtpStartDate.Value > Me.dtpEndDate.Value Then
        Me.dtpStartDate.Value = Me.dtpEndDate.Value
    End If
End Sub

Private Sub Form_Activate()
    Dim i As Long
    If Not blnFirst Then Exit Sub
    PrvPara = Val(GetSetting(appName:="ZLSOFT", Section:="操作", Key:="匹配方法", Default:="0"))
    StrFh = IIf(PrvPara = 0, "%", "")
    
    Me.txt名称.Tag = InDrugId
    Me.txt名称.Text = InDrugName
    Me.cob库房.Clear
    With frmDrugQuery.cob库房
         For i = 0 To .ListCount - 1
            cob库房.AddItem .List(i)
            cob库房.ItemData(cob库房.NewIndex) = .ItemData(i)
            If .ItemData(i) = inDeptId Then
                cob库房.ListIndex = cob库房.NewIndex
            End If
         Next
    End With
    blnFirst = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{Tab}", 1
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim strsql As String
    Dim i As Long
    Me.dtpEndDate.MaxDate = Currentdate()
    Me.dtpEndDate.Value = Me.dtpEndDate.MaxDate
    Me.dtpStartDate.MaxDate = Me.dtpEndDate.Value
    
    blnFirst = True
    Me.dtpStartDate.Value = DateAdd("m", -1, Me.dtpEndDate.Value)
End Sub

Private Sub txt名称_Change()
    If blnFirst Then Exit Sub
    Me.txt名称.Tag = "C"
End Sub

Private Sub txt名称_DblClick()
    txt名称.SelStart = 0
    txt名称.SelLength = Len(txt名称.Text)
End Sub

Private Sub txt名称_GotFocus()
    txt名称.SelStart = 0
    txt名称.SelLength = LenB(StrConv(txt名称, vbFromUnicode))
End Sub

Private Sub txt名称_Validate(Cancel As Boolean)
    Dim intLevel As Integer
    
    
    
    If InStr(Me.txt名称.Text, "'") <> 0 Then
        MsgBox "名称中出现了非法字符:'", vbInformation, gstrSysName
        Cancel = True
        txt名称.SelStart = 0
        txt名称.SelLength = LenB(StrConv(txt名称, vbFromUnicode))
        Exit Sub
    End If
    
    Set rsDrug = New ADODB.Recordset
    If Me.txt名称.Tag <> "C" Or Me.txt名称 = "" Then Exit Sub
    
    sngLeft = Me.Left + fraRangeSelect.Left + txt名称.Left
    sngTop = Me.Top + Me.Height - Me.ScaleHeight + txt名称.Top + txt名称.Height
    If sngTop + 4730 > Screen.Height Then
        sngTop = sngTop - txt名称.Height - 4730
    End If
           
    Dim strkey As String
    
    strkey = txt名称.Text
    If Mid(strkey, 1, 1) = "[" Then
        If InStr(2, strkey, "]") <> 0 Then
            strkey = Mid(strkey, 2, InStr(2, strkey, "]") - 2)
        Else
            strkey = Mid(strkey, 2)
        End If
    End If
           
    Set rsDrug = Frm药品多选选择器.ShowME(Me, 1, cob库房.ItemData(cob库房.ListIndex), , , strkey, sngLeft, sngTop)
    With rsDrug
        If .RecordCount = 0 Then
            Cancel = True
            txt名称.SelStart = 0
            txt名称.SelLength = LenB(StrConv(txt名称, vbFromUnicode))
            Exit Sub
        End If
        intLevel = frmDrugQuery.intChoose级数
        
        If .RecordCount > 1 Then
            .MoveFirst
            InDrugId = !药品ID
            InDrugName = "[" & !药品编码 & "]" & !商品名
            InDrugStAndard = IIf(IsNull(!规格), " ", !规格)
            Select Case intLevel
                Case 1
                    InDrugUnit = !售价单位
                    frmDrugList.Tag = "1"
                Case 2
                    InDrugUnit = !门诊单位
                    frmDrugList.Tag = !门诊包装
                Case 3
                    InDrugUnit = !药库单位
                    frmDrugList.Tag = !药库包装
                Case 4
                    InDrugUnit = !住院单位
                    frmDrugList.Tag = !住院包装
            End Select
         Else
            InDrugId = !药品ID
            InDrugName = "[" & !药品编码 & "]" & !商品名
            InDrugStAndard = IIf(IsNull(!规格), " ", !规格)
            Select Case intLevel
                Case 1
                    InDrugUnit = !售价单位
                    frmDrugList.Tag = "1"
                Case 2
                    InDrugUnit = !门诊单位
                    frmDrugList.Tag = !门诊包装
                Case 3
                    InDrugUnit = !药库单位
                    frmDrugList.Tag = !药库包装
                Case 4
                    InDrugUnit = !住院单位
                    frmDrugList.Tag = !住院包装
            End Select
        End If
    End With
    
    Me.txt名称.Text = InDrugName
    Me.txt名称.Tag = InDrugId
End Sub


Private Function GetLevel(ByVal lng部门id As Long) As Integer
    '判断该部门只是药库而不是药房
    Dim rsTemp As New ADODB.Recordset
    Dim intChoose级数 As Integer
    Dim strsql As String
    
    On Error GoTo errHandle
    strsql = "Select * From 部门性质说明 " & _
        " Where 部门id=[1] And 工作性质 IN ('西药库','中药库','成药库','制剂室','西药房','中药房','成药房') "
    Set rsTemp = zldatabase.OpenSQLRecord(strsql, "GetLevel", lng部门id)
    If Not rsTemp.EOF Then
        Select Case rsTemp!服务对象
            Case 0
                intChoose级数 = 3
            Case 1, 3
                intChoose级数 = 2
            Case 2
                intChoose级数 = 4
            Case Else
                intChoose级数 = 1
        End Select
    Else
        intChoose级数 = 1
    End If
   
    rsTemp.Close
    
    GetLevel = intChoose级数
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

