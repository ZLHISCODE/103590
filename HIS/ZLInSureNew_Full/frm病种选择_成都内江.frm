VERSION 5.00
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.4#0"; "ZL9BILLEDIT.OCX"
Begin VB.Form frm病种选择_成都内江 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "病种选择"
   ClientHeight    =   3645
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6870
   Icon            =   "frm病种选择_成都内江.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   6870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton CancelButton 
      Caption         =   "取消"
      Height          =   375
      Left            =   5370
      TabIndex        =   1
      Top             =   3105
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "确定"
      Height          =   375
      Left            =   3975
      TabIndex        =   0
      Top             =   3090
      Width           =   1215
   End
   Begin ZL9BillEdit.BillEdit msf附加病种 
      Height          =   2730
      Left            =   135
      TabIndex        =   2
      Top             =   240
      Width           =   6540
      _ExtentX        =   11536
      _ExtentY        =   4815
      Enabled         =   -1  'True
      CellAlignment   =   9
      Text            =   ""
      TextMatrix0     =   ""
      MaxDate         =   2958465
      MinDate         =   -53688
      Value           =   36395
      Active          =   -1  'True
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
End
Attribute VB_Name = "frm病种选择_成都内江"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim mlng病人ID   As Long, mstr并发症 As String

Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim i As Long
    Dim vat并发症 As Variant
    
    With msf附加病种
        '设置行数及列数及列标题名称
        .Rows = 4
        .Cols = 2
        .TextMatrix(0, 0) = ""
        .ColWidth(0) = 1000
        .ColWidth(1) = 4800
        .TextMatrix(0, 1) = "病种编码与名称"
        .TextMatrix(1, 0) = "附加诊断"
        .PrimaryCol = 1
        
        '设置各列的列值，以确定哪些列可操作、可编辑或非编辑列
        .ColData(1) = 1  '文本框输入且有命令按钮
        '未设置的列的列值均为0 (默认), 这些列将可以选择但不能修改
    End With
    
    If mstr并发症 <> "" Then
        If InStr(mstr并发症, "|") > 0 Then
            vat并发症 = Split(mstr并发症, "|")
            msf附加病种.Rows = UBound(vat并发症) + 2
            For i = 0 To UBound(vat并发症) - 1
                msf附加病种.TextMatrix(i + 1, 1) = "[" & Split(vat并发症(i), ";")(0) & "]" & Split(vat并发症(i), ";")(1)
            Next
        End If
    End If
    

End Sub

Private Sub msf附加病种_BeforeDeleteRow(Row As Long, Cancel As Boolean)
    If Row < 2 Then Cancel = True
End Sub

Private Sub msf附加病种_CommandClick()
    Dim str病种 As String
    Select Case msf附加病种.ColData(msf附加病种.COL)
        Case 1
            str病种 = msf附加病种.TextMatrix(msf附加病种.Row, msf附加病种.COL)
            str病种 = BZXZ_成都内江(str病种)
            If str病种 = "" Then Exit Sub
            msf附加病种.TextMatrix(msf附加病种.Row, msf附加病种.COL) = str病种
    End Select
End Sub

Private Sub msf附加病种_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim str病种 As String
    If KeyCode <> vbKeyReturn Or msf附加病种.COL = 0 Then Exit Sub
    str病种 = msf附加病种.Text
    
    If str病种 = "" And msf附加病种.Rows = msf附加病种.Row + 1 Then
        SendKeys "{Tab}"
    End If
    
    If str病种 = "" And msf附加病种.Rows = msf附加病种.Row + 2 Then
        If msf附加病种.TextMatrix(msf附加病种.Row + 1, msf附加病种.COL) = "" Then
            SendKeys "{Tab}"
        End If
    End If
    
    'Cancel = True
    str病种 = BZXZ_成都内江(str病种, 1)
    If str病种 <> "" Then
        msf附加病种.Text = str病种
        msf附加病种.TextMatrix(msf附加病种.Row, msf附加病种.COL) = str病种
    End If
End Sub

Function BZXZ_成都内江(ByVal StrInput As String, Optional strLoad As String = 0) As String
    Dim rsTmp As ADODB.Recordset
    Dim strTmpSQL As String
    
    On Error Resume Next
   
    
    If StrInput = "" And strLoad = 1 Then Exit Function
    
    If StrInput = "" Then
        strTmpSQL = "Select ID,编码,名称 from 保险病种"
    Else
        strTmpSQL = "Select ID,编码,名称 from 保险病种" & _
                 " Where 编码 Like '%" & StrInput & "%' OR " & _
                 "名称 like '%" & StrInput & "%' Or " & _
                 "lower(简码) like lower('%" & StrInput & "%')"
    End If
    
    Set rsTmp = frmPubSel.ShowSelect(Me, strTmpSQL, 0, "病种", True, , , , False, gcnOracle)
    If rsTmp Is Nothing Then Exit Function
    BZXZ_成都内江 = "[" & rsTmp!编码 & "]" & rsTmp!名称
End Function

Function GetCode(lng病人ID) As Boolean
    Dim rsTmp As New ADODB.Recordset
    mlng病人ID = 0
    mstr并发症 = ""
    mlng病人ID = lng病人ID
    
    gstrSQL = "Select * from 保险帐户 Where 病人ID=[1] And 险类=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "取并发症", lng病人ID, TYPE_成都内江)
    mstr并发症 = Nvl(rsTmp!附加诊断)
    
    frm病种选择_成都内江.Show 1
    GetCode = True
End Function

Private Sub OKButton_Click()

    Dim rsTmp As New ADODB.Recordset, rsTmp1 As New ADODB.Recordset
    Dim i As Integer
    Dim str病种编码 As String, str病种名称 As String, lng病种ID  As Long, str附加诊断 As String
    Dim str统筹地区码 As String, str住院流水号 As String
    Dim StrInput As String, strOutput As String
    Dim lng入院病种 As Long
    '>Beging 病种处理
    lng入院病种 = 0
    If msf附加病种.Rows < 1 Then
        MsgBox "请输入诊断信息！", vbInformation, gstrSysName
        Exit Sub
    End If
 
    
    For i = 1 To msf附加病种.Rows - 1
        str病种编码 = msf附加病种.TextMatrix(i, 1)
        If str病种编码 <> "" Then
            If InStr(str病种编码, "]") > 0 And InStr(str病种编码, "[") > 0 And InStr(str病种编码, "]") - InStr(str病种编码, "[") > 1 Then
                str病种名称 = Mid(str病种编码, InStr(str病种编码, "]") + 1)
                str病种编码 = Mid(str病种编码, InStr(str病种编码, "[") + 1, InStr(str病种编码, "]") - InStr(str病种编码, "[") - 1)
                '龚智毅 20051029
                If str病种名称 = "平产" Or str病种名称 = "剖宫产" Then
                   gstrSQL = "Select id from 保险病种 where 名称=[1]"
                   Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "生育保险病种", str病种名称)
                   lng入院病种 = rsTmp!ID
                End If
                str附加诊断 = str附加诊断 & str病种编码 & ";" & str病种名称 & "|"
                
                gstrSQL = "Select * from 保险帐户 where 病人ID=[1]"
                Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "保险帐户", mlng病人ID)
                str统筹地区码 = Split(rsTmp!退休证号, "|")(0)
                str住院流水号 = rsTmp!顺序号
                '调附加诊断上传交易
                StrInput = str统筹地区码 & vbTab & str住院流水号 & vbTab & Rpad(Rpad(str病种编码, 20), 200)
                
                If 业务请求_成都内江(并发症申请上传_内江, StrInput, strOutput) Then Exit Sub
            End If
        End If
    Next
    
    If str附加诊断 <> "" Then
        gstrSQL = "zl_保险帐户_更新信息(" & mlng病人ID & "," & TYPE_成都内江 & ",'附加诊断','''" & str附加诊断 & "''')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "保存并发症")
    End If
    If lng入院病种 > 0 Then
       gstrSQL = "zl_保险帐户_更新信息(" & mlng病人ID & "," & TYPE_成都内江 & ",'病种id','''" & lng入院病种 & "''')"
       Call zlDatabase.ExecuteProcedure(gstrSQL, "保存生育病种")
    End If
    '>End 病种处理
  
    Unload Me
End Sub


