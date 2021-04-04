VERSION 5.00
Begin VB.Form frm病种选择_山西 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "病种选择"
   ClientHeight    =   2220
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5025
   Icon            =   "frm病种选择_山西.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   5025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消"
      Height          =   375
      Left            =   3465
      TabIndex        =   7
      Top             =   1635
      Width           =   1215
   End
   Begin VB.CommandButton cmd出院信息 
      Caption         =   "…"
      Height          =   300
      Left            =   4425
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1050
      Width           =   285
   End
   Begin VB.CommandButton cmd入院信息 
      Caption         =   "…"
      Height          =   300
      Left            =   4410
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   510
      Width           =   285
   End
   Begin VB.TextBox txt出院病种 
      Height          =   300
      Left            =   1035
      TabIndex        =   4
      Top             =   1050
      Width           =   3390
   End
   Begin VB.TextBox txt入院病种 
      Height          =   300
      Left            =   1035
      TabIndex        =   1
      Top             =   510
      Width           =   3390
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "确定"
      Height          =   375
      Left            =   2130
      TabIndex        =   6
      Top             =   1650
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "出院病种"
      Height          =   270
      Left            =   240
      TabIndex        =   3
      Top             =   1110
      Width           =   870
   End
   Begin VB.Label Label1 
      Caption         =   "入院病种"
      Height          =   270
      Left            =   255
      TabIndex        =   0
      Top             =   570
      Width           =   870
   End
End
Attribute VB_Name = "frm病种选择_山西"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim mstrSQL As String
Dim mrsTMP As New ADODB.Recordset
Dim mlng病人ID As Long
Dim mblnOK As Boolean '是按的确定钮退出
Dim mstr入院病种编码 As String
Dim mstr入院病种名称 As String
Dim mstr出院病种编码 As String
Dim mstr出院病种名称 As String


Private Function 入出病种选择(Optional strDisName As String = "") As String
    Dim rsTmp As ADODB.Recordset
    Dim strTmpSQL As String
    
    If strDisName <> "" Then
        strTmpSQL = "select rownum as ID,aka120  病种编码,aka121 病种名称,aka066 助记码,aae035 变更日期 from ka06" & _
                    " where aka120 like '%" & Trim(strDisName) & "%' or aka121 like '%" & Trim(strDisName) & "%' or Upper(aka066) like '%" & UCase(Trim(strDisName)) & "%'"
    Else
        strTmpSQL = "select rownum as ID,aka120  病种编码,aka121 病种名称,aka066 助记码,aae035 变更日期 from ka06"
    End If
    Set rsTmp = frmPubSel.ShowSelect(Me, strTmpSQL, 0, "病种", True, , , , , gcnSxDr)
    If rsTmp Is Nothing Then
        入出病种选择 = "|"
        Exit Function
    End If
    入出病种选择 = rsTmp!病种编码 & "|" & rsTmp!病种名称
End Function

Public Function Select病种(lng病人ID As Long, ByRef str入院病种编码, ByRef str入院病种名称, ByRef str出院病种编码, ByRef str出院病种名称) As Boolean

    mstrSQL = "Select * from 保险病种 where ID=(select 病种ID from 保险帐户 where 病人ID=" & lng病人ID & ")"
    Call OpenRecordset(mrsTMP, "入院病种", mstrSQL)
    If mrsTMP.EOF Then
        mstr入院病种名称 = ""
    Else
        mstr入院病种名称 = mrsTMP!名称
        mstr入院病种编码 = mrsTMP!编码
    End If
    
    mstrSQL = "Select * from 保险病种 where ID=(select 出院病种ID from 保险帐户 where 病人ID=" & lng病人ID & ")"
    Call OpenRecordset(mrsTMP, "入院病种", mstrSQL)
    If mrsTMP.EOF Then
        mstr出院病种名称 = ""
    Else
        mstr出院病种名称 = mrsTMP!名称
        mstr出院病种编码 = mrsTMP!编码
    End If
    mlng病人ID = lng病人ID
    Select病种 = mblnOK
    
    frm病种选择_山西.Show 1
    
    str入院病种编码 = mstr入院病种编码
    str入院病种名称 = mstr入院病种名称
    str出院病种编码 = mstr出院病种编码
    str出院病种名称 = mstr出院病种名称
    Select病种 = mblnOK
    
    Unload Me
End Function

Private Sub cmd出院信息_Click()
  ''调病种选择器
    Dim strReturn As String
    strReturn = "|"

    strReturn = 入出病种选择
    If Trim(strReturn) <> "|" Then
        txt出院病种.Text = Split(strReturn, "|")(1)
        txt出院病种.Tag = Split(strReturn, "|")(0)
    End If
End Sub

Private Sub cmd入院信息_Click()
  ''调病种选择器
    Dim strReturn As String
    strReturn = "|"
    strReturn = 入出病种选择
    If Trim(strReturn) <> "|" Then

        txt入院病种.Text = Split(strReturn, "|")(1)
        txt入院病种.Tag = Split(strReturn, "|")(0)
        
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{Tab}", 1
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    '去掉单引号
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    txt入院病种.Text = mstr入院病种名称
    txt入院病种.Tag = mstr入院病种编码
    txt出院病种.Text = mstr出院病种名称
    txt出院病种.Tag = mstr出院病种编码

End Sub

Private Sub OKButton_Click()

    Dim cur病种ID  As Currency  '用currency不容易出现界面未知错误.
    Dim str病种简码 As String
    
    If Trim(txt入院病种.Text) = "" Or Trim(txt出院病种.Text) = "" Then
        MsgBox "必须选择病种！", vbInformation, gstrSysName
        Exit Sub
    End If
    
        '保存病种信息到保险病种表中
      '判断库中有没有这个病种,如有，则直接取得病种ID
    mstrSQL = "select * from 保险病种 where 险类=" & TYPE_山西 & _
                                         " and 编码='" & txt入院病种.Tag & "'"
    Call OpenRecordset(mrsTMP, "查病种ID", mstrSQL)
    If mrsTMP.EOF Then
        mstrSQL = "select 保险病种_ID.NextVal as ID from Dual "
        Call OpenRecordset(mrsTMP, "取病种ID", mstrSQL)
        cur病种ID = 1
        If Not mrsTMP.EOF Then cur病种ID = mrsTMP!ID
        
        mstrSQL = "select zlspellcode('" & txt入院病种.Text & "') as 简码 from dual"
        Call OpenRecordset(mrsTMP, "取病种简码", mstrSQL)
        str病种简码 = mrsTMP!简码
        
        mstrSQL = "zl_保险病种_insert(" & cur病种ID & "," & TYPE_山西 & ",'" & _
                                         txt入院病种.Tag & "','" & _
                                         txt入院病种.Text & "','" & _
                                         str病种简码 & "',1,NULL,NULL)"
        gcnOracle.Execute mstrSQL, , adCmdStoredProc
        
        
        mrsTMP.Close
        Set mrsTMP = Nothing
    Else
       cur病种ID = mrsTMP!ID
    End If

    gstrSQL = " ZL_保险帐户_更新信息(" & mlng病人ID & "," & TYPE_山西 & ",'病种ID','''" & cur病种ID & "''')"
    Call zldatabase.ExecuteProcedure(gstrSQL, "病种ID")

    mstrSQL = "select * from 保险病种 where 险类=" & TYPE_山西 & _
                                         " and 编码='" & txt出院病种.Tag & "'"
    Call OpenRecordset(mrsTMP, "查病种ID", mstrSQL)
    If mrsTMP.EOF Then
        mstrSQL = "select 保险病种_ID.NextVal as ID from Dual "
        Call OpenRecordset(mrsTMP, "取病种ID", mstrSQL)
        cur病种ID = 1
        If Not mrsTMP.EOF Then cur病种ID = mrsTMP!ID
        
        mstrSQL = "select zlspellcode('" & txt出院病种.Text & "') as 简码 from dual"
        Call OpenRecordset(mrsTMP, "取病种简码", mstrSQL)
        str病种简码 = mrsTMP!简码
        
        mstrSQL = "zl_保险病种_insert(" & cur病种ID & "," & TYPE_山西 & ",'" & _
                                         txt出院病种.Tag & "','" & _
                                         txt出院病种.Text & "','" & _
                                         str病种简码 & "',1,NULL,NULL)"
        gcnOracle.Execute mstrSQL, , adCmdStoredProc
        
        mrsTMP.Close
        Set mrsTMP = Nothing
    Else
       cur病种ID = mrsTMP!ID
    End If
    gstrSQL = " ZL_保险帐户_更新信息(" & mlng病人ID & "," & TYPE_山西 & ",'出院病种ID','''" & cur病种ID & "''')"
    Call zldatabase.ExecuteProcedure(gstrSQL, "病种ID")
    
    mstr入院病种名称 = txt入院病种.Text
    mstr入院病种编码 = txt入院病种.Tag
    
    mstr出院病种名称 = txt出院病种.Text
    mstr出院病种编码 = txt出院病种.Tag
    mblnOK = True
    Unload Me
End Sub


Private Sub txt入院病种_KeyPress(KeyAscii As Integer)
  ''调病种选择器
    Dim strReturn As String
    
    If KeyAscii <> vbKeyReturn Then Exit Sub
    strReturn = 入出病种选择(Trim(txt入院病种.Text))
    If Trim(strReturn) <> "|" Then

    txt入院病种.Text = Split(strReturn, "|")(1)
    txt入院病种.Tag = Split(strReturn, "|")(0)
    End If
End Sub

Private Sub txt出院病种_KeyPress(KeyAscii As Integer)
  ''调病种选择器
    Dim strReturn As String
    If KeyAscii <> vbKeyReturn Then Exit Sub
    strReturn = 入出病种选择(Trim(txt出院病种.Text))
    If Trim(strReturn) <> "|" Then

    txt出院病种.Text = Split(strReturn, "|")(1)
    txt出院病种.Tag = Split(strReturn, "|")(0)
    End If
End Sub

