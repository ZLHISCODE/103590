VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmDiseaseReportSet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "疾病报告范围设置"
   ClientHeight    =   4860
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4965
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   4965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3630
      TabIndex        =   13
      Top             =   4335
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2520
      TabIndex        =   12
      Top             =   4335
      Width           =   1100
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   1
      Left            =   -45
      TabIndex        =   14
      Top             =   4155
      Width           =   5730
   End
   Begin VB.ListBox lstFiles 
      Height          =   1320
      Left            =   1110
      Style           =   1  'Checkbox
      TabIndex        =   11
      Top             =   2625
      Width           =   3210
   End
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   300
      Left            =   1380
      TabIndex        =   7
      Top             =   1785
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   0   'False
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   82903043
      CurrentDate     =   38857
   End
   Begin VB.TextBox txtDates 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   3645
      MaxLength       =   2
      TabIndex        =   4
      Text            =   "7"
      Top             =   1110
      Width           =   435
   End
   Begin VB.OptionButton optDates 
      Caption         =   "&2)指定日期范围的疾病报告:"
      Height          =   210
      Index           =   1
      Left            =   1110
      TabIndex        =   6
      Top             =   1485
      Width           =   2565
   End
   Begin VB.OptionButton optDates 
      Caption         =   "&1)最近的疾病报告(默认):"
      Height          =   210
      Index           =   0
      Left            =   1110
      TabIndex        =   3
      Top             =   1155
      Value           =   -1  'True
      Width           =   2565
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   0
      Left            =   -15
      TabIndex        =   1
      Top             =   615
      Width           =   5730
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   300
      Left            =   3045
      TabIndex        =   9
      Top             =   1785
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   0   'False
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   82903043
      CurrentDate     =   38857
   End
   Begin VB.Label lblTime 
      AutoSize        =   -1  'True
      Caption         =   "本次查阅时间范围:"
      Height          =   180
      Left            =   765
      TabIndex        =   2
      Top             =   825
      Width           =   1530
   End
   Begin VB.Label lblFiles 
      AutoSize        =   -1  'True
      Caption         =   "本工作站可管理文件(&F):"
      Height          =   180
      Left            =   765
      TabIndex        =   10
      Top             =   2340
      Width           =   1980
   End
   Begin VB.Label lblTo 
      AutoSize        =   -1  'True
      Caption         =   "从"
      Height          =   180
      Left            =   2775
      TabIndex        =   8
      Top             =   1845
      Width           =   180
   End
   Begin VB.Label lblDates 
      AutoSize        =   -1  'True
      Caption         =   "天"
      Height          =   180
      Left            =   4140
      TabIndex        =   5
      Top             =   1170
      Width           =   180
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "设置本次查看的疾病报告时间范围和需要本工作站管理的疾病报告文件。"
      Height          =   360
      Left            =   780
      TabIndex        =   0
      Top             =   150
      Width           =   3975
      WordWrap        =   -1  'True
   End
   Begin VB.Image imgNote 
      Height          =   480
      Left            =   150
      Picture         =   "frmDiseaseReportSet.frx":0000
      Top             =   90
      Width           =   480
   End
End
Attribute VB_Name = "frmDiseaseReportSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private mblnOK As Boolean

Public Function ShowMe(ByVal frmParent As Object, _
    ByVal blnFiles As Boolean, _
    ByRef strFiles As String, _
    ByRef intDates As Integer, _
    ByRef strFrom As String, _
    ByRef strTo As String) As Boolean
Dim rsTemp As New ADODB.Recordset
Dim lngCount As Long
    
    '功能：显示本窗体并提供用户设置
    '参数： blnFiles,   是否允许设置文件
    '       strFiles,   目前可管理的文件id列表
    '       intDates,   目前设置的最近日期范围
    '       strFrom,    当前设置的开始日期
    '       strTo,      当前设置的截止日期
    Dim strSetFiles As String, intSetDates As Integer, strCurDate As String, strReturn As String
    
    strSetFiles = Trim(GetSetting("ZLSOFT", App.EXEName, "疾病申报文件范围", ""))
    intSetDates = Val(GetSetting("ZLSOFT", App.EXEName, "疾病申报最近天数", 0))
    strCurDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    
    Me.dtpFrom.MaxDate = strCurDate
    Me.dtpTo.MaxDate = strCurDate
    If Trim(strFrom) = "" Or Trim(strTo) = "" Then
        Me.dtpTo.Value = CDate(strCurDate)
        Me.dtpFrom.Value = CDate(strCurDate) - intSetDates
    Else
        Me.dtpTo.Value = Format(strTo, "yyyy-MM-dd")
        Me.dtpFrom.Value = CDate(Format(strFrom, "yyyy-MM-dd")) - intSetDates
    End If
    Me.txtDates.Text = intSetDates
    
    Err = 0: On Error GoTo errHand
    gstrSQL = "Select Id, 编号, 名称 From 病历文件列表 Where 种类 = 5  Order By 编号"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    With rsTemp
        Me.lstFiles.Clear
        Do While Not .EOF
            '为支持新病，空格看不见，同时用于作分隔符
            Me.lstFiles.AddItem !编号 & "-" & !名称 & "                                   " & !ID
            Me.lstFiles.ItemData(Me.lstFiles.NewIndex) = !ID
            If InStr(1, "," & strSetFiles & ",", "," & !ID & ",") > 0 Then
                Me.lstFiles.Selected(Me.lstFiles.NewIndex) = True
            End If
            .MoveNext
        Loop
    End With
    
    If Not gobjEmr Is Nothing Then
        Set rsTemp = New ADODB.Recordset
        gstrSQL = "Select Rawtohex(ID) ID,Code 编号, Title 名称 From Antetype_List Where Kind = '04' Order By Code"
        strReturn = gobjEmr.OpenSQLRecordset(gstrSQL, "", rsTemp)
        If strReturn = "" Then
            With rsTemp
                Do While Not .EOF
                    '为支持新病，空格看不见，同时用于作分隔符
                    Me.lstFiles.AddItem !编号 & "-" & !名称 & "                                   " & !ID
                    Me.lstFiles.ItemData(Me.lstFiles.NewIndex) = 0
                    If InStr(1, "," & strSetFiles & ",", "," & !ID & ",") > 0 Then
                        Me.lstFiles.Selected(Me.lstFiles.NewIndex) = True
                    End If
                    .MoveNext
                Loop
            End With
        End If
    End If
    
    If intDates <> 0 Then
        Me.optDates(0).Value = True
    Else
        Me.optDates(1).Value = True
    End If
    Me.lstFiles.Enabled = blnFiles
    
    '显示窗体
    Me.Show vbModal, frmParent
    
    '返回处理
    If mblnOK Then
        If Me.optDates(0).Value Then
            intDates = Val(Me.txtDates.Text)
            Call SaveSetting("ZLSOFT", App.EXEName, "疾病申报最近天数", intDates)
        Else
            intDates = 0
            strFrom = Format(Me.dtpFrom.Value, "yyyy-MM-dd")
            strTo = Format(Me.dtpTo.Value, "yyyy-MM-dd")
        End If
        If Me.lstFiles.Enabled Then
            strFiles = ""
            For lngCount = 0 To Me.lstFiles.ListCount - 1
                If Me.lstFiles.Selected(lngCount) Then
                    If IsNumeric(Split(lstFiles.List(lngCount), "                                   ")(1)) Then
                        strFiles = strFiles & "," & Me.lstFiles.ItemData(lngCount)
                    Else
                        strFiles = strFiles & "," & Split(lstFiles.List(lngCount), "                                   ")(1)
                    End If
                End If
            Next
            If strFiles <> "" Then strFiles = Mid(strFiles, 2)
            Call SaveSetting("ZLSOFT", App.EXEName, "疾病申报文件范围", strFiles)
        End If
    End If
    ShowMe = mblnOK: Unload Me
    Exit Function

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Unload Me
End Function

Private Sub cmdCancel_Click()
    mblnOK = False: Me.Hide
End Sub

Private Sub cmdOK_Click()
Dim blnSelected As Boolean
Dim lngCount As Long
    If Me.optDates(0).Value And Val(Me.txtDates.Text) <= 0 Then
        MsgBox "最近天数必须大于0！", vbExclamation, gstrSysName: Me.txtDates.SetFocus: Exit Sub
    End If
    If Me.lstFiles.Enabled Then
        For lngCount = 0 To Me.lstFiles.ListCount - 1
            If Me.lstFiles.Selected(lngCount) Then
                blnSelected = True
                Exit For
            End If
        Next
        If Not blnSelected Then
            MsgBox "没有设置本工作站可管理的疾病报告文件！", vbExclamation, gstrSysName: Me.lstFiles.SetFocus: Exit Sub
        End If
    End If
    mblnOK = True: Me.Hide
End Sub

Private Sub dtpFrom_Change()
    If Me.dtpTo.Value < Me.dtpFrom.Value Then Me.dtpTo.Value = Me.dtpFrom.Value
End Sub

Private Sub dtpFrom_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub dtpTo_Change()
    If Me.dtpFrom.Value > Me.dtpTo.Value Then Me.dtpFrom.Value = Me.dtpTo.Value
End Sub

Private Sub dtpTo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub lstFiles_GotFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub lstFiles_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub optDates_Click(Index As Integer)
    If Me.optDates(0).Value Then
        Me.txtDates.Enabled = True
        Me.dtpFrom.Enabled = False: Me.dtpTo.Enabled = False
        If Me.txtDates.Visible Then Me.txtDates.SetFocus
    Else
        Me.txtDates.Enabled = False
        Me.dtpFrom.Enabled = True: Me.dtpTo.Enabled = True
        If Me.dtpFrom.Visible Then Me.dtpFrom.SetFocus
    End If
End Sub

Private Sub optDates_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub txtDates_GotFocus()
    Me.txtDates.SelStart = 0: Me.txtDates.SelLength = 1000
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txtDates_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub



