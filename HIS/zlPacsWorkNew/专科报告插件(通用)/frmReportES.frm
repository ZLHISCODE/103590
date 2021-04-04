VERSION 5.00
Begin VB.Form frmReportES 
   BorderStyle     =   0  'None
   Caption         =   "内镜报告"
   ClientHeight    =   2520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7410
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   2520
   ScaleWidth      =   7410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame frmESItem 
      BorderStyle     =   0  'None
      Height          =   1905
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   7095
      Begin VB.TextBox txt主诉 
         Height          =   350
         Left            =   840
         TabIndex        =   15
         Top             =   0
         Width           =   6015
      End
      Begin VB.TextBox txtPathologyNo 
         Height          =   350
         Left            =   840
         TabIndex        =   12
         Top             =   690
         Width           =   1605
      End
      Begin VB.TextBox txtPathologyDiag 
         Height          =   735
         Left            =   840
         MultiLine       =   -1  'True
         TabIndex        =   11
         Top             =   1065
         Width           =   6015
      End
      Begin VB.TextBox txtHBsAg 
         Height          =   350
         Left            =   5280
         TabIndex        =   4
         Top             =   345
         Width           =   1575
      End
      Begin VB.TextBox txtHP试验 
         Height          =   350
         Left            =   3480
         TabIndex        =   2
         Top             =   345
         Width           =   975
      End
      Begin VB.TextBox txt细胞刷 
         Height          =   350
         Left            =   5280
         TabIndex        =   3
         Top             =   690
         Width           =   1575
      End
      Begin VB.TextBox txt活检块数 
         Height          =   350
         Left            =   3480
         TabIndex        =   1
         Top             =   690
         Width           =   975
      End
      Begin VB.TextBox txt活检部位 
         Height          =   350
         Left            =   840
         TabIndex        =   0
         Top             =   345
         Width           =   1605
      End
      Begin VB.Label Label8 
         Caption         =   "主诉："
         Height          =   195
         Left            =   0
         TabIndex        =   16
         Top             =   75
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "病理编号："
         Height          =   255
         Left            =   0
         TabIndex        =   14
         Top             =   735
         Width           =   975
      End
      Begin VB.Label lbl病理诊断 
         Caption         =   "病理诊断："
         Height          =   615
         Left            =   0
         TabIndex        =   13
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "HBsAg："
         Height          =   195
         Left            =   4560
         TabIndex        =   10
         Top             =   435
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "细胞刷："
         Height          =   195
         Left            =   4560
         TabIndex        =   9
         Top             =   765
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "HP试验："
         Height          =   195
         Left            =   2520
         TabIndex        =   8
         Top             =   435
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "活检块数："
         Height          =   195
         Left            =   2520
         TabIndex        =   7
         Top             =   765
         Width           =   975
      End
      Begin VB.Label lbl活检部位 
         Caption         =   "活检部位："
         Height          =   195
         Left            =   0
         TabIndex        =   6
         Top             =   435
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmReportES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private mblnCheckModity As Boolean      '是否启动内容变化记录

'内镜专科报告要素
Private Const Report_Element_病理序号 = "病理序号"
Private Const Report_Element_病理诊断 = "病理诊断"
Private Const Report_Element_活检部位 = "活检部位"
Private Const Report_Element_活检块数 = "活检块数"
Private Const Report_Element_细胞刷 = "细胞刷"
Private Const Report_Element_HP试验 = "HP试验"
Private Const Report_Element_HBsAg = "HBsAg"
Private Const Report_Element_主诉 = "主诉"


Public Sub zlRefresh()
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    txtPathologyNo.Text = ""
    txtPathologyDiag.Text = ""
    txt活检部位.Text = ""
    txt活检块数.Text = ""
    txt细胞刷.Text = ""
    txtHP试验.Text = ""
    txtHBsAg.Text = ""
    txt主诉.Text = ""
    
    mblnCheckModity = False     '停止内容变化记录
    gModified = False

    strSql = "Select 内容文本,要素名称 From 电子病历内容 Where 文件ID=[1] And 对象类型=4 And 终止版=0 And 替换域=0"
    If gblnMoved = True Then
        strSql = Replace(strSql, "电子病历内容", "H电子病历内容")
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, glngReportId)
    
    While rsTemp.EOF = False
        Select Case Nvl(rsTemp!要素名称)
            Case Report_Element_病理序号
                txtPathologyNo.Text = Nvl(rsTemp!内容文本)
            Case Report_Element_病理诊断
                txtPathologyDiag.Text = Nvl(rsTemp!内容文本)
            Case Report_Element_活检部位
                txt活检部位.Text = Nvl(rsTemp!内容文本)
            Case Report_Element_活检块数
                txt活检块数.Text = Nvl(rsTemp!内容文本)
            Case Report_Element_细胞刷
                txt细胞刷.Text = Nvl(rsTemp!内容文本)
            Case Report_Element_HP试验
                txtHP试验.Text = Nvl(rsTemp!内容文本)
            Case Report_Element_HBsAg
                txtHBsAg.Text = Nvl(rsTemp!内容文本)
            Case Report_Element_主诉
                txt主诉.Text = Nvl(rsTemp!内容文本)
        End Select
        rsTemp.MoveNext
    Wend
    
    '设置界面控件是否可以编辑
    frmESItem.Enabled = gblnEditable
'    frmPathology.Enabled = mblnEditable
    
    mblnCheckModity = True     '启动内容变化记录
End Sub

Public Function GetElementString() As String
    Dim strElements As String
    
    strElements = SPLITER_REPORT & Report_Element_病理序号 & SPLITER_ELEMENT & txtPathologyNo.Text & SPLITER_REPORT & _
                Report_Element_病理诊断 & SPLITER_ELEMENT & txtPathologyDiag.Text & SPLITER_REPORT & _
                Report_Element_活检部位 & SPLITER_ELEMENT & txt活检部位.Text & SPLITER_REPORT & _
                Report_Element_活检块数 & SPLITER_ELEMENT & Val(txt活检块数.Text) & SPLITER_REPORT & _
                Report_Element_细胞刷 & SPLITER_ELEMENT & txt细胞刷.Text & SPLITER_REPORT & _
                Report_Element_HP试验 & SPLITER_ELEMENT & txtHP试验.Text & SPLITER_REPORT & _
                Report_Element_HBsAg & SPLITER_ELEMENT & txtHBsAg.Text & SPLITER_REPORT & _
                Report_Element_主诉 & SPLITER_ELEMENT & txt主诉.Text
    GetElementString = strElements
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            zlCommFun.PressKey vbKeyTab
    End Select
End Sub

Private Sub Form_Resize()
    Dim lngTemp As Long
    
    frmESItem.Left = 0
    frmESItem.Top = 0
    frmESItem.Width = Me.ScaleWidth
'    '摆放控件位置
'    If Me.Width > 10500 Then
'        '横排内容，拍成一排
'        '内镜项目
'        Label4.Top = Label1.Top
'        Label4.Left = txtHP试验.Left + txtHP试验.Width + 50
'
'        txt细胞刷.Top = txt活检部位.Top
'        txt细胞刷.Left = Label4.Left + Label4.Width + 50
'
'        Label5.Top = Label4.Top
'        Label5.Left = txt细胞刷.Left + txt细胞刷.Width + 50
'
'        txtHBsAg.Top = txt细胞刷.Top
'        txtHBsAg.Left = Label5.Left + Label5.Width + 50
'
'        '病理诊断
'        txtPathologyNo.Left = Label7.Left
'        txtPathologyNo.Top = Label7.Top + Label7.Height + 50
'
'        Label6.Left = txtPathologyNo.Left + txtPathologyNo.Width + 50
'        Label6.Top = Label7.Top
'
'        txtPathologyDiag.Left = Label6.Left
'        txtPathologyDiag.Top = Label6.Top + Label6.Height + 50
'    Else
'        '内容排成两排
'        '内镜项目
'        Label4.Top = txt活检部位.Top + txt活检部位.Height + 50
'        Label4.Left = Label1.Left
'
'        txt细胞刷.Top = Label4.Top - 50
'        txt细胞刷.Left = txt活检部位.Left
'
'        Label5.Top = Label4.Top
'        Label5.Left = Label2.Left
'
'        txtHBsAg.Top = txt细胞刷.Top
'        txtHBsAg.Left = txt活检块数.Left
'
'        '病理诊断
'        txtPathologyNo.Left = Label7.Left + Label7.Width + 50
'        txtPathologyNo.Top = Label7.Top - 15
'
'        Label6.Left = Label7.Left
'        Label6.Top = txtPathologyNo.Top + txtPathologyNo.Height + 50
'
'        txtPathologyDiag.Left = txtPathologyNo.Left
'        txtPathologyDiag.Top = Label6.Top
'    End If
'
'    frmESItem.Left = 0
'    frmESItem.Top = 0
'    lngTemp = Me.Width - 100
'    frmESItem.Width = IIf(lngTemp < 0, 0, lngTemp)
'    frmESItem.Height = txtHBsAg.Top + txtHBsAg.Height + 100
'
'    frmPathology.Left = 10
'    frmPathology.Top = frmESItem.Top + frmESItem.Height + 10
'    frmPathology.Width = frmESItem.Width
'    lngTemp = Me.Height - frmESItem.Height - 100
'    frmPathology.Height = IIf(lngTemp < 0, 0, lngTemp)
'
'    lngTemp = frmPathology.Height - txtPathologyDiag.Top - 100
'    txtPathologyDiag.Height = IIf(lngTemp < 0, 0, lngTemp)
'    lngTemp = frmPathology.Width - txtPathologyDiag.Left - 100
'    txtPathologyDiag.Width = IIf(lngTemp < 0, 0, lngTemp)

End Sub

Private Sub Form_Unload(Cancel As Integer)
'    Dim strRegPath As String
'
'    If mblnSingleWindow = True Then
'        strRegPath = "公共模块\" & App.ProductName & "\frmReport\SingleWindow"
'    Else
'        strRegPath = "公共模块\" & App.ProductName & "\frmReport"
'    End If
'
'    SaveSetting "ZLSOFT", strRegPath, "CY22", Me.Height
End Sub

Private Sub frmPathology_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub lbl病理诊断_DblClick()
    On Error GoTo err
    If Not gobjParent Is Nothing Then
       Call gobjParent.WordItemClick(ReportViewType_病理诊断, ReportViewType_病理诊断, txtPathologyDiag.Text)
    End If
err:
    
End Sub

Private Sub lbl活检部位_DblClick()
    On Error GoTo err
    If Not gobjParent Is Nothing Then
        Call gobjParent.WordItemClick(ReportViewType_活检部位, ReportViewType_活检部位, txt活检部位.Text)
    End If
err:
End Sub

Private Sub txtHBsAg_Change()
    If mblnCheckModity = True Then
        gModified = True
    End If
End Sub

Private Sub txtHBsAg_GotFocus()
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txtHBsAg_LostFocus()
    Call zlCommFun.OpenIme
End Sub

Private Sub txtHP试验_Change()
    If mblnCheckModity = True Then
        gModified = True
    End If
End Sub

Private Sub txtPathologyDiag_Change()
     If mblnCheckModity = True Then
        gModified = True
    End If
End Sub

Private Sub txtPathologyDiag_GotFocus()
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txtPathologyDiag_LostFocus()
    Call zlCommFun.OpenIme
End Sub

Private Sub txtPathologyNo_Change()
    If mblnCheckModity = True Then
        gModified = True
    End If
End Sub

Private Sub txtPathologyNo_GotFocus()
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txtPathologyNo_LostFocus()
    Call zlCommFun.OpenIme
End Sub

Private Sub txt活检部位_Change()
    If mblnCheckModity = True Then
        gModified = True
    End If
End Sub

Private Sub txt活检部位_GotFocus()
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt活检部位_LostFocus()
    Call zlCommFun.OpenIme
End Sub

Private Sub txt活检块数_Change()
    If mblnCheckModity = True Then
        gModified = True
    End If
End Sub

Private Sub txt细胞刷_Change()
    If mblnCheckModity = True Then
        gModified = True
    End If
End Sub

Private Sub txt细胞刷_GotFocus()
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt细胞刷_LostFocus()
    Call zlCommFun.OpenIme
End Sub

Private Sub txt主诉_Change()
    If mblnCheckModity = True Then
        gModified = True
    End If
End Sub

Private Sub txt主诉_GotFocus()
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt主诉_LostFocus()
    Call zlCommFun.OpenIme
End Sub

Public Sub zlWriteWord(strWord As String, strReportViewType As String)
    If strReportViewType = ReportViewType_病理诊断 Then
        txtPathologyDiag.Text = txtPathologyDiag.Text & strWord
    ElseIf strReportViewType = ReportViewType_活检部位 Then
        txt活检部位.Text = txt活检部位.Text & strWord
    End If
End Sub
