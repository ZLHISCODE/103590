VERSION 5.00
Begin VB.Form frmReportES 
   BorderStyle     =   0  'None
   Caption         =   "�ھ�����"
   ClientHeight    =   2520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7410
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   2520
   ScaleWidth      =   7410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Frame frmESItem 
      BorderStyle     =   0  'None
      Height          =   1905
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   7095
      Begin VB.TextBox txt���� 
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
      Begin VB.TextBox txtHP���� 
         Height          =   350
         Left            =   3480
         TabIndex        =   2
         Top             =   345
         Width           =   975
      End
      Begin VB.TextBox txtϸ��ˢ 
         Height          =   350
         Left            =   5280
         TabIndex        =   3
         Top             =   690
         Width           =   1575
      End
      Begin VB.TextBox txt������ 
         Height          =   350
         Left            =   3480
         TabIndex        =   1
         Top             =   690
         Width           =   975
      End
      Begin VB.TextBox txt��첿λ 
         Height          =   350
         Left            =   840
         TabIndex        =   0
         Top             =   345
         Width           =   1605
      End
      Begin VB.Label Label8 
         Caption         =   "���ߣ�"
         Height          =   195
         Left            =   0
         TabIndex        =   16
         Top             =   75
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "�����ţ�"
         Height          =   255
         Left            =   0
         TabIndex        =   14
         Top             =   735
         Width           =   975
      End
      Begin VB.Label lbl������� 
         Caption         =   "������ϣ�"
         Height          =   615
         Left            =   0
         TabIndex        =   13
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "HBsAg��"
         Height          =   195
         Left            =   4560
         TabIndex        =   10
         Top             =   435
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "ϸ��ˢ��"
         Height          =   195
         Left            =   4560
         TabIndex        =   9
         Top             =   765
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "HP���飺"
         Height          =   195
         Left            =   2520
         TabIndex        =   8
         Top             =   435
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "��������"
         Height          =   195
         Left            =   2520
         TabIndex        =   7
         Top             =   765
         Width           =   975
      End
      Begin VB.Label lbl��첿λ 
         Caption         =   "��첿λ��"
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


Private mblnCheckModity As Boolean      '�Ƿ��������ݱ仯��¼

'�ھ�ר�Ʊ���Ҫ��
Private Const Report_Element_������� = "�������"
Private Const Report_Element_������� = "�������"
Private Const Report_Element_��첿λ = "��첿λ"
Private Const Report_Element_������ = "������"
Private Const Report_Element_ϸ��ˢ = "ϸ��ˢ"
Private Const Report_Element_HP���� = "HP����"
Private Const Report_Element_HBsAg = "HBsAg"
Private Const Report_Element_���� = "����"


Public Sub zlRefresh()
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    txtPathologyNo.Text = ""
    txtPathologyDiag.Text = ""
    txt��첿λ.Text = ""
    txt������.Text = ""
    txtϸ��ˢ.Text = ""
    txtHP����.Text = ""
    txtHBsAg.Text = ""
    txt����.Text = ""
    
    mblnCheckModity = False     'ֹͣ���ݱ仯��¼
    gModified = False

    strSql = "Select �����ı�,Ҫ������ From ���Ӳ������� Where �ļ�ID=[1] And ��������=4 And ��ֹ��=0 And �滻��=0"
    If gblnMoved = True Then
        strSql = Replace(strSql, "���Ӳ�������", "H���Ӳ�������")
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, glngReportId)
    
    While rsTemp.EOF = False
        Select Case Nvl(rsTemp!Ҫ������)
            Case Report_Element_�������
                txtPathologyNo.Text = Nvl(rsTemp!�����ı�)
            Case Report_Element_�������
                txtPathologyDiag.Text = Nvl(rsTemp!�����ı�)
            Case Report_Element_��첿λ
                txt��첿λ.Text = Nvl(rsTemp!�����ı�)
            Case Report_Element_������
                txt������.Text = Nvl(rsTemp!�����ı�)
            Case Report_Element_ϸ��ˢ
                txtϸ��ˢ.Text = Nvl(rsTemp!�����ı�)
            Case Report_Element_HP����
                txtHP����.Text = Nvl(rsTemp!�����ı�)
            Case Report_Element_HBsAg
                txtHBsAg.Text = Nvl(rsTemp!�����ı�)
            Case Report_Element_����
                txt����.Text = Nvl(rsTemp!�����ı�)
        End Select
        rsTemp.MoveNext
    Wend
    
    '���ý���ؼ��Ƿ���Ա༭
    frmESItem.Enabled = gblnEditable
'    frmPathology.Enabled = mblnEditable
    
    mblnCheckModity = True     '�������ݱ仯��¼
End Sub

Public Function GetElementString() As String
    Dim strElements As String
    
    strElements = SPLITER_REPORT & Report_Element_������� & SPLITER_ELEMENT & txtPathologyNo.Text & SPLITER_REPORT & _
                Report_Element_������� & SPLITER_ELEMENT & txtPathologyDiag.Text & SPLITER_REPORT & _
                Report_Element_��첿λ & SPLITER_ELEMENT & txt��첿λ.Text & SPLITER_REPORT & _
                Report_Element_������ & SPLITER_ELEMENT & Val(txt������.Text) & SPLITER_REPORT & _
                Report_Element_ϸ��ˢ & SPLITER_ELEMENT & txtϸ��ˢ.Text & SPLITER_REPORT & _
                Report_Element_HP���� & SPLITER_ELEMENT & txtHP����.Text & SPLITER_REPORT & _
                Report_Element_HBsAg & SPLITER_ELEMENT & txtHBsAg.Text & SPLITER_REPORT & _
                Report_Element_���� & SPLITER_ELEMENT & txt����.Text
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
'    '�ڷſؼ�λ��
'    If Me.Width > 10500 Then
'        '�������ݣ��ĳ�һ��
'        '�ھ���Ŀ
'        Label4.Top = Label1.Top
'        Label4.Left = txtHP����.Left + txtHP����.Width + 50
'
'        txtϸ��ˢ.Top = txt��첿λ.Top
'        txtϸ��ˢ.Left = Label4.Left + Label4.Width + 50
'
'        Label5.Top = Label4.Top
'        Label5.Left = txtϸ��ˢ.Left + txtϸ��ˢ.Width + 50
'
'        txtHBsAg.Top = txtϸ��ˢ.Top
'        txtHBsAg.Left = Label5.Left + Label5.Width + 50
'
'        '�������
'        txtPathologyNo.Left = Label7.Left
'        txtPathologyNo.Top = Label7.Top + Label7.Height + 50
'
'        Label6.Left = txtPathologyNo.Left + txtPathologyNo.Width + 50
'        Label6.Top = Label7.Top
'
'        txtPathologyDiag.Left = Label6.Left
'        txtPathologyDiag.Top = Label6.Top + Label6.Height + 50
'    Else
'        '�����ų�����
'        '�ھ���Ŀ
'        Label4.Top = txt��첿λ.Top + txt��첿λ.Height + 50
'        Label4.Left = Label1.Left
'
'        txtϸ��ˢ.Top = Label4.Top - 50
'        txtϸ��ˢ.Left = txt��첿λ.Left
'
'        Label5.Top = Label4.Top
'        Label5.Left = Label2.Left
'
'        txtHBsAg.Top = txtϸ��ˢ.Top
'        txtHBsAg.Left = txt������.Left
'
'        '�������
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
'        strRegPath = "����ģ��\" & App.ProductName & "\frmReport\SingleWindow"
'    Else
'        strRegPath = "����ģ��\" & App.ProductName & "\frmReport"
'    End If
'
'    SaveSetting "ZLSOFT", strRegPath, "CY22", Me.Height
End Sub

Private Sub frmPathology_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub lbl�������_DblClick()
    On Error GoTo err
    If Not gobjParent Is Nothing Then
       Call gobjParent.WordItemClick(ReportViewType_�������, ReportViewType_�������, txtPathologyDiag.Text)
    End If
err:
    
End Sub

Private Sub lbl��첿λ_DblClick()
    On Error GoTo err
    If Not gobjParent Is Nothing Then
        Call gobjParent.WordItemClick(ReportViewType_��첿λ, ReportViewType_��첿λ, txt��첿λ.Text)
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

Private Sub txtHP����_Change()
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

Private Sub txt��첿λ_Change()
    If mblnCheckModity = True Then
        gModified = True
    End If
End Sub

Private Sub txt��첿λ_GotFocus()
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt��첿λ_LostFocus()
    Call zlCommFun.OpenIme
End Sub

Private Sub txt������_Change()
    If mblnCheckModity = True Then
        gModified = True
    End If
End Sub

Private Sub txtϸ��ˢ_Change()
    If mblnCheckModity = True Then
        gModified = True
    End If
End Sub

Private Sub txtϸ��ˢ_GotFocus()
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txtϸ��ˢ_LostFocus()
    Call zlCommFun.OpenIme
End Sub

Private Sub txt����_Change()
    If mblnCheckModity = True Then
        gModified = True
    End If
End Sub

Private Sub txt����_GotFocus()
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt����_LostFocus()
    Call zlCommFun.OpenIme
End Sub

Public Sub zlWriteWord(strWord As String, strReportViewType As String)
    If strReportViewType = ReportViewType_������� Then
        txtPathologyDiag.Text = txtPathologyDiag.Text & strWord
    ElseIf strReportViewType = ReportViewType_��첿λ Then
        txt��첿λ.Text = txt��첿λ.Text & strWord
    End If
End Sub
