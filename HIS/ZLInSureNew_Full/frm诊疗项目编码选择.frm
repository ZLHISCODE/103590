VERSION 5.00
Begin VB.Form frm������Ŀ����ѡ�� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "������Ŀ����ѡ��"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5715
   Icon            =   "frm������Ŀ����ѡ��.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   5715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmd���� 
      Caption         =   "��"
      Height          =   285
      Left            =   4860
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1395
      Width           =   255
   End
   Begin VB.Frame fra 
      Height          =   30
      Index           =   1
      Left            =   0
      TabIndex        =   6
      Top             =   915
      Width           =   5715
   End
   Begin VB.Frame fra 
      Height          =   30
      Index           =   0
      Left            =   15
      TabIndex        =   5
      Top             =   2355
      Width           =   5715
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3090
      TabIndex        =   3
      Top             =   2550
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4335
      TabIndex        =   2
      Top             =   2535
      Width           =   1100
   End
   Begin VB.TextBox txt��Ŀ 
      Height          =   300
      Left            =   1230
      TabIndex        =   1
      Top             =   1395
      Width           =   3870
   End
   Begin VB.Label lbl 
      Caption         =   "ѡ��Һ��е����Ʒ������ĵ�ҽ����Ŀ���ж��롣"
      Height          =   225
      Index           =   0
      Left            =   825
      TabIndex        =   7
      Top             =   510
      Width           =   4965
   End
   Begin VB.Image img 
      Height          =   555
      Left            =   45
      Picture         =   "frm������Ŀ����ѡ��.frx":000C
      Stretch         =   -1  'True
      Top             =   240
      Width           =   615
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "��Ŀ����"
      Height          =   180
      Index           =   1
      Left            =   450
      TabIndex        =   4
      Top             =   1470
      Width           =   720
   End
End
Attribute VB_Name = "frm������Ŀ����ѡ��"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mblnOK As Boolean
Dim mcnTest As New ADODB.Connection
Private mblnChange As Boolean
Dim mblnFirst As Boolean
Dim mstrCode As String
 
Public Function ShowCard(strCode As String) As Boolean
    mblnChange = False
    
    Me.Show vbModal
    ShowCard = mblnOK
    strCode = mstrCode
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub cmdCancel_Click()
    mstrCode = ""
    Unload Me
End Sub


Private Sub cmd����_Click()
        '���˺�:20040706
        Dim strCode As String
        Dim STRNAME As String
        
        On Error Resume Next
        If frm������Ŀѡ�������山.GetCode(Me, strCode, STRNAME, True) = True Then
            Me.txt��Ŀ.Text = strCode & "-" & STRNAME
            Me.txt��Ŀ.Tag = strCode
            If cmdOK.Enabled Then cmdOK.SetFocus
        End If
End Sub

Private Sub Form_Activate()
    Dim rsTemp As New ADODB.Recordset
    
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    gstrSQL = "Select * From ���ղ��� where ����=" & TYPE_�����山 & " and ������='������Ŀ����'"
    zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
    With rsTemp
        Do While Not .EOF
            Select Case Nvl(!������)
            Case "������Ŀ����"
                  txt��Ŀ.Text = Nvl(!����ֵ)
                  txt��Ŀ.Tag = txt��Ŀ.Text
            End Select
            .MoveNext
        Loop
    End With
 End Sub

Private Sub Form_Load()
    mblnFirst = True
End Sub


Private Sub cmdOK_Click()
    
    If IsValid = False Then Exit Sub
    mstrCode = txt��Ŀ.Tag
    mblnOK = True
    mblnChange = False
    Unload Me
End Sub

Private Function IsValid() As Boolean
    Dim lngCount As Long
    Dim strTitle As String
    Dim rsTemp As New ADODB.Recordset
    If txt��Ŀ.Tag = "" Then
        ShowMsgbox "������Ŀ����Ӧ��ҽ����Ŀδѡ��!"
        txt��Ŀ.SetFocus
        Exit Function
    End If
        
    IsValid = True
End Function


Private Sub txt��Ŀ_Change()
    txt��Ŀ.Tag = ""
End Sub

Private Sub txt��Ŀ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    Dim rsTemp As New ADODB.Recordset
    Dim strLeft As String
    Dim strTemp As String
    Dim blnReturn As Boolean
    If txt��Ŀ.Text = "" Then Exit Sub
    strLeft = IIf(GetSetting("ZLSOFT", "����ģ��\����", "����ƥ��", 0) = 0, "%", "")
    strTemp = "'" & strLeft & txt��Ŀ.Text & "%'"
    
    gstrSQL = " select  ��Ʒ���� as ҽ������,  ҽԺ�������, ҩƷͨ��������, ҩƷͨ��Ӣ����,��Ʒ��, ��Ʒ������, ������Ŀ���㷽ʽ, ������ʶ, ҽ����ʶ, �Ƿ񴦷���ҩ, ҩƷ��Ӧ֢, ����ҽ��, ����Ȩ��, ����, ��װ���, " & _
             "         ��С��װ��λ, ��С������λ, ÿ���������, ָ���۸�, �б�۸�, ����֧���޼�1, ����֧���޼�2, ����֧���޼�3, ʵ��ִ�м۸�, �Ը�����1, �Ը�����2, �Ը�����3, �Ը�����4, �Ը�����5, �Ը�����6, �Ը�����7, �Ը�����8,  " & _
             "         �Ը�����9, �Ը�����10, �Ը�����11, �Ը�����12, ҽԺʹ��״̬, ����ʹ��״̬, ��׼���,  " & _
             "         ���������1, ���������2, ���������3, ƴ��������1, ƴ��������2, ƴ��������3, ��ע, ҽ���������,������׼���, ҽ�ƻ������, " & _
             "          �޸�ʱ��, Ŀ¼����  " & _
             "  from ҽ��������ĿĿ¼" & _
             "  where ҽԺ�������='61' and ( ��Ʒ���� like " & strTemp & " Or ��Ʒ�� like " & strTemp & " Or " & _
             "        ���������1 like " & UCase(strTemp) & " Or " & _
             "        ƴ��������1 like " & UCase(strTemp) & ")"
    If gcnOracle_CQYB.State = adStateOpen Then
        rsTemp.Open gstrSQL, gcnOracle_CQYB, adOpenStatic, adLockReadOnly
    Else
        'ǿ��ʹ��¼��Ϊ��״̬
        gstrSQL = "Select ����  ҽ������,����,���� FROM ������Ŀ Where Rownum<1"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    End If
                   
    If rsTemp.RecordCount > 0 Then
        '����ѡ����
        If rsTemp.RecordCount > 1 Then
            '�����ֶδ���3�ģ���ʹֻ��һ����¼�ѸöԻ�����ʾ�������Ա����û��õ��������Ϣ
            blnReturn = frmListSel.ShowSelect(TYPE_�����山, rsTemp, "ҽ������", "ҽ����Ŀѡ��", "��ѡ���Ӧ��ҽ����Ŀ��")
        Else
            blnReturn = True
        End If
    Else
        MsgBox "�޴���Ŀ!"
        Exit Sub
    End If
    
    If blnReturn = False Then Exit Sub
        '�϶����м�¼����
    txt��Ŀ.Text = rsTemp("ҽ������") & "-" & Nvl(rsTemp!��Ʒ��)
    txt��Ŀ.Tag = rsTemp("ҽ������")
    If cmdOK.Enabled Then cmdOK.SetFocus
End Sub

Private Sub txt��Ŀ_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txt��Ŀ, KeyAscii, m�ı�ʽ
End Sub


