VERSION 5.00
Object = "{CA73588D-282F-4592-9369-A61CC244FADA}#15.3#0"; "Codejock.SyntaxEdit.v15.3.1.ocx"
Begin VB.Form frmProcEditCommon 
   BackColor       =   &H00FFFFFF&
   Caption         =   "���̱༭"
   ClientHeight    =   7260
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   9090
   Icon            =   "frmProcEditCommon.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7260
   ScaleWidth      =   9090
   StartUpPosition =   1  '����������
   WindowState     =   2  'Maximized
   Begin XtremeSyntaxEdit.SyntaxEdit txtEdit 
      Height          =   5655
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   8655
      _Version        =   983043
      _ExtentX        =   15266
      _ExtentY        =   9975
      _StockProps     =   84
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   1
      EnableSyntaxColorization=   -1  'True
      ShowLineNumbers =   -1  'True
      ShowSelectionMargin=   0   'False
      ShowScrollBarVert=   -1  'True
      ShowScrollBarHorz=   -1  'True
      EnableVirtualSpace=   0   'False
      EnableAutoIndent=   -1  'True
      ShowWhiteSpace  =   0   'False
      ShowCollapsibleNodes=   -1  'True
      AutoCompleteWndWidth=   160
      EnableEditAccelerators=   -1  'True
   End
   Begin VB.PictureBox pctBottom 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   240
      ScaleHeight     =   465
      ScaleWidth      =   8610
      TabIndex        =   3
      Top             =   6600
      Width           =   8610
      Begin VB.CommandButton cmdCancel 
         Appearance      =   0  'Flat
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   1200
         TabIndex        =   2
         Top             =   5
         Width           =   1095
      End
      Begin VB.CommandButton cmdSave 
         Appearance      =   0  'Flat
         Caption         =   "����(&O)"
         Height          =   350
         Left            =   0
         TabIndex        =   1
         Top             =   5
         Width           =   1095
      End
   End
   Begin VB.Label lblProc 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "��������"
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   720
   End
End
Attribute VB_Name = "frmProcEditCommon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngId As Long
Private mstrProcDec As String
Private mstrUser As String

Private mblnSave As Boolean
Private mblnErr As Boolean

Public Function ShowMe(ByVal lngID As Long, ByVal strProcName As String, ByVal strProcTxt As String, _
                                        ByVal strProcSys As String, ByVal strProcDec As String, ByVal strUser As String, _
                                        Optional ByVal bytType As Byte, Optional lngLine As Long) As Boolean
    'bytType:1-���̼����������ã�Ĭ�Ϲ��̱䶯�������
    '���ÿؼ���ʽ
    With txtEdit
        '���ÿؼ�����ʾ��ɫ����Ϊ��SQL
        .SyntaxSet "[Schemes]" & vbCrLf & "SQL" & vbCrLf & "[Themes]" & vbCrLf & "Default" & vbCrLf & "Alternative" & vbCrLf
        .SyntaxScheme = GetSqlColor
        .Text = strProcTxt
    End With
    
    If bytType = 0 Then
        mlngId = lngID
        mstrProcDec = strProcDec
        mstrUser = strUser
        Me.Caption = "���̱༭"
        cmdCancel.Caption = "ȡ��(&C)"
        lblProc.Caption = strProcName & "(" & strProcSys & ")"
    Else
        Me.Caption = "���̲鿴"
        cmdCancel.Caption = "�˳�(&E)"
        lblProc.Caption = strProcName
        cmdSave.Visible = False
        txtEdit.CurrPos.Row = lngLine
        txtEdit.ReadOnly = True
    End If
    Me.Show 1
    If bytType = 0 Then
        ShowMe = mblnSave    '������޸ı���رմ��ھͷ���True,���򷵻�Fasle
    End If
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim strSQL As String
    
    On Error Resume Next    'Ϊ�˱���������,����Ҫ���δ���
    strSQL = txtEdit.Text
    '��д�޸�˵��
    If Not frmProcEditor.ShowMe(mstrUser, mstrProcDec) Then Exit Sub
    
    '��������ı�,ִ��һ�μ���
    gcnOldOra.Execute strSQL '������΢�������ִ��,��ֹ��Ϊ�����ŵ������ַ�����ʧ��
    If err.Number <> 0 Then
        mblnErr = True
        MsgBox "���̱���������������" & vbNewLine & err.Description, , "����"
        Exit Sub
    ElseIf gcnOracle.Errors.Count > 1 Then
        mblnErr = True
        MsgBox "���̱���������������", , "����"
        Exit Sub
    Else
        mblnErr = False
    End If

    '����zlProcedure����
    strSQL = "Update zlProcedure Set ˵�� = '" & mstrProcDec & "',�޸���Ա='" & mstrUser & "',�޸�ʱ��= Sysdate " & vbNewLine & _
                ",�ϴ��޸���Ա = �޸���Ա, �ϴ��޸�ʱ�� = �޸�ʱ��  Where ID = " & mlngId
    gcnOracle.Execute strSQL, "����䶯���̡�"
    If err <> 0 Then
        mblnErr = True
        MsgBox "���̱���ʧ�ܣ����������" & vbNewLine & err.Description, , "��ʾ"
        Exit Sub
    End If
    
    mblnSave = True
    Unload Me
    
End Sub

Private Sub Form_Resize()
    
    On Error Resume Next
    
    pctBottom.Top = Me.ScaleHeight - pctBottom.Height - 100
    pctBottom.Width = Me.ScaleWidth
    
    cmdCancel.Left = pctBottom.ScaleWidth - cmdCancel.Width - 360
    cmdSave.Left = cmdCancel.Left - cmdSave.Width - 120
    
    txtEdit.Width = pctBottom.Width - txtEdit.Left - 240
    txtEdit.Height = pctBottom.Top - txtEdit.Top - 120
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim strMsg As String
    
    strMsg = "�༭��Ĺ��̴��ڴ��󣬲��Ӵ���ᵼ�¸ù���ʧЧ���Ƿ�����˳���"
    
    If mblnErr Then
        If MsgBox(strMsg, vbOKCancel, "�˳�ȷ��") = vbNo Then
            Cancel = 1
        End If
    End If
End Sub

