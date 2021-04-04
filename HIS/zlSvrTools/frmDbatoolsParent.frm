VERSION 5.00
Begin VB.Form frmDbatoolsParent 
   BackColor       =   &H00FFFFFF&
   Caption         =   "�����Ż�����"
   ClientHeight    =   6540
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8085
   ControlBox      =   0   'False
   DrawMode        =   3  'Not Merge Pen
   DrawStyle       =   2  'Dot
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   Picture         =   "frmDbatoolsParent.frx":0000
   ScaleHeight     =   6540
   ScaleWidth      =   8085
   WindowState     =   2  'Maximized
   Begin VB.PictureBox pctContent 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5700
      Left            =   0
      ScaleHeight     =   5700
      ScaleWidth      =   7605
      TabIndex        =   1
      Top             =   480
      Width           =   7605
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�����Ż�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   195
      TabIndex        =   0
      Top             =   105
      Width           =   960
   End
End
Attribute VB_Name = "frmDbatoolsParent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mfrmTools As Object
Attribute mfrmTools.VB_VarHelpID = -1

Public Function SupportPrint() As Boolean
'���ر������Ƿ�֧�ִ�ӡ���������ڵ���
    SupportPrint = False
End Function

Private Sub Form_Resize()

    If mfrmTools Is Nothing Then Exit Sub
    
    On Error Resume Next
    pctContent.Height = Me.ScaleHeight - pctContent.Top
    pctContent.Width = Me.ScaleWidth
    
    mfrmTools.WindowState = 0
    mfrmTools.Move 0, 0, pctContent.ScaleWidth, pctContent.ScaleHeight
End Sub

Public Sub ShowToolsForm(ByVal strMoudle As String)
    Static objTools As Object
    
    On Error GoTo errH
    If objTools Is Nothing Then
        Set objTools = CreateObject("zlDbaTools.clsToolsMain")
    End If
    
    If objTools Is Nothing Then
        Me.Show
        frmMDIMain.stbThis.Panels(2).Text = "DBA���߼���ʧ�ܣ�����zlDbaTools.dll�Ƿ�ɹ�ע�ᡣ"
        Exit Sub
    End If
    
    Set mfrmTools = objTools.GetFrmByMdoudle(strMoudle, gblnDBA, gcnOracle, gstrUserName, gstrPassword)
    
    Select Case strMoudle
    Case "0601"
        Call ShowFlash("���ڼ������ݿ����ܷ�������...")
        lblTitle.Caption = "���ݿ����ܷ���"
    Case "0602"
        Call ShowFlash("���ڼ���SQL���ܷ������Ż�����...")
        lblTitle.Caption = "SQL���ܷ������Ż�"
    Case "0604"
        Call ShowFlash("���ڼ��ػỰ��������...")
        lblTitle.Caption = "�Ự����"
    Case "0605"
        Call ShowFlash("���ڼ��������������...")
        lblTitle.Caption = "�������"
    Case "0606"
        Call ShowFlash("���ڼ��ؿռ������������...")
        lblTitle.Caption = "�ռ����������"
    End Select

    '����Ӧ����һ��ShowMe������
    If mfrmTools Is Nothing Then
        Me.Show
        frmMDIMain.stbThis.Panels(2).Text = "��ǰ�û�����DBA�û���Ȩ�޲��㣬�޷�ʹ�øù��ܡ�"
        Call ShowFlash("")
        Exit Sub
    Else
        LockWindowUpdate Me.hwnd
        SetParent mfrmTools.hwnd, pctContent.hwnd
        mfrmTools.ShowMe
    End If
    
    Form_Resize
    Call ShowFlash("")
    LockWindowUpdate 0
    Exit Sub
errH:
	Call ShowFlash("")
    MsgBox err.Description
End Sub


