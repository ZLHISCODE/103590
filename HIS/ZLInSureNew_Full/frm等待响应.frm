VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm�ȴ���Ӧ 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�ȴ���Ӧ..."
   ClientHeight    =   1785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5355
   ControlBox      =   0   'False
   Icon            =   "frm�ȴ���Ӧ.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   5355
   StartUpPosition =   2  '��Ļ����
   Begin VB.Timer TimeAvi 
      Interval        =   50
      Left            =   2070
      Top             =   660
   End
   Begin VB.Timer TimeSearch 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2070
      Top             =   660
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4050
      TabIndex        =   3
      Top             =   1320
      Width           =   1100
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   75
      Left            =   30
      Picture         =   "frm�ȴ���Ӧ.frx":000C
      ScaleHeight     =   5
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   355
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1140
      Width           =   5325
   End
   Begin MSComCtl2.Animation Avi 
      Height          =   765
      Left            =   120
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   1349
      _Version        =   393216
      AutoPlay        =   -1  'True
      BackColor       =   -2147483643
      FullWidth       =   61
      FullHeight      =   51
   End
   Begin VB.Label LblNote 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "  ���ύ�������ڵȴ�ҽ����������Ӧ..."
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   1380
      TabIndex        =   0
      Top             =   480
      Width           =   3510
   End
End
Attribute VB_Name = "frm�ȴ���Ӧ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mint������ʽ As Integer         '������ʽ
Private mint����Ŀ�� As Integer         '����Ŀ��
Private mlng����ID As Long              '����ID
Private mlng����ID As Long              '����ID
Private mint���� As Integer
Private mblnReturn As Boolean           '�Ƿ�ɹ�����
Private mstrFile As String              '

Private Sub cmdCancel_Click()
    On Error GoTo ErrHand
    Dim objFileSys As New FileSystemObject
    
    TimeSearch.Enabled = False
    'ɾ�������ļ�
    If objFileSys.FileExists(mstrPath_�������� & mint���� & "\" & mstrRequest_��������) Then
        Call objFileSys.DeleteFile(mstrPath_�������� & mint���� & "\" & mstrRequest_��������, True)
    End If
    '�ȼ��Ӧ���ļ��Ƿ���ڣ������������ʾ
    If objFileSys.FileExists(mstrPath_�������� & mint���� & "\" & mstrReply_��������) Then
        If MsgBox("�������Ѿ���Ӧ������ȷ��Ҫ�˳���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            TimeSearch.Enabled = True
            Exit Sub
        End If
        Call objFileSys.DeleteFile(mstrPath_�������� & mint���� & "\" & mstrReply_��������, True)
    End If
    
    Unload Me
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub Form_Activate()
    Dim objFileSys As New FileSystemObject
    
    '���ָ��Ŀ¼�����ڣ��򴴽�
    If mint������ʽ <> ������ʽ.��֤ Then
        If Not objFileSys.FolderExists(mstrPath_�������� & mint����) Then objFileSys.CreateFolder (mstrPath_�������� & mint����)
        LblNote.Caption = "  ������ҽ����������������..."
        If SendRequest(mint������ʽ, mint����Ŀ��, mlng����ID, mlng����ID, mint����) = False Then
            mblnReturn = False
            Unload Me
            Exit Sub
        End If
    End If
    
    TimeSearch.Enabled = True
    LblNote.Caption = "  ���ύ�������ڵȴ�ҽ����������Ӧ..."
End Sub

Private Sub Form_Load()
    Dim strCaption As String
    mstrFile = gstrAviPath & "\" & mstrSearch_��������
    
    LblNote.Caption = "  ���ڼ����ػ���..."
    
    '��������
    Select Case mint����Ŀ��
    Case ����Ŀ��.����
        Select Case mint������ʽ
        Case ������ʽ.�Һ�
            strCaption = "�Һ�����..."
        Case ������ʽ.�շ�
            strCaption = "�����շ�����..."
        Case ������ʽ.����
            strCaption = "סԺ��������..."
        Case ������ʽ.��Ժ
            strCaption = "��Ժ����..."
        Case ������ʽ.��Ժ
            strCaption = "��Ժ����..."
        Case ������ʽ.��¼
            strCaption = "��¼����..."
        End Select
    Case ����Ŀ��.����
        Select Case mint������ʽ
        Case ������ʽ.�Һ�
            strCaption = "�Һų�������..."
        Case ������ʽ.�շ�
            strCaption = "�����շѳ�������..."
        Case ������ʽ.����
            strCaption = "סԺ�����������..."
        Case ������ʽ.��Ժ
            strCaption = "������Ժ����..."
        Case ������ʽ.��Ժ
            strCaption = "������Ժ����..."
        Case ������ʽ.��¼
            strCaption = "�˳�����..."
        End Select
    Case ����Ŀ��.ˢ��
        strCaption = "��ˢ��..."
    End Select
    
    Me.Caption = strCaption
    Call Avi_Play
    mblnReturn = False
End Sub

Private Sub Avi_Play()
    On Error Resume Next
    With Avi
        .Open mstrFile
        .AutoPlay = True
        .Play
    End With
End Sub

Private Sub Avi_Stop()
    Avi.Stop
End Sub

Public Function ShowME(ByVal intInsure As Integer, ByVal ���� As Integer, ByVal Ŀ�� As Integer, Optional ByVal ����ID As Long = 0, _
        Optional ByVal ����ID As Long = 0) As Boolean
    mint������ʽ = ����
    mint����Ŀ�� = Ŀ��
    mlng����ID = ����ID
    mlng����ID = ����ID
    mint���� = intInsure
    Me.Show 1
    ShowME = mblnReturn
End Function

Private Sub TimeSearch_Timer()
    Dim intResult As Integer
    intResult = SearchFile
    If intResult = 0 Then Exit Sub
    
    mblnReturn = (intResult = 1)
    If mblnReturn Then
        If mint����Ŀ�� = ����Ŀ��.ˢ�� And (mint������ʽ = ������ʽ.�Һ� _
        Or mint������ʽ = ������ʽ.�շ� Or mint������ʽ = ������ʽ.��Ժ Or mint������ʽ = ������ʽ.��֤) Then
            mblnReturn = frmIdentify��������.ShowCard(��ȡ����ID(mint����), mint����)
        End If
    End If
    
    TimeSearch.Enabled = False
    Unload Me
    Exit Sub
End Sub

Private Sub TimeAvi_Timer()
    Static i As Long
    i = i + 20
    If i >= Picture1.ScaleWidth Then i = 1
    
    Picture1.PaintPicture Picture1.Picture, i, 0, Picture1.ScaleWidth - i, Picture1.ScaleHeight, 0, 0, Picture1.ScaleWidth - i, Picture1.ScaleHeight
    Picture1.PaintPicture Picture1.Picture, 0, 0, i, Picture1.ScaleHeight, Picture1.ScaleWidth - i, 0, i, Picture1.ScaleHeight
End Sub

Private Function SearchFile() As Integer
    Dim objFileSys As New FileSystemObject
    SearchFile = False
    
    If mint������ʽ <> ������ʽ.��֤ Then
        If Not objFileSys.FileExists(mstrPath_�������� & mint���� & "\" & mstrReply_��������) Then Exit Function
        SearchFile = AnalyseReply(mint������ʽ, mint����Ŀ��, mint����)
    Else
        SearchFile = 1
    End If
End Function
