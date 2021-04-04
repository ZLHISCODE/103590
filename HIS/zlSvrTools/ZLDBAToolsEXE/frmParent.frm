VERSION 5.00
Begin VB.Form frmParent 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "�����Ż�����"
   ClientHeight    =   6555
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   10500
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "frmParent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mobjTools As Object
Private mfrmTmp As Form
Attribute mfrmTmp.VB_VarHelpID = -1
Private mcolFrmList As New Collection

Private Sub Form_load()

    Set mobjTools = CreateObject("zlDbaTools.clsToolsMain")
    If mobjTools Is Nothing Then
        MsgBox "��ʼ��ʧ�ܣ�����zlDbaTools.dll�Ƿ�ɹ�ע�ᡣ"
        ShowFlash ""
        Exit Sub
    End If
    
    '�״μ��أ��������ݿ�����
    ShowForm "0601"
End Sub

Private Sub Form_Resize()
    mfrmTmp.WindowState = 0
    mfrmTmp.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
    Me.Refresh
End Sub

Public Sub ShowForm(ByVal strMoudleNum As String)
    Dim frmNew As Form
    Dim strFormName As String, strTmp As String
    
    If mobjTools Is Nothing Then
        Exit Sub
    End If
    
    '��ѡ�����ǵ�ǰ���壬�����ٴμ���
    '0601-���ݿ�����  0602-SQL����    0604-�Ự����   0605-�������   0606-�ռ�����
    Select Case strMoudleNum
        Case "0601"
            strFormName = "frmMonitorMain"
            strTmp = "���ڼ������ݿ����ܷ�������..."
        Case "0602"
            strFormName = "frmTunning"
            strTmp = "���ڼ���SQL���ܷ������Ż�����..."
        Case "0604"
            strFormName = "frmKillBlockers"
            strTmp = "���ڼ��ػỰ��������..."
        Case "0605"
            strFormName = "frmIdxInfo"
            strTmp = "���ڼ��������������..."
        Case "0606"
            strFormName = "frmReused"
            strTmp = "���ڼ��ؿռ������..."
    End Select
    
    On Error Resume Next
    If Not mfrmTmp Is Nothing Then
        If mfrmTmp.Name = strFormName Then Exit Sub
        mfrmTmp.Visible = False
    End If
    
    Set frmNew = mcolFrmList.Item(strFormName)
    
    ShowFlash strTmp
    If frmNew Is Nothing Then
        On Error GoTo errH
        
        Set mfrmTmp = mobjTools.GetFrmByMdoudle(strMoudleNum, gblnDBA, gcnOracle, gstrUserName, gstrPassword)
        If mfrmTmp Is Nothing Then
            MsgBox "�������ʧ�ܣ���ʹ��DBA�û���¼��"
            ShowFlash ""
            Exit Sub
        End If
        
        '����Ӧ����һ��ShowMe������
        mcolFrmList.Add mfrmTmp, mfrmTmp.Name
        SetParent mfrmTmp.hwnd, Me.hwnd
        LockWindowUpdate Me.hwnd
        mfrmTmp.ShowMe
    Else
        Set mfrmTmp = mcolFrmList.Item(strFormName)
        mfrmTmp.Visible = True
    End If
    
    Call Form_Resize    '�ü��ع��Ĵ��屣��ԭ��С
    LockWindowUpdate 0
    ShowFlash ""
    Exit Sub
errH:
    LockWindowUpdate Me.hwnd
    ShowFlash ""
    MsgBox Err.Description
    If 0 = 1 Then
        Resume
    End If
End Sub

