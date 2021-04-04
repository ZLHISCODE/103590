VERSION 5.00
Begin VB.Form frmParent 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "性能优化工具"
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
        MsgBox "初始化失败，请检查zlDbaTools.dll是否成功注册。"
        ShowFlash ""
        Exit Sub
    End If
    
    '首次加载，加载数据库性能
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
    
    '所选窗体是当前窗体，无需再次加载
    '0601-数据库性能  0602-SQL性能    0604-会话解锁   0605-外键索引   0606-空间整理
    Select Case strMoudleNum
        Case "0601"
            strFormName = "frmMonitorMain"
            strTmp = "正在加载数据库性能分析工具..."
        Case "0602"
            strFormName = "frmTunning"
            strTmp = "正在加载SQL性能分析与优化工具..."
        Case "0604"
            strFormName = "frmKillBlockers"
            strTmp = "正在加载会话解锁工具..."
        Case "0605"
            strFormName = "frmIdxInfo"
            strTmp = "正在加载外键索引工具..."
        Case "0606"
            strFormName = "frmReused"
            strTmp = "正在加载空间管理工具..."
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
            MsgBox "窗体加载失败，请使用DBA用户登录。"
            ShowFlash ""
            Exit Sub
        End If
        
        '窗体应该有一个ShowMe方法。
        mcolFrmList.Add mfrmTmp, mfrmTmp.Name
        SetParent mfrmTmp.hwnd, Me.hwnd
        LockWindowUpdate Me.hwnd
        mfrmTmp.ShowMe
    Else
        Set mfrmTmp = mcolFrmList.Item(strFormName)
        mfrmTmp.Visible = True
    End If
    
    Call Form_Resize    '让加载过的窗体保持原大小
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

