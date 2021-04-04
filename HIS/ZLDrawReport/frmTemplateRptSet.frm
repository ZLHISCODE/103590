VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmTemplateRptSet 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "选项"
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6150
   Icon            =   "frmTemplateRptSet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   3795
      TabIndex        =   0
      Top             =   5985
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4950
      TabIndex        =   1
      Top             =   5970
      Width           =   1100
   End
   Begin XtremeSuiteControls.TabControl tbc 
      Height          =   5790
      Left            =   75
      TabIndex        =   2
      Top             =   105
      Width           =   5970
      _Version        =   589884
      _ExtentX        =   10530
      _ExtentY        =   10213
      _StockProps     =   64
   End
End
Attribute VB_Name = "frmTemplateRptSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'######################################################################################################################

Private mblnOK As Boolean
Private mblnModifyPaper As Boolean
Private mintPaperSize As Integer
Private mfrmMain As Object
Private mstrSavePath As String

Private mfrmChildPrintSet As frmChildPrintSet

'######################################################################################################################

Public Function ShowDialog(ByVal frmMain As Object, Optional ByVal intPaperSize As Integer, Optional ByVal blnModifyPaper As Boolean = True, Optional ByVal strSavePath As String) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    
    mblnOK = False
    mintPaperSize = intPaperSize
    mblnModifyPaper = blnModifyPaper
    
    mstrSavePath = strSavePath
    If mstrSavePath = "" Then mstrSavePath = "私有模块\ZLHIS\" & App.ProductName & "\固定报告"
    
    Set mfrmMain = frmMain
    
    If ExecuteCommand("初始控件") = False Then Exit Function
    If ExecuteCommand("刷新数据") = False Then Exit Function
    
    Call ExecuteCommand("控件状态")
    
    DataChanged = False
    
    Me.Show 1, frmMain
    
    ShowDialog = mblnOK
    
End Function

Private Function ExecuteCommand(ByVal strCommand As String) As Boolean
    '--------------------------------------------------------------------------------------------------------------
    '功能：
    '参数：
    '返回：
    '--------------------------------------------------------------------------------------------------------------
    Dim intLoop As Integer
    Dim strSQL As String
    Dim strTmp As String
    Dim varTmp As Variant
    
    On Error GoTo errHand

    Select Case strCommand
    '--------------------------------------------------------------------------------------------------------------
    Case "初始控件"
        With tbc
            With .PaintManager
                .Appearance = xtpTabAppearancePropertyPage2003
                .BoldSelected = True
                .Color = xtpTabColorDefault
                .ColorSet.ButtonSelected = &H80000005
                .ShowIcons = True
            End With
            
            Set mfrmChildPrintSet = New frmChildPrintSet
            Call mfrmChildPrintSet.InitData(mfrmMain, "")
            
            .InsertItem 0, "打印 ", mfrmChildPrintSet.hWnd, 0
            .Item(0).Selected = True
        End With

        
    '--------------------------------------------------------------------------------------------------------------
    Case "控件状态"
    

        
    '--------------------------------------------------------------------------------------------------------------
    Case "刷新数据"
        
        If Not mfrmChildPrintSet Is Nothing Then Call mfrmChildPrintSet.RefreshData
        
    '--------------------------------------------------------------------------------------------------------------
    Case "校验数据"
        
        If Not mfrmChildPrintSet Is Nothing Then Call mfrmChildPrintSet.ValidData
        
    '--------------------------------------------------------------------------------------------------------------
    Case "保存数据"
        
        If Not mfrmChildPrintSet Is Nothing Then Call mfrmChildPrintSet.SaveData
        
    End Select

    ExecuteCommand = True

    Exit Function
    
    '
    '----------------------------------------------------------------------------------------------------------
errHand:

    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog

End Function

'######################################################################################################################

Private Property Let DataChanged(ByVal blnData As Boolean)

    mfrmChildPrintSet.DataChanged = blnData
            
End Property

Private Property Get DataChanged() As Boolean
    If Not (mfrmChildPrintSet Is Nothing) Then
        DataChanged = mfrmChildPrintSet.DataChanged
    End If
End Property

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    
    If ExecuteCommand("校验数据") Then
        Call ExecuteCommand("保存数据")
    End If

    mblnOK = True
    
    Unload Me
End Sub

