VERSION 5.00
Begin VB.Form frmSave 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "保存图像到服务器"
   ClientHeight    =   3330
   ClientLeft      =   30
   ClientTop       =   510
   ClientWidth     =   4860
   Icon            =   "frmSave.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   4860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame2 
      Caption         =   "保存类型"
      Height          =   816
      Left            =   264
      TabIndex        =   7
      Top             =   2268
      Width           =   3060
      Begin VB.CheckBox chkSave 
         Caption         =   "报告图像"
         Height          =   276
         Index           =   1
         Left            =   1644
         TabIndex        =   9
         Top             =   336
         Width           =   1236
      End
      Begin VB.CheckBox chkSave 
         Caption         =   "DICOM图像"
         Height          =   276
         Index           =   0
         Left            =   252
         TabIndex        =   8
         Top             =   324
         Value           =   1  'Checked
         Width           =   1236
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   375
      Left            =   3480
      TabIndex        =   6
      Top             =   960
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "保存(&O)"
      Default         =   -1  'True
      Height          =   375
      Left            =   3480
      TabIndex        =   5
      Top             =   360
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Caption         =   "保存选项"
      Height          =   1920
      Left            =   276
      TabIndex        =   0
      Top             =   216
      Width           =   3048
      Begin VB.OptionButton OptSave 
         Caption         =   "保存当前图像"
         Height          =   264
         Index           =   3
         Left            =   228
         TabIndex        =   4
         Top             =   1440
         Value           =   -1  'True
         Width           =   1776
      End
      Begin VB.OptionButton OptSave 
         Caption         =   "保存当前序列选择的图像"
         Height          =   264
         Index           =   2
         Left            =   228
         TabIndex        =   3
         Top             =   1056
         Width           =   2736
      End
      Begin VB.OptionButton OptSave 
         Caption         =   "保存所有选择的图像"
         Height          =   264
         Index           =   1
         Left            =   228
         TabIndex        =   2
         Top             =   684
         Width           =   2016
      End
      Begin VB.OptionButton OptSave 
         Caption         =   "保存所有的图像"
         Height          =   264
         Index           =   0
         Left            =   228
         TabIndex        =   1
         Top             =   348
         Width           =   1776
      End
   End
End
Attribute VB_Name = "frmSave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public f As frmViewer

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim imgs As New DicomImages
    Dim im As DicomImage
    Dim v As DicomViewer
    Dim i As Integer
    If f.intSelectedSerial = 0 Then Exit Sub
    If chkSave(0) = 0 And chkSave(1) = 0 Then
        MsgBox "至少必须选择一个保存的文件类型！", vbInformation, gstrSysName
        Exit Sub
    End If
    '''''''''''''''''''[检查是否有重复的序列，并提示]'''''''''''''''''''''''''''''''
    If funIsRepeatSerial(f) Then
        If MsgBox("当前操作中有重复序列，保存可能导致一些有用信息的丢失，是否继续？", vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
    End If
    If f.MSFViewer.TextMatrix(f.intSelectedSerial, 0) <> 0 And (OptSave(3) Or OptSave(2)) Then
        MsgBox "当前序列不是来自影像服务器的图像，不能保存！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If OptSave(3) Then
        imgs.Add f.SelectedImage
        subLabelCopyRebuild f.SelectedImage, imgs(imgs.Count)
        
    ElseIf OptSave(2) Then
        For Each im In f.viewer(f.intSelectedSerial).Images
            If im.Tag <> "" Then
                imgs.Add im
                subLabelCopyRebuild im, imgs(imgs.Count)
            End If
        Next
    ElseIf OptSave(1) Or OptSave(0) Then
        For Each v In f.viewer
            If v.Index <> 0 And Val(f.MSFViewer.TextMatrix(v.Index, 0)) = 0 Then
                For Each im In v.Images
                    If im.Tag <> "" Or OptSave(0) Then
                        imgs.Add im
                        subLabelCopyRebuild im, imgs(imgs.Count)
                    End If
                Next
            End If
        Next
    End If
    For Each im In imgs
        subSaveLabelToImg im
    Next
    If chkSave(0) = 1 And chkSave(1) = 1 Then
        i = 2
    ElseIf chkSave(1) = 1 Then
        i = 1
    Else
        i = 0
    End If
    SaveImages imgs, i
    Unload Me
End Sub

Function funIsRepeatSerial(f As frmViewer) As Boolean
    Dim i As Integer, j  As Integer
    funIsRepeatSerial = False
    For i = 1 To f.MSFViewer.Rows - 2
        For j = i + 1 To f.MSFViewer.Rows - 1
            If f.MSFViewer.TextMatrix(i, 2) = f.MSFViewer.TextMatrix(j, 2) Then
                funIsRepeatSerial = True
                Exit Function
            End If
        Next
    Next
End Function

