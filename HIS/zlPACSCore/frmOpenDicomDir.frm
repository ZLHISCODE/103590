VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOpenDicomDir 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "打开DICOMDIR"
   ClientHeight    =   5955
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8700
   Icon            =   "frmOpenDicomDir.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   8700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdOpenFile 
      Caption         =   "选择DicomDir"
      Height          =   350
      Left            =   1320
      TabIndex        =   3
      Top             =   5400
      Width           =   1275
   End
   Begin MSComDlg.CommonDialog dlgOpenDicomDir 
      Left            =   720
      Top             =   5280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消"
      Height          =   350
      Left            =   6120
      TabIndex        =   2
      Top             =   5400
      Width           =   1275
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "打开图像"
      Height          =   350
      Left            =   3720
      TabIndex        =   1
      Top             =   5400
      Width           =   1275
   End
   Begin MSComctlLib.TreeView treDicomDir 
      Height          =   4815
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   8493
      _Version        =   393217
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
   End
End
Attribute VB_Name = "frmOpenDicomDir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public f As frmViewer
Private strFileList() As String

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOpen_Click()
    Dim arrFileList As OpenFileArray
    Dim objFile As New Scripting.FileSystemObject
    Dim i As Integer
    Dim strPath As String
    Dim iImageCount As Integer
    
    On Error GoTo err
    
    If treDicomDir.SelectedItem Is Nothing Then Exit Sub
    
    '先从DICOMDIR中提取本次要打开的全部图像文件列表
    ReDim strFileList(0) As String
    Call subGetFileNameList(treDicomDir.SelectedItem)
    '处理列表中的图像，同一个目录中的图像一起打开
    
    '提取图像文件路径，判断是否同一个目录中的图像
    ReDim arrFileList.Filename(0)
    For i = 1 To UBound(strFileList)
        strPath = objFile.GetParentFolderName(strFileList(i)) & "\"
        If arrFileList.FilePath = "" Then
            arrFileList.FilePath = strPath
            ReDim arrFileList.Filename(1) As String
            arrFileList.Filename(1) = objFile.GetFileName(strFileList(i))
        ElseIf arrFileList.FilePath = strPath Then
            iImageCount = UBound(arrFileList.Filename)
            ReDim Preserve arrFileList.Filename(iImageCount + 1) As String
            arrFileList.Filename(iImageCount + 1) = objFile.GetFileName(strFileList(i))
        Else
            '先调用图像打开的过程
            Call subOpenFileList(f, arrFileList)
            
            '再继续读取图像列表
            arrFileList.FilePath = strPath
            ReDim arrFileList.Filename(1) As String
            arrFileList.Filename(1) = objFile.GetFileName(strFileList(i))
        End If
    Next i
    
    If arrFileList.FilePath <> "" Then
        Call subOpenFileList(f, arrFileList)
    End If
    
    Unload Me
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub subGetFileNameList(nodeSelected As Node)
'------------------------------------------------
'功能：从TreeView中被选中的节点里，读取该节点下级层次的所有文件名列表。
'      这是一个递归函数，循环以nodeSelected节点为根的所有下级节点。
'参数：nodeSelected－做为根节点，循环其下的所有节点
'返回：无
'上级函数或过程：cmdOpen_Click
'下级函数或过程：subGetFileNameList
'引用的外部参数：直接修改strFileList
'编制人：黄捷
'------------------------------------------------
    Dim nodeChild As Node
    Dim iCount As Integer
    Dim i As Integer
    If (nodeSelected.Children > 0) Then
        Set nodeChild = nodeSelected.Child
        For i = 1 To nodeSelected.Children
            subGetFileNameList nodeChild
            Set nodeChild = nodeChild.Next
        Next
    Else
        iCount = UBound(strFileList)
        ReDim Preserve strFileList(iCount + 1) As String
        strFileList(iCount + 1) = nodeSelected.Text
    End If
End Sub

Private Sub cmdOpenFile_Click()
    Form_Load
End Sub

Private Sub Form_Load()
    Dim strRoot As String
    Dim dsetsOpen As DicomDataSets
    Dim dsetOpen As DicomDataSet
    Dim dsetChild1 As DicomDataSet
    Dim dsetChild2 As DicomDataSet
    Dim dsetChild3 As DicomDataSet
    Dim dsetChild4 As DicomDataSet
    Dim node1 As Node
    Dim node2 As Node
    Dim node3 As Node
    Dim node4 As Node
    Dim imcount As Integer
    Dim strDescription As String
    Dim varPath As Variant
    Dim strPath As String
    Dim i As Integer
    
    
    dlgOpenDicomDir.Filter = "DICOMDIR|DICOMDIR"
    dlgOpenDicomDir.ShowOpen       '打开选择dicomdir文件的对话框
    If dlgOpenDicomDir.Filename = "" Then Exit Sub
    '' -获取dicomdir的路径,因为DicomDir的固定名字一定是dicomdir，因此直接减9
    On Error GoTo errFileName
    strRoot = left(dlgOpenDicomDir.Filename, Len(dlgOpenDicomDir.Filename) - 9)
    Set dsetsOpen = New DicomDataSets
    
    Set dsetOpen = dsetsOpen.ReadDirectory(dlgOpenDicomDir.Filename)
    dlgOpenDicomDir.Filename = ""
    treDicomDir.Nodes.Clear
    On Error GoTo 0
    For Each dsetChild1 In dsetOpen.Children
        imcount = 0
        Set node1 = treDicomDir.Nodes.Add(, , , "姓名：" & dsetChild1.Name)
        node1.Tag = "PATIENT"
        For Each dsetChild2 In dsetChild1.Children
            strDescription = dsetChild2.StudyDescription
            If strDescription = "" Then
                strDescription = "检查"
            Else
                strDescription = "检查：" & strDescription
            End If
            Set node2 = treDicomDir.Nodes.Add(node1, tvwChild, , strDescription)
            node2.Tag = "STUDY"
            For Each dsetChild3 In dsetChild2.Children
                strDescription = dsetChild3.SeriesDescription
                If strDescription = "" Then
                    strDescription = "序列"
                Else
                    strDescription = "序列：" & strDescription
                End If
                Set node3 = treDicomDir.Nodes.Add(node2, tvwChild, , strDescription)
                node3.Tag = "SERIES"
                For Each dsetChild4 In dsetChild3.Children
                    varPath = dsetChild4.Attributes(4, &H1500)       'Referenced File ID，即文件名
                    strPath = strRoot
                    For i = 1 To UBound(varPath)
                        If varPath(i) <> "" Then strPath = strPath & "\" & varPath(i)
                    Next
                    Set node4 = treDicomDir.Nodes.Add(node3, tvwChild, , strPath)
                    node4.Tag = strPath
                    imcount = imcount + 1
                Next
            Next
        Next
        node1.Text = node1.Text & "  (" & imcount & " 幅图像)"
    Next
    Exit Sub
errFileName:
    MsgBox "选择的文件错误，不是DICOMDIR类型的文件。请重新选择。", vbExclamation, gstrSysName
End Sub

