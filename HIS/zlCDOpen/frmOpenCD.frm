VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOpenCD 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "打开DICOMDIR"
   ClientHeight    =   5955
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8700
   Icon            =   "frmOpenCD.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   8700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "退出(&Q)"
      Height          =   350
      Left            =   6120
      TabIndex        =   2
      Top             =   5400
      Width           =   1100
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "打开(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   1650
      TabIndex        =   1
      Top             =   5400
      Width           =   1100
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
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      Appearance      =   1
   End
End
Attribute VB_Name = "frmOpenCD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private strFileList() As String
Private Const STR_PRIVILEGE = "图像操作处理,图像标注测量,独立观片站,单机版"

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOpen_Click()
    ReDim strFileList(0) As String
    Dim o As New clsViewer
    Dim strPath As String
    Dim strTempPath As String
    Dim i As Integer
    Dim j As Integer
    Dim strFileName() As String
    Dim intOpenCount As Integer
    ReDim strFileName(0) As String
    
    If treDicomDir.SelectedItem Is Nothing Then Exit Sub
    subGetFileNameList treDicomDir.SelectedItem
    If UBound(strFileList) < 1 Then Exit Sub
    strPath = Mid(strFileList(1), 1, InStrRev(strFileList(1), "\") - 1)
    
    On Error GoTo OpenObjectError
    '处理图像的路径问题
    For i = 1 To UBound(strFileList)
        strTempPath = Mid(strFileList(i), 1, InStrRev(strFileList(i), "\") - 1)
        If strPath = strTempPath Then
            j = UBound(strFileName) + 1
            ReDim Preserve strFileName(j) As String
            strFileName(j) = Mid(strFileList(i), InStrRev(strFileList(i), "\") + 1)
        Else
            Call o.CallOpenViewerCache(strFileName, Me, strPath, "", STR_PRIVILEGE, , , , , , True)
            intOpenCount = intOpenCount + j
            strPath = strTempPath
            ReDim strFileName(1) As String
            strFileName(1) = Mid(strFileList(i), InStrRev(strFileList(i), "\") + 1)
        End If
    Next i
    
    If intOpenCount <> i Then Call o.CallOpenViewerCache(strFileName, Me, strPath, "", STR_PRIVILEGE, , , , , , True)
    Exit Sub
OpenObjectError:
    MsgBox Err.Description, vbInformation, "提示"
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
'编制人： 黄捷
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
    
    strRoot = App.Path & IIf(Len(App.Path) > 3, "\DICOM", "DICOM") '定义DICOMDIR的路径
    
    Set dsetsOpen = New DicomDataSets
    On Error GoTo errFileName
    Set dsetOpen = dsetsOpen.ReadDirectory(strRoot & "\dicomdir")
    If dsetOpen.Name <> "ZLPACS" Then Exit Sub
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
    MsgBox "选择的文件错误，不是DICOMDIR类型的文件。请重新选择。"
End Sub

