VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOpenDicomDir 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��DICOMDIR"
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
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdOpenFile 
      Caption         =   "ѡ��DicomDir"
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
      Caption         =   "ȡ��"
      Height          =   350
      Left            =   6120
      TabIndex        =   2
      Top             =   5400
      Width           =   1275
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "��ͼ��"
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
    
    '�ȴ�DICOMDIR����ȡ����Ҫ�򿪵�ȫ��ͼ���ļ��б�
    ReDim strFileList(0) As String
    Call subGetFileNameList(treDicomDir.SelectedItem)
    '�����б��е�ͼ��ͬһ��Ŀ¼�е�ͼ��һ���
    
    '��ȡͼ���ļ�·�����ж��Ƿ�ͬһ��Ŀ¼�е�ͼ��
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
            '�ȵ���ͼ��򿪵Ĺ���
            Call subOpenFileList(f, arrFileList)
            
            '�ټ�����ȡͼ���б�
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
'���ܣ���TreeView�б�ѡ�еĽڵ����ȡ�ýڵ��¼���ε������ļ����б�
'      ����һ���ݹ麯����ѭ����nodeSelected�ڵ�Ϊ���������¼��ڵ㡣
'������nodeSelected����Ϊ���ڵ㣬ѭ�����µ����нڵ�
'���أ���
'�ϼ���������̣�cmdOpen_Click
'�¼���������̣�subGetFileNameList
'���õ��ⲿ������ֱ���޸�strFileList
'�����ˣ��ƽ�
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
    dlgOpenDicomDir.ShowOpen       '��ѡ��dicomdir�ļ��ĶԻ���
    If dlgOpenDicomDir.Filename = "" Then Exit Sub
    '' -��ȡdicomdir��·��,��ΪDicomDir�Ĺ̶�����һ����dicomdir�����ֱ�Ӽ�9
    On Error GoTo errFileName
    strRoot = left(dlgOpenDicomDir.Filename, Len(dlgOpenDicomDir.Filename) - 9)
    Set dsetsOpen = New DicomDataSets
    
    Set dsetOpen = dsetsOpen.ReadDirectory(dlgOpenDicomDir.Filename)
    dlgOpenDicomDir.Filename = ""
    treDicomDir.Nodes.Clear
    On Error GoTo 0
    For Each dsetChild1 In dsetOpen.Children
        imcount = 0
        Set node1 = treDicomDir.Nodes.Add(, , , "������" & dsetChild1.Name)
        node1.Tag = "PATIENT"
        For Each dsetChild2 In dsetChild1.Children
            strDescription = dsetChild2.StudyDescription
            If strDescription = "" Then
                strDescription = "���"
            Else
                strDescription = "��飺" & strDescription
            End If
            Set node2 = treDicomDir.Nodes.Add(node1, tvwChild, , strDescription)
            node2.Tag = "STUDY"
            For Each dsetChild3 In dsetChild2.Children
                strDescription = dsetChild3.SeriesDescription
                If strDescription = "" Then
                    strDescription = "����"
                Else
                    strDescription = "���У�" & strDescription
                End If
                Set node3 = treDicomDir.Nodes.Add(node2, tvwChild, , strDescription)
                node3.Tag = "SERIES"
                For Each dsetChild4 In dsetChild3.Children
                    varPath = dsetChild4.Attributes(4, &H1500)       'Referenced File ID�����ļ���
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
        node1.Text = node1.Text & "  (" & imcount & " ��ͼ��)"
    Next
    Exit Sub
errFileName:
    MsgBox "ѡ����ļ����󣬲���DICOMDIR���͵��ļ���������ѡ��", vbExclamation, gstrSysName
End Sub

