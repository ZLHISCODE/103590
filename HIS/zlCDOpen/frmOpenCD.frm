VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOpenCD 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��DICOMDIR"
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
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "�˳�(&Q)"
      Height          =   350
      Left            =   6120
      TabIndex        =   2
      Top             =   5400
      Width           =   1100
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "��(&O)"
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
Private Const STR_PRIVILEGE = "ͼ���������,ͼ���ע����,������Ƭվ,������"

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
    '����ͼ���·������
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
    MsgBox Err.Description, vbInformation, "��ʾ"
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
'�����ˣ� �ƽ�
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
    
    strRoot = App.Path & IIf(Len(App.Path) > 3, "\DICOM", "DICOM") '����DICOMDIR��·��
    
    Set dsetsOpen = New DicomDataSets
    On Error GoTo errFileName
    Set dsetOpen = dsetsOpen.ReadDirectory(strRoot & "\dicomdir")
    If dsetOpen.Name <> "ZLPACS" Then Exit Sub
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
    MsgBox "ѡ����ļ����󣬲���DICOMDIR���͵��ļ���������ѡ��"
End Sub

