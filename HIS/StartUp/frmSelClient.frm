VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSelClient 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��ǰλ��"
   ClientHeight    =   3012
   ClientLeft      =   36
   ClientTop       =   336
   ClientWidth     =   4512
   Icon            =   "frmSelClient.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3012
   ScaleWidth      =   4512
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3060
      TabIndex        =   1
      Top             =   2580
      Width           =   1230
   End
   Begin MSComctlLib.ImageList imgMain 
      Left            =   1245
      Top             =   2355
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelClient.frx":1CFA
            Key             =   "dep"
            Object.Tag             =   "dep"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwMain 
      Height          =   1980
      Left            =   120
      TabIndex        =   0
      Top             =   540
      Width           =   4200
      _ExtentX        =   7408
      _ExtentY        =   3493
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "imgMain"
      SmallIcons      =   "imgMain"
      ColHdrIcons     =   "imgMain"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "��������"
         Object.Width           =   6704
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "վ����"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "��ѡ���㵱ǰ�����λ�����ڵĲ��ţ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.8
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   165
      TabIndex        =   2
      Top             =   240
      Width           =   3570
   End
End
Attribute VB_Name = "frmSelClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mstr���� As String
Dim mstr���� As String
Dim mstrCurIndex As String
Public gstrվ�� As String
Public gstrCurվ�� As String

Private Sub cmdOK_Click()
    If lvwMain.ListItems.Count <> 0 Then
    If ObjPtr(lvwMain.SelectedItem) = 0 Then
        If lvwMain.Enabled Then lvwMain.SetFocus
    End If
    
    gstrվ�� = ""
    gstrCurվ�� = ""
    With lvwMain.SelectedItem
        gstrվ�� = .SubItems(1)
        gstrCurվ�� = .Text

        If gstrվ�� = "" Then
            MsgBox "��ѡ��һ����������ڵĲ���!", vbInformation, "��ʾ"
            If lvwMain.Enabled Then lvwMain.SetFocus
            Exit Sub
        End If
    End With
    End If
    Unload Me
End Sub

Private Sub Form_Activate()
    '����ͷ��Ϣ
    Call LoadListView(mstr����, mstr����, mstrCurIndex)
    
End Sub

Public Sub ShowEdit(ByVal str���� As String, ByVal str���� As String, ByVal strCurIndex As String)
    '--���ܣ���ʾѡ������λ�����ڲ���
    mstr���� = str����
    mstr���� = str����
    mstrCurIndex = strCurIndex
    Me.Show 1
End Sub

Private Sub LoadListView(ByVal str����, str����, strCurIndex As String)
    Dim i As Integer
    Dim strSplit����() As String, strSplit����() As String
    Dim mList As MSComctlLib.ListItem
    On Error Resume Next
    With lvwMain
        .ListItems.Clear
        strSplit���� = Split(mstr����, ",")
        strSplit���� = Split(mstr����, ",")
        For i = LBound(strSplit����) To UBound(strSplit����)
            Set mList = .ListItems.Add(, , strSplit����(i), "dep", "dep")
            mList.SubItems(1) = strSplit����(i)
        Next
        
        If .Enabled Then .SetFocus
        
        If lvwMain.ListItems.Count > 0 Then
            If strCurIndex = "" Then
                lvwMain.ListItems(1).Selected = True
            Else
                lvwMain.ListItems(1).Selected = True
                For i = 1 To lvwMain.ListItems.Count
                    If strCurIndex = lvwMain.ListItems(i).SubItems(1) Then
                        lvwMain.ListItems(i).Selected = True
                        Exit For
                    End If
                Next
            End If
        End If
    End With
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdOK_Click
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        If lvwMain.ListItems.Count <> 0 Then
            If ObjPtr(lvwMain.SelectedItem) = 0 Then
                If lvwMain.Enabled Then lvwMain.SetFocus
            End If
            
            gstrվ�� = ""
            gstrCurվ�� = ""
            With lvwMain.SelectedItem
                gstrվ�� = .SubItems(1)
                gstrCurվ�� = .Text
        
                If gstrվ�� = "" Then
                    MsgBox "��ѡ��һ����������ڵĲ���!", vbInformation, "��ʾ"
                    If lvwMain.Enabled Then lvwMain.SetFocus
                    Cancel = 1
                End If
            End With
        End If
    End If
End Sub

Private Sub lvwMain_DblClick()
    cmdOK_Click
End Sub
