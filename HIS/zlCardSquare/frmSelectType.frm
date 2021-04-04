VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSelectType 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "ҽ�ƿ����ѡ��"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5085
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   5085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin MSComctlLib.ImageList ilt16 
      Left            =   3825
      Top             =   2715
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelectType.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3645
      TabIndex        =   1
      Top             =   225
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3645
      TabIndex        =   0
      Top             =   735
      Width           =   1100
   End
   Begin MSComctlLib.ListView lvwSel 
      Height          =   4755
      Left            =   15
      TabIndex        =   2
      Top             =   60
      Width           =   3510
      _ExtentX        =   6191
      _ExtentY        =   8387
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ilt16"
      SmallIcons      =   "ilt16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "����"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "����"
         Object.Width           =   4304
      EndProperty
   End
End
Attribute VB_Name = "frmSelectType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrCardTypeIDs As String
Private mlngCardTypeID As Long
Private mblnOK As Boolean, mblnFirst As Boolean
Private mcnOracle As ADODB.Connection

Public Function zlSelect(ByVal frmMain As Object, _
    ByVal strCardTypeIDs As String, ByRef lngCardTypeID As Long, _
    Optional strFromCaption As String = "", Optional cnOracle As ADODB.Connection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ѡ��ָ����ҽ�ƿ����
    '���:strCardTypeIDs-��,��������ҽ�ƿ�;����ָ����ҽ�ƿ�
    '       strFromCaption-���������Ĵ������
    '����:lngCardTypeID-��ǰѡ���ҽ�ƿ�
    '����:ѡ��ɹ�������true,���򷵻�False
    '����:���˺�
    '����:2012-03-19 10:23:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    mstrCardTypeIDs = strCardTypeIDs: mlngCardTypeID = 0
    Set mcnOracle = cnOracle
    mblnOK = False
    If strFromCaption <> "" Then Me.Caption = strFromCaption
    If frmMain Is Nothing Then
        Me.Show 1
    Else
        Me.Show 1, frmMain
    End If
    lngCardTypeID = mlngCardTypeID
    zlSelect = mblnOK
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function LoadData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ҽ�ƿ����
    '����:���سɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2012-03-19 10:26:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String, objItem As Object
    Dim objDatabase As New clsDatabase
    On Error GoTo errHandle
    

    If mstrCardTypeIDs <> "" Then
        strSQL = "Select /*+ rule */ A.ID,A.����,A.���� From ҽ�ƿ���� A,Table(f_Str2list([1])) J  Where A.ID=J.Column_Value"
    Else
        strSQL = "Select   A.ID,A.����,A.���� From ҽ�ƿ���� A  Where nvl(�Ƿ�����,0)=1"
    End If
    If Not mcnOracle Is Nothing Then
        objDatabase.InitCommon gcnOracle
        Set rsTemp = objDatabase.OpenSQLRecord(strSQL, "ѡ��ҽ�ƿ����", mstrCardTypeIDs)
    Else
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "ѡ��ҽ�ƿ����", mstrCardTypeIDs)
    End If
    
    With lvwSel
        .ListItems.Clear
        Do While Not rsTemp.EOF
            Set objItem = .ListItems.Add(, "K" & rsTemp!id, Nvl(rsTemp!����), 1, 1)
           objItem.SubItems(1) = Nvl(rsTemp!����)
           If .SelectedItem Is Nothing Then objItem.Selected = True
            rsTemp.MoveNext
        Loop
    End With
    LoadData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub cmdCancel_Click()
    mblnOK = False: mlngCardTypeID = 0
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If lvwSel.SelectedItem Is Nothing Then Exit Sub
    mlngCardTypeID = Val(Mid(lvwSel.SelectedItem.Key, 2))
    mblnOK = True: Unload Me
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    If LoadData = False Then Unload Me: Exit Sub
    If lvwSel.Enabled Then lvwSel.SetFocus
End Sub
Private Sub Form_Load()
    mblnFirst = True
End Sub

Private Sub lvwSel_DblClick()
    cmdOK_Click
End Sub

 Private Sub lvwSel_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then cmdOK_Click
End Sub
