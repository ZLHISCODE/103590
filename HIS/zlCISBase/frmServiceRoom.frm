VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmServiceRoom 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�ⷿѡ��"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4560
   Icon            =   "frmServiceRoom.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   4560
   StartUpPosition =   2  '��Ļ����
   Begin VB.CheckBox chk�ⷿ 
      Appearance      =   0  'Flat
      Caption         =   "ȫѡ"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   3785
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   120
      Width           =   675
   End
   Begin VB.CommandButton cmdȡ�� 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3360
      TabIndex        =   2
      Top             =   2640
      Width           =   1100
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2160
      TabIndex        =   1
      Top             =   2640
      Width           =   1100
   End
   Begin MSComctlLib.ListView lvw�洢�ⷿ 
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   3836
      View            =   2
      Arrange         =   1
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "img16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "����"
         Object.Width           =   3528
      EndProperty
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServiceRoom.frx":6852
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServiceRoom.frx":D0B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServiceRoom.frx":13916
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmServiceRoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrҩƷ���� As String
Private mstr�洢�ⷿ As String
Private mstr�洢�ⷿID As String
Private mstrArr�洢�ⷿ() As String
Private mstrStationNo As String
Private mstrPrivs As String
Private mbln��ҩ��ҩ�����ʲ��� As Boolean

Private Sub Init�洢�ⷿ()
    Dim rsTemp As New ADODB.Recordset
    Dim rsOther As New ADODB.Recordset
    Const str��ҩ As String = "'��ҩ%'"
    Const str��ҩ As String = "'��ҩ%'"
    Const str��ҩ As String = "'��ҩ%'"
    Dim mstrȫ���ⷿID As String
    Dim dbl���пⷿ As Boolean
    
    On Error GoTo ErrHandle
    
    If InStr(1, ";" & mstrPrivs & ";", ";���пⷿ;") > 0 Then dbl���пⷿ = True
    
    '����ҩƷ����;������ȡ������洢�Ŀⷿ
    gstrSql = " Select ID,����,���� From ���ű� " & _
              " Where ID in (select distinct ����id from ��������˵�� where �������� like "
    If mstrҩƷ���� = "����ҩ" Then
        gstrSql = gstrSql & str��ҩ
    ElseIf mstrҩƷ���� = "�г�ҩ" Then
        gstrSql = gstrSql & str��ҩ
    Else
        gstrSql = gstrSql & str��ҩ
    End If
    gstrSql = gstrSql & " or ��������='�Ƽ���')"
    
    gstrSql = gstrSql & " and (����ʱ�� is null or ����ʱ��=to_date('3000-01-01','YYYY-MM-DD')) "
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "����ҩƷ����;������ȡ������洢�Ŀⷿ(�����ⷿ)")
    mstrȫ���ⷿID = ""
    Do While Not rsTemp.EOF
        mstrȫ���ⷿID = mstrȫ���ⷿID & "," & rsTemp!ID
        rsTemp.MoveNext
    Loop
    If mstrȫ���ⷿID <> "" Then
        mbln��ҩ��ҩ�����ʲ��� = False
    Else
        mbln��ҩ��ҩ�����ʲ��� = True
        Exit Sub
    End If
    
    If Not dbl���пⷿ Then
        'ȡ��ǰ�û������ⷿ
        gstrSql = gstrSql & " And Id In(Select ����ID From ������Ա Where ��Աid=[1]) "
    End If
    
    gstrSql = gstrSql & "order by id"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "����ҩƷ����;������ȡ������洢�Ŀⷿ", UserInfo.ID)
    
    lvw�洢�ⷿ.ListItems.Clear

    With rsTemp
        Do While Not .EOF
            lvw�洢�ⷿ.ListItems.Add , "K" & !ID, !����, , 2
            .MoveNext
        Loop
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdOK_Click()
    Dim intItem As Integer
    Dim intItems As Integer
    'ȡ�ô洢�ⷿ
    mstr�洢�ⷿID = ""
    mstr�洢�ⷿ = ""
    intItems = Me.lvw�洢�ⷿ.ListItems.Count
    For intItem = 1 To intItems
        If lvw�洢�ⷿ.ListItems(intItem).Checked Then
            mstr�洢�ⷿID = mstr�洢�ⷿID & "!!" & Mid(lvw�洢�ⷿ.ListItems(intItem).Key, 2) & "|"
            mstr�洢�ⷿ = mstr�洢�ⷿ & "|" & Mid(lvw�洢�ⷿ.ListItems(intItem).Text, 1)
        End If
    Next
    mstr�洢�ⷿ = Mid(mstr�洢�ⷿ, 2)
    mstr�洢�ⷿID = Mid(mstr�洢�ⷿID, 3)
    Call frmBatchUpdate.ShowRoom(mstr�洢�ⷿ, mstr�洢�ⷿID)
    Unload Me
End Sub

Private Sub cmdȡ��_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call Init�洢�ⷿ
    
    If mbln��ҩ��ҩ�����ʲ��� = True Then
        MsgBox "�������þ���ҩ��ҩ�����ʵĲ��š�", vbInformation, gstrSysName
        Unload Me
        Exit Sub
    End If
    
    Call Change�洢�ⷿ
End Sub

Public Sub ShowMe(ByVal frmParent As Object, ByVal strҩƷ���� As String, ByVal str�洢�ⷿ As String, ByVal strPrivs As String)
    mstrҩƷ���� = strҩƷ����
    mstr�洢�ⷿ = str�洢�ⷿ
    mstrPrivs = strPrivs

    Me.Show 1, frmParent
End Sub
Private Sub chk�ⷿ_Click()
'�ⷿȫѡ��ť
    If chk�ⷿ.Value = 2 Then Exit Sub
    Call SetSelect(lvw�洢�ⷿ, chk�ⷿ.Value)
End Sub
Private Sub SetSelect(ByVal lvwObj As Object, Optional ByVal BlnSelect As Boolean = True)
'ȫѡ����
    Dim intSelect As Integer
    With lvwObj
        For intSelect = 1 To .ListItems.Count
            .ListItems(intSelect).Checked = BlnSelect
        Next
    End With
End Sub
Private Sub lvw�洢�ⷿ_ItemCheck(ByVal Item As MSComctlLib.ListItem)
'����ѡ��Ĵ洢�ⷿ
    Call ItemCheck(lvw�洢�ⷿ, Item, chk�ⷿ)
End Sub
Private Sub ItemCheck(ByVal lvwObj As Object, ByVal Item As MSComctlLib.ListItem, ByVal chkObj As CheckBox)
'��¼ѡ��Ŀⷿ
    Dim lngCheck As Long, blnCheck As Boolean, intCount As Integer
    
    intCount = 0
    With lvwObj
        For lngCheck = 1 To .ListItems.Count
            If .ListItems(lngCheck).Checked = True Then
                intCount = intCount + 1
            End If
        Next
        
        If intCount = lvwObj.ListItems.Count Then
            chkObj.Value = 1
        ElseIf intCount > 0 Then
            chkObj.Value = 2
        Else
            chkObj.Value = 0
        End If
    End With
End Sub

Private Sub Change�洢�ⷿ()
    Dim i As Integer, j As Integer
    Dim intSelect As Integer
    mstrArr�洢�ⷿ = Split(mstr�洢�ⷿ, "|")
    
    For i = LBound(mstrArr�洢�ⷿ) To UBound(mstrArr�洢�ⷿ)
        For intSelect = 1 To lvw�洢�ⷿ.ListItems.Count
            If mstrArr�洢�ⷿ(i) = lvw�洢�ⷿ.ListItems(intSelect).Text Then
                lvw�洢�ⷿ.ListItems(intSelect).Checked = True
                j = j + 1
            End If
        Next
    Next
    
    If j = lvw�洢�ⷿ.ListItems.Count Then
        chk�ⷿ.Value = 1
    ElseIf j > 0 And j < lvw�洢�ⷿ.ListItems.Count Then
        chk�ⷿ.Value = 2
    End If
End Sub


