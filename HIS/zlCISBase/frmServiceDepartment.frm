VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmServiceDepartment 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "����ѡ��"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6075
   Icon            =   "frmServiceDepartment.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   6075
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdȡ�� 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4800
      TabIndex        =   7
      Top             =   3960
      Width           =   1100
   End
   Begin VB.PictureBox picDrug 
      Height          =   2940
      Left            =   1320
      ScaleHeight     =   2880
      ScaleWidth      =   4035
      TabIndex        =   2
      Top             =   960
      Visible         =   0   'False
      Width           =   4095
      Begin VB.CommandButton cmdCancel 
         Caption         =   "ȡ��"
         Height          =   350
         Left            =   1200
         TabIndex        =   6
         Top             =   2280
         Width           =   1100
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "ȷ��"
         Height          =   350
         Left            =   120
         TabIndex        =   5
         Top             =   2280
         Width           =   1100
      End
      Begin VB.CheckBox chk���� 
         Appearance      =   0  'Flat
         Caption         =   "ȫѡ"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   3000
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   0
         Width           =   675
      End
      Begin MSComctlLib.ListView lvwItems 
         Height          =   2055
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Visible         =   0   'False
         Width           =   3570
         _ExtentX        =   6297
         _ExtentY        =   3625
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "imgList"
         SmallIcons      =   "imgList"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfDepartment 
      Height          =   3735
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   5655
      _cx             =   9975
      _cy             =   6588
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmServiceDepartment.frx":6852
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   2
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.CommandButton cmdȷ�� 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3600
      TabIndex        =   0
      Top             =   3960
      Width           =   1100
   End
End
Attribute VB_Name = "frmServiceDepartment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstr�洢�ⷿ As String
Private mstr�洢�ⷿID As String
Private mstr������� As String
Private mstr�������ID As String
Private mstr�ⷿ���� As String
Private mstr�ⷿ����ID As String
Private mstrArr�洢�ⷿ() As String
Private mstrArr�洢�ⷿID() As String
Private mstrArr�������() As String
Private mstrArr�ⷿ����ID() As String
Private mrs���� As ADODB.Recordset
Private mstr������� As String

Private Enum mSpecColumn
    �洢�ⷿ = 0
    �洢�ⷿID = 1
    ������� = 2
    �������id = 3
End Enum

Public Sub ShowMe(ByVal frmParent As Object, ByVal str�洢�ⷿ As String, ByVal str�洢�ⷿID As String, ByVal str�ⷿ���� As String, ByVal str�ⷿ����ID As String)
    mstr�洢�ⷿ = str�洢�ⷿ
    mstr�洢�ⷿID = str�洢�ⷿID
    mstr�ⷿ���� = str�ⷿ����
    mstr�ⷿ����ID = str�ⷿ����ID

    Me.Show 1, frmParent
End Sub

Private Sub cmdȡ��_Click()
    Unload Me
End Sub

Private Sub cmdȷ��_Click()
    Dim i As Integer
    
    mstr�ⷿ���� = ""
    mstr�ⷿ����ID = ""
    
    For i = 1 To vsfDepartment.Rows - 1
        If vsfDepartment.TextMatrix(i, 2) = "" Then
            mstr�ⷿ����ID = mstr�ⷿ����ID & "!!" & vsfDepartment.TextMatrix(i, 1) & "|"
        End If
        
        If vsfDepartment.TextMatrix(i, 2) <> "" Then
            mstr�ⷿ���� = mstr�ⷿ���� & "��" & vsfDepartment.TextMatrix(i, 0) & "��" & vsfDepartment.TextMatrix(i, 2)
            mstr�ⷿ����ID = mstr�ⷿ����ID & "!!" & vsfDepartment.TextMatrix(i, 1) & "|" & vsfDepartment.TextMatrix(i, 3)
        End If
    Next
     
     mstr�ⷿ���� = Mid(mstr�ⷿ����, 2)
     mstr�ⷿ����ID = Mid(mstr�ⷿ����ID, 3)
     
    Call frmBatchUpdate.ShowDepartment(mstr�ⷿ����, mstr�ⷿ����ID, 0)
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    chk����.Value = 0
    picDrug.Visible = False
    vsfDepartment.Enabled = True
End Sub

Private Sub Form_Load()
    Call Init��ʼ�����
    Call Init��ʼ���ⷿ����
End Sub

Private Sub Init��ʼ�����()
    
    VsfGridColFormat vsfDepartment, mSpecColumn.�洢�ⷿ, "�洢�ⷿ", 1500, flexAlignLeftCenter, "�洢�ⷿ"
    VsfGridColFormat vsfDepartment, mSpecColumn.�洢�ⷿID, "�洢�ⷿID", 1500, flexAlignCenterCenter, "�洢�ⷿID"
    VsfGridColFormat vsfDepartment, mSpecColumn.�������, "�������", 4000, flexAlignLeftCenter, "�������"
    VsfGridColFormat vsfDepartment, mSpecColumn.�������id, "�������id", 4000, flexAlignCenterCenter, "�������id"
    vsfDepartment.ColComboList(mSpecColumn.�������) = "..."
    
End Sub

Private Sub Init��ʼ���ⷿ����()
    '����һ��������������ѡ���˾���Ĵ洢�ⷿ�󣬳�ʼ��vsfDepartment��Ŀⷿ�ͿⷿID(���ڲ�ѯ�������)
    '���ܶ��������ѡ���˷������֮���ٴε�����ͻ�Ѷ�Ӧ�ķ������Ҳ��ʾ����
    Dim i As Integer, j As Integer
    Dim rsRoom As New ADODB.Recordset
    
    mstrArr�洢�ⷿ = Split(mstr�洢�ⷿ, "|")
    mstrArr�洢�ⷿID = Split(mstr�洢�ⷿID, "!!")
    mstrArr�ⷿ����ID = Split(mstr�ⷿ����ID, "!!")
    
    For i = LBound(mstrArr�洢�ⷿ) To UBound(mstrArr�洢�ⷿ)
        vsfDepartment.Rows = vsfDepartment.Rows + 1
        vsfDepartment.RowHeight(i + 1) = 400
        vsfDepartment.TextMatrix(i + 1, mSpecColumn.�洢�ⷿ) = mstrArr�洢�ⷿ(i)
    Next
     
    For i = LBound(mstrArr�洢�ⷿID) To UBound(mstrArr�洢�ⷿID)
        vsfDepartment.TextMatrix(i + 1, mSpecColumn.�洢�ⷿID) = Split(mstrArr�洢�ⷿID(i), "|")(0)
    Next
    
    For i = LBound(mstrArr�ⷿ����ID) To UBound(mstrArr�ⷿ����ID)
        For j = 1 To vsfDepartment.Rows - 1
            If Split(mstrArr�ⷿ����ID(i), "|")(0) = vsfDepartment.TextMatrix(j, 1) Then
                    vsfDepartment.TextMatrix(j, 3) = Split(mstrArr�ⷿ����ID(i), "|")(1)
                    
                    gstrSql = "select a.���� from ���ű� a where a.id in(Select Column_Value From Table(f_num2list([1])))"
                    Set rsRoom = zlDatabase.OpenSQLRecord(gstrSql, "", vsfDepartment.TextMatrix(j, 3))
                    
                    Do While Not rsRoom.EOF
                        vsfDepartment.TextMatrix(j, 2) = vsfDepartment.TextMatrix(j, 2) & "," & rsRoom!����
                        rsRoom.MoveNext
                    Loop
                    If vsfDepartment.TextMatrix(j, 2) <> "" Then
                        vsfDepartment.TextMatrix(j, 2) = Mid(vsfDepartment.TextMatrix(j, 2), 2)
                    End If
                    
                Exit For
            End If
        Next
    Next
End Sub
Public Sub VsfGridColFormat(ByVal objGrid As VSFlexGrid, ByVal intCol As Integer, ByVal strColName As String, _
    ByVal lngColWidth As Long, ByVal intColAlignment As Integer, _
    Optional ByVal strColKey As String = "", Optional ByVal intFixedColAlignment As Integer = 4)
    'vsf�����ã��������п��ж��뷽ʽ���̶��ж��뷽ʽ��Ĭ��Ϊ���ж��룩

    With objGrid
        .TextMatrix(0, intCol) = strColName
        .ColWidth(intCol) = lngColWidth
        .ColAlignment(intCol) = intColAlignment
        .ColKey(intCol) = strColKey
        .FixedAlignment(intCol) = intFixedColAlignment
    End With
End Sub


Private Sub vsfDepartment_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Select Case Col
        Case mSpecColumn.�������
            If Check������� = False Then
'                Call Init���ؿ�������
            Call frmServiceSelect.ShowMe(frmServiceDepartment, vsfDepartment.Row, mstr�������, 2)
            End If
    End Select
End Sub

Private Function Check�������() As Boolean
    '���ܣ���鵱ǰ�ⷿ�ǲ���ҩ�������Ƿ������ٴ�����
    '����ֵ true ��ǰ�ⷿ����ҩ��Ҳû�������ٴ�����,false ��ǰ�ⷿ��ҩ�����߻����������ٴ�����
    Dim str������� As String
    Dim rsTemp As ADODB.Recordset

    On Error GoTo ErrHandle
    str������� = ""
    gstrSql = "select distinct ������� from ��������˵�� where ����ID=[1] "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "��ȡ�������", vsfDepartment.TextMatrix(vsfDepartment.Row, 1))

    Do While Not rsTemp.EOF
        str������� = str������� & "," & rsTemp!�������
        rsTemp.MoveNext
    Loop
    If str������� <> "" Then
        str������� = Mid(str�������, 2)
        If InStr(1, str�������, 3) <> 0 Then
            str������� = "0,1,2,3"
        ElseIf InStr(1, str�������, 1) <> 0 Or InStr(1, str�������, 2) <> 0 Then
            str������� = str������� & ",3"
        End If
    Else
        str������� = "0"
    End If
    mstr������� = str�������
    
    gstrSql = "Select distinct a.Id, a.����, a.����, a.����" & vbNewLine & _
            "From ���ű� a, ��������˵�� b, �������ʷ��� c" & vbNewLine & _
            "Where a.Id = b.����id And b.�������� = c.���� And Instr('3ABCDEF', c.����) > 0 And" & vbNewLine & _
            "  (a.����ʱ�� Is Null Or a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD')) And Instr([1], ',' || b.������� || ',') > 0 order by id"

    Set mrs���� = zlDatabase.OpenSQLRecord(gstrSql, "��ȡ�������", "," & str������� & ",")

    If mrs����.RecordCount = 0 Then
        MsgBox "��ǰ�ⷿ����ҩ������δ�����ٴ����ң�[���Ź���]", vbInformation, gstrSysName
        vsfDepartment.Text = ""
        vsfDepartment.TextMatrix(vsfDepartment.Row, vsfDepartment.Col) = ""
        Check������� = True
        Exit Function
    End If
    Check������� = False
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Init���ؿ�������()
    '��lvwItems��ʾ������ķ������
    Dim str������� As String
    Dim objItem As ListItem
    Dim intItem As Integer
    Dim i As Integer, j As Integer
    
    Call AddColumnHeader(False)
    Me.lvwItems.ListItems.Clear
    Me.lvwItems.Checkboxes = True
    
    Do While Not mrs����.EOF
        Set objItem = Me.lvwItems.ListItems.Add(, "_" & mrs����!ID, mrs����!����, , 3)
        objItem.SubItems(Me.lvwItems.ColumnHeaders("����").Index - 1) = mrs����!����
        objItem.SubItems(Me.lvwItems.ColumnHeaders("����").Index - 1) = mrs����!����
        mrs����.MoveNext
    Loop
    
    With Me.picDrug
        .Left = Me.vsfDepartment.Left + 1500
        .Top = Me.vsfDepartment.Top + Me.vsfDepartment.CellTop
        If .Top + .Height > Me.ScaleHeight Then
            .Top = Me.ScaleHeight - .Height
        End If

        lvwItems.Move 0, 250, picDrug.Width, picDrug.Height - 670
        chk����.Move 3300, 0
        cmdOk.Move 0, picDrug.Height - 400
        cmdCancel.Move cmdOk.Width, cmdOk.Top
        
        lvwItems.Visible = True
        .ZOrder 0: .Visible = True
        vsfDepartment.Enabled = False
        .SetFocus
    End With
    
    '��ѡ���˷�����Һ��ٴε��������ң�������еķ��������lvwItems��ʾ����
    mstr������� = vsfDepartment.TextMatrix(vsfDepartment.Row, mSpecColumn.�������)
    mstrArr������� = Split(mstr�������, ",")
    
    For i = LBound(mstrArr�������) To UBound(mstrArr�������)
        For intItem = 1 To lvwItems.ListItems.Count
            If mstrArr�������(i) = lvwItems.ListItems(intItem).Text Then
                lvwItems.ListItems(intItem).Checked = True
                j = j + 1
            End If
        Next
    Next
    
    If j = lvwItems.ListItems.Count Then
        chk����.Value = 1
    ElseIf j > 0 And j < lvwItems.ListItems.Count Then
        chk����.Value = 2
    End If
End Sub

Private Sub AddColumnHeader(Optional ByVal blnҩƷ As Boolean = True)
 
        With Me.lvwItems.ColumnHeaders
            .Clear
            .Add , "����", "����", 2000
            .Add , "����", "����", 800
            .Add , "����", "����", 700
        End With
        
        With Me.lvwItems
            .Checkboxes = True
            .ColumnHeaders("����").Position = 1
            .Sorted = False '�ر�������
        End With
    
End Sub

Private Sub cmdOK_Click()
    Dim intItem As Integer, intItems As Integer
    
    '��ѡ��ķ�����Һͷ������ID��ʾ��vsfDepartment
    mstr�������ID = ""
    mstr������� = ""
    intItems = Me.lvwItems.ListItems.Count
    For intItem = 1 To intItems
        If lvwItems.ListItems(intItem).Checked Then
            mstr�������ID = mstr�������ID & "," & Mid(lvwItems.ListItems(intItem).Key, 2)
            mstr������� = mstr������� & "," & Mid(lvwItems.ListItems(intItem).Text, 1)
        End If
    Next
    
    mstr������� = Mid(mstr�������, 2)
    mstr�������ID = Mid(mstr�������ID, 2)
    If vsfDepartment.Row <> 0 Then
        vsfDepartment.TextMatrix(vsfDepartment.Row, mSpecColumn.�������) = mstr�������
        vsfDepartment.TextMatrix(vsfDepartment.Row, mSpecColumn.�������id) = mstr�������ID
    End If
    
    picDrug.Visible = False
    vsfDepartment.Enabled = True
End Sub

Private Sub vsfDepartment_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Col = mSpecColumn.������� Then
        KeyAscii = 0
    End If
End Sub
Private Sub vsfDepartment_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = mSpecColumn.�洢�ⷿ Then
        Cancel = True
    End If
End Sub

Private Sub chk����_Click()
'�ⷿȫѡ��ť
    If chk����.Value = 2 Then Exit Sub
    Call SetSelect(lvwItems, chk����.Value)
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
Private Sub lvwItems_ItemCheck(ByVal Item As MSComctlLib.ListItem)
'����ѡ��Ĵ洢�ⷿ
    Call ItemCheck(lvwItems, Item, chk����)
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
