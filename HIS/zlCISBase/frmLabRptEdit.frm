VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLabRptEdit 
   BorderStyle     =   0  'None
   Caption         =   "����ģ��༭"
   ClientHeight    =   5580
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6945
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5580
   ScaleWidth      =   6945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picEdit 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   900
      Left            =   0
      ScaleHeight     =   900
      ScaleWidth      =   6945
      TabIndex        =   13
      Top             =   4665
      Width           =   6945
      Begin VB.TextBox txt���� 
         Height          =   300
         Left            =   960
         MaxLength       =   60
         TabIndex        =   15
         Top             =   105
         Width           =   5835
      End
      Begin VB.TextBox txt��ע 
         Height          =   300
         Left            =   960
         MaxLength       =   60
         TabIndex        =   17
         Top             =   495
         Width           =   5835
      End
      Begin VB.Label lbl���� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   150
         TabIndex        =   14
         Top             =   165
         Width           =   720
      End
      Begin VB.Label lbl��ע 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "���汸ע"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   150
         TabIndex        =   16
         Top             =   555
         Width           =   720
      End
   End
   Begin VB.PictureBox picName 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   1215
      Left            =   15
      ScaleHeight     =   1215
      ScaleWidth      =   6780
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   15
      Width           =   6780
      Begin VB.TextBox txtItem 
         Height          =   300
         Left            =   1620
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   915
         Width           =   4050
      End
      Begin VB.CommandButton cmdItem 
         Caption         =   "ѡ��"
         Height          =   315
         Left            =   5700
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   915
         Width           =   1100
      End
      Begin VB.TextBox txt˵�� 
         Height          =   300
         Left            =   2520
         MaxLength       =   60
         TabIndex        =   8
         Top             =   495
         Width           =   4245
      End
      Begin VB.TextBox txt���� 
         Height          =   300
         Left            =   2520
         MaxLength       =   60
         TabIndex        =   4
         Top             =   105
         Width           =   4245
      End
      Begin VB.TextBox txt���� 
         Height          =   300
         Left            =   585
         MaxLength       =   13
         TabIndex        =   2
         Top             =   105
         Width           =   1260
      End
      Begin VB.TextBox txt���� 
         Height          =   300
         Left            =   585
         MaxLength       =   10
         TabIndex        =   6
         Top             =   495
         Width           =   1260
      End
      Begin VB.Label lblItem 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��Ӧ���������Ŀ"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   135
         TabIndex        =   9
         Top             =   975
         Width           =   1440
      End
      Begin VB.Label lbl˵�� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "˵��"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   2070
         TabIndex        =   7
         Top             =   555
         Width           =   360
      End
      Begin VB.Label lbl���� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   2070
         TabIndex        =   3
         Top             =   165
         Width           =   360
      End
      Begin VB.Label lbl���� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   135
         TabIndex        =   1
         Top             =   165
         Width           =   360
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   135
         TabIndex        =   5
         Top             =   555
         Width           =   360
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgEdit 
      Height          =   3210
      Left            =   150
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1305
      Width           =   6645
      _cx             =   11721
      _cy             =   5662
      Appearance      =   2
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
      BackColorFixed  =   15790320
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16772055
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
      ScrollTrack     =   -1  'True
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
      Editable        =   0
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
   Begin MSComctlLib.ListView lvwItems 
      Height          =   4065
      Left            =   1635
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1245
      Visible         =   0   'False
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   7170
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
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   0
      Top             =   900
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
            Picture         =   "frmLabRptEdit.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmLabRptEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngItemID As Long          '��ǰ��ʾ����Ŀid
Private mbln΢���� As Boolean

Private Enum mCol
    ID = 0:  ������: Ӣ����: ��λ: ������: ��������
End Enum

Dim objItem As ListItem
Dim lngCount As Long

'--------------------------------------------
'����Ϊ���幫������
'--------------------------------------------
Private Sub RecallReport()
    '���ܣ�����װ�ر�����Ŀ
    Dim rsTemp As New ADODB.Recordset
    
    
    Err = 0: On Error GoTo ErrHand
    
    gstrSql = "Select B.��Ŀ���" & vbNewLine & _
                "From ���鱨����Ŀ A, ������Ŀ B" & vbNewLine & _
                "Where A.������Ŀid = B.������Ŀid And B.��Ŀ��� = 2 And ������Ŀid = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(Me.txtItem.Tag))
    mbln΢���� = Not rsTemp.EOF
    
    If mbln΢���� = False Then
        gstrSql = "Select I.ID, I.������, I.Ӣ����, I.��λ, C.������" & vbNewLine & _
                "From ���鱨����Ŀ R, ����������Ŀ I, (Select ��Ŀid, ������ From ����ģ������ Where ģ��ID = [1]) C" & vbNewLine & _
                "Where R.������ĿID = I.ID And I.ID = C.��Ŀid(+) And R.������Ŀid = [2]" & vbNewLine & _
                "Order By r.�������"
    Else
        gstrSql = "Select I.ID, I.������, I.Ӣ����, '' As ��λ, C.������,C.�������� " & vbNewLine & _
            "From ���鱨����Ŀ R, ����ϸ�� I, (Select ϸ��id, ������,�������� From ����ģ������ Where ģ��id = [1]) C" & vbNewLine & _
            "Where R.ϸ��id = I.ID And R.ϸ��id is not null And I.ID = C.ϸ��id(+) And R.������Ŀid = [2]" & vbNewLine & _
            "Order By R.�������"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngItemID, Val(Me.txtItem.Tag))
    Me.vfgEdit.Clear
    Set Me.vfgEdit.DataSource = rsTemp: Call setListFormat(True)
    If Me.vfgEdit.Rows > Me.vfgEdit.FixedRows Then Me.vfgEdit.Row = Me.vfgEdit.FixedRows
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub setListFormat(Optional blnKeepData As Boolean)
    '���ܣ���ʼ�������б�
    '������ blnKeepData-�Ƿ������ݣ���ֻ���������ø�ʽ
    With Me.vfgEdit
        .Redraw = flexRDNone
        If blnKeepData = False Then
            .Clear
            .Rows = 1: .FixedRows = 1: .Cols = 6: .FixedCols = 0
        End If
        .TextMatrix(0, mCol.ID) = "ID": .TextMatrix(0, mCol.������) = "������": .TextMatrix(0, mCol.Ӣ����) = "Ӣ����"
        .TextMatrix(0, mCol.��λ) = "��λ": .TextMatrix(0, mCol.������) = "������"
'        Call IIf(mbln΢���� = True, .TextMatrix(0, mCol.��������) = "��������", "")
        
        .ColWidth(mCol.ID) = 0: .ColWidth(mCol.������) = 3000: .ColWidth(mCol.Ӣ����) = 1000
        .ColWidth(mCol.��λ) = 700: .ColWidth(mCol.������) = 900
'        Call IIf(mbln΢���� = False, .ColWidth(mCol.��������) = 0, .ColWidth(mCol.��������) = 500)
        For lngCount = 0 To .Cols - 1
            .FixedAlignment(lngCount) = flexAlignCenterCenter
            If .ColWidth(lngCount) = 0 Then .ColHidden(lngCount) = True
        Next
        .Redraw = flexRDDirect
    End With
End Sub

Public Function zlRefresh(lngItemId As Long) As Boolean
    '���ܣ�������Ŀidˢ�µ�ǰ��ʾ����
    Dim rsTemp As New ADODB.Recordset
    
    mlngItemID = lngItemId
    
    '�����ǰ��Ŀ����ʾ
    Me.txt����.Text = "": Me.txt����.Text = "": Me.txt����.Text = "": Me.txt˵��.Text = ""
    Me.txt����.Text = "": Me.txt��ע.Text = ""
    
    '��ȡָ����Ŀ����Ϣ
    Err = 0: On Error GoTo ErrHand
    
    gstrSql = "Select ����, ����, ����, ˵��, ������Ŀid, ��������, ���鱸ע From ����ģ��Ŀ¼ L Where ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngItemId)
    With rsTemp
        Me.txt����.MaxLength = .Fields("����").DefinedSize
        Me.txt����.MaxLength = .Fields("����").DefinedSize
        Me.txt����.MaxLength = .Fields("����").DefinedSize
        Me.txt˵��.MaxLength = .Fields("˵��").DefinedSize
        Me.txt����.MaxLength = .Fields("��������").DefinedSize
        Me.txt��ע.MaxLength = .Fields("���鱸ע").DefinedSize
        If .RecordCount > 0 Then
            Me.txt����.Text = "" & !����
            Me.txt����.Text = "" & !����: Me.txt����.Text = "" & !����: Me.txt˵��.Text = "" & !˵��
            Me.txt����.Text = "" & !��������: Me.txt��ע.Text = "" & !���鱸ע
            For Each objItem In Me.lvwItems.ListItems
                If Mid(objItem.Key, 2) = Val("" & !������Ŀid) Then
                    objItem.Selected = True
                    Me.txtItem.Tag = Mid(objItem.Key, 2)
                    Me.txtItem.Text = objItem.Text
                    
                End If
            Next
        Else
            Me.txtItem.Tag = ""
            Me.txtItem.Text = ""
        End If
    End With
    Call RecallReport
    
    zlRefresh = True: Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlRefresh = False: Exit Function
End Function

Public Function zlEditStart(blnAdd As Boolean, lngItemId As Long) As Boolean
    '���ܣ���ʼ��Ŀ�༭
    '������ blnAdd-�Ƿ����ӣ�����Ϊ�޸�
    '       lngItemId-���ӵĲ�����Ŀ������ָ���༭����Ŀ
    Dim rsTemp As New ADODB.Recordset
    
    If blnAdd Then
        Err = 0: On Error GoTo ErrHand
        gstrSql = "Select Nvl(Max(To_Number(����)), 0) As ����, Nvl(Max(Length(����)), 0) As ���� From ����ģ��Ŀ¼"
'            If .State = adStateOpen Then .Close
'            Call SQLTest(App.ProductName, Me.Caption, gstrSql)
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "zlEditStart")
'            Call SQLTest
        With rsTemp
            If !���� <> 0 And !���� <= Me.txt����.MaxLength Then
                Me.txt����.Text = Format(Val(!����) + 1, String(!����, "0"))
            Else
                Me.txt����.Text = Format(Val(!����) + 1, String(Me.txt����.MaxLength, "0"))
            End If
            
            Me.txt����.Text = "": Me.txt����.Text = "": Me.txt˵��.Text = ""
            Me.txt����.Text = "": Me.txt��ע.Text = ""
            Me.txtItem.Tag = "": Me.txtItem.Text = "": Call setListFormat
        End With
    End If

    Me.Tag = IIf(blnAdd, "����", "�޸�")
    Me.BackColor = RGB(250, 250, 250): Me.picName.BackColor = Me.BackColor: Me.picEdit.BackColor = Me.BackColor
    Me.picName.Enabled = True: Me.picEdit.Enabled = True
    Me.vfgEdit.Editable = flexEDKbd: Me.vfgEdit.FocusRect = flexFocusHeavy
    
    Me.txt����.SetFocus
    zlEditStart = True: Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlEditStart = False: Exit Function
End Function

Public Sub zlEditCancel()
    '���ܣ��������ڽ��еı༭
    Me.Tag = ""
    Me.BackColor = &H8000000F: Me.picName.BackColor = Me.BackColor: Me.picEdit.BackColor = Me.BackColor
    Me.picName.Enabled = False: Me.picEdit.Enabled = False
    Me.vfgEdit.Editable = flexEDNone: Me.vfgEdit.FocusRect = flexFocusNone
    Call Me.zlRefresh(mlngItemID)
End Sub

Public Function zlEditSave() As Long
    '���ܣ��������ڽ��еı༭,���������ڱ༭��Ŀid,����ʧ�ܷ���0
    Dim lngNewId As Long, strLists As String
    Dim str����  As String
    'һ�����Լ��
    If Trim(Me.txt����.Text) = "" Then
        MsgBox "��������룡", vbInformation, gstrSysName
        Me.txt����.SetFocus: zlEditSave = 0: Exit Function
    End If
    If Val(Me.txt����.Text) > Val(String(Me.txt����.MaxLength, "9")) Then
        MsgBox "����̫��", vbInformation, gstrSysName
        Me.txt����.SetFocus: zlEditSave = 0: Exit Function
    End If
    If Trim(Me.txt����.Text) = "" Then
        MsgBox "���������ƣ�", vbInformation, gstrSysName
        Me.txt����.SetFocus: zlEditSave = 0: Exit Function
    End If
    If LenB(StrConv(Trim(Me.txt����.Text), vbFromUnicode)) > Me.txt����.MaxLength Then
        MsgBox "���Ƴ��������" & Me.txt����.MaxLength & "���ַ���ȳ����֣���", vbInformation, gstrSysName
        Me.txt����.SetFocus: zlEditSave = 0: Exit Function
    End If
    If LenB(StrConv(Trim(Me.txt����.Text), vbFromUnicode)) > Me.txt����.MaxLength Then
        MsgBox "��д���������" & Me.txt����.MaxLength & "���ַ�����", vbInformation, gstrSysName
        Me.txt����.SetFocus: zlEditSave = 0: Exit Function
    End If
    If LenB(StrConv(Trim(Me.txt˵��.Text), vbFromUnicode)) > Me.txt˵��.MaxLength Then
        MsgBox "˵�����������" & Me.txt˵��.MaxLength & "���ַ�����", vbInformation, gstrSysName
        Me.txt˵��.SetFocus: zlEditSave = 0: Exit Function
    End If
    If LenB(StrConv(Trim(Me.txt����.Text), vbFromUnicode)) > Me.txt����.MaxLength Then
        MsgBox "���ﳬ�������" & Me.txt����.MaxLength & "���ַ�����", vbInformation, gstrSysName
        Me.txt����.SetFocus: zlEditSave = 0: Exit Function
    End If
    If LenB(StrConv(Trim(Me.txt��ע.Text), vbFromUnicode)) > Me.txt��ע.MaxLength Then
        MsgBox "��ע���������" & Me.txt��ע.MaxLength & "���ַ�����", vbInformation, gstrSysName
        Me.txt��ע.SetFocus: zlEditSave = 0: Exit Function
    End If
    
    strLists = ""
    With Me.vfgEdit
        For lngCount = .FixedRows To .Rows - 1
            If LenB(StrConv(Trim(.TextMatrix(lngCount, mCol.������)), vbFromUnicode)) > 50 Then
                MsgBox "��" & lngCount & "�н����д����", vbInformation, gstrSysName
                .SetFocus: zlEditSave = 0: Exit Function
            End If
            If mbln΢���� Then
                str���� = .TextMatrix(lngCount, mCol.��������)
            Else
                str���� = ""
            End If
            strLists = strLists & "|" & .TextMatrix(lngCount, mCol.ID) & ";" & .TextMatrix(lngCount, mCol.������) & ";" & str����
        Next
    End With
    If strLists = "" Then
        MsgBox "û������ģ��ı������ݣ�", vbInformation, gstrSysName
        Me.txt˵��.SetFocus: zlEditSave = 0: Exit Function
    End If
    strLists = Mid(strLists, 2)
    
    gstrSql = "'" & Trim(Me.txt����.Text) & "','" & Trim(Me.txt����.Text) & "','" & Trim(Me.txt����.Text) & "'"
    gstrSql = gstrSql & ",'" & Trim(Me.txt˵��.Text) & "'," & Val(Me.txtItem.Tag)
    gstrSql = gstrSql & ",'" & Trim(Me.txt����.Text) & "','" & Trim(Me.txt��ע.Text) & "'"
    gstrSql = gstrSql & ",'" & strLists & "'"
    
    '���ݱ��������֯
    
    lngNewId = mlngItemID
    If Me.Tag = "����" Then
        lngNewId = zlDatabase.GetNextId("����ģ��Ŀ¼")
        gstrSql = "Zl_����ģ��Ŀ¼_Edit(1," & lngNewId & "," & gstrSql & ")"
    Else
        gstrSql = "Zl_����ģ��Ŀ¼_Edit(2," & lngNewId & "," & gstrSql & ")"
    End If
    
    Err = 0: On Error GoTo ErrHand
    Call SQLTest(App.ProductName, Me.Caption, gstrSql): gcnOracle.Execute gstrSql, , adCmdStoredProc: Call SQLTest
    
    If Me.Tag = "����" Then mlngItemID = lngNewId
    
    Me.Tag = ""
    Me.BackColor = &H8000000F: Me.picName.BackColor = Me.BackColor: Me.picEdit.BackColor = Me.BackColor
    Me.picName.Enabled = False: Me.picEdit.Enabled = False
    Me.vfgEdit.Editable = flexEDNone: Me.vfgEdit.FocusRect = flexFocusNone
    
    zlEditSave = mlngItemID: Exit Function
    
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlEditSave = 0: Exit Function
End Function

'--------------------------------------------
'����Ϊ����ؼ���Ӧ�¼�
'--------------------------------------------

Private Sub cmdItem_Click()
    Dim rsTemp As New ADODB.Recordset
    With Me.lvwItems
        .Tag = Me.txtItem.Name
        .Left = Me.txtItem.Left
        .Top = Me.txtItem.Top + Me.txtItem.Height
        .ZOrder 0: .Visible = True
        .SetFocus
    End With
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyEscape Then Exit Sub
    If Me.lvwItems.Visible Then
        Me.lvwItems.Visible = False: Me.txtItem.SetFocus: Exit Sub
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim rsTemp As New ADODB.Recordset
    
    mlngItemID = 0
   
    Me.picName.BackColor = Me.BackColor
    Me.picEdit.BackColor = Me.BackColor
    Call setListFormat
    Me.vfgEdit.ZOrder 0

    '------------------------------------------
    With Me.lvwItems.ColumnHeaders
        .Clear
        .Add , "����", "����", 3500
        .Add , "����", "����", 1000
        .Add , "����", "����", 1000
    End With
    With Me.lvwItems
        .ColumnHeaders("����").Position = 1
        .SortKey = .ColumnHeaders("����").Index - 1
        .SortOrder = lvwAscending
    End With
    
    Err = 0: On Error GoTo ErrHand
'        gstrSql = "Select I.ID, I.�������� As ����, I.����, I.����" & vbNewLine & _
            "From ������ĿĿ¼ I, ���Ƽ������� K" & vbNewLine & _
            "Where I.��� = 'C' And I.�������� = K.���� And I.�����Ŀ = 1 And" & vbNewLine & _
            "      (I.����ʱ�� Is Null Or To_Char(I.����ʱ��, 'yyyy-mm-dd') = '3000-01-01')"
    gstrSql = "Select Distinct I.ID, I.�������� As ����, I.����, I.����,decode(N.��Ŀ���,2,2,1) as ��Ŀ��� " & vbNewLine & _
            "From ������ĿĿ¼ I, ���Ƽ������� K, ���鱨����Ŀ M, ������Ŀ N, ����ϸ�� O" & vbNewLine & _
            "Where I.��� = 'C' And I.�������� = K.���� And I.ID = M.������Ŀid And (M.������Ŀid = N.������Ŀid Or M.ϸ��id = O.ID) And" & vbNewLine & _
            "      ((N.��Ŀ��� = 1 And I.�����Ŀ = 1) Or (N.��Ŀ��� = 2 And I.����Ӧ�� = 1)) And" & vbNewLine & _
            "      (I.����ʱ�� Is Null Or To_Char(I.����ʱ��, 'yyyy-mm-dd') = '3000-01-01')"
            
'        If .State = adStateOpen Then .Close
'        Call SQLTest(App.Title, Me.Caption, gstrSql)
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "Form_Load")
'        Call SQLTest
    Me.lvwItems.ListItems.Clear
    With rsTemp
        Do While Not .EOF
            Set objItem = Me.lvwItems.ListItems.Add(, IIf(Val(!��Ŀ���) = 2, "^", "_") & !ID, !����)
            objItem.Icon = 1: objItem.SmallIcon = 1
            objItem.SubItems(Me.lvwItems.ColumnHeaders("����").Index - 1) = !����
            objItem.SubItems(Me.lvwItems.ColumnHeaders("����").Index - 1) = !����
            .MoveNext
        Loop
    End With
    
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog

End Sub

Private Sub Form_Resize()
    Err = 0: 'On Error Resume Next
    Me.picEdit.Top = Me.ScaleHeight - Me.picEdit.Height
    Me.vfgEdit.Height = Me.picEdit.Top - Me.vfgEdit.Top
End Sub

Private Sub lvwItems_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If Me.lvwItems.SortKey = ColumnHeader.Index - 1 Then
        Me.lvwItems.SortOrder = IIf(Me.lvwItems.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        Me.lvwItems.SortKey = ColumnHeader.Index - 1
        Me.lvwItems.SortOrder = lvwAscending
    End If
End Sub

Private Sub lvwItems_DblClick()
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    With Me.lvwItems
        Me.txtItem.Tag = Mid(.SelectedItem.Key, 2)
        Me.txtItem.Text = .SelectedItem.Text
        Me.lvwItems.Visible = False
        Call RecallReport
        Call zlCommFun.PressKey(vbKeyTab)
    End With
End Sub

Private Sub lvwItems_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn, vbKeySpace
        If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
        Call lvwItems_DblClick
    End Select
End Sub

Private Sub lvwItems_LostFocus()
    Me.lvwItems.Visible = False
End Sub

Private Sub txtItem_GotFocus()
    Me.txtItem.SelStart = 0: Me.txtItem.SelLength = 1000
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txtItem_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub txt��ע_GotFocus()
    Me.txt����.SelStart = 0: Me.txt����.SelLength = 1000
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt��ע_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(GCST_INVALIDCHAR, Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt����_GotFocus()
    Me.txt����.SelStart = 0: Me.txt����.SelLength = 1000
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt����_GotFocus()
    Me.txt����.SelStart = 0: Me.txt����.SelLength = 1000
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt����_GotFocus()
    Me.txt����.SelStart = 0: Me.txt����.SelLength = 1000
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Me.txt����.Text = MoveSpecialChar(Me.txt����.Text)
        Me.txt����.Text = zlStr.GetCodeByORCL(Me.txt����.Text, False, Me.txt����.MaxLength)
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    End If
    If InStr(GCST_INVALIDCHAR, Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt����_GotFocus()
    Me.txt����.SelStart = 0: Me.txt����.SelLength = 1000
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(GCST_INVALIDCHAR, Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt˵��_GotFocus()
    Me.txt˵��.SelStart = 0: Me.txt˵��.SelLength = 1000
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt˵��_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(GCST_INVALIDCHAR, Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub vfgEdit_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> mCol.������ And Col <> mCol.�������� Then Cancel = True
End Sub

