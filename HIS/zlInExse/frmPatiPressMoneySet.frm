VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPatiPressMoneySet 
   Caption         =   "������������"
   ClientHeight    =   7875
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11010
   Icon            =   "frmPatiPressMoneySet.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7875
   ScaleWidth      =   11010
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picվ�� 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   390
      Left            =   45
      ScaleHeight     =   390
      ScaleWidth      =   5625
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1260
      Width           =   5625
      Begin VB.ComboBox cboվ�� 
         Height          =   300
         Left            =   480
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   60
         Width           =   3195
      End
      Begin VB.Label lblվ�� 
         AutoSize        =   -1  'True
         Caption         =   "վ��"
         Height          =   180
         Left            =   90
         TabIndex        =   17
         Top             =   120
         Width           =   360
      End
   End
   Begin VB.PictureBox picList��� 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3555
      Left            =   4455
      ScaleHeight     =   3525
      ScaleWidth      =   2310
      TabIndex        =   8
      Top             =   2670
      Visible         =   0   'False
      Width           =   2340
      Begin VB.ListBox lst��� 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   3180
         Left            =   -30
         Style           =   1  'Checkbox
         TabIndex        =   9
         Top             =   360
         Width           =   2355
      End
      Begin XtremeSuiteControls.ShortcutCaption shtCaption 
         Height          =   360
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   2325
         _Version        =   589884
         _ExtentX        =   4101
         _ExtentY        =   635
         _StockProps     =   6
         Caption         =   "���ѡ��"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColorLight=   16744576
         GradientColorDark=   16761024
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsPressMoney 
      Height          =   4110
      Left            =   195
      TabIndex        =   2
      Top             =   2145
      Width           =   10515
      _cx             =   18547
      _cy             =   7250
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
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
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   10
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmPatiPressMoneySet.frx":6852
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
   Begin VB.PictureBox picTop 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1245
      Left            =   0
      ScaleHeight     =   1245
      ScaleWidth      =   11010
      TabIndex        =   4
      Top             =   0
      Width           =   11010
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmPatiPressMoneySet.frx":698B
         Height          =   555
         Left            =   600
         TabIndex        =   7
         Top             =   615
         Width           =   7740
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000010&
         X1              =   -45
         X2              =   11640
         Y1              =   1215
         Y2              =   1215
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����������ÿ�ַ������������������߼�������ʽ����� zl_PatiWarnScheme �������ʹ��"
         Height          =   180
         Left            =   600
         TabIndex        =   6
         Top             =   390
         Width           =   7290
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   210
         TabIndex        =   5
         Top             =   135
         Width           =   1170
      End
   End
   Begin MSComctlLib.TabStrip tab���� 
      Height          =   4650
      Left            =   105
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1725
      Width           =   10755
      _ExtentX        =   18971
      _ExtentY        =   8202
      HotTracking     =   -1  'True
      TabMinWidth     =   0
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "��ͨ����"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picDown 
      BorderStyle     =   0  'None
      Height          =   1425
      Left            =   45
      ScaleHeight     =   1425
      ScaleWidth      =   11025
      TabIndex        =   11
      Top             =   6525
      Width           =   11025
      Begin VB.CommandButton cmdWarnNew 
         Caption         =   "���ӱ�������(&A)"
         Height          =   350
         Left            =   45
         TabIndex        =   15
         Top             =   60
         Width           =   1710
      End
      Begin VB.CommandButton cmdWarnDel 
         Caption         =   "ɾ����������(&D)"
         Height          =   350
         Left            =   1860
         TabIndex        =   14
         Top             =   60
         Width           =   1710
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "���淽��(&O)"
         Height          =   350
         Left            =   9465
         TabIndex        =   3
         Top             =   60
         Width           =   1395
      End
      Begin VB.Frame fraSplit 
         Height          =   90
         Left            =   -60
         TabIndex        =   13
         Top             =   465
         Width           =   11025
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   9570
         TabIndex        =   12
         Top             =   735
         Width           =   1150
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         Caption         =   "ע��:����Ĳ�����վ��."
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   195
         TabIndex        =   18
         Top             =   735
         Visible         =   0   'False
         Width           =   2820
      End
   End
End
Attribute VB_Name = "frmPatiPressMoneySet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrPrivs As String, mlngModule As Long
Private mblnEdit As Boolean, mblnSort As Boolean, mblnChange As Boolean
Private mrsWarn As ADODB.Recordset
Private mrs��� As ADODB.Recordset
Private mstrDel���ò��� As String
Private mblnOK As Boolean
Private mblnNotClick As Boolean
Private mlngPreSelIdx As Long   '�ϴ�����
Private mblnFirst As Boolean

Private Sub LoadClients()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����վ��
    '����:���˺�
    '����:2011-02-11 10:10:37
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strվ�� As String
    
    On Error GoTo errHandle
    
    strվ�� = zlDatabase.GetPara("�ϴ�ѡ��վ��", glngSys, mlngModule, "", Array(cboվ��, lblվ��), InStr(1, mstrPrivs, ";��������;") > 0)
    gstrSQL = "" & _
    "   Select Distinct q.���, Q.���� As վ������ " & _
     "  From ��������˵�� B, ���ű� A ,Zlnodelist Q " & _
     "  Where B.������� In (1, 2, 3) And B.�������� = '����' And B.����id = A.ID And A.վ�� = Q.��� And " & _
     "         (A.����ʱ�� = To_Date('3000-01-01', 'yyyy-mm-dd') Or A.����ʱ�� Is Null) " & _
     "    Order By ���"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    mblnNotClick = True
    With rsTemp
        cboվ��.Clear
        Do While Not .EOF
            cboվ��.AddItem NVL(rsTemp!���) & "-" & NVL(rsTemp!վ������)
            If cboվ��.ListIndex < 0 And NVL(rsTemp!���) = gstrNodeNo Then
                cboվ��.ListIndex = cboվ��.NewIndex
            End If
            If strվ�� = NVL(rsTemp!���) Then
                cboվ��.ListIndex = cboվ��.NewIndex
            End If
            .MoveNext
        Loop
        If cboվ��.ListIndex < 0 And cboվ��.ListCount > 0 Then cboվ��.ListIndex = 0
        picվ��.ZOrder
        picվ��.Visible = cboվ��.ListCount > 0
        lbl����.Visible = cboվ��.ListCount > 0
    End With
    mlngPreSelIdx = cboվ��.ListIndex
    mblnNotClick = False
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Public Function zlShowMe(ByVal frmMain As Form, lngModule As Long, strPrivs As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������
    '���:frmMain-���õ�������
    '        lngModule -ģ���
    '        strPrivs-Ȩ�޴�
    '����:
    '����:���óɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-01-20 09:39:27
    '����:35386
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mstrPrivs = strPrivs: mlngModule = lngModule: mblnOK = False: mblnChange = False
    Me.Show 1, frmMain
    zlShowMe = mblnOK
End Function

Private Sub InitGridData()
    Dim rsTemp As ADODB.Recordset
'    gstrSQL = "" & _
'    "   Select -1 as ID,'Z' as ����,'* ���� * ' as ���� From dual Union All " & _
'    "   Select A.ID,A.���� ,A.����||'-'||A.����  as ����" & _
'    "   From  ��������˵�� b,���ű� a " & _
'    "   Where B.������� in(1,2,3) And B.��������='����'  " & _
'    "           And  b.����ID=a.ID and " & Where����ʱ��("A") & _
'    "   Order by ����"
'    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    With vsPressMoney
        .Clear 1
'        If rsTemp.EOF Then
'            .ColComboList(.ColIndex("����")) = " "
'        Else
'            .ColComboList(.ColIndex("����")) = .BuildComboList(rsTemp, "����", "ID")
'        End If
        .ColComboList(.ColIndex("��������")) = "1-�ۼƷ���|2-ÿ�շ���"
        .ColComboList(.ColIndex("����")) = "..."
        .ColComboList(.ColIndex("������ʽ1")) = "..."
        .ColComboList(.ColIndex("������ʽ2")) = "..."
        .ColComboList(.ColIndex("������ʽ3")) = "..."
        mblnEdit = InStr(1, mstrPrivs, ";������������;") > 0
        If mblnEdit Then .Editable = flexEDKbdMouse
        cmdWarnNew.Enabled = mblnEdit
        cmdWarnDel.Enabled = mblnEdit
        cmdOK.Visible = mblnEdit
        zl_vsGrid_Para_Restore mlngModule, vsPressMoney, Me.Caption, "�����б�", False
    End With
 End Sub
Private Sub InitData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ������
    '����:���˺�
    '����:2011-01-20 09:31:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strCoding As String, i As Long
     
    On Error GoTo errHandle
    
   '���ʱ������
    gstrSQL = "Select RowID as ID,����,��� From �շ���� Order by ����"
    Set mrs��� = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    lst���.Clear
    lst���.AddItem "�������"
    Do While Not mrs���.EOF
        lst���.AddItem mrs���!���
        lst���.ItemData(lst���.NewIndex) = Asc(mrs���!����)
        mrs���.MoveNext
    Loop
    Call LoadScheme
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Sub LoadScheme()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ָ��վ��ķ���
    '����:���˺�
    '����:2011-02-12 10:46:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strCoding As String, i As Long
    
    On Error GoTo errHandle
    
    '�������ʱ�����
    Set mrsWarn = New ADODB.Recordset
    mrsWarn.Fields.Append "����ID", adBigInt, , adFldIsNullable
    mrsWarn.Fields.Append "������", adVarChar, 100, adFldIsNullable
    mrsWarn.Fields.Append "������", adVarChar, 100, adFldIsNullable
    mrsWarn.Fields.Append "���ò���", adVarChar, 100
    mrsWarn.Fields.Append "��������", adSmallInt
    mrsWarn.Fields.Append "����ֵ", adCurrency
    mrsWarn.Fields.Append "������־1", adVarChar, 100, adFldIsNullable
    mrsWarn.Fields.Append "������־2", adVarChar, 100, adFldIsNullable
    mrsWarn.Fields.Append "������־3", adVarChar, 100, adFldIsNullable
    mrsWarn.Fields.Append "�߿�����", adCurrency
    mrsWarn.Fields.Append "�߿��׼", adCurrency
    mrsWarn.CursorLocation = adUseClient
    mrsWarn.LockType = adLockOptimistic
    mrsWarn.CursorType = adOpenStatic
    mrsWarn.Open
    gstrSQL = "" & _
    "   Select a.����ID,B.����,b.���� as ����,a.���ò���,nvl(a.��������,1) as ��������, " & _
    "               a.����ֵ,a.������־1,a.������־2,a.������־3,A.�߿�����,a.�߿��׼ " & _
    "   From ���ʱ����� a,���ű� b " & _
    "   Where a.����ID= b.id(+)  " & IIf(cboվ��.ListCount > 0, " And b.վ��=[1]", "") & _
    "   Order by Decode(a.���ò���,'��ͨ����',1,'ҽ������',2,3),a.���ò���,B.���� Desc"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, CStr(Split(cboվ��.Text & "-", "-")(0)))
    strCoding = ",��ͨ����" '������һ����ͨ����
    Do Until rsTemp.EOF
        mrsWarn.AddNew
        mrsWarn!����ID = rsTemp!����ID
        mrsWarn!������ = rsTemp!����
        mrsWarn!������ = rsTemp!����
        mrsWarn!���ò��� = rsTemp!���ò���
        mrsWarn!�������� = rsTemp!��������
        mrsWarn!����ֵ = rsTemp!����ֵ
        mrsWarn!������־1 = rsTemp!������־1
        mrsWarn!������־2 = rsTemp!������־2
        mrsWarn!������־3 = rsTemp!������־3
        mrsWarn!�߿����� = Val(NVL(rsTemp!�߿�����))
        mrsWarn!�߿��׼ = Val(NVL(rsTemp!�߿��׼))
        mrsWarn.Update
        If InStr(strCoding & ",", "," & rsTemp!���ò��� & ",") = 0 Then
            strCoding = strCoding & "," & rsTemp!���ò���
        End If
        rsTemp.MoveNext
    Loop
    strCoding = Mid(strCoding, 2)
    tab����.Tabs.Clear
    For i = 0 To UBound(Split(strCoding, ","))
        tab����.Tabs.Add , , Split(strCoding, ",")(i)
    Next
    tab����.Tabs(1).Selected = True '֮ǰ���ἤ��Click�¼�,��Ϊ����
   mblnChange = False
   

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Sub AfterDeleteRow()
    'ɾ���к�
End Sub
Private Sub AfterAddRow(Row As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����Ӻ�
    '����:���˺�
    '����:2011-01-18 18:36:42
    '---------------------------------------------------------------------------------------------------------------------------------------------
     With vsPressMoney
        .Cell(flexcpData, Row, 0, Row, .Cols - 1) = ""
        .Cell(flexcpText, Row, 0, Row, .Cols - 1) = ""
        .TextMatrix(Row, .ColIndex("��������")) = "1-�ۼƷ���"
    End With
End Sub
Private Sub BeforeDeleteRow(Row As Long, Cancel As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ɾ����֮ǰ
    '����:���˺�
    '����:2011-01-18 18:37:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
     With vsPressMoney
        If .Editable = flexEDNone Then Exit Sub
        If Val(.Cell(flexcpData, Row, .ColIndex("����"))) <> 0 Then
            If MsgBox("���Ƿ����Ҫɾ������Ϊ��" & .TextMatrix(Row, .ColIndex("����")) & "���ķ�����¼��?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                Cancel = True
                Exit Sub
            End If
        End If
        mblnChange = True
    End With
End Sub

 
Private Sub cboվ��_Click()
    If mblnNotClick = True Then Exit Sub
    If mblnChange Then
        If mlngPreSelIdx <> cboվ��.ListIndex Then
             If MsgBox("ע��:" & vbCrLf & "     ���Ѿ�����������,�����ı�վ��,�����޸ĵķ�����Ϣ" & vbCrLf & _
                "    ���ᶪʧ,���Ƿ����Ҫ�ı�?", vbQuestion + vbDefaultButton2 + vbYesNo, gstrSysName) = vbNo Then
                cboվ��.ListIndex = mlngPreSelIdx: Exit Sub
             End If
        End If
    End If
    mlngPreSelIdx = cboվ��.ListIndex
    Call LoadScheme
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If Check���ʱ��� = False Then Exit Sub
    If SaveData = False Then Exit Sub
    mblnChange = False: mblnOK = True
    If picվ��.Visible Then
        MsgBox "��������ɹ�!", vbInformation + vbOKOnly, gstrSysName
    Else
        Unload Me
    End If
End Sub

Private Sub Form_Activate()
    If Not mblnFirst Then Exit Sub
    mblnFirst = False
    If vsPressMoney.Enabled And vsPressMoney.Visible Then vsPressMoney.SetFocus
    Call picDown_Resize
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyEscape
        If picList���.Visible Then
            picList���.Visible = False: Exit Sub
        End If
        Call cmdCancel_Click
    Case Else
    End Select
End Sub
 
Private Sub Form_Load()
    RestoreWinState Me, App.ProductName
    Call LoadClients
    Call InitGridData
    Call InitData
    mblnFirst = True
    mblnChange = False
     
End Sub
Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    picDown.Left = ScaleLeft
    picDown.Top = ScaleHeight - picDown.Height
    picDown.Width = ScaleWidth
    
    With tab����
        .Top = IIf(picվ��.Visible, picվ��.Top + picվ��.Height + 50, picվ��.Top)
        .Width = ScaleWidth - .Left - 50
        .Height = picDown.Top - .Top
        vsPressMoney.Width = ScaleWidth - vsPressMoney.Left - 120
        vsPressMoney.Height = picDown.Top - vsPressMoney.Top - 100
    End With
    picTop.Width = ScaleWidth - picTop.Left
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange = True Then
        If MsgBox("ע��:" & vbCrLf & "   ���Ѿ����Ĺ�����,�Ƿ����Ҫ�˳�?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = 1: Exit Sub
        End If
    End If
    zl_vsGrid_Para_Save mlngModule, vsPressMoney, Me.Caption, "�����б�", False
    Set mrsWarn = Nothing
    Set mrs��� = Nothing
    Call zlDatabase.SetPara("�ϴ�ѡ��վ��", CStr(Split(cboվ��.Text & "-", "-")(0)), glngSys, mlngModule, InStr(1, mstrPrivs, ";��������;") > 0)
    SaveWinState Me, App.ProductName
End Sub

Private Sub picDown_Resize()
    Err = 0: On Error Resume Next
    With picDown
        cmdCancel.Left = .ScaleWidth - cmdCancel.Width - 100
        If cboվ��.ListCount > 0 Then
            cmdOK.Top = cmdWarnDel.Top
            cmdOK.Left = .ScaleWidth - cmdOK.Width - 100
        Else
            cmdOK.Top = cmdCancel.Top
            cmdOK.Left = cmdCancel.Left - cmdOK.Width - 20
        End If
        fraSplit.Width = .ScaleWidth
    End With
End Sub

Private Sub picList���_Resize()
    Err = 0: On Error Resume Next
    With picList���
        shtCaption.Left = .ScaleLeft
        shtCaption.Width = .ScaleWidth: shtCaption.Top = .ScaleTop
        lst���.Left = .ScaleLeft: lst���.Width = .ScaleWidth
    End With
End Sub
 
Private Sub picTop_Resize()
    Err = 0: On Error Resume Next
    Line4.X2 = picTop.ScaleWidth + 30
End Sub

Private Sub vsPressMoney_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������صĸ�ʽ
    '����:���˺�
    '����:2011-01-18 18:32:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    With vsPressMoney
        Select Case Col
        Case .ColIndex("����")
            .ColComboList(Col) = "..."
'        Case .ColIndex("������ʽ1"), .ColIndex("������ʽ2"), .ColIndex("������ʽ3")
'            .ColComboList(Col) = "..."
        Case .ColIndex("����")
        Case .ColIndex("��������")
            If InStr(.TextMatrix(Row, .ColIndex("��������")), "ÿ�շ���") > 0 Then
                .TextMatrix(Row, .ColIndex("������ʽ2")) = ""   'ÿ�շ����ޱ�����ʽ2
                'Ϊ��ÿ�շ��á�ʱ�ж�һ�½���Ϊ����
                If IsNumeric(.TextMatrix(Row, .ColIndex("����ֵ"))) Then
                    If Val(.TextMatrix(Row, .ColIndex("����ֵ"))) < 0 Then
                        .TextMatrix(Row, .ColIndex("����ֵ")) = "0.00"
                    End If
                Else
                    .TextMatrix(Row, .ColIndex("����ֵ")) = "0.00"
                End If
            End If
        Case .ColIndex("����ֵ"), .ColIndex("�߿�����"), .ColIndex("�߿��׼")
            .TextMatrix(Row, Col) = Format(Val(.TextMatrix(Row, Col)), "###0.00;-###0.00;;")
        End Select
    End With
End Sub
Private Sub vsPressMoney_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
        If mblnSort = True Then Exit Sub
        Call zl_VsGridRowChange(vsPressMoney, OldRow, NewRow, OldCol, NewCol)
End Sub
Private Sub vsPressMoney_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    '--------------------------------------------------------------------------
    '����:��ťѡ��
    '����:
    '--------------------------------------------------------------------------
    Dim lngRow As Long
    With vsPressMoney
        Select Case Col
        Case .ColIndex("����")
             If Select����("") = False Then Exit Sub
            Call zlVsMoveGridCell(vsPressMoney, .ColIndex("����"), , mblnEdit, lngRow)
            If lngRow >= 0 Then AfterAddRow lngRow
        Case .ColIndex("������ʽ1"), .ColIndex("������ʽ2"), .ColIndex("������ʽ3")
            If Select�ı�����ʽ() = False Then Exit Sub
        End Select
    End With
    
End Sub
Private Sub vsPressMoney_ChangeEdit()
    mblnChange = True
End Sub

Private Sub vsPressMoney_DblClick()
    With vsPressMoney
      If .MouseCol <> .Cols - 1 And .MouseCol <> 1 Then Exit Sub
        If mblnEdit = False Then Exit Sub
        If .Col = 1 Then
            .TextMatrix(.Row, .ColIndex("��������")) = IIf(Left(.TextMatrix(.Row, .ColIndex("��������")), 1) = "1", "2-ÿ�շ���", "1-�ۼƷ���")
            If InStr(.TextMatrix(.Row, .ColIndex("��������")), "ÿ�շ���") > 0 Then
                .TextMatrix(.Row, .ColIndex("������ʽ2")) = ""   'ÿ�շ����ޱ�����ʽ2
                'Ϊ��ÿ�շ��á�ʱ�ж�һ�½���Ϊ����
                If IsNumeric(.TextMatrix(.Row, .ColIndex("����ֵ"))) Then
                    If Val(.TextMatrix(.Row, .ColIndex("����ֵ"))) < 0 Then
                        .TextMatrix(.Row, .ColIndex("����ֵ")) = "0.00"
                    End If
                Else
                    .TextMatrix(.Row, .ColIndex("����ֵ")) = "0.00"
                End If
            End If
        End If
        mblnChange = True
    End With
End Sub

Private Sub vsPressMoney_GotFocus()
    Call zl_VsGridGotFocus(vsPressMoney)
End Sub

Private Sub vsPressMoney_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngCol As Long, blnCancel As Boolean, lngRow As Long

    With vsPressMoney
        If KeyCode <> vbKeyReturn And KeyCode <> vbKeyReturn _
            And (KeyCode <> Asc("*")) And KeyCode <> vbKeySpace _
            And KeyCode <> vbKeyShift Then
            If Shift = 1 And (KeyCode = 56 Or KeyCode <> Asc("*")) Then
                vsPressMoney_CellButtonClick .Row, .Col
            Else

            Select Case .Col
            Case .ColIndex("����")  '.ColIndex("������ʽ1"), .ColIndex("������ʽ2"), .ColIndex("������ʽ3"),
                .ColComboList(.Col) = ""
            Case Else
            End Select
            End If
        End If

        If KeyCode = vbKeyDelete Then
            blnCancel = False
            'ɾ����ǰ
            Call BeforeDeleteRow(.Row, blnCancel)
            If blnCancel = True Then Exit Sub
            If .Row = .Rows - 1 And .Row = 1 Then
                For lngCol = 0 To .Cols - 1
                    .TextMatrix(.Row, lngCol) = ""
                    .Cell(flexcpData, .Row, lngCol) = ""
                Next
            Else
                .RemoveItem .Row
            End If
            'ɾ���к�
            Call AfterDeleteRow
        End If
    End With
    If KeyCode <> vbKeyReturn Then Exit Sub
    With vsPressMoney
        If Trim(.TextMatrix(.Row, .ColIndex("����"))) = "" Then
            zlCommFun.PressKey vbKeyTab
            Exit Sub
        End If
        Call zlVsMoveGridCell(vsPressMoney, .ColIndex("����"), , mblnEdit, lngRow)
        If lngRow >= 0 Then
            Call AfterAddRow(lngRow)
        End If
    End With
End Sub

Private Sub vsPressMoney_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    '�༭����
    Dim intCol As Integer, strKey As String, lngRow As Long

    If KeyCode <> vbKeyReturn Then Exit Sub

    With vsPressMoney
        Select Case Col
        Case .ColIndex("����")
            strKey = Trim(.EditText)
            strKey = Replace(strKey, Chr(vbKeyReturn), "")
            strKey = Replace(strKey, Chr(10), "")
            If strKey = "" Then Exit Sub
             If Select����(strKey) = False Then
                .TextMatrix(Row, Col) = .EditText: .Cell(flexcpData, Row, Col) = ""
                 Exit Sub
            End If
            .EditText = .TextMatrix(Row, Col)
'        Case .ColIndex("������ʽ1"), .ColIndex("������ʽ2"), .ColIndex("������ʽ3")
'            strKey = Trim(.EditText)
'            strKey = Replace(strKey, Chr(vbKeyReturn), "")
'            strKey = Replace(strKey, Chr(10), "")
'            If strKey = "" Then Exit Sub
''            If Select��������(strKey) = False Then
''                .TextMatrix(Row, Col) = .EditText: .Cell(flexcpData, Row, Col) = ""
''                Exit Sub
''            End If
'            .EditText = .TextMatrix(Row, Col)
        Case Else
        End Select
        Call zlVsMoveGridCell(vsPressMoney, .ColIndex("����"), -1, mblnEdit, lngRow)
        If lngRow >= 0 Then AfterAddRow lngRow
    End With
End Sub

Private Sub vsPressMoney_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
    With vsPressMoney
        '�л���������
        If .Col = .ColIndex("��������") Then
            Select Case KeyAscii
                Case Asc(" ")
                    '�л������־
                    Select Case Left(.TextMatrix(.Row, .Col), 1)
                        Case "1"
                            .TextMatrix(.Row, .Col) = "2-ÿ�շ���"
                        Case Else
                            .TextMatrix(.Row, .Col) = "1-�ۼƷ���"
                    End Select
                    mblnChange = True
                Case vbKey1
                    .TextMatrix(.Row, .Col) = "1-�ۼƷ���"
                    mblnChange = True
                Case vbKey2
                    .TextMatrix(.Row, .Col) = "2-ÿ�շ���"
                    mblnChange = True
            End Select
            If InStr(.TextMatrix(.Row, .Col), "ÿ�շ���") > 0 Then
                .TextMatrix(.Row, .ColIndex("������ʽ2")) = ""   'ÿ�շ����ޱ�����ʽ2
            End If
        End If
    End With
End Sub

Private Sub vsPressMoney_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    With vsPressMoney
        Select Case .Col
        Case .ColIndex("����")  '.ColIndex("������ʽ1"), .ColIndex("������ʽ2"), .ColIndex("������ʽ3"),
            VsFlxGridCheckKeyPress vsPressMoney, Row, Col, KeyAscii, m�ı�ʽ
        Case .ColIndex("����ֵ"), .ColIndex("�߿�����"), .ColIndex("�߿��׼")
            VsFlxGridCheckKeyPress vsPressMoney, Row, Col, KeyAscii, m���ʽ
        End Select
    End With
End Sub
Private Sub vsPressMoney_LeaveCell()
    If mblnSort Then Exit Sub
    zlCommFun.OpenIme False
End Sub
Private Sub vsPressMoney_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
        '���õ�Ԫ��ı༭����
        With vsPressMoney
           Select Case .Col
               Case .ColIndex("����") ' .ColIndex("������ʽ1"), .ColIndex("������ʽ2"), .ColIndex("������ʽ3")
                   .EditMaxLength = 100
               Case .ColIndex("����ֵ"), .ColIndex("�߿�����"), .ColIndex("�߿��׼")
                   .EditMaxLength = 16
           End Select
    End With
End Sub

Private Sub vsPressMoney_EnterCell()
    If mblnSort = True Then Exit Sub
    '�������޸ĲŴ�������
    If mblnEdit Then Exit Sub
    With vsPressMoney
        zlCommFun.OpenIme (False)
        Select Case .Col
        Case .ColIndex("����"), .ColIndex("������ʽ1"), .ColIndex("������ʽ2"), .ColIndex("������ʽ3")
             .ColComboList(.Col) = "..."
        End Select
    End With
End Sub

 Private Sub vsPressMoney_LostFocus()
    zlCommFun.OpenIme False
     Call zl_VsGridLOSTFOCUS(vsPressMoney)
End Sub
Private Sub vsPressMoney_Validate(Cancel As Boolean)
        Dim lngRow As Long
        If Not mblnChange Then Exit Sub
        If zlControl.MouseInRect(cmdCancel.hWnd) Then Exit Sub
        '�����ʱ�������
        If Not Check���ʱ��� Then Cancel = True: Exit Sub
        '�ռ����ʱ�������
        With mrsWarn
            .Filter = "���ò���='" & tab����.SelectedItem.Caption & "'"
            Do While Not .EOF
                .Delete
                .Update
                .MoveNext
            Loop
            .Filter = 0
        End With
        With vsPressMoney
            For lngRow = 1 To .Rows - 1
                If .TextMatrix(lngRow, .ColIndex("����")) <> "" And .TextMatrix(lngRow, .ColIndex("����ֵ")) <> "" Then
                    mrsWarn.AddNew
                    mrsWarn!���ò��� = tab����.SelectedItem.Caption
                    If Val(.Cell(flexcpData, lngRow, .ColIndex("����"))) <> 0 Then
                        mrsWarn!����ID = Val(.Cell(flexcpData, lngRow, .ColIndex("����")))
                        If mrsWarn!����ID <= 0 Then
                            mrsWarn!����ID = Null
                            mrsWarn!������ = Null
                            mrsWarn!������ = Trim(.TextMatrix(lngRow, .ColIndex("����")))
                        Else
                            mrsWarn!������ = Split(.TextMatrix(lngRow, .ColIndex("����")), "-")(0)
                            mrsWarn!������ = Split(.TextMatrix(lngRow, .ColIndex("����")), "-")(1)
                        End If
                    End If
                    mrsWarn!�������� = CInt(Left(.TextMatrix(lngRow, .ColIndex("��������")), 1))
                    mrsWarn!����ֵ = CCur(.TextMatrix(lngRow, .ColIndex("����ֵ")))

                    mrsWarn!������־1 = Get�����봮(.TextMatrix(lngRow, .ColIndex("������ʽ1")))
                    mrsWarn!������־2 = Get�����봮(.TextMatrix(lngRow, .ColIndex("������ʽ2")))
                    mrsWarn!������־3 = Get�����봮(.TextMatrix(lngRow, .ColIndex("������ʽ3")))
                    mrsWarn!�߿����� = FormatEx(Val(.TextMatrix(lngRow, .ColIndex("�߿�����"))), 2)
                    mrsWarn!�߿��׼ = FormatEx(Val(.TextMatrix(lngRow, .ColIndex("�߿��׼"))), 2)
                    mrsWarn.Update
                End If
            Next
        End With
End Sub

Private Sub vsPressMoney_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim strKey As String, intCol As Integer
    '������֤
    With vsPressMoney
        strKey = Trim(.EditText): strKey = Replace(strKey, Chr(vbKeyReturn), ""): strKey = Replace(strKey, Chr(10), "")
        Select Case Col
            Case .ColIndex("����ֵ"), .ColIndex("�߿�����"), .ColIndex("�߿��׼")
                If zlDblIsValid(strKey, 16, True, False, 0, .ColKey(Col)) = False Then
                    Cancel = True: Exit Sub
                End If
                If strKey <> "" Then
                    .EditText = Format(Val(strKey), "###0.00;-###0.00;;")
                End If
        End Select
    End With
End Sub
Private Sub vsPressMoney_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim i As Long, arrSplit As Variant
    With vsPressMoney
        If mblnEdit = False Then Cancel = True: Exit Sub
        Select Case Col
        Case .ColIndex("����")
        Case .ColIndex("������ʽ1"), .ColIndex("������ʽ3")
        Case .ColIndex("������ʽ2")
            'ÿ�շ��ò��ܱ༭������ʽ2
            If InStr(Trim(.TextMatrix(Row, .ColIndex("��������"))), "ÿ�շ���") > 0 Then Cancel = True: Exit Sub
        Case .ColIndex("����ֵ"), .ColIndex("�߿�����"), .ColIndex("�߿��׼")
        Case Else: Cancel = True
        End Select
    End With
End Sub
Private Sub tab����_Click()
    Dim lngRow As Long
    mrsWarn.Filter = "���ò���='" & tab����.SelectedItem.Caption & "'"
    With vsPressMoney
        If mrsWarn.RecordCount = 0 Then
            .Clear 1
            .Rows = 2: .Row = 1: .Col = 1
            .RowData(1) = ""
            .Cell(flexcpData, 1, 0, 1, .Cols - 1) = ""
            .Cell(flexcpText, 1, 0, 1, .Cols - 1) = ""
        Else
            .Clear 1
            .Rows = mrsWarn.RecordCount + 1: .Row = 1: .Col = 1
            lngRow = 1
            Do Until mrsWarn.EOF
                .RowData(lngRow) = NVL(mrsWarn!����ID, 0)
                .TextMatrix(lngRow, .ColIndex("����")) = IIf(IsNull(mrsWarn!����ID), "*����*", mrsWarn!������ & "-" & mrsWarn!������)
                .Cell(flexcpData, lngRow, .ColIndex("����")) = IIf(IsNull(mrsWarn!����ID), -1, NVL(mrsWarn!����ID, 0))
                .TextMatrix(lngRow, .ColIndex("��������")) = IIf(mrsWarn!�������� = 1, "1-�ۼƷ���", "2-ÿ�շ���")
                .TextMatrix(lngRow, .ColIndex("����ֵ")) = Format(mrsWarn!����ֵ, "###0.00;-###0.00;;")
                .TextMatrix(lngRow, .ColIndex("������ʽ1")) = Get������ƴ�(NVL(mrsWarn!������־1), mrs���)
                .TextMatrix(lngRow, .ColIndex("������ʽ2")) = Get������ƴ�(NVL(mrsWarn!������־2), mrs���)
                .TextMatrix(lngRow, .ColIndex("������ʽ3")) = Get������ƴ�(NVL(mrsWarn!������־3), mrs���)
                .TextMatrix(lngRow, .ColIndex("�߿�����")) = Format(mrsWarn!�߿�����, "###0.00;-###0.00;0.00;0.00")
                .TextMatrix(lngRow, .ColIndex("�߿��׼")) = Format(mrsWarn!�߿��׼, "###0.00;-###0.00;0.00;0.00")
                lngRow = lngRow + 1
                mrsWarn.MoveNext
            Loop
          If .Enabled And .Visible Then .SetFocus
        End If
    End With
End Sub
Private Function Check���ʱ���() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����ʱ����Ƿ���ȷ
    '����:��ȷ,����true,���򷵻�False
    '����:���˺�
    '����:2011-01-18 18:58:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngRow As Long, lngTemp As Long
    Dim lngCol1 As Long, lngCol2 As Long
    Dim arr���() As String

    With vsPressMoney
        For lngRow = 1 To .Rows - 1
            If Trim(.TextMatrix(lngRow, .ColIndex("����"))) <> "" Then
                For lngTemp = lngRow + 1 To .Rows - 1
                    If .TextMatrix(lngRow, .ColIndex("����")) = .TextMatrix(lngTemp, .ColIndex("����")) And .TextMatrix(lngTemp, 2) <> "" Then
                        MsgBox "������" & .TextMatrix(lngTemp, .ColIndex("����")) & "�����ֶ�Ρ�", vbExclamation, gstrSysName
                        .Row = lngTemp: .Col = .ColIndex("����"): .SetFocus: Exit Function
                    End If
                Next
                If Val(.TextMatrix(lngRow, .ColIndex("�߿�����"))) > 999999999 Or Val(.TextMatrix(lngRow, .ColIndex("�߿�����"))) < 0 Then
                    MsgBox "������" & .TextMatrix(lngRow, .ColIndex("����")) & "���еĴ߿�������������(Ӧ����0~999999999)!", vbExclamation, gstrSysName
                    .Row = lngRow: .Col = .ColIndex("�߿�����"): .SetFocus: Exit Function
                End If
                If Val(.TextMatrix(lngRow, .ColIndex("�߿��׼"))) > 999999999 Or Val(.TextMatrix(lngRow, .ColIndex("�߿��׼"))) < 0 Then
                    MsgBox "������" & .TextMatrix(lngRow, .ColIndex("����")) & "���еĴ߿��׼����(Ӧ����0~999999999)!", vbExclamation, gstrSysName
                    .Row = lngRow: .Col = .ColIndex("�߿��׼"): .SetFocus: Exit Function
                End If
            End If
        Next

        '���ͬһ������ͬ������ʽ������Ƿ�һ����û�����û��ظ�
        For lngRow = 1 To .Rows - 1
            If .TextMatrix(lngRow, .ColIndex("����")) <> "" And .TextMatrix(lngRow, .ColIndex("����ֵ")) <> "" Then
                If Trim(.TextMatrix(lngRow, .ColIndex("������ʽ1"))) = "" And Trim(.TextMatrix(lngRow, .ColIndex("������ʽ2"))) = "" And Trim(.TextMatrix(lngRow, .ColIndex("������ʽ3"))) = "" Then
                    MsgBox "������" & .TextMatrix(lngRow, .ColIndex("����")) & "��δ����Ҫ�������շ����", vbExclamation, gstrSysName
                    .Row = lngRow: .Col = .ColIndex("������ʽ1"): .SetFocus: Exit Function
                End If

                If (.TextMatrix(lngRow, .ColIndex("������ʽ1")) = "�������" And (Trim(.TextMatrix(lngRow, .ColIndex("������ʽ2"))) <> "" Or Trim(.TextMatrix(lngRow, .ColIndex("������ʽ3"))) <> "")) _
                    Or (.TextMatrix(lngRow, .ColIndex("������ʽ2")) = "�������" And (Trim(.TextMatrix(lngRow, .ColIndex("������ʽ1"))) <> "" Or Trim(.TextMatrix(lngRow, .ColIndex("������ʽ3"))) <> "")) _
                    Or (.TextMatrix(lngRow, .ColIndex("������ʽ3")) = "�������" And (Trim(.TextMatrix(lngRow, .ColIndex("������ʽ2"))) <> "" Or Trim(.TextMatrix(lngRow, .ColIndex("������ʽ1"))) <> "")) Then

                    MsgBox "������" & .TextMatrix(lngRow, .ColIndex("����")) & "����ͬ�ı�����ʽ������ͬ���շ����", vbExclamation, gstrSysName
                    .Row = lngRow: .Col = .ColIndex("������ʽ1"): .SetFocus: Exit Function
                End If
                If .TextMatrix(lngRow, .ColIndex("������ʽ1")) <> "�������" And Trim(.TextMatrix(lngRow, .ColIndex("������ʽ2"))) <> "�������" And Trim(.TextMatrix(lngRow, .ColIndex("������ʽ3"))) <> "�������" Then
                    For lngCol1 = .ColIndex("������ʽ1") To .ColIndex("������ʽ3")
                        If Trim(.TextMatrix(lngRow, lngCol1)) <> "" Then
                            For lngCol2 = .ColIndex("������ʽ1") To .ColIndex("������ʽ3")
                                If lngCol1 <> lngCol2 Then
                                    arr��� = Split(.TextMatrix(lngRow, lngCol1), ",")
                                    For lngTemp = 0 To UBound(arr���)
                                        If InStr("," & .TextMatrix(lngRow, lngCol2) & ",", "," & arr���(lngTemp) & ",") > 0 Then
                                            MsgBox "������" & .TextMatrix(lngRow, .ColIndex("����")) & "����ͬ�ı�����ʽ������ͬ���շ����", vbExclamation, gstrSysName
                                            .Row = lngRow: .Col = .ColIndex("������ʽ1"): .SetFocus: Exit Function
                                        End If
                                    Next
                                End If
                            Next
                        End If
                    Next
                End If
            End If
        Next
    End With

    Check���ʱ��� = True
End Function


Private Function Get�����봮(str��� As String) As String
'���ܣ���������"���,����"�Ĵ���������"CDEFG"�Ĵ�
    Dim i As Integer, j As Integer
    Dim arr���() As String, strTmp As String

    If Trim(str���) = "" Then Exit Function
    If str��� = "�������" Then
        Get�����봮 = "-"
    Else
        arr��� = Split(str���, ",")
        For i = 0 To UBound(arr���)
            For j = 1 To lst���.ListCount - 1
                If lst���.List(j) = arr���(i) Then
                    strTmp = strTmp & Chr(lst���.ItemData(j))
                    Exit For
                End If
            Next
        Next
        Get�����봮 = strTmp
    End If
End Function

Private Sub cmdWarnNew_Click()
    Dim strName As String, strCopy As String
    Dim strSchemes As String, i As Integer
    Dim rsCopy As ADODB.Recordset
    
    For i = 1 To tab����.Tabs.Count
        strSchemes = strSchemes & "," & tab����.Tabs(i).Caption
    Next
    
    strName = frmWarnEdit.ShowMe(Me, Mid(strSchemes, 2), strCopy)
    If strName = "" Then Exit Sub
    
    '��������
    Set rsCopy = mrsWarn.Clone
    rsCopy.Filter = "���ò���='" & strCopy & "'"
    Do While Not rsCopy.EOF
        mrsWarn.AddNew
        mrsWarn!���ò��� = strName
        mrsWarn!����ID = rsCopy!����ID
        mrsWarn!������ = rsCopy!������
        mrsWarn!������ = rsCopy!������
        mrsWarn!�������� = rsCopy!��������
        mrsWarn!����ֵ = rsCopy!����ֵ
        mrsWarn!������־1 = rsCopy!������־1
        mrsWarn!������־2 = rsCopy!������־2
        mrsWarn!������־3 = rsCopy!������־3
        mrsWarn!�߿����� = rsCopy!�߿�����
        mrsWarn!�߿��׼ = rsCopy!�߿��׼
        mrsWarn.Update
        rsCopy.MoveNext
    Loop
    
    tab����.Tabs.Add , , strName
    tab����.Tabs(tab����.Tabs.Count).Selected = True
    
    mblnChange = True
End Sub
Private Sub cmdWarnDel_Click()
    If tab����.SelectedItem.Caption = "��ͨ����" Then
        MsgBox """" & tab����.SelectedItem.Caption & """��������������ɾ����", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If MsgBox("ȷʵҪɾ��""" & tab����.SelectedItem.Caption & """����������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub

    With mrsWarn
        .Filter = "���ò���='" & tab����.SelectedItem.Caption & "'"
        
        '��¼ɾ�������ò�������
        If InStr(1, mstrDel���ò���, tab����.SelectedItem.Caption) = 0 Then
            mstrDel���ò��� = IIf(mstrDel���ò��� = "", "", mstrDel���ò��� & ";") & tab����.SelectedItem.Caption
        End If
        
        Do While Not .EOF
            .Delete
            .Update
            .MoveNext
        Loop
        .Filter = 0
    End With
    
    tab����.Tabs.Remove tab����.SelectedItem.Index
    tab����.Tabs(1).Selected = True
    
    mblnChange = True
End Sub
Private Function SaveData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݱ���
    '����:���˺�
    '����:2011-01-20 09:35:00
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str���ò��� As String, strTemp As String, i As Long
    Dim strվ�� As String
    
    On Error GoTo errHandle
    If cboվ��.ListCount = 0 Or cboվ��.ListIndex < 0 Then
        strվ�� = "NULL"
    Else
        strվ�� = "'" & Split(cboվ��.Text & "-", "-")(0) & "'"
    End If
    '�����ò��˷�������
    mrsWarn.Filter = 0
    For i = 1 To tab����.Tabs.Count
        strTemp = ""
        str���ò��� = tab����.Tabs.Item(i).Caption
        mrsWarn.Filter = "���ò���='" & str���ò��� & "'"
        Do While Not mrsWarn.EOF
            strTemp = strTemp & NVL(mrsWarn!����ID) & "," & mrsWarn!�������� & "," & _
            mrsWarn!����ֵ & "," & NVL(mrsWarn!������־1) & "," & NVL(mrsWarn!������־2) & "," & NVL(mrsWarn!������־3) & "," & NVL(mrsWarn!�߿�����) & "," & NVL(mrsWarn!�߿��׼) & ","
            mrsWarn.MoveNext
        Loop
        strTemp = str���ò��� & "|" & strTemp
        ' Zl_���ʱ�����_Modify
        gstrSQL = "zl_���ʱ�����_Modify("
        '  ������_In In Varchar2,
        gstrSQL = gstrSQL & "'" & strTemp & "',"
        '  վ��_In Varchar2:=Null
        gstrSQL = gstrSQL & "" & strվ�� & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    Next
    SaveData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Private Function Get������ƴ�(str��� As String, rs��� As ADODB.Recordset) As String
    '���ܣ�������"CDEFG"�����ת��Ϊ����"���,����..."��
    Dim i As Integer, strTmp As String
    If str��� = "" Then
        Get������ƴ� = " " 'Ϊ���ܰ��س�������
        Exit Function
    End If
    If str��� = "-" Then
        Get������ƴ� = "�������"
        Exit Function
    End If
    For i = 1 To Len(str���)
        rs���.Filter = "����='" & Mid(str���, i, 1) & "'"
        If Not rs���.EOF Then strTmp = strTmp & "," & rs���!���
    Next
    Get������ƴ� = Mid(strTmp, 2)
End Function
Private Function Select����(ByVal strSearch As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ѡ����
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-01-20 10:39:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim blnCancel As Boolean, strKey As String, strTittle As String, lngH As Long, strFind As String
    Dim sngX As Single, sngY As Single, bytStyle As Byte
    Dim rsTemp  As ADODB.Recordset
    'zlDatabase.ShowSelect
    '���ܣ��๦��ѡ����
    '������
    '     frmParent=��ʾ�ĸ�����
    '     strSQL=������Դ,��ͬ����ѡ������SQL�е��ֶ��в�ͬҪ��
    '     bytStyle=ѡ�������
    '       Ϊ0ʱ:�б���:ID,��
    '       Ϊ1ʱ:���η��:ID,�ϼ�ID,����,����(���blnĩ��������Ҫĩ���ֶ�)
    '       Ϊ2ʱ:˫����:ID,�ϼ�ID,����,����,ĩ������ListViewֻ��ʾĩ��=1����Ŀ
    '     strTitle=ѡ������������,Ҳ���ڸ��Ի�����
    '     blnĩ��=������ѡ����(bytStyle=1)ʱ,�Ƿ�ֻ��ѡ��ĩ��Ϊ1����Ŀ
    '     strSeek=��bytStyle<>2ʱ��Ч,ȱʡ��λ����Ŀ��
    '             bytStyle=0ʱ,��ID���ϼ�ID֮��ĵ�һ���ֶ�Ϊ׼��
    '             bytStyle=1ʱ,�����Ǳ��������
    '     strNote=ѡ������˵������
    '     blnShowSub=��ѡ��һ���Ǹ����ʱ,�Ƿ���ʾ�����¼������е���Ŀ(��Ŀ��ʱ����)
    '     blnShowRoot=��ѡ������ʱ,�Ƿ���ʾ������Ŀ(��Ŀ��ʱ����)
    '     blnNoneWin,X,Y,txtH=����ɷǴ�����,X,Y,txtH��ʾ���ý�������������(�������Ļ)�͸߶�
    '     Cancel=���ز���,��ʾ�Ƿ�ȡ��,��Ҫ����blnNoneWin=Trueʱ
    '     blnMultiOne=��bytStyle=0ʱ,�Ƿ񽫶Զ�����ͬ��¼����һ���ж�
    '     blnSearch=�Ƿ���ʾ�к�,�����������кŶ�λ
    '���أ�ȡ��=Nothing,ѡ��=SQLԴ�ĵ��м�¼��
    '˵����
    '     1.ID���ϼ�ID����Ϊ�ַ�������
    '     2.ĩ�����ֶβ�Ҫ����ֵ
    'Ӧ�ã������ڸ������������������Ǻܴ��ѡ����,����ƥ���б�ȡ�
    strTittle = "����ѡ��": bytStyle = 0
    strKey = gstrLike & strSearch & "%"
    strFind = ""
    If strSearch <> "" Then
        strFind = "   and  (a.���� like upper([1]) or a.���� like upper([1]) or a.���� like [1] )"
        If IsNumeric(strSearch) Then                         '���������,��ֻȡ����
            strFind = " And (A.���� Like Upper([1]))"
        ElseIf zlCommFun.IsCharAlpha(strSearch) Then           '01,11.����ȫ����ĸʱֻƥ�����
            '0-ƴ����,1-�����,2-����
            '.int���뷽ʽ = Val(zlDatabase.GetPara("���뷽ʽ" ))
            strFind = " And  (a.���� Like Upper([1]))"
        ElseIf zlCommFun.IsCharChinese(strSearch) Then  'ȫ����
            strFind = " And a.���� Like [1] "
        End If
    End If
    If strSearch = "" Then
        gstrSQL = "" & _
            "Select * " & _
            "  From (With M As (Select Distinct A.ID, -10 * Ascii(A.վ��) As �ϼ�id, A.����, A.����, A.����, A.վ��, 1 As ĩ��,Q.���� as վ������" & _
            "                   From ��������˵�� B, ���ű� A,Zlnodelist Q " & _
            "                   Where B.������� In (1, 2, 3) And B.�������� = '����' " & _
                                    IIf(cboվ��.ListCount > 0, " And A.վ��=[2] ", "") & " And B.����id = A.ID And a.վ��=Q.���(+) And " & _
            "                         (A.����ʱ�� = To_Date('3000-01-01', 'yyyy-mm-dd') Or A.����ʱ�� Is Null)) " & _
            "         Select -10 * Ascii(A.���) As ID, -1 * Null As �ϼ�id, To_Char(���) As ����, ����, '' As ����, " & _
            "                ��� as վ�� , 0 As ĩ��,���� As վ������  " & _
            "         From Zlnodelist A " & _
            "         Where Exists (Select 1 From M Where M.վ�� = A.���) " & _
            "         Union All " & _
            "         Select -1 As ID, -1 * Null As �ϼ�id, '-' As ����, '* ���� * ' As ����, 'MZ' As ����, '' As վ��, 1 As ĩ��,'' as վ������ " & _
            "         From Dual " & _
            "         Union All " & _
            "         Select * From M) " & _
            "   "
            bytStyle = 2
    Else
        gstrSQL = "" & _
          "   Select * From ( " & _
          "   Select -1 as ID,'Z' as RID,'-' as ����,'* ���� * ' as ����, 'MZ' as ����,'' as վ������ From dual Union All " & _
          "   Select distinct A.ID,A.���� as RID,A.���� ,A.����,A.����,M.���� as վ������ " & _
          "   From  ��������˵�� b,���ű� a,Zlnodelist M  " & _
          "   Where B.������� in(1,2,3) And B.��������='����'  " & IIf(cboվ��.ListCount > 0, " And A.վ��=[2] ", "") & _
          "           And A.վ��=M.���(+) And  b.����ID=a.ID and " & Where����ʱ��("A") & _
          "     ) A " & IIf(strSearch <> "", " Where 1=1 " & strFind, "") & _
          "   Order by RID"
    End If
    Call CalcPosition(sngX, sngY, vsPressMoney)
    lngH = vsPressMoney.CellHeight
    sngY = sngY - lngH
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, gstrSQL, bytStyle, strTittle, IIf(strSearch = "", True, False), "", "", False, IIf(strSearch = "", True, False), True, sngX, sngY, lngH, blnCancel, False, False, strKey, CStr(Split(cboվ��.Text & "-", "-")(0)))
    If blnCancel = True Then
        vsPressMoney.SetFocus
        Exit Function
    End If
    If rsTemp Is Nothing Then
        ShowMsgbox "û�����������Ĳ���,����!"
        If vsPressMoney.Enabled Then vsPressMoney.SetFocus
        Exit Function
    End If
    '����Ƿ����ظ��Ĳ���
    With vsPressMoney
        For i = 1 To .Rows - 1
            If i <> .Row Then
                If .Cell(flexcpData, i, .Col) = Val(rsTemp!ID) Then
                    MsgBox "�ڵ�: " & i & "���Ѿ�������ͬ�Ĳ���,������ѡ��ò���!", vbInformation + vbDefaultButton1 + vbOKOnly, gstrSysName
                    If vsPressMoney.Enabled Then vsPressMoney.SetFocus
                    Exit Function
                End If
            End If
        Next
    End With
    vsPressMoney.SetFocus
    With vsPressMoney
        .TextMatrix(.Row, .Col) = IIf(NVL(rsTemp!����) = "-", "", NVL(rsTemp!����) & "-") & NVL(rsTemp!����)
        .Cell(flexcpData, .Row, .Col) = Val(rsTemp!ID)
    End With
    'zlVsMoveGridCell vsPressMoney, vsPressMoney.ColIndex("����"), mblnEdit, i
    Select���� = True
End Function
Private Function Select�ı�����ʽ() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������ʽѡ����
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-01-20 10:39:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
    With vsPressMoney
        Call Set���ѡ��(.TextMatrix(.Row, .Col))
        picList���.Left = .Left + .CellLeft
        If .Top + .CellTop + .CellHeight + picList���.Height <= .Container.Height Then
            picList���.Top = .Top + .CellTop + .CellHeight
        Else
            picList���.Top = .Top + .CellTop - .Height - 30
        End If
        picList���.Width = IIf(.CellWidth < 1200, 1200, .CellWidth + 30)
        picList���.ZOrder
        picList���.Visible = True
        lst���.SetFocus
    End With
    Select�ı�����ʽ = True
End Function

Private Sub Set���ѡ��(str��� As String)
'���ܣ���������"���,����..."�Ĵ������б��ѡ�����
    Dim i As Integer, j As Integer
    Dim arr���() As String
    
    For i = 0 To lst���.ListCount - 1
        lst���.Selected(i) = False
    Next
    
    If Trim(str���) = "" Then
        Exit Sub
    ElseIf str��� = "�������" Then
        For i = 0 To lst���.ListCount - 1
            lst���.Selected(i) = (i = 0)
        Next
    Else
        lst���.Selected(0) = False
        arr��� = Split(str���, ",")
        For i = 0 To UBound(arr���)
            For j = 1 To lst���.ListCount - 1
                If lst���.List(j) = arr���(i) Then
                    lst���.Selected(j) = True: Exit For
                End If
            Next
        Next
    End If
    
    For i = 0 To lst���.ListCount - 1
        If lst���.Selected(i) Then
            lst���.TopIndex = i: Exit For
        End If
    Next
End Sub
Private Sub lst���_ItemCheck(Item As Integer)
    Dim i As Integer
    If Item = 0 And lst���.Selected(Item) Then
        For i = 1 To lst���.ListCount - 1
            lst���.Selected(i) = False
        Next
    ElseIf Item > 0 And lst���.Selected(Item) Then
        lst���.Selected(0) = False
    End If
End Sub

Private Sub lst���_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call lst���_Validate(False)
        Call Form_KeyDown(vbKeyEscape, 0)
    End If
End Sub

Private Sub lst���_LostFocus()
    picList���.Visible = False
End Sub
Private Sub lst���_Validate(Cancel As Boolean)
    Dim i As Integer
    With vsPressMoney
        .TextMatrix(.Row, .Col) = Get���ѡ��
        If .TextMatrix(.Row, .Col) = "�������" Then
            For i = .ColIndex("������ʽ1") To .ColIndex("������ʽ3")
                If i <> .Col Then .TextMatrix(.Row, i) = " "
            Next
        End If
    End With
    mblnChange = True
End Sub
Private Function Get���ѡ��() As String
'���ܣ��������ѡ���ѡ��������������"���,����..."�Ĵ�
    Dim i As Integer, strTmp As String
    
    If lst���.Selected(0) Then
        Get���ѡ�� = "�������"
    Else
        For i = 1 To lst���.ListCount - 1
            If lst���.Selected(i) Then
                strTmp = strTmp & "," & lst���.List(i)
            End If
        Next
        Get���ѡ�� = Mid(strTmp, 2)
        If Get���ѡ�� = "" Then Get���ѡ�� = " " 'Ϊ���ܻس�������
    End If
End Function
