VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "VSFLEX8.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmChargeWholeSetItemEdit 
   Caption         =   "�����շ���Ŀ�༭"
   ClientHeight    =   8430
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11745
   Icon            =   "frmChargeWholeSetItemEdit.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8430
   ScaleWidth      =   11745
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picUserDept 
      BorderStyle     =   0  'None
      Height          =   2460
      Left            =   495
      ScaleHeight     =   2460
      ScaleWidth      =   9855
      TabIndex        =   31
      Top             =   4620
      Width           =   9855
      Begin VB.TextBox txt���� 
         Height          =   300
         Left            =   495
         MaxLength       =   50
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   45
         Width           =   3720
      End
      Begin VB.CommandButton cmdSel 
         Caption         =   "&L"
         Height          =   300
         Left            =   4200
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   45
         Width           =   285
      End
      Begin VB.CommandButton cmdDelete 
         Cancel          =   -1  'True
         Caption         =   "ɾ��(&D)"
         Height          =   350
         Left            =   5850
         TabIndex        =   23
         Top             =   0
         Width           =   1100
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "����(&A)"
         Height          =   350
         Left            =   4710
         TabIndex        =   22
         Top             =   0
         Width           =   1100
      End
      Begin MSComctlLib.ListView lvw���� 
         Height          =   4095
         Left            =   0
         TabIndex        =   24
         Top             =   420
         Width           =   11355
         _ExtentX        =   20029
         _ExtentY        =   7223
         View            =   2
         Arrange         =   1
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         _Version        =   393217
         Icons           =   "iltdept"
         SmallIcons      =   "iltdept"
         ForeColor       =   -2147483640
         BackColor       =   -2147483634
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Left            =   75
         TabIndex        =   19
         Top             =   105
         Width           =   360
      End
   End
   Begin VB.PictureBox picWholeSet 
      BorderStyle     =   0  'None
      Height          =   2040
      Left            =   675
      ScaleHeight     =   2040
      ScaleWidth      =   11145
      TabIndex        =   30
      Top             =   3270
      Width           =   11145
      Begin VSFlex8Ctl.VSFlexGrid vsWholeSet 
         Height          =   4680
         Left            =   195
         TabIndex        =   18
         Top             =   195
         Width           =   11355
         _cx             =   20029
         _cy             =   8255
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
         BackColorBkg    =   -2147483634
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   4
         Cols            =   22
         FixedRows       =   1
         FixedCols       =   2
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmChargeWholeSetItemEdit.frx":0442
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
         ExplorerBar     =   2
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
   End
   Begin VB.PictureBox picCmd 
      BorderStyle     =   0  'None
      Height          =   660
      Left            =   0
      ScaleHeight     =   660
      ScaleWidth      =   11745
      TabIndex        =   28
      Top             =   7770
      Width           =   11745
      Begin VB.CommandButton cmdCancel 
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   10560
         TabIndex        =   26
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         Height          =   350
         Left            =   9330
         TabIndex        =   25
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdHelp 
         Caption         =   "����(&H)"
         Height          =   350
         Left            =   120
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   120
         Width           =   1100
      End
   End
   Begin VB.Frame fra���� 
      Caption         =   "������Ŀ��Ϣ"
      Height          =   1950
      Left            =   120
      TabIndex        =   27
      Top             =   255
      Width           =   11460
      Begin VB.TextBox txtMemo 
         Height          =   300
         Left            =   900
         MaxLength       =   100
         TabIndex        =   17
         Tag             =   "��ע"
         Top             =   1485
         Width           =   3795
      End
      Begin VB.ComboBox cbo��Ա 
         Height          =   300
         Left            =   5895
         TabIndex        =   15
         Text            =   "cbo��Ա"
         Top             =   1110
         Width           =   4005
      End
      Begin VB.OptionButton opt��Χ 
         Caption         =   "��Ժ"
         Height          =   315
         Index           =   2
         Left            =   3255
         TabIndex        =   13
         Top             =   1125
         Width           =   1530
      End
      Begin VB.OptionButton opt��Χ 
         Caption         =   "ָ������"
         Height          =   315
         Index           =   1
         Left            =   2025
         TabIndex        =   12
         Top             =   1110
         Width           =   1530
      End
      Begin VB.OptionButton opt��Χ 
         Caption         =   "ָ����Ա"
         Height          =   315
         Index           =   0
         Left            =   900
         TabIndex        =   11
         Top             =   1110
         Value           =   -1  'True
         Width           =   2655
      End
      Begin VB.TextBox txtWB 
         Height          =   300
         Left            =   7860
         MaxLength       =   20
         TabIndex        =   9
         Tag             =   "���"
         Top             =   720
         Width           =   1425
      End
      Begin VB.TextBox txtParent 
         Height          =   300
         Left            =   5895
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   3
         TabStop         =   0   'False
         Text            =   "(��)"
         ToolTipText     =   "��Del����ϼ������ó�������"
         Top             =   285
         Width           =   3720
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "&P"
         Height          =   300
         Left            =   9600
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   285
         Width           =   285
      End
      Begin VB.TextBox txtName 
         Height          =   300
         Left            =   915
         MaxLength       =   100
         TabIndex        =   6
         Tag             =   "����"
         Top             =   750
         Width           =   3795
      End
      Begin VB.TextBox txtSymbol 
         Height          =   300
         Left            =   5895
         MaxLength       =   20
         TabIndex        =   8
         Tag             =   "ƴ��"
         Top             =   720
         Width           =   1425
      End
      Begin VB.TextBox txtCode 
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   930
         MaxLength       =   10
         TabIndex        =   1
         TabStop         =   0   'False
         Tag             =   "����"
         Text            =   "0000"
         Top             =   360
         Width           =   1620
      End
      Begin VB.Label lblMemo 
         AutoSize        =   -1  'True
         Caption         =   "��ע(&M)"
         Height          =   180
         Left            =   180
         TabIndex        =   16
         Top             =   1530
         Width           =   630
      End
      Begin VB.Label lbl��Ա 
         AutoSize        =   -1  'True
         Caption         =   "ָ����Ա"
         Height          =   180
         Left            =   5070
         TabIndex        =   14
         Top             =   1170
         Width           =   720
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "ʹ�÷�Χ"
         Height          =   180
         Left            =   105
         TabIndex        =   10
         Top             =   1170
         Width           =   720
      End
      Begin VB.Label lblParent 
         AutoSize        =   -1  'True
         Caption         =   "����(&U)"
         Height          =   180
         Left            =   5175
         TabIndex        =   2
         Top             =   345
         Width           =   630
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         Caption         =   "����(&N)"
         Height          =   180
         Left            =   195
         TabIndex        =   5
         Top             =   795
         Width           =   630
      End
      Begin VB.Label lblCode 
         AutoSize        =   -1  'True
         Caption         =   "����(&D)"
         Height          =   180
         Left            =   210
         TabIndex        =   0
         Top             =   420
         Width           =   630
      End
      Begin VB.Label lblSymbol 
         AutoSize        =   -1  'True
         Caption         =   "����(&S)                 (ƴ��)                (���)"
         Height          =   180
         Left            =   5160
         TabIndex        =   7
         Top             =   780
         Width           =   4680
      End
   End
   Begin XtremeSuiteControls.TabControl tbPage 
      Height          =   4110
      Left            =   105
      TabIndex        =   32
      Top             =   2355
      Width           =   11595
      _Version        =   589884
      _ExtentX        =   20452
      _ExtentY        =   7250
      _StockProps     =   64
   End
   Begin MSComctlLib.ImageList iltdept 
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
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeWholeSetItemEdit.frx":074A
            Key             =   "Dept"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmChargeWholeSetItemEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Enum EditTypeWhileSetItem
    EdI_���� = 1
    EdI_�޸� = 2
    EdI_�鿴 = 3
End Enum
Private mEditType As EditTypeWhileSetItem
Private mstrPrivs As String, mlngModule As Long
Private mstrWholeItems As String
Private mstrID As String, mlng����ID As Long
Private mintSucces As Integer
Private mblnFirst As Boolean
Private mblnChange As Boolean
Private mblnSort As Boolean
Private mbln�޸� As Boolean
Private Enum mItemPage
    pg_������� = 1
    pg_ʹ�ÿ��� = 2
End Enum
Private mrsDept As ADODB.Recordset
Private Sub zlInitClassPage()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ������ҳ��
    '����:���˺�
    '����:2010-08-24 10:15:11
    '˵��:27327
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, ObjItem As TabControlItem
    Err = 0: On Error GoTo ErrHand:
 
    Set ObjItem = tbPage.InsertItem(mItemPage.pg_�������, "������Ŀ���(&1)", picWholeSet.hwnd, 0)
    ObjItem.Tag = mItemPage.pg_�������
    Set ObjItem = tbPage.InsertItem(mItemPage.pg_ʹ�ÿ���, "ִ�п���(&2)", picUserDept.hwnd, 0)
    ObjItem.Tag = mItemPage.pg_ʹ�ÿ���
    tbPage.Item(0).Selected = True
     With tbPage
        .PaintManager.Appearance = xtpTabAppearancePropertyPage
        .PaintManager.BoldSelected = True
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.StaticFrame = True
        .PaintManager.ClientFrame = xtpTabFrameBorder
    End With
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Public Function ShowCard(ByVal frmMain As Form, ByVal EditType As EditTypeWhileSetItem, ByVal strPrivs As String, ByVal lngModule As Long, _
    Optional lng����id As Long = 0, Optional strID As String = "", Optional strWholeItems As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������,��ʾ��༭��ص���Ϣ����
    '���:frmMain-����������
    '       EditType-�༭����
    '       strWholeItems-�Ӽ��ʵ��д���ĵ�������,��ʽΪ:
    '                              ���,����,�շ�ϸĿID,����,����,ִ�п���|���,����,�շ�ϸĿID,����,����,ִ�п���|��
    '����:strID-���ص�ǰ�༭��ID
    '����:
    '����:���˺�
    '����:2010-08-26 17:00:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mEditType = EditType: mstrPrivs = strPrivs: mlngModule = lngModule: mlng����ID = lng����id: mstrID = strID
    mstrWholeItems = strWholeItems: mintSucces = 0
    Me.Show 1, frmMain
    ShowCard = mintSucces > 0
    strID = mstrID
End Function
Private Sub InitDefaultLen()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��Ĭ�����ݿⳤ��
    '����:���˺�
    '����:2010-08-26 17:08:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    On Error GoTo ErrHandle
    gstrSQL = "" & _
    " Select A.ID,A.����ID,A.����,A.����,A.ƴ��,A.���,A.��ע" & _
    " From �����շ���Ŀ A" & _
    " Where id=0"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mstrID))
    Me.txtCode.MaxLength = rsTemp.Fields("����").DefinedSize
    Me.txtName.MaxLength = rsTemp.Fields("����").DefinedSize
    Me.txtSymbol.MaxLength = rsTemp.Fields("ƴ��").DefinedSize
    Me.txtWB.MaxLength = rsTemp.Fields("���").DefinedSize
    Me.txtMemo.MaxLength = rsTemp.Fields("��ע").DefinedSize
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
Private Sub AnalyzeWholeSetItem()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ֽ⴫��ĳ�����Ŀ����
    '����:���˺�
    '����:2010-08-26 17:17:50
    '˵��:mstrWholeItems�ĸ�ʽΪ:���,����,�շ�ϸĿID,����,����,����,ִ�п���|���,����,�շ�ϸĿID,����,����,ִ�п���|��
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varData As Variant, varTemp As Variant, i As Long, j As Long, m As Long, lngId As Long
    Dim strValue(0 To 10) As String, strSubItem As String, str�շ�ϸĿID As String, strִ�п���ID As String
    Dim strDeptValue(0 To 10) As String, strDeptSub As String, lng���� As Long
    Dim rsItems As ADODB.Recordset, rsDept As ADODB.Recordset, rsOthers As ADODB.Recordset
    Dim lng����ID As Long
    Dim cllTemp As Collection
    
    If mstrWholeItems = "" Then Exit Sub
    
    On Error GoTo ErrHandle
    
    '�ȷֽ����,�ٲ���
    varData = Split(mstrWholeItems, "|")
    For i = 0 To UBound(varData)
        '���,����,�շ�ϸĿID,����,����,����,ִ�п���
        varTemp = Split(varData(i) & ",,,,,", ",")
        If Len(str�շ�ϸĿID) > 1990 And j <= 10 Then
            strValue(j) = Mid(str�շ�ϸĿID, 2)
            strSubItem = strSubItem & " Union ALL " & _
            " Select Column_Value as �շ�ϸĿID From Table(f_Num2List([" & j + 1 & "])) B "
            str�շ�ϸĿID = "": j = j + 1
        End If
        str�շ�ϸĿID = str�շ�ϸĿID & "," & Val(varTemp(2))
        If Len(strִ�п���ID) > 1990 And m <= 10 Then
            strDeptValue(m) = Mid(strִ�п���ID, 2)
            strDeptSub = strDeptSub & " Union ALL " & _
            " Select Column_Value as ִ�в���ID From Table(f_Num2List([" & m + 1 & "])) B "
            m = m + 1
            strִ�п���ID = ""
        Else
            strִ�п���ID = strִ�п���ID & "," & Val(varTemp(6))
        End If
    Next
    If str�շ�ϸĿID <> "" Then
        If j > 10 Then
             strSubItem = strSubItem & " UNION ALL Select ID From �շ���ĿĿ¼ Where id in (" & Mid(str�շ�ϸĿID, 2) & ")"
        Else
            strValue(j) = Mid(str�շ�ϸĿID, 2)
            strSubItem = strSubItem & " Union ALL " & _
            " Select Column_Value as �շ�ϸĿID From Table(f_Num2List([" & j + 1 & "])) B "
        End If
    End If
    If strִ�п���ID <> "" Then
        If m > 10 Then
             strDeptSub = strDeptSub & " UNION ALL Select ID From ���ű� Where id in (" & Mid(strִ�п���ID, 2) & ")"
        Else
            strDeptValue(m) = Mid(strִ�п���ID, 2)
            strDeptSub = strDeptSub & " Union ALL " & _
            " Select Column_Value as ִ�в���ID From Table(f_Num2List([" & m + 1 & "])) B "
        End If
    End If
    
    gstrSQL = "" & _
       "   Select A.����id, A.����id, A.���д���, A.�������� " & _
       "   From �շѴ�����Ŀ A, (" & Mid(strSubItem, 11) & ") D" & _
       "   Where A.����id = D.�շ�ϸĿid"
    Set rsOthers = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strValue(0), strValue(1), strValue(2), strValue(3), strValue(4), strValue(5), strValue(6), strValue(7), strValue(8), strValue(9), strValue(10))
    
    gstrSQL = "" & _
    "   Select A.���, A.ID,A.����,A.����,A.���㵥λ,A.���,B.��ҩ��̬,c.���� as ���Ʊ���," & _
    "             C.���� as ��������,C.���㵥λ as ������λ,B.ҩ��Id,B.����ϵ��,A.ִ�п���,A.�Ƿ���, B1.��������" & _
    "   From �շ���ĿĿ¼ A,ҩƷ��� B,�������� B1,������ĿĿ¼ C,(" & Mid(strSubItem, 11) & ") D" & _
    "   Where A.ID=b.ҩƷID(+) and A.ID=b1.����ID(+) and B.ҩ��Id=C.ID(+) and A.id=d.�շ�ϸĿID "
    Set rsItems = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strValue(0), strValue(1), strValue(2), strValue(3), strValue(4), strValue(5), strValue(6), strValue(7), strValue(8), strValue(9), strValue(10))
    If strDeptSub <> "" Then
        gstrSQL = "" & _
        "   Select A.ID,A.����,A.���� " & _
        "   From ���ű� A,(" & Mid(strDeptSub, 11) & ") D" & _
        "   Where A.id =D.ִ�в���ID"
    Else
        gstrSQL = "" & _
        "   Select A.ID,A.����,A.���� " & _
        "   From ���ű� A " & _
        "   Where A.id =0"
    End If
    Set rsDept = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strDeptValue(0), strDeptValue(1), strDeptValue(2), strDeptValue(3), strDeptValue(4), strDeptValue(5), strDeptValue(6), strDeptValue(7), strDeptValue(8), strDeptValue(9), strDeptValue(10))
    With vsWholeSet
        .Clear 1
        .OutlineBar = flexOutlineBarSimple
        .OutlineCol = .ColIndex("��־"): .SubtotalPosition = flexSTAbove
        .Rows = IIF(UBound(varData) = 0, 1, UBound(varData) + 1) + 1
        str�շ�ϸĿID = "": strִ�п���ID = "": j = 0: m = 0
        Set cllTemp = New Collection
        
        For i = 0 To UBound(varData)
            '���,����,�շ�ϸĿID,����,����,����,ִ�п���|
            varTemp = Split(varData(i) & ",,,,,", ",")
            .TextMatrix(i + 1, .ColIndex("���")) = i + 1
            .Cell(flexcpData, i + 1, .ColIndex("���")) = i + 1 ' Val(varTemp(0))
            cllTemp.Add i + 1, "_" & Val(varTemp(0))
            
            If Val(varTemp(1)) = 0 Then
                .TextMatrix(i + 1, .ColIndex("��������")) = ""
            Else
                .TextMatrix(i + 1, .ColIndex("��������")) = cllTemp("_" & Val(varTemp(1)))
            End If
            .Cell(flexcpData, i + 1, .ColIndex("�շ���Ŀ")) = Val(varTemp(2))
            
            .TextMatrix(i + 1, .ColIndex("ȱʡ����")) = IIF(Val(varTemp(3)) = 0, 1, Val(varTemp(3)))
            
            .TextMatrix(i + 1, .ColIndex("ȱʡ����")) = FormatEx(Val(varTemp(4)), 5)
            .Cell(flexcpData, i + 1, .ColIndex("ȱʡ����")) = Val(varTemp(4))
            .TextMatrix(i + 1, .ColIndex("ȱʡ�۸�")) = FormatEx(Val(varTemp(5)), 8)
            .Cell(flexcpData, i + 1, .ColIndex("ȱʡ�۸�")) = Val(varTemp(5))
            .Cell(flexcpData, i + 1, .ColIndex("ȱʡִ�п���")) = Val(varTemp(6))
            
            lngId = Val(.Cell(flexcpData, i + 1, .ColIndex("�շ���Ŀ")))
            If Val(.TextMatrix(i + 1, .ColIndex("��������"))) = 0 Then
                lng����ID = lngId
            End If
            
            rsItems.Find "ID=" & lngId, , adSearchForward, 1
            If rsItems.EOF = False Then
                .TextMatrix(i + 1, .ColIndex("�շ���Ŀ")) = NVL(rsItems!����) & "-" & NVL(rsItems!����)
                .TextMatrix(i + 1, .ColIndex("���")) = NVL(rsItems!���)
                .TextMatrix(i + 1, .ColIndex("ҩ��ID")) = NVL(rsItems!ҩ��ID)
                .TextMatrix(i + 1, .ColIndex("��������")) = Val(NVL(rsItems!��������))
                If NVL(rsItems!���) = "7" Then
                    '��ҩ,��ʾ��������
                    .TextMatrix(i + 1, .ColIndex("ҩ��")) = NVL(rsItems!���Ʊ���) & "-" & NVL(rsItems!��������)
                    .TextMatrix(i + 1, .ColIndex("��λ")) = NVL(rsItems!������λ)
                    .TextMatrix(i + 1, .ColIndex("ȱʡ����")) = FormatEx(Val(.TextMatrix(i + 1, .ColIndex("ȱʡ����"))) * Val(NVL(rsItems!����ϵ��)), 5)
                    .TextMatrix(i + 1, .ColIndex("ȱʡ�۸�")) = FormatEx(Val(.TextMatrix(i + 1, .ColIndex("ȱʡ�۸�"))) / Val(NVL(rsItems!����ϵ��)), 8)
                    .TextMatrix(i + 1, .ColIndex("��ҩ��̬")) = Val(NVL(rsItems!��ҩ��̬))
                    .TextMatrix(i + 1, .ColIndex("����ϵ��")) = Val(NVL(rsItems!����ϵ��))
                    
'                    If Val(Nvl(rsItems!��ҩ��̬)) = 0 Then   'ɢװ��̬��,����ʾ����Ĺ��
'                        .TextMatrix(i + 1, .ColIndex("���")) = Nvl(rsItems!���)
'                    End If
                Else
                    .TextMatrix(i + 1, .ColIndex("��λ")) = NVL(rsItems!���㵥λ)
                End If
                .TextMatrix(i + 1, .ColIndex("���")) = NVL(rsItems!���)
                .TextMatrix(i + 1, .ColIndex("�Ƿ���")) = NVL(rsItems!�Ƿ���)
                .TextMatrix(i + 1, .ColIndex("ִ�п���")) = NVL(rsItems!ִ�п���)
            End If
            rsDept.Find "ID=" & Val(.Cell(flexcpData, i + 1, .ColIndex("ȱʡִ�п���"))), , adSearchForward, 1
            If Not rsDept.EOF Then
                .TextMatrix(i + 1, .ColIndex("ȱʡִ�п���")) = NVL(rsDept!����) & "-" & NVL(rsDept!����)
            End If
            
            If Not rsOthers Is Nothing And Val(.TextMatrix(i + 1, .ColIndex("��������"))) <> 0 Then
                '  "   Select A.����id, A.����id, A.���д���, A.�������� "
                rsOthers.Filter = "����ID=" & lng����ID & " And ����ID= " & Val(.Cell(flexcpData, i + 1, .ColIndex("�շ���Ŀ")))
                If Not rsOthers.EOF Then
                    .TextMatrix(i + 1, .ColIndex("��������")) = Val(NVL(rsOthers!��������))
                    .Cell(flexcpData, i + 1, .ColIndex("��������")) = Val(NVL(rsOthers!���д���))
                End If
            End If
            
            If Val(.TextMatrix(i + 1, .ColIndex("��������"))) = 0 Then
                    lng���� = i + 1 'Val(.Cell(flexcpData, i + 1, .ColIndex("���")))
                    .IsSubtotal(i + 1) = True: .RowOutlineLevel(i + 1) = 1
            ElseIf lng���� = Val(.TextMatrix(i + 1, .ColIndex("��������"))) Then
                If i + 1 > 2 Then
                   If Val(.TextMatrix(i, .ColIndex("��������"))) <> 0 Then
                        .IsSubtotal(i) = False
                        .RowOutlineLevel(i) = 2
                    End If
                End If
                .IsSubtotal(i + 1) = True
                .RowOutlineLevel(i + 1) = 2
            End If
        Next
    End With
 
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Sub ClearCardData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����Ƭ����
    '����:���˺�
    '����:2010-08-27 10:01:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    txtCode.Text = "": txtName.Text = ""
    txtMemo.Text = "": txtSymbol = ""
    txtWB.Text = ""
    cbo��Ա.Text = "": txt����.Text = ""
    cbo��Ա.ListIndex = -1
    lvw����.ListItems.Clear
    vsWholeSet.Clear 1: vsWholeSet.Rows = 2
    vsWholeSet.TextMatrix(1, vsWholeSet.ColIndex("���")) = 1
    vsWholeSet.ColWidth(vsWholeSet.ColIndex("��־")) = 240
End Sub

Private Sub EditStatusSet()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ñ༭״̬
    '����:���˺�
    '����:2010-08-27 10:24:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCtl As Control
    If mEditType = EdI_���� Then
        With vsWholeSet
            .OutlineBar = flexOutlineBarSimple
            .OutlineCol = .ColIndex("��־"): .SubtotalPosition = flexSTAbove
            .Editable = flexEDKbdMouse
        End With
            opt��Χ(1).Enabled = InStr(1, mstrPrivs, ";���Ƴ��׷���;") > 0
            opt��Χ(2).Enabled = InStr(1, mstrPrivs, ";ȫԺ���׷���;") > 0
'            If opt��Χ(1).Enabled And opt��Χ(1).value = True Then
'               opt��Χ(0).value = True
'            End If
'            If opt��Χ(2).Enabled And opt��Χ(2).value = True Then
'               opt��Χ(0).value = True
'            End If
            txtCode.Enabled = True
    ElseIf mEditType = EdI_�޸� And mbln�޸� Then
        
        With vsWholeSet
            .OutlineBar = flexOutlineBarSimple
            .OutlineCol = .ColIndex("��־"): .SubtotalPosition = flexSTAbove
            .Editable = flexEDKbdMouse
        End With
            opt��Χ(1).Enabled = InStr(1, mstrPrivs, ";���Ƴ��׷���;") > 0
            opt��Χ(2).Enabled = InStr(1, mstrPrivs, ";ȫԺ���׷���;") > 0
'            If opt��Χ(1).Enabled And opt��Χ(1).value = True Then
'               opt��Χ(0).value = True
'            End If
'            If opt��Χ(2).Enabled And opt��Χ(2).value = True Then
'               opt��Χ(0).value = True
'            End If
            txtCode.Enabled = True
    Else
        With vsWholeSet
            .OutlineBar = flexOutlineBarSimple
            .OutlineCol = .ColIndex("���"): .SubtotalPosition = flexSTAbove
            .Editable = flexEDNone
        End With
        
        For Each objCtl In Me.Controls
            Select Case UCase(TypeName(objCtl))
            Case "TEXTBOX"
                objCtl.Enabled = False
                objCtl.BackColor = Me.BackColor
            Case UCase("OptionButton")
                objCtl.Enabled = False
            Case UCase("ComBox")
                objCtl.Enabled = False
                objCtl.BackColor = Me.BackColor
            Case UCase("CommandButton")
                If Not (objCtl Is cmdOK Or objCtl Is cmdCancel Or objCtl Is cmdHelp) Then
                        objCtl.Enabled = False
                End If
            Case UCase("vsFlexGrid")
                objCtl.Editable = flexEDNone
            Case Else
            End Select
        Next
        
        Me.cbo��Ա.Enabled = False
        cbo��Ա.BackColor = Me.BackColor
    End If
End Sub

Private Sub zlDefaultCode()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ȱʡ����
    '����:���˺�
    '����:2010-08-27 12:00:41
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim strTemp As String
    Dim lngLen As String, strUpCode As String
    On Error GoTo ErrHandle
    If Val(txtParent.Tag) = 0 Then
NotNO:
        gstrSQL = "Select Max(����) as ���� From ������Ŀ����  Where �ϼ�ID is null  "
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
        txtParent.Text = "��": txtParent.Tag = ""
        If NVL(rsTemp!����) = "" Then
            txtCode.Text = "001"
        Else
            strTemp = Val(rsTemp!����) + 1
            If Len(strTemp) > Len(rsTemp!����) Then
                txtCode.Text = strTemp
            Else
                 txtCode.Text = String(Len(rsTemp!����) - Len(strTemp), "0") & strTemp
            End If
        End If
        Exit Sub
    End If
    
    gstrSQL = "Select ID,����,���� From ������Ŀ����  Where ID=[1] "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(txtParent.Tag))
    If rsTemp.EOF Then
        GoTo NotNO:
    End If
    
    strUpCode = NVL(rsTemp!����)
    txtParent.Text = NVL(rsTemp!����) & "-" & NVL(rsTemp!����)
    txtParent.Tag = NVL(rsTemp!ID)
    gstrSQL = "select max(����) as ����" & _
            " From �����շ���Ŀ" & _
            " Where   ���� like [1] And ����<> [2] "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strUpCode & "%", strUpCode)
    If rsTemp.EOF Then
         txtCode.Text = strUpCode & "01"
    Else
        strTemp = NVL(rsTemp!����)
        If strTemp = "" Then
            txtCode.Text = strUpCode & "01"
        Else
            txtCode.Text = Val(strTemp) + 1
            lngLen = Len(strTemp) - Len(txtCode)
            If lngLen > 0 Then
                txtCode.Text = String(lngLen, "0") & txtCode.Text
            End If
         End If
   
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub


Private Function ReadCard() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��Ŀ��Ϣ
    '����:��ȡ�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2010-08-26 13:35:22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, rsOthers As ADODB.Recordset
    Dim lngRow As Long, lng���� As Long, i As Long, j As Long
    Dim lng����ID As Long
    
    Call EditStatusSet  '���ñ༭״̬
    If mEditType = EdI_���� Then
        Call ClearCardData
        txtParent.Tag = mlng����ID
        Call zlDefaultCode
        If mstrWholeItems <> "" Then Call AnalyzeWholeSetItem
        Call opt��Χ_Click(0)
        ReadCard = True
        Exit Function
    End If
    Screen.MousePointer = vbHourglass
    Call ClearCardData
    
    On Error GoTo ErrHandle
    If mEditType <> EdI_�鿴 Then
        gstrSQL = "" & _
         "   Select A.����id, A.����id, A.���д���, A.�������� " & _
         "   From �շѴ�����Ŀ A, �����շ���Ŀ��� B " & _
         "   Where ����id = B.�շ�ϸĿid And Nvl(B.��������, 0) = 0 And B.����id = [1]"
        Set rsOthers = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mstrID))
    Else
        Set rsOthers = Nothing
    End If
    
    Dim strWherePriceGrade As String
    If gstr��ͨ�۸�ȼ� = "" And gstrҩƷ�۸�ȼ� = "" And gstr���ļ۸�ȼ� = "" Then
        strWherePriceGrade = " And j.�۸�ȼ� Is Null"
    Else
        strWherePriceGrade = "" & _
            " And ((Instr(';5;6;7;', ';' || k.��� || ';') > 0 And j.�۸�ȼ� = [2])" & vbNewLine & _
            "      Or (Instr(';4;', ';' || k.��� || ';') > 0 And j.�۸�ȼ� = [3])" & vbNewLine & _
            "      Or (Instr(';4;5;6;7;', ';' || k.��� || ';') = 0 And j.�۸�ȼ� = [4])" & vbNewLine & _
            "      Or (j.�۸�ȼ� Is Null" & vbNewLine & _
            "          And Not Exists (Select 1" & vbNewLine & _
            "                          From �շѼ�Ŀ" & vbNewLine & _
            "                          Where j.�շ�ϸĿid = �շ�ϸĿid And Sysdate Between ִ������ And Nvl(��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD'))" & vbNewLine & _
            "                                And ((Instr(';5;6;7;', ';' || k.��� || ';') > 0 And �۸�ȼ� = [2])" & vbNewLine & _
            "                                      Or (Instr(';4;', ';' || k.��� || ';') > 0 And �۸�ȼ� = [3])" & vbNewLine & _
            "                                      Or (Instr(';4;5;6;7;', ';' || k.��� || ';') = 0 And �۸�ȼ� = [4])))))"
    End If
    gstrSQL = "" & _
    "   Select  /*+Rule */ A.����ID,A.�շ�ϸĿID,A.���,A.��������,A.����,A.����,A.����,A.ִ�п���ID, " & _
    "              B.���,B.����,B.����,B.���㵥λ,B.���,C.��ҩ��̬,D.���� as ���Ʊ���, " & _
    "              D.���� as ��������,D.���㵥λ as ������λ,C.����ϵ��, " & _
    "              E.���� As ִ�п��ұ���,E.���� As ִ�п�������, " & _
    "              M.���� As ���ױ���,M.���� As ��������,M.ƴ��,M.���,M.��ע,M.��Χ, " & _
    "              M.����ID,M.��ԱID,G.����,J.���� As �������,J.���� As �������� ,B.�Ƿ���,B.ִ�п���,C.ҩ��ID,B1.��������," & _
    "              Decode(B.�Ƿ���,1,'ʱ��',LTrim(To_Char(J1.�ּ�,'999999999.9999999'))) as �ּ�" & _
    "   From �����շ���Ŀ M,������Ŀ���� J,�����շ���Ŀ��� A,�շ���ĿĿ¼ B,�������� B1,ҩƷ��� C,������ĿĿ¼ D, " & _
    "             ���ű� E,��Ա�� G, " & _
    "             (Select j.�շ�ϸĿid, Sum(j.�ּ�) as �ּ�" & vbNewLine & _
    "              From �շѼ�Ŀ J,�շ���ĿĿ¼ K" & vbNewLine & _
    "              Where j.�շ�ϸĿID = k.ID And Sysdate Between J.ִ������ And Nvl(J.��ֹ����, To_Date('3000-1-1', 'YYYY-MM-DD'))" & vbNewLine & _
                         strWherePriceGrade & vbNewLine & _
    "              Group By j.�շ�ϸĿid ) J1 " & _
    "   Where   M.����ID=J.Id And  M.��ԱID=G.Id(+) And M.Id=A.����ID(+)  " & _
    "               And A.�շ�ϸĿid=J1.�շ�ϸĿID(+)" & _
    "               And A.�շ�ϸĿid=b.Id(+)  and a.�շ�ϸĿID=b1.����ID(+) And a.�շ�ϸĿID=C.ҩƷID(+) And C.ҩ��ID=D.Id(+) " & _
    "               And A.ִ�п���ID=E.Id(+)  " & _
    "               And M.ID=[1] " & _
    "   Order by A.���"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mstrID), gstrҩƷ�۸�ȼ�, gstr���ļ۸�ȼ�, gstr��ͨ�۸�ȼ�)
    
    If rsTemp.EOF Then
        MsgBox "�ó����շ���Ŀ�����Ѿ�������ɾ��,���ܽ����޸Ļ�鿴!", vbInformation + vbDefaultButton1, gstrSysName
        Exit Function
    End If
    mlng����ID = Val(NVL(rsTemp!����id))
    txtParent.Text = IIF(mlng����ID = 0, "��", NVL(rsTemp!�������) & "-" & NVL(rsTemp!��������))
    txtParent.Tag = mlng����ID
    txtCode.Text = NVL(rsTemp!���ױ���)
    txtCode.Tag = NVL(rsTemp!���ױ���)
    txtName.Text = NVL(rsTemp!��������)
    txtName.Tag = NVL(rsTemp!��������)
    txtSymbol.Text = NVL(rsTemp!ƴ��)
    txtWB.Text = NVL(rsTemp!���)
    txtMemo.Text = NVL(rsTemp!��ע)
    
     '0-ȫԺ;1-����;2-����Ա
    If Val(NVL(rsTemp!��Χ)) = 0 Then 'ȫԺ
        opt��Χ(2).value = True
    ElseIf Val(NVL(rsTemp!��Χ)) = 1 Then
        opt��Χ(1).value = True
    Else
        opt��Χ(0).value = True
        cbo��Ա.AddItem NVL(rsTemp!����)
        cbo��Ա.ItemData(cbo��Ա.NewIndex) = Val(NVL(rsTemp!��ԱID))
        cbo��Ա.ListIndex = cbo��Ա.NewIndex
        cbo��Ա.Tag = cbo��Ա.ListIndex
    End If
    
    '�������
    With vsWholeSet
        .Clear 1
        .Rows = IIF(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        lngRow = 1
        .OutlineBar = flexOutlineBarSimple
        .OutlineCol = .ColIndex("��־"): .SubtotalPosition = flexSTAbove
        Do While Not rsTemp.EOF
            .TextMatrix(lngRow, .ColIndex("���")) = IIF(Val(NVL(rsTemp!���)) = 0, lngRow, Val(NVL(rsTemp!���)))
            .Cell(flexcpData, lngRow, .ColIndex("���")) = Val(NVL(rsTemp!���))
            .TextMatrix(lngRow, .ColIndex("��������")) = Val(NVL(rsTemp!��������))
            .Cell(flexcpData, lngRow, .ColIndex("�շ���Ŀ")) = Val(NVL(rsTemp!�շ�ϸĿid))
            .TextMatrix(lngRow, .ColIndex("ȱʡ����")) = IIF(Val(NVL(rsTemp!����)) = 0, 1, Val(NVL(rsTemp!����)))
            .TextMatrix(lngRow, .ColIndex("ȱʡ����")) = FormatEx(Val(NVL(rsTemp!����)), 5)
            .Cell(flexcpData, lngRow, .ColIndex("ȱʡ����")) = Val(NVL(rsTemp!����))
            .TextMatrix(lngRow, .ColIndex("ȱʡ�۸�")) = FormatEx(Val(NVL(rsTemp!����)), 8)
            .Cell(flexcpData, lngRow, .ColIndex("ȱʡ�۸�")) = Val(NVL(rsTemp!����))
            .TextMatrix(lngRow, .ColIndex("ȱʡִ�п���")) = IIF(NVL(rsTemp!ִ�п��ұ���) = "", "", NVL(rsTemp!ִ�п��ұ���) & "-") & NVL(rsTemp!ִ�п�������)
            .Cell(flexcpData, lngRow, .ColIndex("ȱʡִ�п���")) = Val(NVL(rsTemp!ִ�п���ID))
            .TextMatrix(lngRow, .ColIndex("���")) = NVL(rsTemp!���)
            .TextMatrix(lngRow, .ColIndex("�Ƿ���")) = NVL(rsTemp!�Ƿ���)
            .TextMatrix(lngRow, .ColIndex("ִ�п���")) = NVL(rsTemp!ִ�п���)
            .TextMatrix(lngRow, .ColIndex("�ּ�")) = IIF(NVL(rsTemp!�ּ�) = "ʵ��", "ʵ��", FormatEx(Val(NVL(rsTemp!�ּ�)), 5))
            .TextMatrix(lngRow, .ColIndex("�շ���Ŀ")) = NVL(rsTemp!����) & "-" & NVL(rsTemp!����)
            .TextMatrix(lngRow, .ColIndex("���")) = NVL(rsTemp!���)
            .TextMatrix(lngRow, .ColIndex("ҩ��")) = ""
            .TextMatrix(lngRow, .ColIndex("ҩ��ID")) = NVL(rsTemp!ҩ��ID)
            .TextMatrix(lngRow, .ColIndex("��������")) = Val(NVL(rsTemp!��������))
            If NVL(rsTemp!���) = "7" Then
                '��ҩ,��ʾ��������
                .TextMatrix(lngRow, .ColIndex("ҩ��")) = NVL(rsTemp!���Ʊ���) & "-" & NVL(rsTemp!��������)
                .TextMatrix(lngRow, .ColIndex("��λ")) = NVL(rsTemp!������λ)
                .TextMatrix(lngRow, .ColIndex("ȱʡ����")) = FormatEx(Val(.TextMatrix(lngRow, .ColIndex("ȱʡ����"))) * Val(NVL(rsTemp!����ϵ��)), 5)
                .TextMatrix(lngRow, .ColIndex("ȱʡ�۸�")) = FormatEx(Val(.TextMatrix(lngRow, .ColIndex("ȱʡ�۸�"))) / Val(NVL(rsTemp!����ϵ��)), 8)
                .TextMatrix(lngRow, .ColIndex("��ҩ��̬")) = Val(NVL(rsTemp!��ҩ��̬))
                .TextMatrix(lngRow, .ColIndex("����ϵ��")) = Val(NVL(rsTemp!����ϵ��))
'                If Val(Nvl(rsTemp!��ҩ��̬)) = 0 Then   'ɢװ��̬��,����ʾ����Ĺ��
'                    .TextMatrix(lngRow, .ColIndex("���")) = Nvl(rsTemp!���)
'                End If
            Else
                .TextMatrix(lngRow, .ColIndex("��λ")) = NVL(rsTemp!���㵥λ)
            End If
            
            If Val(.TextMatrix(lngRow, .ColIndex("��������"))) = 0 Then
                    lng���� = Val(.Cell(flexcpData, lngRow, .ColIndex("���")))
                    lng����ID = Val(.Cell(flexcpData, lngRow, .ColIndex("�շ���Ŀ")))
                    .IsSubtotal(lngRow) = True: .RowOutlineLevel(lngRow) = 1
            ElseIf lng���� = Val(.TextMatrix(lngRow, .ColIndex("��������"))) Then
                If lngRow > 2 Then
                    If Val(.TextMatrix(lngRow - 1, .ColIndex("��������"))) <> 0 Then
                        .IsSubtotal(lngRow - 1) = False
                        .RowOutlineLevel(lngRow - 1) = 2
                    End If
                End If
                .IsSubtotal(lngRow) = True
                .RowOutlineLevel(lngRow) = 2
            End If
            If Not rsOthers Is Nothing And Val(.TextMatrix(lngRow, .ColIndex("��������"))) <> 0 Then
                '  "   Select A.����id, A.����id, A.���д���, A.�������� "
                rsOthers.Filter = "����ID=" & lng����ID & " And ����ID= " & Val(.Cell(flexcpData, lngRow, .ColIndex("�շ���Ŀ")))
                If Not rsOthers.EOF Then
                    .TextMatrix(lngRow, .ColIndex("��������")) = Val(NVL(rsOthers!��������))
                    .Cell(flexcpData, lngRow, .ColIndex("��������")) = Val(NVL(rsOthers!���д���))
                End If
            End If
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
    End With
    '����ʹ�ò���
    lvw����.ListItems.Clear
    If opt��Χ(1).value = True Then
         gstrSQL = "" & _
         "   Select A.����ID,B.����,b.���� " & _
         "   From ������Ŀʹ�ÿ��� A,���ű�  B  " & _
         "   Where a.����id=b.Id And a.����ID=[1]" & _
         "   Order By ����"
         Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mstrID))
        With lvw����
             Do While Not rsTemp.EOF
                 .ListItems.Add , "K" & rsTemp!����ID, NVL(rsTemp!����) & "-" & NVL(rsTemp!����), "Dept", "Dept"
                 rsTemp.MoveNext
             Loop
        End With
    End If
    Screen.MousePointer = vbDefault
    Call opt��Χ_Click(0)
    
    mbln�޸� = True
    If opt��Χ(0).value = True And InStr(mstrPrivs, "�޸ĸ��˳��׷���") < 1 Then
        mbln�޸� = False
    ElseIf opt��Χ(1).value = True And InStr(mstrPrivs, "�޸Ŀ��ҳ��׷���") < 1 Then
        mbln�޸� = False
    ElseIf opt��Χ(2).value = True And InStr(mstrPrivs, "�޸�ȫԺ���׷���") < 1 Then
        mbln�޸� = False
    End If
    
    Call EditStatusSet  '���ñ༭״̬
    ReadCard = True
    Exit Function
ErrHandle:
    Screen.MousePointer = vbDefault
    If ErrCenter() = 1 Then
        Screen.MousePointer = vbHourglass
        Resume
    End If
End Function

Private Sub cbo��Ա_Change()
    mblnChange = True
End Sub

Private Sub cbo��Ա_Click()
    'ѡ����ص���Ա
    If cbo��Ա.ListIndex <> -1 Then
        If cbo��Ա.ItemData(cbo��Ա.ListIndex) = 0 Then
             If SearchPerson("") = False Then
             End If
        Else
             cbo��Ա.Tag = cbo��Ա.ListIndex
        End If
    End If
End Sub
Private Function SearchPerson(ByVal strInput As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ָ����Ա
    '���:strInput-��������
    '����:
    '����:
    '����:���˺�
    '����:2010-08-30 16:52:22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strKey As String, strWhere As String
    Dim vRect As RECT, bytStyle As Byte, intIdx As Integer
    Dim blnCancel As Boolean
    
    On Error GoTo ErrHandle
    strKey = gstrLike & strInput & "%"
    strWhere = "": bytStyle = 0
    If strInput <> "" Then
        If IsNumeric(strInput) Then
            strWhere = " And A.��� Like [1]"
        ElseIf zlStr.IsCharAlpha(strInput) Then
            strWhere = "  And A.���� Like upper([1])"
        Else
            strWhere = " And (A.��� Like [1] or A.���� Like upper([1]) or A.���� like [1] )"
        End If
        '2010-12-28 �޸�(34325)
'        strWhere = strWhere & IIF(gstrNodeNo <> "-", " And (A.վ��='" & gstrNodeNo & "' or a.վ�� is NULL  )", "")
        
        gstrSQL = "" & _
            "   Select A.ID,A.���,A.����,A.����,A.����,A.�Ա�,A.����,A.��������,A.�칫�ҵ绰,A.ִҵ���,A.����ְ��,A.רҵ����ְ��" & _
            "   From ��Ա�� A " & _
            "   Where   (A.����ʱ�� >= To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null) " & strWhere & _
            "   order by A.���"
    Else
'        gstrSQL = "" & _
'        "   Select id," & IIF(gstrNodeNo <> "-", "1 as ����ID,-1*NULL as �ϼ�ID", "Level as ����ID,�ϼ�id") & " ,����,����,0 ĩ��,'' as ����,'' as ����,''as �Ա�,''as ����, to_date(Null,'yyyy-mm-dd')  as ��������, '' as  �칫�ҵ绰 ,'' ִҵ���, '' ����ְ��,'' רҵ����ְ��" & _
'        "   From ���ű� " & _
'        "   Where ����ʱ�� is null or ����ʱ��>=to_date('3000-01-01','yyyy-mm-dd') " & IIF(gstrNodeNo <> "-", " And (A.վ��='" & gstrNodeNo & "' or a.վ�� is NULL ) ", "") & _
'            IIF(gstrNodeNo <> "-", "", "   Start with �ϼ�id is null connect by prior id=�ϼ�id ") & _
'        "   union all " & _
'        "   Select a.ID,999999 AS ����ID,b.����id as �ϼ�ID,a.���,a.����,1 as ĩ��,����,����,�Ա�,����,��������,�칫�ҵ绰,A.ִҵ���,A.����ְ��,A.רҵ����ְ�� " & _
'        "   From ��Ա�� a,������Ա b  " & _
'        "   Where a.id=b.��Աid and b.ȱʡ=1  " & IIF(gstrNodeNo <> "-", " And (A.վ��='" & gstrNodeNo & "' or a.վ�� is NULL  )", "") & _
'        "         And (a.����ʱ�� >= To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null) " & _
'        "   Order by ����ID,����"
        
        gstrSQL = "" & _
        "   Select id,1 as ����ID,-1*NULL as �ϼ�ID,����,����,0 ĩ��,'' as ����,'' as ����,''as �Ա�,''as ����, to_date(Null,'yyyy-mm-dd')  as ��������, '' as  �칫�ҵ绰 ,'' ִҵ���, '' ����ְ��,'' רҵ����ְ��" & _
        "   From ���ű� " & _
        "   Where ����ʱ�� is null or ����ʱ��>=to_date('3000-01-01','yyyy-mm-dd') " & _
        "   union all " & _
        "   Select a.ID,999999 AS ����ID,b.����id as �ϼ�ID,a.���,a.����,1 as ĩ��,����,����,�Ա�,����,��������,�칫�ҵ绰,A.ִҵ���,A.����ְ��,A.רҵ����ְ�� " & _
        "   From ��Ա�� a,������Ա b  " & _
        "   Where a.id=b.��Աid and b.ȱʡ=1 " & _
        "         And (a.����ʱ�� >= To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null) " & _
        "   Order by ����ID,����"
        
        bytStyle = 2
    End If
    
    vRect = zlControl.GetControlRect(cbo��Ա.hwnd)
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, gstrSQL, bytStyle, "��Աѡ��", False, "", "", False, False, _
        True, vRect.Left, vRect.Top, txtParent.Height, blnCancel, False, True, strKey)
    
    
    If blnCancel = True Then
        If cbo��Ա.Enabled And cbo��Ա.Visible Then cbo��Ա.SetFocus
        Call cbo.SetIndex(cbo��Ա.hwnd, Val(cbo��Ա.Tag))
        Exit Function
    End If
    If rsTemp Is Nothing Then
        MsgBox "δ�ҵ�ƥ���ʹ�ÿ���,����!", vbInformation + vbDefaultButton1, gstrSysName
        If cbo��Ա.Enabled And cbo��Ա.Visible Then txtParent.SetFocus
        Call cbo.SetIndex(cbo��Ա.hwnd, Val(cbo��Ա.Tag))
        Exit Function
    End If
    If rsTemp.State <> 1 Then
        MsgBox "δ�ҵ�ƥ���ʹ�ÿ���,����!", vbInformation + vbDefaultButton1, gstrSysName
        If cbo��Ա.Enabled And cbo��Ա.Visible Then txtParent.SetFocus
        Call cbo.SetIndex(cbo��Ա.hwnd, Val(cbo��Ա.Tag))
        Exit Function
    End If
    If bytStyle = 0 Then
        intIdx = cbo.FindIndex(cbo��Ա, Val(NVL(rsTemp!ID)))
        If intIdx <> -1 Then
            cbo��Ա.ListIndex = intIdx
            cbo��Ա.Tag = cbo��Ա.ListIndex
        Else
            cbo��Ա.AddItem rsTemp!��� & "-" & rsTemp!����, 0
            cbo��Ա.ItemData(cbo��Ա.NewIndex) = rsTemp!ID
            cbo��Ա.ListIndex = cbo��Ա.NewIndex
            cbo��Ա.Tag = cbo��Ա.ListIndex
        End If
    Else
        intIdx = cbo.FindIndex(cbo��Ա, Val(NVL(rsTemp!ID)))
        If intIdx <> -1 Then
            cbo��Ա.ListIndex = intIdx
            cbo��Ա.Tag = cbo��Ա.ListIndex
        Else
            cbo��Ա.AddItem rsTemp!���� & "-" & rsTemp!����, 0
            cbo��Ա.ItemData(cbo��Ա.NewIndex) = rsTemp!ID
            cbo��Ա.ListIndex = cbo��Ա.NewIndex
            cbo��Ա.Tag = cbo��Ա.ListIndex
        End If
    End If
    OS.PressKey vbKeyTab
    SearchPerson = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Sub cbo��Ա_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Trim(cbo��Ա.Text) <> "" And cbo��Ա.ListIndex >= 0 Then
        OS.PressKey vbKeyTab
        Exit Sub
    End If
    If SearchPerson(Trim(cbo��Ա.Text)) = False Then
        Exit Sub
    End If
End Sub

Private Sub cbo��Ա_Validate(Cancel As Boolean)
    If cbo��Ա.ListIndex < 0 Then
        If cbo��Ա.Text <> "" Then
            Call cbo.SetIndex(cbo��Ա.hwnd, Val(cbo��Ա.Tag))
        End If
    End If
End Sub

Private Sub cmdAdd_Click()
    Dim ObjItem As ListItem
    If Val(txt����.Tag) = 0 Then Exit Sub
    If Trim(txt����.Text) = "" Then Exit Sub
    With lvw����
        For Each ObjItem In .ListItems
            If Val(Mid(ObjItem.Key, 2)) = Val(txt����.Tag) Then
                MsgBox "ע��:" & vbCrLf & "    �ÿ����Ѿ�����,����������!", vbInformation + vbOKOnly, gstrSysName
                txt����.SetFocus
                Exit Sub
            End If
        Next
        .ListItems.Add , "K" & txt����.Tag, txt����.Text, "Dept", "Dept"
        txt����.SetFocus
    End With
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    Dim intIndex As Integer, strKey As String
    If lvw����.SelectedItem Is Nothing Then Exit Sub
     With lvw����
        intIndex = .SelectedItem.Index
        .ListItems.Remove .SelectedItem.Key
        If .ListItems.Count > 0 Then
            intIndex = IIF(.ListItems.Count > intIndex, intIndex, .ListItems.Count)
            .ListItems(intIndex).Selected = True
            .ListItems(intIndex).EnsureVisible
        End If
    End With
    Call Setʹ�ÿ���Enable
End Sub
Private Sub Setʹ�ÿ���Enable()
    '����ʹ�ÿ��ҵ���ؿؼ�״̬
    cmdDelete.Enabled = Not Me.lvw����.SelectedItem Is Nothing
End Sub
Private Sub cmdOK_Click()
    If CheckValied = False Then Exit Sub
    If SaveData = False Then Exit Sub
    mintSucces = mintSucces + 1
    If mEditType = EdI_�޸� Then
        Unload Me: mblnChange = False
        Exit Sub
    End If
    Call ClearCardData
    If txtName.Enabled And txtName.Visible Then txtName.SetFocus
End Sub
Private Function CheckValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����������Ч��
    '����:������Ч,����true,���򷵻�Flase
    '����:���˺�
    '����:2010-08-30 15:11:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCtl As Control, lngRow As Long, blnHaveDate As Boolean
    Dim rsTemp As ADODB.Recordset
    
    If mEditType = EdI_�鿴 Then Exit Function
        
    For Each objCtl In Me.Controls
        Select Case UCase(TypeName(objCtl))
        Case UCase("TextBox")
            If objCtl Is txtCode Or objCtl Is txtName Then
                    If Trim(objCtl.Text) = "" Then
                        MsgBox "ע��:" & vbCrLf & "    " & objCtl.Tag & "��������,����!", vbInformation + vbOKOnly, gstrSysName
                        If objCtl.Enabled And objCtl.Visible Then objCtl.SetFocus
                        Exit Function
                    End If
            End If
            If Not objCtl Is txt���� Then
                
                If zlStr.ActualLen(Trim(objCtl.Text)) > objCtl.MaxLength Then
                    MsgBox "ע��:" & vbCrLf & "    " & objCtl.Tag & "���������" & objCtl.MaxLength & "���ַ���" & objCtl.MaxLength \ 2 & "������,����!", vbInformation + vbOKOnly, gstrSysName
                    If objCtl.Enabled And objCtl.Visible Then objCtl.SetFocus
                    Exit Function
                End If
                If InStr(1, Trim(objCtl.Text), "'") > 0 Then
                    MsgBox "ע��:" & vbCrLf & "    " & objCtl.Tag & "���зǷ��ַ�(������),����!", vbInformation + vbOKOnly, gstrSysName
                    If objCtl.Enabled And objCtl.Visible Then objCtl.SetFocus
                    Exit Function
                End If
            End If
        Case Else
        End Select
    Next
    If Val(txtParent.Tag) = 0 Then
        MsgBox "ע��:" & vbCrLf & "    δѡ�������Ϣ,����!", vbInformation + vbOKOnly, gstrSysName
        If txtParent.Enabled And txtParent.Visible Then txtParent.SetFocus
        Exit Function
    End If
    With vsWholeSet
        blnHaveDate = False
        For lngRow = 1 To .Rows - 1
            If Val(.Cell(flexcpData, lngRow, .ColIndex("�շ���Ŀ"))) <> 0 Then
                If IsCheckValiedPrice(lngRow, Val(.TextMatrix(lngRow, .ColIndex("ȱʡ�۸�")))) = False Then
                    .Row = lngRow: .Col = .ColIndex("ȱʡ�۸�")
                    Call .ShowCell(.Row, .Col)
                    tbPage.Item(0).Selected = True
                    If vsWholeSet.Enabled And vsWholeSet.Visible Then vsWholeSet.SetFocus
                    Exit Function
                End If
                blnHaveDate = True
            End If
        Next
    End With
    If blnHaveDate = False Then
        MsgBox "ע��:" & vbCrLf & "    δ��������շ���Ŀ�������Ŀ,����!", vbInformation + vbOKOnly, gstrSysName
        tbPage.Item(0).Selected = True
        If vsWholeSet.Enabled And vsWholeSet.Visible Then vsWholeSet.SetFocus
        Exit Function
    End If
    '����Ƿ�������ʹ�ÿ��ҵ�
    If opt��Χ(1).value Then
        If lvw����.ListItems.Count = 0 Then
            MsgBox "ע��:" & vbCrLf & "    δָ��ʹ�ÿ���,����!", vbInformation + vbOKOnly, gstrSysName
            If tbPage.Item(1).Visible = False Then Exit Function
            tbPage.Item(1).Selected = True
            If txt����.Enabled And txt����.Visible Then txt����.SetFocus
            Exit Function
        End If
    End If
    If opt��Χ(0).value Then
        If cbo��Ա.ListIndex < 0 Then
            MsgBox "ע��:" & vbCrLf & "    δָ��ʹ����Ա,����!", vbInformation + vbOKOnly, gstrSysName
            If cbo��Ա.Enabled And cbo��Ա.Visible Then cbo��Ա.SetFocus
            Exit Function
        End If
    End If
        
    '������������Ƿ��ظ�
    If mEditType = EdI_���� Or (mEditType = EdI_�޸� And Trim(txtCode.Text) <> txtCode.Tag) Then
        gstrSQL = "Select 1 From �����շ���Ŀ Where ���� = [1] "
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Trim(txtCode.Text))
        If rsTemp.RecordCount > 0 Then
            MsgBox "ע��:" & vbCrLf & "    ������ʹ�ã�����!", vbInformation + vbOKOnly, gstrSysName
            If txtCode.Enabled And txtCode.Visible Then txtCode.SetFocus
            Exit Function
        End If
    End If
    
    If mEditType = EdI_���� Or (mEditType = EdI_�޸� And Trim(txtName.Text) <> txtName.Tag) Then
        gstrSQL = "Select 1 From �����շ���Ŀ Where ���� = [1] "
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Trim(txtName.Text))
        If rsTemp.RecordCount > 0 Then
            MsgBox "ע��:" & vbCrLf & "    ������ʹ�ã�����!", vbInformation + vbOKOnly, gstrSysName
            If txtName.Enabled And txtName.Visible Then txtName.SetFocus
            Exit Function
        End If
    End If
    
    CheckValied = True
End Function

Private Sub cmdSel_Click()
    'ѡ��ָ����ʹ�ò���
    If SearchUseDept("") = False Then Exit Sub
    If cmdAdd.Enabled And cmdAdd.Visible Then cmdAdd.SetFocus
End Sub

Private Sub cmdSelect_Click()
    'ѡ�����
    If SearchPreLevel("") = False Then Exit Sub
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    If ReadCard = False Then Unload Me: Exit Sub
    If txtName.Enabled And txtName.Visible Then txtName.SetFocus
End Sub

Private Sub Form_Load()
    Call GetPriceGrade(gstrҩƷ�۸�ȼ�, gstr���ļ۸�ȼ�, gstr��ͨ�۸�ȼ�)
    Call InitDefaultLen 'ȡĬ���ֶγ���
    Call zlInitClassPage
    Call InitData
    RestoreWinState Me, App.ProductName
    zl_vsGrid_Para_Restore mlngModule, vsWholeSet, Me.Caption, "������Ŀ��ɱ���", True, True
    mblnFirst = True
End Sub
Private Function InitData() As Boolean
    Dim strSQL As String
    'ִ�в���
    On Error GoTo ErrHandle
    strSQL = _
    "Select Distinct A.ID,A.����,A.����,A.����,B.��������,B.������� " & _
    " From ���ű� A,��������˵�� B " & _
    " Where (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
    " And B.����ID=A.ID and B.������� IN(2,3) " & _
    " Order by B.�������,A.����"
    Set mrsDept = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    cbo��Ա.AddItem "[ѡ����Ա...]"
    cbo��Ա.Tag = cbo��Ա.ListIndex
    InitData = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    With fra����
        .Left = ScaleLeft + 50: .Top = ScaleTop + 100
        .Width = ScaleWidth - 100
        tbPage.Top = .Top + .Height + 100
        tbPage.Left = .Left: tbPage.Width = .Width
        picCmd.Top = ScaleHeight - picCmd.Height
        picCmd.Width = ScaleWidth
    End With
    With tbPage
        .Height = picCmd.Top - .Top
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Err = 0: On Error Resume Next
  SaveWinState Me, App.ProductName
  Call zlDatabase.SetPara("�ϴγ��׷�������", txtParent.Tag, glngSys, mlngModule, True)
  zl_vsGrid_Para_Save mlngModule, vsWholeSet, Me.Caption, "������Ŀ��ɱ���", True, True
End Sub

Private Sub lvw����_GotFocus()
    Call Setʹ�ÿ���Enable
End Sub

Private Sub opt��Χ_Click(Index As Integer)
    mblnChange = True
    Call Set��Χ״̬
End Sub
Private Sub Set��Χ״̬()
    Dim i As Long
    For i = 0 To opt��Χ.UBound
        If opt��Χ(i).value Then
            Exit For
        End If
    Next
    Select Case i
    Case 0  'ָ����Ա
        cbo��Ա.Enabled = True
        cbo��Ա.BackColor = &H80000005
        If Val(tbPage.Selected.Tag) = mItemPage.pg_ʹ�ÿ��� Then
            tbPage.Item(0).Selected = True
        End If
        For i = 0 To tbPage.ItemCount - 1
            If Val(tbPage.Item(i).Tag) = mItemPage.pg_ʹ�ÿ��� Then
                tbPage.Item(i).Visible = False
            End If
        Next
        If cbo��Ա.ListCount <= 0 Then
            cbo��Ա.Clear
            cbo��Ա.AddItem gstrUserName
            cbo��Ա.ItemData(cbo��Ա.NewIndex) = glngUserId
            cbo��Ա.ListIndex = cbo��Ա.NewIndex
            cbo��Ա.AddItem "ѡ��������Ա..."
        End If
    Case 1  'ָ������
        cbo��Ա.Enabled = False
        cbo��Ա.BackColor = &H8000000F
        For i = 0 To tbPage.ItemCount - 1
            If Val(tbPage.Item(i).Tag) = mItemPage.pg_ʹ�ÿ��� Then
                tbPage.Item(i).Visible = True
            End If
        Next
    Case Else
        If Val(tbPage.Selected.Tag) = mItemPage.pg_ʹ�ÿ��� Then
            tbPage.Item(0).Selected = True
        End If
        cbo��Ա.Enabled = False
        cbo��Ա.BackColor = &H8000000F
        For i = 0 To tbPage.ItemCount - 1
            If Val(tbPage.Item(i).Tag) = mItemPage.pg_ʹ�ÿ��� Then
                tbPage.Item(i).Visible = False
            End If
        Next
    End Select
End Sub
Private Sub opt��Χ_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub

Private Sub picCmd_Resize()
    Err = 0: On Error Resume Next
    cmdCancel.Left = picCmd.ScaleWidth - cmdCancel.Width - 100
    cmdOK.Left = cmdCancel.Left - cmdOK.Width - 50
End Sub
Private Sub picUserDept_Resize()
    Err = 0: On Error Resume Next
    With picUserDept
        lvw����.Left = 50
        lvw����.Width = .ScaleWidth - lvw����.Left * 2
        lvw����.Height = .ScaleHeight - .Top - 50
    End With
End Sub

 
Private Sub picWholeSet_Resize()
    Err = 0: On Error Resume Next
    With picWholeSet
        vsWholeSet.Left = 50
        vsWholeSet.Width = .ScaleWidth - vsWholeSet.Left * 2
        vsWholeSet.Top = 50
        vsWholeSet.Height = .ScaleHeight - .Top - 50
    End With
End Sub

 
Private Sub tbPage_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    If Val(Item.Tag) = mItemPage.pg_ʹ�ÿ��� Then
        Call Setʹ�ÿ���Enable
        If txt����.Enabled And txt����.Visible Then txt����.SetFocus
    Else
        If vsWholeSet.Enabled And vsWholeSet.Visible Then vsWholeSet.SetFocus
    End If
End Sub

Private Sub txtCode_Change()
    mblnChange = True
End Sub

Private Sub txtCode_GotFocus()
    zlControl.TxtSelAll txtCode
End Sub

Private Sub txtCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txtCode, KeyAscii, m����ʽ
End Sub

Private Sub txtMemo_Change()
    mblnChange = True
End Sub

Private Sub txtMemo_GotFocus()
    zlControl.TxtSelAll txtMemo
    OS.OpenIme True
End Sub

Private Sub txtMemo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub

Private Sub txtMemo_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txtMemo, KeyAscii, m�ı�ʽ
End Sub

Private Sub txtMemo_LostFocus()
    OS.OpenIme False
End Sub

Private Sub txtName_Change()
    mblnChange = True
    txtSymbol.Text = zlStr.GetCodeByORCL(txtName, False, 20)
    txtWB.Text = zlStr.GetCodeByORCL(txtName, True, 20)
End Sub

Private Sub txtName_GotFocus()
    zlControl.TxtSelAll txtName
    OS.OpenIme True
End Sub

Private Sub txtName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub
Private Sub txtName_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txtName, KeyAscii, m�ı�ʽ
End Sub

Private Sub txtName_LostFocus()
    OS.OpenIme False
End Sub

Private Sub txtParent_Change()
    mblnChange = True
End Sub

Private Sub txtParent_GotFocus()
    zlControl.TxtSelAll txtParent
End Sub

Private Sub txtParent_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub

Private Sub txtParent_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txtParent, KeyAscii, m�ı�ʽ

End Sub

Private Sub txtSymbol_Change()
    mblnChange = True
End Sub

Private Sub txtSymbol_GotFocus()
    zlControl.TxtSelAll txtSymbol
    OS.OpenIme False
End Sub

Private Sub txtSymbol_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub

Private Sub txtSymbol_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txtSymbol, KeyAscii, m�ı�ʽ
End Sub

Private Sub txtWB_Change()
    mblnChange = True
End Sub

Private Sub txtWB_GotFocus()
    zlControl.TxtSelAll txtWB
    OS.OpenIme False
End Sub

Private Sub txtWB_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub

Private Sub txtWB_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txtWB, KeyAscii, m�ı�ʽ
End Sub
Private Sub txt����_Change()
    mblnChange = True: txt����.Tag = ""
End Sub

Private Sub txt����_GotFocus()
    zlControl.TxtSelAll txt����
    Call Setʹ�ÿ���Enable
End Sub
Private Sub txt����_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode <> vbKeyReturn Then Exit Sub
        If Trim(txt����.Tag) <> "" Then Exit Sub
        If SearchUseDept(Trim(txt����.Text)) = False Then Exit Sub
        If cmdAdd.Enabled And cmdAdd.Visible Then cmdAdd.SetFocus
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txt����, KeyAscii, m�ı�ʽ
End Sub
Private Sub ReCale������Ŀ(ByVal lngRow As Long, ByVal dblNum As Double)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���¼��������Ŀ����
    '���:dblNum-����
    '����:
    '����:
    '����:���˺�
    '����:2010-08-31 11:30:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim int�̶����� As Integer, i As Long, dblTemp As Double
    With vsWholeSet
        If Val(.TextMatrix(lngRow, .ColIndex("��������"))) <> 0 Then Exit Sub
        For i = lngRow + 1 To .Rows - 1
             If Val(.TextMatrix(i, .ColIndex("��������"))) = Val(.Cell(flexcpData, lngRow, .ColIndex("���"))) Then
                int�̶����� = Val(.Cell(flexcpData, i, .ColIndex("��������")))
                If int�̶����� = 0 Then '�ǹ��д���
                    dblTemp = IIF(dblNum < 0, -1, 1) * Val(.TextMatrix(i, .ColIndex("��������")))
                    .TextMatrix(i, .ColIndex("ȱʡ����")) = FormatEx(dblTemp, 5)
                ElseIf int�̶����� = 1 Then '�̶��Ĵ���
                    dblTemp = IIF(dblNum < 0, -1, 1) * IIF(Val(.TextMatrix(i, .ColIndex("��������"))) = 0, 1, Val(.TextMatrix(i, .ColIndex("��������"))))
                    .TextMatrix(i, .ColIndex("ȱʡ����")) = FormatEx(dblTemp, 5)
                ElseIf int�̶����� = 2 Then '����������
                    dblTemp = dblNum * Val(.TextMatrix(i, .ColIndex("��������")))
                    .TextMatrix(i, .ColIndex("ȱʡ����")) = FormatEx(dblTemp, 5)
                End If
             End If
        Next
    End With
End Sub
Private Sub vsWholeSet_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������صĸ�ʽ
    '����:���˺�
    '����:2010-08-27 14:12:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    With vsWholeSet
        Select Case Col
        Case .ColIndex("�շ���Ŀ")
             .ColComboList(Col) = "..."
        Case .ColIndex("ȱʡִ�п���")
             .ColComboList(Col) = "..."
        Case .ColIndex("ȱʡ����")
            Call ReCale������Ŀ(Row, Val(.TextMatrix(Row, .Col)))
        Case .ColIndex("ȱʡ����")
        Case .ColIndex("ȱʡ�۸�")
            
        End Select
    End With
End Sub
Private Function IsCheckValiedPrice(ByVal lngRow As Long, ByVal dbl�۸� As Double) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������ļ۸��Ƿ�����Ч��Χ�ڵļ۸�
    '���:dbl�۸�
    '����:
    '����:��Ч,����true,���򷵻�False
    '����:���˺�
    '����:2010-09-16 15:50:13
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng��Ŀid As Long, dbl����޼� As Double, dbl����޼� As Double, dblȱʡ�۸� As Double
    Dim strMsg As String, rsTemp As ADODB.Recordset
    Dim strSQL  As String, strWherePriceGrade As String
    
    On Error GoTo ErrHandle
    With vsWholeSet
        If dbl�۸� = 0 Then 'Ϊ��,ֱ���˳�,�����
            IsCheckValiedPrice = True: Exit Function
        End If
        lng��Ŀid = Val(.Cell(flexcpData, lngRow, .ColIndex("�շ���Ŀ")))
        If lng��Ŀid = 0 Then
            .TextMatrix(lngRow, .ColIndex("����޼�")) = ""
            .TextMatrix(lngRow, .ColIndex("����޼�")) = ""
            IsCheckValiedPrice = False
            Exit Function
        End If
        If InStr(1, "5,6,7", Trim(.TextMatrix(lngRow, .ColIndex("���")))) > 0 Then 'ҩƷ�����
            IsCheckValiedPrice = True
            Exit Function
        End If
        If Trim(.TextMatrix(lngRow, .ColIndex("���"))) = "4" And Val(.TextMatrix(lngRow, .ColIndex("��������"))) = 1 Then '�������Ĳ����
            IsCheckValiedPrice = True
            Exit Function
        End If
        
        dbl����޼� = Val(.TextMatrix(lngRow, .ColIndex("����޼�")))
        dbl����޼� = Val(.TextMatrix(lngRow, .ColIndex("����޼�")))
        If dbl����޼� <> 0 And dbl����޼� <> 0 Then   '�Ѿ�������,��ֱ�ӷ���
            strMsg = CheckScope(dbl����޼�, dbl����޼�, dbl�۸�)
            If strMsg <> "" Then
                MsgBox "ע��:" & vbCrLf & strMsg, vbInformation + vbOKOnly, gstrSysName
                Exit Function
            End If
            IsCheckValiedPrice = True: Exit Function
        End If
    End With
    
    If gstr��ͨ�۸�ȼ� = "" Or Trim(vsWholeSet.TextMatrix(lngRow, vsWholeSet.ColIndex("���"))) = "4" Then
        strWherePriceGrade = " And b.�۸�ȼ� Is Null"
    Else
        strWherePriceGrade = "" & _
        " And (b.�۸�ȼ� = [2]" & vbNewLine & _
        "    Or (b.�۸�ȼ� Is Null" & vbNewLine & _
        "        And Not Exists(Select 1" & vbNewLine & _
        "                       From �շѼ�Ŀ" & vbNewLine & _
        "                       Where b.�շ�ϸĿid = �շ�ϸĿid And �۸�ȼ� = [2]" & vbNewLine & _
        "                             And Sysdate Between ִ������ And Nvl(��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD')))))"
    End If
    strSQL = _
          " Select  B.�ּ�,B.ԭ��,B.ȱʡ�۸� " & _
          " From  �շѼ�Ŀ B" & _
          " Where  Sysdate Between B.ִ������ And Nvl(B.��ֹ����,To_Date('3000-1-1', 'YYYY-MM-DD')) " & _
          "        And B.�շ�ϸĿID=[1]" & strWherePriceGrade
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng��Ŀid, gstr��ͨ�۸�ȼ�)
    If rsTemp.EOF = False Then
        dbl����޼� = Val(NVL(rsTemp!ԭ��))
        dbl����޼� = Val(NVL(rsTemp!�ּ�))
        dblȱʡ�۸� = Val(NVL(rsTemp!ȱʡ�۸�))
        With vsWholeSet
            .TextMatrix(lngRow, .ColIndex("����޼�")) = dbl����޼�
            .TextMatrix(lngRow, .ColIndex("����޼�")) = dbl����޼�
            .Cell(flexcpData, lngRow, .ColIndex("����޼�")) = dblȱʡ�۸�
        End With
        strMsg = CheckScope(dbl����޼�, dbl����޼�, dbl�۸�)
        If strMsg <> "" Then
            MsgBox "ע��:" & vbCrLf & strMsg, vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
    End If
    IsCheckValiedPrice = True: Exit Function
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub vsWholeSet_AfterMoveColumn(ByVal Col As Long, Position As Long)
  zl_vsGrid_Para_Save mlngModule, vsWholeSet, Me.Caption, "������Ŀ��ɱ���", True, True
End Sub

Private Sub vsWholeSet_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If mblnSort = True Then Exit Sub
    Call zl_VsGridRowChange(vsWholeSet, OldRow, NewRow, OldCol, NewCol)
End Sub
Private Function IsHaveHypotaxisItem(ByVal lngRow As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������Ƿ����������Ŀ
    '����:��������true,���򷵻�False
    '����:���˺�
    '����:2010-08-31 12:00:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    With vsWholeSet
        For i = lngRow + 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("��������"))) = lngRow Then
                IsHaveHypotaxisItem = True: Exit Function
            End If
        Next
    End With
End Function

Private Sub vsWholeSet_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    With vsWholeSet
        Select Case Col
        Case .ColIndex("��־")
        Case Else
        End Select
    End With
  zl_vsGrid_Para_Save mlngModule, vsWholeSet, Me.Caption, "������Ŀ��ɱ���", True, True

End Sub

Private Sub vsWholeSet_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim i As Long, arrSplit As Variant
    With vsWholeSet
        If mEditType = EdI_�鿴 Then
            Cancel = True: Exit Sub
        End If
        Select Case Col
        Case .ColIndex("�շ���Ŀ")
            If Val(.TextMatrix(Row, .ColIndex("��������"))) <> 0 Then
                Cancel = True: Exit Sub
            ElseIf IsHaveHypotaxisItem(Row) Then
                Cancel = True
            End If
        Case .ColIndex("ȱʡִ�п���")
            If Not IsEditִ�п��� Then Cancel = True
        Case .ColIndex("ȱʡ����")
        Case .ColIndex("ȱʡ����")
            If InStr(1, ",7", Trim(.TextMatrix(Row, .ColIndex("���")))) = 0 Then Cancel = True: Exit Sub
        Case .ColIndex("ȱʡ�۸�")
            If Val(.TextMatrix(Row, .ColIndex("�Ƿ���"))) = 0 Then Cancel = True: Exit Sub
            'ҩƷ�͸������õ��������ϵļ۸����������,���Բ�������ȱʡ�۸�
            If InStr(1, "5,6,7", Trim(.TextMatrix(Row, .ColIndex("���")))) > 0 Then Cancel = True: Exit Sub
            If Trim(.TextMatrix(Row, .ColIndex("���"))) = "4" And Val(.TextMatrix(Row, .ColIndex("��������"))) = 1 Then Cancel = True: Exit Sub
        Case Else: Cancel = True
        End Select
    End With
End Sub

Private Sub vsWholeSet_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsWholeSet
        Select Case Col
        Case .ColIndex("��־")
            Cancel = True
            
        Case Else
        End Select
    End With

End Sub

Private Sub vsWholeSet_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim lngRow As Long
    If mEditType = EdI_�鿴 Then Exit Sub
    
    With vsWholeSet
        Select Case Col
        Case .ColIndex("�շ���Ŀ")
            If Select�շ���Ŀ("") = False Then Exit Sub
            Call zlVsMoveGridCell(vsWholeSet, .ColIndex("�շ���Ŀ"), , True, lngRow)
        Case .ColIndex("ȱʡִ�п���")
            If ShowSelectDept("") = False Then Exit Sub
            Call zlVsMoveGridCell(vsWholeSet, .ColIndex("�շ���Ŀ"), , True, lngRow)
        End Select
    End With
End Sub

Private Sub vsWholeSet_CellChanged(ByVal Row As Long, ByVal Col As Long)
  mblnChange = True
End Sub
Private Sub SetInputFormat(ByVal lngRow As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����ʽ����
    '����:���˺�
    '����:2010-08-27 14:42:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim arrSplit As Variant
    If mEditType = EdI_�鿴 Then Exit Sub
    With vsWholeSet
        'ColData(i):����������(1-�̶�,-1-����ѡ,0-��ѡ)||������(0-��������,1-��ֹ����,2-��������,�����س���������)
        .ColData(.ColIndex("���")) = "1||1"
        .ColData(.ColIndex("���")) = "1||1"
        .ColData(.ColIndex("��λ")) = "0||1"
        .ColData(.ColIndex("��������")) = "1||1"
        .ColData(.ColIndex("�Ƿ���")) = "1||1"
        .ColData(.ColIndex("���")) = "1||1"
        .ColData(.ColIndex("ִ�п���")) = "1||1"
        .ColData(.ColIndex("��������")) = "1||1"
    End With
End Sub

Private Sub vsWholeSet_EnterCell()
    If mblnSort = True Then Exit Sub
    If mEditType = EdI_�鿴 Then Exit Sub
    
    '�������޸ĲŴ�������
    With vsWholeSet
        SetInputFormat .Row
        OS.OpenIme (False)
        Select Case .Col
        Case .ColIndex("�շ���Ŀ")
             .ColComboList(.Col) = "..."
        Case .ColIndex("ȱʡִ�п���")
             .ColComboList(.Col) = "..."
        End Select
    End With
End Sub

Private Sub vsWholeSet_GotFocus()
  Call zl_VsGridGotFocus(vsWholeSet)
End Sub

Private Sub vsWholeSet_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngCol As Long, blnCancel As Boolean, lngRow As Long
    Dim blnDeleteSubs As Boolean  '�Ƿ�ɾ��������Ŀ
    Dim blnHaveData As Boolean, lng���� As Long
    With vsWholeSet
        If KeyCode <> vbKeyReturn And KeyCode <> vbKeyReturn _
            And (KeyCode <> Asc("*")) And KeyCode <> vbKeySpace _
            And KeyCode <> vbKeyShift Then
            If Shift = 1 And (KeyCode = 56 Or KeyCode <> Asc("*")) Then
                vsWholeSet_CellButtonClick .Row, .Col
            Else
            Select Case .Col
            Case .ColIndex("�շ���Ŀ")
                .ColComboList(.Col) = ""
            Case .ColIndex("ȱʡִ�п���")
                .ColComboList(.Col) = ""
            Case Else
            End Select
            End If
        End If
 
        If KeyCode = vbKeyDelete Then
            blnCancel = False
            'ɾ����ǰ
            Call BeforeDeleteRow(.Row, blnCancel, blnDeleteSubs)
            If blnCancel = True Then Exit Sub
            If .Row = .Rows - 1 And .Row = 1 Then
                For lngCol = 0 To .Cols - 1
                    .TextMatrix(.Row, lngCol) = ""
                    .Cell(flexcpData, .Row, lngCol) = ""
                Next
            Else
                If Val(.TextMatrix(lngRow, .ColIndex("��������"))) = 0 Then
                    Do While True
                        blnHaveData = False
                         For lngRow = .Row + 1 To .Rows - 1
                            If Val(.TextMatrix(lngRow, .ColIndex("��������"))) = Val(.Cell(flexcpData, .Row, .ColIndex("���"))) Then
                                If blnDeleteSubs Then
                                    .RemoveItem lngRow
                                    blnHaveData = True
                                    Exit For
                                Else
                                    .TextMatrix(lngRow, .ColIndex("��������")) = ""
                                    .IsSubtotal(lngRow) = True
                                    .RowOutlineLevel(lngRow) = 1
                                End If
                            End If
                        Next
                        If blnHaveData = False Then Exit Do
                    Loop
                    If .Row = -1 Then .Row = .Rows - 1
                End If
                If .Row = .Rows - 1 And .Row = 1 Then
                    For lngCol = 0 To .Cols - 1
                        .TextMatrix(.Row, lngCol) = ""
                        .Cell(flexcpData, .Row, lngCol) = ""
                    Next
                    .IsSubtotal(.Row) = True
                    .RowOutlineLevel(.Row) = 1
                Else
                    .RemoveItem .Row
                End If
            End If
            'ɾ���к�
            Call AfterDeleteRow
        End If
    End With
    If KeyCode <> vbKeyReturn Then Exit Sub
    With vsWholeSet
        If Trim(.TextMatrix(.Row, .ColIndex("�շ���Ŀ"))) = "" Then
            OS.PressKey vbKeyTab
            Exit Sub
        End If
        Call zlVsMoveGridCell(vsWholeSet, .ColIndex("�շ���Ŀ"), , IIF(mEditType <> EdI_�鿴, True, False), lngRow)
        If lngRow >= 0 Then
            Call AfterAddRow(lngRow)
        End If
    End With
End Sub
Private Sub AfterAddRow(Row As Long)
    '�����к�
    Call RefreshRowNO(Row)
End Sub
Private Sub AfterDeleteRow()
    'ɾ���к�
    Call RefreshRowNO
End Sub
Private Sub RefreshRowNO(Optional lngRow As Long = 1)
    Dim i As Long, j As Long, lng��� As Long
    '���¼������
    With vsWholeSet
        '���������
        For i = lngRow To .Rows - 1
            .TextMatrix(i, .ColIndex("���")) = i
            lng��� = Val(.Cell(flexcpData, i, .ColIndex("���")))
            For j = i + 1 To .Rows - 1
                If Val(.TextMatrix(j, .ColIndex("��������"))) = lng��� And lng��� <> 0 Then
                    .TextMatrix(j, .ColIndex("��������")) = i
                End If
            Next
            .Cell(flexcpData, i, .ColIndex("���")) = i
            If Trim(.TextMatrix(i, .ColIndex("�շ���Ŀ"))) = "" Then
                .IsSubtotal(i) = True
                .RowOutlineLevel(i) = 1
            End If
        Next
    End With
End Sub

Private Sub vsWholeSet_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    '�༭����
    Dim intCol As Integer, strKey As String, lngRow As Long
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    With vsWholeSet
        Select Case Col
        Case .ColIndex("�շ���Ŀ")
            strKey = Trim(.EditText)
            strKey = Replace(strKey, Chr(vbKeyReturn), "")
            strKey = Replace(strKey, Chr(10), "")
            If strKey = "" Then Exit Sub
            If Select�շ���Ŀ(strKey) = False Then
                .TextMatrix(Row, Col) = .EditText: .Cell(flexcpData, Row, Col) = ""
                Exit Sub
            End If
            .EditText = .TextMatrix(Row, Col)
        Case .ColIndex("ȱʡִ�п���")
            strKey = Trim(.EditText)
            strKey = Replace(strKey, Chr(vbKeyReturn), "")
            strKey = Replace(strKey, Chr(10), "")
            If strKey = "" Then Exit Sub
            If ShowSelectDept(strKey) = False Then
                .TextMatrix(Row, Col) = .EditText: .Cell(flexcpData, Row, Col) = ""
                Exit Sub
            End If
            .EditText = .TextMatrix(Row, Col)
        Case Else
        End Select
        Call zlVsMoveGridCell(vsWholeSet, .ColIndex("�շ���Ŀ"), -1, True, lngRow)
        If lngRow >= 0 Then AfterAddRow lngRow
    End With
End Sub

Private Sub vsWholeSet_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub

Private Sub vsWholeSet_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    With vsWholeSet
        Select Case .Col
            Case .ColIndex("�շ���Ŀ")
                Grid.CheckKeyPress vsWholeSet, Row, Col, KeyAscii, m�ı�ʽ
            Case .ColIndex("ȱʡ����"), .ColIndex("ȱʡ�۸�"), .ColIndex("ȱʡ����")
                Grid.CheckKeyPress vsWholeSet, Row, Col, KeyAscii, m���ʽ
        End Select
    End With
End Sub

Private Sub vsWholeSet_LeaveCell()
    If mblnSort Then Exit Sub
    OS.OpenIme False
End Sub

Private Sub vsWholeSet_LostFocus()
    OS.OpenIme False
    Call zl_VsGridLOSTFOCUS(vsWholeSet)
End Sub

Private Sub vsWholeSet_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
        '���õ�Ԫ��ı༭����
        With vsWholeSet
           Select Case .Col
               Case .ColIndex("�շ���Ŀ")
                   .EditMaxLength = 100
               Case .ColIndex("ȱʡ����"), .ColIndex("ȱʡ�۸�")
                   .EditMaxLength = 16
               Case .ColIndex("ȱʡ����")
                  .EditMaxLength = 3
                  ' .EditMask = "-1234567890"
           End Select
    End With
End Sub

Private Sub vsWholeSet_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim strKey As String, intCol As Integer, strTemp As String
    '������֤
    With vsWholeSet
        strKey = Trim(.EditText): strKey = Replace(strKey, Chr(vbKeyReturn), ""): strKey = Replace(strKey, Chr(10), "")
        Select Case Col
            Case .ColIndex("ȱʡ����")
                If zlNumInputCheck(strKey, 16, True, False, 0, .ColKey(Col)) = False Then
                    Cancel = True: Exit Sub
                End If
                If strKey <> "" Then
                    .EditText = FormatEx(Val(strKey), 5)
                End If
            Case .ColIndex("ȱʡ����")
                If zlNumInputCheck(strKey, 3, True, False, 0, .ColKey(Col)) = False Then
                    Cancel = True: Exit Sub
                End If
                If strKey <> "" Then
                    .EditText = IIF(Val(strKey) = 0, 1, Val(strKey))
                End If
            Case .ColIndex("ȱʡ�۸�")
                If zlNumInputCheck(strKey, 16, True, False, 0, .ColKey(Col)) = False Then
                    Cancel = True: Exit Sub
                End If
                If IsCheckValiedPrice(Row, Val(strKey)) = False Then
                     Cancel = True: Exit Sub
                End If
                If strKey <> "" Then
                    .EditText = FormatEx(Val(strKey), 8)
                End If
        End Select
    End With
End Sub
Private Sub BeforeDeleteRow(Row As Long, Cancel As Boolean, blnDeleteSubs As Boolean)
    If mEditType = EdI_�鿴 Then Cancel = True: Exit Sub
    With vsWholeSet
        If Val(.Cell(flexcpData, Row, .ColIndex("�շ���Ŀ"))) <> 0 Then
            If MsgBox("���Ƿ����Ҫɾ���շ���ĿΪ��" & .TextMatrix(.Row, .ColIndex("�շ���Ŀ")) & "���ļ�¼��?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                Cancel = True
                Exit Sub
            End If
            If IsHaveHypotaxisItem(Row) Then
                If MsgBox("�շ���ĿΪ��" & .TextMatrix(.Row, .ColIndex("�շ���Ŀ")) & "���д�����Ŀ,�Ƿ���ͬ������Ŀһ��ɾ��?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                    blnDeleteSubs = True
                End If
            End If
            If Val(.TextMatrix(Row, .ColIndex("��������"))) <> 0 And .IsSubtotal(Row) Then
                'ɾ�����Ǵ�������
                '����һ���Ƿ�������һ��
                If Row >= 2 Then
                    If .TextMatrix(Row - 1, .ColIndex("��������")) = .TextMatrix(Row, .ColIndex("��������")) Then
                        .IsSubtotal(Row - 1) = True
                        .RowOutlineLevel(Row - 1) = 2
                    End If
                End If
            End If
        End If
    End With
End Sub

Private Function Select�շ���Ŀ(Optional strSearch As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�շ���Ŀѡ����
    '���:strSearch-Ҫ����������(""��ʾ������������)
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2010-08-27 14:38:00
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New Recordset, sngLeft As Single, sngTop As Single
    Dim int������Դ  As Integer, str��� As String, lng��Ŀid As Long, str�ų���� As String
    Dim j As Long
    
    With vsWholeSet
        str�ų���� = ""
        If .TextMatrix(1, .ColIndex("���")) <> "7" And .TextMatrix(1, .ColIndex("���")) <> "" Then
            str�ų���� = "'7'"
        End If
        If .TextMatrix(1, .ColIndex("���")) = "7" Then
            str��� = "'7'"
        End If
        
        int������Դ = 2
        If strSearch = "" Then
            lng��Ŀid = frmItemSelect.ShowSelect(Me, int������Դ, True, str���, , , "", zl��ȡ��ҩ��̬(.Row), str�ų����, , gstrҩƷ�۸�ȼ�, gstr���ļ۸�ȼ�, gstr��ͨ�۸�ȼ�)
        Else
            'Call CalcPosition(sngLeft, sngTop, vsWholeSet)
            lng��Ŀid = frmItemSelect.ShowSelect(Me, int������Դ, True, str���, strSearch, .EditWindow, "", zl��ȡ��ҩ��̬(.Row), str�ų����, , gstrҩƷ�۸�ȼ�, gstr���ļ۸�ȼ�, gstr��ͨ�۸�ȼ�)
        End If
        If lng��Ŀid = 0 Then GoTo NotSel:
        
        If CheckItemsIsExsits(lng��Ŀid, .Row) Then GoTo NotSel:
        
        If LoadWholeItem(lng��Ŀid) = False Then
            GoTo NotSel:
        End If
        Select�շ���Ŀ = True
NotSel:
        vsWholeSet_GotFocus
    End With
End Function
Private Function CheckItemsIsExsits(ByVal lng��Ŀid As Long, ByVal lngNotCheckRow As Long, Optional blnSubItem As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����շ�ϸĿ�Ƿ����
    '���:lngNotCheckRow-��������
    '       blnSubItem-�Ƿ��ײ���Ŀ
    '����:
    '����:������Ŀ,����true,���򷵻�False
    '����:���˺�
    '����:2010-08-31 16:00:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngRow As Long
    CheckItemsIsExsits = True
    With vsWholeSet
        For lngRow = 1 To .Rows - 1
            If Val(.Cell(flexcpData, lngRow, .ColIndex("�շ���Ŀ"))) = lng��Ŀid And lngNotCheckRow <> lngRow Then
                If blnSubItem = True Then
                    MsgBox "ע��:" & vbCrLf & "   �ڵ�" & lngRow & "�����Ѿ�����" & vbCrLf & "  ��" & .TextMatrix(lngRow, .ColIndex("�շ���Ŀ")) & " �� " & vbCrLf & "��Ŀ��,���ܼ��ظô�����!", vbInformation + vbOKOnly, gstrSysName
                    Exit Function
                Else
                    MsgBox "�ڵ�" & lngRow & "�����Ѿ����ڸ���Ŀ��,����!", vbInformation + vbOKOnly, gstrSysName
                    Exit Function
                End If
            End If
        Next
    End With
    CheckItemsIsExsits = False
End Function


Public Function zl��ȡ��ҩ��̬(Optional ByVal lngRow As Long = -1, Optional blnOnly�г�ҩ As Boolean = False) As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�����Ƿ�¼�����в�ҩ��
    '���:blnOnly�г�ҩ-���ж��Ƿ����г�ҩ(���䷽ʱ�ж���Ч):ԭ�����г�ҩ���䷽���Ѿ�����,�Ͳ���Ҫ���
    '     lngRow-��ǰ��������
    '����:
    '����:¼�����в�ҩ��,�򷵻���ҩ��̬����(0-ɢװ,1-��Ƭ,2-����),���򷵻�-1 ��ʾ��û��¼����ҩ��̬��Ŀ
    '����:���˺�
    '����:2010-02-02 11:44:17
    '����:27816
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, strTemp As String
    zl��ȡ��ҩ��̬ = -1
    '���δָ��ҳ,���õ�ǰҳ
    strTemp = IIF(blnOnly�г�ҩ, ",6,", ",6,7,")
    With vsWholeSet
        For i = 1 To .Rows - 1
            If InStr(1, strTemp, "," & .TextMatrix(i, .ColIndex("���")) & ",") > 0 And Val(.Cell(flexcpData, i, .ColIndex("�շ���Ŀ"))) <> 0 And i <> lngRow Then
                zl��ȡ��ҩ��̬ = Val(.TextMatrix(i, .ColIndex("��ҩ��̬")))
                Exit Function
            End If
        Next
    End With
End Function
Private Function zlNumInputCheck(ByVal strInput As String, ByVal intMax As Integer, Optional bln������� As Boolean = True, Optional bln���� As Boolean = True, _
        Optional ByVal hwnd As Long = 0, Optional str��Ŀ As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ַ����Ƿ�Ϸ�����������
    '���::strInput        ������ַ���
    '     intMax          ������λ��
    '     bln�������     �Ƿ���и������
    '     bln����         �Ƿ������ļ��
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2010-08-27 15:03:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblValue As Double
    If bln���� = True Then
        If strInput = "" Then
            ShowMsgBox str��Ŀ & "δ���룬����!"
            If hwnd <> 0 Then SetFocusHwnd hwnd
            Exit Function
        End If
    End If
    If strInput = "" Then zlNumInputCheck = True: Exit Function
    
    If IsNumeric(strInput) = False Then
        MsgBox str��Ŀ & "������Ч�����ָ�ʽ��", vbInformation, gstrSysName
        If hwnd <> 0 Then SetFocusHwnd hwnd              '���ý���
        Exit Function
    End If
    
    dblValue = Val(strInput)
    If dblValue >= 10 ^ intMax - 1 Then
        MsgBox str��Ŀ & "��ֵ���󣬲��ܳ���" & 10 ^ intMax - 1 & "��", vbInformation, gstrSysName
        If hwnd <> 0 Then SetFocusHwnd hwnd              '���ý���
        Exit Function
    End If
    If bln������� = True And dblValue < 0 Then
        MsgBox str��Ŀ & "�������븺����", vbInformation, gstrSysName
        If hwnd <> 0 Then SetFocusHwnd hwnd              '���ý���
        Exit Function
    End If
    
    If Abs(dblValue) >= 10 ^ intMax And dblValue < 0 Then
        MsgBox str��Ŀ & "��ֵ��С������С��-" & 10 ^ intMax - 1 & "λ��", vbInformation, gstrSysName
        If hwnd <> 0 Then SetFocusHwnd hwnd              '���ý���
        Exit Function
    End If
    
    
    If bln���� = True And dblValue = 0 Then
        MsgBox str��Ŀ & "���������㡣", vbInformation, gstrSysName
        If hwnd <> 0 Then SetFocusHwnd hwnd              '���ý���
        Exit Function
    End If
    zlNumInputCheck = True
End Function
Private Function SearchPreLevel(ByVal strInput As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ѡ���ϼ�����
    '����:
    '����:���˺�
    '����:2010-08-26 13:39:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strKey As String, strWhere As String
    Dim vRect As RECT, bytStyle As Byte
    Dim blnCancel As Boolean
    
    On Error GoTo ErrHandle
    strKey = gstrLike & strInput & "%"
    If strInput <> "" Then
        If IsNumeric(strInput) Then
            strWhere = " ���� Like [1]"
        ElseIf zlStr.IsCharAlpha(strInput) Then
            strWhere = " ���� Like upper([1])"
        Else
            strWhere = " ���� Like [1] or ���� Like upper([1]) or ���� like [1]"
        End If
        gstrSQL = "" & _
        " Select ID,�ϼ�ID,����,����,����" & _
        " From ������Ŀ����" & _
        " Where " & strWhere
        bytStyle = 0
    Else
        gstrSQL = "" & _
        " Select ID,�ϼ�ID,����,����,����" & _
        " From ������Ŀ����" & _
        " Start with �ϼ�ID is null Connect by prior ID=�ϼ�ID"
        bytStyle = 1
    End If
    
    vRect = zlControl.GetControlRect(txtParent.hwnd)
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, gstrSQL, bytStyle, "�����շ���Ŀ����", False, "", "", False, False, _
        True, vRect.Left, vRect.Top, txtParent.Height, blnCancel, False, True, strKey)
    
    If blnCancel = True Then
        If txtParent.Enabled And txtParent.Visible Then txtParent.SetFocus
        Exit Function
    End If
    If rsTemp Is Nothing Then
        MsgBox "δ�ҵ�ƥ��ķ�����Ϣ,����!", vbInformation + vbDefaultButton1, gstrSysName
        If txtParent.Enabled And txtParent.Visible Then txtParent.SetFocus
        Exit Function
    End If
    If rsTemp.State <> 1 Then
        MsgBox "δ�ҵ�ƥ��ķ�����Ϣ,����!", vbInformation + vbDefaultButton1, gstrSysName
        If txtParent.Enabled And txtParent.Visible Then txtParent.SetFocus
        Exit Function
    End If
    txtParent.Text = NVL(rsTemp!����) & "-" & NVL(rsTemp!����)
    txtParent.Tag = NVL(rsTemp!ID)
    Call zlDefaultCode
    If txtName.Enabled And txtName.Visible Then txtName.SetFocus
    SearchPreLevel = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function SearchUseDept(ByVal strInput As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ѡ��ʹ�ÿ���
    '����:
    '����:���˺�
    '����:2010-08-26 13:39:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strKey As String, strWhere As String
    Dim vRect As RECT, bytStyle As Byte, str��Χ As String
    Dim blnCancel As Boolean
    
    On Error GoTo ErrHandle
    strKey = gstrLike & strInput & "%"
    strWhere = ""
    If strInput <> "" Then
        If IsNumeric(strInput) Then
            strWhere = " And A.���� Like [3]"
        ElseIf zlStr.IsCharAlpha(strInput) Then
            strWhere = "  And A.���� Like upper([3])"
        Else
            strWhere = " And (A.���� Like [3] or A.���� Like upper([3]) or A.���� like [3] )"
        End If
    End If
    
    str��Χ = "1,2,3"
    'str��Χ = IIF(chk��Χ(0).Value = 1, ",1", "") & IIF(chk��Χ(1).Value = 1, ",2", "") & ",3,"
    If InStr(1, mstrPrivs, ";��Ժ���׷���;") > 0 Then
        '����ָ����ȫԺ����
        gstrSQL = "" & _
        "   Select Distinct A.ID,A.����,A.���� " & _
        "   From ���ű� A,��������˵�� B" & _
        "   Where A.ID=B.����ID  " & strWhere & _
        "        And Instr([1],B.�������)>0" & _
        "       And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is Null)" & _
        "       And B.�������� IN('�ٴ�','����','���','����','����','����','Ӫ��')" & _
        " Order by A.����"
    Else
        'ֻ��ָ�����ѵĿ���
        gstrSQL = "" & _
            "   Select Distinct A.ID,A.����,A.����  " & _
            "   From ���ű� A,��������˵�� B,������Ա C" & _
            "   Where A.ID=B.����ID " & strWhere & _
            "       And Instr([1],B.�������)>0 And A.ID=C.����ID And C.��ԱID=[2]" & _
            "       And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is Null)" & _
            "       And B.�������� IN('�ٴ�','����','���','����','����','����','Ӫ��')" & _
            " Order by A.����"
    End If
    
    vRect = zlControl.GetControlRect(txt����.hwnd)
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, gstrSQL, 0, "ʹ�ÿ���ѡ��", False, "", "", False, False, _
        True, vRect.Left, vRect.Top, txt����.Height, blnCancel, False, True, str��Χ, glngUserId, strKey)
    
    If blnCancel = True Then
        If txt����.Enabled And txt����.Visible Then txt����.SetFocus
        Exit Function
    End If
    If rsTemp Is Nothing Then
        MsgBox "δ�ҵ�ƥ���ʹ�ÿ���,����!", vbInformation + vbDefaultButton1, gstrSysName
        If txt����.Enabled And txt����.Visible Then txt����.SetFocus
        Exit Function
    End If
    If rsTemp.State <> 1 Then
        MsgBox "δ�ҵ�ƥ���ʹ�ÿ���,����!", vbInformation + vbDefaultButton1, gstrSysName
        If txt����.Enabled And txt����.Visible Then txt����.SetFocus
        Exit Function
    End If
    txt����.Text = NVL(rsTemp!����) & "-" & NVL(rsTemp!����)
    txt����.Tag = NVL(rsTemp!ID)
    Call Setʹ�ÿ���Enable
    SearchUseDept = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function SaveData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���濨Ƭ������Ϣ
    '����:����ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2010-08-30 15:35:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngRow As Long, cllPro As Collection, lngId As Long
    Dim ObjItem As ListItem, dblTemp As Double
    On Error GoTo ErrHandle
    Set cllPro = New Collection
    If mEditType = EdI_���� Then
        lngId = Sys.NextId("�����շ���Ŀ")
        'Zl_�����շ���Ŀ_Insert
        gstrSQL = "Zl_�����շ���Ŀ_Insert("
    Else
        lngId = Val(mstrID)
        '��Ҫɾ����صĳ�����Ŀ��ɺ�ʹ�ÿ���
        'Zl_�����շ���Ŀ_Update
        gstrSQL = "Zl_�����շ���Ŀ_Update("
    End If
    '  Id_In     In �����շ���Ŀ.ID%Type,
    gstrSQL = gstrSQL & "" & lngId & ","
    '  ����id_In In �����շ���Ŀ.����id%Type,
    gstrSQL = gstrSQL & "" & IIF(Val(txtParent.Tag) = 0, "NULL", Val(txtParent.Tag)) & ","
    '  ����_In   In �����շ���Ŀ.����%Type,
    gstrSQL = gstrSQL & "'" & Trim(txtCode.Text) & "',"
    '  ����_In   In �����շ���Ŀ.����%Type,
    gstrSQL = gstrSQL & "'" & Trim(txtName.Text) & "',"
    '  ��Χ_In   In �����շ���Ŀ.��Χ%Type,
    gstrSQL = gstrSQL & "" & IIF(opt��Χ(0).value, 2, IIF(opt��Χ(1).value, 1, 0)) & ","
    '  ��Աid_In In �����շ���Ŀ.��Աid%Type,
    If opt��Χ(0).value Then
        gstrSQL = gstrSQL & "" & cbo��Ա.ItemData(cbo��Ա.ListIndex) & ","
    Else
        gstrSQL = gstrSQL & "NULL" & ","
    End If
    '  ���_In   In �����շ���Ŀ.���%Type,
    gstrSQL = gstrSQL & "'" & Trim(txtWB.Text) & "',"
    '  ��ע_In   In �����շ���Ŀ.��ע%Type,
    gstrSQL = gstrSQL & "'" & Trim(txtMemo.Text) & "',"
    '  ƴ��_In   In �����շ���Ŀ.ƴ��%Type
    gstrSQL = gstrSQL & "'" & Trim(txtSymbol.Text) & "')"
    zlDatabase.AddItem cllPro, gstrSQL
    '������ɲ���
    With vsWholeSet
        For lngRow = 1 To .Rows - 1
            If Val(.Cell(flexcpData, lngRow, .ColIndex("�շ���Ŀ"))) <> 0 Then
                'Zl_�����շ���Ŀ���_Insert
                gstrSQL = "Zl_�����շ���Ŀ���_Insert("
                '  ����id_In     In �����շ���Ŀ���.����id%Type,
                gstrSQL = gstrSQL & "" & lngId & ","
                '  �շ�ϸĿid_In In �����շ���Ŀ���.�շ�ϸĿid%Type,
                gstrSQL = gstrSQL & "" & Val(.Cell(flexcpData, lngRow, .ColIndex("�շ���Ŀ"))) & ","
                '  ���_In       In �����շ���Ŀ���.���%Type,
                gstrSQL = gstrSQL & "" & Val(.TextMatrix(lngRow, .ColIndex("���"))) & ","
                '  ��������_In   In �����շ���Ŀ���.��������%Type,
                gstrSQL = gstrSQL & "" & IIF(Val(.TextMatrix(lngRow, .ColIndex("��������"))) = 0, "NULL", Val(.TextMatrix(lngRow, .ColIndex("��������")))) & ","
                '����_IN
                dblTemp = Val(.TextMatrix(lngRow, .ColIndex("ȱʡ����")))
                gstrSQL = gstrSQL & "" & dblTemp & ","
                '  ����_In       In �����շ���Ŀ���.����%Type,
                dblTemp = Val(.TextMatrix(lngRow, .ColIndex("ȱʡ����")))
                If .TextMatrix(lngRow, .ColIndex("���")) = "7" Then
                    If Val(.TextMatrix(lngRow, .ColIndex("����ϵ��"))) <> 0 Then
                        dblTemp = Round(dblTemp / Val(.TextMatrix(lngRow, .ColIndex("����ϵ��"))), 5)
                    End If
                End If
                gstrSQL = gstrSQL & "" & dblTemp & ","
                '  ����_In       In �����շ���Ŀ���.����%Type,
                dblTemp = Val(.TextMatrix(lngRow, .ColIndex("ȱʡ�۸�")))
                If .TextMatrix(lngRow, .ColIndex("���")) = "7" Then
                    If Val(.TextMatrix(lngRow, .ColIndex("����ϵ��"))) <> 0 Then
                        dblTemp = Round(dblTemp * Val(.TextMatrix(lngRow, .ColIndex("����ϵ��"))), 8)
                    End If
                End If
                gstrSQL = gstrSQL & "" & dblTemp & ","
                '  ִ�п���id_In In �����շ���Ŀ���.ִ�п���id%Type
                If Val(.Cell(flexcpData, lngRow, .ColIndex("ȱʡִ�п���"))) <> 0 Then
                        gstrSQL = gstrSQL & "" & Val(.Cell(flexcpData, lngRow, .ColIndex("ȱʡִ�п���"))) & ")"
                Else
                        gstrSQL = gstrSQL & "NULL)"
                End If
                zlDatabase.AddItem cllPro, gstrSQL
            End If
        Next
    End With
    '����ʹ�ÿ���
    If opt��Χ(1).value Then
        With lvw����
                For Each ObjItem In lvw����.ListItems
                    'Zl_������Ŀʹ�ÿ���_Insert
                    gstrSQL = "Zl_������Ŀʹ�ÿ���_Insert ("
                    '  ����id_In In ������Ŀʹ�ÿ���.����id%Type,
                    gstrSQL = gstrSQL & "" & lngId & ","
                    '  ����id_In In ������Ŀʹ�ÿ���.����id%Type
                    gstrSQL = gstrSQL & "" & Val(Mid(ObjItem.Key, 2)) & ")"
                     zlDatabase.AddItem cllPro, gstrSQL
                Next
        End With
    End If
    Err = 0: On Error GoTo ErrCommit:
    zlDatabase.ExecuteProcedureBeach cllPro, Me.Caption
    SaveData = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Exit Function
ErrCommit:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Function LoadWholeItem(ByVal lng��Ŀid As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ָ�����շ�ϸĿ
    '���:lng��ĿID-�շ�ϸĿID
    '����:
    '����:
    '����:���˺�
    '����:2010-08-30 17:57:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim strWherePriceGrade As String
    
    If gstr��ͨ�۸�ȼ� = "" And gstrҩƷ�۸�ȼ� = "" And gstr���ļ۸�ȼ� = "" Then
        strWherePriceGrade = " And j.�۸�ȼ� Is Null"
    Else
        strWherePriceGrade = "" & _
            " And ((Instr(';5;6;7;', ';' || k.��� || ';') > 0 And j.�۸�ȼ� = [2])" & vbNewLine & _
            "      Or (Instr(';4;', ';' || k.��� || ';') > 0 And j.�۸�ȼ� = [3])" & vbNewLine & _
            "      Or (Instr(';4;5;6;7;', ';' || k.��� || ';') = 0 And j.�۸�ȼ� = [4])" & vbNewLine & _
            "      Or (j.�۸�ȼ� Is Null" & vbNewLine & _
            "          And Not Exists (Select 1" & vbNewLine & _
            "                          From �շѼ�Ŀ" & vbNewLine & _
            "                          Where j.�շ�ϸĿid = �շ�ϸĿid And Sysdate Between ִ������ And Nvl(��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD'))" & vbNewLine & _
            "                                And ((Instr(';5;6;7;', ';' || k.��� || ';') > 0 And �۸�ȼ� = [2])" & vbNewLine & _
            "                                      Or (Instr(';4;', ';' || k.��� || ';') > 0 And �۸�ȼ� = [3])" & vbNewLine & _
            "                                      Or (Instr(';4;5;6;7;', ';' || k.��� || ';') = 0 And �۸�ȼ� = [4])))))"
    End If
    strSQL = _
    " Select A.ID,A.���,B.���� as �������,A.����,Nvl(E.����,A.����) as ����,E1.���� as ��Ʒ��,A.���,A.���㵥λ," & _
    "       A.���ηѱ�,A.�Ƿ���,A.�Ӱ�Ӽ�,A.ִ�п���,A.��������,A.����ժҪ,A.�������,0 as Ҫ������," & _
    "       Decode(A.���,'4',D.����ID,C.ҩ��ID) as ҩ��ID," & _
    "       Decode(A.���,'4',D.���÷���,C.ҩ������) as ����," & _
    "       Decode(A.���,'4',1,C.סԺ��װ) as סԺ��װ," & _
    "       Decode(A.���,'4',A.���㵥λ,C.סԺ��λ) as סԺ��λ,D.��������,A.¼������,C.��ҩ��̬," & _
    "       M1.���� as ���Ʊ���,M1.���� as ��������,M1.���㵥λ as ������λ,C.����ϵ��," & _
    "       Decode(A.�Ƿ���,1,'ʱ��',LTrim(To_Char(J1.�ּ�,'999999999.9999999'))) as �ּ�" & _
    " From �շ���ĿĿ¼ A, �շ���Ŀ��� B,ҩƷ��� C,�������� D,�շ���Ŀ���� E,�շ���Ŀ���� E1,������ĿĿ¼ M1," & _
    "             (Select j.�շ�ϸĿid, Sum(j.�ּ�) as �ּ�" & vbNewLine & _
    "              From �շѼ�Ŀ J,�շ���ĿĿ¼ K" & vbNewLine & _
    "              Where j.�շ�ϸĿID = k.ID And Sysdate Between J.ִ������ And Nvl(J.��ֹ����, To_Date('3000-1-1', 'YYYY-MM-DD'))" & vbNewLine & _
                         strWherePriceGrade & vbNewLine & _
    "              Group By j.�շ�ϸĿid ) J1 " & _
    " Where A.ID=J1.�շ�ϸĿID(+)  " & _
    "       And A.���=B.���� And A.ID=C.ҩƷID(+) And C.ҩ��ID=M1.id(+) And A.ID=D.����ID(+)" & _
    "       And A.ID=E.�շ�ϸĿID(+) And E.����(+)=1 And E.����(+)=" & IIF(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1) & _
    "       And A.ID=E1.�շ�ϸĿID(+) And E1.����(+)=1 And E1.����(+)=3" & _
    "       And A.ID=[1] "
      
    
 On Error GoTo errH
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng��Ŀid, gstrҩƷ�۸�ȼ�, gstr���ļ۸�ȼ�, gstr��ͨ�۸�ȼ�)
    If rsTemp.EOF Then Exit Function
    If NVL(rsTemp!���) = "7" Then
        '��Ҫ���ҩ���Ƿ����
        If ItemExist(Val(NVL(rsTemp!ҩ��ID)), vsWholeSet.Row) Then
            ShowMsgBox "ע��:" & vbCrLf & "   ��ҩ��Ϊ" & NVL(rsTemp!��������) & " �Ѿ�����,����������!"
            Exit Function
        End If
    End If
    
    With vsWholeSet
        '��ǰ��:
        .TextMatrix(.Row, .ColIndex("���")) = NVL(rsTemp!���)
        .TextMatrix(.Row, .ColIndex("�շ���Ŀ")) = NVL(rsTemp!����) & "-" & NVL(rsTemp!����)
        .EditText = .TextMatrix(.Row, .ColIndex("�շ���Ŀ"))
        .Cell(flexcpData, .Row, .ColIndex("�շ���Ŀ")) = NVL(rsTemp!ID)
        .TextMatrix(.Row, .ColIndex("���")) = NVL(rsTemp!���)
        .TextMatrix(.Row, .ColIndex("��ҩ��̬")) = Val(NVL(rsTemp!��ҩ��̬))
        .TextMatrix(.Row, .ColIndex("����ϵ��")) = Val(NVL(rsTemp!����ϵ��))
        .TextMatrix(.Row, .ColIndex("ҩ��ID")) = NVL(rsTemp!ҩ��ID)
        .TextMatrix(.Row, .ColIndex("��λ")) = NVL(rsTemp!���㵥λ)
        .TextMatrix(.Row, .ColIndex("��������")) = Val(NVL(rsTemp!��������))
        .TextMatrix(.Row, .ColIndex("����޼�")) = ""
        .TextMatrix(.Row, .ColIndex("����޼�")) = ""
        .TextMatrix(.Row, .ColIndex("�ּ�")) = IIF(NVL(rsTemp!�ּ�) = "ʵ��", "ʵ��", FormatEx(Val(NVL(rsTemp!�ּ�)), 5))
        If Val(.TextMatrix(.Row, .ColIndex("ȱʡ����"))) = 0 Then
            .TextMatrix(.Row, .ColIndex("ȱʡ����")) = 1
        End If
        If NVL(rsTemp!���) = "7" Then
            '��ҩ,��ʾ��������
            .TextMatrix(.Row, .ColIndex("ҩ��")) = NVL(rsTemp!���Ʊ���) & "-" & NVL(rsTemp!��������)
            .TextMatrix(.Row, .ColIndex("��λ")) = NVL(rsTemp!������λ)
            .TextMatrix(.Row, .ColIndex("ȱʡ����")) = FormatEx(Val(.TextMatrix(.Row, .ColIndex("ȱʡ����"))) * Val(NVL(rsTemp!����ϵ��)), 5)
            .TextMatrix(.Row, .ColIndex("ȱʡ�۸�")) = FormatEx(Val(.TextMatrix(.Row, .ColIndex("ȱʡ�۸�"))) / Val(NVL(rsTemp!����ϵ��)), 8)
            .TextMatrix(.Row, .ColIndex("��ҩ��̬")) = Val(NVL(rsTemp!��ҩ��̬))
        End If
        .TextMatrix(.Row, .ColIndex("��������")) = ""
        .TextMatrix(.Row, .ColIndex("���")) = .Row
        .Cell(flexcpData, .Row, .ColIndex("���")) = .Row
        .TextMatrix(.Row, .ColIndex("�Ƿ���")) = NVL(rsTemp!�Ƿ���)
        .TextMatrix(.Row, .ColIndex("ִ�п���")) = NVL(rsTemp!ִ�п���)
        .IsSubtotal(.Row) = True
        .RowOutlineLevel(.Row) = 1
        
        If InStr(",5,6,7,", NVL(rsTemp!���)) = 0 Then
            'ҩƷ�������ô�����Ŀ
            If (gbln��������ۿ� And Val(.TextMatrix(.Row, .ColIndex("��������"))) = 0) Or Not gbln��������ۿ� Then  '(����м���,ֻȡһ��)
                If CheckISGetSubItem(.Row) Then
                     '��������
                     Call LoadWholeSubItems(Val(.Cell(flexcpData, .Row, .ColIndex("�շ���Ŀ"))), .Row)
                End If
            End If
        End If
        .Cell(flexcpData, .Row, .ColIndex("���")) = .Row
    End With
    LoadWholeItem = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
    Resume
    End If
    Call SaveErrLog
End Function

Private Function LoadWholeSubItems(ByVal lng��Ŀid As Long, ByVal lng���� As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����ײ���Ŀ
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2010-08-31 10:29:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String, i As Long, lngRow As Long
    Dim strWherePriceGrade As String
    
    If gstr��ͨ�۸�ȼ� = "" And gstrҩƷ�۸�ȼ� = "" And gstr���ļ۸�ȼ� = "" Then
        strWherePriceGrade = " And j.�۸�ȼ� Is Null"
    Else
        strWherePriceGrade = "" & _
            " And ((Instr(';5;6;7;', ';' || k.��� || ';') > 0 And j.�۸�ȼ� = [2])" & vbNewLine & _
            "      Or (Instr(';4;', ';' || k.��� || ';') > 0 And j.�۸�ȼ� = [3])" & vbNewLine & _
            "      Or (Instr(';4;5;6;7;', ';' || k.��� || ';') = 0 And j.�۸�ȼ� = [4])" & vbNewLine & _
            "      Or (j.�۸�ȼ� Is Null" & vbNewLine & _
            "          And Not Exists (Select 1" & vbNewLine & _
            "                          From �շѼ�Ŀ" & vbNewLine & _
            "                          Where j.�շ�ϸĿid = �շ�ϸĿid And Sysdate Between ִ������ And Nvl(��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD'))" & vbNewLine & _
            "                                And ((Instr(';5;6;7;', ';' || k.��� || ';') > 0 And �۸�ȼ� = [2])" & vbNewLine & _
            "                                      Or (Instr(';4;', ';' || k.��� || ';') > 0 And �۸�ȼ� = [3])" & vbNewLine & _
            "                                      Or (Instr(';4;5;6;7;', ';' || k.��� || ';') = 0 And �۸�ȼ� = [4])))))"
    End If
    strSQL = _
    "Select A.ID,Decode(A.���,'4',E.����ID,D.ҩ��ID) as ҩ��ID,A.���,B.���� as �������," & _
    "       A.��������,A.����,Nvl(F.����,A.����) as ����,E1.���� as ��Ʒ��,A.���,A.���㵥λ,A.���ηѱ�,0 as Ҫ������," & _
    "       Decode(A.���,'4',E.���÷���,D.ҩ������) as ����,A.�Ƿ���," & _
    "       Decode(A.���,'4',1,D.סԺ��װ) as סԺ��װ,A.�������," & _
    "       Decode(A.���,'4',A.���㵥λ,D.סԺ��λ) as סԺ��λ," & _
    "       A.�Ӱ�Ӽ�,A.ִ�п���,C.���д���,C.��������,E.��������,D.��ҩ��̬," & _
    "       M1.���� as ���Ʊ���,M1.���� as ��������,M1.���㵥λ as ������λ,D.����ϵ��," & _
    "       Decode(A.�Ƿ���,1,'ʱ��',LTrim(To_Char(J1.�ּ�,'999999999.9999999'))) as �ּ�" & _
    " From �շ���ĿĿ¼ A, �շ���Ŀ��� B,�շѴ�����Ŀ C,ҩƷ��� D,�������� E,�շ���Ŀ���� F,�շ���Ŀ���� E1,������ĿĿ¼ M1," & _
    "             (Select j.�շ�ϸĿid, Sum(j.�ּ�) as �ּ�" & vbNewLine & _
    "              From �շѼ�Ŀ J,�շ���ĿĿ¼ K" & vbNewLine & _
    "              Where j.�շ�ϸĿID = k.ID And Sysdate Between J.ִ������ And Nvl(J.��ֹ����, To_Date('3000-1-1', 'YYYY-MM-DD'))" & vbNewLine & _
                         strWherePriceGrade & vbNewLine & _
    "              Group By j.�շ�ϸĿid ) J1 " & _
    " Where A.ID=J1.�շ�ϸĿID(+)  " & _
    "   And B.����=A.��� And C.����ID=A.ID And A.ID=D.ҩƷID(+) And D.ҩ��ID=M1.id(+)  And A.ID=E.����ID(+)" & _
    "   And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
    "   And A.ID=F.�շ�ϸĿID(+) And F.����(+)=1 And F.����(+)=" & IIF(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1) & _
    "   And A.ID=E1.�շ�ϸĿID(+) And E1.����(+)=1 And E1.����(+)=3" & _
    "   And C.����ID=[1] " & _
    " Order by A.����"

    On Error GoTo errH
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng��Ŀid, gstrҩƷ�۸�ȼ�, gstr���ļ۸�ȼ�, gstr��ͨ�۸�ȼ�)
    With vsWholeSet
        .RowData(lng����) = 1
        Do While Not rsTemp.EOF
            If CheckItemsIsExsits(Val(NVL(rsTemp!ID)), 0, True) = False Then
                .Rows = .Rows + 1
                lngRow = .Rows - 1
                '��ǰ��:
                .TextMatrix(lngRow, .ColIndex("���")) = lngRow
                .Cell(flexcpData, lngRow, .ColIndex("���")) = lngRow
                .TextMatrix(lngRow, .ColIndex("���")) = NVL(rsTemp!���)
                .TextMatrix(lngRow, .ColIndex("��������")) = lng����
                .Cell(flexcpData, lngRow, .ColIndex("��������")) = Val(NVL(rsTemp!���д���))
                .TextMatrix(lngRow, .ColIndex("��������")) = Val(NVL(rsTemp!��������))
                .TextMatrix(lngRow, .ColIndex("�շ���Ŀ")) = NVL(rsTemp!����) & "-" & NVL(rsTemp!����)
                .Cell(flexcpData, lngRow, .ColIndex("�շ���Ŀ")) = NVL(rsTemp!ID)
                .TextMatrix(lngRow, .ColIndex("���")) = NVL(rsTemp!���)
                .TextMatrix(lngRow, .ColIndex("��ҩ��̬")) = Val(NVL(rsTemp!��ҩ��̬))
                .TextMatrix(lngRow, .ColIndex("��λ")) = NVL(rsTemp!���㵥λ)
                .TextMatrix(lngRow, .ColIndex("��������")) = Val(NVL(rsTemp!��������))
                .TextMatrix(lngRow, .ColIndex("����ϵ��")) = Val(NVL(rsTemp!����ϵ��))
                .TextMatrix(lngRow, .ColIndex("ҩ��ID")) = NVL(rsTemp!ҩ��ID)
                .TextMatrix(lngRow, .ColIndex("����޼�")) = ""
                .TextMatrix(lngRow, .ColIndex("����޼�")) = ""
                .TextMatrix(lngRow, .ColIndex("�ּ�")) = IIF(NVL(rsTemp!�ּ�) = "ʵ��", "ʵ��", FormatEx(Val(NVL(rsTemp!�ּ�)), 5))
                If Val(.TextMatrix(lngRow, .ColIndex("ȱʡ����"))) = 0 Then
                    .TextMatrix(lngRow, .ColIndex("ȱʡ����")) = 1
                End If
                If NVL(rsTemp!���) = "7" Then
                    '��ҩ,��ʾ��������
                    .TextMatrix(lngRow, .ColIndex("ҩ��")) = NVL(rsTemp!���Ʊ���) & "-" & NVL(rsTemp!��������)
                    .TextMatrix(lngRow, .ColIndex("��λ")) = NVL(rsTemp!������λ)
                    .TextMatrix(lngRow, .ColIndex("ȱʡ����")) = FormatEx(Val(.TextMatrix(lngRow, .ColIndex("ȱʡ����"))) * Val(NVL(rsTemp!����ϵ��)), 5)
                    .TextMatrix(lngRow, .ColIndex("ȱʡ�۸�")) = FormatEx(Val(.TextMatrix(lngRow, .ColIndex("ȱʡ�۸�"))) / Val(NVL(rsTemp!����ϵ��)), 8)
                End If
                
                .TextMatrix(lngRow, .ColIndex("�Ƿ���")) = NVL(rsTemp!�Ƿ���)
                .TextMatrix(lngRow, .ColIndex("ִ�п���")) = NVL(rsTemp!ִ�п���)
                .RowData(lngRow) = 0
                
                If lng���� <> 0 Then  '���ϼ����зּ�
                      .IsSubtotal(lng����) = True: .RowOutlineLevel(lng����) = 1
                End If
                
                If Val(.RowData(lngRow - 1)) <> 1 Then
                    .IsSubtotal(lngRow - 1) = False
                    .RowOutlineLevel(lngRow - 1) = 2
                End If
                .IsSubtotal(lngRow) = True
                .RowOutlineLevel(lngRow) = 2
            End If
            rsTemp.MoveNext
        Loop
    End With
    LoadWholeSubItems = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Function CheckISGetSubItem(lngRow As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�жϸ����Ƿ�Ӧ��ȡ������Ŀ(�������շ���Ŀ�д�����Ŀ����δȡ��ȡ��)
    '���:
    '����:
    '����:
    '����:���˺�
    '����:2010-08-31 10:41:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, i As Long, strSQL As String
    On Error GoTo ErrHandle
    strSQL = "Select count(����ID) as �������� From �շѴ�����Ŀ Where ����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(vsWholeSet.Cell(flexcpData, lngRow, vsWholeSet.ColIndex("�շ���Ŀ"))))
    If rsTemp.EOF Then
        CheckISGetSubItem = False: Exit Function
    End If
    If Val(NVL(rsTemp!��������)) = 0 Then
        CheckISGetSubItem = False: Exit Function
    End If
    With vsWholeSet
        For i = lngRow + 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("��������"))) = Val(.Cell(flexcpData, lngRow, .ColIndex("���"))) Then
                CheckISGetSubItem = False: Exit Function
            End If
        Next
    End With
    CheckISGetSubItem = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Function ShowSelectDept(ByVal strInput As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ѡ��ָ����ִ�в���
    '���:strInput-����ļ�鴮
    '����:
    '����:
    '����:���˺�
    '����:2010-08-31 16:40:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, i As Long, strKey As String
    Dim str��� As String, lngִ�п��� As Long, strWhere As String, lng��Ŀid As Long
    Dim sngX As Single, sngY As Single, lngH As Long, blnCancel As Boolean
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandle
    With vsWholeSet
        If .Col <> .ColIndex("ȱʡִ�п���") Then Exit Function
        If Val(.Cell(flexcpData, .Row, .ColIndex("�շ���Ŀ"))) = 0 Then Exit Function
        str��� = Trim(.TextMatrix(.Row, .ColIndex("���")))
        If InStr(",4,5,6,7,", "," & str��� & ",") > 0 Then Exit Function
        '0-����ȷ,1-���˿���,2-���˲���,3-����Ա����,4-ָ������,5-Ժ��ִ��(Ԥ��,������δ��),6-�����˿���
        lngִ�п��� = Val(.TextMatrix(.Row, .ColIndex("ִ�п���")))
        If lngִ�п��� <> 0 And lngִ�п��� <> 4 Then Exit Function
        lng��Ŀid = Val(.Cell(flexcpData, .Row, .ColIndex("�շ���Ŀ")))

        strKey = gstrLike & strInput & "%"
        strWhere = ""
        If strInput <> "" Then
            If IsNumeric(strInput) Then
                strWhere = " And A.���� Like [3]"
            ElseIf zlStr.IsCharAlpha(strInput) Then
                strWhere = " And A.���� Like upper([3])"
            Else
                strWhere = " And (A.���� Like [3] or A.���� Like upper([3]) or A.���� like [3] )"
            End If
        End If
            
        '0-����ȷ,1-���˿���,2-���˲���,3-����Ա����,4-ָ������,5-Ժ��ִ��(Ԥ��,������δ��),6-�����˿���
        If lngִ�п��� = 0 Then
            strSQL = _
            "Select Distinct A.ID,A.����,A.����,A.����,B.��������,B.������� " & _
            " From ���ű� A,��������˵�� B " & _
            " Where (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
            "       And B.����ID=A.ID and B.������� IN(2,3) " & strWhere & _
            " Order by B.�������,A.����"
        Else  '4
            strSQL = "" & _
            " Select Distinct A.ID,A.����, A.����" & _
            " From �շ�ִ�п��� B,���ű� A" & _
            " Where B.�շ�ϸĿID=[1] And B.ִ�п���ID=A.id " & strWhere & _
            "       And (b.������Դ is NULL Or b.������Դ=[2]) " & _
            " Order by A.����" '
        End If
    End With
    Call CalcPosition(sngX, sngY, vsWholeSet)
     lngH = vsWholeSet.CellHeight
     sngY = sngY - lngH
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "ִ�в���ѡ��", False, "", "", False, False, _
        True, sngX, sngY, lngH, blnCancel, False, True, lng��Ŀid, 2, strKey)
    If blnCancel Then Exit Function
    If rsTemp Is Nothing Then
        MsgBox "δ�ҵ�ƥ���ִ�п���,����!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    If rsTemp.State <> 1 Then
        MsgBox "δ�ҵ�ƥ���ִ�п���,����!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    With vsWholeSet
        .TextMatrix(.Row, .ColIndex("ȱʡִ�п���")) = rsTemp!���� & "-" & rsTemp!����
        .Cell(flexcpData, .Row, .ColIndex("ȱʡִ�п���")) = NVL(rsTemp!ID)
    End With
    ShowSelectDept = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function IsEditִ�п���() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�Ƿ�����༭ִ�п���
    '����:������true,���򷵻�False
    '����:���˺�
    '����:2010-08-31 17:09:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, i As Long, strKey As String
    Dim str��� As String, lngִ�п��� As Long
    With vsWholeSet
      If .Col <> .ColIndex("ȱʡִ�п���") Then Exit Function
        If Val(.Cell(flexcpData, .Row, .ColIndex("�շ���Ŀ"))) = 0 Then Exit Function
        str��� = Trim(.TextMatrix(.Row, .ColIndex("���")))
        If InStr(",4,5,6,7,", "," & str��� & ",") > 0 Then Exit Function
        '0-����ȷ,1-���˿���,2-���˲���,3-����Ա����,4-ָ������,5-Ժ��ִ��(Ԥ��,������δ��),6-�����˿���
        lngִ�п��� = Val(.TextMatrix(.Row, .ColIndex("ִ�п���")))
        If lngִ�п��� <> 0 And lngִ�п��� <> 4 Then Exit Function
        IsEditִ�п��� = True
    End With
End Function
Private Sub FillBillComboBox(lngRow As Long, lngCol As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����ִ�п��Ҽ��ص�ָ����combox��
    '���:
    '����:
    '����:
    '����:���˺�
    '����:2010-08-31 17:05:23
    '˵��:��δ�øù���,��Ҫ�ǿ��Ҷ��˲������û�ѡ��
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim str��� As String, lngִ�п��� As Long
    
    On Error GoTo ErrHandle
    With vsWholeSet
        If lngCol <> .ColIndex("ȱʡִ�п���") Then Exit Sub
        
        If Val(.Cell(flexcpData, lngRow, .ColIndex("�շ���Ŀ"))) <> 0 Then
              str��� = Trim(.TextMatrix(lngRow, .ColIndex("���")))
              If InStr(",4,5,6,7,", "," & str��� & ",") > 0 Then
                    'ҩƷ,����
                      .ColComboList(.ColIndex("ȱʡִ�п���")) = ""
              Else
                  '0-����ȷ,1-���˿���,2-���˲���,3-����Ա����,4-ָ������,5-Ժ��ִ��(Ԥ��,������δ��),6-�����˿���
                  lngִ�п��� = Val(.TextMatrix(lngRow, .ColIndex("ִ�п���")))
                  Select Case lngִ�п���
                  Case 0
                         .ColComboList(.ColIndex("ȱʡִ�п���")) = .BuildComboList(mrsDept, "����", "ID", vbRed)
                  Case 4
                        strSQL = "Select Distinct b.ID,B.����, B.����" & _
                            " From �շ�ִ�п��� A,���ű� B" & _
                            " Where A.�շ�ϸĿID=[1] And A.ִ�п���ID=b.id" & _
                            "       And (������Դ is NULL Or ������Դ=[2]) " & _
                            " Order by B.����" '
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.Cell(flexcpData, lngRow, .ColIndex("�շ���Ŀ"))), 2)
                         .ColComboList(.ColIndex("ȱʡִ�п���")) = .BuildComboList(rsTmp, "����", "ID", vbRed)
                  Case Else
                      .ColComboList(.ColIndex("ȱʡִ�п���")) = ""
                  End Select
              End If
        Else
            .ColComboList(.ColIndex("ȱʡִ�п���")) = ""
        End If
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function ItemExist(ByVal lng��ҩID As Long, ByVal lngRow As Long) As Boolean
    '���ܣ��ж���ҩ�䷽��������,ָ������ҩ�Ƿ��Ѿ�����
    Dim i As Long, j As Long, lngTemp As Long
    With vsWholeSet
        For i = 1 To .Rows - 1
            If i <> lngRow Then
                If Val(.TextMatrix(i, .ColIndex("ҩ��ID"))) = lng��ҩID Then
                       ItemExist = True
                       Exit Function
                End If
            End If
        Next
    End With
End Function
Private Function CheckScope(varL As Double, varR As Double, varI As Double) As String
'���ܣ��ж��������Ƿ���ԭ�ۺ��ִ��޶��ķ�Χ��
'������varL=ԭ��,varR=�ּ�,varI=������
'���أ�������ڷ�Χ��,��Ϊ��ʾ��Ϣ,����Ϊ�մ�
    If (varL >= 0 And varR >= 0) Or (varL <= 0 And varR <= 0) Then
        '�����ֵ������ͬ,���þ���ֵ�ж�
        If Abs(varI) < Abs(varL) Or Abs(varI) > Abs(varR) Then
            CheckScope = "����ļ۸����ֵ���ڷ�Χ(" & FormatEx(Abs(varL), 5) & "-" & FormatEx(Abs(varR), 5) & ")��."
        End If
    Else
        '������Ų���ͬ,����ԭʼ��Χ�ж�
        If varI < varL Or varI > varR Then
            CheckScope = "����ļ۸�ֵ���ڷ�Χ(" & FormatEx(varL, 5) & "-" & FormatEx(varR, 5) & ")��."
        End If
    End If
End Function
