VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPACSRec 
   Appearance      =   0  'Flat
   BorderStyle     =   0  'None
   ClientHeight    =   7530
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8580
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7530
   ScaleWidth      =   8580
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picBuffer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   6270
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   73
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   315
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame fraFee 
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   300
      Left            =   0
      TabIndex        =   3
      Top             =   5760
      Width           =   7935
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000C&
         BackStyle       =   0  'Transparent
         Caption         =   " ���뱨��"
         ForeColor       =   &H8000000E&
         Height          =   180
         Left            =   120
         TabIndex        =   4
         Top             =   45
         Width           =   810
      End
   End
   Begin VB.Frame fraSplit1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   0
      MousePointer    =   7  'Size N S
      TabIndex        =   2
      Top             =   5640
      Width           =   7110
   End
   Begin VSFlex8Ctl.VSFlexGrid vsBill 
      Height          =   1260
      Left            =   0
      TabIndex        =   5
      Top             =   6120
      Width           =   7260
      _cx             =   12806
      _cy             =   2222
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
      BackColorSel    =   16772055
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   0
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   2
      RowHeightMin    =   250
      RowHeightMax    =   2000
      ColWidthMin     =   0
      ColWidthMax     =   5000
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmPacsRec.frx":0000
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
      AutoSizeMode    =   1
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
      OwnerDraw       =   1
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   -1  'True
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
      Begin MSComctlLib.ImageList imgFlag 
         Left            =   765
         Top             =   840
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   8
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPacsRec.frx":009B
               Key             =   "δ��"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPacsRec.frx":05B5
               Key             =   "����"
            EndProperty
         EndProperty
      End
   End
   Begin VB.CheckBox chkHistory 
      Caption         =   "��ʾ��ʷ����"
      Height          =   300
      Left            =   3240
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
   End
   Begin MSComctlLib.ListView lvwFile 
      Height          =   1695
      Left            =   0
      TabIndex        =   0
      Top             =   300
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   2990
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "iLsTree32"
      SmallIcons      =   "imgFlag"
      ColHdrIcons     =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "����"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "������¼"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "ҽ��"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "״̬"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "���"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3840
      Top             =   3960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   8
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsRec.frx":0ACF
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmPACSRec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const COL_F���� = 0 '��־��
Private Const COL_F���� = 1
Private Const COL_NO = 2 '�ɼ���
Private Const COL_ҽ������ = 3
Private Const COL_���� = 4
Private Const COL_������ = 5
Private Const COL_����ʱ�� = 6
Private Const COL_����ʱ�� = 7
Private Const COL_������ = 8
Private Const COL_����ʱ�� = 9
Private Const COL_ҽ��ID = 10 '������
Private Const COL_������ĿID = 11
Private Const COL_����ID = 12
Private Const COL_��� = 13
Private Const COL_������ = 14
Private Const COL_����ID = 15
Private Const COL_������ = 16
Private Const COL_����ID = 17
Private Const COL_��¼���� = 18
Private Const COL_ǰ��ID = 19
Private Const COL_ִ��״̬ = 20

Private PatientID As Long '����ID
Private PageID As Variant    '��ҳID
Private AreaID As Long, DeptID As Long, OffHosp As Boolean
Private mlngAdviceID As Long, blnShowAll As Boolean
Private aFrmEdit() As Form '�༭��������
Private mblnMoved As Boolean

Public mstrPrivs As String
Public WithEvents mfrmParent As Form
Attribute mfrmParent.VB_VarHelpID = -1
Private Const COLOR_�鵵 As Long = &H8000&
Private Const COLOR_���� As Long = &HFF&
Private Const COLOR_���� As Long = &H80000008
Private Const COLOR_���� As Long = &H808080
'ˢ�´�����ʾ����
Public Sub zlRefresh(ByVal lngPatientID As Long, ByVal varPageID As Variant, _
    Optional ByVal lngAdviceID As Long = 0, Optional ByVal ifShowAll As Boolean = True, _
    Optional ifShowHistory As Boolean = False, Optional ByVal blnMoved As Boolean = False)
'lngPatientID:����ID��0=ûָ�����ˣ�
'strCheckID:�Һŵ�
    PatientID = lngPatientID: PageID = varPageID
    mlngAdviceID = lngAdviceID: blnShowAll = ifShowAll
    chkHistory.Value = IIf(ifShowHistory, 1, 0)
    mblnMoved = blnMoved
    
    ShowFile chkHistory
    LoadReport
End Sub
'ִ�в˵�����
Public Sub zlMenuClick(mnuClick As Menu)
    Dim strMenu As String
    
    '���Ӳ���
    If UCase(mnuClick.Name) = "FILELIST" Then
        AddFile mnuClick
        Exit Sub
    End If
    If UCase(mnuClick.Name) = "REQLIST" Then
        FuncAddRequest CLng(mnuClick.Tag)
        Exit Sub
    End If
    
    If mnuClick.Caption Like "*(&*)*" Then
        strMenu = Split(mnuClick.Caption, "(&")(0)
    Else
        strMenu = mnuClick.Caption
    End If
    Select Case strMenu
        Case "�޸Ĳ���"
            EditFile
        Case "ɾ������"
            DeleteFile
        Case "�޸����뵥"
            FuncWriteRequest
        Case "ɾ�����뵥"
            FuncDeleteRequest
        Case "��ӡ֪ͨ��"
            FuncPrintRequest
        Case "���ı���"
            FuncViewReport
        Case "Ԥ������"
            FuncPrintReport 1
        Case "��ӡ����"
            FuncPrintReport 2
        Case "Ӱ��Ա�"
            ViewImage
    End Select
End Sub
Public Sub zlButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "����"
            Me.PopupMenu mfrmParent.mnuPFileFunc(0)
        Case "ɾ����"
            DeleteFile
        Case "�����޸�"
            EditFile
    End Select
End Sub
Public Sub zlItemRef()

End Sub
Public Sub zlExcel()

End Sub
Public Sub zlPreview()
    If PatientID > 0 And Not lvwFile.SelectedItem Is Nothing Then _
        PreviewPatientFile mfrmParent, IIf(TypeName(PageID) = "String", 1, 2), False, 1, PatientID, PageID, False, 0, 1
End Sub
Public Sub zlPrint()
    Dim btOption As Byte
    Dim bPrtCurrFile As Boolean, bPrtPatiInfo As Boolean, lngBeginY As Long, iBeginPage As Integer
    Dim rsTmp As New ADODB.Recordset, lngFileSeq As Long
    Dim strSQL As String
    If PatientID = 0 Or lvwFile.SelectedItem Is Nothing Then Exit Sub
    
    
    bPrtCurrFile = True: bPrtPatiInfo = False
    lngBeginY = 0: iBeginPage = 1
    btOption = PrintOptionSetup_Patient(mfrmParent, True, bPrtCurrFile, bPrtPatiInfo, lngBeginY, iBeginPage, CLng(Mid(lvwFile.SelectedItem.Key, 4)))
    If btOption = 0 Then Exit Sub
    
    strSQL = "Select Seq From (Select RowNum As Seq,ID From (Select ID From ���˲�����¼ a,�����ļ�Ŀ¼ b where a.�ļ�ID = b.ID " + _
        " And a.����ID=" & PatientID & IIf(TypeName(PageID) = "String", _
        " And a.�Һŵ�='" & PageID & "' And a.��������=1", _
        " And a.��ҳID=" & PageID & " And a.��������=2") & _
        " Order By b.��ҳ desc,a.��д����)) Where ID= [1]"
    If mblnMoved Then
        strSQL = Replace(strSQL, "���˲�����¼", "H���˲�����¼")
    End If
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, Mid(lvwFile.SelectedItem.Key, 4))

    If rsTmp.EOF Then
        lngFileSeq = 1
    Else
        lngFileSeq = rsTmp(0)
    End If
    Select Case btOption
        Case 1
            PreviewPatientFile mfrmParent, IIf(TypeName(PageID) = "String", 1, 2), bPrtCurrFile, lngFileSeq, PatientID, PageID, bPrtPatiInfo, lngBeginY * 56.7, iBeginPage
        Case 2
            PrintPatientFile mfrmParent, IIf(TypeName(PageID) = "String", 1, 2), bPrtCurrFile, lngFileSeq, PatientID, PageID, bPrtPatiInfo, lngBeginY * 56.7, iBeginPage
    End Select
End Sub
Public Sub zlPrintSetup()
    PrintSetup_Patient mfrmParent
End Sub
Private Sub chkHistory_Click()
    On Error Resume Next
    ShowFile chkHistory: Me.lvwFile.SetFocus
End Sub

'���������ļ�
Private Sub AddFile(mnufile As Menu)
    Dim iNum As Integer
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    If PatientID = 0 Then Exit Sub
    On Error Resume Next
    iNum = -1: iNum = UBound(aFrmEdit)
    On Error GoTo EditFileError
    '�ж�ǰ����Ƿ��ظ���д
    strSQL = "Select * From �����ļ�Ŀ¼ Where ID= [1] "
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, Mid(mnufile.Tag, 2))
    
    If rsTmp.EOF Then Exit Sub
    
    If Not IsNull(rsTmp("��д")) Then
        If rsTmp("��д") = 0 Then
            '�����ظ���д�Ĳ���������Ƿ�����д�ò���
            strSQL = "Select Count(*) From ���˲�����¼ Where ����ID= [1] " & _
                IIf(TypeName(PageID) = "String", _
                " And �Һŵ�= [2] ", _
                " And ��ҳID= [2] ") & _
                " And �������� Is Null And �ļ�ID = [3] "
            Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, PatientID, PageID, Mid(mnufile.Tag, 2))

            If rsTmp(0) > 0 Then MsgBox "�ò����Ѿ����ڣ������ظ���д", vbExclamation, gstrSysName: Exit Sub
        Else
        End If
    End If
    
    ReDim Preserve aFrmEdit(iNum + 1)
    EditPatientFile "", CStr(PatientID), CStr(PageID), IIf(TypeName(PageID) = "String", 0, 1), Mid(mnufile.Tag, 2), , Me, aFrmEdit(UBound(aFrmEdit)), , IIf(TypeName(PageID) = "String", 1, 2), , mlngAdviceID
    Exit Sub
EditFileError:
    If ErrCenter() = 1 Then Resume
    SaveErrLog
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    mfrmParent.SetFocus: zlCommFun.PressKey CByte(KeyCode)
End Sub

Private Sub Form_Load()
    lvwFile.ListItems.Add , , "Temp", , 1
    lvwFile.ListItems.Clear
    
    InitBillTable
    
    mfrmParent.mnuFileExcel.Visible = False
    PatientID = -1: PageID = -1
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.fraSplit1.Top > Me.ScaleHeight Then Me.fraSplit1.Top = Me.ScaleHeight - 1590
    
    With lvwFile
        .Left = 0
        .Top = 0
        .Width = Me.ScaleWidth - .Left
        .Height = Me.fraSplit1.Top - .Top
    End With
    With Me.fraSplit1
        .Left = 0: .Width = Me.ScaleWidth - .Left
    End With
    With Me.fraFee
        .Left = 0: .Top = fraSplit1.Top + fraSplit1.Height
        .Width = Me.ScaleWidth - .Left
    End With
    With Me.vsBill
        .Left = 0: .Top = fraFee.Top + fraFee.Height
        .Width = Me.ScaleWidth - .Left: .Height = Me.ScaleHeight - .Top
    End With
End Sub

'��ʾ���˲����ļ���¼
Private Sub ShowFile(ByVal ShowHistory As Boolean)
    Dim rsTmp As New ADODB.Recordset
    Dim tmpItem As MSComctlLib.ListItem
    Dim i As Integer, lngColor As Long
    Dim strSQL As String
    
    lvwFile.ListItems.Clear
    '���û�в������ٴ���
    If PatientID = 0 Then
        ShowMenu
        Exit Sub
    End If
    If ShowHistory Then
        strSQL = "Select ID,��д����,nvl(��������,'δ����') As ��������,nvl(��д��,'δ����') As ��д��,nvl(�Һŵ�,'0') As �Һŵ�," + _
            "Decode(��������,Null,Decode(�鵵����,Null,' ','�鵵'),'����') As ״̬,Decode(��������,-2,'�����¼','��Ժ����') As ���,ҽ��ID From ���˲�����¼ " + _
            "Where ����ID= [1]  And " & IIf(TypeName(PageID) = "String", "��������=1 ", "�������� In (2,-2) ") & _
            IIf(mlngAdviceID = 0 Or blnShowAll, "", "And ҽ��ID= [3]  ") & _
            "And �������� Is Null Order By ��д���� Desc"
    Else
        strSQL = "Select ID,��д����,nvl(��������,'δ����') As ��������,nvl(��д��,'δ����') As ��д��,nvl(�Һŵ�,'0') As �Һŵ�," + _
            "Decode(��������,Null,Decode(�鵵����,Null,' ','�鵵'),'����') As ״̬,Decode(��������,-2,'�����¼','��Ժ����') As ���,ҽ��ID From ���˲�����¼ " + _
            "Where ����ID= [1] And " & IIf(TypeName(PageID) = "String", "�Һŵ�= [2] And ��������=1 ", "��ҳID= [2] And �������� In (2,-2) ") & _
            IIf(mlngAdviceID = 0 Or blnShowAll, "", "And ҽ��ID= [3] ") & _
            "And �������� Is Null Order By ��д���� Desc"
    End If
    If mblnMoved Then
        strSQL = Replace(strSQL, "���˲�����¼", "H���˲�����¼")
    End If
    
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, PatientID, PageID, mlngAdviceID)
    
    Do While Not rsTmp.EOF
        Set tmpItem = lvwFile.ListItems.Add(, "Key" & rsTmp("ID"), rsTmp("��������"))
        
        tmpItem.Tag = Nvl(rsTmp("ҽ��ID"), "0")
        tmpItem.SubItems(1) = IIf(IsNull(rsTmp("��д����")), "Ժ��", rsTmp("��д����"))
        tmpItem.SubItems(2) = rsTmp("��������")
        tmpItem.SubItems(3) = rsTmp("��д��")
        tmpItem.SubItems(4) = rsTmp("״̬")
        tmpItem.SubItems(5) = rsTmp("���")
        With tmpItem.ListSubItems
            Select Case rsTmp("״̬")
                Case "�鵵"
                    lngColor = COLOR_�鵵
                Case "����"
                    lngColor = COLOR_����
                Case Else
                    lngColor = COLOR_����
            End Select
            If lngColor = COLOR_���� And CLng(tmpItem.Tag) <> mlngAdviceID Then lngColor = COLOR_����
            For i = 1 To lvwFile.ColumnHeaders.Count - 1
                .Item(i).ForeColor = lngColor
            Next
        End With
        
        rsTmp.MoveNext
    Loop
    ShowMenu
End Sub

Private Sub ShowMenu()
    Dim blnEnabled As Boolean
    On Error Resume Next
    
    blnEnabled = Not (Me.lvwFile.SelectedItem Is Nothing)
    mfrmParent.mnuPFileFunc(0).Enabled = Not (TypeName(PageID) = "String" Or PatientID = 0)
    mfrmParent.mnuPFileFunc(1).Enabled = blnEnabled
    mfrmParent.mnuPFileFunc(2).Enabled = blnEnabled
    mfrmParent.tbrMain.Buttons("����").Enabled = mfrmParent.mnuPFileFunc(0).Enabled
    mfrmParent.tbrMain.Buttons("�����޸�").Enabled = mfrmParent.mnuPFileFunc(1).Enabled
    mfrmParent.tbrMain.Buttons("ɾ����").Enabled = mfrmParent.mnuPFileFunc(2).Enabled
    
    blnEnabled = Not (Val(vsBill.TextMatrix(vsBill.Row, COL_ҽ��ID)) = 0)
    mfrmParent.mnuReqFunc(0).Enabled = Not (TypeName(PageID) = "String")
    mfrmParent.mnuReqFunc(1).Enabled = blnEnabled
    mfrmParent.mnuReqFunc(2).Enabled = blnEnabled
    mfrmParent.mnuReqFunc(4).Enabled = blnEnabled
    mfrmParent.mnuReqFunc(6).Enabled = blnEnabled
    mfrmParent.mnuReqFunc(7).Enabled = blnEnabled
    mfrmParent.mnuReqFunc(8).Enabled = blnEnabled
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer, iNum As Integer
    
    'ж�ز����༭����
    On Error Resume Next
    iNum = -1: iNum = UBound(aFrmEdit)
    For i = 0 To iNum
        Unload aFrmEdit(0)
    Next
    
    Call SaveWinState(Me, App.ProductName, mfrmParent.Name)
End Sub

Private Sub fraSplit1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 1 Then Exit Sub
    
    fraSplit1.BackColor = RGB(0, 0, 0)
    On Error Resume Next
    If fraSplit1.Top + y < 2000 Then
        fraSplit1.Top = 2000
    ElseIf Me.ScaleHeight - fraSplit1.Top - y < 2000 Then
        fraSplit1.Top = Me.ScaleHeight - 2000
    Else
        fraSplit1.Top = fraSplit1.Top + y
    End If
End Sub

Private Sub fraSplit1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 1 Then Exit Sub

    fraSplit1.BackColor = Me.BackColor
    Form_Resize
End Sub

Private Sub lvwFile_DblClick()
    EditFile
End Sub

Private Sub lvwFile_KeyDown(KeyCode As Integer, Shift As Integer)
    mfrmParent.SetFocus: zlCommFun.PressKey CByte(KeyCode)
End Sub

Private Sub lvwFile_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton And mfrmParent.mnuPFile.Visible And mfrmParent.mnuPFile.Enabled Then Me.PopupMenu mfrmParent.mnuPFile
End Sub
'ɾ�������ļ�
Private Sub DeleteFile()
    Dim iCurrIndex As Integer
    Dim rsTmp As New ADODB.Recordset
    Dim lngEditHWnd As Long
    Dim strSQL As String
    
    If lvwFile.SelectedItem Is Nothing Then Exit Sub
    
    strSQL = "Select * From ���˲�����¼ Where ID= [1] And Not �鵵���� Is Null"
    
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, Mid(lvwFile.SelectedItem.Key, 4))
    
    If Not rsTmp.EOF Then
        MsgBox "�ò����ļ��ѹ鵵������ɾ����", vbInformation, gstrSysName
        Exit Sub
    End If
    If CLng(lvwFile.SelectedItem.Tag) <> mlngAdviceID Then
        MsgBox "ֻ��ɾ�����μ����д�Ĳ����ļ���", vbInformation, gstrSysName
        Exit Sub
    End If
    If mblnMoved Then
        MsgBox "��ǰ���˵Ĳ�����ת�뱸�ݣ�����ִ�б�������", vbInformation, gstrSysName
        Exit Sub
    End If
    
    On Error Resume Next
    lngEditHWnd = GetEditWindow(CLng(Mid(lvwFile.SelectedItem.Key, 4)))
    
    If lngEditHWnd > 0 Then
        MsgBox "�ò������ڱ༭������ɾ��", vbExclamation, gstrSysName
        Call ShowWindow(lngEditHWnd, SW_RESTORE)
        Call BringWindowToTop(lngEditHWnd)
    Else
        If MsgBox("�Ƿ�ɾ���ò����ļ���", vbDefaultButton2 + vbQuestion + vbYesNo, gstrSysName) = vbYes Then
            With lvwFile
                strSQL = "ZL_���˲���_DELETE(" + Mid(.SelectedItem.Key, 4) + ")"
                ExecuteProc strSQL, Me.Caption
'                zlDatabase.ExecuteProcedure "ZL_���˲���_DELETE(" + Mid(.SelectedItem.Key, 4) + ")", ""
                
                iCurrIndex = .SelectedItem.Index
                .ListItems.Remove iCurrIndex
            End With
            
            ShowMenu
        End If
    End If
End Sub
'�޸Ĳ����ļ�
Private Sub EditFile()
    Dim iNum As Integer, lngEditHWnd As Long
    
    If lvwFile.SelectedItem Is Nothing Then Exit Sub
    If Not mfrmParent.mnuPFile.Visible Then Exit Sub
    
    If mblnMoved Then
        MsgBox "��ǰ���˵Ĳ�����ת�뱸�ݣ�����ִ�б�������", vbInformation, gstrSysName
        Exit Sub
    End If
    
    lngEditHWnd = GetEditWindow(CLng(Mid(lvwFile.SelectedItem.Key, 4)))
    If lngEditHWnd = 0 Then
        On Error Resume Next
        iNum = -1: iNum = UBound(aFrmEdit)
        On Error GoTo EditFileError
        ReDim Preserve aFrmEdit(iNum + 1)
        EditPatientFile Mid(lvwFile.SelectedItem.Key, 4), CStr(PatientID), CStr(PageID), IIf(TypeName(PageID) = "String", 0, 1), , , Me, aFrmEdit(UBound(aFrmEdit)), _
             (Len(Trim(lvwFile.SelectedItem.SubItems(4))) = 0) And _
             mfrmParent.mnuPFile.Visible And _
             CLng(lvwFile.SelectedItem.Tag) = mlngAdviceID, IIf(TypeName(PageID) = "String", 1, 2), , mlngAdviceID
    Else
        Call ShowWindow(lngEditHWnd, SW_RESTORE)
        Call BringWindowToTop(lngEditHWnd)
    End If
    Exit Sub
EditFileError:
    If ErrCenter() = 1 Then Resume
    SaveErrLog
End Sub
'�����ļ��鵵
Private Sub CreateFolder()
    Dim rsTmp As New ADODB.Recordset, i As Integer
    Dim lngColor As Long
    Dim strSQL As String
    
    If lvwFile.SelectedItem Is Nothing Then Exit Sub
    strSQL = "Select * From ���˲�����¼ Where ID= [1] And Not �鵵���� Is Null"
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, Mid(lvwFile.SelectedItem.Key, 4))
    
    If Not rsTmp.EOF Then
        MsgBox "�ò����ļ��ѹ鵵��", vbInformation, gstrSysName
        Exit Sub
    End If
    If MsgBox("�����鵵�󽫲���ɾ�����޸ģ�������", vbDefaultButton2 + vbQuestion + vbYesNo, gstrSysName) = vbYes Then
        With lvwFile
            strSQL = "ZL_���˲���_�鵵(" + Mid(.SelectedItem.Key, 4) + ",'" + UserInfo.���� + "')"
            ExecuteProc strSQL, Me.Caption
'            zlDatabase.ExecuteProcedure "ZL_���˲���_�鵵(" + Mid(.SelectedItem.Key, 4) + ",'" + UserInfo.���� + "')", ""
        End With
        
        lvwFile.SelectedItem.SubItems(4) = "�鵵"
        With lvwFile.SelectedItem.ListSubItems
            lngColor = COLOR_�鵵
            For i = 1 To lvwFile.ColumnHeaders.Count - 1
                .Item(i).ForeColor = lngColor
            Next
        End With
        
        ShowMenu
    End If
End Sub
'�����ļ�����
Private Sub UndoFile()
    Dim iCurrIndex As Integer
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    If lvwFile.SelectedItem Is Nothing Then Exit Sub
    
    strSQL = "Select * From ���˲�����¼ Where ID=  [1] And Not �鵵���� Is Null And �������� Is Null"
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, Mid(lvwFile.SelectedItem.Key, 4))

    If rsTmp.EOF Then
        MsgBox "�ò����ļ��������ϣ�", vbInformation, gstrSysName
        Exit Sub
    End If
    If MsgBox("ȷ�Ͻ��÷ݲ���������", vbDefaultButton2 + vbQuestion + vbYesNo, gstrSysName) = vbYes Then
        With lvwFile
            strSQL = "ZL_���˲���_����(" + Mid(.SelectedItem.Key, 4) + ",'" + UserInfo.���� + "')"
            ExecuteProc strSQL, Me.Caption
'            zlDatabase.ExecuteProcedure "ZL_���˲���_����(" + Mid(.SelectedItem.Key, 4) + ",'" + UserInfo.���� + "')", ""
            
            iCurrIndex = .SelectedItem.Index
            .ListItems.Remove iCurrIndex
        End With
        
        ShowMenu
    End If
End Sub

'���ݲ�����¼ID��ȡ��༭���ڵ�hwnd,Ϊ0���ʾ�ò�����ǰδ�༭
Private Function GetEditWindow(ByVal lngFileID As Long) As Long
    Dim i As Integer, iNum As Integer
    
    On Error Resume Next
    iNum = -1: iNum = UBound(aFrmEdit)
    For i = 0 To iNum
        If aFrmEdit(i).Tag = lngFileID Then GetEditWindow = aFrmEdit(i).Hwnd: Exit For
    Next
End Function
'�����༭���ڹرպ�Ĵ����ɲ����༭���á�
Public Sub EditFile_UnLoad(ByVal lngHwnd As Long)
    Dim i As Integer, iNum As Integer, iTmpIndex As Integer
    
    '�ӱ༭����������ɾ����ǰ�رյĴ���
    On Error Resume Next
    iNum = -1: iNum = UBound(aFrmEdit)
    For i = 0 To iNum
        If aFrmEdit(i).Hwnd = lngHwnd Then Exit For
    Next
    iTmpIndex = i
    For i = iTmpIndex + 1 To iNum
        Set aFrmEdit(i - 1) = aFrmEdit(i)
    Next
    Set aFrmEdit(iNum) = Nothing
    If iNum = 0 Then
        Erase aFrmEdit
    Else
        ReDim Preserve aFrmEdit(iNum - 1)
    End If
    
    ShowFile chkHistory
End Sub

Private Sub mfrmParent_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub FuncPrintRequest()
'���ܣ���ӡ֪ͨ��
    Dim strBill As String
    
    If PatientID = 0 Then Exit Sub
    
    With vsBill
        If Val(.TextMatrix(.Row, COL_ҽ��ID)) = 0 Then Exit Sub
        
        '��������������򲻱�
        If Val(.TextMatrix(.Row, COL_������)) = 0 Then
            MsgBox "�õ��ݲ���Ҫ��д���룬���ܴ�ӡ֪ͨ����", vbInformation, gstrSysName
            Exit Sub
        End If
        '���δ��д����������
        If Val(.TextMatrix(.Row, COL_����ID)) = 0 Then
            MsgBox "�õ��ݻ�û����д���룬���ܴ�ӡ֪ͨ����", vbInformation, gstrSysName
            Exit Sub
        End If
        '�������д����������
        If Val(.TextMatrix(.Row, COL_����ID)) <> 0 Then
            MsgBox "�õ����Ѿ���д���棬���ܴ�ӡ֪ͨ����", vbInformation, gstrSysName
            Exit Sub
        End If
        If .TextMatrix(.Row, COL_ǰ��ID) <> mlngAdviceID Then
            MsgBox "ֻ�ܴ�ӡ���μ����д�����룡", vbInformation, gstrSysName
            Exit Sub
        End If
        
        strBill = "ZLCISBILL" & Format(.TextMatrix(.Row, COL_���), "00000") & "-1"
        If ReportPrintSet(gcnOracle, glngSys, strBill, mfrmParent) Then
            Call ReportOpen(gcnOracle, glngSys, strBill, mfrmParent, "NO=" & .TextMatrix(.Row, COL_NO), "����=" & Val(.TextMatrix(.Row, COL_��¼����)), 2)
        End If
    End With
End Sub

Private Sub FuncPrintReport(ByVal PrtMode As Integer)
'���ܣ���ӡ���浥
    Dim strBill As String
    
    If PatientID = 0 Then Exit Sub
    
    With vsBill
        If Val(.TextMatrix(.Row, COL_ҽ��ID)) = 0 Then Exit Sub
        '����ޱ��������򲻱�
        If Val(.TextMatrix(.Row, COL_������)) = 0 Then
            MsgBox "�õ��ݲ���Ҫ��д���棬���ܴ�ӡ���浥��", vbInformation, gstrSysName
            Exit Sub
        End If

        '���δ��д����������
        If Val(.TextMatrix(.Row, COL_����ID)) = 0 Then
            MsgBox "�õ��ݻ�û����д���棬���ܴ�ӡ���浥��", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If Val(.TextMatrix(.Row, COL_ִ��״̬)) <> 1 Then
            MsgBox "�ñ�����δ��ˣ����ܴ�ӡ��", vbInformation, gstrSysName
            Exit Sub
        End If
        
        PrintDiagRpt_New .TextMatrix(.Row, COL_����ID), mfrmParent, PrtMode, picBuffer, mblnMoved
    End With
End Sub

Private Sub FuncAddRequest(ByVal lng����ID As Long)
'���ܣ��������뵥
    If PatientID = 0 Then Exit Sub
    If lng����ID = 0 Then Exit Sub
    
    '���ýӿ�
    Call AddRequest(mfrmParent, PatientID, PageID, lng����ID, False, , , mlngAdviceID)
    If True Then
        Call LoadReport
    End If
End Sub

Private Sub FuncWriteRequest()
'���ܣ���д��������
    If PatientID = 0 Then Exit Sub
    With vsBill
        If Val(.TextMatrix(.Row, COL_ҽ��ID)) = 0 Then Exit Sub
        If Val(.TextMatrix(.Row, COL_������)) = 0 Then
            MsgBox "�õ��ݲ���Ҫ��д���롣", vbInformation, gstrSysName
            Exit Sub
        End If
        If .TextMatrix(.Row, COL_NO) <> "" Then
            MsgBox "��ҽ���Ѿ����ͣ���������д���롣", vbInformation, gstrSysName
            Exit Sub
        End If
'        If .TextMatrix(.Row, COL_ǰ��ID) <> mlngAdviceID Then
'            MsgBox "ֻ���޸ı��μ����д�����룡", vbInformation, gstrSysName
'            Exit Sub
'        End If
        
        '��д����:ҽ��ID,����ID,����ID,ҽ������
        Call EditRequest(Me, Val(.TextMatrix(.Row, COL_ҽ��ID)), Val(.TextMatrix(.Row, COL_����ID)), Val(.TextMatrix(.Row, COL_����ID)), .TextMatrix(.Row, COL_ҽ������), .TextMatrix(.Row, COL_ǰ��ID) <> mlngAdviceID, DataMoved:=mblnMoved)
        If True Then
            Call LoadReport
        End If
    End With
End Sub

Private Sub FuncViewReport()
'���ܣ���д��������
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    If PatientID = 0 Then Exit Sub
    With vsBill
        If Val(.TextMatrix(.Row, COL_����ID)) = 0 Then Exit Sub
        '�ж��Ƿ������ӡ���ı���,���౨����Բ鿴,
        strSQL = "SELECT a.�Ƿ��ӡ, b.������־ FROM Ӱ�����¼ a ,����ҽ����¼ b where a.ҽ��id=b.id and a.ҽ��id = [1]"
        Set rsTmp = OpenSQLRecord(strSQL, "���ı���", CLng(vsBill.TextMatrix(vsBill.Row, COL_ҽ��ID)))
        If rsTmp.EOF Then Exit Sub
        If (Val(.TextMatrix(.Row, COL_ִ��״̬)) <> 1) And (Nvl(rsTmp(0), 0) = 0 Or Nvl(rsTmp(1), 0) = 0) Then
            MsgBox "�ñ�����δ��ˣ����ܲ��ġ�", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '��д����:ҽ��ID,����ID,����ID,ҽ������
        Call EditReport(Me, .TextMatrix(.Row, COL_NO), Val(.TextMatrix(.Row, COL_��¼����)), _
            Val(.TextMatrix(.Row, COL_����ID)), Val(.TextMatrix(.Row, COL_����ID)), .TextMatrix(.Row, COL_ҽ������), True, DataMoved:=mblnMoved)
        If True Then
            Call LoadReport
        End If
    End With
End Sub

Private Sub FuncDeleteRequest()
'���ܣ�ɾ����ǰ���뵥
    Dim strSQL As String, lngRow As Long
        
    If PatientID = 0 Then Exit Sub
    With vsBill
        If Val(.TextMatrix(.Row, COL_����ID)) = 0 Then Exit Sub
        '�������븽��ĵ���
        If Val(.TextMatrix(.Row, COL_������)) = 0 Then
            MsgBox "����[" & .TextMatrix(.Row, COL_����) & "]û����Ҫ��������ݡ�", vbInformation, gstrSysName
            Exit Sub
        End If
        '����д���뵥
        If Val(.TextMatrix(.Row, COL_����ID)) = 0 Then
            MsgBox "����[" & .TextMatrix(.Row, COL_����) & "]û����д���벿�ݵ����ݡ�", vbInformation, gstrSysName
            Exit Sub
        End If
        '�ѷ��ͺ���ɾ��(����ͨ��ҽ������)
        If .TextMatrix(.Row, COL_NO) <> "" Then
            MsgBox "��ҽ���Ѿ����ͣ���Ӧ�����뵥������ɾ����", vbInformation, gstrSysName
            Exit Sub
        End If
        If .TextMatrix(.Row, COL_ǰ��ID) <> mlngAdviceID Then
            MsgBox "ֻ��ɾ�����μ����д�����룡", vbInformation, gstrSysName
            Exit Sub
        End If
        If mblnMoved Then
            MsgBox "��ǰ���˵�������ת�뱸�ݣ�����ִ�б�������", vbInformation, gstrSysName
            Exit Sub
        End If
        If MsgBox("ȷʵҪɾ�����뵥[" & .TextMatrix(.Row, COL_����) & "]��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        
        '��У�ԶԵ��ڹ����м��
        strSQL = "zl_����ҽ����¼_Delete(" & Val(.TextMatrix(.Row, COL_ҽ��ID)) & ",1)"
    End With
    
    'ɾ�����뵥
    On Error GoTo errH
    gcnOracle.BeginTrans
    Call ExecuteProc(strSQL, Me.Caption)
    gcnOracle.CommitTrans
    On Error GoTo 0
        
    '���½���
    With vsBill
        lngRow = .Row
        .RemoveItem .Row
        If .Rows = .FixedRows Then
            .Rows = .FixedRows + 1
        End If
        If lngRow <= .Rows - 1 Then
            .Row = lngRow
        Else
            .Row = .Rows - 1
        End If
        Call .ShowCell(.Row, .Col)
    End With
    Exit Sub
errH:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub InitBillTable()
'���ܣ���ʼ�������嵥��ʽ
    Dim arrHead As Variant, strHead As String, i As Long
    
    strHead = "���ݺ�,810,1;ҽ������,3000,1;����,1800,1;������,850,1;" & _
        "����ʱ��,1080,1;����ʱ��,1080,1;������,850,1;����ʱ��,1080,1;" & _
        "ҽ��ID;������ĿID;����ID;���;������;����ID;������;����ID;��¼����;ǰ��ID;ִ��״̬"
    arrHead = Split(strHead, ";")
    With vsBill
        .Clear
        .FixedRows = 1: .FixedCols = 2
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
        
        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColHidden(.FixedCols + i) = False
                .ColWidth(.FixedCols + i) = Val(Split(arrHead(i), ",")(1))
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
                'Ϊ��֧��zl9PrintMode
                .Cell(flexcpAlignment, .FixedRows, .FixedCols + i, .Rows - 1, .FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
                .ColWidth(.FixedCols + i) = 0 'Ϊ��֧��zl9PrintMode
            End If
        Next
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = 4
        .ColWidth(0) = 11 * Screen.TwipsPerPixelX
        .ColWidth(1) = 11 * Screen.TwipsPerPixelX
    End With
End Sub

Private Function LoadReport() As Boolean
'���ܣ����ݵ�ǰ����ҽ����ȡ������д�����뵥�򱨸浥
    Dim rsBill As New ADODB.Recordset
    Dim strSQL As String, strBill As String
    Dim strKey As String, lngPreRow As Long, i As Long
    Dim strSqlWhere As String
    
    If PatientID = 0 Then
        With vsBill
            .Rows = .FixedRows
            .Rows = .FixedRows + 1
        End With
        Exit Function
    End If
    
    On Error GoTo errH
    
    Screen.MousePointer = 11
    With vsBill
        If Val(.TextMatrix(.Row, COL_ҽ��ID)) <> 0 Then
            strKey = Val(.TextMatrix(.Row, COL_ҽ��ID)) & "_" & .TextMatrix(.Row, COL_NO)
        End If
        .Redraw = flexRDNone
        .Rows = .FixedRows
    End With
    
    '���Ƶ��ݾ�������򱨸渽���ҽ��
    strBill = "Select A.ID as ҽ��ID," & _
        " B.�����ļ�ID as ����ID,D.���,D.����,D.˵��," & _
        " Max(Decode(C.��дʱ��,1,1,0)) as ������," & _
        " Max(Decode(C.��дʱ��,2,1,0)) as ������" & _
        " From ����ҽ����¼ A,���Ƶ���Ӧ�� B,�����ļ���� C,�����ļ�Ŀ¼ D" & _
        " Where A.������ĿID=B.������ĿID And B.Ӧ�ó���=" & IIf(TypeName(PageID) = "String", 1, 2) & _
        " And B.�����ļ�ID=C.�����ļ�ID And B.�����ļ�ID=D.ID" & _
        " And A.����ID= [1] " & IIf(TypeName(PageID) = "String", " And �Һŵ�= [2] ", " And A.��ҳID= [2] ") & _
        " And (A.������� Not IN('5','6','7') And A.���ID is NULL" & _
        "   Or A.�������='C' And A.���ID is Not NULL)" & _
        IIf(mlngAdviceID = 0 Or blnShowAll, "", " And A.ǰ��ID= [3] ") & _
        " Group by A.ID,B.�����ļ�ID,D.���,D.����,D.˵��"
    
    '��ҩƷ��ص�ҽ��
    strSqlWhere = "Select Distinct ���ID From ����ҽ����¼" & _
        " Where ����ID= [1] " & IIf(TypeName(PageID) = "String", " And �Һŵ�= [2] ", " And A.��ҳID= [2] ") & _
        " And ������� IN('5','6','7')"
        
    'ҽ����Ӧ�ĵ����嵥(����������ҽ��,���������͵Ķ���),���ٰ���һ�ֵ��ݸ���
    'δ���͵�ҽ����ʾһ��,�ѷ��͵�һ�η�����ʾһ��(����ֻ���������)
    strSQL = _
        " Select A.ID,A.������ĿID,A.ҽ������,B.����ʱ��,B.NO,B.��¼����," & _
        " A.����ID,B.����ID,C.���,C.����,C.����ID,C.������,C.������," & _
        " X.��д�� as ������,X.��д���� as ����ʱ��," & _
        " Y.��д�� as ������,Y.��д���� as ����ʱ��,Nvl(A.ǰ��ID,0) As ǰ��ID,Nvl(B.ִ��״̬,0) As ִ��״̬,A.�������,B.���ͺ�" & _
        " From ����ҽ����¼ A,����ҽ������ B,(" & strBill & ") C,���˲�����¼ X,���˲�����¼ Y" & _
        " Where A.����ID= [1] " & IIf(TypeName(PageID) = "String", " And A.�Һŵ�= [2] ", " And A.��ҳID= [2] ") & _
        " And (A.������� Not IN('5','6','7') And A.���ID is NULL" & _
        "   Or A.�������='C' And A.���ID is Not NULL)" & _
        " And A.ID Not IN(" & strSqlWhere & ") And A.ҽ��״̬<>4 And Nvl(A.ִ������,0)<>0" & _
        " And A.ID=B.ҽ��ID(+) And A.ID=C.ҽ��ID And (C.������=1 Or C.������=1)" & _
        " And A.����ID=X.ID(+) And B.����ID=Y.ID(+)" & _
        " Order by Nvl(B.����ʱ��,A.����ʱ��) Desc,A.���"
    If mblnMoved Then
        strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
        strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
        strSQL = Replace(strSQL, "���˲�����¼", "H���˲�����¼")
    End If
    
    'ҽ������,NO,����,������,����ʱ��,����ʱ��,������,����ʱ��
    'ҽ��ID;������ĿID;����ID;���;������;����ID;������;����ID;��¼����
    Set rsBill = OpenSQLRecord(strSQL, Me.Caption, PatientID, PageID, mlngAdviceID)
   
    With vsBill
        .Rows = .FixedRows + rsBill.RecordCount
        For i = 1 To rsBill.RecordCount
            .TextMatrix(i, COL_ҽ������) = rsBill!ҽ������
            .TextMatrix(i, COL_NO) = Nvl(rsBill!NO)
            .TextMatrix(i, COL_����) = rsBill!����
            .TextMatrix(i, COL_������) = Nvl(rsBill!������)
            .TextMatrix(i, COL_����ʱ��) = Format(Nvl(rsBill!����ʱ��), "MM-dd HH:mm")
            .TextMatrix(i, COL_����ʱ��) = Format(Nvl(rsBill!����ʱ��), "MM-dd HH:mm")
            .TextMatrix(i, COL_������) = Nvl(rsBill!������)
            .TextMatrix(i, COL_����ʱ��) = Format(Nvl(rsBill!����ʱ��), "MM-dd HH:mm")
            .TextMatrix(i, COL_ҽ��ID) = rsBill!ID
            .TextMatrix(i, COL_������ĿID) = rsBill!������ĿID
            .TextMatrix(i, COL_����ID) = rsBill!����ID
            .TextMatrix(i, COL_���) = rsBill!���
            .TextMatrix(i, COL_������) = Nvl(rsBill!������, 0)
            .TextMatrix(i, COL_����ID) = Nvl(rsBill!����ID, 0)
            .TextMatrix(i, COL_������) = Nvl(rsBill!������, 0)
            .TextMatrix(i, COL_����ID) = Nvl(rsBill!����ID, 0)
            .TextMatrix(i, COL_��¼����) = Nvl(rsBill!��¼����)
            .TextMatrix(i, COL_ǰ��ID) = rsBill!ǰ��ID
            .TextMatrix(i, COL_ִ��״̬) = rsBill!ִ��״̬
            
            .Cell(flexcpData, i, COL_������ĿID) = Nvl(rsBill!�������)
            .Cell(flexcpData, i, COL_����ʱ��) = rsBill!���ͺ�
            
            '�����뱨��ı�ʶ
            If rsBill!������ = 1 Then
                If Not IsNull(rsBill!����ID) Then
                    Set .Cell(flexcpPicture, i, COL_F����) = imgFlag.ListImages("����").Picture
                Else
                    Set .Cell(flexcpPicture, i, COL_F����) = imgFlag.ListImages("δ��").Picture
                End If
            End If
            If rsBill!������ = 1 Then
                If Not IsNull(rsBill!����ID) Then
                    Set .Cell(flexcpPicture, i, COL_F����) = imgFlag.ListImages("����").Picture
                Else
                    Set .Cell(flexcpPicture, i, COL_F����) = imgFlag.ListImages("δ��").Picture
                End If
            End If
            
            '��λ����ǰ��
            If Val(.TextMatrix(i, COL_ҽ��ID)) & "_" & .TextMatrix(i, COL_NO) = strKey Then
                lngPreRow = i
            End If
            
            If .TextMatrix(i, COL_ǰ��ID) <> mlngAdviceID Then
                .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = COLOR_����
            End If
            
            rsBill.MoveNext
        Next
        If .Rows = .FixedRows Then
            .Rows = .FixedRows + 1
        Else
            .AutoSize COL_ҽ������
        End If
        .Cell(flexcpPictureAlignment, .FixedRows, 0, .Rows - 1, .FixedCols - 1) = 4
        
        .Col = COL_NO
        .Row = IIf(lngPreRow <> 0, lngPreRow, .FixedRows)
        Call .ShowCell(.Row, .Col)
        .Redraw = flexRDDirect
    End With
    ShowMenu
    
    Screen.MousePointer = 0
    LoadReport = True
    Exit Function
errH:
    Screen.MousePointer = 0
    vsBill.Redraw = flexRDDirect
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub vsBill_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton And mfrmParent.mnuReq.Visible And mfrmParent.mnuReq.Enabled Then Me.PopupMenu mfrmParent.mnuReq
End Sub

Private Sub vsBill_RowColChange()
    '���Թ�Ƭ�˵�
    On Error Resume Next
    mfrmParent.mnuReqFunc(10).Enabled = (vsBill.Cell(flexcpData, vsBill.Row, COL_������ĿID) = "D" And _
        Val(vsBill.TextMatrix(vsBill.Row, COL_ִ��״̬)) = 1)
End Sub

Private Sub ViewImage()
'���ܣ����ù�Ƭվ
    Dim aFiles() As String
    Dim objPacsCore As Object
    Dim strFTPHost As String, strDicomPath As String, strLocalPath As String
    Dim strFTPUser As String, strFtpPwd As String
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim strCachePath As String
    Dim objFileSystem As New Scripting.FileSystemObject
    Dim strCheckUID As String
    
    On Error GoTo DBError
    strSQL = "Select ���UID From Ӱ�����¼ Where ҽ��ID=[1]"
    Set rsTmp = OpenSQLRecord(strSQL, "��Ƭ����", CLng(vsBill.TextMatrix(vsBill.Row, COL_ҽ��ID)))
    If rsTmp.EOF Then Exit Sub
    
    strCachePath = App.Path & "\TmpImage\"
    If Not objFileSystem.FolderExists(strCachePath) Then objFileSystem.CreateFolder strCachePath
    
    strCheckUID = Nvl(rsTmp(0))
    aFiles = GetAllImageFiles(strCheckUID, , mblnMoved, strFTPHost, strDicomPath, _
        strLocalPath, strFTPUser, strFtpPwd)
    If UBound(aFiles) > 0 Then
        Set objPacsCore = CreateObject("zl9PacsCore.clsViewer")
        objPacsCore.CallOpenViewerCache aFiles, mfrmParent, strCachePath & strLocalPath, strFTPHost & strDicomPath, mstrPrivs, strCheckUID, strFTPHost, strDicomPath, gcnOracle, strFTPUser, strFtpPwd, True
        Set objPacsCore = Nothing
    End If
    Exit Sub
DBError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
