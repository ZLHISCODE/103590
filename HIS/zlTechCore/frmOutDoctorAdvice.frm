VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOutDoctorAdvice 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   6720
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9195
   Icon            =   "frmOutDoctorAdvice.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   6720
   ScaleWidth      =   9195
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraAdviceUD 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   60
      MousePointer    =   7  'Size N S
      TabIndex        =   3
      Top             =   4455
      Width           =   7275
   End
   Begin VSFlex8Ctl.VSFlexGrid vsAppend 
      Height          =   1380
      Left            =   60
      TabIndex        =   2
      Top             =   4815
      Width           =   7275
      _cx             =   12832
      _cy             =   2434
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
      Rows            =   2
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   2
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
   Begin MSComctlLib.TabStrip tabAppend 
      Height          =   300
      Left            =   60
      TabIndex        =   1
      Top             =   4500
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   529
      Style           =   2
      HotTracking     =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "ҽ���Ƽ���Ŀ(&P)"
            Key             =   "ҽ���Ƽ���Ŀ"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "ҽ��������ϸ(&S)"
            Key             =   "ҽ��������ϸ"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "ҽ��ǩ����¼(&G)"
            Key             =   "ҽ��ǩ����¼"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vsAdvice 
      Height          =   4380
      Left            =   60
      TabIndex        =   0
      Top             =   75
      Width           =   7260
      _cx             =   12806
      _cy             =   7726
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
      FormatString    =   $"frmOutDoctorAdvice.frx":000C
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
      AllowUserFreezing=   1
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
      Begin MSComctlLib.ImageList imgSign 
         Left            =   3285
         Top             =   1170
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   16777215
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOutDoctorAdvice.frx":00A7
               Key             =   "ǩ��"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList imgFlag 
         Left            =   1260
         Top             =   1140
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   8
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   4
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOutDoctorAdvice.frx":03F9
               Key             =   "����"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOutDoctorAdvice.frx":0613
               Key             =   "��¼"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOutDoctorAdvice.frx":0B2D
               Key             =   "δ����"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOutDoctorAdvice.frx":1047
               Key             =   "������"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList imgPass 
         Left            =   2265
         Top             =   1155
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   14
         ImageHeight     =   14
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   5
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOutDoctorAdvice.frx":1561
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOutDoctorAdvice.frx":185B
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOutDoctorAdvice.frx":1B55
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOutDoctorAdvice.frx":1E4F
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOutDoctorAdvice.frx":2149
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox picFocus 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H008080FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   30
      ScaleHeight     =   600
      ScaleWidth      =   630
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   30
      Width           =   630
   End
End
Attribute VB_Name = "frmOutDoctorAdvice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
Public mfrmParent As Object
Public mstrPrivs As String
Private WithEvents mfrmEdit As Form
Attribute mfrmEdit.VB_VarHelpID = -1

'�ϴ�ˢ������ʱ�Ĳ�����Ϣ
Private mlng����ID As Long
Private mstr�Һŵ� As String
Private mint״̬ As Integer '0-���ﲡ��,1-���ﲡ��,2-���ﲡ��
Private mlngǰ��ID As Long
Private mblnShowAll As Boolean

Private mbln�Զ�Ƥ�� As Boolean

Private mblnMoved As Boolean '��ǰ�Һŵ��Ƿ��Ѿ�ת��
Private mvRegDate As Date '�Һŵ��ĹҺ�ʱ��

'ҽ���˵�����
Private Enum Menu_Advice
    mnu�¿�ҽ�� = 0
    mnu�޸�ҽ�� = 1
    mnuɾ��ҽ�� = 2
    mnuƤ�Խ�� = 4 '-
    mnu����ҽ�� = 6 '-
    mnu����ҽ�� = 7
    mnu���Ƶ��ı� = 9 '-
End Enum

'����˵�������
Private Enum Menu_Report
    mnu��ӡ���Ƶ��� = 0
End Enum

Private Enum COLҽ���嵥
    '�̶���
    COL_F��־ = 0
    COL_F���� = 1
    '������
    COL_ID = 2
    COL_���ID = COL_ID + 1
    COL_��ID = COL_ID + 2
    COL_��� = COL_ID + 3
    COL_Ӥ��ID = COL_ID + 4
    COL_ҽ��״̬ = COL_ID + 5
    COL_������� = COL_ID + 6
    COL_�������� = COL_ID + 7
    COL_������� = COL_ID + 8
    COL_��־ = COL_ID + 9
    '�ɼ���
    COL_��ʾ = COL_ID + 10 'Pass
    COL_��ʼʱ�� = COL_ID + 11
    COL_ҽ������ = COL_ID + 12
    COL_Ƥ�� = COL_ID + 13
    COL_���� = COL_ID + 14
    COL_���� = COL_ID + 15
    COL_Ƶ�� = COL_ID + 16
    COL_�÷� = COL_ID + 17
    COL_ҽ������ = COL_ID + 18
    COL_ִ��ʱ�� = COL_ID + 19
    COL_ִ�п��� = COL_ID + 20
    COL_ִ������ = COL_ID + 21
    COL_����ҽ�� = COL_ID + 22
    COL_����ʱ�� = COL_ID + 23
    COL_������ = COL_ID + 24
    COL_����ʱ�� = COL_ID + 25
    '������
    COL_����ID = COL_ID + 26 '��Ӧ�����ļ�Ŀ¼.ID
    COL_������ = COL_ID + 27 '���Ƶ����Ƿ���������
    COL_������ = COL_ID + 28 '���Ƶ����Ƿ��б�����
    COL_����ID = COL_ID + 29 '��Ӧ���˲�����¼.ID
    COL_ǰ��ID = COL_ID + 30
    COL_ǩ���� = COL_ID + 31
End Enum

Private Enum COL�����嵥
    cs���ͺ� = 0
    cs����ʱ�� = 1
    cs����ҽ�� = 2
    cs���ݺ� = 3
    cs�շ���Ŀ = 4
    cs���� = 5
    cs�Ʒ�״̬ = 6
    csִ��״̬ = 7
    csִ�п��� = 8
    cs������ = 9
    cs��¼���� = 10
End Enum

Public Function zlRefresh(lng����ID As Long, str�Һŵ� As String, int״̬ As Integer, varValue As Variant, Optional ByVal lngǰ��ID As Long = 0, Optional ByVal ifShowAll As Boolean = True) As Boolean
'���ܣ�ˢ�»����ҽ���嵥
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    mlng����ID = lng����ID
    mstr�Һŵ� = str�Һŵ�
    mint״̬ = int״̬
    mlngǰ��ID = lngǰ��ID
    mblnShowAll = ifShowAll
        
    '�Һŵ��Ƿ�ת��,���Һ�ʱ��
    mblnMoved = False
    If lng����ID <> 0 Then
        If mint״̬ = 2 Then '����ҵ�����,ֻ�ж����ﲡ��
            mblnMoved = MovedByNO(str�Һŵ�, "���˹Һż�¼")
        End If
        strSQL = "Select �Ǽ�ʱ�� From ���˹Һż�¼ Where NO=[1]"
        If mblnMoved Then strSQL = Replace(strSQL, "���˹Һż�¼", "H���˹Һż�¼")
        On Error GoTo errH
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, str�Һŵ�)
        If Not rsTmp.EOF Then
            mvRegDate = rsTmp!�Ǽ�ʱ��
        Else
            mvRegDate = zlDatabase.Currentdate
        End If
        On Error GoTo 0
    End If
    
    If mlng����ID = 0 Then
        '���ҽ���嵥
        Call ClearAdviceData
        Call ClearAppendData
        mfrmParent.stbThis.Panels(2).Text = ""
    Else
        '��ʾҽ���嵥
        Call LoadAdvice
        Call ShowTotalMoney
    End If
    zlRefresh = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function zlButtonClick(objButton As Button) As Boolean
'���ܣ�ִ��ҽ����ť����
    Select Case objButton.Key
        Case "�¿�"
            Call FuncAdviceAdd
        Case "�޸�"
            Call FuncAdviceModi
        Case "ɾ��"
            Call FuncAdviceDel
        Case "����"
            Call FuncAdviceSend
        Case "����"
            Call FuncAdviceRevoke
        Case "ǩ��"
            Call FuncAdviceSign
    End Select
End Function

Public Function zlMenuClick(objMenu As Menu) As Boolean
'���ܣ�ִ��ҽ���˵�����
    Dim strText As String
    
    If objMenu.Caption Like "*(&*)*" Then
        strText = Split(objMenu.Caption, "(")(0)
    Else
        strText = objMenu.Caption
    End If
        
    If objMenu.Name = "mnuReportClinic" Then
        '��ӡ���Ƶ���
        Call FuncBillPrint(objMenu)
    ElseIf objMenu.Name = "mnuViewAdviceAppend" Then
        '��ʾ/���ظ��ӱ��
        objMenu.Checked = Not objMenu.Checked
        fraAdviceUD.Visible = objMenu.Checked
        tabAppend.Visible = objMenu.Checked
        vsAppend.Visible = objMenu.Checked
        Call Form_Resize
        
        Call vsAdvice_AfterRowColChange(-1, -1, vsAdvice.Row, vsAdvice.Col)
    Else
        Select Case strText
            Case "�¿�ҽ��"
                Call FuncAdviceAdd
            Case "�޸�ҽ��"
                Call FuncAdviceModi
            Case "ɾ��ҽ��"
                Call FuncAdviceDel
            Case "Ƥ�Խ��"
                Call FuncAdviceTest
            Case "����ҽ��"
                Call FuncAdviceSend
            Case "����ҽ��"
                Call FuncAdviceRevoke
            Case "���Ƶ��ı�"
                Call FuncCopyToText
            Case "����ǩ��"
                Call FuncAdviceSign
            Case "ȡ��ǩ��"
                Call FuncAdviceSignErase
            Case "��֤ǩ��"
                Call FuncAdviceSignVerify
        End Select
    End If
End Function

Private Sub FuncCopyToText()
    Dim strCopy As String, intRow As Integer
    strCopy = ""
    With vsAdvice
        For intRow = .FixedRows To .Rows - 1
            If .TextMatrix(intRow, COL_�������) = "5" Or .TextMatrix(intRow, COL_�������) = "6" Then
                strCopy = strCopy & .TextMatrix(intRow, COL_ҽ������) _
                        & " " & .TextMatrix(intRow, COL_����) _
                        & " " & .TextMatrix(intRow, COL_Ƶ��) _
                        & " " & .TextMatrix(intRow, COL_�÷�) _
                        & vbCrLf
            Else
                strCopy = strCopy & .TextMatrix(intRow, COL_ҽ������) & vbCrLf
            End If
        Next
    End With
    If strCopy <> "" Then
        VB.Clipboard.Clear
        VB.Clipboard.SetText strCopy
    End If
End Sub

Public Sub zlItemRef()
'���ܣ��������Ʋο�
    Dim lng������ĿID As Long, i As Long

    With vsAdvice
        If Val(.TextMatrix(.Row, COL_ID)) <> 0 Then
            If .TextMatrix(.Row, COL_�������) = "E" And (RowIs�䷽��(.Row) Or RowIs������(.Row)) Then
                lng������ĿID = Get������ĿID(Val(.TextMatrix(.Row, COL_ID)), True)
            Else
                lng������ĿID = Get������ĿID(Val(.TextMatrix(.Row, COL_ID)), False)
            End If
        End If
    End With
    Call ShowClinicHelp(0, mfrmParent, lng������ĿID)
End Sub

Public Sub zlPrintSetup()
    Call zlPrintSet
End Sub

Public Sub zlExcel()
    Call OutputList(3)
End Sub

Public Sub zlPreview()
    Call OutputList(2)
End Sub

Public Sub zlPrint()
    Call OutputList(1)
End Sub

Private Sub Form_Activate()
    picFocus.SetFocus '�������ú󱾴����ڵĽ���˳�����Ч
    vsAdvice.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim objMenu As Object
    'Ϊ���ⲿϵͳ�������ӣ�By����ͮ��
    On Error Resume Next
    
    If Shift = vbCtrlMask And KeyCode = vbKeyA Then
        Set objMenu = mfrmParent.mnuAdviceFunc(mnu�¿�ҽ��)
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyM Then
        Set objMenu = mfrmParent.mnuAdviceFunc(mnu�޸�ҽ��)
    ElseIf KeyCode = vbKeyDelete Then
        Set objMenu = mfrmParent.mnuAdviceFunc(mnuɾ��ҽ��)
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyT Then
        Set objMenu = mfrmParent.mnuAdviceFunc(mnuƤ�Խ��)
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyS Then
        Set objMenu = mfrmParent.mnuAdviceFunc(mnu����ҽ��)
    ElseIf KeyCode = vbKeyF2 Then '�����涨λ����
        Call mfrmParent.Form_KeyDown(vbKeyF2, 0): Exit Sub
    ElseIf KeyCode = vbKeyF3 Then
        Set objMenu = mfrmParent.mnuAdviceFunc(mnu����ҽ��)
    ElseIf KeyCode = vbKeyF6 Then
        Call zlItemRef
    End If
    If Not objMenu Is Nothing Then
        If objMenu.Enabled And objMenu.Visible Then
            Call zlMenuClick(objMenu)
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    Call mfrmParent.Form_KeyPress(KeyAscii)
End Sub

Private Sub fraAdviceUD_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        If vsAdvice.Height + y < 1000 Or vsAppend.Height - y < 60 Then Exit Sub
        fraAdviceUD.Top = fraAdviceUD.Top + y
        tabAppend.Top = tabAppend.Top + y
        vsAdvice.Height = vsAdvice.Height + y
        vsAppend.Top = vsAppend.Top + y
        vsAppend.Height = vsAppend.Height - y
        Me.Refresh
    End If
End Sub

Private Sub mfrmEdit_Unload(Cancel As Integer)
    If Not Cancel Then
        If frmOutAdviceEdit.mblnOK Then
            Call LoadAdvice
            Call ShowTotalMoney
        End If
        Set mfrmEdit = Nothing
        
        If mfrmParent.TabFile.SelectedItem.Key = "ҽ��" Then
            Call BringWindowToTop(Me.Hwnd)
        End If
    End If
End Sub

Private Function CheckWindow() As Boolean
'���ܣ����ҽ���༭�����Ƿ��Ѿ���
    If Not mfrmEdit Is Nothing Then
        '��ǰ���ڴ���
        MsgBox "ҽ���༭�����Ѿ��򿪣�������ɵ�ǰ��������ִ�С�", vbInformation, gstrSysName
        '��λ����ǰ�Ĵ���
        If mfrmEdit.WindowState = vbMinimized Then mfrmEdit.WindowState = vbNormal
        mfrmEdit.SetFocus
        Exit Function
    Else
        '�������ڴ���
        If Not CheckAdviceWindow("����ҽ���༭") Then Exit Function
    End If
    CheckWindow = True
End Function

Private Sub FuncBillPrint(objMenu As Menu)
'���ܣ���ӡ���Ƶ���
    If objMenu.Tag = "" Then Exit Sub
    If ReportPrintSet(gcnOracle, glngSys, objMenu.Tag, mfrmParent) Then
        With vsAppend
            Call ReportOpen(gcnOracle, glngSys, objMenu.Tag, mfrmParent, "NO=" & .TextMatrix(.Row, cs���ݺ�), "����=" & Val(.TextMatrix(.Row, cs��¼����)), 2)
        End With
    End If
End Sub

Private Sub FuncAdviceSign()
'���ܣ���ҽ�����е���ǩ��
    Dim strSQL As String, strIDs As String, i As Long
    Dim strSource As String, strSign As String
    Dim lngǩ��ID As Long, lng֤��ID As Long
    Dim intRule As Integer
    
    If mint״̬ <> 1 Then Exit Sub '���ﲡ��
    If gobjESign Is Nothing Then Exit Sub
    
    '��ȡǩ��ҽ��Դ��
    intRule = ReadAdviceSignSource(1, mlng����ID, mstr�Һŵ�, strIDs, 0, mblnMoved, strSource, mlngǰ��ID)
    If intRule = 0 Then Exit Sub
    If strSource = "" Then
        MsgBox "�ò���Ŀǰû�п���ǩ����ҽ����", vbInformation, gstrSysName
        Exit Sub
    End If
    
    strSign = gobjESign.Signature(strSource, gstrDBUser, lng֤��ID)
    If strSign <> "" Then
        lngǩ��ID = zlDatabase.GetNextId("ҽ��ǩ����¼")
        strSQL = "zl_ҽ��ǩ����¼_Insert(" & lngǩ��ID & ",1," & intRule & ",'" & Replace(strSign, "'", "''") & "'," & lng֤��ID & ",'" & strIDs & "')"
        On Error GoTo errH
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        On Error GoTo 0
        
        Call LoadAdvice 'ˢ�½���
        MsgBox "����ɵ���ǩ����", vbInformation, gstrSysName
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FuncAdviceSignErase()
'���ܣ�ȡ��ҽ���ĵ���ǩ��
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
        
    If mint״̬ <> 1 Then Exit Sub '���ﲡ��
    If gobjESign Is Nothing Then Exit Sub
    If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)) = 0 Or tabAppend.SelectedItem.Index <> 3 Then Exit Sub
    
    With vsAppend
        If .RowData(.Row) = 0 Then
            MsgBox "��ǰѡ���ҽ��û��ǩ����", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '����ǩ������ȡ��
        If .Cell(flexcpData, .Row, 0) = 4 Then
            MsgBox "����ҽ����ǩ������ȡ����", vbInformation, gstrSysName
            Exit Sub
        End If
        '�¿�ǩ�����������¿�״̬
        If .Cell(flexcpData, .Row, 0) = 1 Then
            If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ҽ��״̬)) <> 1 Then
                MsgBox "����ҽ���Ѿ����ͻ����ϣ���ǩ������ȡ����", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        '����ȡ��ҽ���´��ǩ��
        If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ǰ��ID)) <> 0 Then
            MsgBox "�㲻��ȡ��ҽ�������´�ҽ����ǩ����", vbInformation, gstrSysName
            Exit Sub
        End If
        'ֻ��ȡ������ǩ����
        If .TextMatrix(.Row, 2) <> UserInfo.���� Then
            MsgBox "��ǩ���˲����㱾�ˣ�����ȡ��ǩ����", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If MsgBox("ȷʵҪȡ�����ǩ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        If Not gobjESign.CheckCertificate(gstrDBUser) Then Exit Sub
        
        strSQL = "zl_ҽ��ǩ����¼_Delete(" & .RowData(.Row) & ")"
        On Error GoTo errH
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        On Error GoTo 0
    End With
    
    Call LoadAdvice 'ˢ�½���
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FuncAdviceSignVerify()
'���ܣ�У��ҽ���ĵ���ǩ��(�ɶ���ת�Ƶ�����)
    Dim strSource As String
    
    If gobjESign Is Nothing Then Exit Sub
    If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)) = 0 Or tabAppend.SelectedItem.Index <> 3 Then Exit Sub
    
    With vsAppend
        If .RowData(.Row) = 0 Then
            MsgBox "��ǰѡ���ҽ��û��ǩ����", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '��ȡǩ��ҽ��Դ��
        If ReadAdviceSignSource(.Cell(flexcpData, .Row, 0), 0, 0, "", .RowData(.Row), mblnMoved, strSource) = 0 Then Exit Sub
        
        '��֤ǩ��
        Call gobjESign.VerifySignature(strSource, .RowData(.Row), 1)
    End With
End Sub

Private Sub FuncAdviceAdd()
'���ܣ�����ҽ��
    If Not CheckWindow Then Exit Sub
    If mlng����ID = 0 Then Exit Sub
    If mint״̬ <> 1 Then Exit Sub '���ﲡ��
    
    Set mfrmEdit = frmOutAdviceEdit
    Call frmOutAdviceEdit.ShowMe(mfrmParent, mstrPrivs, mlng����ID, mstr�Һŵ�, mlngǰ��ID)
End Sub

Private Sub FuncAdviceDel()
'ɾ����ɾ����ǰҽ��
'˵������������ɾ��,�Լ�����,�������,��ҩ�䷽,������ɾ��,һ����ҩֻɾ����ǰҩƷ
    Dim strSQL As String, lngҽ��ID As Long
    Dim blnGroup As Boolean, i As Long
    Dim lngRow As Long
    
    If mint״̬ <> 1 Then Exit Sub '���ﲡ��
    
    With vsAdvice
        '����Ƿ����ɾ��
        lngҽ��ID = Val(.TextMatrix(.Row, COL_ID))
        If lngҽ��ID = 0 Then
            MsgBox "�ò���û��ҽ������ɾ����", vbInformation, gstrSysName
            Exit Sub
        End If
        
        'ҽ���´��ҽ��
        If Val(.TextMatrix(.Row, COL_ǰ��ID)) <> mlngǰ��ID Then
            MsgBox "�㲻��ɾ����ҽ����", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If Val(.TextMatrix(.Row, COL_ҽ��״̬)) <> 1 Then
            MsgBox "��ǰѡ���ҽ���Ѿ����ͻ����ϣ�����ɾ����", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '��ǩ����ҽ������ɾ��
        If Val(.TextMatrix(.Row, COL_ǩ����)) = 1 Then
            MsgBox "��ǰѡ���ҽ���Ѿ�ǩ��������ɾ��������ȡ��ǩ����", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If InStr(",5,6,", .TextMatrix(.Row, COL_�������)) > 0 Then
            If .Row - 1 >= .FixedRows Then
                If Val(.TextMatrix(.Row - 1, COL_���ID)) = Val(.TextMatrix(.Row, COL_���ID)) Then blnGroup = True
            End If
            If Not blnGroup And .Row + 1 <= .Rows - 1 Then
                If Val(.TextMatrix(.Row + 1, COL_���ID)) = Val(.TextMatrix(.Row, COL_���ID)) Then blnGroup = True
            End If
            If blnGroup Then
                If MsgBox("ҽ��""" & .TextMatrix(.Row, COL_ҽ������) & """������ҩƷһ����ҩ,ȷʵҪɾ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            End If
        End If
        
        If Not blnGroup Then
            If MsgBox("ȷʵҪɾ��ҽ��""" & .TextMatrix(.Row, COL_ҽ������) & """��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        End If
        
        strSQL = "ZL_����ҽ����¼_Delete(" & lngҽ��ID & ",1)"
    End With
    
    On Error GoTo errH
    gcnOracle.BeginTrans
    zlDatabase.ExecuteProcedure strSQL, Me.Name
    gcnOracle.CommitTrans
    On Error GoTo 0
    
    With vsAdvice
        '������ֱ��ɾ��
        .Redraw = False
        
        'ɾ��һ����ҩ��һ��ʱ����ʾ����
        If blnGroup And .Row + 1 <= .Rows - 1 Then
            If Val(.TextMatrix(.Row, COL_���ID)) = Val(.TextMatrix(.Row + 1, COL_���ID)) Then
                If .TextMatrix(.Row, COL_��ʼʱ��) <> "" And .TextMatrix(.Row + 1, COL_��ʼʱ��) = "" Then
                    .TextMatrix(.Row + 1, COL_��ʼʱ��) = .TextMatrix(.Row, COL_��ʼʱ��)
                    .TextMatrix(.Row + 1, COL_Ƶ��) = .TextMatrix(.Row, COL_Ƶ��)
                    .TextMatrix(.Row + 1, COL_�÷�) = .TextMatrix(.Row, COL_�÷�)
                End If
            End If
        End If
        
        lngRow = .Row
        .RemoveItem .Row
        If .Rows = .FixedRows Then .Rows = .FixedRows + 1
        If lngRow <= .Rows - 1 Then
            .Row = lngRow
        Else
            .Row = .Rows - 1
        End If

        Call .ShowCell(.Row, .Col)
        .Redraw = True
        
        Call vsAdvice_AfterRowColChange(-1, -1, .Row, .Col) '��ɫ���������
    End With
    Call ShowTotalMoney
    Exit Sub
errH:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FuncAdviceRevoke()
'ɾ������ǰҽ������(һ��ҽ������)
    Dim strSQL As String, lngҽ��ID As Long
    Dim lng֤��ID As Long, lngǩ��ID As Long
    Dim strSign As String, intRule As Integer
    Dim strSource As String, strIDs As String
    
    If mint״̬ <> 1 Then Exit Sub '���ﲡ��
    
    With vsAdvice
        '����Ƿ��������
        If Val(.TextMatrix(.Row, COL_���ID)) <> 0 Then
            lngҽ��ID = Val(.TextMatrix(.Row, COL_���ID))
        Else
            lngҽ��ID = Val(.TextMatrix(.Row, COL_ID))
        End If
        If lngҽ��ID = 0 Then
            MsgBox "�ò���û��ҽ���������ϡ�", vbInformation, gstrSysName
            Exit Sub
        End If
                
        If Val(.TextMatrix(.Row, COL_ǰ��ID)) <> mlngǰ��ID Then
            MsgBox "�㲻�����ϸ�ҽ����", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If Val(.TextMatrix(.Row, COL_ҽ��״̬)) <> 8 Then
            MsgBox "��ǰѡ���ҽ����δ���ͻ��Ѿ����ϡ�", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '���з���ת������������
        If MovedByDate(.Cell(flexcpData, .Row, COL_����ʱ��)) Then
            If MovedBySend(lngҽ��ID) Then
                MsgBox "��ҽ���ķ����Ѿ�ȫ���򲿷�ת���������ݿ⣬�����������" & vbCrLf & _
                    "��������ϵͳ����Ա��ϵ������Ӧ���ݳ�ѡ���ء�", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        
        '����ǩ��������ʾ
        If Val(.TextMatrix(.Row, COL_ǩ����)) = "1" Then
            If gobjESign Is Nothing Then
                If gintCA = 0 Then
                    MsgBox "������ǩ��ҽ��ʱ��Ҫ�ٴ�ǩ������ϵͳû������ǩ����֤���ģ��������ϡ�", vbInformation, gstrSysName
                Else
                    MsgBox "������ǩ��ҽ��ʱ��Ҫ�ٴ�ǩ����������ǩ������δ����ȷ��װ���������ϡ�", vbInformation, gstrSysName
                End If
                Exit Sub
            End If
            strSign = vbCrLf & vbCrLf & "��ʾ����ҽ���Ѿ�ǩ��������ʱ����Ҫ�ٴ�ǩ����"
        End If
                
        '�������ҽ����Ӧ�ķ��ý������
        If Not CheckAdviceBalanceRevoke(lngҽ��ID) Then Exit Sub
                
        If RowInһ����ҩ(.Row, 0, 0) Then
            If MsgBox("����һ����ҩ��ҽ������һ�����ϣ�ȷʵҪ������" & strSign, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        Else
            If MsgBox("ȷʵҪ����ҽ��""" & .TextMatrix(.Row, COL_ҽ������) & """��" & strSign, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        End If
        
        strSQL = "ZL_����ҽ����¼_����(" & lngҽ��ID & ")"
        
        '����ʱ���е���ǩ��
        If strSign <> "" Then
            '��ȡǩ��ҽ��Դ��
            strIDs = lngҽ��ID
            intRule = ReadAdviceSignSource(4, mlng����ID, mstr�Һŵ�, strIDs, 0, mblnMoved, strSource)
            If intRule = 0 Then Exit Sub
            If strSource = "" Then
                MsgBox "���ܶ�ȡ��Ҫ���ϵ���ǩ��ҽ��Դ�����ݡ�", vbInformation, gstrSysName
                Exit Sub
            End If
            
            strSign = gobjESign.Signature(strSource, gstrDBUser, lng֤��ID)
            If strSign <> "" Then
                lngǩ��ID = zlDatabase.GetNextId("ҽ��ǩ����¼")
                strSign = "zl_ҽ��ǩ����¼_Insert(" & lngǩ��ID & ",4," & intRule & ",'" & Replace(strSign, "'", "''") & "'," & lng֤��ID & ",'" & strIDs & "')"
            Else
                Exit Sub
            End If
        End If
    End With
    
    On Error GoTo errH
    gcnOracle.BeginTrans
    zlDatabase.ExecuteProcedure strSQL, Me.Name
    If strSign <> "" Then
        zlDatabase.ExecuteProcedure strSign, Me.Name
    End If
    gcnOracle.CommitTrans
    On Error GoTo 0
    
    Call LoadAdvice 'ˢ�½���
    Call ShowTotalMoney
    Exit Sub
errH:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FuncAdviceModi()
'���ܣ��޸ĵ�ǰҽ��
    Dim lngҽ��ID As Long
    
    If Not CheckWindow Then Exit Sub
    
    If mlng����ID = 0 Then Exit Sub
    If mint״̬ <> 1 Then Exit Sub '���ﲡ��
    
    With vsAdvice
        lngҽ��ID = Val(.TextMatrix(.Row, COL_ID))
        If lngҽ��ID = 0 Then Exit Sub
        
        'ҽ���´��ҽ��
        If Val(.TextMatrix(.Row, COL_ǰ��ID)) <> mlngǰ��ID Then
            MsgBox "�㲻���޸ĸ�ҽ����", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '��У�Ի��ѷ�ֹ
        If Val(.TextMatrix(.Row, COL_ҽ��״̬)) <> 1 Then
            MsgBox "��ǰѡ���ҽ���Ѿ����ͻ����ϣ������޸ġ�", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '��ǩ����ҽ�������޸�
        If Val(.TextMatrix(.Row, COL_ǩ����)) = 1 Then
            MsgBox "��ǰѡ���ҽ���Ѿ�ǩ���������޸ġ�����ȡ��ǩ����", vbInformation, gstrSysName
            Exit Sub
        End If
        
        Set mfrmEdit = frmOutAdviceEdit
        Call frmOutAdviceEdit.ShowMe(mfrmParent, mstrPrivs, mlng����ID, mstr�Һŵ�, mlngǰ��ID, _
            Val(.TextMatrix(.Row, COL_Ӥ��ID)), lngҽ��ID)
    End With
End Sub

Private Sub FuncAdviceTest()
'���ܣ���дƤ�Խ��
    Dim rsTmp As New ADODB.Recordset, strSQL As String
    Dim v��� As VbMsgBoxResult, str��� As String
    
    If mlng����ID = 0 Then Exit Sub
    If mint״̬ <> 1 Then Exit Sub '���ﲡ��
    
    If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)) = 0 Then Exit Sub
    If Not (vsAdvice.TextMatrix(vsAdvice.Row, COL_�������) = "E" And vsAdvice.TextMatrix(vsAdvice.Row, COL_��������) = "1") Then
        MsgBox "��ǰҽ�����ݲ��ǹ���������Ŀ��", vbInformation, gstrSysName
        Exit Sub
    End If
    If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ǰ��ID)) <> 0 Then
        MsgBox "�㲻�ܸ��ù���������д�����", vbInformation, gstrSysName
        Exit Sub
    End If
    If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ҽ��״̬)) = 4 Then
        MsgBox "�ù�������ҽ���Ѿ����ϣ�������д�����", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ҽ��״̬)) = 1 Then
        If vsAdvice.TextMatrix(vsAdvice.Row, COL_Ƥ��) = "����" Then
            If MsgBox("�ù�������ҽ���Ѿ����Ϊ���ԣ�Ҫ������Ա����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            str��� = ""
        Else
            If MsgBox("�ù�������ҽ����δ���ͣ���������д������������" & vbCrLf & vbCrLf & _
                "�����Ա��Ϊ���ԣ�ͬʱ��ҽ�������ᷢ�͡�Ҫ���Ϊ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            str��� = "����"
        End If
    Else
        '����Ӧ��ҽ���Ƿ��Ѿ�����
        If mbln�Զ�Ƥ�� Then
            If AdviceSended(Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID))) Then
                MsgBox "��Ƥ�Զ�Ӧ��ҩƷ�Ѿ����ͣ������ٸ���Ƥ�Խ����", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        
        If vsAdvice.TextMatrix(vsAdvice.Row, COL_Ƥ��) <> "" Then
            If MsgBox("�ù�������ҽ���Ѿ���д�˽����Ҫ������д��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        End If
        
        v��� = frmMsgBox.ShowMsgBox(vsAdvice.TextMatrix(vsAdvice.Row, COL_ҽ������) & "��^^����ݹ���������ѡ����Ӧ�İ�ť������", mfrmParent, , 1)
        If v��� = vbCancel Then Exit Sub
        str��� = IIF(v��� = vbYes, "(+)", "(-)")
    End If
    
    strSQL = "ZL_����ҽ����¼_Ƥ��(" & Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)) & ",'" & str��� & "')"
    
    On Error GoTo errH
    gcnOracle.BeginTrans
    zlDatabase.ExecuteProcedure strSQL, Me.Name
    gcnOracle.CommitTrans
    On Error GoTo 0
    
    vsAdvice.TextMatrix(vsAdvice.Row, COL_Ƥ��) = str���
    If str��� = "(+)" Then
        vsAdvice.Cell(flexcpForeColor, vsAdvice.Row, COL_Ƥ��) = vbRed
    ElseIf str��� = "(-)" Then
        vsAdvice.Cell(flexcpForeColor, vsAdvice.Row, COL_Ƥ��) = vbBlue
    End If
    Exit Sub
errH:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function AdviceSended(ByVal lngҽ��ID As Long) As Boolean
'���ܣ��ж�Ƥ�Զ�Ӧ��ҽ���Ƿ��Ѿ�����
'������lngҽ��ID=Ƥ��ҽ����ID
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    '�����ϵĲ���
    strSQL = "Select ������ĿID From ����ҽ����¼ Where ID=[3]"
    strSQL = "Select A.ID From ����ҽ����¼ A,�����÷����� B" & _
        " Where Rownum<2 And A.������� IN('5','6') And A.ҽ��״̬=8" & _
        " And A.������ĿID=B.��ĿID And B.����=0 And B.�÷�ID=(" & strSQL & ")" & _
        " And A.����ID+0=[1] And A.�Һŵ�=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, mlng����ID, mstr�Һŵ�, lngҽ��ID)
    AdviceSended = Not rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub FuncAdviceSend()
'���ܣ����Ͳ���ҽ��(�������üƼ���Ŀ)

    If mlng����ID = 0 Then Exit Sub
    If mint״̬ <> 1 Then Exit Sub '���ﲡ��
    
    If frmOutAdviceSend.ShowMe(Me, mstrPrivs, mlng����ID, mstr�Һŵ�, mlngǰ��ID) Then
        Call LoadAdvice
        Call ShowTotalMoney
    End If
End Sub

Private Sub tabAppend_Click()
    If Val(vsAppend.Tag) = tabAppend.SelectedItem.Index Then Exit Sub
    
    If Visible Then
        Call SaveFlexState(vsAppend, App.ProductName & "\" & Me.Name)
    End If
        
    vsAppend.Tag = tabAppend.SelectedItem.Index
    If tabAppend.SelectedItem.Index = 1 Then
        Call InitPriceTable
    ElseIf tabAppend.SelectedItem.Index = 2 Then
        Call InitSendTable
    ElseIf tabAppend.SelectedItem.Index = 3 Then
        Call InitSignTable
    End If
    
    If Visible Then
        Call RestoreFlexState(vsAppend, App.ProductName & "\" & Me.Name)
    End If
    
    Call vsAdvice_AfterRowColChange(-1, -1, vsAdvice.Row, vsAdvice.Col)
    
    If Visible Then vsAdvice.SetFocus
End Sub

Private Sub vsAdvice_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    'Ϊ���ⲿϵͳ�������ӣ�By����ͮ��
    On Error Resume Next
    
    If NewRow = OldRow Then Exit Sub
    If vsAdvice.Col >= vsAdvice.FixedCols Then
        vsAdvice.ForeColorSel = vsAdvice.Cell(flexcpForeColor, NewRow, COL_��ʼʱ��)
    End If
    If vsAdvice.Redraw <> flexRDNone Then
        If Val(vsAdvice.TextMatrix(NewRow, COL_ID)) <> 0 Then
            '��ʾҽ�����ӱ�������
            If mfrmParent.mnuViewAdviceAppend.Checked Then
                If tabAppend.SelectedItem.Index = 1 Then
                    Call ShowPrice(NewRow)
                ElseIf tabAppend.SelectedItem.Index = 2 Then
                    Call ShowSendList(NewRow)
                ElseIf tabAppend.SelectedItem.Index = 3 Then
                    Call ShowSignList(NewRow)
                End If
            End If
        ElseIf mfrmParent.mnuViewAdviceAppend.Checked Then
            Call ClearAppendData
            vsAppend.Row = vsAppend.FixedRows
        End If
    End If
End Sub

Private Sub vsAdvice_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Dim lngW As Long
    
    If Col = COL_ҽ������ Then
        vsAdvice.AutoSize Col
    ElseIf Row = -1 Then
        lngW = Me.TextWidth(vsAdvice.TextMatrix(vsAdvice.FixedRows - 1, Col) & "A")
        If vsAdvice.ColWidth(Col) < lngW Then
            vsAdvice.ColWidth(Col) = lngW
        ElseIf vsAdvice.ColWidth(Col) > vsAdvice.Width * 0.5 Then
            vsAdvice.ColWidth(Col) = vsAdvice.Width * 0.5
        End If
    End If
End Sub

Private Sub vsAdvice_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Row = -1 Then
        If Col <= vsAdvice.FixedCols - 1 Then
            Cancel = True
        ElseIf Col = COL_Ƥ�� Then
            Cancel = True
        ElseIf Col = COL_��ʾ Then 'Pass
            Cancel = True
        End If
    End If
End Sub

Private Sub vsAdvice_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
'˵����1.OwnerDrawҪ����ΪOver(������Ԫ��������)
'      2.Cell��GridLine�������������ڶ��Ǵӵ�1���߿�ʼ
'      3.Cell��Border�������Ǵӵ�2���߿�ʼ,�����Ǵӵ�1���߿�ʼ
    Dim lngLeft As Long, lngRight As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim vRect As RECT
    
    With vsAdvice
        If Col <= .FixedCols - 1 Then
            '�����̶����еı����
            SetBkColor hDC, SysColor2RGB(.BackColorFixed)

            '����߱����
            vRect.Left = Left
            vRect.Top = Top
            vRect.Right = Left + 1
            vRect.Bottom = Bottom
            If Row = .Rows - 1 Then vRect.Bottom = vRect.Bottom - 1
            ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0

            '���ϱ߱����
            vRect.Left = Left
            vRect.Top = Top
            vRect.Right = Right
            vRect.Bottom = Top + 1
            If Col = .FixedCols - 1 Then vRect.Right = vRect.Right - 1
            ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0

            '���±߱����
            vRect.Left = Left
            vRect.Top = Bottom - 1
            vRect.Right = Right
            vRect.Bottom = Bottom
            If Row = .Rows - 1 Then vRect.Bottom = vRect.Bottom - 1
            If Col = .FixedCols - 1 Then vRect.Right = vRect.Right - 1
            ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0

            '���ұ߱����
            vRect.Left = Right - 1
            vRect.Top = Top
            vRect.Right = Right
            vRect.Bottom = Bottom
            If Row = .Rows - 1 Then vRect.Bottom = vRect.Bottom - 1
            If Col = .FixedCols - 1 Then vRect.Right = vRect.Right - 1
            ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0
        Else
            '����һ����ҩ������еı��߼�����
            lngLeft = COL_��ʼʱ��: lngRight = COL_��ʼʱ��
            If Not Between(Col, lngLeft, lngRight) Then
                lngLeft = COL_Ƶ��: lngRight = COL_�÷�
                If Not Between(Col, lngLeft, lngRight) Then Exit Sub
            End If
            
            If Not RowInһ����ҩ(Row, lngBegin, lngEnd) Then Exit Sub
            
            vRect.Left = Left '������߱����
            vRect.Right = Right - 1 '�����ұ߱����
            If Row = lngBegin Then
                vRect.Top = Bottom - 1 '���б�����������
                vRect.Bottom = Bottom
            Else
                If Row = lngEnd Then
                    vRect.Top = Top
                    vRect.Bottom = Bottom - 1 '���б����±���
                Else
                    vRect.Top = Top
                    vRect.Bottom = Bottom
                End If
                'Ϊ��֧��Ԥ�����
                If .TextMatrix(Row, Col) <> "" Then .TextMatrix(Row, Col) = ""
            End If
            
            If Between(Row, .Row, .RowSel) Then
                SetBkColor hDC, SysColor2RGB(.BackColorSel)
            Else
                SetBkColor hDC, SysColor2RGB(.BackColor)
            End If
            ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0
        End If
        Done = True
    End With
End Sub

Private Sub vsAdvice_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Ϊ���ⲿϵͳ�������ӣ�By����ͮ��
    On Error Resume Next
    
    If Button = 2 And mfrmParent.mnuAdvice.Visible Then PopupMenu mfrmParent.mnuAdvice, 2
End Sub

Private Function GetPatiInfo() As String
'���ܣ���ȡ������Ϣ��(���ڴ�ӡ)
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    'ִ�в���(�ű����)�����˿���
    strSQL = "Select B.����,B.�Ա�,B.����,B.�����," & _
        " B.����,B.��������,C.���� as ִ�в���,A.ִ�в���ID,A.�Ǽ�ʱ��" & _
        " From ���˹Һż�¼ A,������Ϣ B,���ű� C" & _
        " Where A.NO=[2] And A.����ID+0=[1]" & _
        " And A.����ID=B.����ID And A.ִ�в���ID=C.ID"
    If mblnMoved Then
        strSQL = Replace(strSQL, "���˹Һż�¼", "H���˹Һż�¼")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, mlng����ID, mstr�Һŵ�)
    
    GetPatiInfo = _
        "������" & rsTmp!���� & " �Ա�" & Nvl(rsTmp!�Ա�) & _
        " ���䣺" & Nvl(rsTmp!����) & " ����ţ�" & Nvl(rsTmp!�����) & _
        " �Һţ�" & Format(rsTmp!�Ǽ�ʱ��, "MM-dd HH:mm") & _
        " ���ң�" & rsTmp!ִ�в��� & " ���ң�" & Nvl(rsTmp!��������)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub OutputList(bytStyle As Byte)
'���ܣ�������б�
'������bytStyle=1-��ӡ,2-Ԥ��,3-�����Excel
    Dim objOut As New zlPrint1Grd
    Dim objRow As zlTabAppRow
    Dim bytR As Byte, i As Long
    Dim lngRow As Long, lngCol As Long
    Dim strWidth As String
    
    If mlng����ID = 0 Then Exit Sub
    
    '��ͷ
    objOut.Title.Text = "����ҽ���嵥"
    objOut.Title.Font.Name = "����_GB2312"
    objOut.Title.Font.Size = 18
    objOut.Title.Font.Bold = True
    
    '����
    Set objRow = New zlTabAppRow
    objRow.Add GetPatiInfo
    objOut.UnderAppRows.Add objRow
    
    '����
    Set objRow = New zlTabAppRow
    objRow.Add "��ӡ�ˣ�" & UserInfo.����
    objRow.Add "��ӡ���ڣ�" & Format(zlDatabase.Currentdate(), "yyyy��MM��dd��")
    objOut.BelowAppRows.Add objRow
    
    '����
    Set objOut.Body = vsAdvice
    
    '���
    vsAdvice.Redraw = False
    lngRow = vsAdvice.Row: lngCol = vsAdvice.Col
    
    strWidth = ""
    For i = 0 To vsAdvice.FixedCols - 1
        strWidth = strWidth & "," & vsAdvice.ColWidth(i)
        vsAdvice.ColWidth(i) = 0
    Next
        
    If bytStyle = 1 Then
        bytR = zlPrintAsk(objOut)
        Me.Refresh
        If bytR <> 0 Then zlPrintOrView1Grd objOut, bytR
    Else
        zlPrintOrView1Grd objOut, bytStyle
    End If
    
    strWidth = Mid(strWidth, 2)
    For i = 0 To vsAdvice.FixedCols - 1
        vsAdvice.ColWidth(i) = Split(strWidth, ",")(i)
    Next
    
    vsAdvice.Row = lngRow: vsAdvice.Col = lngCol
    vsAdvice.Redraw = True
End Sub

Private Sub Form_Load()
    'Ϊ���ⲿϵͳ�������ӣ�By����ͮ��
    On Error Resume Next
    
    '����ǩ����¼
    If gobjESign Is Nothing Then
        tabAppend.Tabs.Remove 3
    End If
    
    Call InitAdviceTable
    Call tabAppend_Click
    Call RestoreWinState(Me, App.ProductName)
    
    '�Զ�����Ƥ��
    mbln�Զ�Ƥ�� = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "�Զ�����Ƥ��", 0)) <> 0

    fraAdviceUD.Visible = mfrmParent.mnuViewAdviceAppend.Checked
    tabAppend.Visible = mfrmParent.mnuViewAdviceAppend.Checked
    vsAppend.Visible = mfrmParent.mnuViewAdviceAppend.Checked
    
    Set mfrmEdit = Nothing
    
    Call InitSysPar '��ʼ��ϵͳ����
End Sub

Private Sub Form_Resize()
    Dim PriceH As Long

    On Error Resume Next
    
    If WindowState = 1 Then Exit Sub
    
    PriceH = IIF(vsAppend.Visible, vsAppend.Height + fraAdviceUD.Height + tabAppend.Height, 0)
    
    vsAdvice.Left = 0
    vsAdvice.Top = 0
    vsAdvice.Width = Me.ScaleWidth
    vsAdvice.Height = Me.ScaleHeight - PriceH
    
    fraAdviceUD.Left = 0
    fraAdviceUD.Top = vsAdvice.Top + vsAdvice.Height
    fraAdviceUD.Width = Me.ScaleWidth
    
    tabAppend.Left = 0
    tabAppend.Top = fraAdviceUD.Top + fraAdviceUD.Height
    tabAppend.Width = Me.ScaleWidth
    
    vsAppend.Left = 0
    vsAppend.Top = tabAppend.Top + tabAppend.Height
    vsAppend.Width = Me.ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mfrmEdit = Nothing
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub ClearAppendData()
'���ܣ�������ӱ������
    vsAppend.Rows = vsAppend.FixedRows
    vsAppend.Rows = vsAppend.FixedRows + 1
End Sub

Private Sub InitPriceTable()
'���ܣ���ʼ���Ƽ��嵥��ʽ
    Dim arrHead As Variant, strHead As String, i As Long
    
    strHead = "�Ƽ�ҽ��,2000,1;���,650,1;�շ���Ŀ,2500,1;��λ,500,4;����,500,1;����,800,7;ִ�п���,1000,1;��������,800,1;����,450,4"
    arrHead = Split(strHead, ";")
    With vsAppend
        .Clear
        .FixedRows = 1
        .FixedCols = 0
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
        
        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColHidden(.FixedCols + i) = False
                .ColWidth(.FixedCols + i) = Val(Split(arrHead(i), ",")(1))
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
            End If
        Next
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = 4
        .MergeCol(0) = False
        .MergeCol(1) = False
    End With
End Sub

Private Sub InitSendTable()
'���ܣ���ʼ�������嵥��ʽ
    Dim arrHead As Variant, strHead As String, i As Long
    
    strHead = "���ͺ�;����ʱ��,1080,1;����ҽ��,1800,1;���ݺ�,850,1;�շ���Ŀ,1800,1;��������,850,1;�Ʒ�״̬,850,1;ִ��״̬,850,1;ִ�п���,850,1;������,800,1;��¼����"
    arrHead = Split(strHead, ";")
    With vsAppend
        .Clear
        .FixedRows = 1
        .FixedCols = 0
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
        
        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColHidden(.FixedCols + i) = False
                .ColWidth(.FixedCols + i) = Val(Split(arrHead(i), ",")(1))
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
            End If
        Next
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = 4
        .MergeCol(0) = True
        .MergeCol(1) = True
    End With
End Sub

Private Sub InitSignTable()
'���ܣ���ʼ���Ƽ��嵥��ʽ
    Dim arrHead As Variant, strHead As String, i As Long
    
    strHead = "ǩ������,1150,1;ǩ��ʱ��,1900,1;ǩ����,800,1"
    arrHead = Split(strHead, ";")
    With vsAppend
        .Clear
        .FixedRows = 1
        .FixedCols = 0
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
        
        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColHidden(.FixedCols + i) = False
                .ColWidth(.FixedCols + i) = Val(Split(arrHead(i), ",")(1))
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
            End If
        Next
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = 4
        .MergeCol(0) = False
        .MergeCol(1) = False
    End With
End Sub

Private Sub ClearAdviceData()
'���ܣ����ҽ���嵥����
    vsAdvice.Rows = vsAdvice.FixedRows
    vsAdvice.Rows = vsAdvice.FixedRows + 1
    vsAdvice.Editable = flexEDNone
End Sub

Private Sub InitAdviceTable()
'���ܣ���ʼ��ҽ���嵥��ʽ
    Dim arrHead As Variant, strHead As String, i As Long

    strHead = "ID;���ID;��ID;���;Ӥ��ID;ҽ��״̬;�������;��������;�������;��־;" & _
        ",240,4;��ʼʱ��,1080,1;ҽ������,3000,1;,375,4;����,850,1;����,850,1;Ƶ��,1000,1;" & _
        "�÷�,1000,1;ҽ������,1000,1;ִ��ʱ��,1000,1;ִ�п���,850,1;ִ������,850,1;" & _
        "����ҽ��,850,1;����ʱ��,1080,1;������,850,1;����ʱ��,1080,1;" & _
        "����ID;������;������;����ID;ǰ��ID;ǩ����"
    arrHead = Split(strHead, ";")
    With vsAdvice
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
        .ColHidden(COL_��ʾ) = Not (gblnPass And InStr(mstrPrivs, "������ҩ���") > 0) 'Pass
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = 4
        .ColWidth(0) = 9 * Screen.TwipsPerPixelX
        .ColWidth(1) = 11 * Screen.TwipsPerPixelX
    End With
End Sub

Private Function LoadAdvice() As Boolean
'���ܣ����ݵ�ǰ�������ö�ȡ����ʾҽ���嵥
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, lngTop As Long
    Dim strFormat As String, strTmp As String
    Dim bln��ҩ;�� As Boolean, bln��ҩ�÷� As Boolean
        Dim bln�ɼ����� As Boolean, bln������ As Boolean, bln������ As Boolean
    Dim blnFirst As Boolean, lngҽ��ID As Long
    Dim strBill As String, i As Long, j As Long
    
    If mlng����ID = 0 Then Exit Function
    
    Screen.MousePointer = 11
    
    On Error GoTo errH
    
    lngҽ��ID = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)) '��¼��ǰ��
        
    '���Ƶ��ݣ���Ӧ���Ƶ���,��������,������
    strBill = "Select A.ID as ҽ��ID,B.�����ļ�ID as ����ID," & _
        " Max(Decode(C.��дʱ��,1,1,0)) as ������," & _
        " Max(Decode(C.��дʱ��,2,1,0)) as ������" & _
        " From ����ҽ����¼ A,���Ƶ���Ӧ�� B,�����ļ���� C" & _
        " Where A.������ĿID=B.������ĿID And B.Ӧ�ó���=1 And B.�����ļ�ID=C.�����ļ�ID(+)" & _
        " And A.����ID+0=[1] And A.�Һŵ�=[2]" & _
        " And Not(A.������� IN ('F','G','D','E') And A.���ID is Not NULL)" & _
        " Group by A.ID,B.�����ļ�ID"
        
    'ҽ����¼��������������,��������,��鲿λ,��ҩ�巨
    strSQL = _
        "Select /*+ RULE */ A.ID,A.���ID,Nvl(A.���ID,A.ID) as ��ID,Nvl(X.���,A.���) as ���," & _
            " Nvl(A.Ӥ��,0) as Ӥ��ID,A.ҽ��״̬,A.�������,B.��������,C.�������,A.������־ as ��־," & _
            " A.�����,To_Char(A.��ʼִ��ʱ��,'MM-DD HH24:MI') as ��ʼʱ��,A.ҽ������,A.Ƥ�Խ�� as Ƥ��," & _
            " Decode(A.�ܸ�����,NULL,NULL,Decode(A.�������,'E',Decode(B.��������,'4',A.�ܸ�����||'��',A.�ܸ�����||B.���㵥λ),'5',Round(A.�ܸ�����/D.�����װ,5)||D.���ﵥλ,'6',Round(A.�ܸ�����/D.�����װ,5)||D.���ﵥλ,A.�ܸ�����||B.���㵥λ)) as ����," & _
            " Decode(A.��������,NULL,NULL,A.��������||B.���㵥λ) as ����," & _
            " A.ִ��Ƶ�� as Ƶ��,Decode(A.�������,'E',Decode(Instr('246',Nvl(B.��������,'0')),0,NULL,B.����),NULL) as �÷�," & _
            " A.ҽ������,A.ִ��ʱ�䷽�� as ִ��ʱ��,Nvl(E.����,Decode(Nvl(A.ִ������,0),0,'<����>',5,'<Ժ��ִ��>')) as ִ�п���," & _
            " Decode(Instr('567E',A.�������),0,NULL,A.ִ������) as ִ������," & _
            " A.����ҽ��,To_Char(A.����ʱ��,'MM-DD HH24:MI') as ����ʱ��," & _
            " A.ͣ��ҽ�� as ������,A.ͣ��ʱ�� as ����ʱ��," & _
            " Y.����ID,Y.������,Y.������,A.����ID,A.ǰ��ID," & _
            " Decode(S.ǩ��ID,NULL,0,1) as ǩ����" & _
        " From ����ҽ����¼ A,���ű� E,ҩƷ���� C,ҩƷ��� D,������ĿĿ¼ B,����ҽ��״̬ S,����ҽ����¼ X,(" & strBill & ") Y" & _
        " Where A.������ĿID=B.ID And A.ִ�п���ID=E.ID(+) And A.������ĿID=C.ҩ��ID(+)" & _
            " And Nvl(A.ҽ����Ч,0)=1 And A.�շ�ϸĿID=D.ҩƷID(+) And A.���ID=X.ID(+)" & _
            " And Not(A.������� IN ('F','G','D','E') And A.���ID is Not NULL)" & _
            " And A.��ʼִ��ʱ�� is Not NULL And A.������Դ<>3 And A.ID=S.ҽ��ID And S.��������=1" & _
            " And A.ID=Y.ҽ��ID(+) And A.����ID+0=[1] And A.�Һŵ�=[2]" & _
            IIF(mlngǰ��ID = 0 Or mblnShowAll, "", " And A.ǰ��ID=[3]") & _
        " Order by Nvl(A.Ӥ��,0),���,��ID,A.���"
        
    If mblnMoved Then '�Һŵ���ҽ��ͬ�����ݿ�
        strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, mlng����ID, mstr�Һŵ�, mlngǰ��ID)
    
    If Not rsTmp.EOF Then
        With vsAdvice
            .Redraw = False
            
            '��ʱ�����ʱ��FormatString�ָ�һЩȱʡֵ(�̶����������̶��������ּ����ж���,�ߴ�,�ɼ�)
            'FormatString������ʱ��ֵ��Ч
            '���AutoResize=True,�������п���и߱��Զ�����(����AutoSizeMode)
            '���WordWrap=True,���и߻ᱻ�Զ�����
            .WordWrap = False
            strFormat = GetColFormat(vsAdvice)
            Call ClearAdviceData
            .ScrollBars = flexScrollBarNone
            Set .DataSource = rsTmp
            .ScrollBars = flexScrollBarBoth
            If Err.Number = 0 And gcnOracle.Errors.Count > 0 Then
                gcnOracle.Errors.Clear '��,��ʱ�̶��д˴���
            End If
            Call SetColFormat(vsAdvice, strFormat)
            .TextMatrix(0, COL_Ƥ��) = ""
            .TextMatrix(0, COL_��ʾ) = "" 'Pass
            
            '�Զ������и�
            .WordWrap = True
            .AutoSize COL_ҽ������
            
            '����ÿ��ҽ��
            i = .FixedRows
            Do While i <= .Rows - 1
                '������ʱ��
                If .TextMatrix(i, COL_����ʱ��) <> "" Then
                    .Cell(flexcpData, i, COL_����ʱ��) = .TextMatrix(i, COL_����ʱ��)
                    .TextMatrix(i, COL_����ʱ��) = Format(.TextMatrix(i, COL_����ʱ��), "MM-dd HH:mm")
                End If
                
                '��ҩ����ҩ��һЩ����
                bln��ҩ;�� = False: bln��ҩ�÷� = False
                bln�ɼ����� = False: bln������ = False: bln������ = False '�����ڼ������
                If .TextMatrix(i, COL_�������) = "E" Then
                    If Val(.TextMatrix(i - 1, COL_���ID)) = Val(.TextMatrix(i, COL_ID)) Then
                        If InStr(",5,6,", .TextMatrix(i - 1, COL_�������)) > 0 Then
                            bln��ҩ;�� = True
                            For j = i - 1 To .FixedRows Step -1
                                If Val(.TextMatrix(j, COL_���ID)) = Val(.TextMatrix(i, COL_ID)) Then
                                    '��ʾ��ҩ�ĸ�ҩ;��
                                    .TextMatrix(j, COL_�÷�) = .TextMatrix(i, COL_�÷�)
                                    '��ʾ��ҩ��ִ������
                                    If Val(.TextMatrix(j, COL_ִ������)) = 5 And Val(.TextMatrix(i, COL_ִ������)) <> 5 Then
                                        .TextMatrix(j, COL_ִ������) = "�Ա�ҩ"
                                    ElseIf Val(.TextMatrix(j, COL_ִ������)) <> 5 And Val(.TextMatrix(i, COL_ִ������)) = 5 Then
                                        .TextMatrix(j, COL_ִ������) = "��Ժ��ҩ"
                                    Else
                                        .TextMatrix(j, COL_ִ������) = ""
                                    End If
                                Else
                                    Exit For
                                End If
                            Next
                        ElseIf InStr(",7,C,", .TextMatrix(i - 1, COL_�������)) > 0 Then
                            bln��ҩ�÷� = .TextMatrix(i - 1, COL_�������) = "7" '��ҩ�÷���
                            bln�ɼ����� = .TextMatrix(i - 1, COL_�������) = "C" '�ɼ�������

                            '��ʾ��ҩ�䷽�������ϵ�ִ�п���
                            .TextMatrix(i, COL_ִ�п���) = .TextMatrix(i - 1, COL_ִ�п���)
                            
                            If bln��ҩ�÷� Then
                                '��ʾ��ҩ�䷽ִ������
                                If Val(.TextMatrix(i - 1, COL_ִ������)) = 5 And Val(.TextMatrix(i, COL_ִ������)) <> 5 Then
                                    .TextMatrix(i, COL_ִ������) = "�Ա�ҩ"
                                ElseIf Val(.TextMatrix(i - 1, COL_ִ������)) <> 5 And Val(.TextMatrix(i, COL_ִ������)) = 5 Then
                                    .TextMatrix(i, COL_ִ������) = "��Ժ��ҩ"
                                Else
                                    .TextMatrix(i, COL_ִ������) = ""
                                End If
                            Else
                                .TextMatrix(i, COL_ִ������) = ""
                            End If
                            
                            'ɾ����ζ��ҩ��,�Լ���������еļ�����Ŀ;ͬʱ�жϼ�������
                            For j = i - 1 To .FixedRows Step -1
                                If Val(.TextMatrix(j, COL_���ID)) = Val(.TextMatrix(i, COL_ID)) Then
                                    If .TextMatrix(j, COL_�������) = "C" Then
                                        If Val(.TextMatrix(j, COL_������)) = 1 Then
                                            bln������ = True
                                            If Val(.TextMatrix(j, COL_����ID)) <> 0 Then
                                                bln������ = True
                                            End If
                                        End If
                                    End If
                                    .RemoveItem j: i = i - 1
                                Else
                                    Exit For
                                End If
                            Next
                        End If
                    Else
                        .TextMatrix(i, COL_ִ������) = ""
                    End If
                End If
                                                                
                '����ɼ��еĵ�һЩ��ʶ:�ſ����ɼ�����ʱδɾ������
                If Not bln��ҩ;�� And .TextMatrix(i, COL_�������) <> "7" Then
                    
                    '�иߣ�Ϊ��֧��zl9PrintMode:Resize֮��,ȡRowHeight����С��RowHeightMin
                    If .RowHeight(i) < .RowHeightMin Then .RowHeight(i) = .RowHeightMin
                    
                    '����С��������,��δ�뵽�취
                    If Left(.TextMatrix(i, COL_����), 1) = "." Then
                        .TextMatrix(i, COL_����) = "0" & .TextMatrix(i, COL_����)
                    End If
                    If Left(.TextMatrix(i, COL_����), 1) = "." Then
                        .TextMatrix(i, COL_����) = "0" & .TextMatrix(i, COL_����)
                    End If
                    
                    '������ҽ����ʶ(����ҩƷ�����ҽ��,��ֻ����Ҫҽ��)
                    If Not bln��ҩ�÷� And InStr(",5,6,", .TextMatrix(i, COL_�������)) = 0 Then
                        If bln�ɼ����� Then '����ǰ��ȡ�Ľ��
                            If bln������ Then
                                If Not bln������ Then
                                    Set .Cell(flexcpPicture, i, COL_F����) = imgFlag.ListImages("δ����").Picture
                                Else
                                    Set .Cell(flexcpPicture, i, COL_F����) = imgFlag.ListImages("������").Picture
                                End If
                            End If
                        ElseIf Val(.TextMatrix(i, COL_������)) = 1 Then
                            If Val(.TextMatrix(i, COL_����ID)) = 0 Then
                                Set .Cell(flexcpPicture, i, COL_F����) = imgFlag.ListImages("δ����").Picture
                            Else
                                Set .Cell(flexcpPicture, i, COL_F����) = imgFlag.ListImages("������").Picture
                            End If
                        End If
                    End If
                    
                    'ҽ����ɫ
                    If Val(.TextMatrix(i, COL_ҽ��״̬)) = 4 Then
                        '������(���ͺ�����)
                        .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = &H808080 '��ɫ
                    ElseIf Val(.TextMatrix(i, COL_ҽ��״̬)) = 8 Then
                        '�ѷ���(���ͺ��Զ�ֹͣ)
                        .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = &HC00000 '����
                        If lngTop = 0 Then lngTop = i
                    Else
                        If lngTop = 0 Then lngTop = i
                    End If
                    
                    '���龫ҩƷ��ʶ:��ҩ�䷽�����ζ��ҩ������
                    If .TextMatrix(i, COL_�������) <> "" Then
                        If InStr(",����ҩ,����ҩ,����ҩ,", .TextMatrix(i, COL_�������)) > 0 Then
                            .Cell(flexcpFontBold, i, COL_ҽ������) = True
                        End If
                    End If
                    
                    'Ƥ�Խ����ʶ
                    If .TextMatrix(i, COL_Ƥ��) = "(+)" Then
                        .Cell(flexcpForeColor, i, COL_Ƥ��) = vbRed
                    ElseIf .TextMatrix(i, COL_Ƥ��) = "(-)" Then
                        .Cell(flexcpForeColor, i, COL_Ƥ��) = vbBlue
                    End If
                    
                    '������־:һ����ҩֻ��ʾ�ڵ�һ��
                    blnFirst = True
                    If InStr(",5,6,", .TextMatrix(i, COL_�������)) > 0 Then
                        If Val(.TextMatrix(i, COL_���ID)) = Val(.TextMatrix(i - 1, COL_���ID)) Then
                            blnFirst = False
                        End If
                    End If
                    If blnFirst Then
                        If Val(.TextMatrix(i, COL_��־)) = 1 Then
                            Set .Cell(flexcpPicture, i, COL_F��־) = imgFlag.ListImages("����").Picture
                        ElseIf Val(.TextMatrix(i, COL_��־)) = 2 Then
                            Set .Cell(flexcpPicture, i, COL_F��־) = imgFlag.ListImages("��¼").Picture
                        End If
                    End If
                    
                    'Pass:�����������ʾ��ʾ��
                    If .TextMatrix(i, COL_��ʾ) <> "" Then
                        Set .Cell(flexcpPicture, i, COL_��ʾ) = imgPass.ListImages(Val(.TextMatrix(i, COL_��ʾ)) + 1).Picture
                        .TextMatrix(i, COL_��ʾ) = ""
                    End If
                    
                    '����ǩ����ʶ
                    If Val(.TextMatrix(i, COL_ǩ����)) = 1 Then
                        Set .Cell(flexcpPicture, i, COL_ҽ������) = imgSign.ListImages(1).Picture
                    End If
                End If
                
                If bln��ҩ;�� Then
                    .RemoveItem i
                Else
                    i = i + 1
                End If
            Loop
            
            '�̶���ͼ�����:����Ϊ�ж���,��Ȼ���߿�ʱ����������
            .Cell(flexcpPictureAlignment, .FixedRows, 0, .Rows - 1, .FixedCols - 1) = 4
            '����ǩ��ͼ�����
            .Cell(flexcpPictureAlignment, .FixedRows, COL_ҽ������, .Rows - 1, COL_ҽ������) = 0
            .Redraw = True
        End With
    Else
        Call ClearAdviceData
        Call ClearAppendData
    End If
        
    'ȱʡ��λ
    If lngҽ��ID <> 0 Then
        lngҽ��ID = vsAdvice.FindRow(CStr(lngҽ��ID), , COL_ID)
        If lngҽ��ID <> -1 Then vsAdvice.Row = lngҽ��ID
    End If
    If lngҽ��ID = -1 Or lngҽ��ID = 0 Then
        If lngTop <> 0 Then
            vsAdvice.Row = lngTop
            vsAdvice.TopRow = lngTop
        Else
            vsAdvice.Row = vsAdvice.FixedRows
        End If
    End If
    vsAdvice.Col = vsAdvice.FixedCols
    Call vsAdvice.ShowCell(vsAdvice.Row, vsAdvice.Col)
    Call vsAdvice_AfterRowColChange(-1, -1, vsAdvice.Row, vsAdvice.Col)
    vsAdvice.Refresh
    Screen.MousePointer = 0
    LoadAdvice = True
    Exit Function
errH:
    vsAdvice.Redraw = True
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function RowIs�䷽��(ByVal lngRow As Long) As Boolean
'���ܣ��ж�ָ�����Ƿ���ҩ�䷽��
'˵����ָ����Ϊ��ʾ��,�����="E"
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
        
    On Error GoTo errH
    
    strSQL = "Select ID From ����ҽ����¼ Where Rownum=1 And �������='7' And ���ID=[1]"
    If mblnMoved Then
        strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)))
    If Not rsTmp.EOF Then RowIs�䷽�� = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function RowIs������(ByVal lngRow As Long) As Boolean
'���ܣ��ж�ָ�����Ƿ���������
'˵����ָ����Ϊ��ʾ��,�����="E"
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
        
    On Error GoTo errH
    
    strSQL = "Select ID From ����ҽ����¼ Where Rownum=1 And �������='C' And ���ID=[1]"
    If mblnMoved Then
        strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)))
    If Not rsTmp.EOF Then RowIs������ = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function ShowPrice(ByVal lngRow As Long) As Boolean
'���ܣ���ȡָ��ҽ���ļƼ�,�����ݵ�ǰ�������շѹ�ϵ���и���
    Dim rs������Ŀ As New ADODB.Recordset
    Dim rs�շ�ϸĿ As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim strҽ��IDs As String, str�շ�ϸĿIDs As String
    Dim strSQL As String, i As Long, j As Long
    Dim bln�䷽�� As Boolean, bln������ As Boolean, blnLoad As Boolean
    Dim lng���˿���ID As Long, lngִ�п���ID As Long
    Dim dblPrice As Double
    
    On Error GoTo errH
    
    With vsAppend
        .Redraw = False
        .MergeCells = flexMergeNever
        .Rows = .FixedRows
        .Rows = .FixedRows + 1
    
        If Val(vsAdvice.TextMatrix(lngRow, COL_ID)) = 0 Then
            .Redraw = True: ShowPrice = True: Exit Function
        End If
        If vsAdvice.TextMatrix(lngRow, COL_�������) = "E" Then
            bln�䷽�� = RowIs�䷽��(lngRow)
            bln������ = RowIs������(lngRow)
        End If
                                    
        blnLoad = True
        
        'ҩƷ�ļƼ�
        If InStr(",5,6,", vsAdvice.TextMatrix(lngRow, COL_�������)) > 0 Then
            '��,����ҩ:���ܰ������ҽ��,����1�������װ�ĵ���
            strSQL = "Select A.ID,A.���ID,A.���,A.�������,A.������ĿID,NULL as �걾��λ,C.ID as �շ�ϸĿID," & _
                " B.�����װ,B.���ﵥλ as ���㵥λ,1 as ����,Decode(Nvl(C.�Ƿ���,0),1,-NULL,D.�ּ�)*B.�����װ as ����," & _
                " A.ִ�п���ID,0 as ����" & _
                " From ����ҽ����¼ A,ҩƷ��� B,�շ���ĿĿ¼ C,�շѼ�Ŀ D" & _
                " Where Rownum=1 And A.ID=[1]" & _
                " And A.������ĿID=B.ҩ��ID And B.ҩƷID=C.ID And Nvl(A.ִ������,0)<>5" & _
                " And (A.�շ�ϸĿID is NULL Or A.�շ�ϸĿID=B.ҩƷID)" & _
                " And (C.����ʱ�� is NULL Or C.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                " And C.������� IN(1,3) And D.�շ�ϸĿID=C.ID" & _
                " And ((Sysdate Between D.ִ������ and D.��ֹ����) or (Sysdate>=D.ִ������ And D.��ֹ���� is NULL))"
                
                '��һ����ҩ(�����)�ĵ�һ��ҩ�в���ʾ��ҩ;���ļƼ�
                blnLoad = Val(vsAdvice.TextMatrix(lngRow - 1, COL_���ID)) <> Val(vsAdvice.TextMatrix(lngRow, COL_���ID))
        ElseIf bln�䷽�� Then
            '�в�ҩ:һ����Ӧ�й���¼����д���շ�ϸĿID
            strSQL = "Select A.ID,A.���ID,A.���,A.�������,A.������ĿID,NULL as �걾��λ,C.ID as �շ�ϸĿID," & _
                " B.�����װ,B.���ﵥλ as ���㵥λ,1 as ����,Decode(Nvl(C.�Ƿ���,0),1,-NULL,D.�ּ�)*B.�����װ as ����," & _
                " A.ִ�п���ID,0 as ����" & _
                " From ����ҽ����¼ A,ҩƷ��� B,�շ���ĿĿ¼ C,�շѼ�Ŀ D" & _
                " Where A.�������='7' And A.���ID=[1]" & _
                " And A.�շ�ϸĿID=B.ҩƷID And A.�շ�ϸĿID=C.ID And C.������� IN(1,3)" & _
                " And D.�շ�ϸĿID=C.ID And Nvl(A.ִ������,0)<>5" & _
                " And (C.����ʱ�� is NULL Or C.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                " And ((Sysdate Between D.ִ������ and D.��ֹ����) or (Sysdate>=D.ִ������ And D.��ֹ���� is NULL))"
        End If
        
        '��ȡ���мƼ�(ȡ���¼۸�)����ҩƷ��ļƼ�,�������ҽ���Ƽ�
        '���Ƽ�,�ֹ��Ƽ۵�ҽ������ȡ
        '��Union��ʽ������������
        If blnLoad Then
            '�����¿���ҽ�������ݲ���ҽ���Ƽ���ȡ
            strSQL = strSQL & IIF(strSQL = "", "", " Union ALL") & _
                " Select A.ID,A.���ID,A.���,A.�������,A.������ĿID,A.�걾��λ," & _
                " B.�շ�ϸĿID,1 as �����װ,C.���㵥λ,B.����,Decode(C.�Ƿ���,1,B.����,Sum(D.�ּ�)) as ����," & _
                " Nvl(B.ִ�п���ID,A.ִ�п���ID) as ִ�п���ID,Nvl(B.����,0) as ����" & _
                " From ����ҽ����¼ A,����ҽ���Ƽ� B,�շ���ĿĿ¼ C,�շѼ�Ŀ D" & _
                " Where A.������� Not IN('5','6','7') And A.ID=B.ҽ��ID" & _
                " And Nvl(A.�Ƽ�����,0)=0 And Nvl(A.ִ������,0) Not IN(0,5) And B.�շ�ϸĿID=C.ID And B.�շ�ϸĿID=D.�շ�ϸĿID" & _
                " And ((Sysdate Between D.ִ������ and D.��ֹ����) or (Sysdate>=D.ִ������ And D.��ֹ���� is NULL))" & _
                " And (A.ID=[1] Or A.ID=[2] Or A.���ID=[1])" & _
                " Group by A.ID,A.���ID,A.���,A.�������,A.������ĿID,A.�걾��λ,B.�շ�ϸĿID," & _
                " C.���㵥λ,B.����,C.�Ƿ���,B.����,Nvl(B.ִ�п���ID,A.ִ�п���ID),Nvl(B.����,0)"
            '�¿���ҽ�������������շѹ�ϵ��ȡ(��ҩ�����ʾΪ0)
            strSQL = strSQL & IIF(strSQL = "", "", " Union ALL") & _
                " Select A.ID,A.���ID,A.���,A.�������,A.������ĿID,A.�걾��λ,B.�շ���ĿID," & _
                " 1 as �����װ,C.���㵥λ,B.�շ����� as ����,Decode(C.�Ƿ���,1,0,Sum(D.�ּ�)) as ����," & _
                " A.ִ�п���ID,Nvl(B.������Ŀ,0) as ����" & _
                " From ����ҽ����¼ A,�����շѹ�ϵ B,�շ���ĿĿ¼ C,�շѼ�Ŀ D" & _
                " Where A.������� Not IN('5','6','7') And A.ҽ��״̬ IN(1,2) And A.������ĿID=B.������ĿID" & _
                " And Nvl(A.�Ƽ�����,0)=0 And Nvl(A.ִ������,0) Not IN(0,5) And B.�շ���ĿID=C.ID And B.�շ���ĿID=D.�շ�ϸĿID" & _
                " And ((Sysdate Between D.ִ������ and D.��ֹ����) or (Sysdate>=D.ִ������ And D.��ֹ���� is NULL))" & _
                " And (C.����ʱ�� is NULL Or C.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD')) And C.������� IN(1,3)" & _
                " And (A.ID=[1] Or A.ID=[2] Or A.���ID=[1])" & _
                " Group by A.ID,A.���ID,A.���,A.�������,A.������ĿID,A.�걾��λ,B.�շ���ĿID," & _
                " C.���㵥λ,B.�շ�����,C.�Ƿ���,A.ִ�п���ID,Nvl(B.������Ŀ,0)"
        End If
        strSQL = strSQL & " Order by ���,����"
        
        If mblnMoved Then '�Һŵ���ҽ����ͬ�����ݿ�
            strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
            strSQL = Replace(strSQL, "����ҽ���Ƽ�", "H����ҽ���Ƽ�")
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, Val(vsAdvice.TextMatrix(lngRow, COL_ID)), Val(vsAdvice.TextMatrix(lngRow, COL_���ID)))
        
        '��ʾ�Ƽ�����
        If Not rsTmp.EOF Then
            'ȷ����ʾ����
            .Rows = .FixedRows + rsTmp.RecordCount
            
            '��ȡ������Ŀ,�շ�ϸĿ��Ϣ
            For i = 1 To rsTmp.RecordCount
                strҽ��IDs = strҽ��IDs & "," & rsTmp!ID
                str�շ�ϸĿIDs = str�շ�ϸĿIDs & " Union ALL Select " & rsTmp!�շ�ϸĿID & " From Dual"
                rsTmp.MoveNext
            Next
            strҽ��IDs = Mid(strҽ��IDs, 2)
            str�շ�ϸĿIDs = Mid(str�շ�ϸĿIDs, 12)
                        
            strSQL = "Select B.ID,B.���,C.���� as �������,B.����,B.�걾��λ" & _
                " From ����ҽ����¼ A,������ĿĿ¼ B,������Ŀ��� C" & _
                " Where A.ID IN(" & strҽ��IDs & ") And A.������ĿID=B.ID And B.���=C.����"
                
            If mblnMoved Then '�Һŵ���ҽ����ͬ�����ݿ�
                strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
            End If
            Call zlDatabase.OpenRecordset(rs������Ŀ, strSQL, Me.Name) 'In
            
            strSQL = "Select A.ID,A.���,B.���� as �������,A.����," & _
                " A.����,A.���,A.����,A.��������,A.�Ƿ���" & _
                " From �շ���ĿĿ¼ A,�շ���Ŀ��� B" & _
                " Where A.���=B.���� And A.ID IN(" & str�շ�ϸĿIDs & ")"
            strSQL = "Select A.ID,A.���,A.�������,A.����,Nvl(B.����,A.����) as ����," & _
                " A.���,A.����,A.��������,A.�Ƿ���,C.��������" & _
                " From (" & strSQL & ") A,�շ���Ŀ���� B,�������� C" & _
                " Where A.ID=C.����ID(+) And A.ID=B.�շ�ϸĿID(+) And B.����(+)=1 And B.����(+)=" & IIF(gbln��Ʒ��, 3, 1)
            Call zlDatabase.OpenRecordset(rs�շ�ϸĿ, strSQL, Me.Name) 'In
            
            '��ʾÿ������
            rsTmp.MoveFirst
            For i = 1 To rsTmp.RecordCount
                rs������Ŀ.Filter = "ID=" & rsTmp!������ĿID
                rs�շ�ϸĿ.Filter = "ID=" & rsTmp!�շ�ϸĿID
                
                '�Ƽ�ҽ��
                If InStr(",5,6,7,", rsTmp!�������) > 0 Then
                    .TextMatrix(i, 0) = "ҩƷҽ��-" & rs������Ŀ!����
                ElseIf rsTmp!������� = "E" And InStr(",5,6,", vsAdvice.TextMatrix(lngRow, COL_�������)) > 0 Then
                    .TextMatrix(i, 0) = "��ҩ;��-" & rs������Ŀ!����
                ElseIf rsTmp!������� = "E" And (bln�䷽�� Or bln������) Then
                    If bln������ Then
                        .TextMatrix(i, 0) = "�ɼ�����-" & rs������Ŀ!����
                    ElseIf Not IsNull(rsTmp!���ID) Then
                        .TextMatrix(i, 0) = "��ҩ�巨-" & rs������Ŀ!����
                    Else
                        .TextMatrix(i, 0) = "��ҩ�÷�-" & rs������Ŀ!����
                    End If
                ElseIf Not IsNull(rsTmp!���ID) Then
                    If rsTmp!������� = "C" Then
                        .TextMatrix(i, 0) = "������Ŀ-" & rs������Ŀ!����
                    ElseIf rsTmp!������� = "D" Then
                        .TextMatrix(i, 0) = "��鲿λ-" & Nvl(rsTmp!�걾��λ)
                    ElseIf rsTmp!������� = "F" Then
                        .TextMatrix(i, 0) = "��������-" & rs������Ŀ!����
                    ElseIf rsTmp!������� = "G" Then
                        .TextMatrix(i, 0) = "������Ŀ-" & rs������Ŀ!����
                    End If
                Else
                    .TextMatrix(i, 0) = rs������Ŀ!������� & "ҽ��-" & rs������Ŀ!����
                End If
                
                '���
                .TextMatrix(i, 1) = rs�շ�ϸĿ!�������
                '�շ���Ŀ:���/����
                .TextMatrix(i, 2) = rs�շ�ϸĿ!����
                If Not IsNull(rs�շ�ϸĿ!����) Then
                    .TextMatrix(i, 2) = .TextMatrix(i, 2) & "(" & rs�շ�ϸĿ!���� & ")"
                End If
                If Not IsNull(rs�շ�ϸĿ!���) Then
                    .TextMatrix(i, 2) = .TextMatrix(i, 2) & " " & rs�շ�ϸĿ!���
                End If
                
                '���㵥λ:ҩ��ҩƷΪ���ﵥλ,��ҩ��ҩƷΪ�ۼ۵�λ
                .TextMatrix(i, 3) = Nvl(rsTmp!���㵥λ)
                '�Ƽ�����:ҩ��ҩƷΪ1,��ҩ��ҩƷΪ��Ӧ�ۼ���
                .TextMatrix(i, 4) = FormatEx(rsTmp!����, 5)
                
                'ִ�п���
                lngִ�п���ID = Nvl(rsTmp!ִ�п���ID, 0)
                If rs�շ�ϸĿ!��� = "4" And Nvl(rs�շ�ϸĿ!��������, 0) = 1 _
                    Or InStr(",5,6,7,", rs�շ�ϸĿ!���) > 0 And InStr(",5,6,7,", rs������Ŀ!���) = 0 Then
                    lng���˿���ID = UserInfo.����ID
                    lngִ�п���ID = Get�շ�ִ�п���ID(mlng����ID, 0, rs�շ�ϸĿ!���, rs�շ�ϸĿ!ID, 4, lng���˿���ID, 0, 1, lngִ�п���ID)
                End If
                
                '���۴���
                If InStr(",5,6,7,", rs�շ�ϸĿ!���) > 0 Then
                    If Nvl(rs�շ�ϸĿ!�Ƿ���, 0) = 1 Then
                        '��ҩƷʱ��
                        If InStr(",5,6,7,", rs������Ŀ!���) > 0 Then
                            'ҩ��ҩƷ����һ�������װ������ʱ��
                            .TextMatrix(i, 5) = CalcDrugPrice(rs�շ�ϸĿ!ID, lngִ�п���ID, Nvl(rsTmp!�����װ, 1))
                            .TextMatrix(i, 5) = Format(Val(.TextMatrix(i, 5)) * Nvl(rsTmp!�����װ, 0), "0.00000")
                        Else
                            '��ҩ��ҩƷ��������ۼ��������ۼ�ʵ��
                            .TextMatrix(i, 5) = Format(CalcDrugPrice(rs�շ�ϸĿ!ID, lngִ�п���ID, Nvl(rsTmp!����, 0)), "0.00000")
                        End If
                    Else
                        'ҩ��ҩƷΪ���ﵥ��,��ҩҩƷΪ�ۼ�
                        .TextMatrix(i, 5) = Format(Nvl(rsTmp!����), "0.00000")
                    End If
                ElseIf rs�շ�ϸĿ!��� = "4" And Nvl(rs�շ�ϸĿ!��������, 0) = 1 And Nvl(rs�շ�ϸĿ!�Ƿ���, 0) = 1 Then
                    'ʱ�����ĵĵ��ۺ�ҩƷһ������
                    .TextMatrix(i, 5) = Format(CalcDrugPrice(rs�շ�ϸĿ!ID, lngִ�п���ID, Nvl(rsTmp!����, 0)), "0.00000")
                Else
                    .TextMatrix(i, 5) = Format(Nvl(rsTmp!����), "0.00000")
                End If

                'ִ�п���
                If lngִ�п���ID <> 0 Then
                    .TextMatrix(i, 6) = Get��������(lngִ�п���ID)
                End If
                
                '��������
                .TextMatrix(i, 7) = Nvl(rs�շ�ϸĿ!��������)
                
                '������Ŀ
                .TextMatrix(i, 8) = IIF(Nvl(rsTmp!����, 0) = 0, "", "��")
                
                dblPrice = dblPrice + Format(Val(.TextMatrix(i, 4)) * Val(.TextMatrix(i, 5)), "0.00000")
                
                rsTmp.MoveNext
            Next
        End If
        
        '�ϼ���
        If .Rows > 2 Then
            .Rows = .Rows + 1
            .Cell(flexcpText, .Rows - 1, 0, .Rows - 1, 3) = "�ϼ�"
            .Cell(flexcpAlignment, .Rows - 1, 0, .Rows - 1, 3) = 4
            .Cell(flexcpText, .Rows - 1, 4, .Rows - 1, 5) = Format(dblPrice, "0.00000")
            .Cell(flexcpAlignment, .Rows - 1, 4, .Rows - 1, 5) = 7
            .MergeCells = flexMergeFree
            .MergeRow(.Rows - 1) = True
        End If
        
        .Row = 1: .Col = 0
        .Redraw = True
        Call vsAppend_AfterRowColChange(-1, -1, .Row, .Col)
    End With
    
    ShowPrice = True
    Exit Function
errH:
    vsAppend.Redraw = True
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ShowSendList(ByVal lngRow As Long) As Boolean
'���ܣ���ʾָ����ҽ���ķ��ͼ�¼
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strSub As String, i As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim strExe1 As String, strExe2 As String, strState As String
    Dim bln�䷽�� As Boolean, bln������ As Boolean
    
    On Error GoTo errH
    
    With vsAppend
        .Redraw = False
        .MergeCells = flexMergeNever
        .Rows = .FixedRows
        .Rows = .FixedRows + 1
    
        If Val(vsAdvice.TextMatrix(lngRow, COL_ID)) = 0 Then
            .Redraw = True: ShowSendList = True: Exit Function
        End If
        
        If vsAdvice.TextMatrix(lngRow, COL_�������) = "E" Then
            bln�䷽�� = RowIs�䷽��(lngRow)
            bln������ = RowIs������(lngRow)
        End If
                
        strExe1 = "Decode(Nvl(A.ִ��״̬,0),0,'δִ��',1,'��ȫִ��',2,'����ִ��')"
        strExe2 = "Decode(Nvl(B.ִ��״̬,0),0,'δִ��',1,'ִ�����',2,'�ܾ�ִ��',3,'����ִ��')"
        strState = "Decode(A.��¼����,1,Decode(A.��¼״̬,0,'�շѻ���',1,'���շ�',3,'���˷�'),2,Decode(A.��¼״̬,0,'���ʻ���',1,'�Ѽ���',3,'������'),'�ѼƷ�')"
        
        'ҩ����Ӧ��ҩƷ�Ƽ۰������װ��ʾ,��ҩ����Ӧ��ҩƷ�Ƽ۰����۵�λ��ʾ
        If InStr(",5,6,", vsAdvice.TextMatrix(lngRow, COL_�������)) > 0 Then
            If Not RowInһ����ҩ(lngRow, lngBegin, lngEnd) Then lngBegin = lngRow
            '��ҩ����:��д�˷��ͼ�¼,�������޶�Ӧ����(���Ա�ҩ,��ҽ���й��)
            strSub = "Select A.*,B.�����װ,B.���ﵥλ" & _
                " From ���˷��ü�¼ A,ҩƷ��� B" & _
                " Where A.��¼״̬ IN(0,1,3) And A.�۸񸸺� is NULL And A.�շ���� IN('5','6','7')" & _
                " And A.�շ�ϸĿID=B.ҩƷID And A.ҽ�����=[1]"
            If mblnMoved Then
                strSub = Replace(strSub, "���˷��ü�¼", "H���˷��ü�¼")
            ElseIf MovedByDate(mvRegDate) Then
                strSub = strSub & " Union ALL " & Replace(strSub, "���˷��ü�¼", "H���˷��ü�¼")
            End If
            
            strSQL = _
                " Select C.���ID,C.�걾��λ,B.����ʱ��,B.NO,B.��¼����,A.�շ�ϸĿID," & _
                " Nvl(A.���ﵥλ,D.���ﵥλ) as ��λ," & _
                " Nvl(A.����/Nvl(A.�����װ,1),B.��������/Nvl(D.����ϵ��,1)/Nvl(D.�����װ,1)) as ��������," & _
                " Nvl(A.ִ�в���ID,B.ִ�в���ID) as ִ�в���ID," & _
                " Decode(Nvl(Instr(',4,5,6,7,',A.�շ����),0),0," & strExe2 & "," & strExe1 & ") as ִ��״̬,B.�״�ʱ��,B.ĩ��ʱ��," & _
                " Decode(Nvl(B.�Ʒ�״̬,0),-1,'����Ʒ�',0,'δ�Ʒ�',1," & strState & ") as �Ʒ�״̬," & _
                " B.������,B.���ͺ�,B.��¼��� as �������,A.��� as �������,C.������ĿID,C.�������" & _
                " From (" & strSub & ") A,����ҽ������ B,����ҽ����¼ C,ҩƷ��� D" & _
                " Where B.ҽ��ID=C.ID And C.�շ�ϸĿID=D.ҩƷID And C.ID=[1]" & _
                " And A.NO(+)=B.NO And A.��¼����(+)=B.��¼���� And A.ҽ�����(+)=B.ҽ��ID"
            
            '��һ����ҩ�����в���ʾ��ҩ;���ķ���
            If lngRow = lngBegin Then
                '��ҩ;������:��д�˷��ͼ�¼(������),����һ���з���
                strSub = "Select A.*,B.�����װ,B.���ﵥλ" & _
                    " From ���˷��ü�¼ A,ҩƷ��� B" & _
                    " Where A.��¼״̬ IN(0,1,3) And A.�۸񸸺� is NULL" & _
                    " And A.�շ�ϸĿID=B.ҩƷID(+) And A.ҽ�����=[2]"
                If mblnMoved Then
                    strSub = Replace(strSub, "���˷��ü�¼", "H���˷��ü�¼")
                ElseIf MovedByDate(mvRegDate) Then
                    strSub = strSub & " Union ALL " & Replace(strSub, "���˷��ü�¼", "H���˷��ü�¼")
                End If
                    
                strSQL = strSQL & " Union ALL " & _
                    " Select C.���ID,C.�걾��λ,B.����ʱ��,B.NO,B.��¼����,A.�շ�ϸĿID," & _
                    " Decode(Nvl(Instr('567',A.�շ����),0),0,D.���㵥λ,Nvl(A.���ﵥλ,E.���ﵥλ)) as ��λ," & _
                    " Decode(Nvl(Instr('567',A.�շ����),0),0,B.��������," & _
                    "   Nvl(A.����/Nvl(A.�����װ,1),B.��������/Nvl(E.����ϵ��,1)/Nvl(E.�����װ,1))) as ��������," & _
                    " Nvl(A.ִ�в���ID,B.ִ�в���ID) as ִ�в���ID," & _
                    " Decode(Nvl(Instr(',4,5,6,7,',A.�շ����),0),0," & strExe2 & "," & strExe1 & ") as ִ��״̬,B.�״�ʱ��," & _
                    " B.ĩ��ʱ��,Decode(Nvl(B.�Ʒ�״̬,0),-1,'����Ʒ�',0,'δ�Ʒ�',1," & strState & ") as �Ʒ�״̬," & _
                    " B.������,B.���ͺ�,B.��¼��� as �������,A.��� as �������,C.������ĿID,C.�������" & _
                    " From (" & strSub & ") A,����ҽ������ B,����ҽ����¼ C,������ĿĿ¼ D,ҩƷ��� E" & _
                    " Where B.ҽ��ID=C.ID And C.������ĿID=D.ID And C.�շ�ϸĿID=E.ҩƷID(+)" & _
                    " And A.NO(+)=B.NO And A.��¼����(+)=B.��¼���� And 0+A.ҽ�����(+)=B.ҽ��ID And C.ID=[2]"
            End If
            
            If mblnMoved Then
                strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
                strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
            End If
        Else
            '����ҽ��(�����䷽����飬����һ��ҽ��):��д�˷��ͼ�¼(������),����һ���з���
            '��ҩ�Ա�ҩҲ���޶�Ӧ����(��ҽ���й��)
            strSub = _
                " Select A.*,B.�����װ,B.���ﵥλ" & _
                " From ���˷��ü�¼ A,ҩƷ��� B" & _
                " Where A.��¼״̬ IN(0,1,3) And A.�۸񸸺� is NULL" & _
                " And A.�շ�ϸĿID=B.ҩƷID(+) And A.ҽ�����=[1]"
            strSub = strSub & " Union ALL " & _
                " Select A.*,B.�����װ,B.���ﵥλ" & _
                " From ���˷��ü�¼ A,ҩƷ��� B,����ҽ����¼ C" & _
                " Where A.��¼״̬ IN(0,1,3) And A.�۸񸸺� is NULL" & _
                " And A.�շ�ϸĿID=B.ҩƷID(+) And A.ҽ�����=C.ID" & _
                " And C.���ID=[1]"
            If mblnMoved Then
                strSub = Replace(strSub, "���˷��ü�¼", "H���˷��ü�¼")
            ElseIf MovedByDate(mvRegDate) Then
                strSub = strSub & " Union ALL " & Replace(strSub, "���˷��ü�¼", "H���˷��ü�¼")
            End If
            
            strSQL = _
                " Select * From ����ҽ����¼ Where ID=[1]" & _
                " Union ALL " & _
                " Select * From ����ҽ����¼ Where ���ID=[1]"
            strSQL = _
                " Select C.���ID,C.�걾��λ,B.����ʱ��,B.NO,B.��¼����,A.�շ�ϸĿID," & _
                " Decode(Nvl(Instr('567',A.�շ����),0),0,D.���㵥λ,Nvl(A.���ﵥλ,E.���ﵥλ)) as ��λ," & _
                " Decode(Nvl(Instr('567',A.�շ����),0),0,B.��������," & _
                "   Nvl(Nvl(A.����,1)*A.����/Nvl(A.�����װ,1),B.��������/Nvl(E.����ϵ��,1)/Nvl(E.�����װ,1))) as ��������," & _
                " Nvl(A.ִ�в���ID,B.ִ�в���ID) as ִ�в���ID," & _
                " Decode(Nvl(Instr(',4,5,6,7,',A.�շ����),0),0," & strExe2 & "," & strExe1 & ") as ִ��״̬,B.�״�ʱ��,B.ĩ��ʱ��," & _
                " Decode(Nvl(B.�Ʒ�״̬,0),-1,'����Ʒ�',0,'δ�Ʒ�',1," & strState & ") as �Ʒ�״̬," & _
                " B.������,B.���ͺ�,B.��¼��� as �������,A.��� as �������,C.������ĿID,C.�������" & _
                " From (" & strSub & ") A,����ҽ������ B,(" & strSQL & ") C,������ĿĿ¼ D,ҩƷ��� E" & _
                " Where B.ҽ��ID=C.ID And C.������ĿID=D.ID And C.�շ�ϸĿID=E.ҩƷID(+)" & _
                " And A.NO(+)=B.NO And A.��¼����(+)=B.��¼���� And 0+A.ҽ�����(+)=B.ҽ��ID"
            If mblnMoved Then
                strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
                strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
            End If
        End If
        
        strSQL = "Select /*+ RULE */ A.�������,A.�������," & _
            " A.���ID,A.�������,F.���� as �������,D.���� as ������Ŀ,A.�걾��λ,A.����ʱ��,A.NO,A.��¼����," & _
            " Nvl(G.����,B.����)||Decode(B.����,NULL,NULL,'('||B.����||')')||Decode(B.���,NULL,NULL,' '||B.���) as �շ���Ŀ," & _
            " A.��λ,A.�������� as ����,C.���� as ִ�п���,A.ִ��״̬,A.�״�ʱ��,A.ĩ��ʱ��,A.�Ʒ�״̬,A.������,A.���ͺ�" & _
            " From (" & strSQL & ") A,�շ���ĿĿ¼ B,���ű� C,������ĿĿ¼ D,������Ŀ��� F,�շ���Ŀ���� G" & _
            " Where A.�շ�ϸĿID=B.ID(+) And A.ִ�в���ID=C.ID(+)" & _
            " And A.������ĿID=D.ID And A.�������=F.����" & _
            " And A.�շ�ϸĿID=G.�շ�ϸĿID(+) And G.����(+)=1 And G.����(+)=" & IIF(gbln��Ʒ��, 3, 1) & _
            " Order by A.���ͺ� Desc,A.�������,A.�������,A.�������"
            
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, Val(vsAdvice.TextMatrix(lngRow, COL_ID)), Val(vsAdvice.TextMatrix(lngRow, COL_���ID)))
        If Not rsTmp.EOF Then
            .Rows = rsTmp.RecordCount + 1
            For i = 1 To rsTmp.RecordCount
                .TextMatrix(i, cs���ͺ�) = Nvl(rsTmp!���ͺ�, 0)
                .TextMatrix(i, cs����ʱ��) = Format(Nvl(rsTmp!����ʱ��), "MM-dd HH:mm")
                
                '����ҽ��
                If InStr(",5,6,7,", rsTmp!�������) > 0 Then
                    .TextMatrix(i, cs����ҽ��) = "ҩƷҽ��-" & rsTmp!������Ŀ
                ElseIf rsTmp!������� = "E" And InStr(",5,6,", vsAdvice.TextMatrix(lngRow, COL_�������)) > 0 Then
                    .TextMatrix(i, cs����ҽ��) = "��ҩ;��-" & rsTmp!������Ŀ
                ElseIf rsTmp!������� = "E" And (bln�䷽�� Or bln������) Then
                    If bln������ Then
                        .TextMatrix(i, cs����ҽ��) = "�ɼ�����-" & rsTmp!������Ŀ
                    ElseIf Not IsNull(rsTmp!���ID) Then
                        .TextMatrix(i, cs����ҽ��) = "��ҩ�巨-" & rsTmp!������Ŀ
                    Else
                        .TextMatrix(i, cs����ҽ��) = "��ҩ�÷�-" & rsTmp!������Ŀ
                    End If
                ElseIf Not IsNull(rsTmp!���ID) Then
                    If rsTmp!������� = "C" Then
                        .TextMatrix(i, cs����ҽ��) = "������Ŀ-" & rsTmp!������Ŀ
                    ElseIf rsTmp!������� = "D" Then
                        .TextMatrix(i, cs����ҽ��) = "��鲿λ-" & Nvl(rsTmp!�걾��λ)
                    ElseIf rsTmp!������� = "F" Then
                        .TextMatrix(i, cs����ҽ��) = "��������-" & rsTmp!������Ŀ
                    ElseIf rsTmp!������� = "G" Then
                        .TextMatrix(i, cs����ҽ��) = "������Ŀ-" & rsTmp!������Ŀ
                    End If
                Else
                    .TextMatrix(i, cs����ҽ��) = rsTmp!������� & "ҽ��-" & rsTmp!������Ŀ
                End If
               
                .TextMatrix(i, cs���ݺ�) = Nvl(rsTmp!NO)
                .TextMatrix(i, cs�շ���Ŀ) = Nvl(rsTmp!�շ���Ŀ)
                .TextMatrix(i, cs����) = FormatEx(Nvl(rsTmp!����), 5) & Nvl(rsTmp!��λ)
                .TextMatrix(i, cs�Ʒ�״̬) = Nvl(rsTmp!�Ʒ�״̬)
                .TextMatrix(i, csִ��״̬) = Nvl(rsTmp!ִ��״̬)
                .TextMatrix(i, csִ�п���) = Nvl(rsTmp!ִ�п���)
                .TextMatrix(i, cs������) = Nvl(rsTmp!������)
                .TextMatrix(i, cs��¼����) = Nvl(rsTmp!��¼����)
                
                '���շѵĻ��۵�ͻ����ʾ
                If .TextMatrix(i, cs�Ʒ�״̬) = "�ѽɷ�" Then
                    .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = &HC00000 '����
                ElseIf .TextMatrix(i, cs�Ʒ�״̬) = "���˷�" Then
                    .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = &H808080 '��ɫ
                End If
                rsTmp.MoveNext
            Next
        End If
        
        .Row = 1: .Col = cs����ҽ��
        .Redraw = True
        Call vsAppend_AfterRowColChange(-1, -1, .Row, .Col)
    End With
    ShowSendList = True
    Exit Function
errH:
    vsAppend.Redraw = True
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ShowSignList(ByVal lngRow As Long) As Boolean
'���ܣ���ʾָ����ҽ����ǩ����¼
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strSub As String, i As Long
    Dim lngBegin As Long, lngEnd As Long
    
    On Error GoTo errH
    
    With vsAppend
        .Redraw = False
        .MergeCells = flexMergeNever
        .Rows = .FixedRows
        .Rows = .FixedRows + 1
    
        If Val(vsAdvice.TextMatrix(lngRow, COL_ID)) = 0 Then
            .Redraw = True: ShowSignList = True: Exit Function
        End If
        
        strSQL = "Select A.ǩ��ID,A.��������,B.ǩ��ʱ��,B.ǩ����," & _
            " Decode(A.��������,1,'�¿�ҽ��',4,'����ҽ��','��������') as ǩ������" & _
            " From ����ҽ��״̬ A,ҽ��ǩ����¼ B Where A.ҽ��ID=[1] And A.ǩ��ID=B.ID Order by B.ǩ��ʱ��"
        If mblnMoved Then
            strSQL = Replace(strSQL, "����ҽ��״̬", "H����ҽ��״̬")
            strSQL = Replace(strSQL, "ҽ��ǩ����¼", "Hҽ��ǩ����¼")
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, Val(vsAdvice.TextMatrix(lngRow, COL_ID)))
        If Not rsTmp.EOF Then
            .Rows = rsTmp.RecordCount + 1
            For i = 1 To rsTmp.RecordCount
                .RowData(i) = Val(rsTmp!ǩ��ID)
                .TextMatrix(i, 0) = rsTmp!ǩ������
                .Cell(flexcpData, i, 0) = Val(rsTmp!��������)
                .TextMatrix(i, 1) = Format(rsTmp!ǩ��ʱ��, "yyyy-MM-dd HH:mm:ss")
                .TextMatrix(i, 2) = rsTmp!ǩ����
                Set .Cell(flexcpPicture, i, 0) = imgSign.ListImages(1).Picture
                rsTmp.MoveNext
            Next
        End If
        .Cell(flexcpPictureAlignment, .FixedRows, 0, .Rows - 1, 0) = 0
        .Row = 1
        .Redraw = True
        Call vsAppend_AfterRowColChange(-1, -1, .Row, .Col)
    End With
    ShowSignList = True
    Exit Function
errH:
    vsAppend.Redraw = True
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ShowBillList() As Boolean
'���ܣ���ʾָ���е�ҽ�����Ϳ��Դ�ӡ�����Ƶ����ڲ˵���
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim objMenu As Menu, lngҽ��ID As Long
    
    For i = mfrmParent.mnuReportClinic.UBound To 0 Step -1
        mfrmParent.mnuReportClinic(i).Tag = ""
        If i = 0 Then
            mfrmParent.mnuReportClinic(i).Caption = "<�޿��õ���>"
        Else
            Unload mfrmParent.mnuReportClinic(i)
        End If
    Next
    
    If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)) = 0 Then
        ShowBillList = True: Exit Function
    ElseIf tabAppend.SelectedItem.Index <> 2 Then
        ShowBillList = True: Exit Function
    ElseIf vsAppend.TextMatrix(vsAppend.Row, cs���ͺ�) = "" Then
        ShowBillList = True: Exit Function
    End If
        
    On Error GoTo errH
        
    If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_���ID)) <> 0 Then
        lngҽ��ID = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_���ID))
    Else
        lngҽ��ID = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID))
    End If
    
    With vsAppend
        strSQL = "Select Distinct D.���,D.����,D.˵��" & _
            " From ����ҽ������ A,����ҽ����¼ B,���Ƶ���Ӧ�� C,�����ļ�Ŀ¼ D" & _
            " Where A.���ͺ�=[1] And A.NO=[2]" & _
            " And A.ҽ��ID=B.ID And B.������ĿID=C.������ĿID" & _
            " And C.Ӧ�ó���=1 And C.�����ļ�ID=D.ID And D.����=5" & _
            " And D.ǰ�� IN(1,3) And D.��д IN(1,2)" & _
            " And (B.ID=[3] Or B.���ID=[3])" & _
            " Order by D.���"
        If mblnMoved Then
            strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
            strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, Val(.TextMatrix(.Row, cs���ͺ�)), .TextMatrix(.Row, cs���ݺ�), lngҽ��ID)
    End With
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            If i > 1 Then Load mfrmParent.mnuReportClinic(mfrmParent.mnuReportClinic.UBound + 1)
            Set objMenu = mfrmParent.mnuReportClinic(mfrmParent.mnuReportClinic.UBound)
            objMenu.Caption = rsTmp!����
            If i <= 10 Then
                objMenu.Caption = objMenu.Caption & "(&" & i - 1 & ")"
            ElseIf i <= 36 Then
                objMenu.Caption = objMenu.Caption & "(&" & Chr(i - 11 + Asc("A")) & ")"
            End If
            objMenu.Tag = "ZLCISBILL" & Format(rsTmp!���, "00000") & "-1" '��Ӧ���Զ��屨����
            'If i > 1 Then objMenu.Enabled = False 'һ����Ŀֻ������һ�����Ƶ���
            rsTmp.MoveNext
        Next
    End If
    
    ShowBillList = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub vsAppend_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewRow = OldRow Then Exit Sub
    If NewCol >= vsAppend.FixedCols And NewRow >= vsAppend.FixedRows Then
        If vsAppend.Redraw <> flexRDNone Then
            'Ϊ���ⲿϵͳ�������ӣ�By����ͮ��
            On Error Resume Next
            If mfrmParent.mnuReportClinic.UBound < 0 Then Exit Sub
            On Error GoTo 0
        
            Call ShowBillList '��ʾ�ɴ�ӡ�����Ƶ���
        End If
    End If
End Sub

Private Sub vsAppend_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Ϊ���ⲿϵͳ�������ӣ�By����ͮ��
    On Error Resume Next
    
    With vsAppend
        If Button = 2 And tabAppend.SelectedItem.Index = 2 Then
            If Between(.MouseRow, .FixedRows, .Rows - 1) Then
                If Between(.MouseCol, .FixedCols, .Cols - 1) Then
                    If mfrmParent.mnuReportItem(mnu��ӡ���Ƶ���).Enabled Then
                        PopupMenu mfrmParent.mnuReportItem(mnu��ӡ���Ƶ���), 2
                    End If
                End If
            End If
        End If
    End With
End Sub

Private Sub vsAppend_GotFocus()
    vsAppend.BackColorSel = &HFFCC99
End Sub

Private Sub vsAppend_LostFocus()
    vsAppend.BackColorSel = &HFFEBD7
End Sub

Private Sub vsAdvice_GotFocus()
    vsAdvice.BackColorSel = &HFFCC99
End Sub

Private Sub vsAdvice_LostFocus()
    vsAdvice.BackColorSel = &HFFEBD7
End Sub

Private Function RowInһ����ҩ(ByVal lngRow As Long, lngBegin As Long, lngEnd As Long) As Boolean
'���ܣ��ж�ָ�����Ƿ���һ����ҩ�ķ�Χ��,�����,ͬʱ�����кŷ�Χ
    Dim i As Long, blnTmp As Boolean
    With vsAdvice
        If .TextMatrix(lngRow, COL_�������) = "" Then Exit Function
        If InStr(",5,6,", .TextMatrix(lngRow, COL_�������)) = 0 Then Exit Function
        If Val(.TextMatrix(lngRow - 1, COL_���ID)) = Val(.TextMatrix(lngRow, COL_���ID)) Then
            blnTmp = True
        ElseIf lngRow + 1 <= .Rows - 1 Then
            If Val(.TextMatrix(lngRow + 1, COL_���ID)) = Val(.TextMatrix(lngRow, COL_���ID)) Then
                blnTmp = True
            End If
        End If
        If blnTmp Then
            lngBegin = lngRow
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, COL_���ID)) = Val(.TextMatrix(lngRow, COL_���ID)) Then
                    lngBegin = i
                Else
                    Exit For
                End If
            Next
            lngEnd = lngRow
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, COL_���ID)) = Val(.TextMatrix(lngRow, COL_���ID)) Then
                    lngEnd = i
                Else
                    Exit For
                End If
            Next
        End If
        RowInһ����ҩ = blnTmp
    End With
End Function

Private Sub ShowTotalMoney()
'���ܣ�ҽ���ܽ�����ʾ
'˵��������ҩƷʱ�ۣ��͸�ҩ;������ҩ�巨�÷��ȣ��¿�ҽ����һ��׼ȷ
    Dim rsMoney As New ADODB.Recordset, strSQL As String
    Dim curӦ�� As Currency, curʵ�� As Currency
    Dim curҩƷӦ�� As Currency, curҩƷʵ�� As Currency
    Dim cur�¿� As Currency, curҩƷ�¿� As Currency
    Dim curԤ�� As Currency
    
    On Error GoTo errH
    
    strSQL = _
        " Select /*+ RULE */ Sum(A.Ӧ�ս��) as Ӧ�ս��,Sum(A.ʵ�ս��) as ʵ�ս��," & _
        " Sum(Decode(Instr('567',A.�շ����),0,0,A.Ӧ�ս��)) as ҩƷӦ��," & _
        " Sum(Decode(Instr('567',A.�շ����),0,0,A.ʵ�ս��)) as ҩƷʵ��" & _
        " From ���˷��ü�¼ A,����ҽ������ B,����ҽ����¼ C" & _
        " Where A.ҽ�����=B.ҽ��ID And B.ҽ��ID=C.ID" & _
        " And C.����ID+0=[1] And C.�Һŵ�=[2]"
    If mblnMoved Then
        strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
        strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
        strSQL = Replace(strSQL, "���˷��ü�¼", "H���˷��ü�¼")
    ElseIf MovedByDate(mvRegDate) Then
        strSQL = strSQL & " Union ALL " & Replace(strSQL, "���˷��ü�¼", "H���˷��ü�¼")
        strSQL = "Select Sum(Ӧ�ս��) as Ӧ�ս��,Sum(ʵ�ս��) as ʵ�ս��," & _
            " Sum(ҩƷӦ��) as ҩƷӦ��,Sum(ҩƷʵ��) as ҩƷʵ�� From (" & strSQL & ")"
    End If
    Set rsMoney = zlDatabase.OpenSQLRecord(strSQL, Me.Name, mlng����ID, mstr�Һŵ�)
    If Not rsMoney.EOF Then
        curӦ�� = Nvl(rsMoney!Ӧ�ս��, 0)
        curʵ�� = Nvl(rsMoney!ʵ�ս��, 0)
        curҩƷӦ�� = Nvl(rsMoney!ҩƷӦ��, 0)
        curҩƷʵ�� = Nvl(rsMoney!ҩƷʵ��, 0)
    End If
        
    'ʱ��ҩƷȡ"ָ�����ۼ�"
    strSQL = _
        "Select Sum(Round(���," & gbytDec & ")) As ���,Sum(Round(ҩƷ���," & gbytDec & ")) As ҩƷ���" & _
        " From (Select A.�ܸ�����*Decode(I.�Ƿ���,1,S.ָ�����ۼ�,P.�ּ�) As ���," & _
        "              A.�ܸ�����*Decode(I.�Ƿ���,1,S.ָ�����ۼ�,P.�ּ�) As ҩƷ���" & _
        "       From ����ҽ����¼ A,�շ���ĿĿ¼ I,�շѼ�Ŀ P,ҩƷ��� S" & _
        "       Where A.�շ�ϸĿID=I.ID And I.ID=P.�շ�ϸĿID And I.ID=S.ҩƷID" & _
        "             And (Sysdate Between P.ִ������ And P.��ֹ���� Or Sysdate>=P.ִ������ And P.��ֹ���� is Null)" & _
        "             And A.ҽ��״̬=1 And A.������� In ('5','6')" & _
        "             And A.����ID+0=[1] And A.�Һŵ�=[2]" & _
        "       Union All" & _
        "       Select A.�ܸ�����*A.��������/S.����ϵ��*Decode(I.�Ƿ���,1,S.ָ�����ۼ�,P.�ּ�) As ���," & _
        "              A.�ܸ�����*A.��������/S.����ϵ��*Decode(I.�Ƿ���,1,S.ָ�����ۼ�,P.�ּ�) As ҩƷ���" & _
        "       From ����ҽ����¼ A,�շ���ĿĿ¼ I,�շѼ�Ŀ P,ҩƷ��� S" & _
        "       Where A.�շ�ϸĿID=I.ID And I.ID=P.�շ�ϸĿID And I.ID=S.ҩƷID" & _
        "             And (Sysdate Between P.ִ������ And P.��ֹ���� Or Sysdate>=P.ִ������ And P.��ֹ���� is Null)" & _
        "             And A.ҽ��״̬=1 And A.�������='7'" & _
        "             And A.����ID+0=[1] And A.�Һŵ�=[2]" & _
        "       Union All" & _
        "       Select Nvl(A.�ܸ�����,A.Ƶ�ʴ���)*R.�շ�����*P.�ּ� As ���,0 as ҩƷ���" & _
        "       From ����ҽ����¼ A,�����շѹ�ϵ R,�շ���ĿĿ¼ I,�շѼ�Ŀ P" & _
        "       Where A.������ĿID=R.������ĿID And I.ID=R.�շ���ĿID And I.ID=P.�շ�ϸĿID" & _
        "             And (Sysdate Between P.ִ������ And P.��ֹ���� Or Sysdate>=P.ִ������ And P.��ֹ���� is Null)" & _
        "             And Nvl(A.�Ƽ�����,0)=0 And A.ҽ��״̬=1 And A.������� Not In ('5','6','7')" & _
        "             And A.����ID+0=[1] And A.�Һŵ�=[2]) A"
    Set rsMoney = zlDatabase.OpenSQLRecord(strSQL, Me.Name, mlng����ID, mstr�Һŵ�)
    If Not rsMoney.EOF Then
        cur�¿� = Nvl(rsMoney!���, 0)
        curҩƷ�¿� = Nvl(rsMoney!ҩƷ���, 0)
    End If
    
    strSQL = "Select Nvl(Ԥ�����,0)-Nvl(�������,0) as ��� From ������� Where ����=1 And ����ID=[1]"
    Set rsMoney = zlDatabase.OpenSQLRecord(strSQL, Me.Name, mlng����ID)
    If Not rsMoney.EOF Then curԤ�� = Nvl(rsMoney!���, 0)
    
    mfrmParent.stbThis.Panels(2).Text = _
        "ҽ���ѷ���Ӧ��:" & FormatEx(curӦ��, gbytDec) & "(ҩ" & FormatEx(curҩƷӦ��, gbytDec) & ")," & _
        "ʵ��:" & FormatEx(curʵ��, gbytDec) & "(ҩ" & FormatEx(curҩƷʵ��, gbytDec) & ")" & _
        "  �¿�Լ:" & FormatEx(cur�¿�, gbytDec) & "(ҩ" & FormatEx(curҩƷ�¿�, gbytDec) & ")" & _
        IIF(curԤ�� = 0, "", "  Ԥ��:" & FormatEx(curԤ��, "0.00"))
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
