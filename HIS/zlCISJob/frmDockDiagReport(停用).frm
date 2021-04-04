VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDockDiagReport 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   4590
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7485
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4590
   ScaleWidth      =   7485
   ShowInTaskbar   =   0   'False
   Begin VSFlex8Ctl.VSFlexGrid vsBill 
      Height          =   4380
      Left            =   75
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
      FormatString    =   $"frmDockDiagReport.frx":0000
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
               Picture         =   "frmDockDiagReport.frx":009B
               Key             =   "δ��"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDockDiagReport.frx":05B5
               Key             =   "����"
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmDockDiagReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Event Activate() '���Ѽ���ʱ
Public Event RequestRefresh(ByVal RefreshNotify As Boolean) 'Ҫ��������ˢ��
Public Event StatusTextUpdate(ByVal Text As String) 'Ҫ�����������״̬������

Private mMainPrivs As String '��ģ��Ȩ��
Private mfrmParent As Object
Private mcbsMain As CommandBars
Private mint���� As Integer
Private mbln��ʿվ As Boolean
Private mlng����ID As Long
Private mvar����ID As Variant
Private mblnMoved As Boolean '����סԺ�����Ƿ���ת��

Private Enum PATI_TYPE
    'סԺ
    pt��Ժ = 0
    ptԤ�� = 1
    pt��Ժ = 2
    pt���� = 3 'ҽ��վ:�����ﲡ��(��Ժ)
    pt���� = 4 'ҽ��վ:�ѻ��ﲡ��
    '����
    pt��ֹ = 0
    pt���� = 1
End Enum
Private mint���� As PATI_TYPE

'��ŵ�ǰ���õ����б�
Private Type TYPE_Bill
    ID As Long
    ���� As String
End Type
Private marrBill() As TYPE_Bill

'�г���
Private Enum BILL_COL
    COL_F���� = 0 '��־��
    COL_F���� = 1
    COL_NO = 2 '�ɼ���
    COL_ҽ������ = 3
    COL_���� = 4
    COL_������ = 5
    COL_����ʱ�� = 6
    COL_����ʱ�� = 7
    COL_������ = 8
    COL_����ʱ�� = 9
    COL_ҽ��ID = 10 '������
    COL_������ĿID = 11
    COL_����ID = 12
    COL_��� = 13
    COL_������ = 14
    COL_����ID = 15
    COL_������ = 16
    COL_����ID = 17
    COL_��¼���� = 18
End Enum

Public Sub zlRefresh(ByVal lng����ID As Long, ByVal var����ID As Variant, ByVal int���� As Integer, Optional ByVal blnMoved As Boolean)
'���ܣ�ˢ�»���������嵥
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    mlng����ID = lng����ID: mvar����ID = var����ID
    mint���� = int����: mblnMoved = blnMoved
        
    vsBill.Rows = vsBill.FixedRows
    vsBill.Rows = vsBill.FixedRows + 1
    
    Call LoadBillList
    If mlng����ID <> 0 Then
        Call LoadReport
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Sub zlDefCommandBars(ByVal frmParent As Object, ByVal cbsMain As CommandBars, ByVal int���� As Integer, Optional ByVal bln��ʿվ As Boolean)
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl

    mint���� = int����: mbln��ʿվ = bln��ʿվ
    Set mfrmParent = frmParent
    Set mcbsMain = cbsMain
    cbsMain.Icons = frmPubIcons.imgPublic.Icons
    
    '����˵�:���ڹ���˵�(���������û��)���ļ��˵�����
    '-----------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ManagePopup)
    If objMenu Is Nothing Then
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
    End If
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "����(&E)", objMenu.Index + 1, False)
    objMenu.ID = conMenu_EditPopup
    With objMenu.CommandBar.Controls
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Edit_NewItem, "�������뵥(&A)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸����뵥(&M)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ�����뵥(&D)")
                
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Append, "��д���浥(&W)"): objControl.BeginGroup = True
        
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Audit, "�������뵥(&T)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_SendBack, "���ı��浥(&R)")
        
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Adjust, "��ӡ֪ͨ��(&I)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Compend, "��ӡ���浥(&P)")
        
        If Not mbln��ʿվ Then
            Set objControl = .Add(xtpControlButton, conMenu_Edit_MarkMap, "��Ƭ����(&V)"): objControl.BeginGroup = True
        End If
    End With

    '����������:���ļ�������˵������ť֮��ʼ����
    '-----------------------------------------------------
    Set objBar = cbsMain(2)
    For Each objControl In objBar.Controls '�����ǰ������һ��Control
        If Val(Left(objControl.ID, 1)) <> conMenu_FilePopup And Val(Left(objControl.ID, 1)) <> conMenu_ManagePopup Then
            Set objControl = objBar.Controls(objControl.Index - 1): Exit For
        End If
    Next
    With objBar.Controls
        'Set objControl = .Find(, conMenu_File_Preview) '��Ԥ����ť֮��ʼ����
        Set objPopup = .Add(xtpControlPopup, conMenu_Edit_NewItem, "����", objControl.Index + 1): objPopup.BeginGroup = True
        objPopup.ID = conMenu_Edit_NewItem
        objPopup.IconId = conMenu_Edit_NewItem
        objPopup.Style = xtpButtonIconAndCaption
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸�", objPopup.Index + 1): objControl.ToolTipText = "�޸����뵥"
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ��", objControl.Index + 1): objControl.ToolTipText = "ɾ�����뵥"
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Append, "����", objControl.Index + 1): objControl.BeginGroup = True
        If Not mbln��ʿվ Then
            Set objControl = .Add(xtpControlButton, conMenu_Edit_MarkMap, "��Ƭ", objControl.Index + 1): objControl.BeginGroup = True
        End If
    End With
    
    '����Ŀ����
    '-----------------------------------------------------
    With cbsMain.KeyBindings
        .Add FCONTROL, vbKeyA, conMenu_Edit_NewItem '�������뵥
        .Add FCONTROL, vbKeyM, conMenu_Edit_Modify '�޸����뵥
        .Add 0, vbKeyDelete, conMenu_Edit_Delete 'ɾ�����뵥
        .Add FCONTROL, vbKeyR, conMenu_Edit_Append '��д���浥
        .Add FCONTROL, vbKeyW, conMenu_Edit_MarkMap '��Ƭ����
    End With

    '���ò���������
    '-----------------------------------------------------
    With cbsMain.Options
    End With
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As CommandBarControl)
'���ܣ�����Ȩ�ޡ���ǰ���˻�������������ù��ܻ�ɼ��Ϳ�����
'  1.�޲��˵����
'  2.�����ѳ�Ժ(����)�����
'  3.�����ݵ����
    Dim blnBill As Boolean, blnEnabled As Boolean
            
    If vsBill.Redraw = flexRDNone Then Exit Sub
        
    '����Ȩ�����ð�ť�ɼ�״̬
    Call SetControlVisible(Control)
    If Not Control.Visible Then Exit Sub
    
    '�����������
    '------------------------------------------------------------------------------
    '�ܵ��ж�:�޲��˲������κβ���
    If Between(Control.ID, conMenu_Edit_NewItem, conMenu_Edit_NewItem + 999) Then
        Control.Enabled = mlng����ID <> 0
        If Not Control.Enabled Then Exit Sub
    End If
    
    '���ﲿ��
    '------------------------------------------------------------------------------
    blnBill = Val(vsBill.TextMatrix(vsBill.Row, COL_ҽ��ID)) <> 0
    blnEnabled = mint���� = 1 And mint���� = pt���� Or mint���� = 2 And (mint���� = pt��Ժ Or mint���� = pt����)
    Select Case Control.ID
        Case conMenu_Edit_NewItem
            Control.Enabled = blnEnabled And UBound(marrBill) >= 1
        Case conMenu_Edit_NewItem * 100# + 1 To conMenu_Edit_NewItem * 100# + 200 '�������뵥
            Control.Enabled = blnEnabled
        Case conMenu_Edit_Audit '�������뵥
            Control.Enabled = blnBill
        Case conMenu_Edit_Modify '�޸����뵥
            Control.Enabled = blnBill And blnEnabled
        Case conMenu_Edit_Delete 'ɾ�����뵥
            Control.Enabled = blnBill And blnEnabled
        Case conMenu_Edit_Adjust '��ӡ֪ͨ��
            Control.Enabled = blnBill And blnEnabled
        Case conMenu_Edit_SendBack '���ı��浥
            Control.Enabled = blnBill
        Case conMenu_Edit_Append '��д���浥
            Control.Enabled = blnBill And blnEnabled
        Case conMenu_Edit_Compend '��ӡ���浥
            Control.Enabled = blnBill And blnEnabled
        Case conMenu_Edit_MarkMap '��Ƭ����
            Control.Enabled = blnBill And blnEnabled
            If Control.Enabled Then
                Control.Enabled = (vsBill.Cell(flexcpData, vsBill.Row, COL_������ĿID) = "D" And vsBill.Cell(flexcpData, vsBill.Row, COL_����ID) = 1)
            End If
    End Select
    
    '��������
    '------------------------------------------------------------------------------
    Select Case Control.ID
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        Control.Enabled = blnBill
    End Select
End Sub

Private Sub SetControlVisible(ByVal Control As CommandBarControl)
'���ܣ�����Ȩ�����ò˵��͹������Ŀɼ�״̬
    Dim blnVisible As Boolean, strItem As String

    'Ȩ��ֻ���ж�һ��,�Ѿ��жϹ�����������ж�
    If Control.Category = "���ж�" Then Exit Sub

    blnVisible = True
    Select Case Control.ID
        Case conMenu_Edit_NewItem '�������뵥
            If InStr(GetInsidePrivs(p�����¼����), ";������д;") = 0 Then blnVisible = False
        Case conMenu_Edit_Audit '�������뵥
            If InStr(GetInsidePrivs(p�����¼����), ";������д;") = 0 Then blnVisible = False
        Case conMenu_Edit_Modify '�޸����뵥
            If InStr(GetInsidePrivs(p�����¼����), ";������д;") = 0 Then blnVisible = False
        Case conMenu_Edit_Delete 'ɾ�����뵥
            If InStr(GetInsidePrivs(p�����¼����), ";������д;") = 0 Then blnVisible = False
        Case conMenu_Edit_Adjust '��ӡ֪ͨ��
            If InStr(GetInsidePrivs(p�����¼����), ";������д;") = 0 Then blnVisible = False
        Case conMenu_Edit_SendBack '���ı��浥
            If InStr(GetInsidePrivs(p�����¼����), ";�������;") = 0 Then blnVisible = False
        Case conMenu_Edit_Append '��д���浥
            If InStr(GetInsidePrivs(p�����¼����), ";����༭;") = 0 Then blnVisible = False
        Case conMenu_Edit_Compend '��ӡ���浥
            If InStr(GetInsidePrivs(p�����¼����), ";�����ӡ;") = 0 Then blnVisible = False
        Case conMenu_Edit_MarkMap '��Ƭ����
            If InStr(GetInsidePrivs(p�����¼����), ";��Ƭ����;") = 0 Then blnVisible = False
    End Select
    
    Control.Visible = blnVisible
    Control.Category = "���ж�"
End Sub

Public Sub zlPopupCommandBars(ByVal CommandBar As CommandBar)
    Dim objControl As CommandBarControl
    Dim vBill As TYPE_Bill, i As Long
    
    If CommandBar.Parent Is Nothing Then Exit Sub
    
    Select Case CommandBar.Parent.ID
    Case conMenu_Edit_NewItem '�������뵥
        With CommandBar.Controls
            .DeleteAll
            For i = 1 To UBound(marrBill)
                vBill = marrBill(i)
                Set objControl = .Add(xtpControlButton, conMenu_Edit_NewItem * 100# + i, vBill.����)
                If i <= 10 Then
                    objControl.Caption = objControl.Caption & "(&" & i - 1 & ")"
                ElseIf i <= 36 Then
                    objControl.Caption = objControl.Caption & "(&" & Chr(i - 11 + Asc("A")) & ")"
                End If
                objControl.Parameter = vBill.ID
            Next
        End With
    End Select
End Sub

Public Sub zlExecuteCommandBars(ByVal Control As CommandBarControl)
    Select Case Control.ID
        Case conMenu_File_PrintSet '��ӡ����
            Call zlPrintSet
        Case conMenu_File_Preview 'Ԥ�������嵥
            Call OutputList(2)
        Case conMenu_File_Print '��ӡ�����嵥
            Call OutputList(1)
        Case conMenu_File_Excel '��������嵥
            Call OutputList(3)
        Case conMenu_Help_Help '����
            Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100))
        Case conMenu_Tool_Reference_2 '���ƴ��ϲο�
            Call zlItemRef
        '------------------------------------------------------------------------------------
        Case conMenu_Edit_NewItem * 100# + 1 To conMenu_Edit_NewItem * 100# + 200 '�������뵥
            Call FuncAddRequest(Val(Control.Parameter))
        Case conMenu_Edit_Audit '�������뵥
            Call FuncWriteRequest(True)
        Case conMenu_Edit_Modify '�޸����뵥
            Call FuncWriteRequest(False)
        Case conMenu_Edit_Delete 'ɾ�����뵥
            Call FuncDeleteRequest
        Case conMenu_Edit_Adjust '��ӡ֪ͨ��
            Call FuncPrintRequest
        Case conMenu_Edit_SendBack '���ı��浥
            Call FuncWriteReport(True)
        Case conMenu_Edit_Append '��д���浥
            Call FuncWriteReport(False)
        Case conMenu_Edit_Compend '��ӡ���浥
            Call FuncPrintReport
        Case conMenu_Edit_MarkMap '��Ƭ����
            Call ViewImage(Val(vsBill.TextMatrix(vsBill.Row, COL_ҽ��ID)), mfrmParent, mblnMoved)
    End Select
End Sub

Private Sub InitBillTable()
'���ܣ���ʼ�������嵥��ʽ
    Dim arrHead As Variant, strHead As String, i As Long
    
    strHead = "���ݺ�,810,1;ҽ������,3000,1;����,1800,1;������,850,1;" & _
        "����ʱ��,1080,1;����ʱ��,1080,1;������,850,1;����ʱ��,1080,1;" & _
        "ҽ��ID;������ĿID;����ID;���;������;����ID;������;����ID;��¼����"
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

Private Sub Form_Load()
    Call InitBillTable
    Call RestoreWinState(Me, App.ProductName)
    mMainPrivs = gstrPrivs
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    vsBill.Left = 0
    vsBill.Top = 0
    vsBill.Width = Me.ScaleWidth
    vsBill.Height = Me.ScaleHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub FuncPrintRequest()
'���ܣ���ӡ֪ͨ��
    Dim strBill As String
    
    If mlng����ID = 0 Then Exit Sub
    
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
        '���δ����������
        If Len(.TextMatrix(.Row, COL_NO)) = 0 Then
            MsgBox "�����뻹δ���ͣ����ܴ�ӡ֪ͨ����", vbInformation, gstrSysName
            Exit Sub
        End If
        
        strBill = "ZLCISBILL" & Format(.TextMatrix(.Row, COL_���), "00000") & "-1"
        If ReportPrintSet(gcnOracle, glngSys, strBill, mfrmParent) Then
            Call ReportOpen(gcnOracle, glngSys, strBill, mfrmParent, "NO=" & .TextMatrix(.Row, COL_NO), "����=" & Val(.TextMatrix(.Row, COL_��¼����)), 2)
        End If
    End With
End Sub

Private Sub FuncPrintReport(Optional ByVal PrtMode As Integer = 2)
'���ܣ���ӡ���浥
    Dim strBill As String
    
    If mlng����ID = 0 Then Exit Sub
    
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
        
'        strBill = "ZLCISBILL" & Format(.TextMatrix(.Row, COL_���), "00000") & "-2"
'        If ReportPrintSet(gcnOracle, glngSys, strBill, mfrmParent) Then
'            Call ReportOpen(gcnOracle, glngSys, strBill, mfrmParent, "NO=" & .TextMatrix(.Row, COL_NO), "����=" & Val(.TextMatrix(.Row, COL_��¼����)), 2)
'        End If
        Call PrintDiagRpt_New(Val(.TextMatrix(.Row, COL_����ID)), mfrmParent, PrtMode, , mblnMoved)
    End With
End Sub

Private Sub FuncAddRequest(ByVal lng����ID As Long)
'���ܣ��������뵥
    If mlng����ID = 0 Then Exit Sub
    If lng����ID = 0 Then Exit Sub
    
    If mblnMoved Then
        MsgBox "���˵ı��ξ��������Ѿ�ת���������ݿ⣬�����������" & vbCrLf & _
            "��������ϵͳ����Ա��ϵ������Ӧ���ݳ�ѡ���ء�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '���ýӿ�
    Call AddRequest(mfrmParent, mlng����ID, mvar����ID, lng����ID, mbln��ʿվ)
    If True Then
        Call LoadReport
    End If
End Sub

Private Sub FuncWriteRequest(ByVal blnReadOnly As Boolean)
'���ܣ���д��������
    If mlng����ID = 0 Then Exit Sub
    With vsBill
        If Val(.TextMatrix(.Row, COL_ҽ��ID)) = 0 Then Exit Sub
        If Val(.TextMatrix(.Row, COL_������)) = 0 Then
            MsgBox "�õ��ݲ���Ҫ��д���롣", vbInformation, gstrSysName
            Exit Sub
        End If
        If Not blnReadOnly Then
            If .TextMatrix(.Row, COL_NO) <> "" Then
                MsgBox "��ҽ���Ѿ����ͣ���������д���롣", vbInformation, gstrSysName
                Exit Sub
            End If
            
            If mblnMoved Then
                MsgBox "���˵ı��ξ��������Ѿ�ת���������ݿ⣬�����������" & vbCrLf & _
                    "��������ϵͳ����Ա��ϵ������Ӧ���ݳ�ѡ���ء�", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        
        '��д����:ҽ��ID,����ID,����ID,ҽ������
        Call EditRequest(Me, Val(.TextMatrix(.Row, COL_ҽ��ID)), Val(.TextMatrix(.Row, COL_����ID)), _
            Val(.TextMatrix(.Row, COL_����ID)), .TextMatrix(.Row, COL_ҽ������), blnReadOnly, , , mblnMoved)
        If True Then
            Call LoadReport
        End If
    End With
End Sub

Private Sub FuncWriteReport(ByVal blnReadOnly As Boolean)
'���ܣ���д���ﱨ��
    If mlng����ID = 0 Then Exit Sub
    
    With vsBill
        If Val(.TextMatrix(.Row, COL_ҽ��ID)) = 0 Then Exit Sub
        
        If Val(.TextMatrix(.Row, COL_������)) = 0 Then
            MsgBox "�õ��ݲ���Ҫ��д���档", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If .TextMatrix(.Row, COL_NO) = "" Then
            MsgBox "��ҽ����δ���ͣ����ȷ���ҽ����", vbInformation, gstrSysName
            Exit Sub
        End If
'        If .RowData(.Row) > 0 And .RowData(.Row) < 6 Then
        
        If blnReadOnly Then
            If .TextMatrix(.Row, COL_����ID) = 0 Then
                MsgBox "��ҽ����δ���棬���ܲ��ġ�", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        If Not blnReadOnly Then
            If .Cell(flexcpData, .Row, COL_����ID) <> 1 Then
                MsgBox "��ҽ���ı�����δ��ˣ����ܱ༭��", vbInformation, gstrSysName
                Exit Sub
            End If
            
            If mblnMoved Then
                MsgBox "���˵ı���סԺ�����Ѿ�ת���������ݿ⣬�����������" & vbCrLf & _
                    "��������ϵͳ����Ա��ϵ������Ӧ���ݳ�ѡ���ء�", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        
        '��д����:NO,��¼����,����ID,����ID,ҽ������
        Call EditReport(Me, .TextMatrix(.Row, COL_NO), Val(.TextMatrix(.Row, COL_��¼����)), _
            Val(.TextMatrix(.Row, COL_����ID)), Val(.TextMatrix(.Row, COL_����ID)), _
            .TextMatrix(.Row, COL_ҽ������), blnReadOnly Or .Cell(flexcpData, .Row, COL_����ID) = 1, lngҽ��ID:=.TextMatrix(.Row, COL_ҽ��ID))
        If True Then
            Call LoadReport
        End If
    End With
End Sub

Private Sub FuncDeleteRequest()
'���ܣ�ɾ����ǰ���뵥
    Dim strSQL As String, lngRow As Long
        
    If mlng����ID = 0 Then Exit Sub
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
        
        If mblnMoved Then
            MsgBox "���˵ı���סԺ�����Ѿ�ת���������ݿ⣬�����������" & vbCrLf & _
                "��������ϵͳ����Ա��ϵ������Ӧ���ݳ�ѡ���ء�", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If MsgBox("ȷʵҪɾ�����뵥[" & .TextMatrix(.Row, COL_����) & "]��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        
        '��У�ԶԵ��ڹ����м�飻�Լ������,ע�������ҽ��ID�����ǲɼ�������ID
        strSQL = "zl_����ҽ����¼_Delete(" & Val(.TextMatrix(.Row, COL_ҽ��ID)) & ",1)"
    End With
    
    'ɾ�����뵥
    On Error GoTo errH
    gcnOracle.BeginTrans
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
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

Public Sub zlItemRef()
'���ܣ��������Ʋο�
    Dim lng������ĿID As Long
    
    lng������ĿID = Val(vsBill.TextMatrix(vsBill.Row, COL_������ĿID))
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
    objOut.Title.Text = "���˵����嵥"
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
    Set objOut.Body = vsBill
    
    '���
    vsBill.Redraw = False
    lngRow = vsBill.Row: lngCol = vsBill.Col
    
    strWidth = ""
    For i = 0 To vsBill.FixedCols - 1
        strWidth = strWidth & "," & vsBill.ColWidth(i)
        vsBill.ColWidth(i) = 0
    Next
        
    If bytStyle = 1 Then
        bytR = zlPrintAsk(objOut)
        Me.Refresh
        If bytR <> 0 Then zlPrintOrView1Grd objOut, bytR
    Else
        zlPrintOrView1Grd objOut, bytStyle
    End If
    
    strWidth = Mid(strWidth, 2)
    For i = 0 To vsBill.FixedCols - 1
        vsBill.ColWidth(i) = Split(strWidth, ",")(i)
    Next
    
    vsBill.Row = lngRow: vsBill.Col = lngCol
    vsBill.Redraw = True
End Sub

Private Function GetPatiInfo() As String
'���ܣ���ȡ������Ϣ��(���ڴ�ӡ)
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    If mint���� = 1 Then
        'ִ�в���(�ű����)�����˿���
        strSQL = "Select B.����,B.�Ա�,B.����,B.�����," & _
            " B.����,B.��������,C.���� as ִ�в���,A.ִ�в���ID,A.�Ǽ�ʱ��" & _
            " From ���˹Һż�¼ A,������Ϣ B,���ű� C" & _
            " Where A.NO=[2] And A.����ID+0=[1]" & _
            " And A.����ID=B.����ID And A.ִ�в���ID=C.ID"
        If mblnMoved Then
            strSQL = Replace(strSQL, "���˹Һż�¼", "H���˹Һż�¼")
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mvar����ID)
        
        GetPatiInfo = _
            "������" & rsTmp!���� & " �Ա�" & Nvl(rsTmp!�Ա�) & _
            " ���䣺" & Nvl(rsTmp!����) & " ����ţ�" & Nvl(rsTmp!�����) & _
            " �Һţ�" & Format(rsTmp!�Ǽ�ʱ��, "MM-dd HH:mm") & _
            " ���ң�" & rsTmp!ִ�в��� & " ���ң�" & Nvl(rsTmp!��������)
    ElseIf mint���� = 2 Then
        strSQL = "Select B.����,B.�Ա�,B.����,B.סԺ��," & _
            " B.����,C.���� as ����,A.��Ժ����,A.��Ժ����" & _
            " From ������ҳ A,������Ϣ B,���ű� C" & _
            " Where A.��ҳID=[2] And A.����ID=[1]" & _
            " And A.����ID=B.����ID And A.��Ժ����ID=C.ID"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mvar����ID)
        
        GetPatiInfo = _
            "������" & rsTmp!���� & " �Ա�" & Nvl(rsTmp!�Ա�) & _
            " ���䣺" & Nvl(rsTmp!����) & " סԺ�ţ�" & Nvl(rsTmp!סԺ��) & _
            " ���ң�" & rsTmp!���� & " ��Ժ��" & Format(rsTmp!��Ժ����, "MM-dd HH:mm") & _
            IIf(Not IsNull(rsTmp!��Ժ����), " ��Ժ��" & Format(rsTmp!��Ժ����, "MM-dd HH:mm"), "")
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub vsBill_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Dim lngW As Long
    
    If Col = COL_ҽ������ Then
        vsBill.AutoSize Col
    ElseIf Row = -1 Then
        lngW = Me.TextWidth(vsBill.TextMatrix(vsBill.FixedRows - 1, Col) & "A")
        If vsBill.ColWidth(Col) < lngW Then
            vsBill.ColWidth(Col) = lngW
        ElseIf vsBill.ColWidth(Col) > vsBill.Width * 0.5 Then
            vsBill.ColWidth(Col) = vsBill.Width * 0.5
        End If
    End If
End Sub

Private Sub vsBill_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Row = -1 Then
        If Col <= vsBill.FixedCols - 1 Then
            Cancel = True
        End If
    End If
End Sub

Private Sub vsBill_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
    Dim vRect As RECT
    With vsBill
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
        End If
        Done = True
    End With
End Sub

Private Sub vsBill_GotFocus()
    vsBill.BackColorSel = &HFFCC99
End Sub

Private Sub vsBill_LostFocus()
    vsBill.BackColorSel = &HFFEBD7
End Sub

Private Function LoadReport() As Boolean
'���ܣ����ݵ�ǰ����ҽ����ȡ������д�����뵥�򱨸浥
    Dim rsBill As New ADODB.Recordset
    Dim strSQL As String, strBill As String
    Dim strKey As String, lngPreRow As Long
    Dim lngRow As Long, blnRemove As Boolean, i As Long
    
    If mlng����ID = 0 Then Exit Function
    
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
        " Where A.������ĿID=B.������ĿID And B.Ӧ�ó���=[3]" & _
        " And B.�����ļ�ID=C.�����ļ�ID And B.�����ļ�ID=D.ID" & _
        IIf(mint���� = 1, " And A.����ID+0=[1] And A.�Һŵ�=[2]", " And A.����ID=[1] And A.��ҳID=[2]") & _
        " And (A.������� Not IN('5','6','7') And A.���ID is NULL" & _
        "   Or A.�������='C' And A.���ID is Not NULL)" & _
        " Group by A.ID,B.�����ļ�ID,D.���,D.����,D.˵��"
    
    '��ҩƷ��ص�ҽ���Լ��ɼ�����
    strSQL = "Select Distinct ���ID From ����ҽ����¼" & _
        " Where ����ID=[1] And " & IIf(mint���� = 1, "�Һŵ�", "��ҳID") & "=[2]" & _
        " And (������� IN('5','6','7') Or �������='C' And ���ID is Not NULL)"
        
    'ҽ����Ӧ�ĵ����嵥(����������ҽ��,���������͵Ķ���),���ٰ���һ�ֵ��ݸ���
    'δ���͵�ҽ����ʾһ��,�ѷ��͵�һ�η�����ʾһ��(����ֻ���������)
    strSQL = _
        " Select A.ID,A.���ID,A.�������,A.������ĿID,A.ҽ������,A.�걾��λ," & _
        " B.����ʱ��,B.NO,B.��¼����,A.����ID,B.����ID,C.���,C.����,C.����ID,C.������,C.������," & _
        " X.��д�� as ������,X.��д���� as ����ʱ��," & _
        " Y.��д�� as ������,Y.��д���� as ����ʱ��,Nvl(B.ִ�й���,0) As ִ�й���,Nvl(B.ִ��״̬,0) As ִ��״̬" & _
        " From ����ҽ����¼ A,����ҽ������ B,(" & strBill & ") C,���˲�����¼ X,���˲�����¼ Y" & _
        " Where " & IIf(mint���� = 1, " A.����ID+0=[1] And A.�Һŵ�=[2]", " A.����ID=[1] And A.��ҳID=[2]") & _
        " And (A.������� Not IN('5','6','7') And A.���ID is NULL" & _
        "   Or A.�������='C' And A.���ID is Not NULL)" & _
        " And A.ID Not IN(" & strSQL & ") And A.ҽ��״̬<>4 And Nvl(A.ִ������,0)<>0" & _
        " And A.ID=B.ҽ��ID(+) And A.ID=C.ҽ��ID And (C.������=1 Or C.������=1)" & _
        " And A.����ID=X.ID(+) And B.����ID=Y.ID(+)" & _
        " Order by Nvl(B.����ʱ��,A.����ʱ��) Desc,A.���"
            
    If mblnMoved Then
        strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
        strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
        strSQL = Replace(strSQL, "���˲�����¼", "H���˲�����¼")
    End If
            
    'ҽ������,NO,����,������,����ʱ��,����ʱ��,������,����ʱ��
    'ҽ��ID;������ĿID;����ID;���;������;����ID;������;����ID;��¼����
    Set rsBill = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mvar����ID, mint����)
    With vsBill
        .Rows = .FixedRows + rsBill.RecordCount
        lngRow = .FixedRows
        For i = 1 To rsBill.RecordCount
            .TextMatrix(lngRow, COL_ҽ������) = rsBill!ҽ������
                .Cell(flexcpData, lngRow, COL_ҽ������) = Nvl(rsBill!�걾��λ) '����걾
            .TextMatrix(lngRow, COL_NO) = Nvl(rsBill!NO)
            .TextMatrix(lngRow, COL_����) = rsBill!����
            .TextMatrix(lngRow, COL_������) = Nvl(rsBill!������)
            .TextMatrix(lngRow, COL_����ʱ��) = Format(Nvl(rsBill!����ʱ��), "MM-dd HH:mm")
            .TextMatrix(lngRow, COL_����ʱ��) = Format(Nvl(rsBill!����ʱ��), "MM-dd HH:mm")
            .TextMatrix(lngRow, COL_������) = Nvl(rsBill!������)
            .TextMatrix(lngRow, COL_����ʱ��) = Format(Nvl(rsBill!����ʱ��), "MM-dd HH:mm")
            .TextMatrix(lngRow, COL_ҽ��ID) = rsBill!ID '������
            .TextMatrix(lngRow, COL_������ĿID) = rsBill!������ĿID
                .Cell(flexcpData, lngRow, COL_������ĿID) = Nvl(rsBill!�������)
            .TextMatrix(lngRow, COL_����ID) = rsBill!����ID
            .TextMatrix(lngRow, COL_���) = rsBill!���
            .TextMatrix(lngRow, COL_������) = Nvl(rsBill!������, 0)
            .TextMatrix(lngRow, COL_����ID) = Nvl(rsBill!����ID, 0)
            .TextMatrix(lngRow, COL_������) = Nvl(rsBill!������, 0)
            .TextMatrix(lngRow, COL_����ID) = Nvl(rsBill!����ID, 0)
                .Cell(flexcpData, lngRow, COL_����ID) = Nvl(rsBill!ִ��״̬, 0)
            .TextMatrix(lngRow, COL_��¼����) = Nvl(rsBill!��¼����)
            .RowData(lngRow) = Nvl(rsBill!ִ�й���, 0)
            '����˵ı��棬����Ӵ�
            .Cell(flexcpFontBold, lngRow, 1, lngRow, .Cols - 1) = Not (.Cell(flexcpData, lngRow, COL_����ID) <> 1 Or _
                .TextMatrix(lngRow, COL_����ID) = 0)
            
            '�����뱨��ı�ʶ
            If rsBill!������ = 1 Then
                If Not IsNull(rsBill!����ID) Then
                    Set .Cell(flexcpPicture, lngRow, COL_F����) = imgFlag.ListImages("����").Picture
                Else
                    Set .Cell(flexcpPicture, lngRow, COL_F����) = imgFlag.ListImages("δ��").Picture
                End If
            End If
            If rsBill!������ = 1 Then
                If Not IsNull(rsBill!����ID) Then
                    Set .Cell(flexcpPicture, lngRow, COL_F����) = imgFlag.ListImages("����").Picture
                Else
                    Set .Cell(flexcpPicture, lngRow, COL_F����) = imgFlag.ListImages("δ��").Picture
                End If
            End If
            
            'ɾ������һ���ɼ��ļ�����Ŀ��
            blnRemove = False
            If rsBill!������� = "C" And Not IsNull(rsBill!���ID) Then
                .TextMatrix(lngRow, COL_ҽ��ID) = rsBill!���ID 'һ���ɼ��ļ�¼Ϊ���ID
                If Val(.TextMatrix(lngRow - 1, COL_ҽ��ID)) = rsBill!���ID Then
                    '���ҽ������
                    .TextMatrix(lngRow - 1, COL_ҽ������) = Replace(.TextMatrix(lngRow - 1, COL_ҽ������), "(" & .Cell(flexcpData, lngRow - 1, COL_ҽ������) & ")", "")
                    .TextMatrix(lngRow - 1, COL_ҽ������) = .TextMatrix(lngRow - 1, COL_ҽ������) & "," & .TextMatrix(lngRow, COL_ҽ������)
                    If .Cell(flexcpData, lngRow - 1, COL_ҽ������) <> "" Then
                        .TextMatrix(lngRow - 1, COL_ҽ������) = .TextMatrix(lngRow - 1, COL_ҽ������) & "(" & .Cell(flexcpData, lngRow - 1, COL_ҽ������) & ")"
                    End If
                    'ɾ������
                    .RemoveItem lngRow
                    blnRemove = True
                End If
            End If
                        
            If Not blnRemove Then
                '��λ����ǰ��
                If Val(.TextMatrix(lngRow, COL_ҽ��ID)) & "_" & .TextMatrix(lngRow, COL_NO) = strKey Then
                    lngPreRow = lngRow
                End If
                lngRow = lngRow + 1
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
    Dim objPopup As CommandBarPopup
    If Button = 2 Then
        Set objPopup = mcbsMain.ActiveMenuBar.FindControl(, conMenu_EditPopup)
        If Not objPopup Is Nothing Then
            objPopup.CommandBar.ShowPopup
        End If
    End If
End Sub

Private Function LoadBillList() As Boolean
'���ܣ���ȡ��ǰ���õĸ��ﵥ���嵥
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim vBill As TYPE_Bill
    
    On Error GoTo errH
    
    ReDim marrBill(0)
    
    '���ؿ��õ���
    strSQL = "Select Distinct A.ID,A.���,A.����,A.˵��" & _
        " From �����ļ�Ŀ¼ A,�����ļ���� B" & _
        " Where A.����=5 And A.ǰ�� IN([1],3)" & _
        " And A.ID=B.�����ļ�ID And B.��дʱ�� IN(1,2)" & _
        " Order by A.���"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mint����)
    If Not rsTmp.EOF Then
        ReDim marrBill(rsTmp.RecordCount)
        For i = 1 To rsTmp.RecordCount
            With vBill
                .ID = rsTmp!ID
                .���� = rsTmp!����
            End With
            marrBill(i) = vBill '��0�ĸ�����
            rsTmp.MoveNext
        Next
    End If
    LoadBillList = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
