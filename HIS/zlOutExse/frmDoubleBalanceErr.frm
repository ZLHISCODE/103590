VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Begin VB.Form frmDoubleBalanceErr 
   BorderStyle     =   0  'None
   Caption         =   "frmDoubleBalanceErr"
   ClientHeight    =   7500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11460
   LinkTopic       =   "Form1"
   ScaleHeight     =   7500
   ScaleWidth      =   11460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picBalance 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2805
      Left            =   6570
      ScaleHeight     =   2805
      ScaleWidth      =   2715
      TabIndex        =   4
      Top             =   4245
      Width           =   2715
      Begin VSFlex8Ctl.VSFlexGrid vsfBalance 
         Height          =   1845
         Left            =   300
         TabIndex        =   5
         Top             =   330
         Width           =   1800
         _cx             =   3175
         _cy             =   3254
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
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483633
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
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
   Begin VB.PictureBox picDetail 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2805
      Left            =   885
      ScaleHeight     =   2805
      ScaleWidth      =   2715
      TabIndex        =   2
      Top             =   4515
      Width           =   2715
      Begin VSFlex8Ctl.VSFlexGrid vsfDetail 
         Height          =   1845
         Left            =   300
         TabIndex        =   3
         Top             =   330
         Width           =   1800
         _cx             =   3175
         _cy             =   3254
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
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483633
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
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
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2640
      Left            =   4425
      ScaleHeight     =   2640
      ScaleWidth      =   3120
      TabIndex        =   0
      Top             =   510
      Width           =   3120
      Begin VSFlex8Ctl.VSFlexGrid vsfMain 
         Height          =   1830
         Left            =   510
         TabIndex        =   1
         Top             =   270
         Width           =   1800
         _cx             =   3175
         _cy             =   3228
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
         BackColorSel    =   12632256
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483633
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmDoubleBalanceErr.frx":0000
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
   Begin XtremeDockingPane.DockingPane dkpMain 
      Left            =   975
      Top             =   540
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmDoubleBalanceErr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnNOMoved As Boolean
Private mblnPrinting As Boolean

Private Sub SetDockingPanel()
    '-----------------------------------------------------------------------------------------------------------------------
    '����:����DOCKINGPANEL�ؼ�
    '����:������
    '����:2013-09-04
    '��ע:
    '-----------------------------------------------------------------------------------------------------------------------
    Dim objPanel As Pane
    On Error GoTo errHandle
    
    With dkpMain
        .VisualTheme = ThemeOffice2003
        Set objPanel = .CreatePane(1, 2000, 2000, DockTopOf)
        objPanel.Handle = picMain.hWnd
        objPanel.Options = PaneNoCloseable + PaneNoFloatable + PaneNoHideable + PaneNoCaption
        Set objPanel = .CreatePane(2, 1700, 1000, DockBottomOf, objPanel)
        objPanel.Handle = picDetail.hWnd
        objPanel.Options = PaneNoCloseable + PaneNoFloatable + PaneNoHideable + PaneNoCaption
        Set objPanel = .CreatePane(3, 1000, 1000, DockRightOf, objPanel)
        objPanel.Handle = picBalance.hWnd
        objPanel.Options = PaneNoCloseable + PaneNoFloatable + PaneNoHideable + PaneNoCaption
        .Options.HideClient = True
    End With
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub SetActiveList(obj As Object)
    If obj Is vsfMain Then
        vsfMain.BackColorSel = &HC0C0C0
        vsfBalance.BackColorSel = &HE0E0E0
        vsfDetail.BackColorSel = &HE0E0E0
    ElseIf obj Is vsfBalance Then
        vsfBalance.BackColorSel = &HC0C0C0
        vsfMain.BackColorSel = &HE0E0E0
        vsfDetail.BackColorSel = &HE0E0E0
    ElseIf obj Is vsfDetail Then
        vsfDetail.BackColorSel = &HC0C0C0
        vsfBalance.BackColorSel = &HE0E0E0
        vsfMain.BackColorSel = &HE0E0E0
    End If
End Sub

Private Sub vsfBalance_GotFocus()
    SetActiveList vsfBalance
End Sub

Private Sub vsfDetail_GotFocus()
    SetActiveList vsfDetail
End Sub

Private Sub vsfMain_GotFocus()
    SetActiveList vsfMain
End Sub

Private Sub vsfMain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim intRow As Integer
    With vsfMain
        'If .TextMatrix(1, .ColIndex("�������")) = "" Then Exit Sub
        If Button = 2 Then
            If Y <= 300 Then
                Exit Sub
            End If
            intRow = Y \ 300
            If intRow <= .Rows - 1 Then
                If .Enabled And .Visible Then .SetFocus
                .Select intRow, 0
            End If
            Call frmReplenishTheBalanceManage.ShowPopup
        End If
    End With
End Sub

Public Sub ReadData(ByVal intType As Integer, Optional ByVal lngPatiID As Long)
    '-----------------------------------------------------------------------------------------------------------------------
    '����:��ȡ���ղ�������¼
    '����:������
    '���:intType-��ȡ��¼�ķ�ʽ��0Ϊʹ�ù���������ȡ��1Ϊʹ��IDKIND������ȡ
    '����:2014-9-11
    '��ע:
    '-----------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsMain As ADODB.Recordset
    Dim dtStartDate As Date, dtEndDate As Date
    If intType = 0 Then
        Select Case frmReplenishTheBalanceManage.cboDate.ListIndex
            Case 0 '����
                dtStartDate = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd") & " 00:00:00")
                dtEndDate = CDate(Format(dtStartDate, "yyyy-mm-dd") & " 23:59:59")
            Case 1 '�������
                dtStartDate = CDate(Format(DateAdd("d", -1, frmReplenishTheBalanceManage.dtpStartDate.MaxDate), "yyyy-mm-dd") & " 00:00:00")
                dtEndDate = CDate(Format(frmReplenishTheBalanceManage.dtpStartDate.MaxDate, "yyyy-mm-dd") & " 23:59:59")
            Case 2 '�������
                dtStartDate = CDate(Format(DateAdd("d", -2, frmReplenishTheBalanceManage.dtpStartDate.MaxDate), "yyyy-mm-dd") & " 00:00:00")
                dtEndDate = CDate(Format(frmReplenishTheBalanceManage.dtpStartDate.MaxDate, "yyyy-mm-dd") & " 23:59:59")
            Case 3  '���һ��
                dtStartDate = CDate(Format(DateAdd("d", -7, frmReplenishTheBalanceManage.dtpStartDate.MaxDate), "yyyy-mm-dd") & " 00:00:00")
                dtEndDate = CDate(Format(frmReplenishTheBalanceManage.dtpStartDate.MaxDate, "yyyy-mm-dd") & " 23:59:59")
            Case 4  '
                dtStartDate = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-01") & " 00:00:00")
                dtEndDate = CDate(Format(frmReplenishTheBalanceManage.dtpStartDate.MaxDate, "yyyy-mm-dd") & " 23:59:59")
            Case Else
                dtStartDate = CDate(Format(frmReplenishTheBalanceManage.dtpStartDate.Value, "yyyy-mm-dd") & " 00:00:00")
                dtEndDate = CDate(Format(frmReplenishTheBalanceManage.dtpEndDate.Value, "yyyy-mm-dd") & " 23:59:59")
        End Select
        strSQL = " Select A.No, Decode(Nvl(A.���ӱ�־,0),1,'�Һ�','�շ�'), B.����, B.�Ա�, B.����, Sum(B.���ʽ��), A.����Ա����, A.�Ǽ�ʱ��, A.�������, A.����ID " & _
                 " From ���ò����¼ A, ������ü�¼ B " & _
                 " Where A.�Ǽ�ʱ�� Between [1] And [2] And Nvl(A.����״̬,0)=1 And A.�շѽ���ID=B.����ID And A.��¼״̬ In (1,3) " & _
                 "      And A.����Ա����=[3] And Not Exists (Select 1 From ���ò����¼ Where �������=A.������� And ��¼״̬=2)" & _
                 " Group By A.No, Decode(Nvl(A.���ӱ�־,0),1,'�Һ�','�շ�'), B.����, B.�Ա�, B.����, A.����Ա����, A.�Ǽ�ʱ��, A.�������, A.����ID "
        Set rsMain = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, dtStartDate, dtEndDate, UserInfo.����)
        Set vsfMain.DataSource = rsMain
        If rsMain.RecordCount <> 0 Then
            frmReplenishTheBalanceManage.tabMain.Item(1).Caption = "�쳣�����¼(" & rsMain.RecordCount & ")"
        Else
            frmReplenishTheBalanceManage.tabMain.Item(1).Caption = "�쳣�����¼"
        End If
        Call SetMain
    End If
    If intType = 1 Then
        'ʹ��IDKIND������ȡ
        strSQL = " Select A.No, Decode(Nvl(A.���ӱ�־,0),1,'�Һ�','�շ�'), B.����, B.�Ա�, B.����, Sum(B.���ʽ��), A.����Ա����, A.�Ǽ�ʱ��, A.�������, A.����ID " & _
                 " From ���ò����¼ A, ������ü�¼ B " & _
                 " Where B.����ID= [1] And Nvl(A.����״̬,0)=1 And A.�շѽ���ID=B.����ID And A.��¼״̬ In (1,3) " & _
                 "      And A.����Ա����=[2] And Not Exists (Select 1 From ���ò����¼ Where �������=A.������� And ��¼״̬=2)" & _
                 " Group By A.No, Decode(Nvl(A.���ӱ�־,0),1,'�Һ�','�շ�'), B.����, B.�Ա�, B.����, A.����Ա����, A.�Ǽ�ʱ��, A.�������, A.����ID "
        Set rsMain = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngPatiID, UserInfo.����)
        Set vsfMain.DataSource = rsMain
        If rsMain.RecordCount <> 0 Then
            frmReplenishTheBalanceManage.tabMain.Item(1).Caption = "�쳣�����¼(" & rsMain.RecordCount & ")"
        Else
            frmReplenishTheBalanceManage.tabMain.Item(1).Caption = "�쳣�����¼"
        End If
        Call SetMain
    End If
End Sub

Private Sub ReadBalance(Optional ByVal lngBalanceID As Long)
    Dim strSQL As String, i As Long, rsBalance As ADODB.Recordset
    
    strSQL = _
        " Select Nvl(A.���㷽ʽ,'δ����') As ���㷽ʽ,Sum(A.��Ԥ��) As ��Ԥ��,Decode(Nvl(A.У�Ա�־,0),0,'��',2,'��','��') As ��־,Nvl(B.����,0) As ���� " & _
        " From ����Ԥ����¼ A,���㷽ʽ B " & _
        " Where A.������� = [1] And A.���㷽ʽ=B.����(+)" & _
        " Group By Nvl(A.���㷽ʽ,'δ����'),Nvl(A.У�Ա�־,0),Nvl(B.����,0)" & _
        " Having Sum(A.��Ԥ��) <> 0 Order By ����"
    
    Set rsBalance = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngBalanceID)
    
    vsfBalance.Redraw = False
    vsfBalance.Clear
    vsfBalance.Rows = 2
    If Not rsBalance.EOF Then
        Set vsfBalance.DataSource = rsBalance
    End If
    Call SetBalance
    vsfBalance.Redraw = True
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ReadDetail(ByVal lngBalanceID As Long, ByVal bln�ҺŲ��� As Boolean)
    Dim strSQL As String, rsDetail As ADODB.Recordset
'    mblnNOMoved = zlDatabase.NOMoved("���ò����¼", vsfMain.TextMatrix(vsfMain.Row, vsfMain.ColIndex("���㵥��")))
    strSQL = _
            " Select NO As ���ݺ�, ���, ��������, ������, �ѱ�, ���, ����, ��Ʒ��, ���, ��λ, " & _
            "       Sum(����) As ����, ����, Sum(Ӧ�ս��) As Ӧ�ս��, Sum(ʵ�ս��) As ʵ�ս��, ִ�п���, Max(����) As ����, Max(˵��),Max(״̬), Min(�˷�״̬)" & vbNewLine & _
            " From (Select a.����ID,D1.���� as ��������,A.������,a.No,C.���� as ���,Nvl(E.����,B.����) as ����,E1.���� as ��Ʒ��,B.���," & _
                    IIf(gblnҩ����λ, "Decode(X.ҩƷID,NULL,A.���㵥λ,X." & gstrҩ����λ & ")", "A.���㵥λ") & " as ��λ," & _
            "       To_Char(Avg(Nvl(A.����,1)*A.����)" & _
                    IIf(gblnҩ����λ, "/Nvl(X." & gstrҩ����װ & ",1)", "") & ",'9999990.00000') as ����, " & _
            "       a.�ѱ�,To_Char(Sum(A.��׼����)" & _
                    IIf(gblnҩ����λ, "*Nvl(X." & gstrҩ����װ & ",1)", "") & ",'999999" & gstrFeePrecisionFmt & "') as ����, " & _
            "       To_Char(Sum(A.Ӧ�ս��),'9999999" & gstrDec & "') as Ӧ�ս��, " & _
            "       To_Char(Sum(A.ʵ�ս��),'9999999" & gstrDec & "') as ʵ�ս��, " & _
            "       D.���� as ִ�п���,Nvl(A.��������,B.��������) as ����,Max(Decode(A.��¼״̬,2,'��'||ABS(A.ִ��״̬)||'���˷�',Decode(Nvl(A.ִ��״̬,0),0,'δִ��',1,'��ȫִ��',2,'����ִ��'))) As ˵��," & _
            "       Max(A.��¼״̬) As ״̬,Min(A.��¼״̬) As �˷�״̬, Nvl(a.�۸񸸺�, a.���) As ���" & _
            " From " & IIf(mblnNOMoved, "H", "") & "������ü�¼ A,�շ���ĿĿ¼ B,�շ���Ŀ��� C,���ű� D,���ű� D1,�շ���Ŀ���� E,�շ���Ŀ���� E1,ҩƷ��� X," & _
            "       (Select Distinct �շѽ���ID As ����ID From " & IIf(mblnNOMoved, "H", "") & "���ò����¼ Where �������= [1]) F" & _
            " Where A.�շ�ϸĿID=B.ID and A.�շ����=C.���� And A.ִ�в���ID=D.ID(+) And A.�շ�ϸĿID=X.ҩƷID(+)" & _
            "       And Mod(A.��¼����,10)=[2] And A.����ID = F.����ID " & _
            "       And A.�շ�ϸĿID=E.�շ�ϸĿID(+) And E.����(+)=1 And E.����(+)=" & IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1) & _
            "       And A.�շ�ϸĿID=E1.�շ�ϸĿID(+) And A.��������ID=D1.ID(+) And E1.����(+)=1 And E1.����(+)=3" & _
            " Group by a.����id, D1.����, a.������, a.�ѱ�,a.No,Nvl(A.�۸񸸺�,A.���),C.����,Nvl(E.����,B.����),E1.���� , B.���,A.���㵥λ,D.����," & _
            "       Nvl(A.��������,B.��������),X.ҩƷID,X." & gstrҩ����λ & ",Nvl(X." & gstrҩ����װ & ",1) )" & _
            " Group By NO, ���, ��������, ������, �ѱ�, ���, ����, ��Ʒ��, ���, ��λ, ����, ִ�п��� Having Sum(����) <> 0" & _
            " Order By ���ݺ�, ���"
    
    Set rsDetail = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngBalanceID, IIf(bln�ҺŲ���, 4, 1))
    vsfDetail.Redraw = False
    vsfDetail.Clear 1
    vsfDetail.Rows = 2
    If Not rsDetail.EOF Then
        Set vsfDetail.DataSource = rsDetail
    End If
    Call SetDetail
    vsfDetail.Redraw = True
End Sub

Private Sub SetDetail()
    Dim strHead As String
    Dim i As Long
    Dim varData As Variant
    
    strHead = "���ݺ�,1,0|���,1,0|��������,1,0|������,1,0|�ѱ�,1,0|���,4,800|����,1,2000|��Ʒ��,1,2000|" & _
            "���,1,1200|��λ,4,500|����,7,800|����,7,1000|Ӧ�ս��,7,1000|ʵ�ս��,7,1000|ִ�п���,4,1000|" & _
            "����,4,1000|˵��,1,1800|��¼״̬,1,0"
    
    With vsfDetail
        .HighLight = flexHighlightWithFocus
        .Redraw = False
        varData = Split(strHead, "|")
        .COLS = UBound(varData) + 1
        For i = 0 To UBound(varData)
            If Split(varData(i), ",")(0) = "ID" Then .ColHidden(i) = True
            .TextMatrix(0, i) = Split(varData(i), ",")(0)
            .ColKey(i) = .TextMatrix(0, i)
            .ColAlignment(i) = Split(varData(i), ",")(1)
            .ColWidth(i) = Split(varData(i), ",")(2)
            .FixedAlignment(i) = 4
        Next
        
        Call RestoreFlexState(vsfDetail, App.ProductName & "\" & Me.Name)
        
        .RowHeight(0) = 350
        If .Rows < 2 Then .Rows = 2
        .Row = 1: .Col = 0: .ColSel = .COLS - 1

        .Redraw = True
        If .TextMatrix(1, .ColIndex("���ݺ�")) <> "" Then Call DetailSplitGroup
        For i = 1 To .Rows - 1
            If .IsSubtotal(i) = False Then
                If Val(.TextMatrix(i, .ColIndex("��¼״̬"))) = 1 Then .Cell(flexcpForeColor, i, 0, i, .COLS - 1) = vbBlack
                If Val(.TextMatrix(i, .ColIndex("��¼״̬"))) = 2 Then .Cell(flexcpForeColor, i, 0, i, .COLS - 1) = vbRed
                If Val(.TextMatrix(i, .ColIndex("��¼״̬"))) = 3 Then .Cell(flexcpForeColor, i, 0, i, .COLS - 1) = vbBlue
                .RowHeight(i) = 300
            End If
        Next i
        
        If gTy_System_Para.bytҩƷ������ʾ = 0 Then
            .ColHidden(.ColIndex("����")) = False
            .ColHidden(.ColIndex("��Ʒ��")) = True
        End If
        If gTy_System_Para.bytҩƷ������ʾ = 1 Then
            .ColHidden(.ColIndex("����")) = True
            .ColHidden(.ColIndex("��Ʒ��")) = False
        End If
        If gTy_System_Para.bytҩƷ������ʾ = 2 Then
            .ColHidden(.ColIndex("����")) = False
            .ColHidden(.ColIndex("��Ʒ��")) = False
        End If
    End With
End Sub

Private Sub SetBalance()
    Dim strHead As String, strTemp As String
    Dim i As Long, strAcc As String, j As Long
    Dim varData As Variant
    
    strHead = "���㷽ʽ,4,1200|������,7,1000|�����Ƿ�ɹ�,4,1200|����,1,0"
    
    With vsfBalance
        .HighLight = flexHighlightWithFocus
        .Redraw = False
        varData = Split(strHead, "|")
        .COLS = UBound(varData) + 1
        For i = 0 To UBound(varData)
            .TextMatrix(0, i) = Split(varData(i), ",")(0)
            .ColKey(i) = .TextMatrix(0, i)
            .ColAlignment(i) = Split(varData(i), ",")(1)
            .ColWidth(i) = Split(varData(i), ",")(2)
            .FixedAlignment(i) = 4
        Next
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("���㷽ʽ")) Like "*���*" Then
                .Cell(flexcpForeColor, i, 0, i, .COLS - 1) = vbRed
                strTemp = Val(.TextMatrix(i, .ColIndex("������")))
                If InStr(strTemp, ".") = 0 Then
                    strAcc = "0.00"
                Else
                    strTemp = Split(strTemp, ".")(1)
                    strAcc = "0."
                    If Len(strTemp) < 2 Then
                        strAcc = "0.00"
                    Else
                        For j = 1 To Len(strTemp)
                            strAcc = strAcc & "0"
                        Next j
                    End If
                End If
                .TextMatrix(i, .ColIndex("������")) = Format(.TextMatrix(i, .ColIndex("������")), strAcc)
            Else
                If .TextMatrix(i, .ColIndex("������")) <> "" Then .TextMatrix(i, .ColIndex("������")) = Format(.TextMatrix(i, .ColIndex("������")), "0.00")
            End If
            .RowHeight(i) = 300
        Next i
        
        Call RestoreFlexState(vsfBalance, App.ProductName & "\" & Me.Name)
        
        .RowHeight(0) = 350
        
        .Row = 1: .Col = 0: .ColSel = .COLS - 1

        .Redraw = True
    End With
End Sub

Private Sub SetMain()
    Dim i As Long
    With vsfMain
        .RowHeight(0) = 350
        If .Rows = 1 Then .Rows = 2
        For i = 1 To .Rows - 1
            .RowHeight(i) = 300
            .TextMatrix(i, .ColIndex("������")) = Format(.TextMatrix(i, .ColIndex("������")), gstrDec)
        Next i
        If .Rows >= 2 Then .Select 1, 1
    End With
End Sub

Private Sub DetailSplitGroup()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��Է����б���Ϣ���з�����ʾ
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer, j As Integer
    Dim strTemp As String

    On Error GoTo errHandle
    With vsfDetail
        For i = 0 To .COLS - 1
            If i < .ColIndex("���") And i > .ColIndex("˵��") Then
                .ColHidden(i) = True
            End If
        Next
        .OutlineBar = flexOutlineBarComplete
        .Subtotal flexSTClear
        .MultiTotals = True
        '&H8000000F
        .Subtotal flexSTSum, .ColIndex("���ݺ�"), .ColIndex("ʵ�ս��"), gstrDec, &H8000000F, , True, "%s", , True
        .Subtotal flexSTSum, .ColIndex("���ݺ�"), .ColIndex("Ӧ�ս��"), gstrDec, &H8000000F, , True, "%s", , True
        .SubtotalPosition = flexSTAbove

        .Outline .ColIndex("���")
        .OutlineCol = .ColIndex("���")
        
        For i = 1 To .Rows - 1
            .MergeRow(i) = False
            If .IsSubtotal(i) Then
                .IsCollapsed(i) = flexOutlineExpanded
                strTemp = .Cell(flexcpTextDisplay, i, 0)
                .RowHeight(i) = 350

                .Cell(flexcpText, i, .ColIndex("���")) = strTemp
                
                 strTemp = .Cell(flexcpTextDisplay, i + 1, .ColIndex("���ݺ�"))
                 strTemp = strTemp & Space(2) & "�ѱ�:" & .Cell(flexcpTextDisplay, i + 1, .ColIndex("�ѱ�"))
                 strTemp = strTemp & Space(2) & "��������:" & .Cell(flexcpTextDisplay, i + 1, .ColIndex("��������"))
                 strTemp = strTemp & Space(2) & "������:" & .Cell(flexcpTextDisplay, i + 1, .ColIndex("������"))
                 .MergeRow(i) = True
                 .MergeCells = flexMergeRestrictRows
                 .Cell(flexcpAlignment, i, .ColIndex("���"), i, .ColIndex("���")) = 1
'                 If Val(.TextMatrix(i + 1, .ColIndex("��¼״̬"))) = 1 Then .Cell(flexcpForeColor, i, 0, i, .COLS - 1) = vbBlack
'                 If Val(.TextMatrix(i + 1, .ColIndex("��¼״̬"))) = 2 Then .Cell(flexcpForeColor, i, 0, i, .COLS - 1) = vbRed
'                 If Val(.TextMatrix(i + 1, .ColIndex("��¼״̬"))) = 3 Then .Cell(flexcpForeColor, i, 0, i, .COLS - 1) = vbBlue
                 
                 For j = 0 To .COLS - 1
                    If j < .ColIndex("Ӧ�ս��") Then
                        If j >= .ColIndex("���") Then
                            .Cell(flexcpText, i, j) = strTemp
                            .Cell(flexcpFontBold, i, j) = False
                        End If
                    ElseIf .ColIndex("ʵ�ս��") = j Then
                        .TextMatrix(i, j) = Format(Val(.TextMatrix(i, j)), gstrDec)
                        .Cell(flexcpFontBold, i, j) = False
                    ElseIf .ColIndex("Ӧ�ս��") = j Then
                        .TextMatrix(i, j) = " " & Format(Val(.TextMatrix(i, j)), gstrDec)
                        .Cell(flexcpFontBold, i, j) = False
                    End If
                 Next
            Else
                .TextMatrix(i, .ColIndex("����")) = Format(Val(.TextMatrix(i, .ColIndex("����"))), gstrFeePrecisionFmt)
                .TextMatrix(i, .ColIndex("����")) = FormatEx(Val(.TextMatrix(i, .ColIndex("����"))), 5)
                .TextMatrix(i, .ColIndex("Ӧ�ս��")) = Format(Val(.TextMatrix(i, .ColIndex("Ӧ�ս��"))), gstrDec)
                .TextMatrix(i, .ColIndex("ʵ�ս��")) = Format(Val(.TextMatrix(i, .ColIndex("ʵ�ս��"))), gstrDec)
            End If
        Next
        Call .AutoSize(.ColIndex("���"))
        Call .AutoSize(.ColIndex("����"))
        
        For j = 0 To .COLS - 1
            If j < .ColIndex("Ӧ�ս��") Then
                .MergeCol(j) = True
            Else
                .MergeCol(j) = False
            End If
        Next
        
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Load()
    Call SetDockingPanel
    Call SetMain
    Call SetBalance
    Call SetDetail
End Sub

Private Sub PicDetail_Resize()
    With vsfDetail
        .Top = 0
        .Left = 0
        .Height = picDetail.Height
        .Width = picDetail.Width
    End With
End Sub

Private Sub vsfMain_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If mblnPrinting Then Exit Sub
    If vsfMain.TextMatrix(NewRow, vsfMain.ColIndex("�������")) = "" Then Exit Sub
    Call ReadDetail(Val(vsfMain.TextMatrix(NewRow, vsfMain.ColIndex("�������"))), _
                    vsfMain.TextMatrix(NewRow, vsfMain.ColIndex("����")) = "�Һ�")
    Call ReadBalance(Val(vsfMain.TextMatrix(NewRow, vsfMain.ColIndex("�������"))))
End Sub

Private Sub picBalance_Resize()
    With vsfBalance
        .Top = 0
        .Left = 0
        .Height = picBalance.Height
        .Width = picBalance.Width
    End With
End Sub

Private Sub picMain_Resize()
    With vsfMain
        .Top = 0
        .Left = 0
        .Height = picMain.Height
        .Width = picMain.Width
    End With
End Sub

Private Sub vsfMain_DblClick()
    Call frmReplenishTheBalanceManage.ViewBalance(1)
End Sub

Public Sub zlRptPrint(ByVal bytFunc As Integer)
    '����:���д�ӡ,Ԥ���������EXCEL
    '���:bytFunc=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    Dim i As Long, lngCurrentRow As Long
    Dim objPrint As Object, objRow As New zlTabAppRow, bytPrn As Byte

    With vsfMain
        If .Rows = 1 Then Exit Sub
        If .Rows = 2 And Val(.TextMatrix(1, .ColIndex("�������"))) = 0 Then Exit Sub
    End With
    
    Set objPrint = New zlPrint1Grd
    objPrint.Title.Font.Name = "����_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    objPrint.Title.Text = "���ղ�������쳣�����¼�嵥"
    
    objPrint.UnderAppRows.Add objRow
    Set objRow = New zlTabAppRow
    objRow.Add "��ӡ��:" & UserInfo.����
    objRow.Add "��ӡ����:" & Format(zlDatabase.Currentdate, "yyyy��MM��dd��")
    objPrint.BelowAppRows.Add objRow
    
    '���ڴ�ӡ�ؼ�����ʶ������������
    With vsfMain
        .Redraw = flexRDNone
        .GridColor = .ForeColor
        For i = 0 To .COLS - 1
            .Cell(flexcpData, 0, i) = .ColWidth(i)
            If .ColHidden(i) = True Then
                .ColWidth(i) = 0
            End If
        Next
    End With

    Err = 0: On Error GoTo ErrHand:
    mblnPrinting = True
    lngCurrentRow = vsfMain.Row
    Set objPrint.Body = vsfMain
    If bytFunc = 1 Then
        Select Case zlPrintAsk(objPrint)
            Case 1
                zlPrintOrView1Grd objPrint, 1
            Case 2
                zlPrintOrView1Grd objPrint, 2
            Case 3
                zlPrintOrView1Grd objPrint, 3
        End Select
    Else
        zlPrintOrView1Grd objPrint, bytPrn
    End If
    
    '�ָ�
    With vsfMain
        For i = 0 To .COLS - 1
            If .ColHidden(i) = True Then
                .ColWidth(i) = Val(.Cell(flexcpData, 0, i))
            End If
        Next
        .GridColor = &H8000000F
        .Redraw = flexRDBuffered
    End With
    vsfMain.Row = lngCurrentRow
    mblnPrinting = False
    Exit Sub
ErrHand:
    mblnPrinting = False
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
