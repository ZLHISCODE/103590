VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Begin VB.Form frmBalanceTabErr 
   BorderStyle     =   0  'None
   Caption         =   "frmBalanceTabErr"
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
      TabStop         =   0   'False
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
         ForeColorSel    =   -2147483640
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
      TabStop         =   0   'False
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
         ForeColorSel    =   -2147483640
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
         ExplorerBar     =   3
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
      TabStop         =   0   'False
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
         ForeColorSel    =   -2147483640
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
         FormatString    =   $"frmBalanceTabErr.frx":0000
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
         ExplorerBar     =   3
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
Attribute VB_Name = "frmBalanceTabErr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnPrint As Boolean
Private mblnNOMoved As Boolean
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

Private Sub Form_Unload(Cancel As Integer)
    mblnPrint = False
End Sub

Private Sub vsfBalance_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    zl_VsGridRowChange vsfBalance, OldRow, NewRow, OldCol, NewCol
End Sub

Private Sub vsfBalance_GotFocus()
    zl_VsGridGotFocus vsfBalance, &HFFC0C0
End Sub

Private Sub vsfBalance_LostFocus()
    zl_VsGridLOSTFOCUS vsfBalance
End Sub

Private Sub vsfDetail_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    zl_VsGridRowChange vsfDetail, OldRow, NewRow, OldCol, NewCol
End Sub

Private Sub vsfDetail_GotFocus()
    zl_VsGridGotFocus vsfDetail, &HFFC0C0
End Sub

Private Sub vsfDetail_LostFocus()
    zl_VsGridLOSTFOCUS vsfDetail
End Sub

Private Sub vsfMain_GotFocus()
    zl_VsGridGotFocus vsfMain, &HFFC0C0
End Sub

Private Sub vsfMain_LostFocus()
    zl_VsGridLOSTFOCUS vsfMain
End Sub

Private Sub vsfMain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim intRow As Integer
    With vsfMain
        If Button = 2 Then
            If Y <= 300 Then
                Exit Sub
            End If
            Call frmManageBalance.ShowPopup
        End If
    End With
End Sub

Public Sub ReadData()
    '-----------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�쳣�����¼
    '����:������
    '����:2015-01-06
    '��ע:
    '-----------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, rsMain As ADODB.Recordset, strFilter As String, strTable As String
    Dim dtStartDate As Date, dtEndDate As Date, blnAll As Boolean
    Select Case frmManageBalance.cboDate.ListIndex
        Case 0 '�����쳣
            dtStartDate = CDate(Format("1900-01-01", "yyyy-mm-dd") & " 00:00:00")
            dtEndDate = CDate(Format("3000-01-01", "yyyy-mm-dd") & " 23:59:59")
        Case 1 '����
            dtStartDate = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd") & " 00:00:00")
            dtEndDate = CDate(Format(dtStartDate, "yyyy-mm-dd") & " 23:59:59")
        Case 2 'ǰһ��������
            dtStartDate = CDate(Format(DateAdd("d", -1, frmManageBalance.dtpStartDate.MaxDate), "yyyy-mm-dd") & " 00:00:00")
            dtEndDate = CDate(Format(frmManageBalance.dtpStartDate.MaxDate, "yyyy-mm-dd") & " 23:59:59")
        Case 3 'ǰ����������
            dtStartDate = CDate(Format(DateAdd("d", -2, frmManageBalance.dtpStartDate.MaxDate), "yyyy-mm-dd") & " 00:00:00")
            dtEndDate = CDate(Format(frmManageBalance.dtpStartDate.MaxDate, "yyyy-mm-dd") & " 23:59:59")
        Case 4  'ǰһ��������
            dtStartDate = CDate(Format(DateAdd("d", -7, frmManageBalance.dtpStartDate.MaxDate), "yyyy-mm-dd") & " 00:00:00")
            dtEndDate = CDate(Format(frmManageBalance.dtpStartDate.MaxDate, "yyyy-mm-dd") & " 23:59:59")
        Case 5  '����
            dtStartDate = CDate(Format(zlDatabase.Currentdate, "yyyy-mm") & "-01 00:00:00")
            dtEndDate = CDate(Format(frmManageBalance.dtpStartDate.MaxDate, "yyyy-mm-dd") & " 23:59:59")
        Case Else
            dtStartDate = CDate(Format(frmManageBalance.dtpStartDate.Value, "yyyy-mm-dd") & " 00:00:00")
            dtEndDate = CDate(Format(frmManageBalance.dtpEndDate.Value, "yyyy-mm-dd") & " 23:59:59")
    End Select
    strFilter = " And A.��¼״̬ In (1,3) And A.�շ�ʱ�� Between [1] And [2] And A.����Ա���� = [3] "
    strTable = "" & _
            "   Select A.ID ,1 as סԺ��־,0 as �����־,A.NO,A.ʵ��Ʊ��,A.����ID,B.����ID as ���ò���ID,Nvl(D.�Ա�,C.�Ա�) as �Ա�,Nvl(D.����,C.����) as ���� ,A.��ʼ����,A.��������,Max(A.��¼״̬) As ��¼״̬,Sum(B.���ʽ��) As ���ʽ��,A.����Ա����,A.�շ�ʱ��,A.��;����,A.ԭ�� as ��Լ��λ,A.�������� " & _
            "   From ���˽��ʼ�¼ A,סԺ���ü�¼ B,������Ϣ C,������ҳ D " & _
            "   Where A.ID=B.����ID and  B.����ID=C.����ID And A.����ID =D.����ID(+) And A.��ҳID = D.��ҳID(+) And A.����״̬ = 1 And Not Exists (Select 1 From ���˽��ʼ�¼ Where NO = a.No And ��¼״̬ = 2) " & strFilter & _
            "   Group By A.ID ,A.NO,A.ʵ��Ʊ��,A.����ID,B.����ID,Nvl(D.�Ա�,C.�Ա�),Nvl(D.����,C.����),A.��ʼ����,A.��������,A.����Ա����,A.�շ�ʱ��,A.��;����,A.ԭ��,A.�������� "

    strTable = strTable & vbCrLf & " Union ALL " & vbCrLf & Replace(Replace(strTable, "סԺ���ü�¼", "������ü�¼"), "1 as סԺ��־,0 as �����־", "0 as סԺ��־,1 as �����־")
    
    strSql = _
            " Select A.ID ����ID,decode(סԺ��־,1,decode(�����־,1,3,2),1) as ��־,decode(A.��������,1,'�������',2,'סԺ����','') As �������� ,Decode(P.����,NULL,Decode(C.����,NULL,NULL,'��'),'��') as ҽ��,A.NO as ���ݺ�,A.ʵ��Ʊ�� as Ʊ�ݺ�," & _
            "        Decode(A.����ID,Null,' ',A.����ID) ����ID,Decode(Nvl(A.��������,0),2,' ',Decode(A.����ID,Null,' ',C.�����)) �����,Decode(A.����ID,Null,' ',C.סԺ��) סԺ��," & _
            "        Decode(A.����ID,Null,nvl(A.��Լ��λ,Q.����),C.����) ����,Decode(A.����ID,Null,' ',A.�Ա�) �Ա�," & _
            "        Decode(A.����ID,Null,' ',A.����) ����,Decode(A.����ID,Null,' ',Nvl(P.�ѱ�,C.�ѱ�)) as �ѱ�," & _
            "        To_Char(A.��ʼ����,'YYYY-MM-DD') as ��ʼ����,To_Char(A.��������,'YYYY-MM-DD') as ��������," & _
            "        To_Char(Decode(A.��¼״̬,2,-1,1) *A.���ʽ��,'999999999" & gstrDec & "') as ���ʽ��," & _
            "        A.����Ա���� as ����Ա,To_Char(A.�շ�ʱ��,'YYYY-MM-DD HH24:MI:SS') as �շ�ʱ��,Decode(Nvl(A.��;����,0),1,'��',' ') ��;����,A.��¼״̬ as ��¼״̬" & _
            " From ( " & strTable & ") A,������Ϣ C,������ҳ P,��Լ��λ Q,��Ա�� N" & _
            " Where  A.���ò���ID=C.����ID And A.����Ա����=N.���� " & _
            "        And C.����ID=P.����ID(+) And Nvl(C.��ҳID,0)=P.��ҳID(+) And C.��ͬ��λID=Q.ID(+)" & _
            "       And (N.վ��='" & gstrNodeNo & "' Or N.վ�� is Null)" & vbNewLine
            
    strSql = strSql & " Order by �շ�ʱ�� Desc,���ݺ� Desc"
    
    Set rsMain = zlDatabase.OpenSQLRecord(strSql, Me.Caption, dtStartDate, dtEndDate, UserInfo.����)
    Set vsfMain.DataSource = rsMain
    If rsMain.RecordCount <> 0 Then
        frmManageBalance.tabMain.Item(1).Caption = "�쳣�����¼(" & rsMain.RecordCount & ")"
        frmManageBalance.stbThis.Panels(2).Text = "��ǰ����" & rsMain.RecordCount & "���쳣�����¼,�ϼ�:" & Format(GetTotal, gstrDec) & "Ԫ"
    Else
        frmManageBalance.tabMain.Item(1).Caption = "�쳣�����¼"
        frmManageBalance.stbThis.Panels(2).Text = ""
    End If
    Call SetMain
End Sub

Private Sub ReadBalance(Optional ByVal lngBalanceID As Long)
    Dim strSql As String, i As Long, rsBalance As ADODB.Recordset
    
    If mblnPrint Then Exit Sub
    
    strSql = _
        " Select Nvl(A.���㷽ʽ,'δ����') As ���㷽ʽ,Sum(A.��Ԥ��) As ��Ԥ��," & _
        "       Decode(Nvl(A.У�Ա�־,0),0,'��',2,'��','��') As ��־,Nvl(B.����,0) As ���� " & _
        " From ����Ԥ����¼ A,���㷽ʽ B " & _
        " Where A.����ID = [1] And A.���㷽ʽ=B.����(+)" & _
        " Group By Nvl(A.���㷽ʽ,'δ����'),Nvl(A.У�Ա�־,0),Nvl(B.����,0)" & _
        " Having Sum(A.��Ԥ��) <> 0 Order By ����"
    
    Set rsBalance = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngBalanceID)
    
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

Private Sub ReadDetail(ByVal lngBalanceID As Long)
    Dim strSql As String, rsDetail As ADODB.Recordset, rsTmp As ADODB.Recordset
    Dim blnDel As Boolean, strDec As String, int��Դ As Integer
    
    If mblnPrint Then Exit Sub
    
    int��Դ = Val(vsfMain.TextMatrix(vsfMain.Row, vsfMain.ColIndex("��־")))
    blnDel = Val(vsfMain.TextMatrix(vsfMain.Row, vsfMain.ColIndex("��¼״̬"))) = 2
    strDec = gstrDec
    If lngBalanceID <> 0 Then
        Select Case int��Դ
        Case 1 '����
            strSql = "Select Max(Length(Abs(���ʽ��) - Trunc(Abs(���ʽ��))))-1 declen From " & IIf(mblnNOMoved, "H", "") & "������ü�¼ Where ����ID=[1]"
        Case 2 'סԺ
            strSql = "Select Max(Length(Abs(���ʽ��) - Trunc(Abs(���ʽ��))))-1 declen From " & IIf(mblnNOMoved, "H", "") & "סԺ���ü�¼ Where ����ID=[1]"
        Case Else
            
            strSql = "Select Length(Abs(���ʽ��) - Trunc(Abs(���ʽ��)))  as  declen From " & IIf(mblnNOMoved, "H", "") & "������ü�¼ Where ����ID=[1] Union ALL " & _
                     "Select Length(Abs(���ʽ��) - Trunc(Abs(���ʽ��)))   as  declen  From " & IIf(mblnNOMoved, "H", "") & "סԺ���ü�¼ Where ����ID=[1]"
            strSql = "Select Max(declen)-1 as declen  From ( " & strSql & ")"
        End Select
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngBalanceID)
        If rsTmp.RecordCount > 0 Then
            If Len(strDec) < Len("0." & String(rsTmp!declen, "0")) Then
                strDec = "0." & String(rsTmp!declen, "0")
            End If
        End If
    End If
    
    Select Case int��Դ
    Case 1  '����
        strSql = " (Select ����ID,NO,���,��������ID,�շ�ϸĿID,�����־,0 as ��ҳID,�վݷ�Ŀ,Ӥ����,���ʽ��,����ʱ�� From " & IIf(mblnNOMoved, "H", "") & "������ü�¼ A where A.����ID=[1] ) A "
        'strSQL = IIf(mblnNOMoved, "H", "") & "������ü�¼ A "
    Case 2  'סԺ
        strSql = IIf(mblnNOMoved, "H", "") & "סԺ���ü�¼ A"
    Case Else '�����סԺ
        strSql = " (Select ����ID,NO,���,��������ID,�շ�ϸĿID,�����־,0 as ��ҳID,�վݷ�Ŀ,Ӥ����,���ʽ��,����ʱ�� From " & IIf(mblnNOMoved, "H", "") & "������ü�¼ A where A.����ID=[1] Union ALL " & _
                   " Select ����ID,NO,���,��������ID,�շ�ϸĿID,�����־,��ҳID,�վݷ�Ŀ,Ӥ����,���ʽ��,����ʱ�� From " & IIf(mblnNOMoved, "H", "") & "סԺ���ü�¼ A where A.����ID=[1] )  A"
    End Select
    
    strSql = _
    "   Select Decode(�����־,1,'����',4,'����','��'||Nvl(A.��ҳID,0)||'��') as סԺ," & _
    "         A.NO as ���ݺ�,Nvl(B.����,'δ֪') as ��������,Nvl(E.����,D.����) as ��Ŀ," & IIf(gTy_System_Para.bytҩƷ������ʾ = 2, "E1.���� as ��Ʒ��,", "") & _
    "       A.�վݷ�Ŀ as ��Ŀ,Decode(Nvl(A.Ӥ����,0),0,'','��') as Ӥ����," & _
    "       To_Char(" & IIf(blnDel, "-1*", "") & "A.���ʽ��,'999999999" & strDec & "') as ���ʽ��," & _
    "       To_Char(A.����ʱ��,'YYYY-MM-DD HH24:MI:SS') as ����ʱ��" & _
    " From " & strSql & ",���ű� B,�շ���ĿĿ¼ D,�շ���Ŀ���� E" & _
            IIf(gTy_System_Para.bytҩƷ������ʾ = 2, ",�շ���Ŀ���� E1", "") & _
    " Where A.��������ID=B.ID And A.�շ�ϸĿID=D.ID" & _
    "       And A.�շ�ϸĿID=E.�շ�ϸĿID(+) And E.����(+)=1 And E.����(+)=" & IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1) & _
            IIf(gTy_System_Para.bytҩƷ������ʾ = 2, "       And A.�շ�ϸĿID=E1.�շ�ϸĿID(+) And E1.����(+)=1 And E1.����(+)=3", "") & _
    "       And A.����ID=[1]" & _
    " Order by סԺ Desc,����ʱ�� Desc,���ݺ� Desc,A.���"
    Set rsDetail = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngBalanceID)
    
    vsfDetail.Clear 1
    vsfDetail.Rows = 2
    If Not rsDetail.EOF Then
        Set vsfDetail.DataSource = rsDetail
    End If
    Call SetDetail
    
End Sub

Private Sub SetDetail()
    Dim strHead As String
    Dim i As Long
    Dim varData As Variant

    strHead = "סԺ,4,750|���ݺ�,4,850|��������,1,850|��Ŀ,1,1800" & IIf(gTy_System_Para.bytҩƷ������ʾ = 2, "|��Ʒ��,1,1600", "") & "|��Ŀ,1,850|Ӥ����,4,650|���ʽ��,7,850|����ʱ��,1,1850"
    
    With vsfDetail
        .HighLight = flexHighlightWithFocus
        .Redraw = False
        varData = Split(strHead, "|")
        .Cols = UBound(varData) + 1
        For i = 0 To UBound(varData)
            If Split(varData(i), ",")(0) = "ID" Then .ColHidden(i) = True
            .TextMatrix(0, i) = Split(varData(i), ",")(0)
            .ColKey(i) = .TextMatrix(0, i)
            .ColAlignment(i) = Split(varData(i), ",")(1)
            .ColWidth(i) = Split(varData(i), ",")(2)
            .FixedAlignment(i) = 4
        Next
        
        Call RestoreFlexState(vsfDetail, App.ProductName & "\" & Me.Name)
        
        For i = 0 To .Cols - 1
            If .TextMatrix(0, i) = "��Ʒ��" Then
                If gTy_System_Para.bytҩƷ������ʾ = 2 Then
                    If .ColWidth(i) <= 0 Then .ColWidth(i) = 1600
                Else
                    .ColWidth(i) = 0
                End If
            End If
        Next
        
        .RowHeight(0) = 350
        If .Rows < 2 Then .Rows = 2
        .Row = 1: .Col = 0: .ColSel = .Cols - 1
        
        For i = 1 To .Rows - 1
            .RowHeight(i) = 300
        Next i
        
        .Redraw = True
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
        .Cols = UBound(varData) + 1
        For i = 0 To UBound(varData)
            .TextMatrix(0, i) = Split(varData(i), ",")(0)
            .ColKey(i) = .TextMatrix(0, i)
            .ColAlignment(i) = Split(varData(i), ",")(1)
            .ColWidth(i) = Split(varData(i), ",")(2)
            .FixedAlignment(i) = 4
        Next
        For i = 1 To .Rows - 1
            .TextMatrix(i, .ColIndex("������")) = Formatex(Val(.TextMatrix(i, .ColIndex("������"))), 6, , , 2)
            .RowHeight(i) = 300
        Next i
        
        Call RestoreFlexState(vsfBalance, App.ProductName & "\" & Me.Name)
        
        .RowHeight(0) = 350
        
        .Row = 1: .Col = 0: .ColSel = .Cols - 1

        .Redraw = True
    End With
End Sub

Private Function GetTotal() As Double
    Dim dblTotal As Double
    Dim i As Integer
    With vsfMain
        For i = 1 To .Rows - 1
            dblTotal = dblTotal + Val(.TextMatrix(i, .ColIndex("���ʽ��")))
        Next i
    End With
    GetTotal = dblTotal
End Function

Private Sub SetMain()
    Dim i As Long, strHead As String
    Dim dblTotal As Double
    
    strHead = "����ID,1,0|��־,1,0|��������,4,800|ҽ��,4,500|���ݺ�,4,850|Ʊ�ݺ�,4,850|����ID,1,750|�����,1,750|סԺ��,1,750|����,4,800|�Ա�,4,500|����,4,500|�ѱ�,4,750|��ʼ����,4,1000|��������,4,1000|���ʽ��,7,850|����Ա,4,800|�շ�ʱ��,4,1850|��;����,4,800|��¼״̬,1,0"
    vsfDetail.Clear 1
    vsfDetail.Rows = 2
    vsfBalance.Clear 1
    vsfBalance.Rows = 2
    With vsfMain
        .Cols = UBound(Split(strHead, "|")) + 1
        For i = 0 To UBound(Split(strHead, "|"))
            .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
            .ColKey(i) = .TextMatrix(0, i)
            If .TextMatrix(0, i) = "����ID" Then
                .ColHidden(i) = True
            Else
                .ColHidden(i) = False
            End If
            .ColAlignment(i) = Split(Split(strHead, "|")(i), ",")(1)
            .ColWidth(i) = Split(Split(strHead, "|")(i), ",")(2)
            .FixedAlignment(i) = 4
            If .ColKey(i) = "����ID" Or .ColKey(i) = "��־" Or .ColKey(i) = "��¼״̬" Then .ColData(i) = "-1|1"
            If .ColKey(i) = "���ݺ�" Or .ColKey(i) = "�շ�ʱ��" Then .ColData(i) = "1|0"
        Next
        
        .RowHeight(0) = 350
        If .Rows = 1 Then .Rows = 2
        For i = 1 To .Rows - 1
            .RowHeight(i) = 300
            .TextMatrix(i, .ColIndex("���ʽ��")) = Format(.TextMatrix(i, .ColIndex("���ʽ��")), gstrDec)
        Next i
        
        If .Rows >= 2 Then .Select 1, 1
        If .Enabled And .Visible Then .SetFocus
    End With
End Sub

Private Sub Form_Load()
    Call SetDockingPanel
    Call SetMain
    Call SetBalance
    Call SetDetail
End Sub

Private Sub picDetail_Resize()
    With vsfDetail
        .Top = 0
        .Left = 0
        .Height = picDetail.Height
        .Width = picDetail.Width
    End With
End Sub

Private Sub vsfMain_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    On Error Resume Next
    If OldRow <> 0 And NewRow <> 0 Then zl_VsGridRowChange vsfMain, OldRow, NewRow, OldCol, NewCol
    If vsfMain.TextMatrix(NewRow, vsfMain.ColIndex("����ID")) = "" Then Exit Sub
    Call ReadDetail(Val(vsfMain.TextMatrix(NewRow, vsfMain.ColIndex("����ID"))))
    Call ReadBalance(Val(vsfMain.TextMatrix(NewRow, vsfMain.ColIndex("����ID"))))
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
    Call frmManageBalance.ViewBalance(1)
End Sub

Public Sub zlRptPrint(ByVal bytFunc As Integer)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���д�ӡ,Ԥ���������EXCEL
    '���:bytFunc=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    '����:������
    '����:2013-09-12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, r As Long, lngRow As Long, intActive As Integer
    Dim intCol As Long, objPrint As Object, objRow As New zlTabAppRow, bytPrn As Byte
    Dim vsBill As Object, strTittle As String
    
    lngRow = vsfMain.Row
    Set vsBill = vsfMain: strTittle = GetUnitName & "�쳣���ʼ�¼��Ϣ"
    mblnPrint = True
    
    Set objPrint = New zlPrint1Grd
    objPrint.Title.Font.Name = "����_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    objPrint.Title.Text = strTittle
    
    objPrint.UnderAppRows.Add objRow
    Set objRow = New zlTabAppRow
    objRow.Add "��ӡ��:" & UserInfo.����
    objRow.Add "��ӡ����:" & Format(zlDatabase.Currentdate, "yyyy��MM��dd��")
    objPrint.BelowAppRows.Add objRow
    
    If vsBill Is Nothing Then Exit Sub
    '���ڴ�ӡ�ؼ�����ʶ������������
    With vsBill
        .Redraw = flexRDNone
        .GridColor = .ForeColor
        For i = 0 To .Cols - 1
            .Cell(flexcpData, 0, i) = .ColWidth(i)
            If .ColHidden(i) = True Or i = 0 Then
                .ColWidth(i) = 0
            End If
        Next
    End With
    
    Err = 0: On Error GoTo ErrHand:
    Set objPrint.Body = vsBill
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
    With vsBill
        For i = 0 To .Cols - 1
            .Cell(flexcpData, 0, i) = .ColWidth(i)
            If .ColHidden(i) = True Or i = 0 Then
                .ColWidth(i) = Val(.Cell(flexcpData, 0, i))
            End If
        Next
        .GridColor = &H8000000F
        .Redraw = flexRDBuffered
    End With
    
    mblnPrint = False
    vsfMain.Select lngRow, 1
    Exit Sub
ErrHand:
    mblnPrint = False
    vsfMain.Select lngRow, 1
    If ErrCenter = 1 Then Resume
End Sub
