VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmPacsSrvSet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   9120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11625
   Icon            =   "frmPacsSrvSet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9120
   ScaleWidth      =   11625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton CmdCancel 
      Caption         =   "�˳�(&C)"
      Height          =   350
      Left            =   8640
      TabIndex        =   5
      Top             =   8550
      Width           =   1125
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "����(&S)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   6360
      TabIndex        =   4
      Top             =   8550
      Width           =   1125
   End
   Begin VB.Frame FraList 
      Caption         =   "�����б�"
      Height          =   1860
      Left            =   0
      TabIndex        =   1
      ToolTipText     =   "�б�������<Ӱ���豸Ŀ¼>�е�����"
      Top             =   0
      Width           =   11505
      Begin VB.CommandButton CmdDel 
         Caption         =   "ɾ��(&D)"
         Height          =   350
         Left            =   10200
         TabIndex        =   3
         Top             =   645
         Width           =   1100
      End
      Begin VB.CommandButton CmdOK 
         Caption         =   "����(&S)"
         Height          =   350
         Left            =   10200
         TabIndex        =   2
         Top             =   210
         Width           =   1100
      End
      Begin VSFlex8Ctl.VSFlexGrid vfgList 
         Height          =   1560
         Left            =   120
         TabIndex        =   6
         Top             =   210
         Width           =   9990
         _cx             =   17621
         _cy             =   2752
         Appearance      =   0
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
         Rows            =   3
         Cols            =   10
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
         AutoResize      =   0   'False
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
   Begin XtremeSuiteControls.TabControl TabList 
      Height          =   7080
      Left            =   0
      TabIndex        =   0
      Top             =   1905
      Width           =   11535
      _Version        =   589884
      _ExtentX        =   20346
      _ExtentY        =   12488
      _StockProps     =   64
   End
End
Attribute VB_Name = "frmPacsSrvSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum ColList
        ColID = 0
        Col������
        Col������
        ColPACS��ɫ
        Col����IP
        Col����AE
        Col����˿�
        Col�豸IP
        Col�豸AE
        Col�豸�˿�
End Enum

Private mFrmImgSrv As New frmImgSrv
Private mfrmWorkList As New frmWorklist
Private mFrmQRSrv As New frmQrSrv
Private mDevNo As String, mDevIP As String
Private mblnNeedSaveSrv As Boolean, mblnInitOk As Boolean
Public Sub ShowMe(ByVal DevNo As String, ByVal DevName As String, ByVal DevIP As String, ByVal frmobj As Object)
    mDevNo = DevNo
    mDevIP = DevIP
    mblnNeedSaveSrv = False
    Me.Caption = "�豸(" & DevName & ")" & "��������"
    Me.Show , frmobj
End Sub
Private Sub InitFaceScheme()
'��ʼ���沼��
    FraList.Top = 0
    FraList.Left = 0
    With TabList
        .Top = FraList.Top + FraList.Height
        .Left = FraList.Left
        .Height = Me.ScaleHeight - FraList.Height
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.Color = xtpTabColorOffice2003
        .PaintManager.ClientFrame = xtpTabFrameBorder
        .PaintManager.Position = xtpTabPositionTop
        .PaintManager.OneNoteColors = False
        .PaintManager.BoldSelected = True
        .InsertItem 1, "ͼ����շ���", mFrmImgSrv.hWnd, 0
        .InsertItem 2, "WorkList����", mfrmWorkList.hWnd, 0
        .InsertItem 3, "Q/R ��ѯ����", mFrmQRSrv.hWnd, 0
        
        .Item(0).Enabled = False
        .Item(1).Enabled = False
        .Item(2).Enabled = False
    End With
    cmdCancel.Top = TabList.Top + TabList.Height - cmdCancel.Height - 50
    cmdSave.Top = TabList.Top + TabList.Height - cmdSave.Height - 50
End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Sub cmdDel_Click()
    If vfgList.TextMatrix(vfgList.Row, Col������) = "" Then Exit Sub
    If MsgBoxD(Me, "ȷʵҪɾ������(" & vfgList.TextMatrix(vfgList.Row, Col������) & ")��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    If vfgList.TextMatrix(vfgList.Row, ColID) <> "" Then
        gstrSQL = "Zl_Ӱ��DICOM�����_DELETE(" & vfgList.TextMatrix(vfgList.Row, ColID) & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "ɾ������")
    End If
    Call InitSrvList
End Sub
Private Function ValidData() As Boolean
Dim i As Long, j As Integer

    With vfgList
        For i = 1 To .Rows - 1
            If i <> .Rows - 1 Then
                If .TextMatrix(i, Col������) = "" Then MsgBoxD Me, "��" & i & "�� ������ ����Ϊ��", vbInformation, gstrSysName: Exit Function
                If .TextMatrix(i, Col������) = "" Then MsgBoxD Me, "��" & i & "�� ������ ����Ϊ��", vbInformation, gstrSysName: Exit Function
                If UBound(Split(Trim(.TextMatrix(i, Col����IP)), ".")) <> 3 Then
                    MsgBoxD Me, "��" & i & "�� ����IP��ʽ����ȷ�����飡", vbInformation, gstrSysName: Exit Function
                Else
                    For j = 0 To 3
                        If Not IsNumeric(Split(Trim(.TextMatrix(i, Col����IP)), ".")(j)) Then
                            MsgBoxD Me, "��" & i & "�� ����IP��ʽ����ȷ�����飡", vbInformation, gstrSysName: Exit Function
                        Else
                            If Split(Trim(.TextMatrix(i, Col����IP)), ".")(j) < 0 Or Split(Trim(.TextMatrix(i, Col����IP)), ".")(j) >= 256 Then
                                MsgBoxD Me, "��" & i & "�� ����IP��ʽ����ȷ�����飡", vbInformation, gstrSysName: Exit Function
                            End If
                        End If
                    Next
                End If
                If .TextMatrix(i, Col����AE) = "" Then MsgBoxD Me, "��" & i & "�� ����AE ����Ϊ��", vbInformation, gstrSysName: Exit Function
                If .TextMatrix(i, Col����˿�) = "" Then MsgBoxD Me, "��" & i & "�� ���ض˿� ����Ϊ��", vbInformation, gstrSysName: Exit Function
                If Not IsNumeric(.TextMatrix(i, Col����˿�)) Then MsgBoxD Me, "��" & i & "�� ���ض˿� ����Ϊ��ֵ", vbInformation, gstrSysName: Exit Function
                If .TextMatrix(i, Col�豸AE) = "" Then MsgBoxD Me, "��" & i & "�� �豸AE ����Ϊ��", vbInformation, gstrSysName: Exit Function
            Else
                If .TextMatrix(i, Col������) <> "" Then
                    If .TextMatrix(i, Col������) = "" Then MsgBoxD Me, "��" & i & "�� ������ ����Ϊ��", vbInformation, gstrSysName: Exit Function
                    If UBound(Split(Trim(.TextMatrix(i, Col����IP)), ".")) <> 3 Then
                        MsgBoxD Me, "��" & i & "�� ����IP��ʽ����ȷ�����飡", vbInformation, gstrSysName: Exit Function
                    Else
                        For j = 0 To 3
                            If Not IsNumeric(Split(Trim(.TextMatrix(i, Col����IP)), ".")(j)) Then
                                MsgBoxD Me, "��" & i & "�� ����IP��ʽ����ȷ�����飡", vbInformation, gstrSysName: Exit Function
                            Else
                                If Split(Trim(.TextMatrix(i, Col����IP)), ".")(j) < 0 Or Split(Trim(.TextMatrix(i, Col����IP)), ".")(j) >= 256 Then
                                    MsgBoxD Me, "��" & i & "�� ����IP��ʽ����ȷ�����飡", vbInformation, gstrSysName: Exit Function
                                End If
                            End If
                        Next
                    End If
                    If .TextMatrix(i, Col����AE) = "" Then MsgBoxD Me, "��" & i & "�� ����AE ����Ϊ��", vbInformation, gstrSysName: Exit Function
                    If .TextMatrix(i, Col����˿�) = "" Then MsgBoxD Me, "��" & i & "�� ���ض˿� ����Ϊ��", vbInformation, gstrSysName: Exit Function
                    If Not IsNumeric(.TextMatrix(i, Col����˿�)) Then MsgBoxD Me, "��" & i & "�� ���ض˿� ����Ϊ��ֵ", vbInformation, gstrSysName: Exit Function
                    If .TextMatrix(i, Col�豸AE) = "" Then MsgBoxD Me, "��" & i & "�� �豸AE ����Ϊ��", vbInformation, gstrSysName: Exit Function
                ElseIf i = 1 Then
                    Exit Function
                End If
            End If
        Next
    End With
    ValidData = True
End Function
Private Sub cmdOK_Click()
Dim i As Long, Count As Long
    If Not ValidData Then Exit Sub
    On Error GoTo errHandle
    If Trim(vfgList.TextMatrix(vfgList.Rows - 1, Col������)) = "" Then
        Count = vfgList.Rows - 2
    Else
        Count = vfgList.Rows - 1
    End If
    
    For i = 1 To Count
        With vfgList
            Select Case .TextMatrix(i, ColID)
                Case ""
                    gstrSQL = "Zl_Ӱ��DICOM�����_INSERT('" & Trim(.TextMatrix(i, Col������)) & "','" & mDevNo & "','" & _
                                            .TextMatrix(i, Col������) & "','" & .TextMatrix(i, ColPACS��ɫ) & "','" & _
                                            Trim(.TextMatrix(i, Col����IP)) & "','" & Trim(.TextMatrix(i, Col����AE)) & "','" & _
                                            Trim(.TextMatrix(i, Col����˿�)) & "','" & mDevIP & "','" & _
                                            Trim(.TextMatrix(i, Col�豸AE)) & "','" & Trim(.TextMatrix(i, Col����˿�)) & "')"
                Case Else
                    gstrSQL = "Zl_Ӱ��DICOM�����_UPDATE(" & .TextMatrix(i, ColID) & ",'" & Trim(.TextMatrix(i, Col������)) & "','" & mDevNo & "','" & _
                                            .TextMatrix(i, Col������) & "','" & .TextMatrix(i, ColPACS��ɫ) & "','" & _
                                            Trim(.TextMatrix(i, Col����IP)) & "','" & Trim(.TextMatrix(i, Col����AE)) & "','" & _
                                            Trim(.TextMatrix(i, Col����˿�)) & "','" & mDevIP & "','" & _
                                            Trim(.TextMatrix(i, Col�豸AE)) & "','" & Trim(.TextMatrix(i, Col����˿�)) & "')"
            End Select
            Call zlDatabase.ExecuteProcedure(gstrSQL, "�������")
        End With
    Next
    Call InitSrvList(Trim(vfgList.TextMatrix(vfgList.Row, Col������)))
    mblnNeedSaveSrv = False
    MsgBoxD Me, "���񱣴�ɹ�����Ϊ�÷����趨������", vbInformation, gstrSysName
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cmdSave_Click()
    If mblnNeedSaveSrv Then MsgBoxD Me, "�Ϸ������б����б䶯��δ���棬���ȱ�������б�䶯", vbInformation, gstrSysName: Exit Sub
    Select Case TabList.Selected.Caption
        Case "ͼ����շ���"
            Call mFrmImgSrv.SavePara
        Case "WorkList����"
            Call mfrmWorkList.SavePara
        Case "Q/R ��ѯ����"
            Call mFrmQRSrv.SavePara
    End Select
    MsgBoxD Me, "��������ɹ�", vbInformation, gstrSysName
End Sub

Private Sub Form_Load()
    InitFaceScheme '��ʼ������
    InitSrvList '��ʼ�������б�
End Sub
Private Sub InitSrvList(Optional ByVal strSrvName As String)
    mblnInitOk = False
    With vfgList
        .Clear
        .Rows = 2
        .Cols = 10
        .ColWidth(ColID) = 0 'ID
        .ColWidth(Col������) = 1000 '������
        .ColWidth(Col������) = 1200 '������
        .ColWidth(ColPACS��ɫ) = 500  'PACS��ɫ SUC/SUP
        .ColWidth(Col����IP) = 1400 '����IP
        .ColWidth(Col����AE) = 1000  '����AE
        .ColWidth(Col����˿�) = 500 '����˿�
        .ColWidth(Col�豸IP) = 0
        .ColWidth(Col�豸AE) = 1000
        .ColWidth(Col�豸�˿�) = 0
        .TextMatrix(0, ColID) = "ID"
        .TextMatrix(0, Col������) = "������"
        .TextMatrix(0, Col������) = "������"
        .TextMatrix(0, ColPACS��ɫ) = "��ɫ"
        .TextMatrix(0, Col����IP) = "����IP"
        .TextMatrix(0, Col����AE) = "����AE"
        .TextMatrix(0, Col����˿�) = "�˿�"
        .TextMatrix(0, Col�豸IP) = "�豸IP"
        .TextMatrix(0, Col�豸AE) = "�豸AE"
        .TextMatrix(0, Col�豸�˿�) = "�豸�˿�"
        

        .FixedAlignment(ColID) = flexAlignCenterCenter
        .FixedAlignment(Col������) = flexAlignCenterCenter
        .FixedAlignment(Col������) = flexAlignCenterCenter
        .FixedAlignment(ColPACS��ɫ) = flexAlignCenterCenter
        .FixedAlignment(Col����IP) = flexAlignCenterCenter
        .FixedAlignment(Col����AE) = flexAlignCenterCenter
        .FixedAlignment(Col����˿�) = flexAlignCenterCenter
        .FixedAlignment(Col�豸IP) = flexAlignCenterCenter
        .FixedAlignment(Col�豸AE) = flexAlignCenterCenter
        .FixedAlignment(Col�豸�˿�) = flexAlignCenterCenter
        
        .ColAlignment(ColID) = flexAlignLeftCenter
        .ColAlignment(Col������) = flexAlignLeftCenter
        .ColAlignment(Col������) = flexAlignLeftCenter
        .ColAlignment(ColPACS��ɫ) = flexAlignLeftCenter
        .ColAlignment(Col����IP) = flexAlignLeftCenter
        .ColAlignment(Col����AE) = flexAlignLeftCenter
        .ColAlignment(Col����˿�) = flexAlignLeftCenter
        .ColAlignment(Col�豸IP) = flexAlignLeftCenter
        .ColAlignment(Col�豸AE) = flexAlignLeftCenter
        .ColAlignment(Col�豸�˿�) = flexAlignLeftCenter

        .Editable = flexEDKbdMouse
        .ColComboList(Col������) = "ͼ�����|Worklist|Q/R����|��Ƭ����"
    End With
    Call FillBILL(strSrvName)
    mblnInitOk = True
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Unload mFrmImgSrv
    Unload mfrmWorkList
    Unload mFrmQRSrv
End Sub
Private Sub FillBILL(ByVal strSrvName As String)
Dim rsTemp As ADODB.Recordset, i As Integer
    On Error GoTo errHandle
    With vfgList
        gstrSQL = "select B.�豸IP��ַ,B.�豸AE����,B.�豸�˿�,B.����ID,B.������,B.������,B.PACS��ɫ,B.PACSIP��ַ,B.PACSAE����,B.PACS�˿�" & _
                    " from Ӱ���豸Ŀ¼ A,Ӱ��DICOM����� B where A.�豸��=[1] and A.�豸��=B.�豸��"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ������Ϣ", CStr(mDevNo))
        Do Until rsTemp.EOF
            .TextMatrix(rsTemp.AbsolutePosition, ColID) = rsTemp!����ID
            .TextMatrix(rsTemp.AbsolutePosition, Col������) = Nvl(rsTemp!������)
            .TextMatrix(rsTemp.AbsolutePosition, Col������) = Nvl(rsTemp!������)
            .TextMatrix(rsTemp.AbsolutePosition, ColPACS��ɫ) = Nvl(rsTemp!PACS��ɫ)
            .TextMatrix(rsTemp.AbsolutePosition, Col����IP) = Nvl(rsTemp!PACSIP��ַ)
            .TextMatrix(rsTemp.AbsolutePosition, Col����AE) = Nvl(rsTemp!PACSAE����)
            .TextMatrix(rsTemp.AbsolutePosition, Col����˿�) = Nvl(rsTemp!PACS�˿�)
            .TextMatrix(rsTemp.AbsolutePosition, Col�豸IP) = Nvl(rsTemp!�豸IP��ַ)
            .TextMatrix(rsTemp.AbsolutePosition, Col�豸AE) = Nvl(rsTemp!�豸AE����)
            .TextMatrix(rsTemp.AbsolutePosition, Col�豸�˿�) = Nvl(rsTemp!�豸�˿�)
            .Rows = .Rows + 1
            rsTemp.MoveNext
        Loop
        If strSrvName <> "" Then
            For i = 1 To .Rows - 1
                If .TextMatrix(i, Col������) = strSrvName Then .Row = i: .RowSel = i: .Col = 0
            Next
        Else
            .Row = 1
            vfgList_EnterCell
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub vfgList_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If mblnInitOk = True Then '��ʼ����ɲŴ���,��ʼ���ڼ䴥����Ч
        mblnNeedSaveSrv = True
        If Col = Col������ Then
            Select Case vfgList.TextMatrix(Row, Col)
                Case "ͼ�����"
                    vfgList.TextMatrix(vfgList.Row, ColPACS��ɫ) = "SCP"
                Case "Worklist"
                    vfgList.TextMatrix(vfgList.Row, ColPACS��ɫ) = "SCP"
                Case "Q/R����"
                    vfgList.TextMatrix(vfgList.Row, ColPACS��ɫ) = "SCU"
                Case "��Ƭ����"
                    vfgList.TextMatrix(vfgList.Row, ColPACS��ɫ) = "SCP"
            End Select
        End If
    End If
End Sub

Private Sub vfgList_DblClick()
    With vfgList
        If .Col = ColPACS��ɫ Then
            mblnNeedSaveSrv = True
            .Editable = flexEDNone
            If (.TextMatrix(.Row, Col������) <> "��Ƭ����") Then
                If .TextMatrix(.Row, .Col) = "SCP" Then
                    .TextMatrix(.Row, .Col) = "SCU"
                Else
                    .TextMatrix(.Row, .Col) = "SCP"
                End If
            End If
        Else
            .Editable = flexEDKbdMouse
        End If
    End With
End Sub

Private Sub vfgList_EnterCell()
Dim i As Long

    If Trim(vfgList.TextMatrix(vfgList.Row, Col������) = "") And vfgList.Rows <= 2 Then
        TabList.Enabled = False
    Else
        TabList.Enabled = True
    End If
    TabList.Visible = True
    Select Case vfgList.TextMatrix(vfgList.Row, Col������)
        Case "ͼ�����", "��Ƭ����"
            TabList.Item(0).Enabled = True
            TabList.Item(0).Visible = True
            TabList.Item(0).Selected = True
            TabList.Item(1).Enabled = False
            TabList.Item(2).Enabled = False
            TabList.Item(1).Visible = False
            TabList.Item(2).Visible = False
            cmdSave.Enabled = True
            Call mFrmImgSrv.ShowRefresh(IIf(vfgList.TextMatrix(vfgList.Row, ColID) = "", 0, vfgList.TextMatrix(vfgList.Row, ColID)))
        Case "Worklist"
            TabList.Item(0).Enabled = False
            TabList.Item(0).Visible = False
            TabList.Item(1).Enabled = True
            TabList.Item(1).Visible = True
            TabList.Item(1).Selected = True
            TabList.Item(2).Enabled = False
            TabList.Item(2).Visible = False
            cmdSave.Enabled = True
            Call mfrmWorkList.ShowRefresh(IIf(vfgList.TextMatrix(vfgList.Row, ColID) = "", 0, vfgList.TextMatrix(vfgList.Row, ColID)))
        Case "Q/R����"
            TabList.Item(0).Enabled = False
            TabList.Item(0).Visible = False
            TabList.Item(1).Enabled = False
            TabList.Item(1).Visible = False
            TabList.Item(2).Selected = True
            TabList.Item(2).Enabled = True
            TabList.Item(2).Visible = True
            cmdSave.Enabled = True
            Call mFrmQRSrv.ShowRefresh(IIf(vfgList.TextMatrix(vfgList.Row, ColID) = "", 0, vfgList.TextMatrix(vfgList.Row, ColID)))
        Case Else
            TabList.Enabled = False
            TabList.Visible = False
            cmdSave.Enabled = False
    End Select
End Sub

Private Sub vfgList_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Col = Col����IP Or Col = Col����˿� Then
        If InStr("0123456789." & Chr(8) & Chr(13), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    ElseIf Col = Col������ Or Col = ColPACS��ɫ Then
        If KeyAscii <> 13 Then KeyAscii = 0
    End If
End Sub
