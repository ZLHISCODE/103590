VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMedicalStationReport 
   BackColor       =   &H8000000A&
   BorderStyle     =   0  'None
   Caption         =   "��챨��"
   ClientHeight    =   5250
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7365
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   5250
   ScaleWidth      =   7365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picContainer 
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000C&
      Height          =   3135
      Left            =   135
      ScaleHeight     =   3075
      ScaleWidth      =   5400
      TabIndex        =   1
      Top             =   1860
      Width           =   5460
   End
   Begin VSFlex8Ctl.VSFlexGrid vsf 
      Height          =   1530
      Left            =   135
      TabIndex        =   0
      Top             =   120
      Width           =   5430
      _cx             =   9578
      _cy             =   2699
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
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   12698049
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   2
      HighLight       =   0
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   255
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
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
      Begin VB.Line lnY 
         Index           =   1
         Visible         =   0   'False
         X1              =   270
         X2              =   270
         Y1              =   420
         Y2              =   1635
      End
      Begin VB.Line lnX 
         Index           =   1
         Visible         =   0   'False
         X1              =   -555
         X2              =   1230
         Y1              =   555
         Y2              =   555
      End
   End
   Begin MSComctlLib.ImageList ils13 
      Left            =   6345
      Top             =   810
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationReport.frx":0000
            Key             =   "����"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationReport.frx":039A
            Key             =   "����"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationReport.frx":0734
            Key             =   "״̬"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationReport.frx":0ACE
            Key             =   "����"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationReport.frx":0E68
            Key             =   "����"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationReport.frx":1202
            Key             =   "up"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationReport.frx":13C4
            Key             =   "down"
         EndProperty
      EndProperty
   End
   Begin VB.Image imgX 
      Height          =   135
      Left            =   1080
      MousePointer    =   7  'Size N S
      Top             =   1695
      Width           =   5115
   End
End
Attribute VB_Name = "frmMedicalStationReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'���������弶��������**************************************************************************************************
Private mblnStartUp As Boolean
Private mfrmReport As Object
Private mclsCore As New clsCISCore
Private mlngKey As Long
Private mfrmMain As Object
Private mvarParam As Variant
Private mblnNoAllowChange As Boolean
Private mblnDataMoved As Boolean
Private mblnShow As Boolean         '�Ƿ���ʾ����

Private Enum mCol
    ����
    ״̬
    ��Ŀ
    ִ�п���
    ִ��״̬
    ������
    ʱ��
    ����id
    ����id
    No
    ����;��
End Enum

Public Function zlMenuClick(ByVal frmMain As Object, ByVal strMenuItem As String, Optional ByVal strParam As String = "") As Boolean
    '--------------------------------------------------------------------------------------------------------
    '���ܣ�
    '������lngKey ����ID
    '--------------------------------------------------------------------------------------------------------
    Dim lngSvrKey As Long
    Dim strNO As String
    Dim lng����id As Long
    Dim lng����id As Long
    Dim lng��¼���� As Long
    
    On Error GoTo errHand
    
    mvarParam = Split(strParam, "'")
    
    mlngKey = Val(mvarParam(0))
    
    Set mfrmMain = frmMain
    
    Select Case strMenuItem
    Case "ˢ��"
        
        lngSvrKey = Val(vsf.RowData(vsf.Row))
        Call zlClearData
        Call RefreshData(strMenuItem)
        Call RestoreRow(vsf, lngSvrKey)
        Call vsf_AfterRowColChange(0, 0, vsf.Row, vsf.Col)
        
    Case "��д����", "�鿴����"
        
        If Val(vsf.RowData(vsf.Row)) <= 0 Then Exit Function
        
        strNO = vsf.TextMatrix(vsf.Row, GetCol(vsf, "No"))
        lng����id = Val(vsf.TextMatrix(vsf.Row, GetCol(vsf, "����id")))
        lng����id = Val(vsf.TextMatrix(vsf.Row, GetCol(vsf, "����id")))
        lng��¼���� = IIf(Val(vsf.TextMatrix(vsf.Row, GetCol(vsf, "����;��"))) = 1, 2, 1)
        
        If strNO = "" Then Exit Function
        If lng����id = 0 And lng����id = 0 Then Exit Function
                
        Call EditReport(frmMain, strNO, lng��¼����, lng����id, lng����id, "", IIf(strMenuItem = "��д����", False, True), True, , , , False, , mblnDataMoved, "001")
                            
        '�˳������ˢ��
        mblnNoAllowChange = True
        
        lngSvrKey = Val(vsf.RowData(vsf.Row))
        Call zlClearData
        Call RefreshData("ˢ��")
        Call RestoreRow(vsf, lngSvrKey)
        
        mblnNoAllowChange = False
        
        Call vsf_AfterRowColChange(0, 0, vsf.Row, vsf.Col)

    End Select
    
    zlMenuClick = True
    
    Exit Function
    
errHand:
    'If ErrCenter = 1 Then Resume
End Function

Public Sub zlClearData(Optional ByVal strPart As String = "����")
    '--------------------------------------------------------------------------------------------------
    '���ܣ�
    '������
    '--------------------------------------------------------------------------------------------------
    Dim blnSvr As Boolean
    
    blnSvr = mblnNoAllowChange
    
    mblnNoAllowChange = True
    
    Call ResetVsf(vsf)
    Call AppendSapceRows(vsf, lnX, lnY)
        
    On Error Resume Next
    If Not (mfrmReport Is Nothing) Then mfrmReport.zlClearData
    
    mblnNoAllowChange = blnSvr
End Sub

Public Property Get Body(ByVal lngIndex As Long) As Object
    Set Body = vsf
End Property

Public Property Let ShowResult(ByVal v_Data As Boolean)
    
    mblnShow = v_Data

    picContainer.Visible = mblnShow
    imgX.Visible = mblnShow
    
    Call Form_Resize
        
    If mblnShow Then
        
        If mfrmReport Is Nothing Then Set mfrmReport = mclsCore.ShowFileObject(Me, Me.picContainer, 0, 0, gcnOracle, "", glngSys, "", "")
                
        Call RefreshData("����")
        
    Else

        On Error Resume Next
        
        If Not (mfrmReport Is Nothing) Then mfrmReport.zlClearData
        
    End If
    
End Property

Private Function RefreshData(ByVal strMenu As String) As Boolean
    
    Dim rs As New ADODB.Recordset
    
    Select Case strMenu
    Case "ˢ��"
        If Val(mvarParam(1)) = 0 Then
            'δ��ʼ֮ǰ�Ĳ�ѯ
            mblnDataMoved = False
            gstrSQL = GetPublicSQL(SQL.����������Ŀ)
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mvarParam(0)))
        Else
            
            gstrSQL = "Select X. *, " & _
                               "Y.���� As ִ�п���, " & _
                               "Z.���� As ��Ŀ, " & _
                               "Decode(X.����id, Null, Decode(D.�����ļ�id, Null, '', '����'), Decode(H.��д��, Null, '����', '����')) As ״̬, " & _
                               "D.�����ļ�id As ����id, " & _
                               "H.��д�� As ������, " & _
                               "To_Char(H.��д����, 'yyyy-mm-dd hh24:mi') As ʱ��,Decode(x.�����嵥id,0,0,Null,0,255) As ǰ��ɫ " & _
                        "From ( Select E.ID, " & _
                                      "B.ִ�п���id,b.�����嵥id, " & _
                                      "A.������Ŀid, " & _
                                      "A.����;��, " & _
                                      "Decode(G.ִ��״̬, 1, '��ȫִ��', 2, 'ȡ��ִ��', 3, '����ִ��', '') As ִ��״̬, G.����id, G.NO, " & _
                                      "Decode(A.����id, Null, '', '����') As ���� " & _
                               "From �����Ŀҽ�� B, �����Ŀ�嵥 A, �����Ա���� C, ���ǼǼ�¼ D, ����ҽ����¼ E, ����ҽ������ G " & _
                               "Where A.ID = B.�嵥id " & _
                                     "And B.����id = C.����id " & _
                                     "And C.�Ǽ�id = A.�Ǽ�id " & _
                                     "AND D.ID=C.�Ǽ�id And d.����=E.�Һŵ� And e.����id=c.����id " & _
                                     "AND E.������Դ=4 " & _
                                     "AND E.ҽ��״̬<>4 " & _
                                     "And E.������� In ('C', 'D') " & _
                                     "And G.ҽ��id = E.ID And b.ҽ��id In (e.ID,e.���id) "
            gstrSQL = gstrSQL & _
                                     "And C.ID = [1] " & _
                               ") X, ���ű� Y, ������ĿĿ¼ Z, ���Ƶ���Ӧ�� D, ���˲�����¼ H " & _
                        "Where x.ִ�п���id = y.ID " & _
                              "And Z.ID = X.������Ŀid " & _
                              "And X.����id = H.ID(+) " & _
                              "And D.Ӧ�ó���(+) = 4 " & _
                              "And X.������Ŀid = D.������Ŀid(+) " & _
                        "Order By Y.����"
            
            '����ת������
            '----------------------------------------------------------------------------------------------------------
            mblnDataMoved = DataMove(mlngKey, 2)
            If mblnDataMoved Then
                gstrSQL = Replace(gstrSQL, "�����Ŀҽ��", "H�����Ŀҽ��")
                gstrSQL = Replace(gstrSQL, "�����Ŀ�嵥", "H�����Ŀ�嵥")
                gstrSQL = Replace(gstrSQL, "�����Ա����", "H�����Ա����")
                gstrSQL = Replace(gstrSQL, "���ǼǼ�¼", "H���ǼǼ�¼")
                gstrSQL = Replace(gstrSQL, "����ҽ����¼", "H����ҽ����¼")
                gstrSQL = Replace(gstrSQL, "����ҽ������", "H����ҽ������")
                gstrSQL = Replace(gstrSQL, "���˲�����¼", "H���˲�����¼")
            End If
            
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngKey)
        End If
        
        If rs.BOF = False Then
            
            Call LoadGrid(vsf, rs, , , ils13)
            Call AppendSapceRows(vsf, lnX, lnY)
            
        End If
    
    Case "����"
    
        If Not (mfrmReport Is Nothing) Then Call mfrmReport.zlMenuClick(Me, Val(vsf.TextMatrix(vsf.Row, GetCol(vsf, "����id"))), "ˢ��")
        
    End Select
    
End Function

Private Function InitLoad() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:��ʼ�����ݣ������ڴ����Load�¼�
    '------------------------------------------------------------------------------------------------------------------
    Dim strVsf As String
    
    On Error GoTo errHand
            
    picContainer.Visible = False
    imgX.Visible = False
    
    Set mfrmReport = Nothing
    
    strVsf = ",255,4,1,1,[����];,255,4,1,1,[״̬];��Ŀ,2400,1,1,1,;ִ�п���,1080,1,1,1,;ִ��״̬,900,1,1,1,;������,900,1,1,1,;ʱ��,1670,1,1,1,;����id,0,1,1,1,;����id,0,1,1,1,;No,0,1,1,1,;����;��,0,1,1,1,"
    
    Call CreateVsf(vsf, strVsf)
    vsf.Cols = vsf.Cols + 1
    vsf.ColWidth(vsf.Cols - 1) = 15
    
    Set vsf.Cell(flexcpPicture, 0, 0) = ils13.ListImages("����").Picture
    Set vsf.Cell(flexcpPicture, 0, 1) = ils13.ListImages("״̬").Picture
    
    Call InitCISCore(gcnOracle)
    
    Call AppendSapceRows(vsf, lnX, lnY)
        
    InitLoad = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
End Function

'���������弰��ؼ����¼�����******************************************************************************************
Private Sub Form_Activate()
    
    If mblnStartUp = False Then Exit Sub
    mblnStartUp = False
    
End Sub

Private Sub Form_Load()
    
    mblnStartUp = True
        
    Call InitLoad
       
        
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    If Me.WindowState = 1 Then Exit Sub
    If imgX.Top > Me.ScaleHeight - 1000 Then imgX.Top = Me.ScaleHeight - 1000
    
    With vsf
        .Left = 0
        .Top = 0
        .Width = Me.ScaleWidth
        .Height = IIf(mblnShow, imgX.Top, Me.ScaleHeight)
    End With
    
    If mblnShow Then
        With imgX
            .Left = vsf.Left
            .Width = vsf.Width
            .Height = 45
            .BorderStyle = 0
        End With
    
        With picContainer
            .Left = 0
            .Top = imgX.Top + imgX.Height
            .Width = vsf.Width
            .Height = Me.ScaleHeight - .Top
        End With
    End If
    
    Call AppendSapceRows(vsf, lnX, lnY)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mfrmReport = Nothing
End Sub

Private Sub imgX_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    
    On Error Resume Next
    
    imgX.Top = imgX.Top + Y
    
    If imgX.Top < 1500 Then imgX.Top = 1500
    If Me.Height - imgX.Top - imgX.Height < 1000 Then imgX.Top = Me.Height - imgX.Height - 1000
    
            
    Form_Resize
End Sub

Private Sub picContainer_Resize()
    On Error Resume Next
    
    If Not (mfrmReport Is Nothing) Then
        mfrmReport.Width = picContainer.Width
        mfrmReport.Height = picContainer.Height
    End If
    
End Sub

Private Sub vsf_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If mblnNoAllowChange Then Exit Sub
    If OldRow = NewRow Then Exit Sub
    
    Call SelectRow(vsf, OldRow, NewRow)
    
    If mblnShow Then Call RefreshData("����")
    
    On Error GoTo errHand
    Call mfrmMain.ActiveFormEnabled
    
errHand:
    
End Sub

Private Sub vsf_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    Call AppendSapceRows(vsf, lnX, lnY)
End Sub

Private Sub vsf_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Call AppendSapceRows(vsf, lnX, lnY)
End Sub

Private Sub vsf_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = (Col < 2)
End Sub

Private Sub vsf_DblClick()
    '
    Dim strNO As String
    Dim lng����id As Long
    Dim lng����id As Long
    Dim lng��¼���� As Long
    
    If Val(vsf.RowData(vsf.Row)) <= 0 Then Exit Sub
        
    strNO = vsf.TextMatrix(vsf.Row, GetCol(vsf, "No"))
    lng����id = Val(vsf.TextMatrix(vsf.Row, GetCol(vsf, "����id")))
    lng����id = Val(vsf.TextMatrix(vsf.Row, GetCol(vsf, "����id")))
    lng��¼���� = IIf(Val(vsf.TextMatrix(vsf.Row, GetCol(vsf, "����;��"))) = 1, 2, 1)
    
    If strNO = "" Or lng����id = 0 Then Exit Sub
                
    Call EditReport(mfrmMain, strNO, lng��¼����, lng����id, lng����id, "", True, True, , , , False, , , "001")
    
End Sub

Private Sub vsf_GotFocus()
    vsf.BackColorSel = COLOR.����
    Call SelectRow(vsf, 1, vsf.Row)
End Sub

Private Sub vsf_LostFocus()
    vsf.BackColorSel = COLOR.�ǽ���
    Call SelectRow(vsf, 1, vsf.Row)
End Sub

