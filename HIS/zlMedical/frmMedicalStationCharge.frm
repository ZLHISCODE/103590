VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMedicalStationCharge 
   BackColor       =   &H8000000A&
   BorderStyle     =   0  'None
   ClientHeight    =   6390
   ClientLeft      =   2715
   ClientTop       =   3420
   ClientWidth     =   8580
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   8580
   ShowInTaskbar   =   0   'False
   Begin VB.Frame picState 
      Height          =   765
      Left            =   300
      TabIndex        =   2
      Top             =   270
      Width           =   7845
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "�ϼ�:0.00(���м���:0.00 �շ�:0.00)"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   0
         Left            =   90
         TabIndex        =   4
         Top             =   225
         Width           =   3060
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "δ��:0.00 δ��:0.00"
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   1710
      End
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3135
      Left            =   600
      ScaleHeight     =   3135
      ScaleWidth      =   5460
      TabIndex        =   0
      Top             =   2865
      Width           =   5460
   End
   Begin VSFlex8Ctl.VSFlexGrid vsf 
      Height          =   1530
      Left            =   1845
      TabIndex        =   1
      Top             =   2085
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
      Left            =   6120
      Top             =   1215
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
            Picture         =   "frmMedicalStationCharge.frx":0000
            Key             =   "����"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationCharge.frx":039A
            Key             =   "����"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationCharge.frx":0734
            Key             =   "״̬"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationCharge.frx":0ACE
            Key             =   "����"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationCharge.frx":0E68
            Key             =   "����"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationCharge.frx":1202
            Key             =   "up"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationCharge.frx":13C4
            Key             =   "down"
         EndProperty
      EndProperty
   End
   Begin VB.Image imgX 
      Height          =   135
      Left            =   2505
      MousePointer    =   7  'Size N S
      Top             =   1680
      Width           =   5115
   End
End
Attribute VB_Name = "frmMedicalStationCharge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'���������弶��������**************************************************************************************************
Private mblnStartUp As Boolean
Private mfrmCharge As Object
Private mclsWork As New clsCISWork
Private mlngKey As Long
Private mlng�Ǽ�id As Long
Private mfrmMain As Object
Private mblnDataMoved As Boolean

Public Function zlMenuClick(ByVal frmMain As Object, ByVal lngKey As Long, ByVal strMenuItem As String, Optional ByVal lng�Ǽ�id As Long = 0) As Boolean
    '--------------------------------------------------------------------------------------------------------
    '���ܣ�
    '������lngKey ����ID
    '--------------------------------------------------------------------------------------------------------
    Dim lngSvrKey As Long

    On Error GoTo errHand
    
    mlngKey = lngKey
    mlng�Ǽ�id = lng�Ǽ�id
    
    Set mfrmMain = frmMain
    
    Select Case strMenuItem
    Case "ˢ��"
        
        lngSvrKey = Val(vsf.RowData(vsf.Row))
        Call zlClearData
        Call RefreshData(strMenuItem)
        Call RestoreRow(vsf, lngSvrKey)
        Call vsf_AfterRowColChange(0, 0, vsf.Row, vsf.Col)
        
        Call SumCharge
        
        zlMenuClick = True
        
    Case "����������"
                
        zlMenuClick = mfrmCharge.zlMenuClick("����������")
        If zlMenuClick Then Call zlMenuClick(mfrmMain, mlngKey, "ˢ��")
                
    Case "�շѵ���"
                
        zlMenuClick = mfrmCharge.zlMenuClick("�շѵ���")
        If zlMenuClick Then Call zlMenuClick(mfrmMain, mlngKey, "ˢ��")
        
    Case "���ʵ���"
            
        zlMenuClick = mfrmCharge.zlMenuClick("���ʵ���")
        If zlMenuClick Then Call zlMenuClick(mfrmMain, mlngKey, "ˢ��")
        
    Case "��Ѻ��õǼ�"
        
        zlMenuClick = mfrmCharge.zlMenuClick("��Ѻ��õǼ�")
                
    Case "�޸ĸ��ӷ���"
        
        zlMenuClick = mfrmCharge.zlMenuClick("�޸ĸ��ӷ���")
        If zlMenuClick Then Call zlMenuClick(mfrmMain, mlngKey, "ˢ��")
        
    Case "ɾ�����ӷ���"
                
        zlMenuClick = mfrmCharge.zlMenuClick("ɾ�����ӷ���")
        If zlMenuClick Then Call zlMenuClick(mfrmMain, mlngKey, "ˢ��")
        
    End Select
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Sub zlClearData(Optional ByVal strPart As String = "����")
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ�
    '������
    '------------------------------------------------------------------------------------------------------------------
    
    Call ResetVsf(vsf)
    Call AppendSapceRows(vsf, lnX, lnY)
    
End Sub

Public Property Get Body(ByVal lngIndex As Long) As Object
    Set Body = vsf
End Property

Private Sub SumCharge()
    '------------------------------------------------------------------------------------------------------------------
    '����:���û������
    '------------------------------------------------------------------------------------------------------------------
    
    Dim lngLoop As Long
    Dim sglSum As Single
    Dim strSQL As String
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Dim strTmp As String
        
    Call InitSysPara
    
    lbl(0).Caption = "ʵ�ս��:0.00(����:0.00 �շ�:0.00)��Ӧ�ս��:0.00(����:0.00 �շ�:0.00)��"
    lbl(1).Caption = "δ����:0.00(����:0.00 �շ�:0.00)"
'    lbl(1).Visible = False
    
    '��ȡ�ܵķ������

    strSQL = GetPublicSQL(SQL.���˷��øſ�)
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngKey)
    If CalcCharge(rsData, rs) Then
        strTmp = ""
        
        strTmp = strTmp & "ʵ�ս��:" & Format(zlCommFun.NVL(rs("ʵ�ս��").Value, 0), gstrDec) & "(����:" & Format(zlCommFun.NVL(rs("���ʽ��").Value, 0), gstrDec)
        strTmp = strTmp & " �շ�:" & Format(zlCommFun.NVL(rs("�շѽ��").Value, 0), gstrDec) & ")"
        
        strTmp = strTmp & "��Ӧ�ս��:" & Format(Val(zlCommFun.NVL(rs("Ӧ�ս��_��").Value, 0)) + Val(zlCommFun.NVL(rs("Ӧ�ս��_��").Value, 0)), gstrDec) & "(����:" & Format(zlCommFun.NVL(rs("Ӧ�ս��_��").Value, 0), gstrDec)
        strTmp = strTmp & " �շ�:" & Format(zlCommFun.NVL(rs("Ӧ�ս��_��").Value, 0), gstrDec) & ")"
        
        lbl(0).Caption = strTmp
        
        If zlCommFun.NVL(rs("δ����ϼ�").Value, 0) > 0 Then
            strTmp = ""
            strTmp = strTmp & "δ����:" & Format(zlCommFun.NVL(rs("δ����ϼ�").Value, 0), gstrDec) & "(����:" & Format(zlCommFun.NVL(rs("δ����").Value, 0), gstrDec)
            strTmp = strTmp & " �շ�:" & Format(zlCommFun.NVL(rs("δ�ս��").Value, 0), gstrDec) & ")"
            
            lbl(1).Caption = strTmp
'            lbl(1).Visible = True
        End If
    End If
    
End Sub

Private Function RefreshData(ByVal strMenu As String) As Boolean
    Dim strSQL As String
    Dim strTmp As String
    Dim rs As New ADODB.Recordset
    
    Select Case strMenu
    Case "ˢ��"

        gstrSQL = _
            "Select X.ҽ��id As ID, X.����id, X.���ͺ�, Y.ִ�п���id, Z.������Ŀid, Z.����;��, X.Ӧ�ս��, X.ʵ�ս��," & vbNewLine & _
            "       Decode(X.ִ��״̬, 1, '��ȫִ��', 2, 'ȡ��ִ��', 3, '����ִ��', '') As ִ��״̬," & vbNewLine & _
            "       Decode(Z.����id, Null, '', '����') As ����, U.���� As ִ�п���, P.���� As ��Ŀ," & vbNewLine & _
            "       Decode(X.����id, Null, Decode(T.�����ļ�id, Null, '', '����'), Decode(K.��д��, Null, '����', '����')) As ״̬" & vbNewLine & _
            "From (Select Decode(B.���id, Null, B.ID, B.���id) As ҽ��id, D.����id, D.���ͺ�, D.ִ��״̬," & vbNewLine & _
            "              Sum(A.Ӧ�ս��) As Ӧ�ս��, Sum(A.ʵ�ս��) As ʵ�ս��" & vbNewLine & _
            "       From ���˷��ü�¼ A, ����ҽ����¼ B, ���ǼǼ�¼ C, ����ҽ������ D, �����Ա���� E" & vbNewLine & _
            "       Where A.��¼״̬ In (0, 1) And A.ҽ�����(+) = B.ID And C.���� = B.�Һŵ� And B.������Դ = 4 And" & vbNewLine & _
            "             B.����id = E.����id And E.ID = [1] And D.ҽ��id = B.ID AND C.ID=E.�Ǽ�id " & vbNewLine & _
            "       Group By Decode(B.���id, Null, B.ID, B.���id), D.����id, D.���ͺ�, D.ִ��״̬) X, �����Ŀҽ�� Y," & vbNewLine & _
            "     �����Ŀ�嵥 Z, ���Ƶ���Ӧ�� T, ���˲�����¼ K, ������ĿĿ¼ P, ���ű� U" & vbNewLine & _
            "Where Y.ҽ��id = X.ҽ��id And Y.�嵥id = Z.ID And T.Ӧ�ó���(+) = 4 And T.������Ŀid(+) = Z.������Ŀid And" & vbNewLine & _
            "      K.ID(+) = X.����id And P.ID = Z.������Ŀid And U.ID = Y.ִ�п���id"
        
        '����ת������
        '--------------------------------------------------------------------------------------------------------------
        mblnDataMoved = DataMove(mlngKey, 2)
        If mblnDataMoved Then
            gstrSQL = Replace(gstrSQL, "���˷��ü�¼", "H���˷��ü�¼")
            gstrSQL = Replace(gstrSQL, "����ҽ����¼", "H����ҽ����¼")
            gstrSQL = Replace(gstrSQL, "���ǼǼ�¼", "H���ǼǼ�¼")
            gstrSQL = Replace(gstrSQL, "����ҽ������", "H����ҽ������")
            gstrSQL = Replace(gstrSQL, "�����Ա����", "H�����Ա����")
            gstrSQL = Replace(gstrSQL, "�����Ŀ�嵥", "H�����Ŀ�嵥")
            gstrSQL = Replace(gstrSQL, "���˲�����¼", "H���˲�����¼")
        Else
            '��ʱ���ܷ����Ѳ��ݻ���ȫת��
            strSQL = "Select a.���ʱ�� From ���ǼǼ�¼ a,�����Ա���� b Where a.ID=b.�Ǽ�id And b.ID=[1]"
            Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngKey)
            If rs.BOF = False Then
                If zlDatabase.DateMoved(Format(rs("���ʱ��").Value, "yyyy-MM-dd HH:mm:ss"), , glngSys, Me.Caption) Then
                    strTmp = strSQL
                    strTmp = Replace(strTmp, "���˷��ü�¼", "H���˷��ü�¼")
                    strSQL = strSQL & " Union All " & strTmp
                End If
            End If
        End If
        
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngKey)
        If rs.BOF = False Then
            
            Call InitSysPara
            
            Call LoadGrid(vsf, rs, Array("", "", "", "", "", gstrDec, gstrDec), , ils13)
            Call AppendSapceRows(vsf, lnX, lnY)
            
        End If
    
    Case "����"
                    
        If vsf.TextMatrix(vsf.Row, GetCol(vsf, "�������")) = "E" Then
            Set mfrmCharge = mclsWork.ListChargeInObject(Me, picContainer, Val(vsf.RowData(vsf.Row)), Val(vsf.TextMatrix(vsf.Row, GetCol(vsf, "���ͺ�"))), gcnOracle, gstrDBUser, glngSys, "", "����", "���")
        Else
            Set mfrmCharge = mclsWork.ListChargeInObject(Me, picContainer, Val(vsf.RowData(vsf.Row)), Val(vsf.TextMatrix(vsf.Row, GetCol(vsf, "���ͺ�"))), gcnOracle, gstrDBUser, glngSys, "", "���", "���")
        End If
        
    End Select
    
End Function

Private Function InitLoad() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:��ʼ�����ݣ������ڴ����Load�¼�
    '------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHand
    Dim strVsf As String
            
    strVsf = ",255,4,1,1,[����];,255,4,1,1,[״̬];��Ŀ,2400,1,1,1,;ִ�п���,1080,1,1,1,;ִ��״̬,900,1,1,1,;Ӧ�ս��,1080,7,1,1,;ʵ�ս��,1080,7,1,1,;���ͺ�,0,1,1,1,;������Դ,0,1,1,1,;�������,0,1,1,1,"

    Call CreateVsf(vsf, strVsf)
    vsf.Cols = vsf.Cols + 1
    vsf.ColWidth(vsf.Cols - 1) = 15
        
    Set vsf.Cell(flexcpPicture, 0, 0) = ils13.ListImages("����").Picture
    Set vsf.Cell(flexcpPicture, 0, 1) = ils13.ListImages("״̬").Picture
    
    Call AppendSapceRows(vsf, lnX, lnY)
    
    lbl(0).Caption = ""
    lbl(1).Caption = ""
    
    InitLoad = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
End Function



'���������弰��ؼ����¼�����******************************************************************************************

Private Sub Form_Load()
    
    mblnStartUp = True
    
    Call InitLoad
        
    Set mfrmCharge = mclsWork.ListChargeInObject(Me, picContainer, 0, 0, gcnOracle, gstrDBUser, glngSys, "", "���")
        
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    If Me.WindowState = 1 Then Exit Sub
    If imgX.Top > Me.ScaleHeight - 3000 Then imgX.Top = Me.ScaleHeight - 3000
    
    With picState
        .Left = 0
        .Top = -90
        .Width = Me.ScaleWidth
    End With
    
    With vsf
        .Left = 0
        .Top = picState.Top + picState.Height + 15
        .Width = Me.ScaleWidth
        .Height = imgX.Top - .Top
    End With
    
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
        
    Call AppendSapceRows(vsf, lnX, lnY)
End Sub

Private Sub imgX_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    
    On Error Resume Next
    
    imgX.Top = imgX.Top + Y
    
    If imgX.Top < 1500 Then imgX.Top = 1500
    If Me.Height - imgX.Top - imgX.Height < 3000 Then imgX.Top = Me.Height - imgX.Height - 3000
                
    Form_Resize
End Sub

Private Sub picContainer_Resize()
    On Error Resume Next
    
    If Not (mfrmCharge Is Nothing) Then
        mfrmCharge.Width = picContainer.Width
        mfrmCharge.Height = picContainer.Height
    End If
    
End Sub

Private Sub vsf_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If OldRow = NewRow Then Exit Sub
    
    Call SelectRow(vsf, OldRow, NewRow)
    
    Call RefreshData("����")
    
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

Private Sub vsf_GotFocus()
    vsf.BackColorSel = COLOR.����
    Call SelectRow(vsf, 1, vsf.Row)
End Sub

Private Sub vsf_LostFocus()
    vsf.BackColorSel = COLOR.�ǽ���
    Call SelectRow(vsf, 1, vsf.Row)
End Sub

Private Sub vsf_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Call SendLMouseButton(vsf.hWnd, X, Y)
        
        mfrmMain.mbytPopMenu = 4
        Set mfrmMain.mobjPopMenu = New clsPopMenu
        Call mfrmMain.mobjPopMenu.ShowPopupMenuByCursor
        
    End If
End Sub

Private Sub vsf_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 1 Then Cancel = True
End Sub
