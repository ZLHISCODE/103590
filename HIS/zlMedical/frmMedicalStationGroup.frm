VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMedicalStationGroup 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7485
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9480
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7485
   ScaleWidth      =   9480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Frame picState 
      Height          =   750
      Left            =   405
      TabIndex        =   3
      Top             =   195
      Width           =   5850
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "δ��:0.00 δ��:0.00"
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   1
         Left            =   90
         TabIndex        =   5
         Top             =   465
         Width           =   1710
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "�ϼ�:0.00(���м���:0.00 �շ�:0.00)"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   0
         Left            =   60
         TabIndex        =   4
         Top             =   210
         Width           =   3060
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsf 
      Height          =   1530
      Left            =   600
      TabIndex        =   0
      Top             =   1500
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
      HighLight       =   1
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
      Begin VB.Line lnX 
         Index           =   0
         Visible         =   0   'False
         X1              =   -555
         X2              =   1230
         Y1              =   555
         Y2              =   555
      End
      Begin VB.Line lnY 
         Index           =   0
         Visible         =   0   'False
         X1              =   270
         X2              =   270
         Y1              =   420
         Y2              =   1635
      End
   End
   Begin MSComctlLib.ImageList ils13 
      Left            =   6045
      Top             =   1170
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
            Picture         =   "frmMedicalStationGroup.frx":0000
            Key             =   "����"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationGroup.frx":039A
            Key             =   "����"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationGroup.frx":0734
            Key             =   "״̬"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationGroup.frx":0ACE
            Key             =   "����"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationGroup.frx":0E68
            Key             =   "����"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationGroup.frx":1202
            Key             =   "up"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationGroup.frx":13C4
            Key             =   "down"
         EndProperty
      EndProperty
   End
   Begin VB.Frame picCount 
      Height          =   555
      Left            =   3300
      TabIndex        =   1
      Top             =   3180
      Width           =   3255
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "����ͳ��"
         ForeColor       =   &H80000007&
         Height          =   180
         Index           =   2
         Left            =   105
         TabIndex        =   2
         Top             =   240
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmMedicalStationGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'���������弶��������**************************************************************************************************
Private mblnStartUp As Boolean
Private mvarParam As Variant
Private mfrmMain As Object
Private mblnDataMoved As Boolean

Public Function zlMenuClick(ByVal frmMain As Object, ByVal strMenuItem As String, Optional ByVal strParam As String = "") As Boolean
    '--------------------------------------------------------------------------------------------------------
    '���ܣ�
    '������
    '--------------------------------------------------------------------------------------------------------
    Dim lngSvrKey As Long

    On Error GoTo errHand
    
    Set mfrmMain = frmMain
    mvarParam = Split(strParam, "'")
    
    Select Case strMenuItem
    Case "ˢ��"
        
        Call zlClearData
        Call RefreshData(strMenuItem)
        
        Call SumCharge
    
    End Select
    
    zlMenuClick = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Sub zlClearData(Optional ByVal strPart As String = "����")
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ�
    '������
    '------------------------------------------------------------------------------------------------------------------
    
    Call ResetVsf(vsf)
    Call AppendSapceRows(vsf, lnX, lnY)
        
End Sub

Public Property Get Body(Optional ByVal lngIndex As Long) As Object
    Set Body = vsf
End Property

Private Sub SumCharge()
    '------------------------------------------------------------------------------------------------------------------
    '����:���û������
    '------------------------------------------------------------------------------------------------------------------
    Dim rsData As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim strTmp As String
    Dim strSQL As String
    
    Call InitSysPara
    
    lbl(0).Caption = "ʵ�ս��:0.00(����:0.00 �շ�:0.00)��Ӧ�ս��:0.00(����:0.00 �շ�:0.00)��"
    lbl(1).Caption = "δ����:0.00(����:0.00 �շ�:0.00)"
    
    '��ȡ�ܵķ������
    
    gstrSQL = GetPublicSQL(SQL.������øſ�)
    
    '����ת������
    '--------------------------------------------------------------------------------------------------------------
    If DataMove(Val(mvarParam(0))) Then
        gstrSQL = Replace(gstrSQL, "����ҽ����¼", "H����ҽ����¼")
        gstrSQL = Replace(gstrSQL, "���ǼǼ�¼", "H���ǼǼ�¼")
        gstrSQL = Replace(gstrSQL, "���˷��ü�¼", "H���˷��ü�¼")
    Else
        '��ʱ���ܷ����Ѳ��ݻ���ȫת��
        strSQL = "Select a.���ʱ�� From ���ǼǼ�¼ a,�����Ա���� b Where a.ID=b.�Ǽ�id And b.ID=[1]"
        Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(mvarParam(0)))
        If rs.BOF = False Then
            If zlDatabase.DateMoved(Format(rs("���ʱ��").Value, "yyyy-MM-dd HH:mm:ss"), , glngSys, Me.Caption) Then
                strTmp = gstrSQL
                strTmp = Replace(strTmp, "���˷��ü�¼", "H���˷��ü�¼")
                strSQL = gstrSQL & " Union All " & strTmp
            End If
        End If
    End If
                    
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mvarParam(0)))
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

        End If
        
    End If

End Sub

Private Function RefreshData(ByVal strMenu As String) As Boolean
    Dim lngLoop As Long
    Dim lngCount0 As Long
    Dim lngCount1 As Long
    Dim lngCount2 As Long
    Dim rs As New ADODB.Recordset
    
    On Error GoTo errHand
    
    lbl(2).Caption = ""
    
    Select Case strMenu
    Case "ˢ��"
        
        gstrSQL = "SELECT A.������� AS ���,A.����id AS ID,A.����,B.�����,a.�����," & _
                          "A.��챨�� AS ����," & _
                          "DECODE(A.��첡��ID, Null, 0, 1) As �ܼ�, " & _
                          "DECODE(A.���״̬, 5, 1, 0) As ���,c.Ӧ�ս��,c.ʵ�ս�� " & _
                     "FROM �����Ա���� A," & _
                          "������Ϣ B, " & _
                          "(Select ����id,Sum(D.Ӧ�ս��) As Ӧ�ս��,Sum(D.ʵ�ս��) As ʵ�ս�� " & _
                             "FROM ���˷��ü�¼ D, " & _
                                  "(SELECT C.ID " & _
                                     "FROM ���ǼǼ�¼ B, ����ҽ����¼ C " & _
                                    "WHERE C.������Դ = 4 AND " & _
                                          "C.ҽ��״̬ <> 4 AND B.���� = C.�Һŵ� AND B.ID = [1]) E " & _
                            "WHERE D.��¼״̬ IN (0, 1) AND D.ҽ����� = E.ID Group By d.����id) C " & _
                    "WHERE A.����ID = B.����ID AND A.�Ǽ�ID = [1] And a.����id=c.����id(+)"
                            
                    
        If Trim(mvarParam(1)) <> "" Then
            gstrSQL = gstrSQL & " And a.�������=[2] "
        End If
        
        gstrSQL = gstrSQL & " ORDER BY A.�������,B.�����,a.����� "
        
        '����ת������
        '--------------------------------------------------------------------------------------------------------------
        If DataMove(Val(mvarParam(0))) Then
            gstrSQL = Replace(gstrSQL, "�����Ա����", "H�����Ա����")
        End If
        
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mvarParam(0)), CStr(mvarParam(1)))
        If rs.BOF = False Then

            Call LoadGrid(vsf, rs)
            Call vsf_AfterRowColChange(0, 0, vsf.Row, vsf.Col)
            Call AppendSapceRows(vsf, lnX, lnY)
            
            'ͳ������������δ��������δ��������
            For lngLoop = 1 To vsf.Rows - 1
                
                'δ����ͳ��
                If Abs(Val(vsf.TextMatrix(lngLoop, 4))) <> 1 Then
                    lngCount0 = lngCount0 + 1
                Else
                    '����ͳ��
                    If Abs(Val(vsf.TextMatrix(lngLoop, 6))) = 1 Then
                        lngCount1 = lngCount1 + 1
                    Else
                        'δ��ͳ��
                        lngCount2 = lngCount2 + 1
                    End If
                End If
            Next
            
            lbl(2).Caption = "Ӧ��:" & lngCount0 + lngCount1 + lngCount2 & "��;ʵ��:" & lngCount1 + lngCount2 & "��(���:" & lngCount1 & "��;δ��:" & lngCount2 & "��);δ��:" & lngCount0 & "��"
                        
        End If
    End Select
    
    Exit Function
    
errHand:
    
    If ErrCenter = 1 Then Resume
    
End Function

Private Function InitLoad() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:��ʼ�����ݣ������ڴ����Load�¼�
    '------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHand
    Dim strVsf As String
            
    strVsf = "���,1500,1,1,1,;����,900,1,1,1,;�����,900,7,1,1,;�����,990,1,1,1,;����,600,4,1,1,;�ܼ�,600,4,1,1,;���,600,4,1,1,;Ӧ�ս��,1080,7,1,1,;ʵ�ս��,1080,7,1,1,"
    Call CreateVsf(vsf, strVsf)
    vsf.Cols = vsf.Cols + 1
    vsf.MergeCells = flexMergeFree
    vsf.MergeCol(0) = True
    vsf.ColWidth(vsf.Cols - 1) = 15
    vsf.ColDataType(4) = flexDTBoolean
    vsf.ColDataType(5) = flexDTBoolean
    vsf.ColDataType(6) = flexDTBoolean
    Call AppendSapceRows(vsf, lnX, lnY)
    vsf.ColFormat(7) = "0.00"
    vsf.ColFormat(8) = "0.00"
    
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
        
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    With picState
        .Left = 0
        .Top = -90
        .Width = Me.ScaleWidth
    End With
    
    With vsf
        .Left = 0
        .Top = picState.Top + picState.Height
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - .Top - picCount.Height + 90
    End With
                
    With picCount
        .Left = vsf.Left
        .Top = vsf.Top + vsf.Height - 90
        .Width = picState.Width
    End With
    
    Call AppendSapceRows(vsf, lnX, lnY)
    
End Sub

Private Sub vsf_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If OldRow = NewRow Then Exit Sub

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
    Cancel = (Col > 0)
End Sub

Private Sub vsf_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 1 Then Cancel = True
End Sub



