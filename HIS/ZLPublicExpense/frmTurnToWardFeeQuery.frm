VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Begin VB.Form frmTurnToWardFeeQuery 
   Caption         =   "ת�������ò�ѯ"
   ClientHeight    =   7515
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11580
   BeginProperty Font 
      Name            =   "����"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTurnToWardFeeQuery.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7515
   ScaleWidth      =   11580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.PictureBox picBack 
      BorderStyle     =   0  'None
      Height          =   915
      Index           =   4
      Left            =   -450
      ScaleHeight     =   915
      ScaleWidth      =   12015
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   6510
      Width           =   12015
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00C0C0C0&
         Caption         =   "ȡ��(&C)"
         Height          =   420
         Left            =   10290
         TabIndex        =   14
         ToolTipText     =   "�ȼ�:Esc"
         Top             =   300
         Width           =   1380
      End
      Begin VB.CommandButton cmdOK 
         BackColor       =   &H00C0C0C0&
         Caption         =   "ȷ��(&O)"
         Height          =   420
         Left            =   8790
         TabIndex        =   13
         ToolTipText     =   "�ȼ���F2"
         Top             =   300
         Width           =   1380
      End
   End
   Begin VB.PictureBox picBack 
      BorderStyle     =   0  'None
      Height          =   2055
      Index           =   3
      Left            =   870
      ScaleHeight     =   2055
      ScaleWidth      =   3615
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   5040
      Width           =   3615
      Begin VSFlex8Ctl.VSFlexGrid vsfGrid 
         Height          =   1485
         Index           =   3
         Left            =   150
         TabIndex        =   11
         Top             =   450
         Width           =   2565
         _cx             =   4524
         _cy             =   2619
         Appearance      =   2
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   12
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
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   5
         Cols            =   10
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
      Begin VB.Label lblShow 
         AutoSize        =   -1  'True
         Caption         =   "��������ȡ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   3
         Left            =   150
         TabIndex        =   17
         Tag             =   "������������ȡ��"
         Top             =   120
         Width           =   1260
      End
   End
   Begin VB.PictureBox picBack 
      BorderStyle     =   0  'None
      Height          =   2055
      Index           =   2
      Left            =   960
      ScaleHeight     =   2055
      ScaleWidth      =   3615
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2910
      Width           =   3615
      Begin VSFlex8Ctl.VSFlexGrid vsfGrid 
         Height          =   1485
         Index           =   2
         Left            =   270
         TabIndex        =   9
         Top             =   540
         Width           =   2565
         _cx             =   4524
         _cy             =   2619
         Appearance      =   2
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   12
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
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   5
         Cols            =   10
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
      Begin VB.Label lblShow 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   240
         TabIndex        =   16
         Tag             =   "��������"
         Top             =   210
         Width           =   1260
      End
   End
   Begin VB.PictureBox picBack 
      BorderStyle     =   0  'None
      Height          =   2055
      Index           =   1
      Left            =   900
      ScaleHeight     =   2055
      ScaleWidth      =   6645
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   780
      Width           =   6645
      Begin VB.PictureBox picCboBack 
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   3030
         ScaleHeight     =   315
         ScaleWidth      =   2655
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   390
         Visible         =   0   'False
         Width           =   2655
         Begin VB.ComboBox cboDate 
            Height          =   360
            Left            =   -30
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   -30
            Width           =   2715
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfGrid 
         Height          =   1485
         Index           =   1
         Left            =   360
         TabIndex        =   7
         Top             =   480
         Width           =   2565
         _cx             =   4524
         _cy             =   2619
         Appearance      =   2
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   12
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
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   6
         Cols            =   10
         FixedRows       =   2
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
      Begin VB.Label lblShow 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�䶯ʱ��: 2017-06-12"
         Height          =   240
         Index           =   1
         Left            =   450
         TabIndex        =   15
         Tag             =   "�䶯ʱ��: "
         Top             =   120
         Width           =   2400
      End
   End
   Begin VB.PictureBox picBack 
      BorderStyle     =   0  'None
      Height          =   675
      Index           =   0
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   12135
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   12135
      Begin VB.Label lblShow 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "סԺ����: ��1��"
         Height          =   240
         Index           =   8
         Left            =   8280
         TabIndex        =   18
         Tag             =   "סԺ����: "
         Top             =   180
         Width           =   1800
      End
      Begin VB.Shape shp������Ϣ 
         BorderColor     =   &H8000000A&
         Height          =   105
         Left            =   60
         Top             =   0
         Width           =   11715
      End
      Begin VB.Label lblShow 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ѱ�: ��ͨ"
         Height          =   240
         Index           =   9
         Left            =   10380
         TabIndex        =   5
         Tag             =   "�ѱ�: "
         Top             =   180
         Width           =   1200
      End
      Begin VB.Label lblShow 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "סԺ��: 99999999"
         Height          =   240
         Index           =   7
         Left            =   6090
         TabIndex        =   4
         Tag             =   "סԺ��: "
         Top             =   180
         Width           =   1920
      End
      Begin VB.Label lblShow 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����: 26��3��"
         Height          =   240
         Index           =   6
         Left            =   4020
         TabIndex        =   3
         Tag             =   "����: "
         Top             =   180
         Width           =   1560
      End
      Begin VB.Label lblShow 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Ա�: ��"
         Height          =   240
         Index           =   5
         Left            =   2940
         TabIndex        =   2
         Tag             =   "�Ա�: "
         Top             =   180
         Width           =   960
      End
      Begin VB.Line lineShow 
         BorderColor     =   &H8000000A&
         Index           =   0
         X1              =   900
         X2              =   2400
         Y1              =   435
         Y2              =   435
      End
      Begin VB.Label lblShow 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����: ������"
         Height          =   240
         Index           =   4
         Left            =   180
         TabIndex        =   1
         Tag             =   "����: "
         Top             =   180
         Width           =   1440
      End
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Left            =   0
      Top             =   660
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmTurnToWardFeeQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'��ڲ���
Private mbyt���� As Fun_Index
Private mlng����ID As Long
Private mlng��ҳID As Long
Private mlng�䶯id As Long
Private mlngԭ����id As Long
Private mlngĿ�겡��id As Long
Private mcllSQL As Collection
'----------------------------------------------------------------------------
Private mblnOK As Boolean

Private Enum IndexDef
    Lbl_�䶯ʱ�� = 1
    Lbl_�������� = 2
    Lbl_����ȡ�� = 3
    Lbl_���� = 4
    Lbl_�Ա� = 5
    Lbl_���� = 6
    Lbl_סԺ�� = 7
    Lbl_סԺ���� = 8
    Lbl_�ѱ� = 9
    
    Pane_������Ϣ = 0
    Pane_�䶯���� = 1
    Pane_�������� = 2
    Pane_����ȡ�� = 3
    Pane_���ܰ�ť = 4
End Enum

Private Enum Fun_Index
    Fun_ת���� = 0
    Fun_����ת���� = 1
    Fun_ת�������� = 2
    Fun_��ʷת����ѯ = 3
End Enum

Public Function TurnToWard_Fee_Query(ByVal frmMain As Object, ByVal byt���� As Byte, _
    ByVal lng����ID As Long, ByVal lng��ҳId As Long, _
    Optional ByVal lng�䶯id As Long, _
    Optional ByVal lngԭ����id As Long, Optional lngĿ�겡��id As Long, _
    Optional ByRef cllSql As Collection) As Boolean
     '------------------------------------------------------------------------------------
    '����:ת�������ò�ѯ
    '���:frmMain-���õ�������
    '     byt����- 0-ת����;1-��������;2-ת��������;3-��ʷת����ѯ
    '     cllSQL - ��Ҫִ�е�SQL��䣬0-ת����/1-����������Ч��
    '               0-ת����ʱ���ڸ��ı䶯��¼֮��ִ�У�1-��������ʱ���ڸ��ı䶯��¼֮ǰִ��
    '   1��int����=0(ת����)ʱ
    '       lng�䶯id: ԭ�����ı䶯��¼��ID
    '       lngԭ����id��ԭ����ID
    '       lngĿ�겡��id:Ŀ�겡��ID
    '   2��int����=1(��������)ʱ
    '       lng�䶯id: �ָ���ԭ�����ı䶯��¼��ID
    '       lngԭ����id���������Ĳ���ID
    '       lngĿ�겡��id:�ָ���ԭʼ����ID
    '  3. int����=3(��ʷת����ѯ)
    '����:
    '����:���ò�ѯʱ����Ա��ȷ�Ϸ���true,���򷵻�False
    '-----------------------------------------------------------------------------------
    mbyt���� = byt����
    mlng����ID = lng����ID: mlng��ҳID = lng��ҳId
    mlng�䶯id = lng�䶯id
    mlngԭ����id = lngԭ����id: mlngĿ�겡��id = lngĿ�겡��id
    
    mblnOK = False
    On Error Resume Next
    Me.Show 1, frmMain
    
    If mblnOK Then
        If Not mcllSQL Is Nothing Then
            Set cllSql = mcllSQL
        End If
        TurnToWard_Fee_Query = True
    End If
End Function

Private Sub Form_Load()
    Dim objPane As Pane
    
    Me.Width = 1024 * Screen.TwipsPerPixelX
    Me.Height = 768 * Screen.TwipsPerPixelY
    picCboBack.Visible = (mbyt���� = Fun_��ʷת����ѯ)
    
    If mbyt���� = Fun_ת�������� Or mbyt���� = Fun_ת���� Then
        If Upgradeҽ��ִ�мƼ�ִ��״̬(mlng����ID, mlng��ҳID) = False Then
            MsgBox "ҽ��ִ�мƼ���������ʧ�ܣ����ܼ�����", vbInformation, gstrSysName
            Unload Me: Exit Sub
        End If
    End If
    
    If mbyt���� = Fun_��ʷת����ѯ Then
        cmdOK.Visible = False: cmdOK.Enabled = False
        cmdCancel.Caption = "�˳�(&E)"
        Me.Caption = "ת�������ñ䶯��ѯ"
    Else
        Me.Caption = "ת�������ñ䶯"
    End If
    
    If ShowPatientInfo(mlng����ID, mlng��ҳID) = False Then Unload Me: Exit Sub
    
    If InitPanel() = False Then Unload Me: Exit Sub
    If InitGrid() = False Then Unload Me: Exit Sub
    If mbyt���� <> Fun_����ת���� Then
        Set objPane = dkpMain.FindPane(Pane_����ȡ��)
        If Not objPane Is Nothing Then
            If Not objPane.Closed Then objPane.Close
        End If
    ElseIf mbyt���� = Fun_����ת���� Then
        Set objPane = dkpMain.FindPane(Pane_��������)
        If Not objPane Is Nothing Then
            If Not objPane.Closed Then objPane.Close
        End If
    End If
    
    If mbyt���� = Fun_��ʷת����ѯ Then
        If LoadHistory(mlng����ID, mlng��ҳID) = False Then Unload Me: Exit Sub
        lblShow(Lbl_�䶯ʱ��).Caption = lblShow(Lbl_�䶯ʱ��).Tag
    Else
        lblShow(Lbl_�䶯ʱ��).Caption = lblShow(Lbl_�䶯ʱ��).Tag & Format(gobjDatabase.Currentdate, "yyyy-mm-dd hh:mm:ss")
        If mbyt���� = Fun_����ת���� Then
            If LoadFeeData(mlng����ID, mlng��ҳID, mlngԭ����id, mlngĿ�겡��id, True, mlng�䶯id) = False Then Unload Me: Exit Sub
        Else
            If LoadFeeData(mlng����ID, mlng��ҳID, mlngԭ����id, mlngĿ�겡��id) = False Then Unload Me: Exit Sub
        End If
    End If
End Sub

Private Function LoadHistory(ByVal lng����ID As Long, ByVal lng��ҳId As Long) As Boolean
    '��ʾ��ʷת������¼
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    Err = 0: On Error GoTo ErrHandler
    strSql = "Select Distinct �䶯ʱ��, ��¼״̬ From ���ñ䶯��¼ Where ����id = [1] And ��ҳid = [2] Order By �䶯ʱ��"
    Set rsData = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, lng����ID, lng��ҳId)
    If rsData.EOF Then
        MsgBox "��ǰ����û��ת�������ñ䶯��¼��", vbExclamation, gstrSysName
        Exit Function
    End If
    
    cboDate.Clear
    cboDate.Tag = ""
    Do While Not rsData.EOF
        cboDate.AddItem Format(rsData!�䶯ʱ��, "yyyy-mm-dd hh:mm:ss")
        cboDate.ItemData(cboDate.NewIndex) = Val(Nvl(rsData!��¼״̬)) '1-�����䶯��¼;2-�����ı䶯��¼
        rsData.MoveNext
    Loop
    cboDate.ListIndex = 0
    
    LoadHistory = True
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Private Function InitPanel() As Boolean
    '��ʼ�����沼��
    Dim objPane As Pane

    Err = 0: On Error GoTo ErrHandler
    dkpMain.DestroyAll
    Set objPane = dkpMain.CreatePane(Pane_������Ϣ, 100, 35, DockTopOf, Nothing)
    objPane.Handle = picBack(Pane_������Ϣ).hWnd
    objPane.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoHideable Or PaneNoFloatable
    objPane.MinTrackSize.Height = 38
    objPane.MaxTrackSize.Height = 38

    Set objPane = dkpMain.CreatePane(Pane_�䶯����, 100, 135, DockBottomOf, objPane)
    objPane.Handle = picBack(Pane_�䶯����).hWnd
    objPane.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoHideable Or PaneNoFloatable
    
    Set objPane = dkpMain.CreatePane(Pane_��������, 100, 100, DockBottomOf, objPane)
    objPane.Handle = picBack(Pane_��������).hWnd
    objPane.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoHideable Or PaneNoFloatable
    
    Set objPane = dkpMain.CreatePane(Pane_����ȡ��, 100, 80, DockBottomOf, objPane)
    objPane.Handle = picBack(Pane_����ȡ��).hWnd
    objPane.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoHideable Or PaneNoFloatable
    
    Set objPane = dkpMain.CreatePane(Pane_���ܰ�ť, 100, 135, DockBottomOf, objPane)
    objPane.Handle = picBack(Pane_���ܰ�ť).hWnd
    objPane.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoHideable Or PaneNoFloatable
    objPane.MinTrackSize.Height = 60
    objPane.MaxTrackSize.Height = 60

    With dkpMain
        .VisualTheme = ThemeOffice2003
        .Options.ThemedFloatingFrames = True
        .Options.UseSplitterTracker = True 'ʵʱ�϶�
        .Options.AlphaDockingContext = True
        .Options.HideClient = True
    End With
    
    InitPanel = True
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Private Function InitGrid() As Boolean
    '��ʼ�����
    Dim strHead(1 To 3) As String, strHeadSub As String
    Dim varData As Variant, varDataSub As Variant
    Dim i As Integer, k As Integer
    
    On Error GoTo ErrHandler
    strHead(1) = "��������,1,1050|NO,4,1200|�շ�ϸĿID,0,0|�շ���Ŀ,1,2800|ת��,1,1600|ת��,7,1000|ת��,1,1600|ת��,7,1000|����,7,1300|���,7,1300"
    strHeadSub = "��������,NO,�շ�ϸĿID,�շ���Ŀ,����,����,����,����,����,���"
    strHead(2) = "��������,1,1050|NO,4,1200|�շ�ϸĿID,0,0|�շ���Ŀ,1,2800|����,1,1600|����,7,1000|������ʽ,1,2400"
    strHead(3) = "����ʱ��,4,2400|������,1,1000|NO,4,1200|�շ�ϸĿID,0,0|�շ���Ŀ,1,2800|����,7,1000|���벿��,1,1600|��˲���,1,1600|�������,1,1600"
    
    For k = 1 To 3
        varData = Split(strHead(k), "|")
        If (k = 1) Then varDataSub = Split(strHeadSub, ",")
        
        With vsfGrid(k)
            .Cols = UBound(varData) + 1
            For i = 0 To UBound(varData)
                .TextMatrix(0, i) = Split(varData(i), ",")(0)
                If (k = 1) Then .TextMatrix(1, i) = varDataSub(i)
                .ColAlignment(i) = Split(varData(i), ",")(1)
                .ColWidth(i) = Split(varData(i), ",")(2)
                If k = 1 And (Split(varData(i), ",")(0) = "ת��" Or Split(varData(i), ",")(0) = "ת��") Then
                    .ColKey(i) = Split(varData(i), ",")(0) & "-" & varDataSub(i)
                Else
                    .ColKey(i) = Split(varData(i), ",")(0)
                End If
            Next
            
            .FixedAlignment(-1) = flexAlignCenterCenter '�������ı�����
            If k = 1 Then
                .MergeCells = flexMergeFixedOnly
                .MergeRow(0) = True: .MergeCol(-1) = True
            End If
            
            .ColHidden(.ColIndex("�շ�ϸĿID")) = True
        End With
        Call gobjComlib.RestoreFlexState(vsfGrid(k), App.ProductName & "\" & Me.Name & "_" & k)
    Next
    
    InitGrid = True
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Private Function ShowPatientInfo(ByVal lng����ID As Long, ByVal lng��ҳId As Long) As Boolean
    '��������
    '��Σ�
    '
    Dim strSql As String, rsTemp As ADODB.Recordset
    
    Err = 0: On Error GoTo ErrHandler
    strSql = "Select Nvl(a.����, b.����) As ����, Nvl(a.�Ա�, b.�Ա�) As �Ա�, Nvl(a.����, b.����) As ����," & vbNewLine & _
            "        Nvl(a.סԺ��, b.סԺ��) As סԺ��, a.��ҳID, Nvl(a.�ѱ�, b.�ѱ�) As �ѱ�," & vbNewLine & _
            "        Nvl(��˱�־, 0) As  ��˱�־, Nvl(a.״̬, 0) As ״̬" & vbNewLine & _
            " From ������ҳ A, ������Ϣ B" & vbNewLine & _
            " Where a.����id = b.����id And a.����id = [1] And a.��ҳid = [2]"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, lng����ID, lng��ҳId)
    If rsTemp.EOF Then
        MsgBox "δ���ֲ�����Ϣ��", vbExclamation, gstrSysName
        Exit Function
    End If
    
    lblShow(Lbl_����).Caption = lblShow(Lbl_����).Tag & Nvl(rsTemp!����)
    lblShow(Lbl_�Ա�).Caption = lblShow(Lbl_�Ա�).Tag & Nvl(rsTemp!�Ա�)
    lblShow(Lbl_����).Caption = lblShow(Lbl_����).Tag & Nvl(rsTemp!����)
    lblShow(Lbl_סԺ��).Caption = lblShow(Lbl_סԺ��).Tag & Nvl(rsTemp!סԺ��)
    lblShow(Lbl_סԺ��).Tag = Val(Nvl(rsTemp!״̬)) 'סԺ״̬
    lblShow(Lbl_סԺ����).Caption = lblShow(Lbl_סԺ����).Tag & "��" & Nvl(rsTemp!��ҳID) & "��"
    lblShow(Lbl_סԺ����).Tag = Val(Nvl(rsTemp!��˱�־)) '��˱�־
    lblShow(Lbl_�ѱ�).Caption = lblShow(Lbl_�ѱ�).Tag & Nvl(rsTemp!�ѱ�)
    
    ShowPatientInfo = True
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Private Function LoadFeeData(ByVal lng����ID As Long, ByVal lng��ҳId As Long, _
    ByVal lngԭ����id As Long, ByVal lngĿ�겡��id As Long, _
    Optional ByVal blnCancel As Boolean, Optional ByVal lngĿ��䶯id As Long) As Boolean
    '��ʾ����
    '��Σ�
    '   blnCancel - �Ƿ���תԺ
    Dim strSql As String, rsBill As ADODB.Recordset
    Dim lngRow As Long, blnNotData As Boolean
    Dim strDeptNameSql As String, strFeeSql As String
    Dim dblת������ As Double, strAdviceSql As String
    
    Err = 0: On Error GoTo ErrHandler
    blnNotData = True
    '��ս�������
    vsfGrid(Pane_�䶯����).Clear 1
    vsfGrid(Pane_�䶯����).Rows = vsfGrid(Pane_�䶯����).FixedRows
    vsfGrid(Pane_��������).Clear 1
    vsfGrid(Pane_��������).Rows = vsfGrid(Pane_��������).FixedRows
    vsfGrid(Pane_����ȡ��).Clear 1
    vsfGrid(Pane_����ȡ��).Rows = vsfGrid(Pane_����ȡ��).FixedRows
    
    '������ϸ
    strFeeSql = _
        "Select a.No, a.���, Max(a.�շ�ϸĿid) As �շ�ϸĿid, Max(a.ҽ�����) As ҽ��id, Sum(����) As ʣ������," & vbNewLine & _
        "        Max(Decode(a.��¼״̬, 2, 0, a.����id)) As ����id, Max(a.��¼״̬) As ��¼״̬, Max(a.ִ��״̬) As ִ��״̬, Max(a.��׼����) As ��׼����" & vbNewLine & _
        "From (Select a.No, ��¼״̬, a.ҽ�����, Nvl(a.�۸񸸺�, ���) As ���, �շ�ϸĿid, a.ִ��״̬, Avg(Nvl(a.����, 1) * a.����) As ����," & vbNewLine & _
        "               Sum(a.��׼����) As ��׼����, Max(Decode(a.�۸񸸺�, Null, a.Id, 0)) As ����id" & vbNewLine & _
        "        From סԺ���ü�¼ A, ����ҽ������ B, �������� C" & vbNewLine & _
        "        Where a.��¼���� = b.��¼���� And a.ҽ����� = b.ҽ��id And a.No = b.No And a.�շ�ϸĿid = c.����id(+) And a.��¼���� = 2 And" & vbNewLine & _
        "              a.ִ�в���id = [3] And a.ҽ����� Is Not Null And a.����id = [1] And a.��ҳid = [2] And" & vbNewLine & _
        "              Instr(',5,6,7,', ',' || a.�շ���� || ',') = 0 And Nvl(c.��������, 0) = 0" & vbNewLine & _
        "        Group By a.No, ��¼״̬, a.ҽ�����, Nvl(a.�۸񸸺�, ���), �շ�ϸĿid, a.ִ��״̬) A" & vbNewLine & _
        "Group By a.No, a.���" & vbNewLine
    'ҽ����Ϣ
    strAdviceSql = _
        "Select b.No, b.ҽ��id, b.���ͺ�, c.�շ�ϸĿid, Nvl(c.����id, 0) As ����id, Sum(Decode(Nvl(c.ִ��״̬, 0), 1, c.����, 0)) As ��ִ����" & vbNewLine & _
        "From ����ҽ����¼ A, ����ҽ������ B, ҽ��ִ�мƼ� C, �������� D" & vbNewLine & _
        "Where a.Id = b.ҽ��id And b.���ͺ� = c.���ͺ�(+) And b.ҽ��id = c.ҽ��id(+) And b.��¼���� = 2 And Nvl(b.ִ��״̬, 0) <> 1 And" & vbNewLine & _
        "      c.�շ�ϸĿid = d.����id(+) And Nvl(d.��������, 0) = 0 And a.����id = [1] And a.��ҳid = [2]" & vbNewLine & _
        "Group By b.No, b.ҽ��id, b.���ͺ�, c.�շ�ϸĿid, Nvl(c.����id, 0)" & vbNewLine
    
    '��������
    strDeptNameSql = _
        " Select Max(Decode(ID, [3], ����, Null)) As ԭ����, Max(Decode(ID, [4], ����, Null)) As Ŀ�겡��" & vbNewLine & _
        " From ���ű� Where ID In ([3], [4])"

    strSql = _
        " Select Decode(a.��¼״̬, 0, '���ʻ���', '����') As ��������, a.No, Sum(Nvl(ʣ������, 0) - Nvl(��ִ����, 0)) As ׼����," & vbNewLine & _
        "        a.�շ�ϸĿID, c.���� As �շ���Ŀ, a.��׼����, n.ԭ����, n.Ŀ�겡��, Nvl(Sum(ʣ������), 0) As ʣ������" & vbNewLine & _
        " From (" & strFeeSql & ") A, (" & strAdviceSql & ") B, �շ���ĿĿ¼ C, (" & strDeptNameSql & ") N" & vbNewLine & _
        " Where a.No = b.No And a.ҽ��id = b.ҽ��id And a.�շ�ϸĿid = b.�շ�ϸĿid And a.�շ�ϸĿId = c.Id" & vbNewLine & _
        "       And (a.����id = b.����id Or Nvl(b.����id, 0) = 0) And Nvl(ʣ������, 0) - Nvl(��ִ����, 0) > 0" & vbNewLine & _
        " Group By Decode(a.��¼״̬, 0, '���ʻ���', '����'), a.No, a.�շ�ϸĿid, c.����, a.��׼����, n.ԭ����, n.Ŀ�겡��" & vbNewLine & _
        " Having Sum(Nvl(ʣ������, 0) - Nvl(��ִ����, 0)) > 0" & vbNewLine & _
        " Order By NO"
    Set rsBill = gobjDatabase.OpenSQLRecord(strSql, "��ȡ������Ϣ", lng����ID, lng��ҳId, lngԭ����id, lngĿ�겡��id)
    If rsBill.RecordCount > 0 Then
        blnNotData = False
        With vsfGrid(Pane_�䶯����)
            .Redraw = flexRDNone
            .Rows = .FixedRows + rsBill.RecordCount
            lngRow = .FixedRows
            Do While Not rsBill.EOF
                If Nvl(rsBill!��������) = "���ʻ���" Then '���˻��۵����޸�ԭ��¼����˱䶯�����Ƿ��õ�ʣ������
                    dblת������ = Val(Nvl(rsBill!ʣ������))
                Else
                    dblת������ = Val(Nvl(rsBill!׼����))
                End If
                If dblת������ > 0 Then
                    .TextMatrix(lngRow, .ColIndex("��������")) = Nvl(rsBill!��������)
                    .TextMatrix(lngRow, .ColIndex("NO")) = Nvl(rsBill!NO)
                    .TextMatrix(lngRow, .ColIndex("�շ�ϸĿID")) = Nvl(rsBill!�շ�ϸĿID)
                    .TextMatrix(lngRow, .ColIndex("�շ���Ŀ")) = Nvl(rsBill!�շ���Ŀ)
                    .TextMatrix(lngRow, .ColIndex("ת��-����")) = Nvl(rsBill!ԭ����)
                    .TextMatrix(lngRow, .ColIndex("ת��-����")) = FormatEx(dblת������, 5)
                    .TextMatrix(lngRow, .ColIndex("ת��-����")) = Nvl(rsBill!Ŀ�겡��)
                    .TextMatrix(lngRow, .ColIndex("ת��-����")) = FormatEx(dblת������, 5)
                    .TextMatrix(lngRow, .ColIndex("����")) = Format(Val(Nvl(rsBill!��׼����)), gSysPara.Price_Decimal.strFormt_VB)
                    .TextMatrix(lngRow, .ColIndex("���")) = Format(Val(Nvl(rsBill!��׼����)) * dblת������, gSysPara.Money_Decimal.strFormt_VB)
                    
                    lngRow = lngRow + 1
                End If
                rsBill.MoveNext
            Loop
            .Redraw = flexRDBuffered
        End With
    End If
        
    '������������
    '�������õ�����ִ���˵�Ҳ������������
    strSql = _
        "With סԺ���� As (" & _
        " Select a.No, a.���, Max(a.�շ�ϸĿid) As �շ�ϸĿid, Sum(����) As ʣ������, Max(Decode(a.��¼״̬, 2, 0, a.����id)) As ����id," & vbNewLine & _
        "         Max(a.��¼״̬) As ��¼״̬" & vbNewLine & _
        " From (Select a.No, ��¼״̬, Nvl(a.�۸񸸺�, ���) As ���, �շ�ϸĿid, Avg(Nvl(a.����, 1) * a.����) As ����," & vbNewLine & _
        "                Max(Decode(a.�۸񸸺�, Null, a.Id, 0)) As ����id" & vbNewLine & _
        "         From סԺ���ü�¼ A, ����ҽ������ B, �������� C" & vbNewLine & _
        "         Where a.��¼���� = b.��¼���� And a.ҽ����� = b.ҽ��id And a.No = b.No And a.��¼���� = 2 And a.ִ�в���id = [3] And" & vbNewLine & _
        "               a.ҽ����� Is Not Null And a.����id = [1] And a.��ҳid = [2] And Instr(',5,6,7,', ',' || a.�շ���� || ',') = 0 And" & vbNewLine & _
        "               a.�շ�ϸĿid = c.����id And Nvl(c.��������, 0) = 1" & vbNewLine & _
        "         Group By a.No, ��¼״̬, a.ҽ�����, Nvl(a.�۸񸸺�, ���), �շ�ϸĿid, a.ִ��״̬) A" & vbNewLine & _
        " Group By a.No, a.���)"
    
    strSql = _
        " Select Decode(a.��¼״̬, 0, '���ʻ���', '����') As ��������, a.No, a.�շ�ϸĿid, ׼����, b.���� As �շ���Ŀ, n.ԭ����" & vbNewLine & _
        " From (" & vbNewLine & _
        "        " & strSql & vbNewLine & _
        "        Select Nvl(Sum(a.ʣ������), 0) - Nvl(Sum(b.����), 0) As ׼����, a.No, Max(a.��¼״̬) As ��¼״̬, a.�շ�ϸĿid" & vbNewLine & _
        "        From סԺ���� A," & vbNewLine & _
        "             (Select b.����id, Nvl(Sum(b.����), 0) As ����" & vbNewLine & _
        "               From סԺ���� A, ���˷������� B" & vbNewLine & _
        "               Where a.����id = B.����id And Nvl(B.״̬, 0) = 0" & vbNewLine & _
        "               Group By b.����id" & vbNewLine & _
        "               Having Nvl(Sum(b.����), 0) <> 0) B" & vbNewLine & _
        "        Where a.����id = b.����id(+)" & vbNewLine & _
        "        Group By a.No, a.�շ�ϸĿid" & vbNewLine & _
        "        Having Nvl(Sum(a.ʣ������), 0) - Nvl(Sum(b.����), 0) <> 0) A, �շ���ĿĿ¼ B,(" & strDeptNameSql & ") N" & vbNewLine & _
        " Where a.�շ�ϸĿID = B.ID" & vbNewLine & _
        " Order By NO, �շ�ϸĿid"
    Set rsBill = gobjDatabase.OpenSQLRecord(strSql, "��ȡ������Ϣ", lng����ID, lng��ҳId, lngԭ����id, lngĿ�겡��id)
    If rsBill.RecordCount > 0 Then
        blnNotData = False
        With vsfGrid(Pane_��������)
            .Redraw = flexRDNone
            .Rows = .FixedRows + rsBill.RecordCount
            lngRow = .FixedRows
            Do While Not rsBill.EOF
                .TextMatrix(lngRow, .ColIndex("��������")) = Nvl(rsBill!��������)
                .TextMatrix(lngRow, .ColIndex("NO")) = Nvl(rsBill!NO)
                .TextMatrix(lngRow, .ColIndex("�շ�ϸĿID")) = Nvl(rsBill!�շ�ϸĿID)
                .TextMatrix(lngRow, .ColIndex("�շ���Ŀ")) = Nvl(rsBill!�շ���Ŀ)
                .TextMatrix(lngRow, .ColIndex("����")) = Nvl(rsBill!ԭ����)
                .TextMatrix(lngRow, .ColIndex("����")) = FormatEx(Nvl(rsBill!׼����), 5)
                If Nvl(rsBill!��������) = "���ʻ���" Then
                    .TextMatrix(lngRow, .ColIndex("������ʽ")) = "��ֹת����"
                    '��ɫ������
                    .Cell(flexcpForeColor, lngRow, 0, lngRow, .Cols - 1) = vbRed
                Else
                    .TextMatrix(lngRow, .ColIndex("������ʽ")) = "��������"
                End If
                
                lngRow = lngRow + 1
                rsBill.MoveNext
            Loop
            .Redraw = flexRDBuffered
        End With
    End If
    
    '��������ȡ��
    If blnCancel Then
        strSql = "Select ID From ���˱䶯��¼ Where ����ID = [1] And ��ҳID = [2] And ��ʼԭ�� = 15 And ��ֹʱ�� Is Null"
        Set rsBill = gobjDatabase.OpenSQLRecord(strSql, "��ȡԭ�䶯ID", lng����ID, lng��ҳId)
        If rsBill.EOF Then
            MsgBox "δ�ҵ����˵�ԭʼ�䶯��¼����ֹ������", vbInformation, gstrSysName
            Exit Function
        End If
        
        strSql = _
            " Select a.����ʱ��, a.������, b.No, a.�շ�ϸĿID, c.���� As �շ���Ŀ, Sum(a.����) As ����, e.���� As ���벿��, f.���� As ��˲���," & vbNewLine & _
            "        Decode(a.�������, 1, '��ִ��', 'δִ��') As �������" & vbNewLine & _
            " From ���˷������� A, ���ñ䶯��¼ B, �շ���ĿĿ¼ C, ���ű� E, ���ű� F" & vbNewLine & _
            " Where a.����id = b.����id And a.�շ�ϸĿid = c.Id And a.���벿��id = e.Id And a.��˲���id = f.Id" & vbNewLine & _
            "       And b.ԭ�䶯id = [1] And b.Ŀ��䶯ID = [2] And b.״̬ = 2 And a.״̬ In (0, 2)" & vbNewLine & _
            " Group By a.����ʱ��, a.������, b.No, a.�շ�ϸĿid, c.����, e.����, f.����, Decode(a.�������, 1, '��ִ��', 'δִ��')" & vbNewLine & _
            " Order By No, �շ�ϸĿID"
        Set rsBill = gobjDatabase.OpenSQLRecord(strSql, "��ȡ��������ȡ������", lngĿ��䶯id, Val(Nvl(rsBill!ID)))
        If rsBill.RecordCount > 0 Then
            blnNotData = False
            With vsfGrid(Pane_����ȡ��)
                .Redraw = flexRDNone
                .Rows = .FixedRows + rsBill.RecordCount
                lngRow = .FixedRows
                Do While Not rsBill.EOF
                    .TextMatrix(lngRow, .ColIndex("����ʱ��")) = Format(Nvl(rsBill!����ʱ��), "yyyy-mm-dd hh:mm:ss")
                    .TextMatrix(lngRow, .ColIndex("������")) = Nvl(rsBill!������)
                    .TextMatrix(lngRow, .ColIndex("NO")) = Nvl(rsBill!NO)
                    .TextMatrix(lngRow, .ColIndex("�շ�ϸĿID")) = Nvl(rsBill!�շ�ϸĿID)
                    .TextMatrix(lngRow, .ColIndex("�շ���Ŀ")) = Nvl(rsBill!�շ���Ŀ)
                    .TextMatrix(lngRow, .ColIndex("����")) = FormatEx(Val(Nvl(rsBill!����)), 5)
                    .TextMatrix(lngRow, .ColIndex("���벿��")) = Nvl(rsBill!���벿��)
                    .TextMatrix(lngRow, .ColIndex("��˲���")) = Nvl(rsBill!��˲���)
                    .TextMatrix(lngRow, .ColIndex("�������")) = Nvl(rsBill!�������)
                    
                    lngRow = lngRow + 1
                    rsBill.MoveNext
                Loop
                .Redraw = flexRDBuffered
            End With
        End If
    End If
    
    If blnNotData Then
        'MsgBox "��ǰ����û�п�ת���ķ��ã�", vbInformation, gstrSysName
        mblnOK = True
        Exit Function
    End If
    LoadFeeData = True
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Private Sub Form_Unload(Cancel As Integer)
    Dim k As Integer
    
    For k = 1 To 3
        Call gobjComlib.SaveFlexState(vsfGrid(k), App.ProductName & "\" & Me.Name & "_" & k)
    Next
End Sub

Private Sub picBack_Resize(Index As Integer)
    On Error Resume Next
    Select Case Index
    Case Pane_������Ϣ
        shp������Ϣ.Move 10, 10, picBack(Index).ScaleWidth - 30, picBack(Index).ScaleHeight - 10
        lblShow(Lbl_����).Left = lblShow(Lbl_�Ա�).Left + lblShow(Lbl_�Ա�).Width + 500
        lblShow(Lbl_סԺ��).Left = lblShow(Lbl_����).Left + lblShow(Lbl_����).Width + 500
        lblShow(Lbl_סԺ����).Left = lblShow(Lbl_סԺ��).Left + lblShow(Lbl_סԺ��).Width + 500
        lblShow(Lbl_�ѱ�).Left = lblShow(Lbl_סԺ����).Left + lblShow(Lbl_סԺ����).Width + 500
    Case Pane_�䶯����, Pane_��������, Pane_����ȡ��
        lblShow(Index).Left = 50
        lblShow(Index).Top = 50
        With vsfGrid(Index)
            .Left = 10
            .Top = lblShow(Index).Top + lblShow(Index).Height + 30
            .Width = picBack(Index).ScaleWidth - .Left - 20
            .Height = picBack(Index).ScaleHeight - .Top
        End With
        
        If Index = Pane_�䶯���� Then
            picCboBack.Left = lblShow(Index).Left + lblShow(Index).Width
            picCboBack.Top = lblShow(Index).Top - 50
        End If
    Case Pane_���ܰ�ť
        cmdCancel.Left = picBack(Index).ScaleWidth - cmdCancel.Width - 800
        cmdCancel.Top = (picBack(Index).ScaleHeight - cmdCancel.Height) / 2 - 50
        cmdOK.Top = cmdCancel.Top
        cmdOK.Left = cmdCancel.Left - cmdOK.Width - 200
    End Select
End Sub

Public Function Upgradeҽ��ִ�мƼ�ִ��״̬(ByVal lng����ID As Long, ByVal lng��ҳId As Long) As Boolean
    '���ܣ�����"ҽ��ִ�мƼ�.ִ��״̬"
    '��Σ�
    '   lng����ID
    '   lng��ҳID
    '���أ��������򷵻�True�����򷵻�False
    '�����:99715
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandler
    strSql = _
        " Select Distinct a.Id As ҽ��id, b.No" & vbNewLine & _
        " From ����ҽ����¼ A, ����ҽ������ B, ҽ��ִ�мƼ� C" & vbNewLine & _
        " Where a.Id = b.ҽ��id And b.ҽ��id = c.ҽ��id And b.���ͺ� = c.���ͺ�" & vbNewLine & _
        "       And b.��¼���� = 2 And a.����id = [1] And a.��ҳid = [2]" & vbNewLine & _
        "       And c.ִ��״̬ Is Null"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "�ж�ҽ��ִ�мƼ�ִ��״̬�Ƿ�������", lng����ID, lng��ҳId)
    If rsTemp.RecordCount = 0 Then
        Upgradeҽ��ִ�мƼ�ִ��״̬ = True
        Exit Function
    End If
    
    '��������
    Do While Not rsTemp.EOF
        'Zl_ҽ��ִ�мƼ�_����(
        strSql = "Zl_ҽ��ִ�мƼ�_����("
        '  ҽ��id_In   ����ҽ��ִ��.ҽ��id%Type,
        strSql = strSql & "" & Nvl(rsTemp!ҽ��ID) & ","
        '  No_In       ����ҽ������.No%Type,
        strSql = strSql & "'" & Nvl(rsTemp!NO) & "',"
        '  ��¼����_In ����ҽ������.��¼����%Type
        strSql = strSql & "" & 2 & ")"
        gobjDatabase.ExecuteProcedure strSql, "��������"
        
        rsTemp.MoveNext
    Loop
    
    Upgradeҽ��ִ�мƼ�ִ��״̬ = True
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Private Sub cboDate_Click()
    Dim strSql As String, rsBill As ADODB.Recordset
    Dim lngRow As Long
    
    On Error GoTo ErrHandler
    If cboDate.Tag = cboDate.Text Then Exit Sub
    cboDate.Tag = cboDate.Text
    If cboDate.ListIndex < 0 Then Exit Sub
    
    '��ս�������
    vsfGrid(Pane_�䶯����).Clear 1
    vsfGrid(Pane_�䶯����).Rows = vsfGrid(Pane_�䶯����).FixedRows
    vsfGrid(Pane_��������).Clear 1
    vsfGrid(Pane_��������).Rows = vsfGrid(Pane_��������).FixedRows
    
    strSql = _
        " Select Decode(h.��¼״̬, 0, '���ʻ���', '����') As ��������, a.No, a.�շ�ϸĿID, b.���� As �շ���Ŀ," & vbNewLine & _
        "        c.���� As ԭ����, d.���� As Ŀ�겡��, a.����, a.����, a.ʵ�ս��" & vbNewLine & _
        " From סԺ���ü�¼ H, ���ñ䶯��¼ A, �շ���ĿĿ¼ B, ���ű� C, ���ű� D" & vbNewLine & _
        " Where h.Id = a.����id And a.�շ�ϸĿid = b.Id And a.ԭ����id = c.Id And a.Ŀ�겡��id = d.Id" & vbNewLine & _
        "       And a.����id = [1] And a.��ҳid = [2] And a.�䶯ʱ�� = [3] And ״̬ In (0, 1)" & vbNewLine & _
        " Order By No"
    Set rsBill = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, mlng����ID, mlng��ҳID, CDate(cboDate.Text))
    If rsBill.RecordCount > 0 Then
        With vsfGrid(Pane_�䶯����)
            .Redraw = flexRDNone
            .Rows = .FixedRows + rsBill.RecordCount
            lngRow = .FixedRows
            Do While Not rsBill.EOF
                .TextMatrix(lngRow, .ColIndex("��������")) = Nvl(rsBill!��������)
                .TextMatrix(lngRow, .ColIndex("NO")) = Nvl(rsBill!NO)
                .TextMatrix(lngRow, .ColIndex("�շ�ϸĿID")) = Nvl(rsBill!�շ�ϸĿID)
                .TextMatrix(lngRow, .ColIndex("�շ���Ŀ")) = Nvl(rsBill!�շ���Ŀ)
                .TextMatrix(lngRow, .ColIndex("ת��-����")) = Nvl(rsBill!ԭ����)
                .TextMatrix(lngRow, .ColIndex("ת��-����")) = FormatEx(Val(Nvl(rsBill!����)), 5)
                .TextMatrix(lngRow, .ColIndex("ת��-����")) = Nvl(rsBill!Ŀ�겡��)
                .TextMatrix(lngRow, .ColIndex("ת��-����")) = FormatEx(Val(Nvl(rsBill!����)), 5)
                .TextMatrix(lngRow, .ColIndex("����")) = Format(Nvl(rsBill!����), gSysPara.Price_Decimal.strFormt_VB)
                .TextMatrix(lngRow, .ColIndex("���")) = Format(Nvl(rsBill!ʵ�ս��), gSysPara.Money_Decimal.strFormt_VB)
                
                lngRow = lngRow + 1
                rsBill.MoveNext
            Loop
            .Redraw = flexRDBuffered
        End With
    End If
    
    strSql = _
        " Select Decode(h.��¼״̬, 0, '���ʻ���', '����') As ��������, a.No, a.�շ�ϸĿID, b.���� As �շ���Ŀ, c.���� As ԭ����," & vbNewLine & _
        "        Sum(a.����) As ����, Decode(a.״̬, 2, '��������', 'ȡ������') As ������ʽ" & vbNewLine & _
        " From סԺ���ü�¼ H, ���ñ䶯��¼ A, �շ���ĿĿ¼ B, ���ű� C" & vbNewLine & _
        " Where h.Id = a.����id And a.�շ�ϸĿid = b.Id And a.ԭ����id = c.Id And a.����id = [1] And a.��ҳid = [2]" & vbNewLine & _
        "       And a.�䶯ʱ�� = [3] And a.״̬ In (2, 3)" & vbNewLine & _
        " Group By Decode(h.��¼״̬, 0, '���ʻ���', '����'), a.No, a.�շ�ϸĿid, b.����, c.����, a.����, Decode(a.״̬, 2, '��������', 'ȡ������')" & vbNewLine & _
        " Order By No"
    Set rsBill = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, mlng����ID, mlng��ҳID, CDate(cboDate.Text))
    If rsBill.RecordCount > 0 Then
        With vsfGrid(Pane_��������)
            .Redraw = flexRDNone
            .Rows = .FixedRows + rsBill.RecordCount
            lngRow = .FixedRows
            Do While Not rsBill.EOF
                .TextMatrix(lngRow, .ColIndex("��������")) = Nvl(rsBill!��������)
                .TextMatrix(lngRow, .ColIndex("NO")) = Nvl(rsBill!NO)
                .TextMatrix(lngRow, .ColIndex("�շ�ϸĿID")) = Nvl(rsBill!�շ�ϸĿID)
                .TextMatrix(lngRow, .ColIndex("�շ���Ŀ")) = Nvl(rsBill!�շ���Ŀ)
                .TextMatrix(lngRow, .ColIndex("����")) = Nvl(rsBill!ԭ����)
                .TextMatrix(lngRow, .ColIndex("����")) = FormatEx(Val(Nvl(rsBill!����)), 5)
                .TextMatrix(lngRow, .ColIndex("������ʽ")) = Nvl(rsBill!������ʽ)
                
                lngRow = lngRow + 1
                rsBill.MoveNext
            Loop
            .Redraw = flexRDBuffered
        End With
    End If
    Exit Sub
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strSql As String
    Dim lngRow As Long
    Dim bytסԺ״̬ As Byte, byt��˱�־ As Byte
    
    On Error GoTo ErrHandler
    
     '���������������δ��˵Ĳ�����ת����
    With vsfGrid(Pane_��������)
        For lngRow = .FixedRows To .Rows - 1
            If .TextMatrix(lngRow, .ColIndex("������ʽ")) = "��ֹת����" Then
                MsgBox "���ݡ�" & .TextMatrix(lngRow, .ColIndex("NO")) & "����δ������ˣ���ֹת����������", vbInformation, gstrSysName
                Exit Sub
            End If
        Next
    End With
    
    bytסԺ״̬ = Val(lblShow(Lbl_סԺ��).Tag)
    If gSysPara.blnδ��ƽ�ֹ���� And bytסԺ״̬ = 1 Then
        MsgBox "����δ��ƣ���ֹ�Բ�����ط��õĲ�������˽�ֹת����������", vbInformation, gstrSysName
        Exit Sub
    End If
    If gSysPara.byt������˷�ʽ = 1 Then
        byt��˱�־ = Val(lblShow(Lbl_סԺ����).Tag)
        If byt��˱�־ = 1 Then
            MsgBox "�ò���Ŀǰ������˷��ã����ܽ��з�����ص�������˽�ֹת����������", vbInformation, gstrSysName
            Exit Sub
        ElseIf byt��˱�־ = 2 Then
            MsgBox "�ò���Ŀǰ�Ѿ�����˷�����ˣ����ܽ��з�����ص�������˽�ֹת����������", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    If mbyt���� = Fun_ת�������� Then 'ת���������˳�
        mblnOK = True
        Unload Me
        Exit Sub
    End If
    
    Set mcllSQL = New Collection
    '  --����:����ת�������õ�ת�룬ת������
    '  --���:����_In: 0-�����䶯,1-���������䶯
    '  --   1������_IN=0(�����䶯)ʱ
    '  --       �䶯ID_In: ԭ�����ı䶯��¼��ID
    '  --       ԭ����ID_IN��ԭ����ID
    '  --       Ŀ�겡��ID_IN:Ŀ�겡��ID
    '  --   2������_IN=1(���������䶯)ʱ
    '  --       �䶯ID_In: �ָ���ԭʼ�����ı䶯��¼��ID
    '  --       ԭ����ID_IN���������Ĳ���ID
    '  --       Ŀ�겡��ID_IN:�ָ���ԭʼ����ID
    '  --ת�룬ת������:
    '  --1.����ִ�еķ�ҩƷ���������ϣ��������Ϊ
    '  --   1)��ԭ��¼�������ʴ���
    '  --   2)����һ���²����ķ��ã����˿��ң�����ʱ�䲻��
    '  --2.����ִ�е�ҩƷ����������
    '  --   ��������˵Ĵ�����ת����ʱ�Ľ����н���ȷ��(���Դ�ӡ�˲��嵥)����ת���������ʱ��ȷ�ϡ�
    '  --   a)������ԭ����ͨ�����������������²����ֹ������ģ�
    '  --   b)����ת����ʱ���Զ������������룬����Ѿ���������ˣ���ѯ����ʾ���Ҳ������ķ��ô����ֹ�ȥ����
    'Zl_Turntoward_Fee
    strSql = "Zl_Turntoward_Fee("
    '(
    '  ����_In       Number,
    strSql = strSql & "" & mbyt���� & ","
    '  ����id_In     ������ҳ.����id%Type,
    strSql = strSql & "" & mlng����ID & ","
    '  ��ҳid_In     ������ҳ.����id%Type,
    strSql = strSql & "" & mlng��ҳID & ","
    '  �䶯id_In   ���˱䶯��¼.Id%Type,
    strSql = strSql & "" & mlng�䶯id & ","
    '  ԭ����id_In   ������ҳ.��ǰ����id%Type,
    strSql = strSql & "" & mlngԭ����id & ","
    '  Ŀ�겡��id_In ������ҳ.��ǰ����id%Type,
    strSql = strSql & "" & mlngĿ�겡��id & ","
    '  ����Ա���_In סԺ���ü�¼.����Ա���%Type,
    strSql = strSql & "'" & UserInfo.��� & "',"
    '  ����Ա����_In סԺ���ü�¼.����Ա����%Type,
    strSql = strSql & "'" & UserInfo.���� & "')"
    '  �䶯ʱ��_In   סԺ���ü�¼.�Ǽ�ʱ��%Type := Null
    
    mcllSQL.Add strSql
    mblnOK = True
    Unload Me
    Exit Sub
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Sub

