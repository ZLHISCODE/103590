VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmPayNoEdit 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9960
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   9960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picPayNO 
      AutoRedraw      =   -1  'True
      Height          =   6030
      Left            =   0
      ScaleHeight     =   5970
      ScaleWidth      =   9420
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   9480
      Begin VB.PictureBox picDown 
         BorderStyle     =   0  'None
         Height          =   1305
         Left            =   -60
         ScaleHeight     =   1305
         ScaleWidth      =   9930
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   4785
         Width           =   9930
         Begin VB.TextBox txtInfo 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   0
            Left            =   810
            TabIndex        =   11
            Top             =   135
            Width           =   8820
         End
         Begin VB.TextBox txtInfo 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   1
            Left            =   810
            Locked          =   -1  'True
            TabIndex        =   13
            Top             =   510
            Width           =   3240
         End
         Begin VB.TextBox txtInfo 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   2
            Left            =   6390
            Locked          =   -1  'True
            TabIndex        =   15
            Top             =   510
            Width           =   3240
         End
         Begin VB.TextBox txtInfo 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   3
            Left            =   825
            Locked          =   -1  'True
            TabIndex        =   17
            Top             =   870
            Width           =   3240
         End
         Begin VB.TextBox txtInfo 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   4
            Left            =   6390
            Locked          =   -1  'True
            TabIndex        =   19
            Top             =   885
            Width           =   3240
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            Caption         =   "����˵��"
            Height          =   180
            Index           =   4
            Left            =   0
            TabIndex        =   10
            Top             =   195
            Width           =   750
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            Caption         =   "������"
            Height          =   180
            Index           =   5
            Left            =   180
            TabIndex        =   12
            Top             =   570
            Width           =   570
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            Caption         =   "��������"
            Height          =   180
            Index           =   6
            Left            =   5580
            TabIndex        =   14
            Top             =   570
            Width           =   750
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            Caption         =   "�����"
            Height          =   180
            Index           =   7
            Left            =   180
            TabIndex        =   16
            Top             =   945
            Width           =   570
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            Caption         =   "�������"
            Height          =   180
            Index           =   8
            Left            =   5580
            TabIndex        =   18
            Top             =   945
            Width           =   750
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsPayEdit 
         Height          =   2610
         Left            =   165
         TabIndex        =   6
         Top             =   1500
         Width           =   5055
         _cx             =   8916
         _cy             =   4604
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
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483644
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   16777215
         GridColor       =   -2147483633
         GridColorFixed  =   12632256
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
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
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmPayNoEdit.frx":0000
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
         ExplorerBar     =   7
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
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
      End
      Begin VSFlex8Ctl.VSFlexGrid vs��Ԥ�� 
         Height          =   2625
         Left            =   5205
         TabIndex        =   8
         Top             =   1500
         Width           =   4605
         _cx             =   8123
         _cy             =   4630
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
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483644
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   16777215
         GridColor       =   -2147483633
         GridColorFixed  =   12632256
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
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
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmPayNoEdit.frx":00A8
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
         ExplorerBar     =   7
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
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
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����֪ͨ��"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   345
         Left            =   30
         TabIndex        =   22
         Top             =   90
         Width           =   9780
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "���θ���:"
         Height          =   180
         Index           =   4
         Left            =   7950
         TabIndex        =   5
         Top             =   1260
         Width           =   810
      End
      Begin VB.Label txtNo 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   8355
         TabIndex        =   21
         Top             =   390
         Width           =   1425
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "NO."
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   210
         Index           =   10
         Left            =   8055
         TabIndex        =   20
         Top             =   450
         Width           =   315
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "˰��ǼǺ�:"
         Height          =   180
         Index           =   3
         Left            =   210
         TabIndex        =   4
         Top             =   1290
         Width           =   990
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "��������:"
         Height          =   180
         Index           =   2
         Left            =   405
         TabIndex        =   3
         Top             =   1050
         Width           =   810
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "��ַ�绰:"
         Height          =   180
         Index           =   1
         Left            =   390
         TabIndex        =   2
         Top             =   825
         Width           =   810
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "��λ����:"
         Height          =   180
         Index           =   7
         Left            =   390
         TabIndex        =   1
         Top             =   600
         Width           =   810
      End
      Begin VB.Label lblEdit 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "���γ�Ԥ����:"
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   5
         Left            =   5205
         TabIndex        =   9
         Top             =   4110
         Width           =   4605
      End
      Begin VB.Label lblEdit 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "�ϼ�:"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   6
         Left            =   165
         TabIndex        =   7
         Top             =   4095
         Width           =   5055
      End
   End
End
Attribute VB_Name = "frmPayNoEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '
Private mlngModule As Long
Private mstrPrivs As String
Private mfrmMain As Form
Private mblnChange As Boolean

Private mEditType As gEditType  '�༭����
Private mstrNo As String        '���ݺ�
Private mint��¼״̬ As Integer '��¼״̬
Private mlng������� As Long    '�������
Private mlng��λID As Long      '��λID
Private mdbl����Ӧ�� As Double, mdbl����Ԥ�� As Double

Private mblnEdit As Boolean     '�Ƿ�����༭
Private Enum mlblIdx
    idx_lbl��ַ�绰 = 1
    idx_lbl�������� = 2
    idx_lbl˰��ǼǺ� = 3
    idx_lbl���θ��� = 4
    idx_lbl��Ԥ���ϼ� = 5
    idx_lbl����ϼ� = 6
    idx_lbl��λ���� = 7
End Enum
Private mrs���㷽ʽ As ADODB.Recordset
Private mint��� As Integer

Public Event initCard(ByVal lng������� As Long, ByVal lng��λID As Long, ByVal str��λ���� As String)
Public Event zlChangeData(ByVal blnChange As Boolean)

Private Sub SetEditPro()
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:���ñ༭����
    '--�����:
    '--������:
    '--��  ��:
    '-----------------------------------------------------------------------------------------------------------
    Dim i As Integer
    For i = 0 To 4
        txtInfo(i).Enabled = mblnEdit And i = 0
    Next
    If mEditType = g��� Then
        vsPayEdit.Editable = flexEDKbdMouse
    Else
        vsPayEdit.Editable = IIf(mblnEdit, flexEDKbdMouse, flexEDNone)
    End If
End Sub

Private Sub InitvsPayEdit()
    '-----------------------------------------------------------------------------------------------------------
    '����:��ʼ������ؼ���Ĭ������
    '���:
    '����:
    '����:
    '����:���˺�
    '����:2008-08-19 11:55:12
    '-----------------------------------------------------------------------------------------------------------
    'Dim rsTemp As New ADODB.Recordset
    '����27930 by lesfeng 2010-03-23
    If mint��� = 0 Then
        gstrSQL = "Select ���㷽ʽ From ���㷽ʽӦ�� Where Ӧ�ó���='������' Order by ȱʡ��־ desc"
    Else
        gstrSQL = "Select '���' As ���㷽ʽ From dual "
'        gstrSQL = "Select '  ' As ���㷽ʽ From dual Union All " & _
'                  "Select ���㷽ʽ From ���㷽ʽӦ�� Where Ӧ�ó���='������'"
    End If
    On Error GoTo errHandle
    Set mrs���㷽ʽ = New ADODB.Recordset
    zlDatabase.OpenRecordset mrs���㷽ʽ, gstrSQL, Me.Caption
    With vsPayEdit
        .ColComboList(.ColIndex("���ʽ")) = .BuildComboList(mrs���㷽ʽ, "���㷽ʽ", "���㷽ʽ")
    End With
    Call vs��Ԥ��_LostFocus
    Call vsPayEdit_LostFocus
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub picDown_Resize()
    Err = 0: On Error Resume Next
    With picDown
        txtInfo(0).Width = .ScaleWidth - txtInfo(0).Left
        txtInfo(2).Left = .ScaleWidth - txtInfo(2).Width
        txtInfo(4).Left = txtInfo(2).Left
        lblInfo(6).Left = txtInfo(2).Left - lblInfo(6).Width
        lblInfo(8).Left = lblInfo(6).Left
    End With
End Sub

Private Sub txtInfo_Change(Index As Integer)
    mblnChange = True
    RaiseEvent zlChangeData(mblnChange)
End Sub

Private Sub txtInfo_GotFocus(Index As Integer)
    zlCommFun.OpenIme True
    zlControl.TxtSelAll txtInfo(Index)
End Sub

Private Sub txtInfo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    zlCommFun.PressKey vbKeyTab
End Sub

Private Sub Form_Load()
    mblnEdit = False
    '����27930 by lesfeng 2010-03-23
'    lblTitle.Caption = GetUnitName & lblTitle.Caption
    RestoreWinState Me, App.ProductName
    zl_vsGrid_Para_Restore mlngModule, vsPayEdit, Me.Caption, "�����б�"
    zl_vsGrid_Para_Restore mlngModule, vs��Ԥ��, Me.Caption, "��Ԥ���б�"
'    Call InitvsPayEdit
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    With picPayNO
        .Left = ScaleLeft
        .Width = ScaleWidth
        .Top = ScaleTop
        .Height = ScaleHeight
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    zl_vsGrid_Para_Save mlngModule, vsPayEdit, Me.Caption, "�����б�"
    zl_vsGrid_Para_Save mlngModule, vs��Ԥ��, Me.Caption, "��Ԥ���б�"
End Sub

Private Sub picPayNO_Resize()
    Err = 0: On Error Resume Next
    With picPayNO
        txtNo.Left = .ScaleWidth - txtNo.Width - 50
        lblInfo(10).Left = txtNo.Left - lblInfo(10).Width
        lblTitle.Left = .ScaleLeft
        lblTitle.Width = .ScaleWidth
        picDown.Top = .ScaleHeight - picDown.Height
        picDown.Width = .ScaleWidth
        picDown.Left = .ScaleLeft
        '����27930 by lesfeng 2010-03-23
        If mint��� = 0 Then
            lblEdit(mlblIdx.idx_lbl����ϼ�).Top = picDown.Top - lblEdit(mlblIdx.idx_lbl����ϼ�).Height - 30
            lblEdit(mlblIdx.idx_lbl��Ԥ���ϼ�).Top = lblEdit(mlblIdx.idx_lbl����ϼ�).Top
            lblEdit(mlblIdx.idx_lbl��Ԥ���ϼ�).Width = .ScaleWidth - lblEdit(mlblIdx.idx_lbl��Ԥ���ϼ�).Left - 100
            lblEdit(mlblIdx.idx_lbl��Ԥ���ϼ�).Height = lblEdit(mlblIdx.idx_lbl����ϼ�).Height
            vs��Ԥ��.Width = .ScaleWidth - vs��Ԥ��.Left - 100
            vs��Ԥ��.Height = lblEdit(mlblIdx.idx_lbl��Ԥ���ϼ�).Top - vs��Ԥ��.Top + 10
            
            vsPayEdit.Top = vs��Ԥ��.Top
            vsPayEdit.Height = vs��Ԥ��.Height
            lblEdit(mlblIdx.idx_lbl���θ���).Left = .ScaleWidth - lblEdit(mlblIdx.idx_lbl���θ���).Width - 50
            
            lblTitle.Caption = GetUnitName & lblTitle.Caption
        Else
            lblEdit(mlblIdx.idx_lbl����ϼ�).Top = picDown.Top - lblEdit(mlblIdx.idx_lbl����ϼ�).Height - 30
            lblEdit(mlblIdx.idx_lbl����ϼ�).Width = .ScaleWidth - lblEdit(mlblIdx.idx_lbl����ϼ�).Left - 100
            
            vsPayEdit.Width = .ScaleWidth - vsPayEdit.Left - 100
            vsPayEdit.Height = lblEdit(mlblIdx.idx_lbl����ϼ�).Top - vsPayEdit.Top + 10
            
            lblEdit(mlblIdx.idx_lbl���θ���).Left = .ScaleWidth - lblEdit(mlblIdx.idx_lbl���θ���).Width - 50
            
            lblTitle.Caption = GetUnitName & "��Ǹ��"
            vs��Ԥ��.Visible = False
            lblEdit(mlblIdx.idx_lbl��Ԥ���ϼ�).Visible = False
            lblEdit(mlblIdx.idx_lbl���θ���).Caption = "���α�Ǹ���:"
            lblInfo(4).Caption = "���˵��"
        End If
    End With
End Sub

Private Function initCard(ByRef intErrInfor As Integer) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:��ʼ����Ƭ��Ϣ
    '���:
    '����:intErrInfor-���ش�����Ϣ����(1-�Ѿ�ɾ��,2-�Ѿ����)
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-08-19 13:00:28
    '-----------------------------------------------------------------------------------------------------------
    Dim i As Integer
    Dim rsTemp As New Recordset
    '��ʼ���
    With vsPayEdit
        .Clear 1
        .Rows = 2
        '����27930 by lesfeng 2010-03-23
        If mint��� = 1 Then
            .ColHidden(.ColIndex("���ʽ")) = True: .ColWidth(.ColIndex("���ʽ")) = 0
        Else
            .ColHidden(.ColIndex("���ʽ")) = False: .ColWidth(.ColIndex("���ʽ")) = 1200
        End If
    End With
    With vs��Ԥ��
        .Clear 1
        .Rows = 2
    End With
    On Error GoTo errHandle
    Select Case mEditType
        Case g����
                txtInfo(1).Text = UserInfo.����
                txtInfo(2).Text = Format(zlDatabase.Currentdate, "yyyy-MM-dd hh:mm:ss")
                txtInfo(3).Text = ""
                txtInfo(4).Text = ""
        Case g���, g�޸�, g�鿴, gȡ��, gԤ��
            '��ȡ�������
            'by lesfeng 2009-12-2 �����Ż� '����27930 by lesfeng 2010-03-23
            gstrSQL = "Select ID,��¼״̬,NO,���,Ԥ����,��λID,���,���㷽ʽ,�������,ժҪ,������,��������,�����,�������,�������," & _
                      "       decode(�ܸ���־,0,'����',1,'���','���') as ��Ǹ���" & _
                      " From �����¼ Where NO=[1] And ��¼״̬=[2] order by ���"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mstrNo, mint��¼״̬)
            
            If rsTemp.EOF Then
                intErrInfor = 1
                Exit Function
            End If
            
            mlng������� = Nvl(rsTemp!�������, 0)
            mlng��λID = Nvl(rsTemp!��λID, 0)
            
            txtInfo(0).Text = Nvl(rsTemp!ժҪ)
            txtInfo(1).Text = Nvl(rsTemp!������)
            txtInfo(2).Text = Format(rsTemp!��������, "yyyy-MM-dd hh:mm:ss")
            txtInfo(3).Text = Nvl(rsTemp!�����)
            txtInfo(4).Text = Format(rsTemp!�������, "yyyy-MM-dd hh:mm:ss")
            txtNo = Nvl(rsTemp!NO)
            txtNo.Tag = Nvl(rsTemp!NO)
            If Nvl(rsTemp!�����) <> "" And mEditType = g��� Then
                intErrInfor = 2
                Exit Function
            End If
            
            If mEditType = g��� Or mEditType = gȡ�� Then
                txtInfo(3).Text = UserInfo.����
                txtInfo(4).Text = Format(zlDatabase.Currentdate, "yyyy-MM-dd hh:mm:ss")
            End If
            
            With vsPayEdit
                .Rows = rsTemp.RecordCount + 1
                i = 1
                Do While Not rsTemp.EOF
                    .TextMatrix(i, .ColIndex("��Ǹ���")) = Nvl(rsTemp!��Ǹ���)
                    .TextMatrix(i, .ColIndex("���ʽ")) = Nvl(rsTemp!���㷽ʽ)
                    .Cell(flexcpData, i, .ColIndex("���ʽ")) = Nvl(rsTemp!ID)
                    .TextMatrix(i, .ColIndex("������")) = Format(Val(Nvl(rsTemp!���)), gVbFmtString.FM_���)
                    .TextMatrix(i, .ColIndex("�������")) = Nvl(rsTemp!�������)
                    .Cell(flexcpData, i, .ColIndex("�������")) = Nvl(rsTemp!�������)
                    i = i + 1
                    rsTemp.MoveNext
                Loop
            End With
            '��ȡԤ����¼
    End Select
    
    Call zlLoadPrivder(mlng��λID)
    RaiseEvent initCard(mlng�������, mlng��λID, lblEdit(mlblIdx.idx_lbl��λ����).Tag)
    Call SetEditPro
    Call ����ϼ�
    initCard = True
    Exit Function
    
errHandle:
    If ErrCenter = 1 Then Resume
End Function

Public Function zlLoadPrivder(ByVal lng��λID As Long) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:���ع�Ӧ����Ϣ
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-08-19 13:06:23
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    '����ṩ�˹�Ӧ��ID���ȡ�ù�Ӧ����Ϣ
    On Error GoTo errHandle
    gstrSQL = "Select ����,��ַ,�绰,��������,˰��ǼǺ� From ��Ӧ�� Where ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng��λID)
    mlng��λID = lng��λID
    If Not rsTemp.EOF Then
        lblEdit(mlblIdx.idx_lbl��λ����).Caption = "��λ����:" & rsTemp!����
        lblEdit(mlblIdx.idx_lbl��λ����).Tag = Nvl(rsTemp!����)
        lblEdit(mlblIdx.idx_lbl��ַ�绰).Caption = "��ַ�绰:" & Nvl(rsTemp!��ַ) & Nvl(rsTemp!��ַ)
        lblEdit(mlblIdx.idx_lbl��������).Caption = "��������:" & Nvl(rsTemp!��������)
        lblEdit(mlblIdx.idx_lbl˰��ǼǺ�).Caption = "˰��ǼǺ�:" & Nvl(rsTemp!˰��ǼǺ�)
    Else
        lblEdit(mlblIdx.idx_lbl��λ����).Caption = "��λ����:"
        lblEdit(mlblIdx.idx_lbl��λ����).Tag = ""
        lblEdit(mlblIdx.idx_lbl��ַ�绰).Caption = "��ַ�绰:"
        lblEdit(mlblIdx.idx_lbl��������).Caption = "��������:"
        lblEdit(mlblIdx.idx_lbl˰��ǼǺ�).Caption = "˰��ǼǺ�:"
        zlLoadPrivder = False: Exit Function
    End If
    zlLoadPrivder = True
    Exit Function
    
errHandle:
    If ErrCenter = 1 Then Resume
End Function

Public Sub zlInitPara(ByVal FrmMain As Form, ByVal lngModuel As Long, ByVal strPrivs As String, ByVal int��� As Integer)
    '-----------------------------------------------------------------------------------------------------------
    '����:��ʼ����������(��һ�������ʼ��)
    '���:frmMain-���õ�������
    '     lngModuel-ģ���
    '     strPrivs-Ȩ�޴�
    '����:
    '����:
    '����:���˺�
    '����:2008-08-19 11:47:39
    '-----------------------------------------------------------------------------------------------------------
    Set mfrmMain = FrmMain: mlngModule = lngModuel: mstrPrivs = strPrivs: mint��� = int���
    Call InitvsPayEdit
End Sub

Public Function zlLoadData(ByVal EditType As gEditType, ByVal lng��λID As Long, _
    ByVal strNO As String, ByVal int��¼״̬ As Integer, _
    intErrInfor As Integer) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:�������ݽӿ�
    '���:
    '����:intErrInfor-���ش�����Ϣ����(1-�Ѿ�ɾ��,2-�Ѿ����)
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-08-19 12:53:05
    '-----------------------------------------------------------------------------------------------------------
    mEditType = EditType: mstrNo = strNO: mint��¼״̬ = int��¼״̬: mlng��λID = lng��λID
    mblnEdit = mEditType = g���� Or mEditType = g�޸�
    zlLoadData = initCard(intErrInfor)
End Function

Private Sub txtInfo_KeyPress(Index As Integer, KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txtInfo(Index), KeyAscii, m�ı�ʽ
End Sub

Private Sub txtInfo_LostFocus(Index As Integer)
    zlCommFun.OpenIme False
End Sub

Private Sub Init�������()
    '����������
    Dim dbl��� As Double, i As Integer
    If mEditType <> g���� And mEditType <> g�޸� Then Exit Sub
    dbl��� = 0
    With vsPayEdit
        For i = 1 To .Rows - 1
            dbl��� = dbl��� + Val(.TextMatrix(i, .ColIndex("������")))
        Next
        If (mdbl����Ӧ�� - mdbl����Ԥ��) - dbl��� <> 0 Then
            If .Row = .Rows - 1 And Val(.TextMatrix(.Row, .ColIndex("������"))) = 0 Then
                .TextMatrix(.Row, .ColIndex("������")) = Format((mdbl����Ӧ�� - mdbl����Ԥ��) - dbl���, gVbFmtString.FM_���)
            End If
        End If
    End With
    Call ����ϼ�
End Sub

Private Sub vsPayEdit_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    '--------------------------------------------------------------------------------
    '������صĸ�ʽ
    '���˺�:2007/09/17
    '--------------------------------------------------------------------------------
    With vsPayEdit
        Select Case Col
        Case .ColIndex("������")
            .TextMatrix(Row, .Col) = Format(Val(.TextMatrix(Row, .Col)), gVbFmtString.FM_���)
        Case .ColIndex("�������")
            If mEditType = g��� Or mEditType = g�޸� Then
                If Trim(.TextMatrix(Row, Col)) <> Trim(.Cell(flexcpData, Row, Col)) Then
                    .Cell(flexcpForeColor, Row, Col) = vbRed
                Else
                    .Cell(flexcpForeColor, Row, Col) = .ForeColor
                End If
            End If
        End Select
    End With
End Sub

Private Sub vsPayEdit_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call zl_VsGridRowChange(vsPayEdit, OldRow, NewRow, OldCol, NewCol)
End Sub

Private Sub vsPayEdit_AfterSort(ByVal Col As Long, Order As Integer)
    Call zl_VsGridAfterSort(vsPayEdit, Col, Order)
End Sub

Private Sub vsPayEdit_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsPayEdit
        Select Case Col
        Case .ColIndex("���ʽ"), .ColIndex("������")
            If mEditType <> g���� And mEditType <> g�޸� Then
                Cancel = True: Exit Sub
            End If
        Case .ColIndex("�������")
            If mEditType <> g���� And mEditType <> g�޸� And mEditType <> g��� Then
                Cancel = True: Exit Sub
            End If
        Case Else
            Cancel = True
        End Select
    End With
End Sub

Private Sub vsPayEdit_ChangeEdit()
    mblnChange = True
    RaiseEvent zlChangeData(mblnChange)
End Sub

Private Sub vsPayEdit_EnterCell()
    If mEditType <> g���� And mEditType <> g�޸� Then Exit Sub
    
    With vsPayEdit
        .EditMaxLength = 0
        Select Case .Col
        Case .ColIndex("���㷽ʽ")
            '             .ColComboList(.Col) = "..."
        Case .ColIndex("������")
            .EditMaxLength = 16
        Case .ColIndex("�������")
            .EditMaxLength = 10
        End Select
    End With
End Sub

Private Sub vsPayEdit_GotFocus()
    zl_VsGridGotFocus vsPayEdit
End Sub

Private Sub vsPayEdit_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngCol As Long
    
    With vsPayEdit
        If (.Col = .ColIndex("���㷽ʽ")) And KeyCode <> vbKeyReturn Then
           ' .ColComboList(.Col) = ""
        End If
        
        If KeyCode = vbKeyDelete And (mEditType = g���� Or mEditType = g�޸�) Then
            If .Row = .Rows - 1 And .Row = 1 Then
                For lngCol = 0 To .Cols - 1
                    .TextMatrix(.Row, lngCol) = ""
                    .Cell(flexcpData, .Row, lngCol) = ""
                Next
            Else
                .RemoveItem .Row
            End If
            Call Init�������
        End If
    End With
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    With vsPayEdit
        Select Case .Col
        Case .ColIndex("���ʽ")
            If Trim(.TextMatrix(.Row, .Col)) = "" Then
               If zlControl.IsCtrlSetFocus(txtInfo(0)) Then
                  zlControl.IsCtrlSetFocus txtInfo(0)
               Else
                  zlCommFun.PressKey vbKeyTab
               End If
                Exit Sub
            End If
        End Select
        Call zlVsMoveGridCell(vsPayEdit, vsPayEdit.ColIndex("���ʽ"), vsPayEdit.Cols - 1, mblnEdit)
        If mblnEdit Then
            Call Init�������
            '����Ĭ�ϵĽ��㷽ʽ
            Call Local���㷽ʽ
        End If
    End With
End Sub

Private Sub Local���㷽ʽ()
    '-----------------------------------------------------------------------------------------------------------
    '����:���㷽ʽ��λ
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-08-21 13:45:52
    '-----------------------------------------------------------------------------------------------------------
    If mrs���㷽ʽ Is Nothing Then Exit Sub
    If mrs���㷽ʽ.State <> 1 Then Exit Sub
    With vsPayEdit
        If mrs���㷽ʽ.RecordCount <> 0 Then mrs���㷽ʽ.MoveFirst
        If Val(.TextMatrix(.Row, .ColIndex("������"))) <> 0 _
            And Trim(.TextMatrix(.Row, .ColIndex("���ʽ"))) = "" Then
            If mrs���㷽ʽ.EOF = False Then mrs���㷽ʽ.MoveFirst
            If .Row > 1 Then
                If Trim(.TextMatrix(.Row - 1, .ColIndex("���ʽ"))) <> "" Then
                    mrs���㷽ʽ.Find "���㷽ʽ='" & Trim(.TextMatrix(.Row - 1, .ColIndex("���ʽ"))) & "'"
                    If mrs���㷽ʽ.EOF = False Then mrs���㷽ʽ.MoveNext
                    If mrs���㷽ʽ.EOF = False Then
                        .TextMatrix(.Row, .ColIndex("���ʽ")) = Nvl(mrs���㷽ʽ!���㷽ʽ)
                    End If
                End If
            ElseIf mrs���㷽ʽ.EOF = False Then
                .TextMatrix(.Row, .ColIndex("���ʽ")) = Nvl(mrs���㷽ʽ!���㷽ʽ)
            End If
        End If
    End With
End Sub

Private Sub vsPayEdit_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Dim intCol As Integer
    Dim strKey As String
    If KeyCode <> vbKeyReturn Then Exit Sub
    With vsPayEdit
        Select Case Col
        Case .ColIndex("���ʽ")
        Case .ColIndex("������")
            .TextMatrix(Row, Col) = Format(Val(strKey), gVbFmtString.FM_���)
        Case Else
        End Select
        Call zlVsMoveGridCell(vsPayEdit, .ColIndex("���ʽ"), .Cols - 1, mblnEdit)
        Call Init�������
        '����Ĭ�ϵĽ��㷽ʽ
        Call Local���㷽ʽ
    End With
 End Sub
 
Private Sub vsPayEdit_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub

Private Sub vsPayEdit_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0: Exit Sub
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Exit Sub
    End If
    With vsPayEdit
        Select Case Col
        Case .ColIndex("�������")
            Call VsFlxGridCheckKeyPress(vsPayEdit, Row, Col, KeyAscii, m�ı�ʽ)
        Case .ColIndex("������")
            '��Ҫ���ܴ����˿����
            Call VsFlxGridCheckKeyPress(vsPayEdit, Row, Col, KeyAscii, m�����ʽ)
        Case Else
        End Select
    End With
End Sub

Private Sub vsPayEdit_LostFocus()
    zlCommFun.OpenIme False
    Call zl_VsGridLOSTFOCUS(vsPayEdit)
End Sub

Private Sub vsPayEdit_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim strKey As String, intCol As Integer, dbl��� As Double, i As Long
    
    If mEditType <> g���� And mEditType <> g�޸� Then
        If Col <> vsPayEdit.ColIndex("�������") And mEditType <> g��� Then
            Cancel = True: Exit Sub
        End If
    End If
    With vsPayEdit
        strKey = Trim(.EditText): strKey = Replace(strKey, Chr(vbKeyReturn), ""): strKey = Replace(strKey, Chr(10), "")
        Select Case Col
        Case .ColIndex("�������")
            If strKey <> "" Then
                If LenB(StrConv(strKey, vbFromUnicode)) > 10 Then
                    ShowMsgbox "������볬��,���������5�����ֻ�10���ַ�!"
                    zlControl.IsCtrlSetFocus vsPayEdit
                    Cancel = True
                    Exit Sub
                End If
                If InStr(1, strKey, "'") <> 0 Then
                    ShowMsgbox "������벻�����뵥����!"
                    zlControl.IsCtrlSetFocus vsPayEdit
                    Cancel = True
                    Exit Sub
                End If
            End If
        Case .ColIndex("������")
            If strKey <> "" Then
                If Not IsNumeric(strKey) Then
                    ShowMsgbox "�������������,������!"
                    zlControl.IsCtrlSetFocus vsPayEdit
                    Cancel = True
                    Exit Sub
                End If
    '            If Val(strKey) < 0 Then
    '                ShowMsgbox "�������С����,������!"
    '                zlCtlSetFocus vsPayEdit, True
    '                Cancel = True
    '                Exit Sub
    '            End If
                If Abs(Val(strKey)) > 10 ^ 12 - 1 Then
                    ShowMsgbox "������ֻ����-" & 10 ^ 12 - 1 & "��" & 10 ^ 12 - 1 & "֮�������,������!"
                    zlControl.IsCtrlSetFocus vsPayEdit
                    Cancel = True
                    Exit Sub
                End If
                
                dbl��� = 0
                For i = 1 To .Rows - 1
                    If i <> .Row Then
                        dbl��� = dbl��� + Val(.TextMatrix(i, .ColIndex("������")))
                    End If
                Next
    '            dbl��� = (mdbl����Ӧ�� - mdbl����Ԥ��) - dbl���
    '            dbl��� = dbl��� - Val(strKey)
    '            If dbl��� < 0 Then
    '                ShowMsgbox "��������ܶ�!"
    '                zlCtlSetFocus vsPayEdit, True
    '                Cancel = True
    '                Exit Sub
    '            End If
            End If
        End Select
    End With
End Sub

Private Sub ����ϼ�()
    Dim lngRow As Long, dblCount As Double
   '��ȡ����ϼ���
    With vsPayEdit
        For lngRow = 1 To .Rows - 1
            dblCount = dblCount + Val(.TextMatrix(lngRow, .ColIndex("������")))
        Next
    End With
    '����27930 by lesfeng 2010-03-23
    If mint��� = 0 Then
        lblEdit(mlblIdx.idx_lbl����ϼ�).Caption = "����ϼ�:" & Format(dblCount, gVbFmtString.FM_���) & "Ԫ"
    Else
        lblEdit(mlblIdx.idx_lbl����ϼ�).Caption = "��ǽ���ϼ�:" & Format(dblCount, gVbFmtString.FM_���) & "Ԫ"
    End If
End Sub

Private Sub vs��Ԥ��_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call zl_VsGridRowChange(vs��Ԥ��, OldRow, NewRow, OldCol, NewCol)
End Sub

Private Sub vs��Ԥ��_GotFocus()
    Call zl_VsGridGotFocus(vs��Ԥ��)
End Sub

Private Sub vs��Ԥ��_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    With vs��Ԥ��
        Select Case .Col
        Case .ColIndex("���ʽ")
        Case .Cols - 1
            If .Row = .Rows - 1 Then
                zlCommFun.PressKey vbKeyTab
                Exit Sub
            End If
        End Select
        End With
    Call zlVsMoveGridCell(vs��Ԥ��, vs��Ԥ��.ColIndex("���ʽ"), vs��Ԥ��.Cols - 1, False)
    
End Sub
'��������
Public Property Get zldbl����Ӧ��() As Double
    zldbl����Ӧ�� = mdbl����Ӧ��
End Property

Public Property Let zldbl����Ӧ��(ByVal vNewValue As Double)
    mdbl����Ӧ�� = vNewValue
     
    Call InitPayData
End Property

Public Property Get zldbl����Ԥ��() As Double
    zldbl����Ԥ�� = mdbl����Ԥ��
End Property

Public Property Let zldbl����Ԥ��(ByVal vNewValue As Double)
    mdbl����Ԥ�� = vNewValue
    Call InitPayData
End Property

Private Sub InitPayData()
    '-----------------------------------------------------------------------------------------------------------
    '����:��ʼ���������
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-08-19 15:48:58
    '-----------------------------------------------------------------------------------------------------------
    lblEdit(mlblIdx.idx_lbl��Ԥ���ϼ�).Caption = "��Ԥ����ϼƣ�" & Format(mdbl����Ԥ��, gVbFmtString.FM_���) & "Ԫ"
    '����27930 by lesfeng 2010-03-23
    If mint��� = 0 Then
        lblEdit(mlblIdx.idx_lbl���θ���).Caption = "���θ��" & Format(mdbl����Ӧ��, gVbFmtString.FM_���) & "Ԫ"
    Else
        lblEdit(mlblIdx.idx_lbl���θ���).Caption = "���α�Ǹ��" & Format(mdbl����Ӧ��, gVbFmtString.FM_���) & "Ԫ"
    End If
    With vsPayEdit
    
        If .Rows = 2 Then
            .Row = 1
'            If .TextMatrix(.Row, .ColIndex("���ʽ")) = "" Then
                .TextMatrix(.Row, .ColIndex("������")) = Format(mdbl����Ӧ�� - mdbl����Ԥ��, gVbFmtString.FM_���)
'            End If
            '����27930 by lesfeng 2010-03-23
            If mint��� = 0 Then
                .TextMatrix(.Row, .ColIndex("��Ǹ���")) = "����"
            Else
                .TextMatrix(.Row, .ColIndex("��Ǹ���")) = "���"
            End If
        End If
        '����Ĭ�ϵĽ��㷽ʽ
        Call Local���㷽ʽ
    End With
End Sub

Public Function zlValidData() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��֤���ݵĺϷ���
    '--�����:
    '--������:
    '--��  ��:��֤�Ϸ�,����True,����=false
    '-----------------------------------------------------------------------------------------------------------
    Dim strTemp As String
    Dim intIndex As Integer, lngRow As Long, dblCount As Double
    
    With vsPayEdit
        For lngRow = 1 To .Rows - 1
            If Trim(.TextMatrix(lngRow, .ColIndex("���ʽ"))) <> "" Then
                strTemp = Trim(.TextMatrix(lngRow, .ColIndex("������")))
                If strTemp = "" Then
                    ShowMsgbox "�������������!"
                    zlControl.IsCtrlSetFocus vsPayEdit
                    Exit Function
                End If
                If Not IsNumeric(strTemp) Then
                    ShowMsgbox "�������������,������!"
                    zlControl.IsCtrlSetFocus vsPayEdit
                    Exit Function
                End If
                If Abs(Val(strTemp)) > 10 ^ 12 - 1 Then
                    ShowMsgbox "������ֻ����-" & 10 ^ 12 - 1 & "��" & 10 ^ 12 - 1 & "֮�������,������!"
                    zlControl.IsCtrlSetFocus vsPayEdit
                    Exit Function
                End If
                
                dblCount = dblCount + Val(strTemp)
                strTemp = Trim(.TextMatrix(lngRow, .ColIndex("�������")))
                If strTemp <> "" Then
                    If LenB(StrConv(strTemp, vbFromUnicode)) > 10 Then
                        ShowMsgbox "������볬��,���������5�����ֻ�10���ַ�!"
                        zlControl.IsCtrlSetFocus vsPayEdit
                        Exit Function
                    End If
                    If InStr(1, strTemp, "'") <> 0 Then
                        ShowMsgbox "������벻�����뵥����!"
                        zlControl.IsCtrlSetFocus vsPayEdit
                        Exit Function
                    End If
                End If
            End If
        Next
    End With
    
    If Round(mdbl����Ӧ�� - (dblCount + mdbl����Ԥ��), g_С��λ��.���С��) <> 0 Then
        ShowMsgbox "�����ƽ,���鸶��������ⵥ" & vbCrLf & "��Ʊ����Ԥ����֮���Ƿ���ͬ!"
        zlControl.IsCtrlSetFocus vsPayEdit
        Exit Function
    End If
    If mdbl����Ӧ�� = 0 Then
        ShowMsgbox "���β������κ�Ӧ����¼,����!"
        zlControl.IsCtrlSetFocus vsPayEdit
        Exit Function
    End If
    
    If LenB(StrConv(txtInfo(0).Text, vbFromUnicode)) > 50 Then
        ShowMsgbox "����˵���ĳ��ȳ���!(���Ϊ50���ַ���25������)"
        zlControl.IsCtrlSetFocus txtInfo(0)
        Exit Function
    End If
    zlValidData = True
End Function

Public Function zlSaveCard(ByRef cllPro As Collection, ByRef lng������� As Long, ByRef strNO As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:���ݱ���
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-08-19 15:06:38
    '-----------------------------------------------------------------------------------------------------------
    Dim int���_IN As Integer
    Dim dbl���_IN As Double
    Dim str���㷽ʽ_IN As String
    Dim str�������_IN As String
    
    Dim str������_IN As String
    Dim str��������_IN As String
    Dim lng�������_IN As Long
    Dim strժҪ_IN As String
    Dim lngRow As Long
    Dim intCol As Integer
    
    zlSaveCard = False
    'txtNo = NextNo(31)
    
    str������_IN = UserInfo.����
    str��������_IN = Format(zlDatabase.Currentdate, "yyyy-mm-dd hh:mm:ss")
    
    strժҪ_IN = txtInfo(0).Text
    
    
    On Error GoTo errHandle:
    strNO = txtNo.Caption
    If mEditType = g���� Then
        lng�������_IN = zlDatabase.GetNextId("�����¼")
        lng������� = lng�������_IN
        strNO = NextNo(31)
        txtNo.Tag = strNO
    Else
        lng�������_IN = mlng�������
        lng������� = lng�������_IN
        gstrSQL = "zl_�����¼_DELETE('" & strNO & "')"
        AddArray cllPro, gstrSQL
    End If
     Dim blnData As Boolean
     blnData = False
    'ѭ������ÿ������
    With vsPayEdit
        For lngRow = 1 To .Rows - 1
            If Val(.TextMatrix(lngRow, .ColIndex("������"))) <> 0 _
                And Trim(.TextMatrix(lngRow, .ColIndex("���ʽ"))) <> "" Then
                blnData = True
                dbl���_IN = .TextMatrix(lngRow, .ColIndex("������"))
                '����27930 by lesfeng 2010-03-23
                If mint��� = 0 Then
                    str���㷽ʽ_IN = .TextMatrix(lngRow, .ColIndex("���ʽ"))
                Else
                    str���㷽ʽ_IN = ""
                End If
                str�������_IN = .TextMatrix(lngRow, .ColIndex("�������"))
                            
                'Zl_�������_Insert
                gstrSQL = " zl_�������_INSERT("
                '  No_In       IN �����¼.NO%TYPE,
                gstrSQL = gstrSQL & "'" & strNO & "',"
                '  ���_In     IN �����¼.���%TYPE,
                gstrSQL = gstrSQL & "" & lngRow & ","
                '  Ԥ����_In   IN �����¼.Ԥ����%TYPE := 0,
                gstrSQL = gstrSQL & "" & 0 & ","
                '  ��λid_In   IN �����¼.��λid%TYPE,
                gstrSQL = gstrSQL & "" & mlng��λID & ","
                '  ���_In     IN �����¼.���%TYPE,
                gstrSQL = gstrSQL & "" & dbl���_IN & ","
                '  ���㷽ʽ_In IN �����¼.���㷽ʽ%TYPE,
                gstrSQL = gstrSQL & "'" & str���㷽ʽ_IN & "',"
                '  �������_In IN �����¼.�������%TYPE := NULL,
                gstrSQL = gstrSQL & "'" & str�������_IN & "',"
                '  ������_In   IN �����¼.������%TYPE,
                gstrSQL = gstrSQL & "'" & str������_IN & "',"
                '  ��������_In IN �����¼.��������%TYPE,
                gstrSQL = gstrSQL & "to_date('" & str��������_IN & "','yyyy-mm-dd HH24:MI:SS'),"
                '  �������_In IN �����¼.�������%TYPE := NULL,
                gstrSQL = gstrSQL & "" & lng�������_IN & ","
                '  ժҪ_In     IN �����¼.ժҪ%TYPE := NULL
                gstrSQL = gstrSQL & "'" & strժҪ_IN & "',"
                '����27930 by lesfeng 2010-03-23
                '  �ܸ���־_In IN �����¼.�ܸ���־%TYPE := 0
                gstrSQL = gstrSQL & "" & mint��� & ")"
                AddArray cllPro, gstrSQL
            End If
        Next
    End With
    
    If blnData = False Then
        'Zl_�������_Insert
        gstrSQL = " zl_�������_INSERT("
        '  No_In       IN �����¼.NO%TYPE,
        gstrSQL = gstrSQL & "'" & strNO & "',"
        '  ���_In     IN �����¼.���%TYPE,
        gstrSQL = gstrSQL & "" & lngRow & ","
        '  Ԥ����_In   IN �����¼.Ԥ����%TYPE := 0,
        gstrSQL = gstrSQL & "" & 0 & ","
        '  ��λid_In   IN �����¼.��λid%TYPE,
        gstrSQL = gstrSQL & "" & mlng��λID & ","
        '  ���_In     IN �����¼.���%TYPE,
        gstrSQL = gstrSQL & "" & dbl���_IN & ","
        '  ���㷽ʽ_In IN �����¼.���㷽ʽ%TYPE,
        gstrSQL = gstrSQL & "'" & "" & "',"
        '  �������_In IN �����¼.�������%TYPE := NULL,
        gstrSQL = gstrSQL & "'" & "" & "',"
        '  ������_In   IN �����¼.������%TYPE,
        gstrSQL = gstrSQL & "'" & str������_IN & "',"
        '  ��������_In IN �����¼.��������%TYPE,
        gstrSQL = gstrSQL & "to_date('" & str��������_IN & "','yyyy-mm-dd HH24:MI:SS'),"
        '  �������_In IN �����¼.�������%TYPE := NULL,
        gstrSQL = gstrSQL & "" & lng�������_IN & ","
        '  ժҪ_In     IN �����¼.ժҪ%TYPE := NULL
        gstrSQL = gstrSQL & "'" & strժҪ_IN & "',"
        '����27930 by lesfeng 2010-03-23
        '  �ܸ���־_In IN �����¼.�ܸ���־%TYPE := 0
        gstrSQL = gstrSQL & "" & mint��� & ")"
        AddArray cllPro, gstrSQL
    End If
    zlSaveCard = True
    Exit Function
errHandle:
    If ErrCenter = 1 Then Exit Function
End Function

Public Function ClearData()
    '-----------------------------------------------------------------------------------------------------------
    '����:��ʼ������ؼ���Ĭ������
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-08-19 15:57:36
    '-----------------------------------------------------------------------------------------------------------
    txtInfo(0).Text = ""
    vsPayEdit.Clear 1
    vsPayEdit.Rows = 2
    vs��Ԥ��.Clear 1
    vs��Ԥ��.Rows = 2
    mlng��λID = 0
    Call zlLoadPrivder(0)
    mblnChange = False
End Function

Private Sub vs��Ԥ��_LostFocus()
    Call zl_VsGridLOSTFOCUS(vs��Ԥ��)
End Sub

Public Function zlCheck(ByRef cllPro As Collection) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:�������
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-08-19 15:06:38
    '-----------------------------------------------------------------------------------------------------------
    Dim lngRow As Long, str�������_IN As String
    'ѭ������ÿ������
    Err = 0: On Error GoTo ErrHand:
    With vsPayEdit
        For lngRow = 1 To .Rows - 1
            If Val(.Cell(flexcpData, lngRow, .ColIndex("���ʽ"))) <> 0 Then
                str�������_IN = .TextMatrix(lngRow, .ColIndex("�������"))
                If Trim(.Cell(flexcpData, lngRow, .ColIndex("�������"))) <> str�������_IN Then
                     If str�������_IN <> "" Then
                        If LenB(StrConv(str�������_IN, vbFromUnicode)) > 10 Then
                            ShowMsgbox "������볬��,���������5�����ֻ�10���ַ�!"
                            .Col = .ColIndex("�������"): .Row = lngRow: .TopRow = lngRow
                            Exit Function
                        End If
                        If InStr(1, str�������_IN, "'") <> 0 Then
                            ShowMsgbox "������벻�����뵥����!"
                            .Col = .ColIndex("�������"): .Row = lngRow: .TopRow = lngRow
                            Exit Function
                        End If
                    End If
                   
                    ' Zl_�����¼_�����update
                    gstrSQL = " Zl_�����¼_�����update("
                    '  Id_In       �����¼.ID%Type,
                    gstrSQL = gstrSQL & "" & Val(.Cell(flexcpData, lngRow, .ColIndex("���ʽ"))) & ","
                    '  �������_In In �����¼.�������%Type
                    gstrSQL = gstrSQL & "'" & str�������_IN & "')"
                    AddArray cllPro, gstrSQL
                End If
            End If
        Next
    End With
    zlCheck = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Exit Function
End Function
