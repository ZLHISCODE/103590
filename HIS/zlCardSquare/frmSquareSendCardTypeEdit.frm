VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmSquareSendCardTypeEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���ѿ����༭"
   ClientHeight    =   7170
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7800
   Icon            =   "frmSquareSendCardTypeEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   7800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   6015
      TabIndex        =   48
      Top             =   6600
      Width           =   1100
   End
   Begin VB.Frame fra 
      BorderStyle     =   0  'None
      Height          =   6375
      Index           =   1
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7815
      Begin VB.Frame fra 
         Height          =   855
         Index           =   11
         Left            =   75
         TabIndex        =   12
         Top             =   1710
         Width           =   5625
         Begin VB.CheckBox chkEdit 
            Caption         =   "���(&X)"
            Height          =   180
            Index           =   13
            Left            =   3870
            TabIndex        =   16
            Top             =   360
            Value           =   1  'Checked
            Width           =   975
         End
         Begin VB.CheckBox chkEdit 
            Caption         =   "סԺ(&Z)"
            Height          =   180
            Index           =   12
            Left            =   2460
            TabIndex        =   15
            Top             =   360
            Value           =   1  'Checked
            Width           =   975
         End
         Begin VB.CheckBox chkEdit 
            Caption         =   "����(&M)"
            Height          =   180
            Index           =   11
            Left            =   1140
            TabIndex        =   14
            Top             =   360
            Value           =   1  'Checked
            Width           =   975
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "ʹ�ó��ϣ�"
            Height          =   180
            Index           =   1
            Left            =   120
            TabIndex        =   13
            Top             =   360
            Width           =   900
         End
      End
      Begin VB.Frame fra 
         Caption         =   "ȱʡ�������"
         Height          =   1200
         Index           =   13
         Left            =   75
         TabIndex        =   26
         Top             =   2670
         Width           =   7665
         Begin VSFlex8Ctl.VSFlexGrid vsf������� 
            Height          =   945
            Left            =   30
            TabIndex        =   27
            Top             =   210
            Width           =   7575
            _cx             =   13361
            _cy             =   1667
            Appearance      =   2
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
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483643
            GridColorFixed  =   -2147483643
            TreeColor       =   -2147483643
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   0
            HighLight       =   2
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   0
            Cols            =   0
            FixedRows       =   0
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
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   3
         Left            =   3735
         MaxLength       =   3
         TabIndex        =   8
         Tag             =   "���ų���"
         Top             =   780
         Width           =   1935
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   2
         Left            =   3735
         MaxLength       =   100
         TabIndex        =   4
         Tag             =   "����"
         Top             =   225
         Width           =   1935
      End
      Begin VB.CheckBox chkEdit 
         Caption         =   "�������ѿ�(&S)"
         Height          =   180
         Index           =   0
         Left            =   3735
         TabIndex        =   11
         Top             =   1395
         Value           =   1  'Checked
         Width           =   1530
      End
      Begin VB.Frame fra 
         Height          =   1140
         Index           =   14
         Left            =   75
         TabIndex        =   28
         Top             =   3870
         Width           =   7665
         Begin VB.CheckBox chkEdit 
            Caption         =   "ˢ��"
            Height          =   180
            Index           =   8
            Left            =   1215
            TabIndex        =   30
            Top             =   270
            Value           =   1  'Checked
            Width           =   885
         End
         Begin VB.CheckBox chkEdit 
            Caption         =   "ɨ�迨"
            Height          =   180
            Index           =   9
            Left            =   2250
            TabIndex        =   31
            Top             =   270
            Width           =   960
         End
         Begin VB.OptionButton opt���̿��� 
            Caption         =   "ʹ���ַ������"
            Height          =   180
            Index           =   2
            Left            =   5040
            TabIndex        =   35
            Top             =   780
            Width           =   2055
         End
         Begin VB.OptionButton opt���̿��� 
            Caption         =   "ʹ�����������"
            Height          =   180
            Index           =   1
            Left            =   3090
            TabIndex        =   34
            Top             =   780
            Width           =   1800
         End
         Begin VB.OptionButton opt���̿��� 
            Caption         =   "��ֹʹ�������"
            Height          =   180
            Index           =   0
            Left            =   1170
            TabIndex        =   33
            Top             =   780
            Value           =   -1  'True
            Width           =   1650
         End
         Begin VB.Line Line1 
            BorderColor     =   &H8000000A&
            X1              =   0
            X2              =   7635
            Y1              =   600
            Y2              =   600
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "�������ʣ�"
            Height          =   180
            Index           =   5
            Left            =   90
            TabIndex        =   29
            Top             =   270
            Width           =   900
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "���̿��ƣ�"
            Height          =   180
            Index           =   8
            Left            =   90
            TabIndex        =   32
            Top             =   750
            Width           =   900
         End
      End
      Begin VB.Frame fra 
         Caption         =   "���ѿ�����"
         Height          =   2445
         Index           =   12
         Left            =   5790
         TabIndex        =   17
         Top             =   120
         Width           =   1950
         Begin VB.CheckBox chkEdit 
            Caption         =   "������(&6)"
            Height          =   180
            Index           =   7
            Left            =   150
            TabIndex        =   23
            Top             =   1625
            Width           =   1335
         End
         Begin VB.CheckBox chkEdit 
            Caption         =   "��������˿�(&8)"
            Height          =   180
            Index           =   3
            Left            =   150
            TabIndex        =   25
            Top             =   2160
            Width           =   1695
         End
         Begin VB.CheckBox chkEdit 
            Caption         =   "�ض�����(&5)"
            Height          =   180
            Index           =   2
            Left            =   150
            TabIndex        =   22
            Top             =   1360
            Width           =   1335
         End
         Begin VB.CheckBox chkEdit 
            Caption         =   "������(&7)"
            Enabled         =   0   'False
            Height          =   180
            Index           =   1
            Left            =   150
            TabIndex        =   24
            Top             =   1890
            Width           =   1320
         End
         Begin VB.CheckBox chkEdit 
            Caption         =   "�����˿�(&2)"
            Height          =   180
            Index           =   5
            Left            =   150
            TabIndex        =   19
            Top             =   565
            Value           =   1  'Checked
            Width           =   1335
         End
         Begin VB.CheckBox chkEdit 
            Caption         =   "��������(&4)"
            Height          =   180
            Index           =   4
            Left            =   150
            TabIndex        =   20
            Top             =   830
            Width           =   1335
         End
         Begin VB.CheckBox chkEdit 
            Caption         =   "��������(&1)"
            Height          =   180
            Index           =   6
            Left            =   150
            TabIndex        =   18
            Top             =   300
            Value           =   1  'Checked
            Width           =   1335
         End
         Begin VB.CheckBox chkEdit 
            Caption         =   "�ϸ����(&3)"
            Height          =   180
            Index           =   10
            Left            =   150
            TabIndex        =   21
            Top             =   1095
            Width           =   1320
         End
      End
      Begin VB.Frame fra 
         Caption         =   "������������"
         Height          =   1290
         Index           =   15
         Left            =   75
         TabIndex        =   36
         Top             =   5085
         Width           =   7665
         Begin VB.Frame fra 
            BorderStyle     =   0  'None
            Height          =   255
            Index           =   16
            Left            =   1335
            TabIndex        =   44
            Top             =   863
            Width           =   3915
            Begin VB.OptionButton optRule 
               Caption         =   "�����ַ�������"
               Height          =   180
               Index           =   0
               Left            =   0
               TabIndex        =   45
               Top             =   30
               Value           =   -1  'True
               Width           =   1560
            End
            Begin VB.OptionButton optRule 
               Caption         =   "�����ַ�ֻ��Ϊ����"
               Height          =   180
               Index           =   1
               Left            =   1620
               TabIndex        =   46
               Top             =   30
               Width           =   2070
            End
         End
         Begin VB.TextBox txtEdit 
            Enabled         =   0   'False
            Height          =   270
            Index           =   4
            Left            =   5565
            MaxLength       =   2
            TabIndex        =   42
            Text            =   "0"
            Top             =   375
            Width           =   300
         End
         Begin VB.OptionButton optPassInput 
            Caption         =   "�̶�����10λ"
            Height          =   210
            Index           =   1
            Left            =   2955
            TabIndex        =   40
            Top             =   405
            Width           =   1545
         End
         Begin VB.TextBox txtEdit 
            Height          =   270
            Index           =   5
            Left            =   435
            MaxLength       =   2
            TabIndex        =   38
            Text            =   "10"
            Top             =   330
            Width           =   300
         End
         Begin VB.OptionButton optPassInput 
            Caption         =   "���벻�̶�"
            Height          =   210
            Index           =   0
            Left            =   1335
            TabIndex        =   39
            Top             =   390
            Width           =   1380
         End
         Begin VB.OptionButton optPassInput 
            Caption         =   "��������    λ��������"
            Height          =   210
            Index           =   2
            Left            =   4545
            TabIndex        =   41
            Top             =   405
            Width           =   2295
         End
         Begin VB.Line Line2 
            BorderColor     =   &H8000000A&
            X1              =   0
            X2              =   7635
            Y1              =   750
            Y2              =   750
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "����    λ��"
            Height          =   180
            Index           =   4
            Left            =   60
            TabIndex        =   37
            Top             =   375
            Width           =   1080
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "�������"
            Height          =   180
            Index           =   9
            Left            =   240
            TabIndex        =   43
            Top             =   900
            Width           =   900
         End
      End
      Begin VB.ComboBox cbo���㷽ʽ 
         Height          =   300
         Left            =   1065
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1335
         Width           =   1545
      End
      Begin VB.TextBox txtEdit 
         Height          =   315
         Index           =   1
         Left            =   1065
         MaxLength       =   2
         TabIndex        =   6
         Tag             =   "ǰ׺�ı�"
         Top             =   765
         Width           =   1545
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   0
         Left            =   1065
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "����"
         Top             =   225
         Width           =   1545
      End
      Begin VB.Label lblEdit 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "���㷽ʽ(&J)"
         Height          =   180
         Index           =   7
         Left            =   60
         TabIndex        =   9
         Top             =   1395
         Width           =   990
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "���ų���(&L)"
         Height          =   180
         Index           =   6
         Left            =   2715
         TabIndex        =   7
         Top             =   840
         Width           =   990
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "ǰ׺�ı�(&T)"
         Height          =   180
         Index           =   3
         Left            =   60
         TabIndex        =   5
         Top             =   840
         Width           =   990
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "����(&N)"
         Height          =   180
         Index           =   2
         Left            =   3075
         TabIndex        =   3
         Top             =   285
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "����(&B)"
         Height          =   180
         Index           =   0
         Left            =   420
         TabIndex        =   1
         Top             =   285
         Width           =   630
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   4875
      TabIndex        =   47
      Top             =   6600
      Width           =   1100
   End
End
Attribute VB_Name = "frmSquareSendCardTypeEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'-----------------------------------------------------------------------------------------
'��ڲ���
Public Enum gSendCardEdit
    Card_���� = 0
    Card_�޸� = 1
    Card_ɾ�� = 2
    Card_ͣ�� = 3
    Card_���� = 4
    Card_�鿴 = 5
End Enum
Private mlngModule As Long
Private mstrPrivs As String
Private mEditType As gSendCardEdit
Private mlngCardTypeID As Long
'-----------------------------------------------------------------------------------------
Private mintSucces As Integer
Private mblnFirst As Boolean
Private Enum mtxtIdx
     idx_��� = 0
     idx_���� = 2
     idx_ǰ׺�ı� = 1
     idx_���ų��� = 3
     idx_���볤�� = 5
     idx_����λ�� = 4
End Enum

Private Enum mchkIdx
    idx_���� = 0
    idx_���� = 6
    idx_ȫ�� = 5
    idx_���� = 4
    idx_�ϸ���� = 10
    idx_�ض����� = 2
    idx_���� = 7
    idx_���� = 1
    idx_����˿� = 3
    
    idx_ˢ�� = 8
    idx_ɨ�迨 = 9
    
    idx_���� = 11
    idx_סԺ = 12
    idx_��� = 13
End Enum

Private Type Ty_CardType
    lng���ų��� As Long
    bln�̶� As Boolean
    bln�ѷ��� As Boolean '�Ƿ��Ѿ�������
    strǰ׺�ı� As String
End Type
Private mCardType As Ty_CardType
Private mblnNotClick As Boolean
Private mblnChange As Boolean

Public Function zlEditSendCard(ByVal frmMain As Object, ByVal lngModule As Long, ByVal strPrivs As String, _
    ByVal EditType As gSendCardEdit, Optional lngCardTypeID As Long = 0) As Boolean
    '����:ҽ�ƿ����༭
    '���:EditType-�༭����
    '        lngCardTypeID-����ʱΪ0
    '����:
    '����:ֻҪ�ɹ�һ��,����true,���򷵻�Flase
    mEditType = EditType: mlngModule = lngModule: mstrPrivs = strPrivs
    mlngCardTypeID = lngCardTypeID
    
    On Error Resume Next
    mintSucces = 0
    Me.Show 1, frmMain
    zlEditSendCard = mintSucces > 0
End Function

Private Sub Form_Load()
    Dim ty_Temp As Ty_CardType
    
    mblnFirst = True
    mCardType = ty_Temp '�Զ���Type��ʼ��
    
    If InitData() = False Then Unload Me: Exit Sub
    If LoadCardData() = False Then Unload Me: Exit Sub
    Call SetCtrlEnable
    
    If mEditType = dt_�鿴 Then
        cmdOK.Visible = False
    End If
    mblnChange = False
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    
    If mEditType = Card_���� Then
        zlControl.ControlSetFocus txtEdit(mtxtIdx.idx_����)
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("|,'��~;", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnFirst Or mblnChange = False Then Exit Sub
    If mEditType = Card_���� Or mEditType = gEd_�޸� Then
        If MsgBox("�����������˳��Ļ������е��޸Ķ�������Ч��" & vbCrLf & "�Ƿ�ȷ���˳���", _
            vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Cancel = 1
    End If
End Sub
 
 Private Function InitData() As Boolean
    '����:��ʼ������
    '����:��ʼ���ɹ�������true,���򷵻�False
    Dim rsTemp As ADODB.Recordset, strSQL As String
    
    On Error GoTo errHandle
    If Not (mEditType = Card_���� Or mEditType = Card_�޸�) Then InitData = True: Exit Function
    
    If mEditType = Card_���� Then
        txtEdit(mtxtIdx.idx_���).Text = zlDatabase.GetMax("���ѿ����Ŀ¼", "���", txtEdit(mtxtIdx.idx_���).MaxLength)
    End If
    
    strSQL = "Select ���� From ���㷽ʽ Where ���� = 8 And Nvl(Ӧ����, 0) = 0"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    With cbo���㷽ʽ
        .Clear
        Do While Not rsTemp.EOF
            If NVL(rsTemp!����) <> "" Then .AddItem NVL(rsTemp!����)
            rsTemp.MoveNext
        Loop
    End With
    
    Set rsTemp = zlGet�շ����()
    With vsf�������
        .Clear
        Do While Not rsTemp.EOF
           ZL_vsGrid_AddCell vsf�������, NVL(rsTemp!����) & "-" & NVL(rsTemp!����), NVL(rsTemp!����), True
           rsTemp.MoveNext
        Loop
        ZL_vsGrid_AutoSetGridRowAndCol vsf�������
    End With
    
    txtEdit(mtxtIdx.idx_���).MaxLength = 6
    txtEdit(mtxtIdx.idx_����).MaxLength = 50
    txtEdit(mtxtIdx.idx_ǰ׺�ı�).MaxLength = 2
    txtEdit(mtxtIdx.idx_���ų���).MaxLength = 2
    InitData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
 End Function
 
Private Function LoadCardData() As Boolean
    '����:���ؿ�Ƭ����
    '����:���سɹ�������true�����򷵻�False
    Dim rsTemp As ADODB.Recordset, strSQL As String
    Dim rs����� As ADODB.Recordset
    Dim strValue As String, intIndx As Integer
    Dim i As Long, j As Long
    
    On Error GoTo errHandle
    If mEditType = Card_���� Then LoadCardData = True: Exit Function
    
    Set rs����� = zlGet���ѿ��ӿ�(, True)
    rs�����.Filter = "���=" & mlngCardTypeID
    If rs�����.EOF Then
        MsgBox "δ�ҵ����ѿ������Ϣ�������Ѿ�������ɾ����", vbInformation, gstrSysName
        Exit Function
    End If
    
    txtEdit(mtxtIdx.idx_���).Text = NVL(rs�����!���)
    txtEdit(mtxtIdx.idx_����).Text = NVL(rs�����!����)
    txtEdit(mtxtIdx.idx_ǰ׺�ı�).Text = NVL(rs�����!ǰ׺�ı�)
    txtEdit(mtxtIdx.idx_���ų���).Text = IIf(Val(NVL(rs�����!���ų���)) = 0, 1, Val(NVL(rs�����!���ų���)))
    
    cbo.SeekIndex cbo���㷽ʽ, NVL(rs�����!���㷽ʽ)
    If cbo���㷽ʽ.ListIndex < 0 Then
        cbo���㷽ʽ.AddItem NVL(rs�����!���㷽ʽ)
        cbo���㷽ʽ.ListIndex = cbo���㷽ʽ.NewIndex
    End If
    chkEdit(mchkIdx.idx_����).value = IIf(Val(NVL(rs�����!����)) = 1, 1, 0)
    
    chkEdit(mchkIdx.idx_����).value = IIf(Val(NVL(rs�����!�Ƿ�����)) = 1, 1, 0)
    chkEdit(mchkIdx.idx_ȫ��).value = IIf(Val(NVL(rs�����!�Ƿ�ȫ��)) = 1, 0, 1)
    chkEdit(mchkIdx.idx_����).value = IIf(Val(NVL(rs�����!�Ƿ�����)) = 1, 1, 0)
    chkEdit(mchkIdx.idx_�ϸ����).value = IIf(Val(NVL(rs�����!�Ƿ��ϸ����)) = 1, 1, 0)
    chkEdit(mchkIdx.idx_�ض�����).value = IIf(Val(NVL(rs�����!�Ƿ��ض�����)) = 1, 1, 0)
    chkEdit(mchkIdx.idx_����).value = IIf(Val(NVL(rs�����!�Ƿ�������)) = 1, 1, 0)
    chkEdit(mchkIdx.idx_����).value = IIf(Val(NVL(rs�����!�Ƿ�������)) = 1, 1, 0)
    chkEdit(mchkIdx.idx_����˿�).value = IIf(Val(NVL(rs�����!�Ƿ���������˿�)) = 1, 1, 0)
    
    strValue = NVL(rs�����!Ӧ�ó���, "000")
    chkEdit(mchkIdx.idx_����).value = IIf(Val(Mid(strValue, 1, 1)) = 0, vbChecked, vbUnchecked)
    chkEdit(mchkIdx.idx_סԺ).value = IIf(Val(Mid(strValue, 2, 1)) = 0, vbChecked, vbUnchecked)
    chkEdit(mchkIdx.idx_���).value = IIf(Val(Mid(strValue, 3, 1)) = 0, vbChecked, vbUnchecked)
    
    With vsf�������
        strValue = NVL(rs�����!�������)
        .Tag = strValue
        For i = 0 To .Rows - 1
            For j = 0 To .Cols - 1
                If .TextMatrix(i, j) <> "" Then
                    If InStr("," & strValue & ",", "," & .Cell(flexcpData, i, j) & ",") > 0 Then
                        .Cell(flexcpChecked, i, j) = 1
                    Else
                        .Cell(flexcpChecked, i, j) = 2
                    End If
                End If
            Next
        Next
    End With
    
    strValue = NVL(rs�����!��������, "10")
    chkEdit(mchkIdx.idx_ˢ��).value = Val(Mid(strValue, 1, 1))
    chkEdit(mchkIdx.idx_ɨ�迨).value = Val(Mid(strValue, 2, 1))
    
    intIndx = Val(NVL(rs�����!���̿��Ʒ�ʽ))
    If intIndx < 0 Or intIndx > 2 Then intIndx = 0
    opt���̿���(intIndx).value = True
    
    txtEdit(mtxtIdx.idx_���볤��).Text = Val(NVL(rs�����!���볤��))
    Select Case Val(NVL(rs�����!���볤������))
    Case 0
        optPassInput(0).value = True
    Case 1
        optPassInput(1).value = True
    Case Else '����
        optPassInput(2).value = True
        txtEdit(mtxtIdx.idx_����λ��).Text = Abs(Val(NVL(rs�����!���볤������)))
    End Select
    intIndx = Val(NVL(rs�����!�������))
    If intIndx < 0 Or intIndx > 1 Then intIndx = 0
    optRule(intIndx).value = True
    
    With mCardType
        .lng���ų��� = Val(NVL(rs�����!���ų���))
        .bln�̶� = Val(NVL(rs�����!ϵͳ)) = 1
        .strǰ׺�ı� = NVL(rs�����!ǰ׺�ı�)
        
        strSQL = "Select 1 From ���ѿ���Ϣ Where �ӿڱ��=[1] And Rownum < 2"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngCardTypeID)
        .bln�ѷ��� = Not rsTemp.EOF
    End With
    
    LoadCardData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub SetCtrlEnable()
    '����:���ÿؼ��ı༭����
    '����:���˺�
    Dim i As Long, blnEdit As Boolean
    
    On Error GoTo ErrHandler
    blnEdit = (mEditType = Card_���� Or mEditType = Card_�޸�)
    For i = 0 To txtEdit.UBound
        Select Case i
        Case mtxtIdx.idx_���
            txtEdit(i).Enabled = mEditType = Card_����
        Case mtxtIdx.idx_����
            txtEdit(i).Enabled = blnEdit And Not mCardType.bln�̶�
        Case mtxtIdx.idx_����λ��
            txtEdit(i).Enabled = False
        Case Else
            txtEdit(i).Enabled = blnEdit
        End Select
    Next
    
    For i = 0 To chkEdit.UBound
        chkEdit(i).Enabled = blnEdit
    Next
    Call chkEdit_Click(mchkIdx.idx_�ض�����)
    
    cbo���㷽ʽ.Enabled = blnEdit
    vsf�������.Enabled = blnEdit
    vsf�������.Editable = IIf(blnEdit, flexEDKbdMouse, flexEDNone)
    
    optPassInput(0).Enabled = blnEdit
    optPassInput(1).Enabled = blnEdit
    optPassInput(2).Enabled = blnEdit
    
    optRule(0).Enabled = blnEdit
    optRule(1).Enabled = blnEdit
    
    Call SetEnabledBackColor(Me)
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cbo���㷽ʽ_Change()
    mblnChange = True
End Sub

Private Sub cbo���㷽ʽ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub chkEdit_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub lvw�������_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    mblnChange = True
End Sub

Private Sub lvw�������_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optRule_Click(Index As Integer)
    mblnChange = True
End Sub

Private Sub opt���̿���_Click(Index As Integer)
    mblnChange = True
End Sub

Private Sub txtEdit_Change(Index As Integer)
    mblnChange = True
    If Index = mtxtIdx.idx_���볤�� Then
        optPassInput(1).Caption = "�̶�����" & Val(txtEdit(Index).Text) & "λ"
    End If
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtEdit(Index)
    If Index = mtxtIdx.idx_���� Then
        zlCommFun.OpenIme True
    End If
End Sub

Private Sub txtEdit_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then txtEdit(Index).Text = ""
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = mtxtIdx.idx_���ų��� Or Index = mtxtIdx.idx_��� _
        Or Index = mtxtIdx.idx_����λ�� Or Index = mtxtIdx.idx_���볤�� Then
        zlControl.TxtCheckKeyPress txtEdit(Index), KeyAscii, m����ʽ
    ElseIf Index = mtxtIdx.idx_ǰ׺�ı� Then
        If zlStr.IsCharChinese(Chr(KeyAscii)) Then KeyAscii = 0
    End If
End Sub

Private Sub txtEdit_LostFocus(Index As Integer)
    If Index = mtxtIdx.idx_���� Then
        zlCommFun.OpenIme False
    ElseIf Index = mtxtIdx.idx_ǰ׺�ı� Then
        txtEdit(Index).Text = UCase(txtEdit(Index).Text)
    End If
End Sub

Private Sub txtEdit_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        glngTXTProc = GetWindowLong(txtEdit(Index).hWnd, GWL_WNDPROC)
        Call SetWindowLong(txtEdit(Index).hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txtEdit_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        Call SetWindowLong(txtEdit(Index).hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub chkEdit_Click(Index As Integer)
    Dim blnEnabled As Boolean
    
    mblnChange = True
    
    '���ٱ���һ��
    Select Case Index
    Case mchkIdx.idx_����, mchkIdx.idx_סԺ, mchkIdx.idx_���
        CheckCheckboxValue Array(chkEdit(mchkIdx.idx_����), chkEdit(mchkIdx.idx_סԺ), chkEdit(mchkIdx.idx_���))
    Case mchkIdx.idx_ˢ��, mchkIdx.idx_ɨ�迨
        'CheckCheckboxValue Array(chkEdit(mchkIdx.idx_ˢ��), chkEdit(mchkIdx.idx_ɨ�迨))
        If chkEdit(Index).value = vbUnchecked Then
            If Index = mchkIdx.idx_ˢ�� Then
                chkEdit(mchkIdx.idx_ɨ�迨).value = vbChecked
            Else
                chkEdit(mchkIdx.idx_ˢ��).value = vbChecked
            End If
        End If
    Case mchkIdx.idx_�ض�����
        blnEnabled = chkEdit(mchkIdx.idx_�ض�����).value
        chkEdit(mchkIdx.idx_����).Enabled = blnEnabled
    End Select
End Sub

Private Sub CheckCheckboxValue(ByVal varCheckBox As Variant)
    '����һ��CkeckBox�����뱣֤����һ���ǹ�ѡ��
    Dim i As Integer, blnChecked As Boolean
    
    On Error GoTo ErrHandler
    For i = 0 To UBound(varCheckBox)
        If varCheckBox(i).value Then
            blnChecked = True: Exit For
        End If
    Next
    
    If blnChecked = False Then
        varCheckBox(0).value = vbChecked
    End If
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub optPassInput_Click(Index As Integer)
    mblnChange = True
    
    txtEdit(mtxtIdx.idx_����λ��).Enabled = optPassInput(2).value
    zl_SetCtlBackColor txtEdit(mtxtIdx.idx_����λ��), Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    On Error GoTo ErrHandler
    If isValied() = False Then Exit Sub
    If SaveData() = False Then Exit Sub
    
    mintSucces = mintSucces + 1
    mblnChange = False
    Unload Me
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
 
Private Function isValied() As Boolean
    '����:������ݵ���Ч��
    '����:������Ч������true,���򷵻�False
    Dim i As Long, strTemp As String
    Dim rsTemp As ADODB.Recordset, strSQL As String
    
    On Error GoTo errHandle
    If zlControl.FormCheckInput(Me) = False Then Exit Function
    If zlControl.TxtCheckInput(txtEdit(mtxtIdx.idx_���), "���") = False Then Exit Function
    If zlControl.TxtCheckInput(txtEdit(mtxtIdx.idx_����), "����") = False Then Exit Function
    If zlControl.TxtCheckInput(txtEdit(mtxtIdx.idx_ǰ׺�ı�), "ǰ׺�ı�", , True) = False Then Exit Function
    If zlControl.TxtCheckInput(txtEdit(mtxtIdx.idx_���ų���), "���ų���") = False Then Exit Function
    
    If zlStr.IsCharChinese(txtEdit(mtxtIdx.idx_ǰ׺�ı�)) Then
        ShowMsgbox "ǰ׺�ı����ܰ������֣�"
        zlControl.ControlSetFocus txtEdit(mtxtIdx.idx_ǰ׺�ı�)
        Exit Function
    End If
    
    If Val(txtEdit(mtxtIdx.idx_���ų���).Text) < 1 Then
        ShowMsgbox "���ų��ȱ�����ڵ���1λ��"
        zlControl.ControlSetFocus txtEdit(mtxtIdx.idx_���ų���)
        Exit Function
    End If
    
    If zlCommFun.ActualLen(Trim(txtEdit(idx_ǰ׺�ı�))) + Val(txtEdit(mtxtIdx.idx_���ų���).Text) > 20 Then
        ShowMsgbox "���ŵ���󳤶�(ǰ׺+���ų���)���ܴ���20λ�����飡"
        zlControl.ControlSetFocus txtEdit(mtxtIdx.idx_���ų���)
        Exit Function
    End If
    
    If mCardType.bln�ѷ��� Then
        If Val(txtEdit(idx_���ų���).Text) + Len(Trim(txtEdit(idx_ǰ׺�ı�))) < mCardType.lng���ų��� + Len(NVL(mCardType.strǰ׺�ı�)) Then
            ShowMsgbox "���ڷ����˷�����Ϣ,�������ѿ�ǰ׺�ı������ų��Ȳ��ܼ�С,���飡"
            zlControl.ControlSetFocus txtEdit(mtxtIdx.idx_���ų���)
            Exit Function
        End If
    End If
    
    If cbo���㷽ʽ.ListIndex < 0 Then
        ShowMsgbox "���㷽ʽ����ѡ��"
        zlControl.ControlSetFocus cbo���㷽ʽ
        Exit Function
    End If
    
    strSQL = _
        "Select ���� From ҽ�ƿ���� Where ���㷽ʽ = [2]" & vbNewLine & _
        "Union All" & vbNewLine & _
        "Select ���� From ���ѿ����Ŀ¼ Where ��� <> [1] And ���㷽ʽ = [2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngCardTypeID, cbo���㷽ʽ.Text)
    If Not rsTemp.EOF Then
        ShowMsgbox "���㷽ʽ��" & cbo���㷽ʽ.Text & "���ѱ�" & NVL(rsTemp!����) & "ʹ�ã�" & _
                   "�ظ�ʹ�û���ɲ����������ң�������ѡ��һ�ֽ��㷽ʽ��"
        zlControl.ControlSetFocus cbo���㷽ʽ
        Exit Function
    End If
    
    If Val(txtEdit(mtxtIdx.idx_���볤��).Text) = 0 Then
        ShowMsgbox "���볤�Ȳ�������Ϊ�㣡"
        zlControl.ControlSetFocus txtEdit(mtxtIdx.idx_���볤��)
        Exit Function
    End If
    If Val(txtEdit(mtxtIdx.idx_���볤��).Text) > 50 Then
        ShowMsgbox "���볤�Ȳ��ܴ���50λ��"
        zlControl.ControlSetFocus txtEdit(mtxtIdx.idx_���볤��)
        Exit Function
    End If
    If optPassInput(2).value Then
        If Val(txtEdit(mtxtIdx.idx_���볤��).Text) < Val(txtEdit(mtxtIdx.idx_����λ��).Text) Then
            ShowMsgbox "������������볤�Ȳ��ܴ����ܵ����볤��(" & Val(txtEdit(mtxtIdx.idx_���볤��).Text) & ")λ��"
            zlControl.ControlSetFocus txtEdit(mtxtIdx.idx_����λ��)
            Exit Function
        End If
    End If
    isValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function SaveData() As Boolean
    '����:��������
    '����:����ɹ�,����true,���򷵻�False
    Dim strSQL As String
    Dim strValue As String

    On Error GoTo errHandle
    'Zl_���ѿ����Ŀ¼_Update
    strSQL = "Zl_���ѿ����Ŀ¼_Update("
    '  ����_In         In ���ѿ����Ŀ¼.���%Type,
    strSQL = strSQL & "'" & Trim(txtEdit(mtxtIdx.idx_���).Text) & "',"
    '  ����_In         In ���ѿ����Ŀ¼.����%Type,
    strSQL = strSQL & "'" & Trim(txtEdit(mtxtIdx.idx_����).Text) & "',"
    '  ���㷽ʽ_In     In ���ѿ����Ŀ¼.���㷽ʽ%Type,
    strSQL = strSQL & "'" & cbo���㷽ʽ.Text & "',"
    '  ǰ׺�ı�_In     In ���ѿ����Ŀ¼.ǰ׺�ı�%Type,
    strSQL = strSQL & "'" & Trim(txtEdit(mtxtIdx.idx_ǰ׺�ı�).Text) & "',"
    '  ���ų���_In     In ���ѿ����Ŀ¼.���ų���%Type,
    strSQL = strSQL & "" & Val(txtEdit(mtxtIdx.idx_���ų���).Text) & ","
    '  ��������_In     In ���ѿ����Ŀ¼.�Ƿ�����%Type,
    strSQL = strSQL & "" & IIf(chkEdit(idx_����).value = vbChecked, "1", "0") & ","
    '  �Ƿ�����_In     In ���ѿ����Ŀ¼.�Ƿ�ȫ��%Type,
    strSQL = strSQL & "" & IIf(chkEdit(idx_����).value = vbChecked, "1", "0") & ","
    '  �Ƿ�ȫ��_In     In ���ѿ����Ŀ¼.�Ƿ�ȫ��%Type,
    strSQL = strSQL & "" & IIf(chkEdit(idx_ȫ��).value = vbChecked, "0", "1") & ","
    '  ����_In         In ���ѿ����Ŀ¼.����%Type,
    strSQL = strSQL & "" & IIf(chkEdit(idx_����).value = vbChecked, "1", "0") & ","
    '  ���볤��_In     In ���ѿ����Ŀ¼.���볤��%Type,
    strSQL = strSQL & "" & Val(txtEdit(mtxtIdx.idx_���볤��).Text) & ","
    '  ���볤������_In In ���ѿ����Ŀ¼.���볤������%Type,
    If optPassInput(0).value Then
        strSQL = strSQL & "" & 0 & ","
    ElseIf optPassInput(1).value Then
        strSQL = strSQL & "" & 1 & ","
    Else
        strSQL = strSQL & "" & -1 * Val(txtEdit(mtxtIdx.idx_����λ��).Text) & ","
    End If
    '  �������_In     In ���ѿ����Ŀ¼.�������%Type,
    If optRule(0).value Then
        strSQL = strSQL & "" & 0 & ","
    Else
        strSQL = strSQL & "" & 1 & ","
    End If
    '  ������ʽ_In     In Integer := 0
    strSQL = strSQL & "" & IIf(mEditType = Card_����, 0, 1) & ","
    '  ��������_In         In ���ѿ����Ŀ¼.��������%Type,
    strValue = IIf(chkEdit(mchkIdx.idx_ˢ��).value = vbChecked, "1", "0")
    strValue = strValue & IIf(chkEdit(mchkIdx.idx_ɨ�迨).value = 1, "1", "0")
    strSQL = strSQL & "'" & strValue & "',"
    '  ���̿��Ʒ�ʽ_In     In ���ѿ����Ŀ¼.���̿��Ʒ�ʽ%Type,
    If opt���̿���(0).value Then
        strSQL = strSQL & "" & 0 & ","
    ElseIf opt���̿���(1).value Then
        strSQL = strSQL & "" & 1 & ","
    ElseIf opt���̿���(2).value Then
        strSQL = strSQL & "" & 2 & ","
    End If
    '  �������_In         In ���ѿ����Ŀ¼.�������%Type,
    strSQL = strSQL & "'" & Get�������() & "',"
    '  �Ƿ��ϸ����_In     In ���ѿ����Ŀ¼.�Ƿ��ϸ����%Type,
    strSQL = strSQL & "" & IIf(chkEdit(idx_�ϸ����).value = vbChecked, "1", "0") & ","
    '  �Ƿ��ض�����_In     In ���ѿ����Ŀ¼.�Ƿ��ض�����%Type,
    strSQL = strSQL & "" & IIf(chkEdit(idx_�ض�����).value = vbChecked, "1", "0") & ","
    '  �Ƿ�������_In     In ���ѿ����Ŀ¼.�Ƿ�������%Type,
    strSQL = strSQL & "" & IIf(chkEdit(idx_����).value = vbChecked, "1", "0") & ","
    '  �Ƿ�������_In     In ���ѿ����Ŀ¼.�Ƿ�������%Type,
    strSQL = strSQL & "" & IIf(chkEdit(idx_�ض�����).value = vbChecked And chkEdit(idx_����).value = vbChecked, "1", "0") & ","
    '  �Ƿ���������˿�_In In ���ѿ����Ŀ¼.�Ƿ���������˿�%Type,
    strSQL = strSQL & "" & IIf(chkEdit(idx_����˿�).value = vbChecked, "1", "0") & ","
    '  Ӧ�ó���_In         In ���ѿ����Ŀ¼.Ӧ�ó���%Type
    strValue = IIf(chkEdit(mchkIdx.idx_����).value = vbChecked, "0", "1")
    strValue = strValue & IIf(chkEdit(mchkIdx.idx_סԺ).value = vbChecked, "0", "1")
    strValue = strValue & IIf(chkEdit(mchkIdx.idx_���).value = vbChecked, "0", "1")
    strSQL = strSQL & "'" & strValue & "')"
    
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    SaveData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Get�������() As String
    '��ȡ�������
    Dim strType As String, i As Long, j As Long
    
    On Error GoTo ErrHandler
    With vsf�������
         For i = 0 To .Rows - 1
            For j = 0 To .Cols - 1
                If Abs(Val(.Cell(flexcpChecked, i, j))) = 1 Then
                    strType = strType & "," & .Cell(flexcpData, i, j)
                End If
            Next
         Next
         If strType <> "" Then strType = Mid(strType, 2)
    End With
    Get������� = strType
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub vsf�������_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    mblnChange = True
End Sub

Private Sub vsf�������_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Row < 0 Or Col < 0 Then Exit Sub
    If vsf�������.TextMatrix(Row, Col) = "" Then Cancel = True
End Sub

Private Sub vsf�������_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    If NewRow < 0 Or NewCol < 0 Then Exit Sub
    If vsf�������.TextMatrix(NewRow, NewCol) = "" Then Cancel = True
End Sub

Private Sub vsf�������_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub
