VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmBillInDamnifyEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ʊ�ݱ���"
   ClientHeight    =   7050
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9600
   Icon            =   "frmBillInDamnifyEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7050
   ScaleWidth      =   9600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame Frame1 
      Height          =   7350
      Left            =   7980
      TabIndex        =   28
      Top             =   -60
      Width           =   30
   End
   Begin VB.Frame fra 
      Caption         =   "���α������"
      Height          =   5145
      Left            =   165
      TabIndex        =   11
      Top             =   1695
      Width           =   7680
      Begin VB.TextBox txtEdit 
         Enabled         =   0   'False
         Height          =   360
         Index           =   10
         Left            =   1035
         MaxLength       =   20
         TabIndex        =   22
         Top             =   4500
         Width           =   2145
      End
      Begin VB.TextBox txtEdit 
         Enabled         =   0   'False
         Height          =   360
         Index           =   9
         Left            =   5415
         MaxLength       =   20
         TabIndex        =   24
         Top             =   4500
         Width           =   2145
      End
      Begin VB.TextBox txtEdit 
         Height          =   360
         Index           =   8
         Left            =   1050
         MaxLength       =   200
         TabIndex        =   20
         Top             =   4065
         Width           =   6525
      End
      Begin VB.CommandButton cmdRemove 
         Cancel          =   -1  'True
         Caption         =   "ɾ��(&R)"
         Height          =   375
         Left            =   6720
         TabIndex        =   17
         Top             =   285
         Width           =   825
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "����(&N)"
         Height          =   375
         Left            =   5835
         TabIndex        =   16
         Top             =   285
         Width           =   840
      End
      Begin VB.TextBox txtEdit 
         Height          =   360
         Index           =   7
         Left            =   3660
         MaxLength       =   20
         TabIndex        =   15
         Top             =   300
         Width           =   2145
      End
      Begin VSFlex8Ctl.VSFlexGrid vsMemo 
         Height          =   3210
         Left            =   105
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   765
         Width           =   7470
         _cx             =   13176
         _cy             =   5662
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
         BackColorBkg    =   -2147483634
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483648
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   350
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmBillInDamnifyEdit.frx":058A
         ScrollTrack     =   0   'False
         ScrollBars      =   2
         ScrollTips      =   0   'False
         MergeCells      =   1
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
      Begin VB.TextBox txtEdit 
         Height          =   360
         Index           =   6
         Left            =   1185
         MaxLength       =   20
         TabIndex        =   13
         Top             =   285
         Width           =   2145
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "������"
         Height          =   180
         Left            =   465
         TabIndex        =   21
         Top             =   4590
         Width           =   540
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����ʱ��"
         Height          =   180
         Index           =   2
         Left            =   4590
         TabIndex        =   23
         Top             =   4590
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "����ԭ��(&M)"
         Height          =   180
         Left            =   60
         TabIndex        =   19
         Top             =   4155
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         Height          =   240
         Index           =   1
         Left            =   3405
         TabIndex        =   14
         Top             =   390
         Width           =   240
      End
      Begin VB.Label lblE 
         AutoSize        =   -1  'True
         Caption         =   "����Ʊ��(&B)"
         Height          =   180
         Left            =   195
         TabIndex        =   12
         Top             =   390
         Width           =   990
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   420
      Left            =   8145
      TabIndex        =   27
      Top             =   6450
      Width           =   1200
   End
   Begin VB.Frame fraUse 
      Caption         =   "��������Ϣ"
      Height          =   1380
      Left            =   150
      TabIndex        =   0
      Top             =   210
      Width           =   7710
      Begin VB.TextBox txtEdit 
         Enabled         =   0   'False
         Height          =   360
         Index           =   1
         Left            =   3180
         MaxLength       =   2
         TabIndex        =   4
         Top             =   375
         Width           =   375
      End
      Begin VB.TextBox txtEdit 
         Enabled         =   0   'False
         Height          =   360
         Index           =   5
         Left            =   1125
         MaxLength       =   20
         TabIndex        =   10
         Top             =   870
         Width           =   6300
      End
      Begin VB.TextBox txtEdit 
         Enabled         =   0   'False
         Height          =   360
         Index           =   2
         Left            =   3570
         MaxLength       =   20
         TabIndex        =   5
         Top             =   375
         Width           =   1530
      End
      Begin VB.TextBox txtEdit 
         Enabled         =   0   'False
         Height          =   360
         Index           =   3
         Left            =   5520
         MaxLength       =   2
         TabIndex        =   7
         Top             =   375
         Width           =   375
      End
      Begin VB.TextBox txtEdit 
         Enabled         =   0   'False
         Height          =   360
         Index           =   4
         Left            =   5910
         MaxLength       =   20
         TabIndex        =   8
         Top             =   375
         Width           =   1530
      End
      Begin VB.TextBox txtEdit 
         Enabled         =   0   'False
         Height          =   360
         Index           =   0
         Left            =   1125
         MaxLength       =   20
         TabIndex        =   2
         Top             =   390
         Width           =   915
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ע"
         Height          =   180
         Index           =   0
         Left            =   660
         TabIndex        =   9
         Top             =   960
         Width           =   360
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���뷶Χ"
         Height          =   180
         Index           =   6
         Left            =   2355
         TabIndex        =   3
         Top             =   465
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         Height          =   240
         Index           =   5
         Left            =   5250
         TabIndex        =   6
         Top             =   435
         Width           =   240
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������"
         Height          =   180
         Index           =   7
         Left            =   285
         TabIndex        =   1
         Top             =   480
         Width           =   720
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   420
      Left            =   8145
      TabIndex        =   26
      Top             =   780
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   420
      Left            =   8145
      TabIndex        =   25
      Top             =   285
      Width           =   1200
   End
End
Attribute VB_Name = "frmBillInDamnifyEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Enum EditDamnifyType
    EdS_���� = 0
    EdS_�鿴 = 2
End Enum
Private mstrPrivs As String, mlngModule As Long
Private mEditType As EditDamnifyType '�༭����
Private mblnChange As Boolean     'Ϊ��ʱ��ʾ�Ѹı���
Private mintSucceed As Integer
Private mlng���� As Long
Private mlng���ID  As Long, mintƱ�� As Integer 'Ʊ��
Private mlng����ID As Long
Private mblnFirst As Boolean
Private Enum mTxtIdx
    idx_���� = 0
    idx_��ʼǰ׺ = 1
    idx_��ʼ���� = 2
    idx_��ֹǰ׺ = 3
    idx_��ֹ���� = 4
    idx_��ע = 5
    idx_����ʼ = 6
    idx_������� = 7
    idx_����ԭ�� = 8
    idx_����ʱ�� = 9
    idx_������ = 10
End Enum

Public Function zlBillEdit(ByVal frmMain As Form, ByVal EditType As EditDamnifyType, ByVal strPrivs As String, _
    ByVal lngModule As Long, ByVal intƱ�� As gBillType, ByVal lng���ID As Long, Optional lng����ID As Long = 0) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������,Ʊ����ⱨ����(�������ӺͲ鿴)
    '���:frmMain-����������
    '       BillEditType-���ݲ�������
    '       strPrivs-Ȩ�޴�
    '       lngModule-ģ���
    '       lng���ID-����ָ�����ε����
    '       lng����ID-�޸Ļ�鿴ʱ�ı���IDֵ.
    '����:
    '����:����һ�����ϳɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2010-11-16 10:29:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mEditType = EditType: mstrPrivs = strPrivs: mlngModule = lngModule
    mintƱ�� = intƱ��: mlng���ID = lng���ID: mlng����ID = lng����ID
    mintSucceed = False
    Me.Show 1, frmMain
    zlBillEdit = mintSucceed > 0
End Function

Private Function LoadCardData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ؿ�Ƭ����
    '����:���سɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2010-11-16 10:35:27
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, lngLen As Long
    
    If UserInfo.���� = "" Then
        MsgBox "�㻹δ������Ա�Ķ��չ�ϵ������ϵͳ����Ա��ϵ�����ú����ʹ�ñ����ܡ�", vbExclamation, gstrSysName
        Exit Function
    End If
    Call ClearData  '����ؼ�����
    Err = 0: On Error GoTo errHandle
    If mEditType <> EdS_���� Then
        If mintƱ�� = gBillType.���ѿ� Then
            gstrSQL = _
                "Select Id, ���id, ��ʼ���� As ��ʼ����, ��ֹ���� As ��ֹ����, ����, ����ԭ��, ������, ����ʱ�� " & _
                "From ���ѿ������¼ where id=[1]"
        Else
            gstrSQL = _
                "Select Id, ���id, ��ʼ����, ��ֹ����, ����, ����ԭ��, ������, ����ʱ�� " & _
                "From Ʊ�ݱ����¼ where id=[1]"
        End If
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng����ID)
        If rsTemp.RecordCount = 0 Then
            MsgBox "ע��:" & vbCrLf & "    �����εı��𵥾ݿ����Ѿ�������ɾ�������飡", vbOKOnly + vbInformation + vbDefaultButton1, gstrSysName
            Exit Function
        End If
        txtEdit(mTxtIdx.idx_������) = Nvl(rsTemp!������)
        txtEdit(mTxtIdx.idx_����ԭ��) = Nvl(rsTemp!����ԭ��)
        txtEdit(mTxtIdx.idx_����ʱ��) = Format(rsTemp!����ʱ��, "yyyy-mm-dd HH:MM:SS")
        With vsMemo
            .Clear 1
            .Rows = 2
            .TextMatrix(1, .ColIndex("���")) = 1
            If Nvl(rsTemp!��ʼ����) <> Nvl(rsTemp!��ֹ����) And Nvl(rsTemp!��ֹ����) <> "" Then
                .TextMatrix(1, .ColIndex("����Ʊ��")) = Nvl(rsTemp!��ʼ����) & "-" & Nvl(rsTemp!��ֹ����)
            Else
                .TextMatrix(1, .ColIndex("����Ʊ��")) = Nvl(rsTemp!��ʼ����)
            End If
            .TextMatrix(1, .ColIndex("��������")) = Nvl(rsTemp!����)
        End With
        mlng���ID = Val(Nvl(rsTemp!���ID))
    End If
    
    If mintƱ�� = gBillType.���ѿ� Then
        gstrSQL = _
            "Select Id, ǰ׺�ı�, ��ʼ���� As ��ʼ����, ��ֹ���� As ��ֹ����, �������, ʣ������, ��ע, �Ǽ���, �Ǽ�ʱ��  " & _
            "From ���ѿ�����¼ " & _
            "Where Id=[1]"
    Else
        gstrSQL = _
            "Select Id, ǰ׺�ı�, ��ʼ����, ��ֹ����, �������, ʣ������, ��ע, �Ǽ���, �Ǽ�ʱ��  " & _
            "From Ʊ������¼ " & _
            "Where Id=[1]"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng���ID)
    If rsTemp.RecordCount = 0 Then
        MsgBox "ע��:" & vbCrLf & "    �����ε�����¼�Ѿ�������ɾ�������飡", vbOKOnly + vbInformation + vbDefaultButton1, gstrSysName
        Exit Function
    End If
    txtEdit(mTxtIdx.idx_����).Text = Nvl(rsTemp!ID)
    txtEdit(mTxtIdx.idx_��ʼǰ׺).Text = Nvl(rsTemp!ǰ׺�ı�)
    lngLen = Len(Trim(txtEdit(mTxtIdx.idx_��ʼǰ׺).Text))
    txtEdit(mTxtIdx.idx_��ʼ����).Text = Mid(Nvl(rsTemp!��ʼ����), lngLen + 1)
    txtEdit(mTxtIdx.idx_��ʼ����).Tag = txtEdit(mTxtIdx.idx_��ʼ����).Text
    txtEdit(mTxtIdx.idx_��ֹǰ׺).Text = Nvl(rsTemp!ǰ׺�ı�)
    txtEdit(mTxtIdx.idx_��ֹ����).Text = Mid(Nvl(rsTemp!��ֹ����), lngLen + 1)
    txtEdit(mTxtIdx.idx_��ֹ����).Tag = txtEdit(mTxtIdx.idx_��ֹ����).Text
    txtEdit(mTxtIdx.idx_����ʼ).MaxLength = Len(Nvl(rsTemp!��ʼ����))
    txtEdit(mTxtIdx.idx_�������).MaxLength = txtEdit(mTxtIdx.idx_����ʼ).MaxLength
    If mEditType = Ed_���� Then
        txtEdit(mTxtIdx.idx_������) = UserInfo.����
        txtEdit(mTxtIdx.idx_����ʱ��) = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
        LoadCardData = True
        Exit Function
    End If
    Call RefreshNo
    LoadCardData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function GetBillNum(ByVal str��ʼ���� As String, ByVal str�տ����� As String, Optional ByRef strErrMsg As String = "") As Long
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡƱ������
    '���:str��ʼ����-����Ϊ����;
    '       str�տ�����-����Ϊ����
    '����:strErrMsg-���ش���ļ�����Ϣ
    '����:Ʊ��������
    '����:���˺�
    '����:2010-11-16 11:06:59
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo errHandle
    strErrMsg = ""
    If (str��ʼ���� = "" And str�տ����� <> "") Or (str�տ����� = "" And str��ʼ���� <> "") Then
        GetBillNum = 1: Exit Function
    End If
    GetBillNum = CDec(str�տ�����) - CDec(str��ʼ����) + 1
    Exit Function
errHandle:
    strErrMsg = "�������򳬳��˼��㷶Χ"
    GetBillNum = 0
End Function

Private Sub ClearData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ؼ�����
    '����:���˺�
    '����:2010-11-16 10:35:55
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer
    For i = 0 To txtEdit.UBound
        txtEdit(i).Text = ""
        If txtEdit(i).Enabled = False Then
            txtEdit(i).BackColor = Me.BackColor
        Else
            txtEdit(i).BackColor = &H80000005
        End If
    Next
    vsMemo.Clear 1
    vsMemo.Rows = 2
End Sub

Private Sub cmdAdd_Click()
    '��������
    Dim i As Long, lngRow As Long, str��ʼƱ�� As String, str����Ʊ�� As String
    Dim lngǰ׺ As Long, lng���� As Long
    
    On Error GoTo errHandle
    If CheckInputValied = False Then Exit Sub
    With vsMemo
        .Redraw = flexRDNone
        For i = 1 To .Rows - 1
            .TextMatrix(i, .ColIndex("���")) = i
            If Trim(.TextMatrix(i, .ColIndex("����Ʊ��"))) = "" Then
                lngRow = i: Exit For
            End If
        Next
        If lngRow = 0 Then
            .Rows = .Rows + 1
            lngRow = .Rows - 1
            .TextMatrix(lngRow, .ColIndex("���")) = lngRow
        End If
        str��ʼƱ�� = Trim(txtEdit(mTxtIdx.idx_����ʼ))
        str����Ʊ�� = Trim(txtEdit(mTxtIdx.idx_�������))
        lngǰ׺ = Len(Trim(txtEdit(mTxtIdx.idx_��ʼǰ׺)))
        If str��ʼƱ�� = str����Ʊ�� Then
            .TextMatrix(lngRow, .ColIndex("����Ʊ��")) = str��ʼƱ��
            .Cell(flexcpData, lngRow, .ColIndex("����Ʊ��")) = Mid(str��ʼƱ��, lngǰ׺ + 1)
            .TextMatrix(lngRow, .ColIndex("��������")) = 1
        Else
            lng���� = GetBillNum(Mid(str��ʼƱ��, lngǰ׺ + 1), Mid(str����Ʊ��, lngǰ׺ + 1))
            
            .TextMatrix(lngRow, .ColIndex("����Ʊ��")) = str��ʼƱ�� & IIf(str��ʼƱ�� = "" Or str����Ʊ�� = "", "", "-") & str����Ʊ��
            .Cell(flexcpData, lngRow, .ColIndex("����Ʊ��")) = Mid(str��ʼƱ��, lngǰ׺ + 1) & "-" & Mid(str����Ʊ��, lngǰ׺ + 1)
            .TextMatrix(lngRow, .ColIndex("��������")) = lng����
        End If
        .Row = lngRow
        .Redraw = flexRDBuffered
    End With
    txtEdit(mTxtIdx.idx_�������).Text = "": txtEdit(mTxtIdx.idx_����ʼ).Text = ""
    zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_����ʼ)
    Exit Sub
errHandle:
    vsMemo.Redraw = flexRDBuffered
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub cmdRemove_Click()
    Dim lngRow As Long
    On Error GoTo errHandle
    'ɾ����ӵı���
    With vsMemo
        .Redraw = flexRDNone
        lngRow = .Row
        If lngRow < .Rows - 1 Then
            .Row = lngRow + 1
            .RemoveItem lngRow
        ElseIf lngRow = .Rows - 1 And lngRow = 1 Then
            .Cell(flexcpText, lngRow, 0, lngRow, .Cols - 1) = ""
            .Cell(flexcpData, lngRow, 0, lngRow, .Cols - 1) = ""
        Else
            .Row = lngRow - 1
            .RemoveItem lngRow
        End If
        .Redraw = flexRDBuffered
        Call RefreshNo
    End With
    Exit Sub
errHandle:
    vsMemo.Redraw = flexRDBuffered
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub RefreshNo()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ˢ�����
    '����:���˺�
    '����:2010-11-17 11:58:00
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    With vsMemo
        For i = 1 To .Rows - 1
            .TextMatrix(i, .ColIndex("���")) = i
        Next
    End With
End Sub

Private Sub cmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    If LoadCardData = False Then Unload Me: Exit Sub
    Call SetCtrlEnable
    zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_����ʼ)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab: Exit Sub
    If InStr("'[]����������,.'�ۣ�", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub SetCtrlEnable()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ÿؼ���Enable����
    '����:���˺�
    '����:2010-11-17 17:18:26
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    For i = 0 To txtEdit.UBound
        If mEditType = EdS_�鿴 Then
            txtEdit(i).Enabled = False
        End If
        If txtEdit(i).Enabled = False Then
            txtEdit(i).BackColor = Me.BackColor
        Else
            txtEdit(i).BackColor = &H80000005
        End If
    Next
End Sub

Private Sub Form_Load()
    Dim blnBill As Boolean
    
    mblnFirst = True
    blnBill = CurrentIsBill(mintƱ��)
    lbl(6).Caption = IIf(blnBill, "���뷶Χ", "���ŷ�Χ")
    lblE.Caption = IIf(blnBill, "����Ʊ��(&B)", "���𿨺�(&B)")
    vsMemo.TextMatrix(0, vsMemo.ColIndex("����Ʊ��")) = IIf(blnBill, "����Ʊ��", "���𿨺�")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange = False Then Exit Sub
    If MsgBox("�����������˳��Ļ������е��޸Ķ�������Ч��" & vbCrLf & "�Ƿ�ȷ���˳���", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Function CheckInputValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������ı���ʼ�Ż�������Ƿ�Ϸ�
    '���:
    '����:
    '����:
    '����:���˺�
    '����:2010-11-16 17:48:55
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, strStartNo As String, strEndNo As String, strTemp As String
    Dim lngLen As Integer, str���� As String, str���� As String, rsTemp As ADODB.Recordset
    Dim strName As String
    
    On Error GoTo errHandle
    strName = IIf(CurrentIsBill(mintƱ��), "����", "����")
    If Trim(txtEdit(mTxtIdx.idx_����ʼ).Text) = "" And Trim(txtEdit(mTxtIdx.idx_�������).Text) = "" Then
        ShowMsgbox "ע��" & vbCrLf & "    ����Χ�еĿ�ʼ" & strName & "�����" & strName & "��������,����!"
        zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_����ʼ): Exit Function
    End If
    lngLen = Len(txtEdit(mTxtIdx.idx_��ʼǰ׺))
    strTemp = Mid(txtEdit(mTxtIdx.idx_����ʼ).Text, lngLen + 1)
    If strTemp <> "" Then
        If zlIsOnlyNum(strTemp) = False Then
            MsgBox "����Χ�еĿ�ʼ" & strName & "�к��з������ַ�����ĸֻ����Ϊǰ׺��", vbExclamation, gstrSysName
            zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_����ʼ): Exit Function
        End If
    End If
    strTemp = Mid(txtEdit(mTxtIdx.idx_�������).Text, lngLen + 1)
    If strTemp <> "" Then
        If zlIsOnlyNum(strTemp) = False Then
                MsgBox "����Χ�е���ֹ" & strName & "�к��з������ַ�����ĸֻ����Ϊǰ׺��", vbExclamation, gstrSysName
                zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_�������): Exit Function
        End If
    End If
    mlng���� = zlCommFun.ActualLen(txtEdit(mTxtIdx.idx_��ʼǰ׺) & txtEdit(mTxtIdx.idx_��ʼ����))
    If txtEdit(mTxtIdx.idx_����ʼ).Text <> "" Then
        If zlCommFun.ActualLen(txtEdit(mTxtIdx.idx_����ʼ).Text) <> mlng���� Then
            ShowMsgbox "ע��" & vbCrLf & "    ����Χ�еĿ�ʼ" & strName & "���Ȳ���(ӦΪ" & mlng���� & "λ()),����!"
            zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_����ʼ): Exit Function
        End If
    End If
    If txtEdit(mTxtIdx.idx_�������).Text <> "" Then
        If zlCommFun.ActualLen(txtEdit(mTxtIdx.idx_�������).Text) <> mlng���� Then
            ShowMsgbox "ע��" & vbCrLf & "    ����Χ�еĽ���" & strName & "���Ȳ���(ӦΪ" & mlng���� & "λ),����!"
            zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_�������): Exit Function
        End If
        If txtEdit(mTxtIdx.idx_�������).Text < txtEdit(mTxtIdx.idx_����ʼ) _
            And Trim(txtEdit(mTxtIdx.idx_�������).Text) <> "" And txtEdit(mTxtIdx.idx_����ʼ) <> "" Then
            ShowMsgbox "ע��" & vbCrLf & "    ����Χ�еĽ���" & strName & "С���˿�ʼ" & strName & ",����!"
            zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_�������): Exit Function
        End If
    End If
    '����Ƿ��������������
    Dim varTemp As Variant
    With vsMemo
        strStartNo = Trim(txtEdit(mTxtIdx.idx_����ʼ))
        strEndNo = Trim(txtEdit(mTxtIdx.idx_�������))
        For i = 1 To .Rows - 1
             If .TextMatrix(i, .ColIndex("����Ʊ��")) <> "" Then
                varTemp = Split(.TextMatrix(i, .ColIndex("����Ʊ��")) & "-", "-")
                If varTemp(1) = "" Then varTemp(1) = varTemp(0)
                If varTemp(0) <> "" And varTemp(1) <> "" Then
                    If strStartNo >= varTemp(0) And strStartNo <= varTemp(1) Then
                        ShowMsgbox "ע��" & vbCrLf & "    ����Χ�еĿ�ʼ" & strName & "�Ѿ��������˵�" & i & "������,����!"
                        .Row = i
                        If .RowIsVisible(.Row) = False Then Call .ShowCell(.Row, .Col)
                        zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_����ʼ): Exit Function
                    End If
                    
                    If strEndNo >= varTemp(0) And strEndNo <= varTemp(1) And strEndNo <> "" Then
                        ShowMsgbox "ע��" & vbCrLf & "    ����Χ�еĽ���" & strName & "�Ѿ��������˵�" & i & "������,����!"
                        .Row = i
                        If .RowIsVisible(.Row) = False Then Call .ShowCell(.Row, .Col)
                        zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_�������): Exit Function
                    End If
                ElseIf varTemp(0) <> "" Then
                    If strStartNo = varTemp(0) Then
                        ShowMsgbox "ע��" & vbCrLf & "    ����Χ�еĿ�ʼ" & strName & "�Ѿ��������˵�" & i & "������,����!"
                        .Row = i
                        If .RowIsVisible(.Row) = False Then Call .ShowCell(.Row, .Col)
                        zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_����ʼ): Exit Function
                    End If
                    If strEndNo = varTemp(0) Then
                        ShowMsgbox "ע��" & vbCrLf & "    ����Χ�еĽ���" & strName & "�Ѿ��������˵�" & i & "������,����!"
                        .Row = i
                        If .RowIsVisible(.Row) = False Then Call .ShowCell(.Row, .Col)
                        zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_�������): Exit Function
                    End If
                ElseIf varTemp(1) <> "" Then
                    If strStartNo = varTemp(1) Then
                        ShowMsgbox "ע��" & vbCrLf & "    ����Χ�еĿ�ʼ" & strName & "�Ѿ��������˵�" & i & "������,����!"
                        .Row = i
                        If .RowIsVisible(.Row) = False Then Call .ShowCell(.Row, .Col)
                        zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_����ʼ): Exit Function
                    End If
                    If strEndNo = varTemp(1) Then
                        ShowMsgbox "ע��" & vbCrLf & "    ����Χ�еĽ���" & strName & "�Ѿ��������˵�" & i & "������,����!"
                        .Row = i
                        If .RowIsVisible(.Row) = False Then Call .ShowCell(.Row, .Col)
                        zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_�������): Exit Function
                    End If
                End If
             End If
        Next
    End With
    
    '����Ƿ����ʹ�õ����
    If mintƱ�� = gBillType.���ѿ� Then
        gstrSQL = "" & _
        "   Select 1 as ���,��ʼ���� As ��ʼ����,��ֹ���� As ��ֹ���� " & _
        "   From ���ѿ������¼ " & _
        "   Where (([1] between ��ʼ���� and ��ֹ����) or ([2] between ��ʼ���� and ��ֹ����)) and ���ID=[3] " & _
        "   Union ALL " & _
        "   Select 2 as ���,��ʼ���� As ��ʼ����, ��ֹ���� As ��ֹ���롡" & _
        "   From ���ѿ����ü�¼ " & _
        "   Where (([1] between ��ʼ���� and ��ֹ����) or ([2] between ��ʼ���� and ��ֹ����)) and ����=[3]  "
    Else
        gstrSQL = "" & _
        "   Select 1 as ���,��ʼ����,��ֹ���� " & _
        "   From Ʊ�ݱ����¼ " & _
        "   Where (([1]  between ��ʼ����  and ��ֹ����  ) or ([2] between ��ʼ����  and ��ֹ����  )) and ���ID=[3] " & _
        "   Union ALL " & _
        "   Select 2 as ���,��ʼ����, ��ֹ���롡" & _
        "   From Ʊ�����ü�¼ " & _
        "   Where (([1]  between ��ʼ����  and ��ֹ����  ) or ([2] between ��ʼ����  and ��ֹ����  )) and ����=[3] and Ʊ��=[4]  "
    End If
    If strStartNo = "" Then strStartNo = strEndNo
    If strEndNo = "" Then strEndNo = strStartNo
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strStartNo, strEndNo, mlng���ID, mintƱ��)
    
    If Not rsTemp.EOF Then
        str���� = "": str���� = ""
        Do While Not rsTemp.EOF
            If Nvl(rsTemp!��ֹ����) = Nvl(rsTemp!��ʼ����) Then
                strTemp = Nvl(rsTemp!��ʼ����)
            Else
                strTemp = Nvl(rsTemp!��ʼ����) & "-" & Nvl(rsTemp!��ֹ����)
            End If
            If rsTemp!��� = 1 Then
               If Len(str����) <= 50 Then
                    str���� = str���� & vbCrLf & strTemp
               Else
                  If InStr(1, str����, "...") = 0 Then str���� = str���� & vbCrLf & "..."
               End If
            Else
               If Len(str����) <= 50 Then
                    str���� = str���� & vbCrLf & strTemp
               Else
                  If InStr(1, str����, "...") = 0 Then str���� = str���� & vbCrLf & "..."
               End If
            End If
            rsTemp.MoveNext
        Loop
        If str���� <> "" Then
            ShowMsgbox "ע��:" & vbCrLf & "    ��ǰ����Χ�е�" & strName & "�Ѿ�������,�ѱ����" & strName & "����:" & vbCrLf & str����
            zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_����ʼ): Exit Function
            Exit Function
        End If
        If str���� <> "" Then
            ShowMsgbox "ע��:" & vbCrLf & "    ��ǰ����Χ�е�" & strName & "�Ѿ�������,�����õ�" & strName & "����:" & vbCrLf & str����
            zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_����ʼ): Exit Function
            Exit Function
        End If
    End If
    
    CheckInputValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function isValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������������Ƿ�Ϸ�
    '����:�Ϸ�����true,���򷵻�False
    '����:���˺�
    '����:2010-11-16 15:04:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varTemp As Variant, varData As Variant, i As Long, str���� As String, str���� As String, strTemp As String
    Dim rsTemp As ADODB.Recordset, blnHaveData As Boolean '�Ƿ��������
    Dim blnBill As Boolean
    
    On Error GoTo errHandle
    blnBill = CurrentIsBill(mintƱ��)
    If zlCommFun.ActualLen(Trim(txtEdit(mTxtIdx.idx_����ԭ��))) > 200 Then
        ShowMsgbox "ע��" & vbCrLf & "    ����ԭ�����ֻ������200���ַ���100������,����!"
        zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_����ԭ��): Exit Function
    End If
    
    If InStr(1, txtEdit(mTxtIdx.idx_����ԭ��), "'") > 0 Then
        ShowMsgbox "ע��" & vbCrLf & "    ����ԭ���к��зǷ��ַ�������,����!"
        zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_����ԭ��): Exit Function
    End If
    
    If txtEdit(mTxtIdx.idx_��ʼ����).Text = String("0", mlng����) And txtEdit(mTxtIdx.idx_��ֹǰ׺).Text = String("9", mlng����) Then
        MsgBox "����ʹ��" & String("0", mlng����) & "-" & String("9", mlng����) & "��" & IIf(blnBill, "Ʊ��", "����") & "��Χ��", vbExclamation, gstrSysName
        zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_��ֹ����): Exit Function
    End If
    
    '����Ƿ�ú����Ѿ�ʹ�û����Ѿ����õģ������ٱ���ģ��Ѿ������˵�,Ҳ���ܱ���
    With vsMemo
        blnHaveData = False
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("����Ʊ��")) <> "" Then
                blnHaveData = True
                varTemp = Split(.TextMatrix(i, .ColIndex("����Ʊ��")) & "-", "-")
                If varTemp(1) = "" Then varTemp(1) = varTemp(0)
                '1.��鱨�����
                gstrSQL = "" & _
                "   Select 1 as ���,��ʼ����,��ֹ���� " & _
                "   From Ʊ�ݱ����¼ " & _
                "   Where (��ʼ����>=[1]  and ��ֹ���� <=[1]) or (��ʼ����>=[2] and ��ֹ����<=[2] ) and ���ID=[3] " & _
                "   Union ALL " & _
                "   Select 2 as ���,��ʼ����, ��ֹ���롡" & _
                "   From Ʊ�����ü�¼ " & _
                "   Where (��ʼ����>=[1]  and ��ֹ���� <=[1]) or (��ʼ����>=[2] and ��ֹ����<=[2] ) and ����=[3] and Ʊ��=[4] " & _
                "   "
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, CStr(varTemp(0)), CStr(varTemp(1)), mlng���ID, mintƱ��)
                If Not rsTemp.EOF Then
                    str���� = "": str���� = ""
                    Do While Not .EOF
                        If Nvl(rsTemp!��ֹ����) = Nvl(rsTemp!��ʼ����) Then
                            strTemp = Nvl(rsTemp!��ʼ����)
                        Else
                            strTemp = Nvl(rsTemp!��ʼ����) & "-" & Nvl(rsTemp!��ֹ����)
                        End If
                        If rsTemp!��� = 1 Then
                           If Len(str����) <= 50 Then
                                str���� = str���� & vbCrLf & strTemp
                           Else
                              If InStr(1, str����, "...") = 0 Then str���� = str���� & vbCrLf & "..."
                           End If
                        Else
                           If Len(str����) <= 50 Then
                                str���� = str���� & vbCrLf & strTemp
                           Else
                              If InStr(1, str����, "...") = 0 Then str���� = str���� & vbCrLf & "..."
                           End If
                        End If
                        rsTemp.MoveNext
                    Loop
                    If str���� <> "" Then
                        ShowMsgbox "ע��:" & vbCrLf & "    �ڵ�" & i + 1 & "�м�¼�а������Ѿ����������" & IIf(blnBill, "Ʊ��", "����") & _
                            ",�ѱ����" & IIf(blnBill, "Ʊ��", "����") & "����:" & vbCrLf & str����
                        Exit Function
                    End If
                    If str���� <> "" Then
                        ShowMsgbox "ע��:" & vbCrLf & "    �ڵ�" & i + 1 & "�м�¼�а������Ѿ����õ�" & IIf(blnBill, "Ʊ��", "����") & _
                            ",�����õ�" & IIf(blnBill, "Ʊ��", "����") & "����:" & vbCrLf & str����
                        Exit Function
                    End If
                End If
            End If
        Next
    End With
    If Not blnHaveData Then
        ShowMsgbox "ע��:" & vbCrLf & "    ��û��ѡ��Ҫ�����" & IIf(blnBill, "Ʊ��", "��Ƭ") & "�����ܼ�����"
        Exit Function
    End If
    isValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function SaveData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������
    '����:���ݱ���ɹ�,����true,���򷵻�ΪFalse
    '����:���˺�
    '����:2010-11-16 15:15:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngID As Long, lng�������� As Long, strDate As String
    Dim i As Long, cllPro As Collection, varTemp As Variant, varData As Variant
    
    On Error GoTo errHandle
    Set cllPro = New Collection
    With vsMemo
        strDate = "to_Date('" & Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd HH24:mi:ss')"
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("����Ʊ��")) <> "" Then
                varTemp = Split(.TextMatrix(i, .ColIndex("����Ʊ��")) & "-", "-")
                If varTemp(1) = "" Then varTemp(1) = varTemp(0)
                lng�������� = Val(.TextMatrix(i, .ColIndex("��������")))
                If mintƱ�� = gBillType.���ѿ� Then
                    '    Zl_���ѿ������¼_Insert
                    gstrSQL = "Zl_���ѿ������¼_Insert("
                    '      ���id_In   In ���ѿ������¼.���id%Type,
                    gstrSQL = gstrSQL & "" & mlng���ID & ","
                    '      ��ʼ����_In In ���ѿ������¼.��ʼ����%Type,
                    gstrSQL = gstrSQL & "'" & varTemp(0) & "',"
                    '      ��ֹ����_In In ���ѿ������¼.��ֹ����%Type,
                    gstrSQL = gstrSQL & "'" & varTemp(1) & "',"
                    '      ����_In     In ���ѿ������¼.����%Type,
                    gstrSQL = gstrSQL & "" & lng�������� & ","
                    '      ����ԭ��_In In ���ѿ������¼.����ԭ��%Type,
                    gstrSQL = gstrSQL & "" & _
                        IIf(Trim(txtEdit(mTxtIdx.idx_����ԭ��).Text) = "", "NULL", _
                        "'" & Trim(txtEdit(mTxtIdx.idx_����ԭ��).Text) & "'") & ","
                    '      ������_In   In ���ѿ������¼.������%Type,
                    gstrSQL = gstrSQL & "'" & UserInfo.���� & "',"
                    '      ����ʱ��_In In ���ѿ������¼.����ʱ��%Type
                    gstrSQL = gstrSQL & "" & strDate & ")"
                Else
                    '    Zl_Ʊ�ݱ����¼_Insert
                    gstrSQL = "Zl_Ʊ�ݱ����¼_Insert("
                    '      ���id_In   In Ʊ�ݱ����¼.���id%Type,
                    gstrSQL = gstrSQL & "" & mlng���ID & ","
                    '      ��ʼ����_In In Ʊ�ݱ����¼.��ʼ����%Type,
                    gstrSQL = gstrSQL & "'" & varTemp(0) & "',"
                    '      ��ֹ����_In In Ʊ�ݱ����¼.��ֹ����%Type,
                    gstrSQL = gstrSQL & "'" & varTemp(1) & "',"
                    '      ����_In     In Ʊ�ݱ����¼.����%Type,
                    gstrSQL = gstrSQL & "" & lng�������� & ","
                    '      ����ԭ��_In In Ʊ�ݱ����¼.����ԭ��%Type,
                    gstrSQL = gstrSQL & "" & IIf(Trim(txtEdit(mTxtIdx.idx_����ԭ��)) = "", "NULL", "'" & Trim(txtEdit(mTxtIdx.idx_����ԭ��)) & "'") & ","
                    '      ������_In   In Ʊ�ݱ����¼.������%Type,
                    gstrSQL = gstrSQL & "'" & UserInfo.���� & "',"
                    '      ����ʱ��_In In Ʊ�ݱ����¼.����ʱ��%Type
                    gstrSQL = gstrSQL & "" & strDate & ")"
                End If
                AddArray cllPro, gstrSQL
            End If
        Next
    End With
    ExecuteProcedureArrAy cllPro, Me.Caption
    SaveData = True
    Exit Function
errHandle:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub cmdOK_Click()
    If mEditType = Ed_�鿴 Then
        mblnChange = False
        Unload Me: Exit Sub
    End If
    If isValied() = False Then Exit Sub
    If SaveData() = False Then Exit Sub
    mintSucceed = mintSucceed + 1
    mblnChange = False
    Unload Me
End Sub

Private Sub txtEdit_Change(Index As Integer)
    If mEditType = Ed_�鿴 Then Exit Sub
    mblnChange = True
End Sub
Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtEdit(Index)
    If idx_����ԭ�� = Index Then zlCommFun.OpenIme True
End Sub

Private Sub txtEdit_LostFocus(Index As Integer)
    Dim lngLen As Long, lngǰ׺Len As Long, strTemp As String, strǰ׺ As String
    Dim strNum As String
    Dim strChr As String
    Dim i As Long
    
    If Index = mTxtIdx.idx_����ʼ Or Index = mTxtIdx.idx_������� Then
        '���Ȳ���ʱ����Ҫ��λ
        strTemp = Trim(txtEdit(Index).Text)
        If strTemp = "" Then Exit Sub
        lngLen = Len(txtEdit(mTxtIdx.idx_��ʼ����))
        strǰ׺ = Trim(txtEdit(mTxtIdx.idx_��ʼǰ׺))
        lngǰ׺Len = Len(strǰ׺)
        If Len(txtEdit(Index)) < lngLen Then
            If zlIsOnlyNum(strTemp) Then
                strTemp = strǰ׺ & zlStr.Lpad(strTemp, lngLen, "0", True)
            ElseIf UCase(Mid(strTemp, 1, lngǰ׺Len)) = strǰ׺ Then
                  strTemp = strǰ׺ & zlStr.Lpad(Mid(strTemp, lngǰ׺Len + 1), lngLen, "0", True)
            Else
                strNum = ""
                For i = 1 To Len(strTemp)
                    strChr = Mid(strTemp, i, 1)
                    If InStr(1, "0123456789", strChr) > 0 Then
                        strNum = strNum & strChr
                    End If
                Next
                strTemp = strǰ׺ & zlStr.Lpad(strNum, lngLen, "0", True)
            End If
        ElseIf UCase(Mid(strTemp, 1, lngǰ׺Len)) = strǰ׺ Then
                strTemp = strǰ׺ & Right(Mid(strTemp, lngǰ׺Len + 1), lngLen)
        Else
                strTemp = Left(strTemp, lngǰ׺Len + lngLen)
        End If
        txtEdit(Index).Text = UCase(strTemp)
    End If
    txtEdit(Index).Text = Trim(txtEdit(Index).Text)
    If idx_����ԭ�� = Index Then zlCommFun.OpenIme False
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txtEdit(Index), KeyAscii, m�ı�ʽ
End Sub
