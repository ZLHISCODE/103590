VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmBillInEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ʊ�����༭"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8805
   Icon            =   "frmBillInEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   8805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   420
      Left            =   7455
      TabIndex        =   22
      Top             =   5580
      Width           =   1200
   End
   Begin VB.Frame fraUse 
      Caption         =   "��������Ϣ"
      Height          =   2490
      Left            =   135
      TabIndex        =   21
      Top             =   390
      Width           =   6990
      Begin VB.ComboBox cbo��� 
         Height          =   300
         Left            =   1110
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   390
         Width           =   2670
      End
      Begin VB.TextBox txtEdit 
         Enabled         =   0   'False
         Height          =   360
         Index           =   7
         Left            =   4605
         MaxLength       =   20
         TabIndex        =   15
         Top             =   1785
         Width           =   2265
      End
      Begin VB.TextBox txtEdit 
         Enabled         =   0   'False
         Height          =   360
         Index           =   6
         Left            =   1110
         MaxLength       =   20
         TabIndex        =   13
         Top             =   1785
         Width           =   2655
      End
      Begin VB.TextBox txtEdit 
         Height          =   330
         Index           =   0
         Left            =   4935
         MaxLength       =   20
         TabIndex        =   3
         Top             =   375
         Width           =   1950
      End
      Begin VB.TextBox txtEdit 
         Height          =   360
         IMEMode         =   2  'OFF
         Index           =   4
         Left            =   4605
         MaxLength       =   20
         TabIndex        =   9
         Top             =   855
         Width           =   2295
      End
      Begin VB.TextBox txtEdit 
         Height          =   360
         IMEMode         =   2  'OFF
         Index           =   3
         Left            =   4215
         MaxLength       =   2
         TabIndex        =   8
         Top             =   855
         Width           =   375
      End
      Begin VB.TextBox txtEdit 
         Height          =   360
         IMEMode         =   2  'OFF
         Index           =   2
         Left            =   1515
         MaxLength       =   20
         TabIndex        =   6
         Top             =   855
         Width           =   2295
      End
      Begin VB.TextBox txtEdit 
         Enabled         =   0   'False
         Height          =   360
         Index           =   5
         Left            =   1110
         MaxLength       =   20
         TabIndex        =   11
         Top             =   1320
         Width           =   5775
      End
      Begin VB.TextBox txtEdit 
         Height          =   360
         IMEMode         =   2  'OFF
         Index           =   1
         Left            =   1125
         MaxLength       =   2
         TabIndex        =   5
         Top             =   855
         Width           =   375
      End
      Begin VB.Label lblUserType 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ʹ�����(&L)"
         Height          =   180
         Left            =   120
         TabIndex        =   0
         Top             =   450
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������"
         Height          =   180
         Index           =   7
         Left            =   4200
         TabIndex        =   2
         Top             =   450
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         Height          =   240
         Index           =   5
         Left            =   3945
         TabIndex        =   7
         Top             =   945
         Width           =   240
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���뷶Χ(&B)"
         Height          =   180
         Index           =   6
         Left            =   120
         TabIndex        =   4
         Top             =   945
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Ǽ�ʱ��"
         Height          =   180
         Index           =   3
         Left            =   3870
         TabIndex        =   14
         Top             =   1875
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Ǽ���"
         Height          =   180
         Index           =   2
         Left            =   540
         TabIndex        =   12
         Top             =   1875
         Width           =   540
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ע(&G)"
         Height          =   180
         Index           =   0
         Left            =   480
         TabIndex        =   10
         Top             =   1410
         Width           =   630
      End
   End
   Begin VB.Frame Frame1 
      Height          =   6735
      Left            =   7215
      TabIndex        =   20
      Top             =   -15
      Width           =   30
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   420
      Left            =   7455
      TabIndex        =   19
      Top             =   690
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   420
      Left            =   7455
      TabIndex        =   18
      Top             =   210
      Width           =   1200
   End
   Begin VSFlex8Ctl.VSFlexGrid vsMemo 
      Height          =   3150
      Left            =   150
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   3330
      Width           =   6945
      _cx             =   12250
      _cy             =   5556
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
      GridColor       =   -2147483643
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   0
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   2
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmBillInEdit.frx":058A
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
   Begin VB.Label Label2 
      Caption         =   "��ϸ���"
      Height          =   255
      Left            =   135
      TabIndex        =   16
      Top             =   3090
      Width           =   975
   End
End
Attribute VB_Name = "frmBillInEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Enum BillInEditType
    Ed_���� = 0
    Ed_�޸� = 1
    Ed_�鿴 = 2
End Enum
Private mstrPrivs As String, mlngModule As Long
Private mEditType As BillInEditType '�༭����
Private mblnChange As Boolean     'Ϊ��ʱ��ʾ�Ѹı���
Private mstrƱ�ݳ��� As String '��ʾ����Ʊ�ݵĺ��볤�ȣ���λ�ֱ�Ϊ1-�շ�,2-Ԥ��,3-����,4-�Һ�,5-���￨  77777
Private mlng���� As Long       '��ǰƱ������ĳ���
Private mintSucceed As Integer
Private mlng���ID  As Long
Private mstrDrawBill As String, mstrDrawNum As String '���÷ֶ���Ϣ
Private mstrDamnifyBill As String, mlngDamnifyNum As Long  '���÷ֶ���Ϣ,���������ϼ�
Private mintƱ�� As gBillType  'Ʊ��
Private mblnIsBIll As Boolean '��ǰƱ���Ƿ�ΪƱ��
Private mblnFirst As Boolean
Private mstr��� As String 'ȱʡ�������
Private mstrPreType(1 To 6) As String
Private mcllCardProperty As Collection
Private mblnNotClick As Boolean
Private Enum mTxtIdx
    idx_���� = 0
    idx_��ʼǰ׺ = 1
    idx_��ʼ���� = 2
    idx_��ֹǰ׺ = 3
    idx_��ֹ���� = 4
    idx_��ע = 5
    idx_�Ǽ��� = 6
    idx_�Ǽ�ʱ�� = 7
End Enum
Public Function zlBillEdit(ByVal frmMain As Form, ByVal intƱ�� As gBillType, ByVal EditType As BillInEditType, ByVal strPrivs As String, _
    ByVal lngModule As Long, Optional ByVal lng���ID As Long = 0, Optional str��� As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������,Ʊ������ѯ���������(�������Ӻ��޸�)
    '���:frmMain-����������
    '       BillEditType-���ݲ�������
    '       strPrivs-Ȩ�޴�
    '       lngModule-ģ���
    '       lng���ID-�޸�ʱ,ת����������.
    '       str���:ʹ���������(27559)
    '����:
    '����:����һ�����ϳɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2010-11-16 10:29:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTemp As String
    mstr��� = str���: mintƱ�� = intƱ��: mblnIsBIll = CurrentIsBill(intƱ��)
    mstrPreType(mintƱ��) = mstr���
    mEditType = EditType: mstrPrivs = strPrivs: mlngModule = lngModule: mlng���ID = lng���ID
    mstrƱ�ݳ��� = zlDatabase.GetPara(20, glngSys, , "7|7|7|7|7")
    mintSucceed = False
    If mEditType = Ed_�鿴 Then
        cmdOK.Caption = "�˳�(&O)"
        cmdCancel.Visible = False
    End If
    strTemp = Decode(mintƱ��, gBillType.�շ��վ�, "�շ��վ�", gBillType.Ԥ���վ�, "Ԥ���վ�", gBillType.�����վ�, "�����վ�", _
        gBillType.�Һ��վ�, "�Һ��վ�", gBillType.���￨, "���￨", gBillType.���ѿ�, "���ѿ�", "���￨")
    Me.Caption = strTemp & "���"
    fraUse.Caption = "��" & strTemp & "����������Ϣ"
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
    Dim i As Long, blnFind As Boolean
    
    If mintƱ�� = gBillType.���￨ Or mintƱ�� = gBillType.���ѿ� Then
         mlng���� = mcllCardProperty(cbo���.ListIndex + 1)(0)
    Else
        mlng���� = Split(mstrƱ�ݳ���, "|")(mintƱ�� - 1)
    End If
    If UserInfo.���� = "" Then
        MsgBox "�㻹δ������Ա�Ķ��չ�ϵ������ϵͳ����Ա��ϵ�����ú����ʹ�ñ����ܡ�", vbExclamation, gstrSysName
        Exit Function
    End If
    
    Call ClearData  '����ؼ�����
    Err = 0: On Error GoTo errHandle
    If mEditType = Ed_���� Then
        If mintƱ�� = gBillType.���￨ Or mintƱ�� = gBillType.���ѿ� Then
            txtEdit(mTxtIdx.idx_��ʼǰ׺).Text = mcllCardProperty(cbo���.ListIndex + 1)(1)
        End If
        txtEdit(mTxtIdx.idx_�Ǽ���) = UserInfo.����
        txtEdit(mTxtIdx.idx_�Ǽ�ʱ��) = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
        
        If mintƱ�� = gBillType.���ѿ� Or mintƱ�� = gBillType.���ѿ� Then
            Call Setǰ׺(mcllCardProperty(cbo���.ListIndex + 1)(1))
        End If
        LoadCardData = True
        Exit Function
    End If
    
    If mintƱ�� = gBillType.���ѿ� Then
        gstrSQL = _
            "Select Id, �ӿڱ�� As ʹ�����, ǰ׺�ı�, ��ʼ���� As ��ʼ����, ��ֹ���� As ��ֹ����, �������, ʣ������, ��ע, �Ǽ���, �Ǽ�ʱ��, ����  " & _
            "From ���ѿ�����¼ " & _
            "Where Id=[1]"
    Else
        gstrSQL = _
            "Select Id, ʹ�����, ǰ׺�ı�, ��ʼ����, ��ֹ����, �������, ʣ������, ��ע, �Ǽ���, �Ǽ�ʱ��, ����  " & _
            "From Ʊ������¼ " & _
            "Where Id=[1]"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng���ID)
    If rsTemp.RecordCount = 0 Then
        MsgBox "ע��:" & vbCrLf & "    �����ε����" & IIf(mblnIsBIll, "Ʊ��", "��Ƭ") & "�Ѿ�������ɾ�������飡", vbOKOnly + vbInformation + vbDefaultButton1, gstrSysName
        Exit Function
    End If
    
    With cbo���
        blnFind = False
        For i = 0 To .ListCount - 1
            If mintƱ�� = gBillType.Ԥ���վ� Then
                 If .ItemData(i) = Val(NVL(rsTemp!ʹ�����)) + 1 Then
                    .ListIndex = i: blnFind = True: Exit For
                 End If
            ElseIf mintƱ�� = gBillType.���￨ Or mintƱ�� = gBillType.���ѿ� Then
                 If .ItemData(i) = Val(NVL(rsTemp!ʹ�����)) Then
                    .ListIndex = i: blnFind = True: Exit For
                 End If
            Else
                If .List(i) = NVL(rsTemp!ʹ�����) Then
                   .ListIndex = i: blnFind = True: Exit For
                End If
            End If
        Next
        '58071
        If blnFind = False And Not (mintƱ�� = gBillType.���￨ Or mintƱ�� = gBillType.���ѿ�) Then
            .AddItem NVL(rsTemp!ʹ�����, " ")
            .ListIndex = .NewIndex
        End If
        .Tag = NVL(rsTemp!ʹ�����, " ")
        .Enabled = NVL(rsTemp!�������) = NVL(rsTemp!ʣ������)
    End With
    
    txtEdit(mTxtIdx.idx_����).Text = NVL(rsTemp!����)
    txtEdit(mTxtIdx.idx_����).Tag = NVL(rsTemp!����)
    txtEdit(mTxtIdx.idx_��ʼǰ׺).Text = NVL(rsTemp!ǰ׺�ı�)
    lngLen = Len(Trim(txtEdit(mTxtIdx.idx_��ʼǰ׺).Text))
    txtEdit(mTxtIdx.idx_��ʼ����).Text = Mid(NVL(rsTemp!��ʼ����), lngLen + 1)
    txtEdit(mTxtIdx.idx_��ʼ����).Tag = txtEdit(mTxtIdx.idx_��ʼ����).Text
    txtEdit(mTxtIdx.idx_��ֹǰ׺).Text = NVL(rsTemp!ǰ׺�ı�)
    txtEdit(mTxtIdx.idx_��ֹ����).Text = Mid(NVL(rsTemp!��ֹ����), lngLen + 1)
    txtEdit(mTxtIdx.idx_��ֹ����).Tag = txtEdit(mTxtIdx.idx_��ֹ����).Text
    txtEdit(mTxtIdx.idx_��ע).Text = NVL(rsTemp!��ע)
    txtEdit(mTxtIdx.idx_�Ǽ���).Text = NVL(rsTemp!�Ǽ���)
    txtEdit(mTxtIdx.idx_�Ǽ�ʱ��).Text = Format(rsTemp!�Ǽ�ʱ��, "yyyy-mm-dd HH:MM:SS")
    
    '������ϸ������Ϣ
    vsMemo.Tag = Val(NVL(rsTemp!�������)) & "-" & Val(NVL(rsTemp!ʣ������))
    If mintƱ�� = gBillType.���ѿ� Then
        gstrSQL = _
            "Select A.�Ǽ�ʱ��,A.��ʼ���� As ��ʼ����,A.��ֹ���� As ��ֹ���� " & _
            "From ���ѿ����ü�¼ A " & _
            "Where A.����=[1] " & _
            "Order By �Ǽ�ʱ��"
    Else
        gstrSQL = _
            "Select A.�Ǽ�ʱ��,A.��ʼ����,A.��ֹ���� " & _
            "From Ʊ�����ü�¼ A " & _
            "Where A.����=[1] " & _
            "Order By �Ǽ�ʱ��"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng���ID)
    With rsTemp
        mstrDrawNum = "0"
        Do While Not .EOF
            If NVL(rsTemp!��ʼ����) = NVL(rsTemp!��ֹ����) Then
                mstrDrawBill = mstrDrawBill & "," & NVL(rsTemp!��ʼ����)
            Else
                mstrDrawBill = mstrDrawBill & "," & NVL(rsTemp!��ʼ����) & "-" & NVL(rsTemp!��ֹ����)
            End If
            'mstrDrawNum = mlngDrawNum + GetBillNum(Mid(Nvl(rsTemp!��ʼ����), lngLen + 1), Mid(Nvl(rsTemp!��ֹ����), lngLen + 1))
            '�����:54259
            '77390:���ϴ�,2014/9/3 09:33:32,����Ʊ������
             mstrDrawNum = zlStr.ExpressValue(GetBillNum(Mid(NVL(rsTemp!��ʼ����), lngLen + 1), Mid(NVL(rsTemp!��ֹ����), lngLen + 1)) & "+" & mstrDrawNum)
            .MoveNext
        Loop
        If mstrDrawBill <> "" Then mstrDrawBill = Mid(mstrDrawBill, 2)
    End With
    
    '������Ϣ
    If mintƱ�� = gBillType.���ѿ� Then
        gstrSQL = _
            "Select A.��ֹ���� As ��ֹ����, A.��ʼ���� As ��ʼ����,A.����ʱ��,A.���� " & _
            "From ���ѿ������¼ A " & _
            "Where ���ID=[1] " & _
            "Order By ��ʼ����,����ʱ��"
    Else
        gstrSQL = _
            "Select A.��ֹ����, A.��ʼ����,A.����ʱ��,A.���� " & _
            "From Ʊ�ݱ����¼ A " & _
            "Where ���ID=[1] " & _
            "Order By ��ʼ����,����ʱ��"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng���ID)
    mstrDamnifyBill = ""
     With rsTemp
        mlngDamnifyNum = 0
        Do While Not .EOF
            If NVL(rsTemp!��ʼ����) = NVL(rsTemp!��ֹ����) Then
                mstrDamnifyBill = mstrDamnifyBill & "," & NVL(rsTemp!��ʼ����)
            ElseIf NVL(rsTemp!��ʼ����) = "" And NVL(rsTemp!��ֹ����) <> "" Then
                mstrDamnifyBill = mstrDamnifyBill & "," & NVL(rsTemp!��ֹ����)
            ElseIf NVL(rsTemp!��ʼ����) <> "" And NVL(rsTemp!��ֹ����) = "" Then
                mstrDamnifyBill = mstrDamnifyBill & "," & NVL(rsTemp!��ʼ����)
            Else
                mstrDamnifyBill = mstrDamnifyBill & "," & NVL(rsTemp!��ʼ����) & "-" & NVL(rsTemp!��ֹ����)
            End If
            mlngDamnifyNum = mlngDamnifyNum + Val(NVL(rsTemp!����))
            .MoveNext
        Loop
        If mstrDamnifyBill <> "" Then mstrDamnifyBill = Mid(mstrDamnifyBill, 2)
    End With
    Call SetCtlEnable
    Call SetMemo
    If mintƱ�� = gBillType.���￨ Or mintƱ�� = gBillType.���ѿ� Then
        Call Setǰ׺(mcllCardProperty(cbo���.ListIndex + 1)(1))
    End If
    LoadCardData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub SetCtlEnable()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ÿؼ��ɱ༭����
    '����:���˺�
    '����:2010-11-17 16:03:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer
    Select Case mEditType
    Case Ed_����
        txtEdit(mTxtIdx.idx_��ʼǰ׺).Enabled = True
        txtEdit(mTxtIdx.idx_��ֹǰ׺).Enabled = True
        txtEdit(mTxtIdx.idx_��ʼ����).Enabled = True
        txtEdit(mTxtIdx.idx_��ֹ����).Enabled = True
        txtEdit(mTxtIdx.idx_��ע).Enabled = True
    Case Ed_�޸�
       'If mlng���� > 2 Then
            txtEdit(mTxtIdx.idx_��ʼǰ׺).Enabled = True
            txtEdit(mTxtIdx.idx_��ֹǰ׺).Enabled = True
        ' End If
        txtEdit(mTxtIdx.idx_��ʼ����).Enabled = True
        txtEdit(mTxtIdx.idx_��ֹ����).Enabled = True
        txtEdit(mTxtIdx.idx_��ע).Enabled = True
        If mstrDamnifyBill <> "" Or mstrDrawBill <> "" Then
            '���ܸ���ǰ׺
            txtEdit(mTxtIdx.idx_��ʼǰ׺).Enabled = False: txtEdit(mTxtIdx.idx_��ֹǰ׺).Enabled = False:
        End If
    Case Else
        For i = 0 To txtEdit.UBound
            txtEdit(i).Enabled = False
        Next
    End Select
End Sub


Private Sub SetMemo()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����˵����Ϣ
    '����:���˺�
    '����:2010-11-16 10:55:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngRow As Long, sngY As Single, intTittleFontSize As Integer, intTextFontSize As Integer
    Dim strTmp As String, strTemp As String, strText As String, i As Long
    Dim varTemp As Variant
    
    With vsMemo
        .Redraw = flexRDNone
        .Clear
        lngRow = 1
        '-----------------------------------------------------------------------
        '���Ʊ�ݴ���
        .Cell(flexcpFontBold, lngRow, 0, lngRow, .Cols - 1) = True  '������ʾ
        .Cell(flexcpFontSize, lngRow, 0, lngRow, .Cols - 1) = intTittleFontSize  '������ʾ
        .Cell(flexcpText, lngRow, 0, lngRow, .Cols - 1) = "���:"
        
        lngRow = lngRow + 1
        .Cell(flexcpFontBold, lngRow, 0, lngRow, .Cols - 1) = False  '
        .Cell(flexcpFontSize, lngRow, 0, lngRow, .Cols - 1) = intTextFontSize
        .Cell(flexcpText, lngRow, 1, lngRow, .Cols - 1) = IIf(mblnIsBIll, "Ʊ�ŷ�Χ:", "���ŷ�Χ:") & Trim(txtEdit(mTxtIdx.idx_��ʼǰ׺).Text) & Trim(txtEdit(mTxtIdx.idx_��ʼ����)) & "��" & Trim(txtEdit(mTxtIdx.idx_��ʼǰ׺).Text) & Trim(txtEdit(mTxtIdx.idx_��ֹ����))
        '�����:54259
        strTmp = "0"
        If mEditType = Ed_�鿴 Then
            strText = Val(Split(vsMemo.Tag & "-", "-")(0))
        Else
            strTmp = GetBillNum(Trim(txtEdit(mTxtIdx.idx_��ʼ����)), Trim(txtEdit(mTxtIdx.idx_��ֹ����)), strTemp)
            strText = strTmp
            If strTemp <> "" Then
                strText = strTemp
            End If
        End If
        
        If Not IsNumeric(strText) Then
            .Cell(flexcpForeColor, lngRow, 0, lngRow, .Cols - 1) = vbRed
        Else
            .Cell(flexcpForeColor, lngRow, 0, lngRow, .Cols - 1) = .ForeColor
        End If
        lngRow = lngRow + 1
        .Cell(flexcpFontBold, lngRow, 0, lngRow, .Cols - 1) = False  '
        .Cell(flexcpFontSize, lngRow, 0, lngRow, .Cols - 1) = intTextFontSize
        .Cell(flexcpText, lngRow, 1, lngRow, .Cols - 1) = "�������:" & strText & "��"
        
        lngRow = lngRow + 1
        If mEditType = Ed_���� Then GoTo goEnd:
        varTemp = Split(vsMemo.Tag & "-", "-")
        strText = Val(varTemp(1))
        '�����:54259
        If strTmp <> "0" Then    '�޸�ʱ,����ʣ������Ҫ�����仯
'            lngTemp = lngTemp - (Val(varTemp(0)) - Val(varTemp(1)))
'            strText = lngTemp
            '77390:���ϴ�,2014/9/3 09:33:32,����Ʊ������
            strTmp = GetBillNum(GetBillNum(varTemp(1), varTemp(0)), strTmp)
            strText = strTmp
        End If
        .Cell(flexcpForeColor, lngRow, 0, lngRow, .Cols - 1) = .ForeColor
        .Cell(flexcpFontBold, lngRow, 0, lngRow, .Cols - 1) = False  '
        .Cell(flexcpFontSize, lngRow, 0, lngRow, .Cols - 1) = intTextFontSize
        .Cell(flexcpText, lngRow, 1, lngRow, .Cols - 1) = "ʣ������:" & strText & "��"
        If Val(strText) < 0 Then
            .Cell(flexcpForeColor, lngRow, 0, lngRow, .Cols - 1) = vbRed
        Else
            .Cell(flexcpForeColor, lngRow, 0, lngRow, .Cols - 1) = .ForeColor
        End If
        '-----------------------------------------------------------------------
        '2.����Ʊ�ݴ���
        lngRow = lngRow + 1
        .Cell(flexcpFontBold, lngRow, 0, lngRow, .Cols - 1) = True  '������ʾ
        .Cell(flexcpFontSize, lngRow, 0, lngRow, .Cols - 1) = intTittleFontSize  '������ʾ
        .Cell(flexcpText, lngRow, 0, lngRow, .Cols - 1) = "����:"
        
        lngRow = lngRow + 1
        .Cell(flexcpFontBold, lngRow, 0, lngRow, .Cols - 1) = False  '
        .Cell(flexcpFontSize, lngRow, 0, lngRow, .Cols - 1) = intTextFontSize
        .Cell(flexcpText, lngRow, 1, lngRow, .Cols - 1) = IIf(mblnIsBIll, "����Ʊ��:", "���ÿ���:") & mstrDrawBill
        
        lngRow = lngRow + 1
        .Cell(flexcpFontBold, lngRow, 0, lngRow, .Cols - 1) = False  '
        .Cell(flexcpFontSize, lngRow, 0, lngRow, .Cols - 1) = intTextFontSize
        .Cell(flexcpText, lngRow, 1, lngRow, .Cols - 1) = "��������:" & mstrDrawNum & "��" '�����:54259
              
      '-----------------------------------------------------------------------
        '3.����Ʊ�ݴ���
        lngRow = lngRow + 1
        .Cell(flexcpFontBold, lngRow, 0, lngRow, .Cols - 1) = True  '������ʾ
        .Cell(flexcpFontSize, lngRow, 0, lngRow, .Cols - 1) = intTittleFontSize  '������ʾ
        .Cell(flexcpText, lngRow, 0, lngRow, .Cols - 1) = "����:"
        
        lngRow = lngRow + 1
        .Cell(flexcpFontBold, lngRow, 0, lngRow, .Cols - 1) = False
        .Cell(flexcpFontSize, lngRow, 0, lngRow, .Cols - 1) = intTextFontSize
        .Cell(flexcpText, lngRow, 1, lngRow, .Cols - 1) = IIf(mblnIsBIll, "����Ʊ��:", "���𿨺�:") & mstrDamnifyBill
        
        lngRow = lngRow + 1
        .Cell(flexcpFontBold, lngRow, 0, lngRow, .Cols - 1) = False  '
        .Cell(flexcpFontSize, lngRow, 0, lngRow, .Cols - 1) = intTextFontSize
        .Cell(flexcpText, lngRow, 1, lngRow, .Cols - 1) = "��������:" & mlngDamnifyNum & "��"
goEnd:
        .AutoSizeMode = flexAutoSizeRowHeight
        .AutoSize 1, 1
        .MergeCells = flexMergeFree
        For i = 0 To .Rows - 1
            .MergeRow(i) = True
        Next
        For i = 0 To .Cols - 1
            .MergeCol(i) = True
        Next
        .Redraw = flexRDBuffered
    End With
End Sub

Private Function GetBillNum(ByVal str��ʼ���� As String, ByVal str�տ����� As String, Optional ByRef strErrMsg As String = "") As String
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
'    If (str��ʼ���� <> "" And str�տ����� = "") Or (str��ʼ���� = "" And str�տ����� <> "") Then
'        GetBillNum = 1: Exit Function
'    End If
'    GetBillNum = CDec(str�տ�����) - CDec(str��ʼ����) + 1

    GetBillNum = zlStr.ExpressValue(str�տ����� & "-" & str��ʼ����) + 1
    Exit Function
errHandle:
    strErrMsg = "�������򳬳��˼��㷶Χ"
    GetBillNum = "0"
End Function


Private Sub ClearData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ؼ�����
    '����:���˺�
    '����:2010-11-16 10:35:55
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer
    mstrDrawBill = "": mstrDamnifyBill = ""

    For i = 0 To txtEdit.UBound
        txtEdit(i).Text = ""
        If i = mTxtIdx.idx_��ʼǰ׺ Or i = mTxtIdx.idx_��ֹǰ׺ Then
            Call txtEdit_Change(i)  '����:38021
        End If
        If txtEdit(i).Enabled = False Then
            txtEdit(i).BackColor = Me.BackColor
        Else
            txtEdit(i).BackColor = &H80000005
        End If
    Next
    
    vsMemo.Clear
    vsMemo.Rows = 11
End Sub

Private Sub cbo���_Click()
    If mintƱ�� = gBillType.���￨ Or mintƱ�� = gBillType.���ѿ� Then
        mlng���� = mcllCardProperty(cbo���.ListIndex + 1)(0)
        Call Setǰ׺(mcllCardProperty(cbo���.ListIndex + 1)(1))
        If mlng���� < 3 Then
            txtEdit(mTxtIdx.idx_��ʼǰ׺).Text = "": txtEdit(mTxtIdx.idx_��ʼǰ׺).Enabled = False
            txtEdit(mTxtIdx.idx_��ֹǰ׺).Enabled = False
        End If
        txtEdit(mTxtIdx.idx_��ʼ����).MaxLength = mlng���� - zlCommFun.ActualLen(txtEdit(mTxtIdx.idx_��ʼǰ׺).Text)
        txtEdit(mTxtIdx.idx_��ֹ����).MaxLength = txtEdit(mTxtIdx.idx_��ʼ����).MaxLength
    End If
End Sub

Private Sub Setǰ׺(Optional ByVal strǰ׺ As String = "")
    Me.txtEdit(mTxtIdx.idx_��ʼǰ׺).Enabled = strǰ׺ <> ""
    Me.txtEdit(mTxtIdx.idx_��ֹǰ׺).Enabled = Me.txtEdit(mTxtIdx.idx_��ʼǰ׺).Enabled
    Me.txtEdit(mTxtIdx.idx_��ʼǰ׺).BackColor = Me.txtEdit(mTxtIdx.idx_��ʼ����).BackColor
    Me.txtEdit(mTxtIdx.idx_��ֹǰ׺).BackColor = Me.txtEdit(mTxtIdx.idx_��ʼ����).BackColor
    If strǰ׺ = "" And mlng���� > 2 Then Exit Sub
    Me.txtEdit(mTxtIdx.idx_��ʼǰ׺).Text = UCase(strǰ׺)
    Me.txtEdit(mTxtIdx.idx_��ʼǰ׺).BackColor = Me.BackColor
    Me.txtEdit(mTxtIdx.idx_��ֹǰ׺).BackColor = Me.BackColor
End Sub

Private Sub cmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    If LoadCombox = False Then Unload Me: Exit Sub
    mblnFirst = False
    Call SetCtlEnable
    If LoadCardData = False Then Unload Me: Exit Sub
    If zlControl.IsCtrlSetFocus(txtEdit(mTxtIdx.idx_��ʼǰ׺)) Then
        txtEdit(mTxtIdx.idx_��ʼǰ׺).SetFocus
    Else
        zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_��ʼ����)
    End If
    
    mblnChange = False
End Sub

Private Sub Form_Load()
    mblnFirst = True
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

Private Function isValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������������Ƿ�Ϸ�
    '����:�Ϸ�����true,���򷵻�False
    '����:���˺�
    '����:2010-11-16 15:04:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str��С���� As String, str������ As String, varTemp As Variant, varData As Variant
    Dim str��ʼ���� As String, str�������� As String, i As Long, strTemp As String
    Dim str��� As String, str������� As String
    Dim rsTemp As ADODB.Recordset
    Dim byt�������� As Byte, blnDefult As Boolean
    Dim strName1 As String, strName2 As String
    
    On Error GoTo errHandle
    strName1 = IIf(mblnIsBIll, "Ʊ��", "��Ƭ")
    strName2 = IIf(mblnIsBIll, "����", "����")
    '�����:54259
    If Len(GetBillNum(Trim(txtEdit(mTxtIdx.idx_��ʼ����)), Trim(txtEdit(mTxtIdx.idx_��ֹ����)))) > 25 Then
        ShowMsgbox "ע��" & vbCrLf & "    ���" & strName1 & "����λ�����ó���" & 25 & "λ!"
        Exit Function
    End If
    
    If zlCommFun.ActualLen(Trim(txtEdit(mTxtIdx.idx_��ע))) > 200 Then
        ShowMsgbox "ע��" & vbCrLf & "    ��ע���ֻ������200���ַ���100������,����!"
        zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_��ע): Exit Function
    End If
    If zlCommFun.ActualLen(Trim(txtEdit(mTxtIdx.idx_��ʼǰ׺))) > 2 Then
        ShowMsgbox "ע��" & vbCrLf & "   " & strName2 & "ǰ׺���ֻ������2���ַ���1������,����!"
        zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_��ʼǰ׺): Exit Function
    End If
    If InStr(1, txtEdit(mTxtIdx.idx_��ע), "'") > 0 Then
        ShowMsgbox "ע��" & vbCrLf & "    ��ע�к��зǷ��ַ�������,����!"
        zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_��ע): Exit Function
    End If
    If Trim(txtEdit(mTxtIdx.idx_��ʼ����).Text) = "" Then
        ShowMsgbox "ע��" & vbCrLf & "    " & strName2 & "��Χ�еĿ�ʼ" & strName2 & "��������,����!"
        zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_��ʼ����): Exit Function
    End If
    If Trim(txtEdit(mTxtIdx.idx_��ֹ����).Text) = "" Then
        ShowMsgbox "ע��" & vbCrLf & "    " & strName2 & "��Χ�еĽ���" & strName2 & "��������,����!"
        zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_��ֹ����): Exit Function
    End If
    If Not IsNumeric(txtEdit(mTxtIdx.idx_��ʼ����).Text) Then
        ShowMsgbox "ע��" & vbCrLf & "    " & strName2 & "��Χ�еĿ�ʼ" & strName2 & "������������,����!"
        zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_��ʼ����): Exit Function
    End If
    If Not IsNumeric(txtEdit(mTxtIdx.idx_��ֹ����).Text) Then
        ShowMsgbox "ע��" & vbCrLf & "    " & strName2 & "��Χ�еĽ���" & strName2 & "������������,����!"
        zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_��ֹ����): Exit Function
    End If
    '103428:���ϴ���2017/2/15����鿨�ų���
    If zlCommFun.ActualLen(txtEdit(mTxtIdx.idx_��ʼǰ׺) & txtEdit(mTxtIdx.idx_��ʼ����).Text) <> mlng���� Then
        If mintƱ�� = gBillType.���￨ Or mintƱ�� = gBillType.���ѿ� Then
            byt�������� = mcllCardProperty(cbo���.ListIndex + 1)(3)
            Select Case byt��������
                Case 0
                    ShowMsgbox "ע��" & vbCrLf & "    ���ŷ�Χ�еĿ�ʼ���ų��Ȳ���(ӦΪ" & mlng���� & "λ),����!"
                Case 2
                    ShowMsgbox "ע��" & vbCrLf & "    ���ŷ�Χ�еĿ�ʼ���ų���δ�ﵽ���λ��,�Ƿ������", True, blnDefult
                    If Not blnDefult Then byt�������� = 0
            End Select
        Else
            ShowMsgbox "ע��" & vbCrLf & "    ���뷶Χ�еĿ�ʼ���볤�Ȳ���(ӦΪ" & mlng���� & "λ),����!"
            byt�������� = 0
        End If
        If byt�������� = 0 Then
            zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_��ʼ����): Exit Function
        End If
    End If
    If zlCommFun.ActualLen(txtEdit(mTxtIdx.idx_��ֹǰ׺) & txtEdit(mTxtIdx.idx_��ֹ����).Text) <> zlCommFun.ActualLen(txtEdit(mTxtIdx.idx_��ʼǰ׺) & txtEdit(mTxtIdx.idx_��ʼ����).Text) Then
        ShowMsgbox "ע��" & vbCrLf & "    " & strName2 & "��Χ�еĽ���" & strName2 & "�뿪ʼ" & strName2 & "�ĳ��Ȳ�һ��,����!"
        zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_��ֹ����): Exit Function
    End If
    If txtEdit(mTxtIdx.idx_��ֹ����).Text < txtEdit(mTxtIdx.idx_��ʼ����) Then
        ShowMsgbox "ע��" & vbCrLf & "    " & strName2 & "��Χ�еĽ���" & strName2 & "С���˿�ʼ" & strName2 & ",����!"
        zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_��ֹ����): Exit Function
    End If
    
    If zlIsOnlyNum(Trim(txtEdit(mTxtIdx.idx_��ʼ����))) = False Then
        MsgBox "��ʼ" & strName2 & "�к��з������ַ�����ĸֻ����Ϊǰ׺��", vbExclamation, gstrSysName
        zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_��ʼ����): Exit Function
    End If
    
    If zlIsOnlyNum(Trim(txtEdit(mTxtIdx.idx_��ֹ����))) = False Then
        MsgBox "��ֹ" & strName2 & "�к��з������ַ�����ĸֻ����Ϊǰ׺��", vbExclamation, gstrSysName
        zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_��ֹ����): Exit Function
    End If
    
    '105916:���ϴ�,2017/4/10�����Ʊ�ݺ���
    If txtEdit(mTxtIdx.idx_��ʼ����).Text = String(mlng����, "0") And txtEdit(mTxtIdx.idx_��ֹ����).Text = String(mlng����, "9") Then
        MsgBox "����ʹ��" & String(mlng����, "0") & "-" & String(mlng����, "9") & "��" & strName2 & "��Χ��", vbExclamation, gstrSysName
        zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_��ֹ����): Exit Function
    End If
    
    str������� = GetBillNum(Trim(txtEdit(mTxtIdx.idx_��ʼ����)), Trim(txtEdit(mTxtIdx.idx_��ֹ����)))
    If InStr(str�������, "E") > 0 Or Len(str�������) > 10 Then '����̫���Ѿ���ɿ�ѧ���㷨
        MsgBox "����" & strName1 & "�������ܳ��� 10000000000 �ţ��������⡣", vbExclamation, gstrSysName
        zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_��ʼ����): Exit Function
    End If
    
    '����Ƿ��Ѿ�ʹ�ù�,ʹ�ù���Ʊ�ݲ��ܸı��䳤��
    If mEditType = Ed_�޸� And (mstrDrawBill <> "" Or mstrDamnifyBill <> "") Then
        If Len(txtEdit(mTxtIdx.idx_��ʼ����).Text) <> Len(txtEdit(mTxtIdx.idx_��ʼ����).Tag) Then
            MsgBox "��������" & strName1 & "�Ѿ���ʹ�ù�," & strName2 & "���Ȳ��ܸı�," & vbCrLf & _
                strName2 & "����Ӧ����" & Len(txtEdit(mTxtIdx.idx_��ʼǰ׺).Text & txtEdit(mTxtIdx.idx_��ʼ����).Tag) & "λ��", vbExclamation, gstrSysName
            zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_��ֹ����): Exit Function
        End If
        
        varData = Split(mstrDrawBill, ",")
        For i = 0 To UBound(varData)
            If InStr(varData(i), "-") > 0 Then
                varTemp = Split(varData(i), "-")
                If str��С���� = "" Or str��С���� > varTemp(0) Then
                    str��С���� = varTemp(0)
                End If
                If str������ = "" Or str������ < varTemp(1) Then
                    str������ = varTemp(1)
                End If
            Else
                If str��С���� = "" Or str��С���� > varData(i) Then
                    str��С���� = varData(i)
                End If
                If str������ = "" Or str������ < varData(i) Then
                    str������ = varData(i)
                End If
            End If
        Next
        varData = Split(mstrDamnifyBill, ",")
        For i = 0 To UBound(varData)
            If InStr(varData(i), "-") > 0 Then
                varTemp = Split(varData(i), "-")
                If str��С���� = "" Or str��С���� > varTemp(0) Then
                    str��С���� = varTemp(0)
                End If
                If str������ = "" Or str������ < varTemp(1) Then
                    str������ = varTemp(1)
                End If
            Else
                If str��С���� = "" Or str��С���� > varData(i) Then
                    str��С���� = varData(i)
                End If
                If str������ = "" Or str������ < varData(i) Then
                    str������ = varData(i)
                End If
            End If
        Next
        
        If txtEdit(mTxtIdx.idx_��ʼǰ׺).Text & txtEdit(mTxtIdx.idx_��ʼ����).Text > str��С���� Then
            MsgBox "��������" & strName1 & "�Ѿ�ʹ�ã�" & vbCrLf & "��ʼ" & strName2 & "ֻ����С��" & str��С���� & "��", vbExclamation, gstrSysName
            zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_��ʼ����): Exit Function
        End If
        If txtEdit(mTxtIdx.idx_��ֹǰ׺).Text & txtEdit(mTxtIdx.idx_��ֹ����).Text < str������ Then
            MsgBox "��������" & strName1 & "�Ѿ�ʹ�ã�" & vbCrLf & strName2 & "�Ѿ��õ�" & str������ & ",��ֹ" & strName2 & "�����������", vbExclamation, gstrSysName
            zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_��ֹ����): Exit Function
        End If
    End If
    
    '����Ƿ���ʹ�����
    If mintƱ�� = 1 Or mintƱ�� = 3 Then
        If cbo���.ListIndex < 0 Then
            MsgBox "ע��:" & vbCrLf & "    ʹ�����û��ѡ��,��ѡ��!", vbInformation + vbOKOnly, gstrSysName
            zlControl.ControlSetFocus cbo���: Exit Function
            Exit Function
        End If
    End If
    
    str��� = Getʹ�����(mintƱ��)
    
    '��������Ƿ��ظ�
    If mEditType = Ed_���� Then
        If Trim(txtEdit(idx_����).Text) <> "" Then
            If batchValied = False Then
                ShowMsgbox "�����뱾�����������ͬ������¼�����ܼ��������飡"
                Exit Function
            End If
        End If
    ElseIf mEditType = Ed_�޸� Then
        If Trim(txtEdit(idx_����).Text) = "" Then
            ShowMsgbox "δ����������Σ����ܼ��������飡"
            Exit Function
        ElseIf Trim(txtEdit(mTxtIdx.idx_����).Tag) <> Trim(txtEdit(idx_����).Text) Then
            If batchValied = False Then
                ShowMsgbox "�����뱾�����������ͬ������¼�����ܼ��������飡"
                Exit Function
            End If
        End If
    End If
    
    '����Ƿ��Ѿ����ò���ʹ������뵱ǰ�޸ĵĲ�һ��ʱ
    If mEditType = Ed_�޸� And str��� <> Trim(cbo���.Tag) Then
        If mintƱ�� = gBillType.���ѿ� Then
            gstrSQL = _
                "Select b.���� As ʹ����� " & _
                "From ���ѿ����ü�¼ A,���ѿ����Ŀ¼ B " & _
                "Where Nvl(a.�ӿڱ��,0)=b.���(+) And a.����=[1] And Nvl(a.�ӿڱ��,0)<>Nvl([3],0) And Nvl(a.ʣ������,0) >0 And Rownum < 2 "
        ElseIf mintƱ�� = gBillType.���￨ Then
            gstrSQL = _
                "Select b.���� As ʹ����� " & _
                "From Ʊ�����ü�¼ A,ҽ�ƿ���� B " & _
                "Where To_Number(Nvl(a.ʹ�����,0))=b.ID(+) And a.����=[1] And a.Ʊ��=[2] " & _
                "      And Nvl(a.ʹ�����,'LXH')<>Nvl([3],'LXH') And Nvl(a.ʣ������,0) >0 And Rownum < 2 "
        Else
            gstrSQL = _
                "Select " & IIf(mintƱ�� = gBillType.Ԥ���վ�, "Decode(ʹ�����,'2','סԺԤ��','����Ԥ��') As ʹ����� ", "ʹ����� ") & _
                "From Ʊ�����ü�¼ " & _
                "Where ����=[1] And Ʊ��=[2] And Nvl(ʹ�����,'LXH')<>Nvl([3],'LXH') And Nvl(ʣ������,0) >0 And Rownum < 2 "
        End If
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng���ID, mintƱ��, str���)
        If rsTemp.EOF = False Then
            If rsTemp.RecordCount >= 0 Then
                If MsgBox("ע��:" & vbCrLf & "     �㽫ԭʹ�����Ϊ��" & IIf(Trim(NVL(rsTemp!ʹ�����)) = "", "���������", NVL(rsTemp!ʹ�����)) & "������Ϊ" & vbCrLf & _
                                  "    ��" & IIf(Trim(cbo���.Text) = "", "���������", cbo���.Text) & "��������¼�Ѿ�������, " & vbCrLf & _
                                  "     �Ƿ����õ�" & strName1 & "һ�����? ", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    zlControl.ControlSetFocus cbo���: Exit Function
                End If
            End If
        End If
    End If
    
    '�ж�����Ƿ��ظ�
    str��ʼ���� = Trim(txtEdit(mTxtIdx.idx_��ʼǰ׺).Text) & Trim(txtEdit(mTxtIdx.idx_��ʼ����).Text)
    str�������� = Trim(txtEdit(mTxtIdx.idx_��ֹǰ׺).Text) & Trim(txtEdit(mTxtIdx.idx_��ֹ����).Text)
    '115348:���ϴ���2017/10/24��ҽ�ƿ�Ҫ������飬��ͬҽ�ƿ����ܿ��Ż����ظ�
    If mintƱ�� = gBillType.���ѿ� Then
        gstrSQL = _
            "Select ID,nvl(ʣ������,0) ʣ������ " & _
            "From ���ѿ�����¼ " & _
            "Where ID<>[3] And Nvl(�ӿڱ��,0)=Nvl([5],0) " & _
            "      And (([1] Between ��ʼ���� And ��ֹ����) Or ([2] Between ��ʼ���� And ��ֹ����))  " & _
            "      And Length(��ʼ����)=Length([1]) And ����=[6]"
    Else
        gstrSQL = _
            "Select ID,nvl(ʣ������,0) ʣ������ " & _
            "From Ʊ������¼ " & _
            "Where ID<>[3] And Ʊ��=[4] And nvl(ʹ�����,'LXH')=nvl([5],'LXH') " & _
            "      And (([1] between ��ʼ���� and  ��ֹ����) or  ([2] between ��ʼ����  and ��ֹ����)) " & _
            "      And Length(��ʼ����)=Length([1]) And ����=[6]"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, str��ʼ����, str��������, mlng���ID, mintƱ��, str���, Trim(txtEdit(idx_����).Text))
    If rsTemp.RecordCount > 0 Then
        If mblnIsBIll Then
            If MsgBox("�����뱾������ص�������¼" & IIf(Val(NVL(rsTemp!ʣ������)) > 0, "�����һ���δʹ�����Ʊ�ݡ�", "��") & vbCrLf & _
                "�㻹��Ҫ������", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
        Else
            ShowMsgbox "�����뱾����⿨���ص�������¼�����ܼ��������飡"
            Exit Function
        End If
    End If
    
    '102996:���ϴ�,2016/11/23,ҽ�Ʒ�Ʊ���ӻ�����
    If (mEditType = Ed_���� Or mEditType = Ed_�޸�) And gblnBillPrint Then
        Dim strErrMsg As String, strExpended As String
        On Error Resume Next
        If gobjBillPrint.zlBillInCheckValied(mEditType + 1, mintƱ��, str���, mlng���ID, str��ʼ����, str��������, strExpended) = False Then
            zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_��ʼ����): Exit Function
        End If
        Err = 0: On Error GoTo errHandle
    End If
    
    isValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function batchValied() As Boolean
    '��������Ƿ��ظ�
    Dim str�ӿڱ�� As String, rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    
    If mintƱ�� = gBillType.���ѿ� Then
        str�ӿڱ�� = Getʹ�����(mintƱ��)
        gstrSQL = "Select 1 From ���ѿ�����¼ Where �ӿڱ�� = [1] And ���� = [2] And ID <> ���� And Rownum < 2"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, str�ӿڱ��, Trim(txtEdit(idx_����).Text))
    Else
        gstrSQL = "Select 1 From Ʊ������¼ Where Ʊ�� = [1] And ���� = [2] And ID <> ���� And Rownum < 2"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mintƱ��, Trim(txtEdit(idx_����).Text))
    End If
    If rsTemp.EOF = False Then
        Exit Function
    End If
    batchValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Getʹ�����(ByVal intƱ�� As gBillType) As String
    '��ȡʹ�����
    Dim str��� As String
    
    On Error GoTo errHandle
    Select Case intƱ��
    Case gBillType.�շ��վ�, gBillType.�����վ�
        '�շѺͽ���
        str��� = Trim(cbo���.Text)
    Case gBillType.Ԥ���վ�
        str��� = cbo���.ItemData(cbo���.ListIndex) - 1
        If Val(str���) = 0 Then str��� = ""
    Case gBillType.���￨, gBillType.���ѿ�
        str��� = cbo���.ItemData(cbo���.ListIndex)
        If Val(str���) = 0 Then str��� = ""
    Case Else
        str��� = ""
    End Select
    Getʹ����� = str���
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function SaveData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������
    '����:���ݱ���ɹ�,����true,���򷵻�ΪFalse
    '����:���˺�
    '����:2010-11-16 15:15:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '�����:54259
    Dim lngID As Long, str������� As String, strʣ������ As String
    Dim varTemp As Variant, str��� As String
    Dim blnTrans As Boolean, strErrMsg As String, strExpended As String
    
    On Error GoTo errHandle
    
    str������� = GetBillNum(Trim(txtEdit(mTxtIdx.idx_��ʼ����)), Trim(txtEdit(mTxtIdx.idx_��ֹ����)))
    strʣ������ = "0"
    If Len(str�������) <= 0 Then
       If Val(str�������) <= 0 Then Exit Function
    End If
    str��� = Getʹ�����(mintƱ��)
        
    If mEditType = Ed_���� Then
        If mintƱ�� = gBillType.���ѿ� Then
            lngID = zlDatabase.GetNextId("���ѿ�����¼")
        Else
            lngID = zlDatabase.GetNextId("Ʊ������¼")
        End If
        strʣ������ = str�������
    Else
        lngID = mlng���ID
        '77390:���ϴ�,2014/9/3 09:33:32,����Ʊ������
        varTemp = Split(vsMemo.Tag & "-", "-")
        strʣ������ = GetBillNum(varTemp(1), varTemp(0))
        If Val(strʣ������) < 0 Then strʣ������ = "0"
        
        strʣ������ = GetBillNum(strʣ������, str�������)
        If Val(strʣ������) < 0 Then Exit Function
    End If
    
    If mintƱ�� = gBillType.���ѿ� Then
        ' Zl_���ѿ�����¼_Insert
        gstrSQL = "Zl_���ѿ�����¼_Insert("
        '  Id_In       In ���ѿ�����¼.ID%Type,
        gstrSQL = gstrSQL & "" & lngID & ","
        '  �ӿڱ��_In In ���ѿ�����¼.�ӿڱ��%Type,
        gstrSQL = gstrSQL & "" & Val(str���) & ","
        '  ǰ׺�ı�_In In ���ѿ�����¼.ǰ׺�ı�%Type,
        gstrSQL = gstrSQL & "'" & Trim(txtEdit(mTxtIdx.idx_��ʼǰ׺)) & "',"
        '  ��ʼ����_In In ���ѿ�����¼.��ʼ����%Type,
        gstrSQL = gstrSQL & "'" & Trim(txtEdit(mTxtIdx.idx_��ʼǰ׺)) & Trim(txtEdit(mTxtIdx.idx_��ʼ����)) & "',"
        '  ��ֹ����_In In ���ѿ�����¼.��ֹ����%Type,
        gstrSQL = gstrSQL & "'" & Trim(txtEdit(mTxtIdx.idx_��ֹǰ׺)) & Trim(txtEdit(mTxtIdx.idx_��ֹ����)) & "',"
        '  �������_In In ���ѿ�����¼.�������%Type,
        gstrSQL = gstrSQL & "'" & str������� & "',"
        '  ʣ������_In In ���ѿ�����¼.ʣ������%Type,
        gstrSQL = gstrSQL & "'" & strʣ������ & "',"
        '  ��ע_In     In ���ѿ�����¼.��ע%Type,
        gstrSQL = gstrSQL & "'" & Trim(txtEdit(mTxtIdx.idx_��ע)) & "',"
        '  �Ǽ���_In   In ���ѿ�����¼.�Ǽ���%Type,
        gstrSQL = gstrSQL & "'" & UserInfo.���� & "',"
        '  �޸ı�־_In Integer:=0
        gstrSQL = gstrSQL & "" & IIf(mEditType = Ed_����, 0, 1) & ","
        '  ����_In     In Ʊ������¼.����%Type := Null
        gstrSQL = gstrSQL & "" & IIf(Trim(txtEdit(idx_����).Text) = "", lngID, Trim(txtEdit(idx_����).Text)) & ")"
    Else
        ' Zl_Ʊ������¼_Insert
        gstrSQL = "Zl_Ʊ������¼_Insert("
        '  Id_In       In Ʊ������¼.ID%Type,
        gstrSQL = gstrSQL & "" & lngID & ","
        '  Ʊ��_In     In Ʊ������¼.Ʊ��%Type,
        gstrSQL = gstrSQL & "" & mintƱ�� & ","
        '  ʹ�����_In In Ʊ������¼.ʹ�����%Type,
        gstrSQL = gstrSQL & "" & IIf(str��� = "", "NULL", "'" & str��� & "'") & ","
        '  ǰ׺�ı�_In In Ʊ������¼.ǰ׺�ı�%Type,
        gstrSQL = gstrSQL & "'" & Trim(txtEdit(mTxtIdx.idx_��ʼǰ׺)) & "',"
        '  ��ʼ����_In In Ʊ������¼.��ʼ����%Type,
        gstrSQL = gstrSQL & "'" & Trim(txtEdit(mTxtIdx.idx_��ʼǰ׺)) & Trim(txtEdit(mTxtIdx.idx_��ʼ����)) & "',"
        '  ��ֹ����_In In Ʊ������¼.��ֹ����%Type,
        gstrSQL = gstrSQL & "'" & Trim(txtEdit(mTxtIdx.idx_��ֹǰ׺)) & Trim(txtEdit(mTxtIdx.idx_��ֹ����)) & "',"
        '  �������_In In Ʊ������¼.�������%Type,
        gstrSQL = gstrSQL & "'" & str������� & "',"
        '  ʣ������_In In Ʊ������¼.ʣ������%Type,
        gstrSQL = gstrSQL & "'" & strʣ������ & "',"
        '  ��ע_In     In Ʊ������¼.��ע%Type,
        gstrSQL = gstrSQL & "'" & Trim(txtEdit(mTxtIdx.idx_��ע)) & "',"
        '  �Ǽ���_In   In Ʊ������¼.�Ǽ���%Type,
        gstrSQL = gstrSQL & "'" & UserInfo.���� & "',"
        '  �޸ı�־_In Integer:=0
        gstrSQL = gstrSQL & "" & IIf(mEditType = Ed_����, 0, 1) & ","
         '  ����_In     In Ʊ������¼.����%Type := Null
        gstrSQL = gstrSQL & "" & IIf(Trim(txtEdit(idx_����).Text) = "", lngID, Trim(txtEdit(idx_����).Text)) & ")"
    End If
    
    '102996:���ϴ�,2016/11/23,ҽ�Ʒ�Ʊ���ӻ�����
    gcnOracle.BeginTrans: blnTrans = True
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    
    If gblnBillPrint Then
        On Error Resume Next
        If gobjBillPrint.zlBillIn(mEditType + 1, mintƱ��, str���, lngID, strExpended) = False Then
            gcnOracle.RollbackTrans: blnTrans = False
            zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_��ʼ����): Exit Function
        End If
        Err = 0: On Error GoTo errHandle
    End If
    gcnOracle.CommitTrans: blnTrans = False
    SaveData = True
    Exit Function
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
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
    If mEditType = Ed_�޸� Then
        mblnChange = False
        Unload Me
        Exit Sub
    End If
    Call ClearData: mblnChange = False
    zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_��ʼǰ׺)
End Sub

Private Sub txtEdit_Change(Index As Integer)
    If mEditType = Ed_�鿴 Then Exit Sub
    
    mblnChange = True
    If Index = mTxtIdx.idx_��ʼǰ׺ And txtEdit(mTxtIdx.idx_��ʼǰ׺).Text <> txtEdit(mTxtIdx.idx_��ֹǰ׺).Text Then
        txtEdit(mTxtIdx.idx_��ֹǰ׺).Text = txtEdit(mTxtIdx.idx_��ʼǰ׺).Text
    End If
    If Index = mTxtIdx.idx_��ֹǰ׺ And txtEdit(mTxtIdx.idx_��ʼǰ׺).Text <> txtEdit(mTxtIdx.idx_��ֹǰ׺).Text Then
        txtEdit(mTxtIdx.idx_��ʼǰ׺).Text = txtEdit(mTxtIdx.idx_��ֹǰ׺).Text
    End If
    If Index = mTxtIdx.idx_��ʼǰ׺ Or Index = mTxtIdx.idx_��ֹǰ׺ Then
        txtEdit(mTxtIdx.idx_��ʼ����).MaxLength = mlng���� - zlCommFun.ActualLen(txtEdit(mTxtIdx.idx_��ʼǰ׺).Text)
        txtEdit(mTxtIdx.idx_��ֹ����).MaxLength = txtEdit(mTxtIdx.idx_��ʼ����).MaxLength
    End If
    If Index = mTxtIdx.idx_��ʼ���� Or Index = mTxtIdx.idx_��ֹ���� Then
        Call SetMemo
    End If
End Sub
Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtEdit(Index)
    If idx_��ע = Index Then
        zlCommFun.OpenIme True
    Else
        zlCommFun.OpenIme False
    End If
End Sub

Private Sub txtEdit_LostFocus(Index As Integer)
    If Index = mTxtIdx.idx_��ʼǰ׺ Or Index = mTxtIdx.idx_��ֹǰ׺ Then
        txtEdit(Index).Text = UCase(txtEdit(Index).Text)
    End If
    txtEdit(Index).Text = Trim(txtEdit(Index).Text)
    If idx_��ע = Index Then zlCommFun.OpenIme False
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab: Exit Sub
    If Index = mTxtIdx.idx_��ʼǰ׺ Or Index = mTxtIdx.idx_��ֹǰ׺ Then
        If KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Then
            KeyAscii = KeyAscii - Asc("a") + Asc("A")
        End If
    ElseIf Index = mTxtIdx.idx_���� Then
        If InStr("'[]����������,.'�ۣ�", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    Else
        If Index <> mTxtIdx.idx_��ע Then
            If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack Then
                KeyAscii = 0
            End If
        Else
            zlControl.TxtCheckKeyPress txtEdit(Index), KeyAscii, m�ı�ʽ
        End If
    End If
End Sub

Private Function LoadCombox() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����Combox����
    '����:���سɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-04-27 10:22:29
    '����:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim str��� As String
    
    On Error GoTo errHandle
    str��� = mstrPreType(mintƱ��)
    lblUserType.Caption = "ʹ�����(&L)"
    lbl(6).Caption = "���뷶Χ(&B)"
    Select Case mintƱ��
    Case gBillType.�շ��վ�, gBillType.�����վ�
        strSQL = "Select ����,����,����,ȱʡ��־ From Ʊ��ʹ����� "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        With cbo���
            .Clear
            Do While Not rsTemp.EOF
                .AddItem NVL(rsTemp!����)
                .ItemData(.NewIndex) = 1
                If Val(NVL(rsTemp!ȱʡ��־)) = 1 And .ListIndex < 0 Then .ListIndex = .NewIndex
                If mstr��� = NVL(rsTemp!����) Then .ListIndex = .NewIndex
                rsTemp.MoveNext
            Loop
            .AddItem " "    '��������Ϊ��
            If mstr��� = "" Then .ListIndex = .NewIndex
            If .ListIndex < 0 Then .ListIndex = 0
            .Visible = True: lblUserType.Visible = True
        End With
  Case gBillType.Ԥ���վ�
        mblnNotClick = True
        With cbo���
            .Clear
            If InStr(1, mstrPrivs, ";Ԥ������Ʊ��;") > 0 Then
                .AddItem "����Ԥ��": .ItemData(.NewIndex) = 2
                If Val(str���) = 2 Then .ListIndex = .NewIndex
            End If
            If InStr(1, mstrPrivs, ";Ԥ��סԺƱ��;") > 0 Then
                .AddItem "סԺԤ��": .ItemData(.NewIndex) = 3
                If Val(str���) = 3 Then .ListIndex = .NewIndex
            End If
            '58071
            If InStr(1, mstrPrivs, ";Ԥ��סԺƱ��;") > 0 And InStr(1, mstrPrivs, ";Ԥ������Ʊ��;") > 0 Then
                .AddItem " "
                .ItemData(.NewIndex) = 1
            End If
            If .ListIndex < 0 And .ListCount > 0 Then .ListIndex = 0
        End With
        mblnNotClick = False
    Case gBillType.���￨
        '103428:���ϴ���2017/2/15����鿨�ų���
        strSQL = "Select ID,����,����,ȱʡ��־,���ų���,��������,ǰ׺�ı�,�������� From ҽ�ƿ���� where nvl(�Ƿ�����,0) >=1 Order by ���� "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        mblnNotClick = True
        Set mcllCardProperty = New Collection
        With cbo���
            .Clear
            Do While Not rsTemp.EOF
                .AddItem NVL(rsTemp!����) & "-" & NVL(rsTemp!����)
                .ItemData(.NewIndex) = Val(NVL(rsTemp!ID))
                mcllCardProperty.Add Array(Val(NVL(rsTemp!���ų���)), CStr(NVL(rsTemp!ǰ׺�ı�)), CStr(NVL(rsTemp!��������)), Val(NVL(rsTemp!��������))), "K" & Val(NVL(rsTemp!ID))
                If Val(NVL(rsTemp!ȱʡ��־)) = 1 And .ListIndex < 0 Then .ListIndex = .NewIndex
                If Val(str���) = Val(NVL(rsTemp!ID)) Then .ListIndex = .NewIndex
                rsTemp.MoveNext
            Loop
            If .ListIndex < 0 And .ListCount > 0 Then .ListIndex = 0
            lblUserType.Caption = "�����(&L)"
            lbl(6).Caption = "���ŷ�Χ(&B)"
            mblnNotClick = False
        End With
    Case gBillType.���ѿ�
        strSQL = "Select ���,����,���ų���,�Ƿ����� As ��������,ǰ׺�ı� From ���ѿ����Ŀ¼ Where Nvl(����,0) >=1 Order By ��� "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        mblnNotClick = True
        Set mcllCardProperty = New Collection
        With cbo���
            .Clear
            Do While Not rsTemp.EOF
                .AddItem NVL(rsTemp!���) & "-" & NVL(rsTemp!����)
                .ItemData(.NewIndex) = Val(NVL(rsTemp!���))
                mcllCardProperty.Add Array(Val(NVL(rsTemp!���ų���)), CStr(NVL(rsTemp!ǰ׺�ı�)), CStr(NVL(rsTemp!��������)), 0), "K" & Val(NVL(rsTemp!���))
                If Val(str���) = Val(NVL(rsTemp!���)) Then .ListIndex = .NewIndex
                rsTemp.MoveNext
            Loop
            If .ListIndex < 0 And .ListCount > 0 Then .ListIndex = 0
            lblUserType.Caption = "�����(&L)"
            lbl(6).Caption = "���ŷ�Χ(&B)"
            mblnNotClick = False
        End With
    Case Else
            cbo���.Visible = False: lblUserType.Visible = False
    End Select
    LoadCombox = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

