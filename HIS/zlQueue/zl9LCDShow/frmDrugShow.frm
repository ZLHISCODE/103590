VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmDrugShow 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8730
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11640
   LinkTopic       =   "Form1"
   ScaleHeight     =   8730
   ScaleWidth      =   11640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox PicMsg 
      Height          =   1215
      Left            =   0
      ScaleHeight     =   1155
      ScaleWidth      =   11595
      TabIndex        =   0
      Top             =   7440
      Width           =   11655
      Begin VB.Timer TimerCall 
         Interval        =   1000
         Left            =   1680
         Top             =   240
      End
      Begin VB.Timer timerLCD 
         Interval        =   10000
         Left            =   10680
         Top             =   120
      End
      Begin VB.Timer timerPage 
         Interval        =   5000
         Left            =   9000
         Top             =   360
      End
      Begin VB.Label lblmsg 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "����"
            Size            =   24
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   11655
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgCallingData 
      Height          =   7335
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   11895
      _cx             =   20981
      _cy             =   12938
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   0
      ForeColor       =   65280
      BackColorFixed  =   0
      ForeColorFixed  =   0
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   0
      BackColorAlternate=   0
      GridColor       =   65280
      GridColorFixed  =   65280
      TreeColor       =   -2147483633
      FloodColor      =   0
      SheetBorder     =   0
      FocusRect       =   1
      HighLight       =   0
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   0
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   10
      Cols            =   2
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmDrugShow.frx":0000
      ScrollTrack     =   0   'False
      ScrollBars      =   0
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
Attribute VB_Name = "frmDrugShow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private mintCols As Integer
Private mstrWins As String
Private mintRows As Integer
Private mrsData() As Recordset
Private mrsCallingData As Recordset
Private mrsPreparingData() As Recordset
Private mRowRec As Integer
Private mlngҩ��ID As Long
Private mIntCallRow As Integer
Private mIntPraRow As Integer
Private mbln��ҩ As Boolean
Private mbln��ҩȷ�� As Boolean
Private mstrSendNames() As String
Private mstrPraNames() As String
Private mIntSendPages() As Integer

Private Type Type_para
    bln��������ʾģʽ As Boolean             '������ʾģʽ�������壺�ര��
    Str���� As String
    dblLeft As Double
    dblTop As Double
    dblWidth As Double
    dblHeight As Double
    
    lng������������ɫ As Long
    
    bln��ʾ����ҩ As Boolean
    int����ҩ���� As Integer
    int����ҩ���� As Integer
    int����ҩ���� As Integer
    lng����ҩ������ɫ As Long
    
    bln��ʾ����ҩ As Boolean
    int����ҩ���� As Integer
    int����ҩ���� As Integer
    int����ҩ���� As Integer
    lng����ҩ������ɫ As Long
    
    bln��ʾ���� As Boolean
    lng����������ɫ As Long
    
    bln��ʾ�������� As Boolean
    lng��������������ɫ As Long
    
    
    intRowPeople  As Integer
    intPage As Integer
    intRefTime As Integer
    
    str��ʾ���� As String
End Type

Private mType_para As Type_para

Public Sub SetFacePostion()
'************************************************************************************
'
'���ý������ʾλ��
'
'************************************************************************************
    Dim strReg As String
    
    On Error Resume Next
        
    '��ע����У���ȡ��ʾ����
    strReg = "����ģ��\ҩ���Ŷӽк�\Һ������"
    
    '������ʾ����
    Me.Left = GetSetting("ZLSOFT", strReg, "��", "1024") * Screen.TwipsPerPixelX
    Me.Top = GetSetting("ZLSOFT", strReg, "��", "0") * Screen.TwipsPerPixelY
    Me.Width = GetSetting("ZLSOFT", strReg, "���", "1024") * Screen.TwipsPerPixelX
    Me.Height = GetSetting("ZLSOFT", strReg, "�߶�", "768") * Screen.TwipsPerPixelY
End Sub
Private Sub LoadPara()
    Dim strReg As String
    Dim i As Integer
    Dim strWin As String
    
    strReg = "����ģ��\ҩ���Ŷӽк�\Һ������"
    
    With mType_para
        .bln��������ʾģʽ = (Val(GetSetting("ZLSOFT", strReg, "����ģʽ", "0")) = 0)
        
        '���ش���
        .Str���� = GetSetting("ZLSOFT", strReg, "����", "1,2,3")
        
        '������Ļ��Ϣ
        .dblLeft = GetSetting("ZLSOFT", strReg, "��", "1024")
        .dblTop = GetSetting("ZLSOFT", strReg, "��", "0")
        .dblWidth = GetSetting("ZLSOFT", strReg, "���", "1024")
        .dblHeight = GetSetting("ZLSOFT", strReg, "�߶�", "768")
        
        '�����е�������ɫ
        .lng������������ɫ = GetSetting("ZLSOFT", strReg, "��������ɫ", vbGreen)
        
        '����ҩ�б������
        .bln��ʾ����ҩ = (Val(GetSetting("ZLSOFT", strReg, "��ʾ����ҩ", "1")) = 1)
        .int����ҩ���� = Val(GetSetting("ZLSOFT", strReg, "����ҩ����", "9"))
        .int����ҩ���� = Val(GetSetting("ZLSOFT", strReg, "����ҩ����", "3"))
        .int����ҩ���� = Val(GetSetting("ZLSOFT", strReg, "����ҩ����", "3"))
        .lng����ҩ������ɫ = GetSetting("ZLSOFT", strReg, "����ҩ��ɫ", vbGreen)
        
        '����ҩ�б������
        .bln��ʾ����ҩ = (Val(GetSetting("ZLSOFT", strReg, "��ʾ����ҩ", "1")) = 1)
        .int����ҩ���� = Val(GetSetting("ZLSOFT", strReg, "����ҩ����", "9"))
        .int����ҩ���� = Val(GetSetting("ZLSOFT", strReg, "����ҩ����", "9"))
        .int����ҩ���� = Val(GetSetting("ZLSOFT", strReg, "����ҩ����", "9"))
        .lng����ҩ������ɫ = GetSetting("ZLSOFT", strReg, "����ҩ��ɫ", vbGreen)
        
        .intRowPeople = 3
        .intPage = GetSetting("ZLSOFT", strReg, "��ҳʱ��", "5")
        .intRefTime = GetSetting("ZLSOFT", strReg, "ˢ��ʱ��", "10")
        
        .bln��ʾ���� = (Val(GetSetting("ZLSOFT", strReg, "��ʾ����", "1")) = 1)
        .lng����������ɫ = GetSetting("ZLSOFT", strReg, "������ɫ", vbGreen)
        
        .bln��ʾ�������� = (Val(GetSetting("ZLSOFT", strReg, "��ʾ��������", "1")) = 1)
        .lng��������������ɫ = GetSetting("ZLSOFT", strReg, "����������ɫ", vbBlack)
        
        .str��ʾ���� = GetSetting("ZLSOFT", strReg, "��ʾ����", "")
    End With
End Sub


Private Sub InitData(ByVal intPage As Integer, ByVal blnRef As Boolean)
'***********************************************************************
'
'ˢ�����ݣ�intPage=1Ϊ�������ʱ���������ݣ�intpage=2Ϊtimer�¼�ˢ������
'
'************************************************************************
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim strpeople As String
    Dim count As Integer
    Dim intcol As Integer
    Dim intSum As Integer
    Dim strTemp As String
    Dim intTemp As Integer
    Dim intCurPage As Integer
    Dim intPraPage As Integer
    Dim intCallPage As Integer
    
    '���ƴ���ҩ�б�ı߿�
     
    If vfgCallingData.Cols = 0 Then Exit Sub
    If mType_para.bln��ʾ����ҩ Then
        vfgCallingData.Select mIntCallRow + 2, 0, mIntCallRow + 2, (mintCols) * mRowRec - 1
        vfgCallingData.CellBorder &HFF00&, -1, -1, -1, 1, 0, 1
    End If
    
    For k = 0 To mintCols - 1
        If intPage = 1 Or blnRef Then
            loadCalling (Split(mstrWins, ",")(k))
            For j = 1 To mRowRec
                intSum = intSum + 1
                Me.vfgCallingData.TextMatrix(0, intSum - 1) = Split(mstrWins, ",")(k)
                
                If k Mod 2 = 0 Then
                    strTemp = String(0, " ")
                Else
                    strTemp = String(1, " ")
                End If
                
                If Not mrsCallingData.EOF Then
                    Me.vfgCallingData.TextMatrix(1, intSum - 1) = strTemp & "�� " & mrsCallingData!���� & " ��ҩ"
                Else
                    Me.vfgCallingData.TextMatrix(1, intSum - 1) = strTemp & "�޺�����Ա"
                End If
            Next
        End If
        If blnRef = False Then
            '��ʾ����ҩ��Ϣ
            If mType_para.bln��ʾ����ҩ Then
                loadData (Split(mstrWins, ",")(k)), k
                
                ShowSend k, intPage
            End If
        

            '��ʾ����ҩ��Ϣ
            ShowPra k, intPage
        End If
        
        '���߿�
        If k <> mintCols - 1 And vfgCallingData.Rows > 2 Then
            vfgCallingData.Select 2, (k + 1) * mRowRec - 1, mintRows - 1, (k + 1) * mRowRec
            vfgCallingData.CellBorder &HFF00&, -1, -1, -1, 0, 1, 0
            
            If mType_para.bln��ʾ����ҩ Then
                vfgCallingData.Select mIntCallRow + 2, (k + 1) * mRowRec - 1, mIntCallRow + 2, (k + 1) * mRowRec
                vfgCallingData.CellBorder &HFF00&, -1, -1, -1, 1, 1, 1
            End If
        End If
    Next
    
    '�ϲ����ںͽк���Ϣ
    vfgCallingData.MergeRow(0) = True
    vfgCallingData.MergeRow(1) = True
    vfgCallingData.Refresh
    
    vfgCallingData.Select 0, 0, 1, mintCols * mRowRec - 1
    vfgCallingData.CellBorder &HFF00&, 0, 0, 0, 1, 1, 1
End Sub


Private Sub Form_Load()
    '���ز���
    LoadPara
    
    SetFacePostion
    
    Me.vfgCallingData.Move 0, 0, Me.ScaleWidth, IIf(mType_para.bln��ʾ��������, Round(Me.ScaleHeight * 0.9), Round(Me.ScaleHeight))
    
    '����ģʽȷ���������ʾ���ڣ��������Ǵ��Σ��ര�����ڲ�������ʱ����ѡ��
    If mType_para.bln��������ʾģʽ = False Then
        mstrWins = mType_para.Str����
        If mstrWins = "" Then
            Exit Sub
        End If
    Else
        Me.TimerCall.Enabled = False
    End If
    
    mintCols = UBound(Split(mstrWins, ",")) + 1
'    If mintCols = 0 Then Exit Sub
    mRowRec = mType_para.intRowPeople
    
    'ȷ�����ݼ�����ĳ���
    ReDim mrsData(mintCols)
    ReDim mrsPreparingData(mintCols)
    ReDim mstrSendNames(mintCols)
    ReDim mstrPraNames(mintCols)
    
    '��ʼ�����
    InitVSF
    
    InitData 1, False
    
    Me.timerPage.Interval = mType_para.intPage * 1000
    Me.timerLCD.Interval = mType_para.intRefTime * 1000
    
    Me.PicMsg.Visible = mType_para.bln��ʾ��������
    Me.lblmsg.Caption = IIf(mType_para.str��ʾ���� = "", "ף�����տ�����  " & Format(zlDatabase.Currentdate, "yyyy-mm-dd  hh:mm"), mType_para.str��ʾ����)
End Sub

Private Sub loadData(ByVal strWin As String, ByVal Index As Integer)
'************************************************************************
'
'���ش���ҩ�б������
'
'************************************************************************
    Dim strSql As String
    Dim date��ʼ���� As Date
    Dim date�������� As Date
        
    On Error GoTo errHandle
    date��ʼ���� = zlDatabase.Currentdate
    date��ʼ���� = CDate(Format(date��ʼ����, "yyyy-mm-dd") & " 00:00:00")
    
    date�������� = zlDatabase.Currentdate
    date�������� = CDate(Format(date��������, "yyyy-mm-dd") & " 23:59:59")

    strSql = "Select distinct A.����,B.��ҩ����,B.ǩ��ʱ�� From δ��ҩƷ��¼ A,ҩƷ�շ���¼ B,������ü�¼ C" & _
             " Where A.����=B.���� And A.No=B.NO And A.�ⷿid=B.�ⷿid and B.����id=C.id and (A.����=8 or A.����=9 or A.����=10) "
    
    If mbln��ҩ Then
        strSql = strSql & " and (A.�Ŷ�״̬=2 or A.�Ŷ�״̬=4) and A.�ⷿid=[1] and A.��ҩ����=[2] and A.�������� between [3] and [4] And (B.��¼״̬=1 Or Mod(B.��¼״̬,3)=0)"
    ElseIf mbln��ҩȷ�� And mbln��ҩ = False Then
        strSql = strSql & " and (A.�Ŷ�״̬=1 or A.�Ŷ�״̬=2 or A.�Ŷ�״̬=4) and A.�ⷿid=[1] and A.��ҩ����=[2] and A.�������� between [3] and [4] And (B.��¼״̬=1 Or Mod(B.��¼״̬,3)=0)"
    ElseIf mbln��ҩ = False And mbln��ҩȷ�� = False Then
        strSql = strSql & "  and (A.�Ŷ�״̬<>3 or A.�Ŷ�״̬ is null) and A.�ⷿid=[1] and A.��ҩ����=[2] and A.�������� between [3] and [4] And (B.��¼״̬=1 Or Mod(B.��¼״̬,3)=0)"
    End If
    
    Set mrsData(Index) = zlDatabase.OpenSQLRecord(strSql, "", mlngҩ��ID, strWin, date��ʼ����, date��������)
    
    
    If Not mrsData(Index).EOF Then
        If Nvl(mrsData(Index)!��ҩ����) <> "" Then
            mrsData(Index).Sort = "��ҩ����"
        Else
            mrsData(Index).Sort = "ǩ��ʱ��"
        End If
    End If
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub loadCalling(ByVal strWin As String)
'************************************************************************
'
'���ص�ǰ���е�����
'
'************************************************************************
    Dim strSql As String
    Dim date��ʼ���� As Date
    Dim date�������� As Date
    
    On Error GoTo errHandle
    date��ʼ���� = zlDatabase.Currentdate
    date��ʼ���� = CDate(Format(date��ʼ����, "yyyy-mm-dd") & " 00:00:00")
    
    date�������� = zlDatabase.Currentdate
    date�������� = CDate(Format(date��������, "yyyy-mm-dd") & " 23:59:59")
    
    strSql = "select ���� from δ��ҩƷ��¼ where �Ŷ�״̬=3 and �ⷿid=[1] and ��ҩ����=[2] and �������� between [3] and [4]"
    Set mrsCallingData = zlDatabase.OpenSQLRecord(strSql, "", mlngҩ��ID, strWin, date��ʼ����, date��������)
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub


Private Sub loadPreparing(ByVal strWin As String, ByVal intIndex As Integer)
'************************************************************************
'
'���ش���ҩ�б������
'
'************************************************************************
    Dim strSql As String
    
    On Error GoTo errHandle
    Dim date��ʼ���� As Date
    Dim date�������� As Date
        
    On Error GoTo errHandle
    date��ʼ���� = zlDatabase.Currentdate
    date��ʼ���� = CDate(Format(date��ʼ����, "yyyy-mm-dd") & " 00:00:00")
    
    date�������� = zlDatabase.Currentdate
    date�������� = CDate(Format(date��������, "yyyy-mm-dd") & " 23:59:59")
    
    strSql = "Select distinct A.����,B.��ҩ����,B.ǩ��ʱ�� From δ��ҩƷ��¼ A,ҩƷ�շ���¼ B,������ü�¼ C" & _
             " Where A.����=B.���� And A.No=B.NO And A.�ⷿid=B.�ⷿid and B.����id=C.id and (A.����=8 or A.����=9 or A.����=10) "
    If mbln��ҩȷ�� Then
        strSql = strSql & "and A.�Ŷ�״̬=1 and A.�ⷿid=[1] and A.��ҩ����=[2] and A.�������� between [3] and [4] And (B.��¼״̬=1 Or Mod(B.��¼״̬,3)=0)"
    Else
        strSql = strSql & "and (A.�Ŷ�״̬=1 or A.�Ŷ�״̬=0 or A.�Ŷ�״̬ is null) and A.�ⷿid=[1] and A.��ҩ����=[2] and A.�������� between [3] and [4] And (B.��¼״̬=1 Or Mod(B.��¼״̬,3)=0)"
    End If
    
    If mbln��ҩ = False Then strSql = strSql & " And 1=2"
    
    Set mrsPreparingData(intIndex) = zlDatabase.OpenSQLRecord(strSql, "", mlngҩ��ID, strWin, date��ʼ����, date��������)
    
    If Not mrsPreparingData(intIndex).EOF Then
        If Nvl(mrsPreparingData(intIndex)!��ҩ����) <> "" Then
             mrsPreparingData(intIndex).Sort = "��ҩ����"
        Else
             mrsPreparingData(intIndex).Sort = "ǩ��ʱ��"
        End If
    End If
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub
Private Sub InitVSF()
'************************************************************************
'
'��ʼ�����
'
'************************************************************************
    Dim intColWidth As Integer
    Dim intRowheight As Integer
    Dim i As Integer
    Dim strReg As String
    Dim dblHeight As Double
    
    strReg = "����ģ��\ҩ���Ŷӽк�\Һ������"
    
    mintRows = 2
'    dblHeight = (20 * Val(GetSetting("ZLSOFT", strReg, "�ֺ�(0)", "14"))) * 1.5
'
'    If mType_para.bln��ʾ���� Then
'        dblHeight = dblHeight + (20 * Val(GetSetting("ZLSOFT", strReg, "�ֺ�(1)", "14"))) * 1.5
'        mintRows = 2
'    End If
'
'    If mType_para.bln��ʾ����ҩ = True Then
'        mIntCallRow = (Me.vfgCallingData.Height - dblHeight) * 0.6 \ ((20 * Val(GetSetting("ZLSOFT", strReg, "�ֺ�(2)", "14")) * 2))
'    End If
'
'    If mType_para.bln��ʾ����ҩ = True Then
'        mIntPraRow = (Me.vfgCallingData.Height - dblHeight) * 0.4 \ ((20 * Val(GetSetting("ZLSOFT", strReg, "�ֺ�(3)", "14")) * 2))
'    End If
'
'    If Val(GetSetting("ZLSOFT", strReg, "�ֺ�(2)", "14")) > Val(GetSetting("ZLSOFT", strReg, "�ֺ�(3)", "14")) Then
'        mRowRec = mType_para.int����ҩ���� \ mIntCallRow + 1
'    Else
'        mRowRec = mType_para.int����ҩ���� \ mIntPraRow + 1
'    End If
'
'    If mType_para.int����ҩ���� Mod mRowRec = 0 Then
'        mIntCallRow = mType_para.int����ҩ���� \ mRowRec
'    Else
'        mIntCallRow = mType_para.int����ҩ���� \ mRowRec + 1
'    End If
'
'    If mType_para.int����ҩ���� Mod mRowRec = 0 Then
'        mIntPraRow = mType_para.int����ҩ���� \ mRowRec
'    Else
'        mIntPraRow = mType_para.int����ҩ���� \ mRowRec + 1
'    End If
    
    mIntCallRow = IIf(mType_para.bln��ʾ����ҩ, mType_para.int����ҩ����, 0)
    mIntPraRow = IIf(mType_para.bln��ʾ����ҩ, mType_para.int����ҩ����, 0)
    mRowRec = mType_para.int����ҩ����
    mintRows = mintRows + mIntCallRow + mIntPraRow + IIf(mType_para.bln��ʾ����ҩ, 1, 0) + IIf(mType_para.bln��ʾ����ҩ, 1, 0)
    With vfgCallingData
        .Rows = mintRows
        .Cols = mintCols * mRowRec
        
        If .Cols = 0 Then Exit Sub
        
'        If mintCols = 0 Then
'            Unload Me
'            Exit Sub
'        End If
        '���ñ��Ϊ���ɺϲ�
        .MergeCells = flexMergeFree

         '���������������ɫ��С
        SetFont
        
        intColWidth = Me.ScaleWidth / (mintCols * mRowRec)
        '�������ݾ�����ʾ
        For i = 0 To mintCols * mRowRec - 1
            .ColWidth(i) = intColWidth
            vfgCallingData.ColAlignment(i) = flexAlignCenterCenter
        Next
        
        If mType_para.bln��ʾ����ҩ = False And mType_para.bln��ʾ����ҩ = False Then
            .RowHeight(0) = (20 * Val(GetSetting("ZLSOFT", strReg, "�ֺ�(1)", "14"))) * 1.5
            .RowHeight(1) = IIf(mType_para.bln��ʾ����, .Height - .RowHeight(0), .Height)
        ElseIf mType_para.bln��ʾ����ҩ = False Or mType_para.bln��ʾ����ҩ = False Then
            .RowHeight(0) = (20 * Val(GetSetting("ZLSOFT", strReg, "�ֺ�(1)", "14"))) * 1.5
            .RowHeight(1) = (20 * Val(GetSetting("ZLSOFT", strReg, "�ֺ�(0)", "14"))) * 2
        Else
            .RowHeight(0) = (20 * Val(GetSetting("ZLSOFT", strReg, "�ֺ�(1)", "14"))) * 1.5
            .RowHeight(1) = (20 * Val(GetSetting("ZLSOFT", strReg, "�ֺ�(0)", "14"))) * 2
        End If
        
        If Not mType_para.bln��ʾ���� Then
            .RowHeight(0) = 0
        End If
        If vfgCallingData.Rows > 2 Then
            intRowheight = (.Height - .RowHeight(0) - .RowHeight(1)) / (mintRows - 2)
            For i = 2 To mintRows - 1
                .RowHeight(i) = intRowheight
            Next
        End If
    End With
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Me.vfgCallingData.Move 0, 0, Me.ScaleWidth, IIf(mType_para.bln��ʾ��������, Round(Me.ScaleHeight * 0.9), Round(Me.ScaleHeight))
    Me.PicMsg.Move 0, Me.vfgCallingData.Height, Me.vfgCallingData.Width, Round(Me.ScaleHeight * 0.1)
    Me.lblmsg.Move 0, Me.PicMsg.Height / 5, Me.PicMsg.Width, Me.PicMsg.Height
    
    InitVSF
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer
    
    For i = 0 To mintCols - 1
        mstrSendNames(i) = ""
        mstrPraNames(i) = ""
        Set mrsData(i) = Nothing
        Set mrsData(i) = Nothing
    Next
End Sub

Private Sub TimerCall_Timer()
    InitData 2, True
End Sub

Private Sub timerPage_Timer()
'************************************************************************
'
'�Դ���ҩ�б�����ݽ��з�ҳ
'
'************************************************************************
    Dim i As Integer
    Dim intcol As Integer
    Dim k As Integer
    Dim count As Integer
    Dim intPage As Integer
    Dim strTemp As String
    Dim intCallPage As Integer
    Dim intPraPage As Integer

    If mType_para.bln��ʾ����ҩ = False And mType_para.bln��ʾ����ҩ = False Then Exit Sub

'    Me.timerLCD.Enabled = False

    For k = 0 To mintCols - 1
        intcol = k * mRowRec
        '����������ڷ�ҳ֮���ҳ��
'        For intcol = k * mRowRec To (k + 1) * mRowRec - 1
        If (intcol \ mRowRec) Mod 2 = 0 Then
            strTemp = String(0, " ")
        Else
            strTemp = String(1, " ")
        End If

        If mType_para.bln��ʾ����ҩ = True Then
            '����ҩ����ҳ��
            intCallPage = (mrsData(k).RecordCount \ (mRowRec * mIntCallRow) + IIf(mrsData(k).RecordCount Mod (mRowRec * mIntCallRow) = 0, 0, 1))

            If intCallPage = 0 Then intCallPage = 1

            '��ǰҳ
            intPage = Val(Mid(Me.vfgCallingData.TextMatrix(mIntCallRow + 2, intcol), 4, InStr(1, Me.vfgCallingData.TextMatrix(mIntCallRow + 2, intcol), "/"))) + 1

            If intPage > intCallPage Then intPage = 1

            Me.vfgCallingData.Cell(flexcpText, mIntCallRow + 2, intcol, mIntCallRow + 2, (k + 1) * mRowRec - 1) = "����ҩ   " & strTemp & intPage & "/" & intCallPage & " ��" & mrsData(k).RecordCount & "��"
        End If

    
        If mType_para.bln��ʾ����ҩ = True Then
            '�����ݼ������ݣ������Ѿ��Ƶ����һ�������˾��Ƶ���һ��
            If mrsData(k).EOF And mrsData(k).RecordCount > 0 Then mrsData(k).MoveFirst
            '��մ���ҩ�б������
            Me.vfgCallingData.Cell(flexcpText, 2, k * mRowRec, mIntCallRow + 1, (k + 1) * mRowRec - 1) = ""

            i = 2
            count = 0
            intcol = k * mRowRec
            If mstrSendNames(k) = "" Or intPage = 1 Then mstrSendNames(k) = ","
            Do While Not mrsData(k).EOF
                If InStr(1, mstrSendNames(k), "," & mrsData(k)!���� & ",") = 0 Then
                    count = count + 1
                    Me.vfgCallingData.TextMatrix(i, intcol) = Nvl(mrsData(k)!����)
                    
                    '����ÿ������ÿ�����һ��ʱ��������һ����ʾ
                    
                    intcol = intcol + 1
'                    If count = mRowRec Then
'                        count = 0
'                        intcol = k * mRowRec
'                        If Not mrsData(k).EOF Then i = i + 1
'                    End If
                End If

                mstrSendNames(k) = mstrSendNames(k) & mrsData(k)!���� & ","
                mrsData(k).MoveNext
                
                 '�����ݵ���ʾ�Ѿ���������ֵʱ���˳�ѭ��
                If (i - 2) * mRowRec + count = mType_para.int����ҩ���� Then
                    Exit Do
                End If
                '����ÿ������ÿ�����һ��ʱ��������һ����ʾ
                If count = mRowRec Then
                    count = 0
                    intcol = k * mRowRec
                    If Not mrsData(k).EOF Then i = i + 1
                End If
            Loop
        End If
        
        intcol = 0
        If mType_para.bln��ʾ����ҩ = True Then
             '��ǰ��ҳ��
            intPraPage = (mrsPreparingData(k).RecordCount \ (mIntPraRow * mRowRec) + IIf(mrsPreparingData(k).RecordCount Mod (mIntPraRow * mRowRec) = 0, 0, 1))

            If intPraPage = 0 Then intPraPage = 1

            intPage = Val(Mid(Me.vfgCallingData.TextMatrix(mintRows - 1, intcol), 4, InStr(1, Me.vfgCallingData.TextMatrix(mintRows - 1, intcol), "/"))) + 1

            If intPage > intPraPage Then intPage = 1
            
            Me.vfgCallingData.Cell(flexcpText, mintRows - 1, intcol, mintRows - 1, (k + 1) * mRowRec - 1) = "����ҩ   " & strTemp & intPage & "/" & intPraPage & " ��" & mrsPreparingData(k).RecordCount & "��"
        End If

        If mType_para.bln��ʾ����ҩ = True Then
            count = 0
            '�����ݼ������ݣ������Ѿ��Ƶ����һ�������˾��Ƶ���һ��
            If mrsPreparingData(k).EOF And mrsPreparingData(k).RecordCount > 0 Then mrsPreparingData(k).MoveFirst
            '��մ���ҩ�б������
            Me.vfgCallingData.Cell(flexcpText, mintRows - mIntPraRow - 1, k * mRowRec, mintRows - 2, (k + 1) * mRowRec - 1) = ""

            i = mintRows - mIntPraRow - 1
            intcol = k * mRowRec
            If mstrPraNames(k) = "" Or intPage = 1 Then mstrPraNames(k) = ","
            Do While Not mrsPreparingData(k).EOF
                If InStr(1, mstrPraNames(k), "," & mrsPreparingData(k)!���� & ",") = 0 Then
                    count = count + 1
                    Me.vfgCallingData.TextMatrix(i, intcol) = mrsPreparingData(k)!����

                    '����ÿ������ÿ�����һ��ʱ��������һ����ʾ
                    intcol = intcol + 1
                    
'                    If count = mRowRec Then
'                        count = 0
'                        intcol = k * mRowRec
'                        If Not mrsPreparingData(k).EOF Then i = i + 1
'                    End If
                    mstrPraNames(k) = mstrPraNames(k) & mrsPreparingData(k)!���� & ","
                    
                    If (i - (mintRows - mIntPraRow - 1)) * mRowRec + count = mType_para.int����ҩ���� Then
                        Exit Do
                    End If
                    If count = mRowRec Then
                        count = 0
                        intcol = k * mRowRec
                        If Not mrsPreparingData(k).EOF Then i = i + 1
                    End If
                End If
                
                mrsPreparingData(k).MoveNext
                
                '�����ݵ���ʾ�Ѿ���������ֵʱ���˳�ѭ��
                
            Loop
        End If

'        intcol = k * mRowRec
'        '����������ڷ�ҳ֮���ҳ��
''        For intcol = k * mRowRec To (k + 1) * mRowRec - 1
'        If (intcol \ mRowRec) Mod 2 = 0 Then
'            strTemp = String(0, " ")
'        Else
'            strTemp = String(1, " ")
'        End If
'
'        If mType_para.bln��ʾ����ҩ = True Then
'            '����ҩ����ҳ��
'            intCallPage = (mrsData(k).RecordCount \ (mRowRec * mIntCallRow) + IIf(mrsData(k).RecordCount Mod (mRowRec * mIntCallRow) = 0, 0, 1))
'
'            If intCallPage = 0 Then intCallPage = 1
'
'            '��ǰҳ
'            intPage = Val(Mid(Me.vfgCallingData.TextMatrix(mIntCallRow + 2, intcol), 4, InStr(1, Me.vfgCallingData.TextMatrix(mIntCallRow + 2, intcol), "/"))) + 1
'
'            If intPage > intCallPage Then intPage = 1
'
'            Me.vfgCallingData.Cell(flexcpText, mIntCallRow + 2, intcol, mIntCallRow + 2, (k + 1) * mRowRec - 1) = "����ҩ   " & strTemp & intPage & "/" & intCallPage & " ��" & mrsData(k).RecordCount & "��"
'        End If
'
'        If mType_para.bln��ʾ����ҩ = True Then
'             '��ǰ��ҳ��
'            intPraPage = (mrsPreparingData(k).RecordCount \ (mIntPraRow * mRowRec) + IIf(mrsPreparingData(k).RecordCount Mod (mIntPraRow * mRowRec) = 0, 0, 1))
'
'            If intPraPage = 0 Then intPraPage = 1
'
'            intPage = Val(Mid(Me.vfgCallingData.TextMatrix(mintRows - 1, intcol), 4, InStr(1, Me.vfgCallingData.TextMatrix(mintRows - 1, intcol), "/"))) + 1
'
'            If intPage > intPraPage Then intPage = 1
'
'
'            Me.vfgCallingData.Cell(flexcpText, mintRows - 1, intcol, mintRows - 1, (k + 1) * mRowRec - 1) = "����ҩ   " & strTemp & intPage & "/" & intPraPage & " ��" & mrsPreparingData(k).RecordCount & "��"
'        End If
'        Next
    Next
End Sub

Public Sub ShowMe(ByVal lngҩ��ID As Long, ByVal strWins As String, ByVal bln��ҩ As Boolean, ByVal bln��ҩȷ�� As Boolean)
'**************************************************************************
'�򿪴���Ľӿڣ�lngҩ��ID����ǰ��ҩ��id��strWins���������Ӵ�
'**************************************************************************
    mlngҩ��ID = lngҩ��ID
    mstrWins = strWins
    mbln��ҩ = bln��ҩ
    mbln��ҩȷ�� = bln��ҩȷ��
    Dim strTemp As String
    Dim strReg As String
    Dim cls As New clsLCDShow
    
    strReg = "����ģ��\ҩ���Ŷӽк�\Һ������"
    strTemp = GetSetting("ZLSOFT", strReg, "����", "1,2,3")
    If strTemp = "" And strWins = "" Then
        cls.zlClose
        Exit Sub
    End If
    
    Me.Show
End Sub

Private Sub timerLCD_Timer()
'************************************************************************
'
'ˢ�´���ҩ�б������
'
'************************************************************************
    InitData 2, False
    
    Me.lblmsg.Caption = IIf(mType_para.str��ʾ���� = "", "ף�����տ�����  " & Format(zlDatabase.Currentdate, "yyyy-mm-dd  hh:mm"), mType_para.str��ʾ����)
End Sub

Public Sub ChangeCall(ByVal strWin As String, ByVal strName As String)
'****************************************************************************
'
'���µ�ǰ������Ϣ
'
'**************************************************************************

    InitData 2, True
End Sub

Private Sub ShowSend(ByVal Index As Integer, ByVal intPage As Integer)
'******************************************************************************
'
'������ҩ�����ݼ��ص���������
'
'******************************************************************************
    Dim count As Integer
    Dim i As Integer
    Dim intcol As Integer
    Dim intCallPage As Integer
    Dim intCurPage As Integer
    Dim strTemp As String
    Dim strNames As String
    
    '��ʾ����ҩ��Ϣ
    If mType_para.bln��ʾ����ҩ Then
        '�����ܵ�ҳ��
        intCallPage = (mrsData(Index).RecordCount \ (mRowRec * mIntCallRow) + IIf(mrsData(Index).RecordCount Mod (mRowRec * mIntCallRow) = 0, 0, 1))
        If intCallPage = 0 Then intCallPage = 1
        
        '�ж��Ƿ�Ϊ�������
        If intPage <> 1 Then
            For i = 0 To Me.vfgCallingData.Cols - 1
                '���㵱ǰҳ��
                If vfgCallingData.TextMatrix(0, i) = (Split(mstrWins, ",")(Index)) Then
                    intCurPage = Val(Mid(Me.vfgCallingData.TextMatrix(mIntCallRow + 2, i), 4, InStr(1, Me.vfgCallingData.TextMatrix(mIntCallRow + 2, i), "/")))
                End If
            Next
            
            '����¼�����α�����ǰҳ��ʾ������
            For i = 1 To mIntCallRow * mRowRec * (intCurPage - 1)
                If Not mrsData(Index).EOF Then
                    mrsData(Index).MoveNext
                End If
            Next
            
        Else
            intCurPage = 1
        End If
        
        count = 0
        i = 2
        intcol = Index * mRowRec
        
        'ѭ����¼������ʾ����
        mstrSendNames(Index) = ","
        Do While Not mrsData(Index).EOF
            If InStr(1, mstrSendNames(Index), "," & mrsData(Index)!���� & ",") = 0 Then
                count = count + 1
                Me.vfgCallingData.TextMatrix(i, intcol) = Nvl(mrsData(Index)!����)
                
'                'ÿ����ʾ���ƶ�������������һ��
                intcol = intcol + 1
'                If count = mRowRec Then
'                    count = 0
'                    intcol = Index * mRowRec
'                    If Not mrsData(Index).EOF Then i = i + 1
'                End If
            End If
            
            mstrSendNames(Index) = mstrSendNames(Index) & Nvl(mrsData(Index)!����) & ","
            '������һ����¼
            mrsData(Index).MoveNext
            
            '������ʾ��ָ���������˳�ѭ��
            If (i - 2) * mRowRec + count = mType_para.int����ҩ���� Then
                Exit Do
            End If
'            intcol = intcol + 1
            If count = mRowRec Then
                count = 0
                intcol = Index * mRowRec
                If Not mrsData(Index).EOF Then i = i + 1
            End If
            
        Loop
        
        '��ʾ��ҳ��Ϣ
        intcol = Index * mRowRec
        For intcol = Index * mRowRec To (Index + 1) * mRowRec - 1
            If (intcol \ mRowRec) Mod 2 = 0 Then
                strTemp = String(0, " ")
            Else
                strTemp = String(1, " ")
            End If
            
            Me.vfgCallingData.TextMatrix(mIntCallRow + 2, intcol) = "����ҩ   " & strTemp & intCurPage & "/" & intCallPage & " ��" & mrsData(Index).RecordCount & "��"
        Next
        '�ϲ���ʾ��ҳ��Ϣ
        vfgCallingData.MergeRow(mIntCallRow + 2) = True
    End If
End Sub

Private Sub ShowPra(ByVal Index As Integer, ByVal intPage As Integer)
    Dim count As Integer
    Dim i As Integer
    Dim intPraPage As Integer
    Dim intCurPage As Integer
    Dim intcol As Integer
    Dim strTemp As String
    
    If mType_para.bln��ʾ����ҩ Then
        '���ش���ҩ��Ϣ
        loadPreparing (Split(mstrWins, ",")(Index)), Index
        
        '������ҳ��
        intPraPage = (mrsPreparingData(Index).RecordCount \ (mIntPraRow * mRowRec) + IIf(mrsPreparingData(Index).RecordCount Mod (mIntPraRow * mRowRec) = 0, 0, 1))
        If intPraPage = 0 Then intPraPage = 1
        
        '�ж��Ƿ��Ǵ������
        If intPage <> 1 Then
            For i = 0 To Me.vfgCallingData.Cols - 1
                '�õ���ǰҳ��
                If vfgCallingData.TextMatrix(0, i) = (Split(mstrWins, ",")(Index)) Then
                    intCurPage = Val(Mid(Me.vfgCallingData.TextMatrix(vfgCallingData.Rows - 1, i), 4, InStr(1, Me.vfgCallingData.TextMatrix(vfgCallingData.Rows - 1, i), "/")))
                    Exit For
                End If
            Next
            
            '����¼���Ƶ���ǰҳ����λ��
            For i = 1 To mIntPraRow * mRowRec * (intCurPage - 1)
                If Not mrsPreparingData(Index).EOF Then
                    mrsPreparingData(Index).MoveNext
                End If
            Next
            
        Else
            intCurPage = 1
        End If
        
        count = 0
        i = Me.vfgCallingData.Rows - mIntPraRow - 1
        intcol = Index * mRowRec
        mstrPraNames(Index) = ","
        'ѭ����¼������������ʾ������
        Do While Not mrsPreparingData(Index).EOF
            If InStr(1, mstrPraNames(Index), "," & Nvl(mrsPreparingData(Index)!����) & ",") = 0 Then
                count = count + 1
                Me.vfgCallingData.TextMatrix(i, intcol) = Nvl(mrsPreparingData(Index)!����)
                
                'һ�����ݼ�����֮��������һ��
                intcol = intcol + 1
'                If count = mRowRec Then
'                    count = 0
'                    intcol = Index * mRowRec
'                End If
                mstrPraNames(Index) = mstrPraNames(Index) & Nvl(mrsPreparingData(Index)!����) & ","
            End If
            
            
            mrsPreparingData(Index).MoveNext
            
             '�������������ʾ�ƶ��������˳�ѭ��
            If (i - (Me.vfgCallingData.Rows - mIntPraRow - 1)) * mRowRec + count = mType_para.int����ҩ���� Then
                Exit Do
            End If
            'һ�����ݼ�����֮��������һ��
'            intcol = intcol + 1
            If count = mRowRec Then
                count = 0
                intcol = Index * mRowRec
                i = i + 1
            End If
        Loop
        
        '��ʾ��ҳ��Ϣ
        intcol = Index * mRowRec
        For intcol = Index * mRowRec To (Index + 1) * mRowRec - 1
            If (intcol \ mRowRec) Mod 2 = 0 Then
                strTemp = String(0, " ")
            Else
                strTemp = String(1, " ")
            End If
            
            Me.vfgCallingData.Cell(flexcpText, mintRows - 1, intcol, mintRows - 1, (Index + 1) * mRowRec - 1) = "����ҩ   " & strTemp & intCurPage & "/" & intPraPage & " ��" & mrsPreparingData(Index).RecordCount & "��"
        Next
        '�ϲ���ʾ��ҳ��Ϣ
        vfgCallingData.MergeRow(Me.vfgCallingData.Rows - 1) = True
    End If
End Sub

Public Sub SetFont()
    Dim strReg As String
    
    strReg = "����ģ��\ҩ���Ŷӽк�\Һ������"
    With Me.vfgCallingData
        If .Cols = 0 Then Exit Sub
        '���������������ɫ��С
        .Cell(flexcpFontSize, 1, 0, 1, .Cols - 1) = Val(GetSetting("ZLSOFT", strReg, "�ֺ�(0)", "14"))
        .Cell(flexcpForeColor, 1, 0, 1, .Cols - 1) = GetSetting("ZLSOFT", strReg, "��������ɫ", vbGreen)
        .Cell(flexcpFontName, 1, 0, 1, .Cols - 1) = GetSetting("ZLSOFT", strReg, "����(0)", "����")
        .Cell(flexcpFontBold, 1, 0, 1, .Cols - 1) = GetSetting("ZLSOFT", strReg, "����(0)", "false")
        .Cell(flexcpFontItalic, 1, 0, 1, .Cols - 1) = GetSetting("ZLSOFT", strReg, "б��(0)", "false")
        If mType_para.bln��ʾ���� Then
            .Cell(flexcpFontSize, 0, 0, 0, .Cols - 1) = Val(GetSetting("ZLSOFT", strReg, "�ֺ�(1)", "14"))
            .Cell(flexcpForeColor, 0, 0, 0, .Cols - 1) = GetSetting("ZLSOFT", strReg, "������ɫ", vbGreen)
            .Cell(flexcpFontName, 0, 0, 0, .Cols - 1) = GetSetting("ZLSOFT", strReg, "����(1)", "����")
            .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = GetSetting("ZLSOFT", strReg, "����(1)", "false")
            .Cell(flexcpFontItalic, 0, 0, 0, .Cols - 1) = GetSetting("ZLSOFT", strReg, "б��(1)", "false")
        End If
        
        If mType_para.bln��ʾ����ҩ = True Then
            .Cell(flexcpFontSize, 2, 0, mIntCallRow + 2, .Cols - 1) = Val(GetSetting("ZLSOFT", strReg, "�ֺ�(2)", "14"))
            .Cell(flexcpForeColor, 2, 0, mIntCallRow + 2, .Cols - 1) = GetSetting("ZLSOFT", strReg, "����ҩ��ɫ", vbGreen)
            .Cell(flexcpFontName, 2, 0, mIntCallRow + 2, .Cols - 1) = GetSetting("ZLSOFT", strReg, "����(2)", "����")
            .Cell(flexcpFontBold, 2, 0, mIntCallRow + 2, .Cols - 1) = GetSetting("ZLSOFT", strReg, "����(2)", "false")
            .Cell(flexcpFontItalic, 2, 0, mIntCallRow + 2, .Cols - 1) = GetSetting("ZLSOFT", strReg, "б��(2)", "false")
        End If
        
        If mType_para.bln��ʾ����ҩ = True Then
            .Cell(flexcpFontSize, mintRows - mIntPraRow - 1, 0, mintRows - 1, .Cols - 1) = Val(GetSetting("ZLSOFT", strReg, "�ֺ�(3)", "14"))
            .Cell(flexcpForeColor, mintRows - mIntPraRow - 1, 0, mintRows - 1, .Cols - 1) = GetSetting("ZLSOFT", strReg, "����ҩ��ɫ", vbGreen)
            .Cell(flexcpFontName, mintRows - mIntPraRow - 1, 0, mintRows - 1, .Cols - 1) = GetSetting("ZLSOFT", strReg, "����(3)", "����")
            .Cell(flexcpFontBold, mintRows - mIntPraRow - 1, 0, mintRows - 1, .Cols - 1) = GetSetting("ZLSOFT", strReg, "����(3)", "false")
            .Cell(flexcpFontItalic, mintRows - mIntPraRow - 1, 0, mintRows - 1, .Cols - 1) = GetSetting("ZLSOFT", strReg, "б��(3)", "false")
        End If
    End With
    
    Me.lblmsg.ForeColor = GetSetting("ZLSOFT", strReg, "����������ɫ", vbBlack)
    Me.lblmsg.FontSize = Val(GetSetting("ZLSOFT", strReg, "�ֺ�(4)", "14"))
    Me.lblmsg.FontName = GetSetting("ZLSOFT", strReg, "����(4)", "����")
    Me.lblmsg.FontBold = GetSetting("ZLSOFT", strReg, "����(4)", "false")
    Me.lblmsg.FontItalic = GetSetting("ZLSOFT", strReg, "б��(4)", "false")
End Sub

'Private Sub GetTotal(ByVal intType As Integer)
'    Dim i As Integer
'    Dim strTemp As String
'
'    If intType = 0 Then
'        For i = 0 To mintCols - 1
'            strTemp = ","
'            If Not mrsData(i) Is Nothing Then mrsData(i).MoveFirst
'            Do While mrsData(i).EOF
'
'                If InStr(1, strTemp, "," & mrsData(i)!���� & ",") Then
'                    mintSenpages(i) = mintSenpages(i) + 1
'                End If
'                strTemp = strTemp & mrsData(i)!���� & ","
'                mrsData(i).MoveNext
'            Loop
'            mrsData(i).MoveFirst
'        Next
'    Else
'
'    End If
'End Sub


