VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmIllImport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�������뵼��"
   ClientHeight    =   8445
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10815
   Icon            =   "frmIllImport.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8445
   ScaleWidth      =   10815
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   10815
      TabIndex        =   15
      Top             =   7710
      Width           =   10815
      Begin VB.CommandButton cmdCancel 
         Caption         =   "�˳�(&C)"
         Height          =   350
         Left            =   9600
         TabIndex        =   17
         Tag             =   "����"
         Top             =   240
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "����(&O)"
         Height          =   350
         Left            =   8400
         TabIndex        =   16
         Tag             =   "����"
         Top             =   240
         Width           =   1100
      End
      Begin MSComctlLib.ProgressBar prg 
         Height          =   225
         Left            =   240
         TabIndex        =   18
         Top             =   360
         Width           =   6450
         _ExtentX        =   11377
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   240
         TabIndex        =   19
         Top             =   120
         Width           =   90
      End
   End
   Begin MSComDlg.CommonDialog dlgOpenFile 
      Left            =   0
      Top             =   7680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "*.xls|*.xls|*.xlsx|*.xlsx"
   End
   Begin TabDlg.SSTab sstType 
      Height          =   7455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   13150
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "EXCEL"
      TabPicture(0)   =   "frmIllImport.frx":6852
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "�ű�"
      TabPicture(1)   =   "frmIllImport.frx":686E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame1 
         Height          =   6855
         Left            =   -74880
         TabIndex        =   6
         Top             =   420
         Width           =   10335
         Begin VB.CommandButton cmd 
            Caption         =   "�ļ�(&F)"
            Height          =   350
            Index           =   0
            Left            =   9000
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   615
            Width           =   1100
         End
         Begin VB.TextBox txt 
            Height          =   375
            Index           =   0
            Left            =   960
            TabIndex        =   9
            ToolTipText     =   "�����������EXCEL���·��"
            Top             =   600
            Width           =   7935
         End
         Begin VSFlex8Ctl.VSFlexGrid vsItem 
            Height          =   2295
            Index           =   0
            Left            =   120
            TabIndex        =   7
            Top             =   1440
            Width           =   10095
            _cx             =   17806
            _cy             =   4048
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
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   4
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   500
            ColWidthMax     =   10000
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
         Begin VSFlex8Ctl.VSFlexGrid vsItem 
            Height          =   2535
            Index           =   1
            Left            =   120
            TabIndex        =   12
            Top             =   4170
            Width           =   10095
            _cx             =   17806
            _cy             =   4471
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
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483633
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   255
            RowHeightMax    =   300
            ColWidthMin     =   500
            ColWidthMax     =   10000
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
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "˵��:EXCEL�ļ�����������������ࡿ�͡���������Ŀ¼��������������ʽ���·����ʾ����"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Index           =   1
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   7740
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "��������Ŀ¼���ʾ��"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Index           =   3
            Left            =   120
            TabIndex        =   13
            Top             =   3930
            Width           =   1800
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "�ļ�λ��"
            Height          =   180
            Index           =   0
            Left            =   120
            TabIndex        =   11
            Top             =   690
            Width           =   720
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "�������������ʾ��"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Index           =   2
            Left            =   120
            TabIndex        =   8
            Top             =   1200
            Width           =   1800
         End
      End
      Begin VB.Frame Frame3 
         Height          =   6735
         Left            =   120
         TabIndex        =   1
         Top             =   420
         Width           =   10335
         Begin VB.CommandButton cmd 
            Caption         =   "�ļ�(&F)"
            Height          =   350
            Index           =   1
            Left            =   9000
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   750
            Width           =   1100
         End
         Begin VB.TextBox txt 
            Height          =   375
            Index           =   1
            Left            =   960
            TabIndex        =   2
            ToolTipText     =   "�����������EXCEL���·��"
            Top             =   750
            Width           =   7935
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "˵��:��ѡ����Ҫִ�еĽű��ļ����ļ����ݱ������ɡ�����������������Ľű���ʽ��"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Index           =   4
            Left            =   120
            TabIndex        =   5
            Top             =   360
            Width           =   7110
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "�ļ�λ��"
            Height          =   180
            Index           =   5
            Left            =   120
            TabIndex        =   4
            Top             =   840
            Width           =   720
         End
      End
   End
End
Attribute VB_Name = "frmIllImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const mstrTYPE As String = "��,���뷶Χ,����"
Private Const mstrCONTENT As String = "����,����,����"

Private mconn As ADODB.Connection
Private mRsType As ADODB.Recordset
Private mrsContent As ADODB.Recordset
Private mrs��� As ADODB.Recordset
Private mbytModel  As Byte   '0-EXCEL��ʽ,���������;1-Excel��ʽ:�������;2-�Ų�����,���������;3-�Ų�����,�������

Private Enum E_ITEM
    E_���� = 0
    E_Ŀ¼ = 1
End Enum

Private Enum E_PAGE
    E_EXCEL = 0
    E_SCRIPT = 1
End Enum

Private Sub InitVsItem()
'����:��ʼ��ʾ�����
    Dim strHead As String
    Dim strRowContent As String 'strRowContent=����Ԥ����������,��ʽΪ����1,����1,��2,����2:��1;��1,����1,��2,����2:��2;
    
    strHead = "��,2000,4;���뷶Χ,2000,1;����,6000,1"
    
    strRowContent = "0,��һ��,1,A00-B99,2,ĳЩ��Ⱦ���ͼ����没;0,,1,A00-A09,2,������Ⱦ��;0,,1,A15-Al9,2,��˲�;0,,1,A20-A28,2,ĳЩ����Դ��ϸ���Լ���;" & vbCrLf & _
                    "0,,1,B99-B99,2,������Ⱦ��;" & vbCrLf & _
                    "0,�ڶ���,1,C00-D48,2,����;0,,1,C00-C14,2,������ǻ���ʶ�������;0,,1,C15-C26,2,�������ٶ�������;0,,1,C30-C39,2,��������ǻ�����ٶ�������"
    Grid.Init vsItem(E_����), strHead, strRowContent, 0, 1
    strHead = "����,2000,1;����,2000,1;����,6000,1"

    strRowContent = "0,A00.000,1,,2,�ŵ������ͻ���;0,A00.100,1,,2,�������ͻ���;" & vbCrLf & _
                    "0,A01.001+,1,K77.0*,2,�˺��Ը���;0,A01.100,1,,2,���˺���"
    Grid.Init vsItem(E_Ŀ¼), strHead, strRowContent, 0, 1
End Sub

Private Sub cmd_Click(Index As Integer)
    If Index = E_EXCEL Then
        OpenFile "EXCEL Files(*.xls,*.xlsx)|*.xls;*.xlsx", Index
    Else
        OpenFile "SQL Files(*.sql)|*.SQL", Index
    End If
    Set mconn = Nothing
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strFile As String, StrInfo As String
    Dim objFile As New FileSystemObject
    Dim i As Byte
    Dim objExcel As Object      'Excel.Application '����Excel��
    Dim objBook As Object        'Excel.Workbook '���幤������
    Dim objsheet As Object      'Excel.Worksheet '���幤������
    Dim arrTmp As Variant
    
    On Error GoTo errH
    Me.MousePointer = 11
    Debug.Print Now & vbCrLf
    
    StrInfo = "ִ�д˲���֮ǰ��ԡ�����������ࡿ�͡���������Ŀ¼�������ݽ������ݱ��ݡ�����һ���ǳ����������,���ȱ�����ִ�б�������" & vbCrLf & vbCrLf & _
             " �Ƿ����?"
    If MsgBox(StrInfo, vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbNo Then
        GoTo errHandle
    End If
        
    gstrSQL = "Select ����,�Ƿ���� From ����������� where ���� IN ('D','Y','M','S') order by ���ȼ�"
    Set mrs��� = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption)
        
    If sstType.Tab = E_EXCEL Then
    'EXCEL
    '����ļ��Ƿ�ѡ��
        If Trim(txt(E_EXCEL).Text) = "" Then
            MsgBox "��ѡ�����������������ࡿ�͡���������Ŀ¼����EXCEL�ļ���", vbOKOnly + vbInformation, gstrSysName
            cmd(E_EXCEL).SetFocus
            GoTo errHandle
        End If
        
        If Not objFile.FileExists(Trim(txt(E_EXCEL).Text)) Then
            MsgBox "���ļ������ڡ��ļ�λ��:" & Trim(txt(E_EXCEL).Text), vbOKOnly + vbInformation, gstrSysName
            txt(E_EXCEL).SetFocus
            GoTo errHandle
        End If
        On Error Resume Next
        Set objExcel = CreateObject("Excel.Application") '����ExcelӦ����
        If Err.Number <> 0 Then
            MsgBox "���鱾���Ƿ���ȷ��װEXCEL��", vbInformation + vbOKOnly, gstrSysName
            GoTo errHandle
        End If
        Err.Clear: On Error GoTo 0
        
        On Error GoTo errH
        
        arrTmp = Split("0,0", ",")
        objExcel.Visible = False '����Excel�ɼ�
        Set objBook = objExcel.Workbooks.Open(Trim(txt(E_EXCEL).Text)) '��Excel������
        For i = 1 To objBook.Worksheets.Count
            Set objsheet = objBook.Worksheets(i)  '��Excel������
            If objsheet.Name = "��������Ŀ¼" Then
                arrTmp(1) = 1
            ElseIf objsheet.Name = "�����������" Then
                arrTmp(0) = 1
            End If
        Next
        objBook.Close
        objExcel.Quit
        Set objExcel = Nothing
        
        If arrTmp(0) = 1 And arrTmp(1) = 1 Then
            mbytModel = 1
        ElseIf arrTmp(0) = 0 And arrTmp(1) = 1 Then
            If MsgBox("���ļ������ƽ��С���������Ŀ¼��,û�С�����������ࡿ������������ֻ���롾��������Ŀ¼�������ݡ�" & vbCrLf & vbCrLf & _
                        "�Ƿ������", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                GoTo errHandle
            End If
            mbytModel = 0
        Else
            MsgBox "����Exele�ļ�" & Trim(txt(E_EXCEL).Text) & vbCrLf & _
                "���ļ������Ʋ���������������Ŀ¼��,�޷�������һ��������", vbInformation + vbOKOnly, gstrSysName
            GoTo errHandle
        End If
        
        If Not InitOLEConn(Trim(txt(E_EXCEL).Text)) Then GoTo errHandle
        '�򿪼�¼��
        Set mRsType = New ADODB.Recordset
        Set mrsContent = New ADODB.Recordset
        On Error Resume Next
        If mbytModel = 1 Then
            mRsType.Open "Select [��],[���뷶Χ],[����] FROM [�����������$]", mconn, adOpenStatic, adLockOptimistic
            If Err.Number <> 0 Then
               MsgBox "�����:" & Err.Number & vbCrLf & "������Ϣ:" & vbCrLf & Err.Description, vbInformation + vbOKOnly, gstrSysName
               Err.Clear
               GoTo errHandle
            End If
        End If
        mrsContent.Open "Select [����],[����],[����] FROM [��������Ŀ¼$]", mconn, adOpenStatic, adLockOptimistic
        If Err.Number <> 0 Then
           MsgBox "�����:" & Err.Number & vbCrLf & "������Ϣ:" & vbCrLf & Err.Description, vbInformation + vbOKOnly, gstrSysName
           Err.Clear
           GoTo errHandle
        End If
        Err.Clear: On Error GoTo 0
        
        On Error GoTo errH
        If Not CheckRS() Then GoTo errHandle
        If Not FuncUpdateRS() Then GoTo errHandle
        Call SaveData(E_EXCEL)
    ElseIf sstType.Tab = E_SCRIPT Then
        '�Ų�
        strFile = Trim(txt(E_SCRIPT).Text)
        If strFile = "" Then
            MsgBox "��ѡ��������Ľű��ļ���", vbInformation, gstrSysName
            cmd(E_SCRIPT).SetFocus
            GoTo errHandle
        End If
        If strFile <> "" Then
            If Not FuncCreateRSBySQL(strFile) Then GoTo errHandle
            Call SaveData(E_SCRIPT)
        End If
    End If
    
    Debug.Print Now
errHandle:
    prg.Visible = False
    lblInfo.Caption = ""
    Me.MousePointer = 0
    Exit Sub
    
errH:
    lblInfo.Caption = ""
    prg.Visible = False
    Me.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    prg.Visible = False
    Call InitVsItem
End Sub

Public Sub ShowMe(ByVal frmParent As Form)

    Me.Show 1, frmParent
End Sub

Private Sub OpenFile(ByVal strFilter As String, ByVal intIndex As Integer)
    dlgOpenFile.Filter = strFilter
    dlgOpenFile.ShowOpen
    If dlgOpenFile.FileName <> "" Then
        txt(intIndex).Text = dlgOpenFile.FileName
    End If
    
End Sub

Public Function CheckRS() As Boolean
    Dim lngRow As Long, lngCol As Long
    Dim reg As New RegExp
    
    '���ݸ�ʽ���
    On Error GoTo errH
    prg.Visible = True
    If mbytModel = 1 Then
        With mRsType
            lblInfo.Caption = "������������ࡿ����ʽ���..."
            reg.Pattern = "^[A-Z]{1}[0-9]{2}-[A-Z]{1}[0-9]{2}$"
            reg.IgnoreCase = False
            '���뷶Χ,����,�����Ϊ��
            For lngRow = 1 To .RecordCount
                If Trim(!���뷶Χ & "") = "" Then
                    MsgBox "EXCEL�ļ�������������ࡿ,�����뷶Χ����ֵ����Ϊ�գ�" & "������:" & lngRow + 1, vbInformation, gstrSysName
                    CheckRS = False
                    Exit Function
                End If
    
                If Trim(!���� & "") = "" Then
                    MsgBox "EXCEL�ļ�������������ࡿ,�����ơ���ֵ����Ϊ�գ�" & "������:" & lngRow + 1, vbInformation, gstrSysName
                    CheckRS = False
                    Exit Function
                End If
                '���뷶Χ�����ʽ����
                If Not reg.Test(!���뷶Χ & "") Then
                    MsgBox "EXCEL�ļ�������������ࡿ,�����뷶Χ���С���" & lngRow + 1 & "��:" & vbCrLf & _
                            "���뷶Χ���ɼ��������ǰ��λ��1λ��д��ĸ��2λ���֣��ӷָ���""-""��ɡ�ʾ��:A00-B99", vbInformation, gstrSysName
                        CheckRS = False
                    Exit Function
                End If
                prg.value = Int((lngRow / .RecordCount) * 100)
                .MoveNext
            Next
        End With
    End If
    With mrsContent
        '��������˳����
        lblInfo.Caption = "����������Ŀ¼������ʽ���..."
        reg.Pattern = "^([A-Z]{1}[0-9]{2}.(([Xx0-9]{0,2}\+?)|([0-9]{0,3}\+?)))|(M[8-9]{1}[0-9]{4}/[01236])|([0-9]{2}.([0-9]{4}|[0-9]{5}))$"
        reg.IgnoreCase = False
        '����,����,�����Ϊ��
        For lngRow = 1 To .RecordCount
            If Trim(!���� & "") = "" Then
                MsgBox "EXCEL�ļ�����������Ŀ¼��,�����롿��ֵ����Ϊ�գ�" & "������:" & lngRow + 1, vbInformation, gstrSysName
                CheckRS = False
                Exit Function
            End If

            If Trim(!���� & "") = "" Then
                MsgBox "EXCEL�ļ�����������Ŀ¼��,�����ơ���ֵ����Ϊ�գ�" & "������:" & lngRow + 1, vbInformation, gstrSysName
                CheckRS = False
                Exit Function
            End If
            '���뷶Χ�����ʽ����
            If Not reg.Test(!���� & "") Then
                MsgBox "EXCEL�ļ�����������Ŀ¼��,�����롿�С���" & lngRow + 1 & "��:" & vbCrLf & _
                        "�����ʽ�������顣", vbInformation, gstrSysName
                    CheckRS = False
                Exit Function
            End If
            
            prg.value = Int((lngRow / .RecordCount) * 100)
            .MoveNext
        Next
    End With
    prg.Visible = False
    CheckRS = True
    Exit Function
errH:
    prg.Visible = False
    MsgBox Err.Description, vbInformation, gstrSysName
End Function

Public Function InitRS(Optional ByVal bytFunc As Byte = 0) As ADODB.Recordset
'����:����ҽ����¼
    Dim rs As ADODB.Recordset
    Dim strFields As String
    Dim strFieldName As String
    Dim lngLen As Long
    Dim FieldType As DataTypeEnum
    Dim i As Long, j As Long
    
    Dim arrField As Variant
    Dim arrSubFeld As Variant '�ֶ�����|�ֶ�����|�ֶγ��� ȱʡ�ֶ����� ΪadVarChar
    
    Select Case bytFunc
    
    Case 0
        strFields = "ID|adBigInt|18,�ϼ�ID|adBigInt|18,���||1,����||4000,���|adInteger|1"    '��� 0-Insert Into �� ;1-Select ��; 2-Select ��β��
    Case 1
        strFields = "ID|adBigInt|18,����ID|adBigInt|18,���||1,����||100,����||4000,���|adInteger|1,�Ƿ����|adInteger|1"  '�Ƿ���� 0-����ʾ����;1-��ʾ����
    End Select
    
    Set rs = New ADODB.Recordset
    '-----------------------------------------
    With rs.Fields
        arrField = Split(strFields, ",")
        For i = LBound(arrField) To UBound(arrField)
            arrSubFeld = Split(arrField(i), "|")
            strFieldName = arrSubFeld(0)
            If UCase(arrSubFeld(1) & "") = UCase("adVarChar") Then
                FieldType = adVarChar
            ElseIf UCase(arrSubFeld(1) & "") = UCase("adBigInt") Then
                FieldType = adBigInt
            ElseIf UCase(arrSubFeld(1) & "") = UCase("adInteger") Then
                FieldType = adInteger
            Else
                FieldType = adVarChar
            End If
            lngLen = Val(arrSubFeld(2))
            .Append strFieldName, FieldType, lngLen
        Next
    End With
    '---------------------------------------
    rs.CursorLocation = adUseClient
    rs.LockType = adLockOptimistic
    rs.CursorType = adOpenStatic
    rs.Open
    '----------------------------------
    Set InitRS = rs
End Function

Private Function FuncUpdateRS() As Boolean
    Dim lngRow As Long, lngLevel As Long
    Dim i As Long, lngPos As Long, j As Long, k As Long
    Dim rsTmp As ADODB.Recordset
    
    Dim lngNum As Long
    Dim strType As String, strTypes As String
    Dim strCode As String, strCodeA As String, strcodeB As String
    Dim strGroup As String
    Dim lngGroupId As Long
    Dim arrTmp As Variant
    Dim arrList As Variant
    Dim arrLevel As Variant
    
    On Error GoTo errH
    lblInfo.Caption = "������֯������������ࡿ������..."
    prg.Visible = True
    If mbytModel = 0 Or mbytModel = 2 Then
        gstrSQL = "Select a.Id, a.�ϼ�id, ���, Level, a.����, a.���뷶Χ, a.���, a.�Ƿ���, 0 As ����" & vbNewLine & _
                "From ����������� A, ����������� B" & vbNewLine & _
                "Where a.��� = b.���� And (a.����ʱ�� Is Null Or Trunc(a.����ʱ��) = To_Date('3000-01-01', 'YYYY-MM-dd')) And" & vbNewLine & _
                "      a.��� In ('D', 'Y', 'M', 'S') And b.�Ƿ���� = 1" & vbNewLine & _
                "Start With a.�ϼ�id Is Null" & vbNewLine & _
                "Connect By Prior ID = a.�ϼ�id"
        Call zldatabase.OpenRecordset(mRsType, gstrSQL, Me.Caption, adOpenStatic, adLockOptimistic)
        Set mRsType = zldatabase.CopyNewRec(mRsType) '���Ʊ��ں�������ֶ�ֵ
    End If
    
    If mbytModel = 1 Then
        Set mRsType = zldatabase.CopyNewRec(mRsType, , , Array("ID", adBigInt, 18, Empty, "�ϼ�ID", adBigInt, 18, Empty, "���", adInteger, 6, Empty, "����A", adVarChar, 60, Empty, _
                        "����B", adVarChar, 60, Empty, "���", adVarChar, 1, Empty, "�Ƿ���", adInteger, 1, Empty, "����", adInteger, 1, Empty, "Level", adInteger, 1, Empty))
        '׷���ֶ�
        strTypes = ""
        With mRsType
            .Filter = ""
            For lngRow = 1 To .RecordCount
                '���Ʊ���
                !ID = lngRow
                !���� = FuncGetStr(!����)
                !���뷶Χ = UCase(FuncGetStr(!���뷶Χ))
                
                !����A = UCase(Split(!���뷶Χ & "", "-")(0))
                !����B = UCase(Split(!���뷶Χ & "", "-")(1))
                !�Ƿ��� = 1 'Ĭ��Ϊ1; 0-������Чֻ��������
                !���� = 1
                !Level = 1
                !��� = FuncCheckType(!����A & "")
                If InStr("," & strTypes & ",", "," & !��� & ",") = 0 Then
                    strTypes = strTypes & "," & !���
                End If
                prg.value = Int((lngRow / .RecordCount) * 100)
                .MoveNext
            Next
            strTypes = Mid(strTypes, 2)
            
            If strTypes <> "" Then
                arrTmp = Split(strTypes, ",")
                For i = LBound(arrTmp) To UBound(arrTmp)
                    mRsType.Filter = "���='" & arrTmp(i) & "'"
                    lngNum = 0
                    For j = 1 To mRsType.RecordCount
                        lngNum = lngNum + 1
                        mRsType!��� = lngNum
                        mRsType.MoveNext
                    Next
                Next
            End If
            '���ݱ��뷶Χ�����ϼ�ID
            .Filter = ""
            For lngRow = 1 To .RecordCount
                lngPos = .AbsolutePosition
                If Trim(!�� & "") <> "" And strGroup <> Trim(!�� & "") Then
                    strGroup = Trim(!�� & "")
                    !�ϼ�id = 0
                    !Level = 1
                Else
                    strCodeA = !����A
                    strcodeB = !����B
                    Do While Not .BOF
                        .MovePrevious
                        If !�ϼ�id = 0 Then
                            lngGroupId = !ID
                            lngLevel = 1
                            .AbsolutePosition = lngPos
                            !�ϼ�id = lngGroupId
                            !Level = (lngLevel + 1)
                            Exit Do
                        End If
                        If strCodeA >= !����A And strcodeB <= !����B Then
                            lngGroupId = !ID
                            lngLevel = !Level
                            .AbsolutePosition = lngPos
                            !�ϼ�id = lngGroupId
                            !Level = (lngLevel + 1)
                            Exit Do
                        End If
                    Loop
                End If
                prg.value = Int((lngRow / .RecordCount) * 100)
                .MoveNext
            Next
        End With
    End If
    
    If mbytModel = 0 Or mbytModel = 1 Then
        lblInfo.Caption = "������֯����������Ŀ¼��������..."
        strCode = ""
        Set mrsContent = zldatabase.CopyNewRec(mrsContent, , , Array("ID", adBigInt, 18, Empty, "���", adInteger, 10, Empty, "����ID", adBigInt, 18, Empty, "�Ƿ����", adInteger, 3, _
                        Empty, "���", adVarChar, 1, Empty))
        strTypes = ""
        With mrsContent
            If .RecordCount > 0 Then .MoveFirst
            For lngRow = 1 To .RecordCount
                !ID = lngRow
                !��� = 1
                !����id = 0
                !���� = FuncGetStr(!����)
                !���� = UCase(FuncGetStr(!����))
                !��� = FuncCheckType(!���� & "")
                If strCode = !��� & "_" & !���� Then
                    !��� = lngNum + 1
                End If
                '��¼��һ�����뼰���
                strCode = !��� & "_" & !����
                lngNum = Val(!��� & "")
                If InStr("," & strTypes & ",", "," & !��� & ",") = 0 Then
                    strTypes = strTypes & "," & !���
                End If
                prg.value = Int((lngRow / .RecordCount) * 100)
                .MoveNext
            Next
        End With
        strTypes = strTypes & ","
    End If
    If mbytModel = 2 Then
        With mrsContent
            .Filter = "���>0"
            For i = 1 To .RecordCount
                !����id = 0
                prg.value = Int((lngRow / .RecordCount) * 100)
                .MoveNext
            Next
        End With
    End If
    lblInfo.Caption = "���ڸ��¡���������Ŀ¼���ġ�����ID��..."
    mrs���.Filter = ""
    Do While Not mrs���.EOF
        If Val(mrs���!�Ƿ���� & "") = 0 Then
            gstrSQL = "Select ID" & vbNewLine & _
                    "From ����������� A" & vbNewLine & _
                    "Where a.��� = [1] And (a.����ʱ�� Is Null Or Trunc(a.����ʱ��) = To_Date('3000-01-01', 'YYYY-MM-DD'))"
            Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, mrs���!����)
            If Not rsTmp.EOF Then lngNum = rsTmp!ID
            mrsContent.Filter = "��� ='" & mrs���!���� & "'"
            For j = 1 To mrsContent.RecordCount
                mrsContent!����id = lngNum
                mrsContent!�Ƿ���� = 1   '����������,����֮ǰ����
                mrsContent.MoveNext
            Next
        Else
            mRsType.Filter = "��� = '" & mrs���!���� & "'"
            mRsType.Sort = "Level desc"
            For i = 1 To mRsType.RecordCount
                arrList = Split(Trim(mRsType!���뷶Χ & ""), ",")
                'A15.0-A15.3,A16.0-A16.2
                For j = LBound(arrList) To UBound(arrList)
                    arrTmp = Split(Trim(arrList(j)), "-")
                    If UBound(arrTmp) = 1 Then
                        mrsContent.Filter = "���� >= '" & arrTmp(0) & "' And ���� <= '" & arrTmp(1) & "' And ����ID = 0 And  ��� ='" & mRsType!��� & "'"
                        For k = 1 To mrsContent.RecordCount
                            mrsContent!����id = mRsType!ID
                            mrsContent.MoveNext
                        Next
                        
                        mrsContent.Filter = "��� ='" & mRsType!��� & "' And ����ID = 0 And ���� like '" & arrTmp(1) & "%'"
                        For k = 1 To mrsContent.RecordCount
                            mrsContent!����id = mRsType!ID
                            mrsContent.MoveNext
                        Next
                        prg.value = Int((i / mRsType.RecordCount) * 100)
                    ElseIf UBound(arrTmp) = 0 Then
                        mrsContent.Filter = "��� ='" & mRsType!��� & "' And ����ID = 0 And ���� like '" & arrTmp(0) & "%'"
                        For k = 1 To mrsContent.RecordCount
                            mrsContent!����id = mRsType!ID
                            mrsContent.MoveNext
                        Next
                    End If
                Next
                mRsType.MoveNext
            Next
        End If
        mrs���.MoveNext
    Loop
    If mbytModel = 0 Or mbytModel = 1 Then
        mrsContent.Filter = "����ID = 0"
    ElseIf mbytModel = 2 Then
        mrsContent.Filter = "���>0 And ����ID = 0"
    End If
    If mrsContent.RecordCount > 0 Then
        strCode = ""
        For i = 1 To mrsContent.RecordCount
            strCode = strCode & "," & mrsContent!����
            If i > 10 Then strCode = strCode & "...": Exit For
            mrsContent.MoveNext
        Next
        strCode = Mid(strCode, 2)
        If MsgBox("����" & mrsContent.RecordCount & "�С���������Ŀ¼���ķ���ID�޷�ȷ��,��Ӧ�����롿����:" & strCode & vbCrLf & _
                "�Ƿ������" & vbCrLf & _
                "ѡ���ǡ��Ὣ�޷�ȷ�ϡ�����ID������Ŀ��ӵ�ָ�����ࡾ�������в�������һ��������" & vbCrLf & _
                "ѡ�񡾷񡿻���ֹ���β��������顾����������ࡿ�ġ����뷶Χ���е�ֵ�ܷ���������������Ŀ¼���ı���ֵ������", vbYesNo + vbDefaultButton2 + vbQuestion, gstrSysName) = vbYes Then
            mrsContent.MoveFirst

            For i = 1 To mrsContent.RecordCount
                mRsType.Filter = "���='" & mrsContent!��� & "' And ����='����' And �ϼ�ID = 0"
                If mRsType.RecordCount > 0 Then
                    lngGroupId = mRsType!ID
                Else
                    If mbytModel = 0 Or mbytModel = 2 Then
                        mRsType.Filter = "���='" & mrsContent!��� & "'"
                        mRsType.Sort = "��� desc"
                        If Not mRsType.EOF Then
                            lngNum = Val(mRsType!��� & "") + 1
                        Else
                            lngNum = (mRsType.RecordCount + 1)
                        End If
                    Else
                        mRsType.Filter = "���='" & mrsContent!��� & "'"
                        lngNum = (mRsType.RecordCount + 1)
                    End If
                    mRsType.Filter = ""
                    lngGroupId = mRsType.RecordCount + 1
                    mRsType.AddNew
                    mRsType!ID = lngGroupId
                    mRsType!�ϼ�id = 0
                    mRsType!��� = lngNum
                    mRsType!���� = "����"
                    mRsType!��� = mrsContent!���
                    mRsType!�Ƿ��� = 1
                    mRsType!���� = 1 '1-����
                    mRsType.Update
                End If
                mrsContent!����id = lngGroupId
                mrsContent.MoveNext
            Next
        Else
            prg.Visible = False
            Me.MousePointer = 0
            Exit Function
        End If
    End If
    mRsType.Filter = ""
    mRsType.Sort = ""
    prg.Visible = False
    FuncUpdateRS = True
    Exit Function
errH:
    Resume
    prg.Visible = False
    MsgBox Err.Description, vbInformation, gstrSysName
End Function

Private Function SaveData(ByVal bytFunc As Byte) As Boolean
'����:
'����:bytFunc=0 Excel����;=1 �Ų��ļ�����
    Dim colType As New Collection
    Dim colSQL As New Collection
    Dim lngId As Long, i As Long
    Dim strValue As String
    Dim strTitle As String
    Dim strTemp As String
    Dim blnOver As Boolean
    Dim datCurr As Date
    Dim arrTmp As Variant
    Dim strDate As String, strType As String
    Dim rsTmp As New ADODB.Recordset
    Dim blnTrans As Boolean
    Dim lngMin As Long, lngMax As Long
    
    On Error GoTo errH
    'GET ID
    prg.Visible = True
    lblInfo.Caption = "�������ɡ�����������ࡿ�ġ�ID�������ϼ�ID��..."
    If mbytModel < 3 Then
        mRsType.Filter = "���� =1 "
    Else
        mRsType.Filter = "���>0"
    End If
    '�����������δ�����������
    lngId = FuncGetNo("�����������", mRsType.RecordCount)  '��ǰ��ȡID
    For i = 1 To mRsType.RecordCount
        'Debug.Print mRsType!ID & "_" & mRsType!�ϼ�ID & "_" & mRsType!����
        colType.Add lngId, "_" & mRsType!ID: lngId = lngId + 1
        If mbytModel = 1 Or mbytModel = 3 Then
            If Not InStr("," & strType & ",", "," & UCase(mRsType!��� & "") & ",") > 0 Then
                strType = strType & "," & UCase(mRsType!��� & "")
            End If
        End If
        prg.value = Int((i / mRsType.RecordCount) * 100)
        mRsType.MoveNext
    Next
    
    If mbytModel = 0 Or mbytModel = 2 Then
    '���þɵķ���ID���������µ�ID
        mRsType.Filter = "���� = 0 "
        For i = 1 To mRsType.RecordCount
            colType.Add Val(mRsType!ID & ""), "_" & mRsType!ID
            prg.value = Int((i / mRsType.RecordCount) * 100)
            mRsType.MoveNext
        Next
    End If

    If mbytModel < 3 Then
        mRsType.Filter = "���� =1 "
    Else
        mRsType.Filter = "���>0"
    End If

    For i = 1 To mRsType.RecordCount
        mRsType!ID = colType("_" & mRsType!ID)
        If Val(mRsType!�ϼ�id & "") <> 0 Then
            mRsType!�ϼ�id = colType("_" & mRsType!�ϼ�id)
        End If
        prg.value = Int((i / mRsType.RecordCount) * 100)
        mRsType.MoveNext
    Next
    
    datCurr = zldatabase.Currentdate
    strDate = Format(datCurr, "yyyy-MM-dd HH:mm:ss")
    strType = Mid(strType, 2)
    If strType <> "" Then
        arrTmp = Split(strType, ",")
        For i = LBound(arrTmp) To UBound(arrTmp)
            gstrSQL = "Update ����������� " & vbNewLine & _
                        " Set ����ʱ�� = " & "To_Date('" & DateAdd("n", -1, datCurr) & "','YYYY-MM-DD HH24:MI:SS')" & vbNewLine & _
                        " Where ��� = '" & arrTmp(i) & "' And (����ʱ�� Is Null Or Trunc(����ʱ��) = To_Date('3000-01-01', 'YYYY-MM-DD'))"
            colSQL.Add gstrSQL
        Next
    End If
    
    strType = "": lblInfo.Caption = "�������ɡ���������Ŀ¼���ġ�ID��..."
    If bytFunc = 0 Then
        mrsContent.Filter = ""
    Else
        mrsContent.Filter = "���>0"
    End If
    lngId = FuncGetNo("��������Ŀ¼", mrsContent.RecordCount)  '��ǰ��ȡID
    For i = 1 To mrsContent.RecordCount
        mrsContent!ID = lngId: lngId = lngId + 1
        'û�з������Ŀ�Ѿ���ǰ����:M-������̬ѧ����
        If Val(mrsContent!�Ƿ���� & "") = 0 Then
            mrsContent!����id = colType("_" & mrsContent!����id)
        End If
        If Not InStr("," & strType & ",", "," & UCase(mrsContent!��� & "") & ",") > 0 Then
            strType = strType & "," & UCase(mrsContent!��� & "")
        End If
        prg.value = Int((i / mrsContent.RecordCount) * 100)
        mrsContent.MoveNext
    Next
    strType = Mid(strType, 2)
    If strType <> "" Then
        arrTmp = Split(strType, ",")
        For i = LBound(arrTmp) To UBound(arrTmp)
            gstrSQL = "Update ��������Ŀ¼ " & vbNewLine & _
                      "Set ����ʱ�� = " & "To_Date('" & DateAdd("n", -1, datCurr) & "','YYYY-MM-DD HH24:MI:SS')" & vbNewLine & _
                      "Where ��� = '" & arrTmp(i) & "' And (����ʱ�� Is Null Or Trunc(����ʱ��) = To_Date('3000-01-01', 'YYYY-MM-DD'))"
            colSQL.Add gstrSQL
        Next
    End If
    
    strDate = "To_Date('" & datCurr & "','YYYY-MM-DD HH24:MI:SS')"
    If mbytModel < 3 Then
        '���켲����������SQL
        strTitle = "Insert Into �����������(ID, �ϼ�id, ���, ����, ����, ���, ���뷶Χ, �Ƿ���, ����ʱ��) " & vbCrLf
        mRsType.Filter = "���� =1"
        lblInfo.Caption = "�������ɡ�����������ࡿ��SQL���..."
        strTemp = "": strValue = ""
        With mRsType
            For i = 1 To .RecordCount
                strTemp = "Select " & !ID & "," & IIF(Val(!�ϼ�id & "") = 0, "Null", !�ϼ�id) & "," & !��� & _
                    ",'" & Trim(Replace(!����, "'", "''")) & "','" & _
                    Mid(zlcommfun.SpellCode(Trim(Replace(!����, "'", "''")) & "��0"), 1, 20) & "','" & Trim(!���) & "','" & Trim(Replace(!���뷶Χ & "", "'", "''")) & "'," & _
                    Val(!�Ƿ��� & "") & "," & strDate & " From Dual UNION ALL" & vbCrLf
                If Len(strTitle & strValue & strTemp) > 100000 Then
                    strValue = Mid(strValue, 1, InStrRev(strValue, "UNION ALL") - 1) '& ";"
                    colSQL.Add strTitle & strValue
                    strValue = strTemp
                    blnOver = True
                Else
                    blnOver = False
                    strValue = strValue & strTemp
                End If
                .MoveNext
                If .EOF Then
                    If Not blnOver Then
                        strValue = Mid(strValue, 1, InStrRev(strValue, "UNION ALL") - 1) '& ";"
                        colSQL.Add strTitle & strValue
                        Exit For
                    End If
                End If
                prg.value = Int((i / .RecordCount) * 100)
            Next
        End With
    Else
        '���켲����������SQL
        mRsType.Filter = "���=0"
        strTitle = mRsType!���� & vbCrLf
        mRsType.Filter = "���>0"
        lblInfo.Caption = "�������ɡ�����������ࡿ��SQL���..."
        strValue = "": lngMin = 0: lngMax = 0
        With mRsType
            For i = 1 To .RecordCount
                If i = 1 Then lngMin = Val(!ID & "")
                If i = .RecordCount Then lngMax = Val(!ID & "")
                strValue = strValue & "Select " & !ID & "," & IIF(Val(!�ϼ�id & "") = 0, "Null", !�ϼ�id) & "," & !���� & vbCrLf
                If !��� = 2 Then
                    colSQL.Add strTitle & strValue
                    strValue = ""
                End If
                .MoveNext
                prg.value = Int((i / .RecordCount) * 100)
            Next
        End With
        If lngMin <> lngMax Then
            colSQL.Add "Update ����������� Set ����ʱ�� = " & strDate & " Where ID Between " & lngMin & " And " & lngMax
        End If
    End If
    
    If bytFunc = 0 Then
        'Update ��������Ŀ¼ Set ����� = ZLTOOLS.zlWbCode(����, 20);
        '���켲������Ŀ¼��SQL
        strTitle = "Insert Into ��������Ŀ¼ (ID, ����id, ����, ���, ����, ����, ����, �����, ���, ����ʱ��)" & vbCrLf
        
        lblInfo.Caption = "�������ɡ���������Ŀ¼����SQL���..."
        strTemp = "": strValue = ""
        With mrsContent
            .Filter = ""
            For i = 1 To .RecordCount
                strTemp = "Select " & !ID & "," & !����id & ",'" & !���� & "'," & !��� & ",'" & !���� & "','" & !���� & "','" & Mid(zlcommfun.SpellCode(Trim(Replace(!����, "'", "''")) & "��0"), 1, 20) & "','" & _
                            Mid(zlcommfun.SpellCode(Trim(Replace(!����, "'", "''")) & "��1"), 1, 20) & "','" & !��� & "'," & strDate & " From Dual UNION ALL" & vbCrLf
                If Len(strTitle & strValue & strTemp) > 100000 Then
                    strValue = Mid(strValue, 1, InStrRev(strValue, "UNION ALL") - 1) '& ";"
                    colSQL.Add strTitle & strValue
                    strValue = strTemp
                    blnOver = True
                Else
                    blnOver = False
                    strValue = strValue & strTemp
                End If
                .MoveNext
                If .EOF Then
                    If Not blnOver Then
                        strValue = Mid(strValue, 1, InStrRev(strValue, "UNION ALL") - 1) '& ";"
                        colSQL.Add strTitle & strValue
                        Exit For
                    End If
                End If
                prg.value = Int((i / .RecordCount) * 100)
            Next
         End With
    Else
        'Update ��������Ŀ¼ Set ����� = ZLTOOLS.zlWbCode(����, 20);
        '���켲������Ŀ¼��SQL
        mrsContent.Filter = "���=0"
        strTitle = mrsContent!���� & vbCrLf
        
        lblInfo.Caption = "�������ɡ���������Ŀ¼����SQL���..."
        strValue = "": lngMin = 0: lngMax = 0
        With mrsContent
            .Filter = "���>0"
            For i = 1 To .RecordCount
                If i = 1 Then lngMin = Val(!ID & "")
                If i = .RecordCount Then lngMax = Val(!ID & "")
                strValue = strValue & "Select " & !ID & "," & !����id & "," & !���� & vbCrLf
                If !��� = 2 Then
                    colSQL.Add strTitle & strValue
                    strValue = ""
                End If
                 .MoveNext
                prg.value = Int((i / .RecordCount) * 100)
            Next
            If lngMin <> lngMax Then
                colSQL.Add "Update ��������Ŀ¼ Set ����ʱ�� = " & strDate & " Where ID Between " & lngMin & " And " & lngMax
            End If
         End With
    End If

    '��������Ŀ¼ ��������� ����ͬ������²�����������,����ԭ���ݡ����ԭ���ݵķ����Ѿ�ͣ�����޸ķ���ID
    gstrSQL = "Zl_��������Ŀ¼_Redo(" & strDate & ",To_Date('" & DateAdd("n", -1, datCurr) & "','YYYY-MM-DD HH24:MI:SS'))"
    lblInfo.Caption = "�����ύ������������ࡿ������������Ŀ¼��������..."
    gcnOracle.BeginTrans: blnTrans = True
    For i = 1 To colSQL.Count
        Call zldatabase.OpenRecordset(rsTmp, CStr(colSQL(i)), Me.Caption)
        WriteLog vbCrLf & colSQL(i)
        prg.value = (Int(i / colSQL.Count) * 100)
    Next
    
    Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    WriteLog vbCrLf & gstrSQL
    gcnOracle.CommitTrans: blnTrans = False

        
    MsgBox "����ɹ�!", vbInformation, Me.Caption
    
    SaveData = True
    lblInfo.Caption = ""
    Me.MousePointer = 0
    prg.Visible = False
    Exit Function
errH:
    prg.Visible = False
    lblInfo.Caption = ""
    If blnTrans Then gcnOracle.RollbackTrans
    MsgBox Err.Description, vbInformation, gstrSysName
End Function

Public Function InitOLEConn(ByVal strFilePath As String) As Boolean
    Dim strConnect As String
    Dim objFile As New FileSystemObject
    
    On Error GoTo errH
    
    If mconn Is Nothing Then
        Set mconn = New ADODB.Connection
        If UCase(objFile.GetExtensionName(strFilePath)) = "XLS" Then
            strConnect = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & strFilePath & ";Extended Properties=""Excel 12.0;HDR=YES"""
        Else
            strConnect = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & strFilePath & ";Extended Properties=""Excel 12.0;HDR=YES"""
        End If
        mconn.ConnectionString = strConnect
    End If
    If mconn.State = adStateClosed Then mconn.Open
    InitOLEConn = True
    Exit Function
errH:
    If Err.Number = -2147467259 Then
        MsgBox "����EXCEL�ļ��Ƿ��Ѿ���", vbInformation, gstrSysName
    Else
        MsgBox Err.Description, vbInformation, gstrSysName
    End If
    Set mconn = Nothing
End Function

Private Sub Form_Unload(Cancel As Integer)
    If Not mconn Is Nothing Then
        If mconn.State = adStateOpen Then mconn.Close
        Set mconn = Nothing
    End If
    Set mrsContent = Nothing
    Set mRsType = Nothing
End Sub


Private Function FuncGetNo(ByVal strTable As String, ByVal lngCount As Long) As Long
'����:��ȡ����,����������
    Dim rsTemp As New ADODB.Recordset
    Dim lngMax As Long
    Dim lngCurr As Long
    Dim strOwner As String
    
    On Error GoTo errH
    gstrSQL = "Select Sequence_Owner From All_Sequences Where Sequence_Name = Upper('" & strTable & "_ID')"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    If Not rsTemp.EOF Then strOwner = rsTemp!Sequence_Owner & ""

    '��������
    gstrSQL = "Select Max(ID) as ID From " & strTable
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    If Not rsTemp.EOF Then lngMax = rsTemp!ID
    gstrSQL = "Select " & strOwner & "." & strTable & "_ID.Nextval As CurrID From Dual"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    If Not rsTemp.EOF Then lngCurr = rsTemp!CurrID
    If Abs(lngMax - lngCurr) > 1 Then
        '--�����ɷ�������
        gstrSQL = "Alter Sequence " & strOwner & "." & strTable & "_ID Increment By " & (lngMax - lngCurr)
        gcnOracle.Execute gstrSQL
        ' --�ƶ�һ������
        gstrSQL = "Select " & strOwner & "." & strTable & "_ID.Nextval From Dual"
        gcnOracle.Execute gstrSQL
        '--�ָ�ԭʼ����
        gstrSQL = "Alter Sequence " & strOwner & "." & strTable & "_ID Increment By 1"
        gcnOracle.Execute gstrSQL
    End If
    gstrSQL = "Select " & strOwner & "." & strTable & "_ID.Nextval AS NO From Dual Connect By Rownum <= [1]"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngCount)
    If Not rsTemp.EOF Then FuncGetNo = rsTemp!NO
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function FuncCheckType(ByVal strCode As String) As String
'����:���ݱ���ȷ�����
'
    Dim strTest As String
    Dim re As New RegExp
    
    '������̬ѧ����
    On Error Resume Next
    re.Pattern = "^(M[8-9]{1}[0-9]{4}/[01236])$"
    If re.Test(strCode) Then
        FuncCheckType = "M": Exit Function
    End If
    Err.Clear: On Error GoTo 0
    
    'Y-�����ж����ⲿԭ��(V01��Y98)
    strTest = Left(strCode, 3)
    If (strTest >= "V01" And strTest <= "Y98") Then
        FuncCheckType = "Y": Exit Function
    End If
    
    'ICD-10�ų������ж����ⲿԭ��
    If (strTest >= "A00" And strTest <= "Z99") And Not (strTest >= "V01" And strTest <= "Y98") Then
        FuncCheckType = "D": Exit Function
    End If

    'ICD-9-CM3��������
    strTest = Left(strCode, 2)
    If strTest >= "00" And strTest <= 99 Then
        FuncCheckType = "S": Exit Function
    End If
End Function

Private Function FuncCreateRSBySQL(ByVal strFile As String) As Boolean

    Dim StrInfo As String, strTXT As String
    Dim objFile As New FileSystemObject
    Dim objStream As TextStream
    Dim bytType As Byte '0-�����������;1-��������Ŀ¼
    Dim arrItem As Variant
    Dim lngNum As Long
    Dim j As Long
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errH
    
    Set mRsType = InitRS(0)
    Set mrsContent = InitRS(1)
    If Not objFile.FileExists(strFile) Then
        MsgBox "��ǰ�ļ�:" & strFile & "�����ڡ�", vbInformation, gstrSysName
        Exit Function
    End If
            
    Set objStream = objFile.OpenTextFile(strFile, ForReading)
    Do While Not objStream.AtEndOfStream
        strTXT = Trim(objStream.ReadLine)
        If InStr(UCase(strTXT), UCase("Insert Into")) > 0 And InStr(strTXT, "��������Ŀ¼") > 0 Then
            bytType = 1
            mrsContent.Filter = "���=0"
            If mrsContent.RecordCount = 0 Then
                mrsContent.AddNew
                mrsContent!��� = 0
                mrsContent!���� = Trim(strTXT)
            End If
        ElseIf InStr(UCase(strTXT), UCase("Insert Into")) > 0 And InStr(strTXT, "�����������") > 0 Then
            bytType = 0
            mRsType.Filter = "���=0"
            If mRsType.RecordCount = 0 Then
                mRsType.AddNew
                mRsType!��� = 0
                mRsType!���� = Trim(strTXT)
            End If
        ElseIf InStr(UCase(strTXT), UCase("Select")) > 0 And InStr(UCase(strTXT), UCase("From Dual UNION ALL")) > 0 Then
            arrItem = Split(strTXT, ",")
            If bytType = 0 Then
                mRsType.AddNew
                mRsType!ID = Val(Replace(UCase(arrItem(0)), UCase("Select "), ""))
                mRsType!�ϼ�id = IIF(UCase(arrItem(1)) = "NULL", 0, Val(arrItem(1)))
                mRsType!��� = Replace(arrItem(2), "'", "")
                mRsType!���� = Mid(strTXT, InStr(strTXT, arrItem(0) & "," & arrItem(1) & ",") + Len(arrItem(0) & "," & arrItem(1) & ","))
                mRsType!��� = 1
            Else
                mrsContent.AddNew
                mrsContent!ID = Val(Replace(UCase(arrItem(0)), UCase("Select"), ""))
                mrsContent!����id = Val(arrItem(1))
                mrsContent!��� = Replace(arrItem(2), "'", "")
                mrsContent!���� = Replace(arrItem(3), "'", "")
                mrsContent!���� = Mid(strTXT, InStr(strTXT, arrItem(0) & "," & arrItem(1) & ",") + Len(arrItem(0) & "," & arrItem(1) & ","))
                mrsContent!��� = 1
            End If
        ElseIf InStr(UCase(strTXT), UCase("Select")) > 0 And InStr(UCase(strTXT), UCase("From Dual")) > 0 And InStr(UCase(strTXT), UCase("UNION ALL")) = 0 Then
            If Right(strTXT, 1) = ";" Then strTXT = Left(strTXT, Len(strTXT) - 1)
            arrItem = Split(strTXT, ",")
            If bytType = 0 Then
                mRsType.AddNew
                mRsType!ID = Val(Replace(UCase(arrItem(0)), UCase("Select"), ""))
                mRsType!�ϼ�id = IIF(UCase(arrItem(1)) = "NULL", 0, Val(arrItem(1)))
                mRsType!��� = Replace(arrItem(2), "'", "")
                mRsType!���� = Mid(strTXT, InStr(strTXT, arrItem(0) & "," & arrItem(1) & ",") + Len(arrItem(0) & "," & arrItem(1) & ","))
                mRsType!��� = 2
                mRsType.UpdateBatch
            Else
                mrsContent.AddNew
                mrsContent!ID = Val(Replace(UCase(arrItem(0)), UCase("Select"), ""))
                mrsContent!����id = Val(arrItem(1))
                mrsContent!��� = Replace(arrItem(2), "'", "")
                mrsContent!���� = Replace(arrItem(3), "'", "")
                mrsContent!���� = Mid(strTXT, InStr(strTXT, arrItem(0) & "," & arrItem(1) & ",") + Len(arrItem(0) & "," & arrItem(1) & ","))
                mrsContent!��� = 2
                mrsContent.UpdateBatch
            End If
        End If
    Loop
    mRsType.Filter = ""
    If mRsType.RecordCount = 0 Then
        mbytModel = 2
        If Not FuncUpdateRS() Then Exit Function
    Else
        mbytModel = 3
    End If
    
    FuncCreateRSBySQL = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

