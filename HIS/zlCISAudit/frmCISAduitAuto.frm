VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCISAduitAuto 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�Զ����"
   ClientHeight    =   7560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10770
   Icon            =   "frmCISAduitAuto.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   10770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.PictureBox picPane 
      BackColor       =   &H80000015&
      BorderStyle     =   0  'None
      Height          =   5175
      Index           =   2
      Left            =   90
      ScaleHeight     =   5175
      ScaleWidth      =   2880
      TabIndex        =   19
      Top             =   780
      Width           =   2880
      Begin MSComctlLib.TreeView tvw 
         Height          =   5145
         Left            =   15
         TabIndex        =   20
         Top             =   15
         Width           =   2850
         _ExtentX        =   5027
         _ExtentY        =   9075
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   494
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         Appearance      =   0
      End
   End
   Begin VB.CommandButton CmdNot 
      Caption         =   "ȫ��"
      Height          =   495
      Left            =   2370
      Picture         =   "frmCISAduitAuto.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   18
      TabStop         =   0   'False
      ToolTipText     =   "ȫ��"
      Top             =   6015
      Width           =   570
   End
   Begin VB.CommandButton cmdAll 
      Caption         =   "ȫѡ"
      Height          =   495
      Left            =   1755
      Picture         =   "frmCISAduitAuto.frx":00E0
      Style           =   1  'Graphical
      TabIndex        =   17
      TabStop         =   0   'False
      ToolTipText     =   "ȫѡ"
      Top             =   6015
      Width           =   570
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   9570
      TabIndex        =   2
      Top             =   7020
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   8280
      TabIndex        =   1
      Top             =   7020
      Width           =   1100
   End
   Begin VB.Frame fraDetail 
      Caption         =   "��������"
      Height          =   690
      Left            =   75
      TabIndex        =   0
      Top             =   30
      Width           =   10590
      Begin VB.TextBox txtסԺ�� 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         IMEMode         =   3  'DISABLE
         Left            =   8775
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   12
         Top             =   315
         Width           =   1410
      End
      Begin VB.TextBox txt���� 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Left            =   4800
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   6
         Top             =   315
         Width           =   1410
      End
      Begin VB.TextBox txt���� 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Left            =   645
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   5
         Top             =   315
         Width           =   1410
      End
      Begin VB.TextBox txtסԺ���� 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         IMEMode         =   3  'DISABLE
         Left            =   6855
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   315
         Width           =   285
      End
      Begin VB.TextBox txt�Ա� 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         IMEMode         =   3  'DISABLE
         Left            =   2760
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   3
         Top             =   315
         Width           =   1410
      End
      Begin VB.Line Line2 
         X1              =   6855
         X2              =   7140
         Y1              =   495
         Y2              =   495
      End
      Begin VB.Line Line1 
         X1              =   8760
         X2              =   10200
         Y1              =   495
         Y2              =   495
      End
      Begin VB.Line Line5 
         X1              =   4815
         X2              =   6255
         Y1              =   495
         Y2              =   495
      End
      Begin VB.Line Line4 
         X1              =   2745
         X2              =   4185
         Y1              =   495
         Y2              =   495
      End
      Begin VB.Line Line3 
         X1              =   630
         X2              =   2070
         Y1              =   495
         Y2              =   495
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         ForeColor       =   &H00404040&
         Height          =   180
         Index           =   7
         Left            =   4320
         TabIndex        =   11
         Top             =   345
         Width           =   360
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Ա�"
         ForeColor       =   &H00404040&
         Height          =   180
         Index           =   5
         Left            =   2265
         TabIndex        =   10
         Top             =   345
         Width           =   360
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         ForeColor       =   &H00404040&
         Height          =   180
         Index           =   4
         Left            =   180
         TabIndex        =   9
         Top             =   345
         Width           =   360
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��    ��סԺ"
         ForeColor       =   &H00404040&
         Height          =   180
         Index           =   3
         Left            =   6645
         TabIndex        =   8
         Top             =   345
         Width           =   1080
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "סԺ��"
         ForeColor       =   &H00404040&
         Height          =   180
         Index           =   0
         Left            =   8145
         TabIndex        =   7
         Top             =   345
         Width           =   540
      End
   End
   Begin MSComctlLib.ProgressBar pbrBar 
      Height          =   345
      Left            =   2250
      TabIndex        =   13
      Top             =   7020
      Width           =   4560
      _ExtentX        =   8043
      _ExtentY        =   609
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.CommandButton cmdAuto 
      Caption         =   "�Զ�(&A)"
      Height          =   350
      Left            =   6930
      TabIndex        =   14
      Top             =   7020
      Width           =   1100
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "��ֹ(&S)"
      Height          =   350
      Left            =   6930
      TabIndex        =   15
      Top             =   7020
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfFeedback 
      Height          =   5730
      Left            =   3030
      TabIndex        =   21
      Top             =   780
      Width           =   7635
      _cx             =   13467
      _cy             =   10107
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
      BackColorSel    =   16772055
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
      ScrollTrack     =   -1  'True
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
   Begin VB.Label labShow 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   90
      TabIndex        =   22
      Top             =   6030
      Width           =   1275
   End
   Begin VB.Shape shpStatus 
      BorderColor     =   &H00000000&
      FillColor       =   &H000000FF&
      FillStyle       =   7  'Diagonal Cross
      Height          =   480
      Left            =   90
      Top             =   6030
      Width           =   1275
   End
   Begin VB.Line Line7 
      BorderColor     =   &H80000005&
      X1              =   -210
      X2              =   10770
      Y1              =   6720
      Y2              =   6735
   End
   Begin VB.Line Line6 
      X1              =   -225
      X2              =   10770
      Y1              =   6705
      Y2              =   6705
   End
   Begin VB.Label LabStatus 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   180
      Left            =   150
      TabIndex        =   16
      Top             =   7095
      Visible         =   0   'False
      Width           =   2025
   End
End
Attribute VB_Name = "frmCISAduitAuto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlng�ύId              As Long             '����Id
Private mlng����ID              As Long             '����Id
Private mlng��ҳID              As Long             '��ҳId
Private mlng����ID              As Long             '����Id
Private mblnCancel              As Boolean          'ȷ�� or ȡ��
Private zlCheck                 As New clsCheck     '���
Private mstrTreeSelect          As String           '��ѡ�е�Key
Private mblnStop                As Boolean          '�Զ�ʱֹͣ
Private mblnChecked             As Boolean          'ѡ��/ȡ��ѡ��
Private mstrSortID              As String
Private mstrLink                As String           '1����� 2�����

'�����ֶ�
Const con_vsfField = "/*+ rule */rownum as Id,'' as ѡ��,���Id,�ύId,����Id,��ҳID,��������,�ļ�Id,ҽ��Id,����Id,��¼����,��¼״̬,������,����ʱ��,��������,�������,������ĿID,���ĵ�ID"

Public Property Get blnCancel() As Boolean
    blnCancel = mblnCancel
End Property

Public Property Let blnCancel(ByVal vNewValue As Boolean)
    mblnCancel = vNewValue
End Property

Public Property Get lng�ύId() As Long
    lng�ύId = mlng�ύId
End Property

Public Property Let lng�ύId(ByVal vNewValue As Long)
    mlng�ύId = vNewValue
End Property

Public Property Get lng����id() As Long
    lng����id = mlng����ID
End Property

Public Property Let lng����id(ByVal vNewValue As Long)
    mlng����ID = vNewValue
End Property

Public Property Get lng��ҳID() As Long
    lng��ҳID = mlng��ҳID
End Property

Public Property Let lng��ҳID(ByVal vNewValue As Long)
    mlng��ҳID = vNewValue
End Property

Public Property Get lng����ID() As Long
    lng����ID = mlng����ID
End Property

Public Property Let lng����ID(ByVal vNewValue As Long)
    mlng����ID = vNewValue
End Property

Public Property Get strTreeSelect() As String
    strTreeSelect = mstrTreeSelect
End Property

Public Property Let strTreeSelect(ByVal vNewValue As String)
    mstrTreeSelect = vNewValue
End Property

Public Property Get strLink() As String
    strLink = mstrLink
End Property

Public Property Let strLink(ByVal vNewValue As String)
    mstrLink = vNewValue
End Property

'==============================================================================
'=���ܣ� ȫѡ
'==============================================================================
Private Sub cmdAll_Click()
    Call AllNot
End Sub

'==============================================================================
'=���ܣ� ȫ��
'==============================================================================
Private Sub CmdNot_Click()
    Call AllNot(False)
End Sub

'==============================================================================
'=���ܣ� ȫѡ��ȫ��
'==============================================================================
Private Sub AllNot(Optional blnAll As Boolean = True)
    Dim i           As Long
    On Error GoTo ErrH
    For i = 1 To tvw.Nodes.count
        tvw.Nodes.Item(i).Checked = blnAll
    Next
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


'==============================================================================
'=���ܣ� ��������
'==============================================================================
Private Sub CmdOK_Click()
    Dim i       As Long

    On Error GoTo ErrH
    If vsfFeedback.Rows <= 1 Then
        zlCheck.Msg_OK "û���������ݣ�����ȡ����"
        Exit Sub
    End If
    gstrSQL = ""
    With gcnOracle
        .BeginTrans
        With vsfFeedback
            For i = 1 To .Rows - 1
                If .TextMatrix(i, .ColIndex("ѡ��")) Then
                    gstrSQL = "zl_����������¼_Update (" & zlDatabase.GetNextId("����������¼") & ",Null," & IIf(lng�ύId <= 0, "Null", lng�ύId) & "," & lng����id & "," & _
                              "" & lng��ҳID & "," & AppObject(.TextMatrix(i, .ColIndex("��������")), False) & ",'" & .TextMatrix(i, .ColIndex("�ļ�ID")) & "','" & .TextMatrix(i, .ColIndex("�������")) & "'," & _
                              "" & .TextMatrix(i, .ColIndex("������ĿID")) & ",'" & .TextMatrix(i, .ColIndex("������")) & "',to_date('" & .TextMatrix(i, .ColIndex("����ʱ��")) & "','yyyy-mm-dd hh24:mi:ss'),to_date('" & .TextMatrix(i, .ColIndex("��������")) & "','yyyy-mm-dd hh24:mi:ss')," & _
                              "" & "Null," & .TextMatrix(i, .ColIndex("����Id")) & ",null,null,null,null,null,null,'" & .TextMatrix(i, .ColIndex("���ĵ�ID")) & "')"
                    zlDatabase.ExecuteProcedure gstrSQL, Me.Name
                End If
            Next
        End With
        .CommitTrans
    End With
    blnCancel = False
    Unload Me
    Exit Sub
ErrH:
    If gcnOracle.Errors.count > 0 Then
        gcnOracle.RollbackTrans
    End If
    If ErrCenter() = 1 Then Resume
   
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ֹͣ
'==============================================================================
Private Sub cmdStop_Click()
    mblnStop = True
End Sub

'==============================================================================
'=���ܣ� �Զ�����
'==============================================================================
Private Sub cmdAuto_Click()
Dim i   As Integer, j   As Integer
Dim strKey          As String
Dim varSplit        As Variant, varTmp         As Variant
Dim rsTmp   As ADODB.Recordset, rsFeed  As ADODB.Recordset
Dim strSql          As String
Dim datFeed         As Date, datFeedBack    As Date
Dim intRow          As Integer, strSource      As String
Dim strDocid As String, strSubDocid As String, strReturn As String, strMid As String, strAlidin As String
    
    On Error GoTo ErrH
    
    If zlCheck.Msg_OKC("ȷ�Ͻ����Զ�������" & vbCrLf & "�Զ�������ձ�����������ݣ�") Then Exit Sub
    LockWindowUpdate vsfFeedback.hWnd
    
    cmdAuto.Visible = False
    cmdStop.Visible = True
    cmdOK.Enabled = False
    cmdCancel.Enabled = False
    pbrBar.Visible = True
    LabStatus.Visible = True
    
    
    datFeed = zlDatabase.Currentdate
    i = Val(GetPara("������������", 1560))
    datFeedBack = DateAdd("D", i, datFeed)
    LabStatus.Caption = "���ڶ�ȡ���ö���..."
    LabStatus.BackColor = vbRed
    DoEvents
    Sleep (1000)
    
    pbrBar.Max = tvw.Nodes.count
    For i = 1 To tvw.Nodes.count
        pbrBar.Value = i
        DoEvents
        If mblnStop Then
            cmdAuto.Visible = True
            cmdOK.Enabled = True
            cmdCancel.Enabled = True
            LabStatus.Visible = False
            cmdStop.Visible = False
            pbrBar.Visible = False
            mblnStop = False
            Call zlCheck.Msg_OK("�����Զ�������;ȡ��������ɲ�������", vbCritical)
            LockWindowUpdate 0
            Exit Sub
        End If
        If tvw.Nodes.Item(i).Checked Then
            strKey = strKey & tvw.Nodes.Item(i).Key & strSplitCmb
        End If
    Next
    If strKey = "" Then
        zlCheck.Msg_OK ("��ǰδѡ�����ö���")
        cmdAuto.Visible = True
            cmdOK.Enabled = True
            cmdCancel.Enabled = True
            LabStatus.Visible = False
            cmdStop.Visible = False
            pbrBar.Visible = False
            mblnStop = False
        Exit Sub
    End If
    varSplit = Split(strKey, strSplitCmb)
    
    strKey = ""
    LabStatus.Caption = "���ڷ�����Ӧ����..."
    LabStatus.BackColor = vbYellow
    DoEvents
    Sleep (1000)
    
    If UBound(varSplit) > 1 Then pbrBar.Max = UBound(varSplit) - 1
    For i = 0 To UBound(varSplit) - 1
        pbrBar.Value = i
        DoEvents
        If mblnStop Then
            cmdAuto.Visible = True
            cmdOK.Enabled = True
            cmdCancel.Enabled = True
            
            LabStatus.Visible = False
            cmdStop.Visible = False
            pbrBar.Visible = False
            mblnStop = False
            Call zlCheck.Msg_OK("�����Զ�������;ȡ��������ɲ�������", vbCritical)
            LockWindowUpdate 0
            Exit Sub
        End If
        If Len(varSplit(i)) = 2 Then
            strKey = strKey & varSplit(i) & "[O]" & Mid(varSplit(i), 2, 1) & "[F][D],"
             
        ElseIf InStr(1, varSplit(i), "R4") > 0 Then
            '���˻����¼��ֱ��Ϊ�ļ�Id
            varSplit(i) = Replace(varSplit(i), "K", ",")
            varTmp = Split(varSplit(i), ",")
            If UBound(varTmp) > 1 Then
                strKey = strKey & Left(varSplit(i), 2) & "_" & varTmp(1) & "[O]" & Mid(varSplit(i), 2, 1) & "[F]" & varTmp(1) & "[D]" & varTmp(3) & ","
            End If
        ElseIf InStr("R2R3R6R7R8", Left(varSplit(i), 2)) > 0 Then
            '���Ӳ�����¼�в���Id
            varSplit(i) = Replace(varSplit(i), "K", ",")
            varTmp = Split(varSplit(i), ",")
            If UBound(varTmp) > 1 Then
                strSql = "select �ļ�Id from ���Ӳ�����¼ where Id = [1] "
                Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Name, Val(varTmp(1)))
                If Not zlCheck.Connection_ChkRsState(rsTmp) Then
                    strKey = strKey & Left(varSplit(i), 2) & "_" & rsTmp.Fields(0) & "[O]" & Mid(varSplit(i), 2, 1) & "[F]" & varTmp(1) & "[D],"
                End If
            End If
        ElseIf InStr(varSplit(i), "R") = 0 Then
            If gobjEmr Is Nothing Then
                MsgBox "����δ��װ������������ܽ����Զ���飬���飡", vbInformation, gstrSysName
                mblnStop = True
            Else
                If InStr(tvw.Nodes(varSplit(i)).Tag, "|") = 0 Then
                    strDocid = varSplit(i)
                    strSubDocid = ""
                Else
                    strDocid = Split(tvw.Nodes(varSplit(i)).Tag, "|")(0)
                    strSubDocid = Split(tvw.Nodes(varSplit(i)).Tag, "|")(1)
                End If
                strSql = "Select RawtoHex(Antetype_id) as ID From bz_doc_Tasks Where Real_Doc_Id = Hextoraw(:rdid)" & IIf(strSubDocid = "", "", " And subdoc_id=:sdid")
                strReturn = gobjEmr.OpenSQLRecordset(strSql, strDocid & "^" & DbType.T_String & "^rdid" & IIf(strSubDocid = "", "", "|" & strSubDocid & "^" & DbType.T_String & "^sdid"), rsTmp)
                If strReturn = "" Then
                If Not rsTmp.EOF Then
                    strKey = strKey & tvw.Nodes(varSplit(i)).Parent.Key & "_" & rsTmp.Fields(0) & "[O]" & Mid(tvw.Nodes(varSplit(i)).Parent.Key, 2, 1) & "[F]" & tvw.Nodes(varSplit(i)).Tag & "[D],"
                End If
                End If
            End If
        End If
    Next
    strKey = Left(strKey, Len(strKey) - 1)
    '��ȡ�������
    strSource = "" & _
            "select x.id,x.����id,x.����,x.����,x.����,x.˵��,x.���ö���,x.���û���,x.�������,'' as �ļ�ID  from �������Ŀ¼  x,������鷽�� C,���������� B where nvl(�ļ�ID,'') is null And  B.����id = C.id and B.id =x.����ID And  C.����ʱ�� is not null And x.���û��� = 0 or x.���û��� = [2] union all " & vbCrLf & _
            "select x.id,x.����id,x.����,x.����,x.����,x.˵��,x.���ö���,x.���û���,x.�������,y.column_value as �ļ�ID  from �������Ŀ¼  x , ������鷽�� C,���������� B,table (Cast(f_Str2List( x.�ļ�ID) As zlTools.t_StrList)) y" & vbCrLf & _
            "Where (X.���û��� = 0 Or X.���û��� = [2])  and  B.����id = C.id and B.id =X.����ID And  C.����ʱ�� is not null "
    
    strSql = "" & _
            "Select a.Id,a.�������,b.���ö���,Decode(Length(b.�ļ�id), 65, Substr(b.�ļ�id, 1, 32), b.�ļ�id) As �ļ�id,b.����Id,Decode(Length(b.�ļ�id), 65, Substr(b.�ļ�id, 34), b.�ļ�id) As ���ĵ�id from (" & strSource & ") a," & vbCrLf & _
            "(" & vbCrLf & _
            "   Select" & vbCrLf & _
            "   SUBSTR(COLUMN_VALUE,1,INSTR(COLUMN_VALUE,'[O]')-1) AS Id," & vbCrLf & _
            "   SUBSTR(COLUMN_VALUE,INSTR(COLUMN_VALUE,'[O]')+length('[O]'),case when (INSTR(COLUMN_VALUE,'[F]'))-(INSTR(COLUMN_VALUE,'[O]')+length('[O]'))<0 then 1000 else (INSTR(COLUMN_VALUE,'[F]'))-(INSTR(COLUMN_VALUE,'[O]')+length('[O]')) end) as ���ö���," & vbCrLf & _
            "   SUBSTR(COLUMN_VALUE,INSTR(COLUMN_VALUE,'[F]')+length('[F]'),case when (INSTR(COLUMN_VALUE,'[D]'))-(INSTR(COLUMN_VALUE,'[F]')+length('[F]'))<0 then 1000 else (INSTR(COLUMN_VALUE,'[D]'))-(INSTR(COLUMN_VALUE,'[F]')+length('[F]')) end) as �ļ�Id," & vbCrLf & _
            "   SUBSTR(COLUMN_VALUE,INSTR(COLUMN_VALUE,'[D]')+length('[D]')) AS ����Id" & vbCrLf & _
            "   From Table (Cast(f_Str2List([1]) As zlTools.t_StrList))" & vbCrLf & _
            ")b" & vbCrLf & _
            "Where 'R' || to_char(a.���ö���) || Case When nvl(a.�ļ�Id,'0')='0' Then '' Else '_' || a.�ļ�Id End = b.Id And a.������� is not null"
    If Len(strKey) >= 4000 Then
        cmdAuto.Visible = True
        cmdOK.Enabled = True
        cmdCancel.Enabled = True
        cmdStop.Visible = False
        pbrBar.Visible = False
        LabStatus.Visible = False
        zlCheck.Msg_OK "ѡ����Ŀ���࣬��ȡ��������Ŀ��"
        LockWindowUpdate 0
        Exit Sub
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Name, strKey, mstrLink)
    i = 0
    LabStatus.Caption = "�������ɷ�����Ϣ..."
    LabStatus.BackColor = vbGreen
    DoEvents
    Sleep (1000)
    
    If rsTmp.RecordCount > 0 Then pbrBar.Max = rsTmp.RecordCount
    
    intRow = 0
    Do While Not zlCheck.Connection_ChkRsState(rsTmp)
        pbrBar.Value = Val(rsTmp.Bookmark)
        DoEvents
        If mblnStop Then
            cmdAuto.Visible = True
            cmdOK.Enabled = True
            cmdCancel.Enabled = True
            
            LabStatus.Visible = False
            LabStatus.Visible = False
            cmdStop.Visible = False
            pbrBar.Visible = False
            mblnStop = False
            Call zlCheck.Msg_OK("�����Զ�������;ȡ��������ɲ�������", vbCritical)
            LockWindowUpdate 0
            Exit Sub
        End If
        With vsfFeedback
            .Rows = rsTmp.RecordCount + 10
            If Len(NVL(rsTmp!�ļ�ID)) < 32 Or InStr(NVL(rsTmp!ID), "R") > 0 Then
                strSql = CheckAuditSql_OUT(rsTmp!�������, lng����id, lng��ҳID)
                Set rsFeed = zlDatabase.OpenSQLRecord("select ZL_FUN_ExecSql('" & Replace(strSql, "'", "''") & "') from dual", "mdlCISAudit")
            ElseIf Not gobjEmr Is Nothing Then
                If strMid = "" Then Call GetEMR_MID_ALIDIN(lng����id, lng��ҳID, strMid, strAlidin) 'ȡ�²�������ID,�ID
                strSql = Replace(rsTmp!�������, "[MID]", ":mid")
                strSql = Replace(rsTmp!�������, "[ALIDIN]", ":alidin")
                strReturn = gobjEmr.OpenSQLRecordset(strSql, IIf(strMid = "", "", strMid & "^" & DbType.T_String & "^mid") & IIf(strAlidin = "", "", IIf(strMid = "", "", "|") & strAlidin & "^" & DbType.T_String & "^alidin"), rsFeed)
                If strReturn <> "" Then Set rsFeed = New ADODB.Recordset
            End If
            
            If Not zlCheck.Connection_ChkRsState(rsFeed) Then
            If InStr(1, rsFeed.Fields(0), "[zlsoft]Error[zlsoft]") = 0 Then
                If Trim("" & rsFeed.Fields(0)) <> "" Then

                    intRow = intRow + 1
                   
                    .Cell(flexcpAlignment, intRow, .ColIndex("ѡ��")) = flexAlignCenterCenter
                    .Cell(flexcpAlignment, intRow, .ColIndex("�������")) = flexAlignLeftCenter
                    .Cell(flexcpAlignment, intRow, .ColIndex("��������")) = flexAlignLeftCenter
                    .Cell(flexcpAlignment, intRow, .ColIndex("�ļ�Id")) = flexAlignLeftCenter
                    .Cell(flexcpAlignment, intRow, .ColIndex("����Id")) = flexAlignLeftCenter
                    .Cell(flexcpAlignment, intRow, .ColIndex("����Id")) = flexAlignLeftCenter
                    .Cell(flexcpAlignment, intRow, .ColIndex("��ҳId")) = flexAlignLeftCenter
                    .Cell(flexcpAlignment, intRow, .ColIndex("������")) = flexAlignLeftCenter
                    .Cell(flexcpAlignment, intRow, .ColIndex("����ʱ��")) = flexAlignLeftCenter
                    .Cell(flexcpAlignment, intRow, .ColIndex("��������")) = flexAlignLeftCenter
                    
                    .TextMatrix(intRow, .ColIndex("ѡ��")) = True
                    .TextMatrix(intRow, .ColIndex("�������")) = "" & rsFeed.Fields(0)
                    .TextMatrix(intRow, .ColIndex("������ĿID")) = 0 & rsTmp.Fields("Id")
                    .TextMatrix(intRow, .ColIndex("��������")) = AppObject(rsTmp.Fields("���ö���"), True)
                    .TextMatrix(intRow, .ColIndex("�ļ�Id")) = NVL(rsTmp.Fields("�ļ�Id"))
                    .TextMatrix(intRow, .ColIndex("����Id")) = 0 & rsTmp.Fields("����Id")
                    .TextMatrix(intRow, .ColIndex("����Id")) = lng����id
                    .TextMatrix(intRow, .ColIndex("��ҳId")) = lng��ҳID
                    .TextMatrix(intRow, .ColIndex("Id")) = "" & intRow
                    .TextMatrix(intRow, .ColIndex("��¼����")) = 1
                    .TextMatrix(intRow, .ColIndex("��¼״̬")) = 1
                    .TextMatrix(intRow, .ColIndex("������")) = UserInfo.����
                    .TextMatrix(intRow, .ColIndex("����ʱ��")) = datFeed
                    .TextMatrix(intRow, .ColIndex("��������")) = datFeedBack
                    .TextMatrix(intRow, .ColIndex("���ĵ�ID")) = NVL(rsTmp.Fields("���ĵ�Id"))
                End If
            End If
            End If
        End With
        rsTmp.MoveNext
    Loop
    vsfFeedback.Rows = intRow + 1
    zlCheck.Msg_OK ("�����Զ�����ɹ���")
    cmdAuto.Visible = True
    cmdOK.Enabled = True
    cmdCancel.Enabled = True
    cmdStop.Visible = False
    pbrBar.Visible = False
    LabStatus.Visible = False
    LockWindowUpdate 0
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Call zlCheck.Msg_OK("�����Զ�����ʧ�ܣ�", vbCritical)
    cmdAuto.Visible = True
    cmdOK.Enabled = True
    cmdCancel.Enabled = True
    cmdStop.Visible = False
    pbrBar.Visible = False
    LabStatus.Visible = False
    LockWindowUpdate 0
End Sub

'==============================================================================
'=���ܣ� ȡ��
'==============================================================================
Private Sub CmdCancel_Click()
    On Error GoTo ErrH
    mblnCancel = True
    Unload Me
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ��ʼ�������б�
'==============================================================================
Private Sub InitTreeView()
Dim objNode     As Node, rsTemp As ADODB.Recordset
Dim strIcon     As String, strKey     As String
Dim strSql As String
Dim blnOldData As Boolean, strTemp As String
    On Error GoTo ErrH
    
    Set tvw.ImageList = frmPubResource.ils16
        
    If Not (tvw.SelectedItem Is Nothing) Then strKey = tvw.SelectedItem.Key
    If InStr(strKey, "K") = 0 And strKey <> "R1" And strKey <> "R5" Then strKey = ""
    
    LockWindowUpdate tvw.hWnd
    
    tvw.Nodes.Clear
    DoEvents
    
    strSql = "Select 1 From ���˻����¼ A Where a.����id = [1] And a.��ҳid = [2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "����Ƿ�����ϰ�����", mlng����ID, mlng��ҳID)
    blnOldData = IIf(rsTemp.RecordCount > 0, True, False)
    Set rsTemp = gclsPackage.GetCISStruct(mlng����ID, mlng��ҳID, mlng����ID, False)
    
    Do While Not zlCheck.Connection_ChkRsState(rsTemp)
        strIcon = zlCommFun.NVL(rsTemp("ͼ��").Value)
        
        If zlCommFun.NVL(rsTemp("�ϼ�Id").Value) = "" Then
            Set objNode = tvw.Nodes.Add(, , rsTemp("Id").Value, rsTemp("����").Value, strIcon, strIcon)
            objNode.Tag = zlCommFun.NVL(rsTemp("����").Value)
        Else
            If rsTemp("�ϼ�ID").Value = "R4" Then
                strTemp = IIf(blnOldData, rsTemp("Id").Value, rsTemp("EPRID").Value)
                If mstrTreeSelect = rsTemp("Id").Value Then mstrTreeSelect = IIf(blnOldData, rsTemp("Id").Value, rsTemp("EPRID").Value)
            Else
                strTemp = rsTemp("Id").Value
            End If
            Set objNode = tvw.Nodes.Add(rsTemp("�ϼ�Id").Value, tvwChild, strTemp, rsTemp("����").Value, strIcon, strIcon)
            objNode.Tag = zlCommFun.NVL(rsTemp("����").Value)
        End If
        
        rsTemp.MoveNext
    Loop
    
    Set rsTemp = New ADODB.Recordset '�°没��
    Set rsTemp = gclsPackage.GetEmrCISStruct(mlng����ID, mlng��ҳID)
    If Not rsTemp Is Nothing Then
    If rsTemp.State = ADODB.adStateOpen Then
    If rsTemp.RecordCount > 0 Then
        rsTemp.MoveFirst
        Do Until rsTemp.EOF
            Set objNode = tvw.Nodes.Add(rsTemp!�ϼ�ID.Value, tvwChild, rsTemp!ID.Value, rsTemp!����.Value, rsTemp!ͼ��.Value, rsTemp!ͼ��.Value)
            objNode.Tag = NVL(rsTemp!����) '�ĵ�ID[|���ĵ�ID]
            rsTemp.MoveNext
        Loop
    End If
    End If
    End If
    
    If tvw.Nodes.count > 0 Then
        strKey = mstrTreeSelect
        If strKey <> "" Then
            tvw.Nodes(strKey).Selected = True
            tvw.Nodes(strKey).Expanded = True
            tvw.Nodes(strKey).Checked = True
        End If
    End If
    
    LockWindowUpdate 0
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ������Ϣ����
'==============================================================================
Private Sub InitBase()
Dim rsTemp As ADODB.Recordset
    On Error GoTo ErrH
    
    gstrSQL = "Select סԺ��,סԺ����,����,�Ա�,���� From ������Ϣ Where ����Id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Name, mlng����ID)
    If Not zlCheck.Connection_ChkRsState(rsTemp) Then
        txtסԺ��.Text = "" & rsTemp.Fields!סԺ��
        txtסԺ����.Text = "" & rsTemp.Fields!סԺ����
        txt����.Text = "" & rsTemp.Fields!����
        txt�Ա�.Text = "" & rsTemp.Fields!�Ա�
        txt����.Text = "" & rsTemp.Fields!����
    End If
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function GetPersonSet() As Boolean
    
    On Error GoTo ErrH
    GetPersonSet = False
    If Val(zlDatabase.GetPara("ʹ�ø��Ի����")) = 1 Then GetPersonSet = True

    Exit Function
ErrH:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

'==============================================================================
'=���ܣ� ��ʼ������ VsflexGrId
'==============================================================================
Private Sub InitVsflexGrid()
Dim strField As String, strFieldWidth  As String, varField As Variant, varFieldWidth  As Variant, i As Integer
Dim rsTemp As New ADODB.Recordset

    On Error GoTo ErrH
    
    vsfFeedback.FocusRect = flexFocusNone
    vsfFeedback.ExtendLastCol = True
    vsfFeedback.ExplorerBar = flexExSortShowAndMove
    vsfFeedback.AutoResize = False
    vsfFeedback.Editable = flexEDKbdMouse
    
    gstrSQL = "Select " & con_vsfField & vbCrLf & _
                "From ����������¼" & vbCrLf & _
                "Where ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Name, -1)
    Set vsfFeedback.DataSource = rsTemp
    With vsfFeedback
        .FrozenCols = 3
        .ColWidth(.ColIndex("����ʱ��")) = 1000
        .ColWidth(.ColIndex("��������")) = 1000
        If GetPersonSet Then
            'ʹ�ø��Ի����á����ѱ���ĸ�ʽ��
            strField = GetSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\��������\" & Me.Name & "\VSFlexGrId", vsfFeedback.Name & "����", "")
            strFieldWidth = GetSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\��������\" & Me.Name & "\VSFlexGrId", vsfFeedback.Name & "���", "")
            varField = Split(strField, ",")
            varFieldWidth = Split(strFieldWidth, ",")
            For i = 0 To UBound(varField)
                If varField(i) <> "" Then
                    .ColPosition(.ColIndex(varField(i))) = i
                    .ColWidth(i) = Val(varFieldWidth(i))
                End If
            Next
        End If
        .ColWidth(0) = 250
        .ColWidth(.ColIndex("ѡ��")) = 450
        .ColDataType(.ColIndex("ѡ��")) = flexDTBoolean
        
        .Cell(flexcpData, 0, .ColIndex("ѡ��")) = "[ѡ��]"
        
        .TextMatrix(0, .ColIndex("ѡ��")) = ""
        .Cell(flexcpPicture, 0, .ColIndex("ѡ��")) = frmPubResource.ils16.ListImages(4).Picture
        
        .Cell(flexcpPictureAlignment, 0, .ColIndex("ѡ��")) = flexAlignCenterCenter
        
        .MergeCol(.ColIndex("����Id")) = True
        .ColWidth(0) = 0:  .ColHidden(0) = True
        .ColWidth(.ColIndex("Id")) = 0: .ColHidden(.ColIndex("Id")) = True
        .ColWidth(.ColIndex("���Id")) = 0: .ColHidden(.ColIndex("���Id")) = True
        .ColWidth(.ColIndex("�ύId")) = 0: .ColHidden(.ColIndex("�ύId")) = True
        .ColWidth(.ColIndex("����Id")) = 0: .ColHidden(.ColIndex("����Id")) = True
        .ColWidth(.ColIndex("��ҳId")) = 0: .ColHidden(.ColIndex("��ҳId")) = True
        .ColWidth(.ColIndex("�ļ�Id")) = 0: .ColHidden(.ColIndex("�ļ�Id")) = True
        .ColWidth(.ColIndex("ҽ��Id")) = 0: .ColHidden(.ColIndex("ҽ��Id")) = True
        .ColWidth(.ColIndex("����Id")) = 0: .ColHidden(.ColIndex("����Id")) = True
        .ColWidth(.ColIndex("��¼����")) = 0: .ColHidden(.ColIndex("��¼����")) = True
        .ColWidth(.ColIndex("��¼״̬")) = 0: .ColHidden(.ColIndex("��¼״̬")) = True
        .ColWidth(.ColIndex("������ĿID")) = 0: .ColHidden(.ColIndex("������ĿID")) = True
        .ColWidth(.ColIndex("���ĵ�Id")) = 0: .ColHidden(.ColIndex("���ĵ�Id")) = False
        For i = 1 To .Cols - 1
            .ColAlignment(i) = flexAlignCenterCenter
        Next
        '���޸���
    End With
    DoEvents
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ҳ���ʼ��
'==============================================================================
Private Sub Form_Load()
    
    On Error GoTo ErrH
    zlCheck.Sys_System Me
    
    txtסԺ��.BackColor = fraDetail.BackColor
    txtסԺ����.BackColor = fraDetail.BackColor
    txt����.BackColor = fraDetail.BackColor
    txt�Ա�.BackColor = fraDetail.BackColor
    txt����.BackColor = fraDetail.BackColor
    Call InitTreeView
    Call InitVsflexGrid
    Call InitBase
    If mstrLink = "1" Then
        labShow.Caption = "���"
        labShow.ForeColor = vbBlue
    Else
        labShow.Caption = "���"
        labShow.ForeColor = vbBlack
    End If
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    SaveFlexState vsfFeedback, Me.Name
    Set zlCheck = Nothing
End Sub

Private Sub tvw_NodeCheck(ByVal Node As MSComctlLib.Node)
    On Error GoTo ErrH
    mblnChecked = Node.Checked
    NoteChildChecked Node
    NotePrentChecked Node
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
    
Private Sub NoteChildChecked(nodex As Node)
    Dim count           As Integer
    Dim ChildNode       As Node
    Dim i               As Integer
    
    On Error GoTo ErrH
    
    count = nodex.Children
    '�Խڵ���в���
    nodex.Checked = mblnChecked
    If count > 0 Then
        Set ChildNode = nodex.Child
        NoteChildChecked ChildNode
        For i = 2 To count
            Set ChildNode = ChildNode.Next
            NoteChildChecked ChildNode
        Next
    End If
    
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub NotePrentChecked(nodex As Node)
    On Error GoTo ErrH
    If mblnChecked And (Not nodex.Parent Is Nothing) Then nodex.Parent.Checked = True
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function AppObject(strApp As String, Optional blnApp As Boolean = True) As String
    Dim strReturn       As String
    
    On Error GoTo ErrH
    
    If blnApp Then
        Select Case strApp
            Case "1"
                strReturn = "סԺҽ��"
            Case "2"
                strReturn = "סԺ����"
            Case "3"
                strReturn = "������"
            Case "4"
                strReturn = "�����¼"
            Case "5"
                strReturn = "��ҳ��¼"
            Case "6"
                strReturn = "ҽ������"
            Case "7"
                strReturn = "����֤��"
            Case "8"
                strReturn = "֪���ļ�"
            Case "9"
                strReturn = "�ٴ�·��"
        End Select
    Else
        Select Case strApp
            Case "סԺҽ��"
                strReturn = "1"
            Case "סԺ����"
                strReturn = "2"
            Case "������"
                strReturn = "3"
            Case "�����¼"
                strReturn = "4"
            Case "��ҳ��¼"
                strReturn = "5"
            Case "ҽ������"
                strReturn = "6"
            Case "����֤��"
                strReturn = "7"
            Case "֪���ļ�"
                strReturn = "8"
            Case "�ٴ�·��"
                strReturn = "9"
        End Select
    End If
    AppObject = strReturn
    Exit Function
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub vsfFeedback_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    On Error GoTo ErrH
    vsfFeedback.TextMatrix(Row, Col) = ConvertString(vsfFeedback.TextMatrix(Row, Col))
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsfFeedback_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    On Error GoTo ErrH
    With vsfFeedback
        Select Case Col
            Case .ColIndex("�������")
                vsfFeedback.ComboList = "|..."
            Case .ColIndex("ѡ��")
                .ComboList = ""
            Case Else
                .ComboList = ""
                Cancel = True
        End Select
    End With
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� �����λ��¼ vsfFeedback
'==============================================================================
Private Sub vsfFeedback_AfterSort(ByVal Col As Long, Order As Integer)
    Dim lngRow      As Long
    On Error GoTo ErrH
    lngRow = vsfFeedback.FindRow(mstrSortID, -1, vsfFeedback.ColIndex("ID"), False, True)
    If lngRow > 0 Then vsfFeedback.Row = lngRow
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsfFeedback_BeforeMoveColumn(ByVal Col As Long, Position As Long)
    If Col = vsfFeedback.ColIndex("ѡ��") Then
        Position = -1
    Else
        If Position <= vsfFeedback.ColIndex("ѡ��") Then Position = Col
    End If
End Sub

'==============================================================================
'=���ܣ� ĳ�в����϶���С vsfAuditItem[ͼ��]
'==============================================================================
Private Sub vsfFeedback_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = vsfFeedback.ColIndex("ѡ��") Then Cancel = True
End Sub

'==============================================================================
'=���ܣ� ����ǰ��¼ID vsfFeedback
'==============================================================================
Private Sub vsfFeedback_BeforeSort(ByVal Col As Long, Order As Integer)
    Dim i           As Long
    Dim blnCheck    As Boolean
    On Error GoTo ErrH
    If Col = vsfFeedback.ColIndex("ѡ��") Then
        Order = -1
        With vsfFeedback
            If .Rows <= 1 Then Exit Sub
            blnCheck = Not (.TextMatrix(1, .ColIndex("ѡ��")) = "True")
            If blnCheck Then
                .Cell(flexcpPicture, 0, .ColIndex("ѡ��")) = frmPubResource.ils16.ListImages(4).Picture
            Else
                .Cell(flexcpPicture, 0, .ColIndex("ѡ��")) = frmPubResource.ils16.ListImages(25).Picture
            End If
            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("ѡ��")) = blnCheck
            Next
        End With
    End If
    mstrSortID = "" & vsfFeedback.TextMatrix(vsfFeedback.Row, vsfFeedback.ColIndex("ID"))
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsfFeedback_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    On Error GoTo ErrH
    If vsfFeedback.ColIndex("�������") = Col Then
        vsfFeedback.TextMatrix(Row, Col) = Big_Note(vsfFeedback.TextMatrix(Row, Col), vsfFeedback.ColKey(Col) & "���༭����", False)
    End If
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsfFeedback_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrH
    If vsfFeedback.ColIndex("�������") = vsfFeedback.Col Then
        '�ո�༭
        If KeyAscii = vbKeySpace Then
            'KeyAscii = 39
            KeyAscii = 0
            SendKeys "{f2}"
        End If
        '�س� ��һ���༭
        If KeyAscii = 13 Then
            KeyAscii = 0
            SendKeys "{down}"
            SendKeys "{f2}"
        End If
    End If
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsfFeedback_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    On Error GoTo ErrH
    If KeyAscii = Asc("'") Then
       KeyAscii = 0
    End If
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


