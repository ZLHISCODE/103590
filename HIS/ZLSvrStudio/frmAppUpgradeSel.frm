VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAppUpgradeSel 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��ʷ��ѡ��"
   ClientHeight    =   6780
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   11280
   Icon            =   "frmAppUpgradeSel.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   11280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdExit 
      Caption         =   "�˳�(&X)"
      Height          =   350
      Left            =   9720
      TabIndex        =   9
      Top             =   6000
      Width           =   1100
   End
   Begin VB.CommandButton cmdNotSel 
      Caption         =   "ȫ��(&R)"
      Height          =   350
      Left            =   8520
      TabIndex        =   6
      Top             =   6000
      Width           =   1100
   End
   Begin VB.CommandButton cmdSelALl 
      Caption         =   "ȫѡ(&A)"
      Height          =   350
      Left            =   7440
      TabIndex        =   4
      Top             =   6000
      Width           =   1100
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   6405
      Width           =   11280
      _ExtentX        =   19897
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmAppUpgradeSel.frx":058A
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   16325
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1111
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "10:54"
            Key             =   "STANUM"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame fraMain 
      Height          =   5985
      Left            =   0
      TabIndex        =   1
      Top             =   -60
      Width           =   11280
      Begin VSFlex8Ctl.VSFlexGrid vsReport 
         Height          =   2415
         Left            =   0
         TabIndex        =   8
         Top             =   3360
         Visible         =   0   'False
         Width           =   11220
         _cx             =   19791
         _cy             =   4260
         Appearance      =   2
         BorderStyle     =   1
         Enabled         =   0   'False
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
         ForeColor       =   0
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   0
         BackColorSel    =   14737632
         ForeColorSel    =   0
         BackColorBkg    =   16777215
         BackColorAlternate=   16777215
         GridColor       =   14737632
         GridColorFixed  =   12632256
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   3
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   250
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmAppUpgradeSel.frx":0E1C
         ScrollTrack     =   -1  'True
         ScrollBars      =   2
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
         Editable        =   2
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
         AutoSizeMouse   =   0   'False
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
      Begin VSFlex8Ctl.VSFlexGrid vsOptional 
         Height          =   2535
         Left            =   0
         TabIndex        =   7
         Top             =   1440
         Visible         =   0   'False
         Width           =   11220
         _cx             =   19791
         _cy             =   4471
         Appearance      =   2
         BorderStyle     =   1
         Enabled         =   0   'False
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
         ForeColor       =   0
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   0
         BackColorSel    =   14737632
         ForeColorSel    =   0
         BackColorBkg    =   16777215
         BackColorAlternate=   16777215
         GridColor       =   14737632
         GridColorFixed  =   12632256
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   0
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmAppUpgradeSel.frx":0EDC
         ScrollTrack     =   -1  'True
         ScrollBars      =   2
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   1
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
         Editable        =   2
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
      Begin VSFlex8Ctl.VSFlexGrid vsHis 
         Height          =   2715
         Left            =   0
         TabIndex        =   5
         Top             =   720
         Visible         =   0   'False
         Width           =   11220
         _cx             =   19791
         _cy             =   4789
         Appearance      =   2
         BorderStyle     =   1
         Enabled         =   0   'False
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
         ForeColor       =   0
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   0
         BackColorSel    =   14737632
         ForeColorSel    =   0
         BackColorBkg    =   16777215
         BackColorAlternate=   16777215
         GridColor       =   14737632
         GridColorFixed  =   12632256
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   3
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   350
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmAppUpgradeSel.frx":0F92
         ScrollTrack     =   -1  'True
         ScrollBars      =   2
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
         Editable        =   2
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
         AutoSizeMouse   =   0   'False
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
      Begin VB.Frame fraTop 
         Height          =   120
         Left            =   15
         TabIndex        =   2
         Top             =   570
         Width           =   11280
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "ȷ������Ҫ��Ǩ����ʷ���ݿռ���û�����ʷ���ݿռ�������ߣ������뼰������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   180
         TabIndex        =   3
         Top             =   225
         Width           =   8100
      End
   End
End
Attribute VB_Name = "frmAppUpgradeSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'====================================================================
'==����
'====================================================================
'����ѡ������
Public Enum AupgradeSelType
    AST_His = 0 '��ʷ��ѡ��
    AST_OptProc = 1 '��ѡ����ѡ��
    AST_Report = 2 '���뱨��ѡ��
End Enum
'��ʷ��ѡ����
Private Enum HisCols
    HC_ID = 0
    HC_ϵͳ = 1
    HC_HisDB = 2
    HC_IsCur = 3
    HC_Server = 4
    HC_PWD = 5
    HC_CurVer = 6
    HC_AimVer = 7
    HC_WarnInfo = 8
    HC_Sel = 9
End Enum
'��ѡ����ѡ����
Private Enum ProcCols
    PC_ID = 0
    PC_ϵͳ = 1
    PC_ProcExector = 2
    PC_ProcVer = 3
    PC_ProcInfo = 4
    PC_Sel = 5
End Enum
'���뱨��ѡ����
Private Enum ReportCols
    RC_ID = 0
    RC_ϵͳ = 1
    RC_RptNo = 2
    RC_RptName = 3
    RC_AllImp = 4
    RC_SourceImp = 5
End Enum
Private mastSelType As AupgradeSelType '����ѡ������
Private mrsSource As ADODB.Recordset '��ʼ��������Ҫ����Դ������¼������Ŀ��ѡ��״̬
Private mblnExecBef As Boolean '�Ƿ���ǰִ��
Private mrsSysFiles As ADODB.Recordset '��Ǩ��Ҫִ�еĽű���¼��
Private mblnOk As Boolean
'====================================================================
'==�����ӿ�
'====================================================================
Public Function ShowMe(frmParent As Object, ByVal astSelType As AupgradeSelType, Optional ByRef rsSource As ADODB.Recordset, Optional ByRef rsSysFiles As ADODB.Recordset, Optional ByVal blnExecBef As Boolean) As Boolean
'���ܣ�չʾѡ�����
'������ frmParent=������
'           astSelType=����ѡ������
'           rsSource=��ʼ��������Ҫ����Դ
'���أ�rsSource=����ѡ��״̬
'         ShowMe=�Ƿ��˳�����ʱδʹ��

    mastSelType = astSelType
    rsSource.Filter = ""
    Set mrsSource = rsSource
    mblnExecBef = blnExecBef
    Set mrsSysFiles = rsSysFiles
    Me.Show 1, frmParent
    Set rsSource = mrsSource
    Set rsSysFiles = mrsSysFiles
    ShowMe = mblnOk
End Function
'====================================================================
'==�ؼ��¼�
'====================================================================

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdNotSel_Click()
    Call SetSelBeach
End Sub

Private Sub cmdSelAll_Click()
    Call SetSelBeach(True)
End Sub

Private Sub Form_Load()
    Call ApplyOEM(stbThis)
    vsHis.Visible = mastSelType = AST_His: vsHis.Enabled = mastSelType = AST_His
    vsOptional.Visible = mastSelType = AST_OptProc: vsOptional.Enabled = mastSelType = AST_OptProc
    vsReport.Visible = mastSelType = AST_Report: vsReport.Enabled = mastSelType = AST_Report
    lblInfo.Caption = Decode(mastSelType, AST_His, "ȷ������Ҫ��Ǩ����ʷ���ݿռ���û�����ʷ���ݿռ�������ߣ������뼰��������", _
                                                                AST_OptProc, "����ϸ�Ķ�ÿ�����̵�˵���������ʵ�����������ȷ���������Щ������Ҫ�ڱ�����Ǩ�������Զ�ִ�С�", _
                                                                AST_Report, "����Ҫ�Զ�����ı����ɸ����������Ҫ����ı���Ҳ����ȡ�������Ժ��ֹ�����")
    Me.Caption = Decode(mastSelType, AST_His, "��ʷ����֤�Լ�ѡ��", AST_OptProc, "��ѡ����ѡ��", AST_Report, "����������")
    Call LoadData
End Sub

Private Sub Form_Resize()
    vsHis.Top = fraTop.Top + fraTop.Height + 60: vsOptional.Top = vsHis.Top: vsReport.Top = vsHis.Top
    vsHis.Left = 30: vsOptional.Left = 30: vsReport.Left = 30
    vsHis.Width = Me.ScaleWidth - 60: vsOptional.Width = vsHis.Width: vsReport.Width = vsHis.Width
    vsHis.Height = fraMain.Height - vsHis.Top - 30: vsOptional.Height = vsHis.Height: vsReport.Height = vsHis.Height
End Sub

Private Sub vsHis_AfterEdit(ByVal Row As Long, ByVal Col As Long)
        Select Case Col
        Case HC_PWD
            vsHis.Cell(flexcpData, Row, Col) = IIf(InStr(1, vsHis.TextMatrix(Row, Col), "*") <> 0, vsHis.Cell(flexcpData, Row, Col), vsHis.TextMatrix(Row, Col))
            vsHis.TextMatrix(Row, Col) = String(Len(vsHis.TextMatrix(Row, Col)), "*")
        Case HC_Sel
            Call RecUpdate(mrsSource, "ID=" & Val(vsHis.TextMatrix(Row, HC_ID)), "����", IIf(Val(vsHis.TextMatrix(Row, HC_Sel)) = 0, 0, 1))
        End Select
        Call RefreshColor(Row)
End Sub

Private Sub vsHis_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Col = HC_PWD Then
        If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
           If KeyAscii < Asc("a") Or KeyAscii > Asc("z") Then
              If KeyAscii < Asc("A") Or KeyAscii > Asc("Z") Then
                  If InStr(1, Chr(KeyAscii), "_") = 0 Then
                      If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyBack Then
                      Else
                          KeyAscii = 0
                      End If
                      Exit Sub
                  End If
              End If
           End If
        End If
        vsHis.Cell(flexcpData, Row, Col) = vsHis.Cell(flexcpData, Row, Col) & Chr(KeyAscii)
    End If
End Sub

Private Sub vsHis_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    '���˺�:����
    '���ñ༭������
    If Col = HC_PWD Then
        SendMessage vsHis.EditWindow, EM_SETPASSWORDCHAR, Asc("*"), 0
    End If
End Sub

Private Sub vsHis_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    '����ʱ������ǵ�ǰ��ֵ�����ֵ���ܸ���
    Cancel = Col <> HC_Sel And Col <> HC_PWD And Col <> HC_Server Or Col = HC_Sel And Trim(vsHis.TextMatrix(Row, HC_IsCur)) <> ""
End Sub

Private Sub vsHis_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim strUserName As String, strBakName As String, strPassword As String, strServer As String
    Dim cnTmp As ADODB.Connection, rsTmp As ADODB.Recordset
    Dim strFilter As String, strMaxVer As String
    Dim strDbLink As String
    
    Select Case Col
        Case HC_Server, HC_PWD
            '���Ƿ������Ƿ���Ч
            If Col = HC_Server Then
                strPassword = vsHis.Cell(flexcpData, Row, HC_PWD)
                strServer = vsHis.EditText
            Else
                If InStr(1, vsHis.EditText, "*") > 0 Then
                    strPassword = vsHis.Cell(flexcpData, Row, Col)
                Else
                    strPassword = vsHis.EditText
                End If
                strServer = vsHis.TextMatrix(Row, HC_Server)
            End If
            strPassword = Trim(strPassword)
            strServer = UCase(Trim(strServer))
            strBakName = UCase(Trim(vsHis.TextMatrix(Row, HC_HisDB)))
            strUserName = UCase(Trim(vsHis.Cell(flexcpData, Row, HC_HisDB)))
            strDbLink = UCase(Trim(vsHis.Cell(flexcpData, Row, HC_Server)))
            '�жϷ������Ƿ�ı䣬�ı䣬��ȥ����֤��־�Լ���ؽű�
            mrsSource.Filter = "ID=" & Val(vsHis.TextMatrix(Row, HC_ID))
            If strPassword = "" And mrsSource!���� & "" <> "" Then 'û���������룬ȡ�ϴ���֤������
                strPassword = mrsSource!���� & ""
            End If
            If strServer <> mrsSource!������ & "" Then
                If strServer <> "" Then '����������˾�������֤
                    mrsSource.Update Array("����", "������", "��ǰ�汾", "����", "��֤", "������", "��ֹ��Ϣ", "�����", "����ǰ����", "��ǰ��ֹ��Ϣ", "��ǰ�����"), _
                                                 Array(mrsSource!���� * -1, strServer, Null, Null, 0, 1, Null, Null, 1, Null, Null)
                    'ɾ���ű�
                    Call RecDelete(mrsSysFiles, "ϵͳ���=" & mrsSource!ϵͳ��� & " And ������='" & strBakName & "'")
                Else 'û�������������ʹ����ǰ�ķ�����
                    strServer = mrsSource!������ & ""
                End If
            End If
            
            If strPassword <> "" And strUserName <> "" And strServer <> "" Then
                Set cnTmp = gobjRegister.GetConnection(strServer, strUserName, strPassword, False, OraOLEDB, "", False)
                If cnTmp.State = adStateOpen Then
                    Set rsTmp = ReadHisUpgrade(cnTmp, strUserName, True, , strDbLink <> "")
                    Call RecUpdate(mrsSource, "������='" & strUserName & "' And ������='" & strServer & "' And ��֤=0", "��֤", 1)
                    rsTmp.Sort = ""
                    If rsTmp.EOF Then
                        Call RecUpdate(mrsSource, "������='" & strUserName & "' And ������='" & strServer & "'", "����", strPassword, "������", 0, "����ǰ����", 0, "�����", "��ʷ��ռ����ݽṹȱʧ�����޷�������")
                    Else
                        Do While Not rsTmp.EOF
                            mrsSource.Filter = "ϵͳ���=" & rsTmp!ϵͳ��� & " And ������='" & strUserName & "' And ������='" & strServer & "'"
                            Do While Not mrsSource.EOF
                                If mrsSource!��֤ = 1 Then mrsSource.Update "��֤", 2
                                mrsSource.Update Array("����", "��ǰ�汾", "��ֹ��Ϣ", "��ǰ��ֹ��Ϣ"), Array(strPassword, rsTmp!��ǰ�汾, rsTmp!��ֹ��Ϣ, rsTmp!��ǰ��ֹ��Ϣ)
                                '�ж��ܷ���Ǩ
                                If Not IsVerSion(rsTmp!��ǰ�汾 & "") Then
                                    mrsSource.Update Array("������", "�����", "����ǰ����"), Array(0, "��ʷ���ݿռ�İ汾����ʶ�����飡", 0)
                                ElseIf VerFull(rsTmp!��ǰ�汾 & "") >= VerFull(mrsSource!Ŀ��汾 & "") Then
                                    mrsSource.Update Array("������", "�����", "����ǰ����"), Array(0, "��ʷ���ݿռ�İ汾���ڱ�����ǨĿ��汾��������Ǩ��", 0)
                                Else
                                    Set mrsSysFiles = GetUpgradeFiles(mrsSysFiles, rsTmp!ϵͳ���, rsTmp!��ǰ�汾, mrsSource!�����ļ�, rsTmp!��ֹ��Ϣ, rsTmp!��ǰ��ֹ��Ϣ, mrsSource!Ŀ��汾, , strBakName)
                                    '��ȡ��ǰִ�е�Ŀ��汾
                                    If mblnExecBef Then
                                        strFilter = "������='" & strBakName & "' And FileType=" & FT_Before
                                        mrsSysFiles.Filter = strFilter: mrsSysFiles.Sort = "FullSPVer Desc": strMaxVer = ""
                                        If Not mrsSysFiles.EOF Then
                                            strMaxVer = mrsSysFiles!SPVer
                                            mrsSysFiles.Filter = strFilter & " And ���ð汾>'" & VerFull(rsTmp!��ǰ�汾 & "") & "'": mrsSysFiles.Sort = "FullSPVer"
                                            If Not mrsSysFiles.EOF Then
                                                mrsSysFiles.Filter = strFilter & " And FullSPVer<'" & mrsSysFiles!FullSPVer & "'": mrsSysFiles.Sort = "FullSPVer Desc"
                                                If Not mrsSysFiles.EOF Then
                                                    strMaxVer = mrsSysFiles!SPVer
                                                Else
                                                    strMaxVer = ""
                                                    mrsSource.Update Array("����ǰ����", "��ǰ�����"), Array(0, "û�п�ִ�е���ǰ�����ű���������ǰ��Ǩ��")
                                                End If
                                            End If
                                        Else
                                            mrsSource.Update Array("����ǰ����", "��ǰ�����"), Array(0, "û����ǰ�����ű���������ǰ��Ǩ��")
                                        End If
                                        mrsSource.Update "��ǰĿ��汾", strMaxVer
                                        'ɾ������ǰִ�нű�
                                        Call RecDelete(mrsSysFiles, "������='" & strBakName & "' And FileType<>" & FT_Before)
                                        'ɾ��������ǰĿ��汾����ǰ�����ű�
                                        Call RecDelete(mrsSysFiles, strFilter & " And FullSPVer>'" & VerFull(strMaxVer) & "'")
                                    End If
                                End If
                                mrsSource.MoveNext
                            Loop
                            rsTmp.MoveNext
                        Loop
                    End If
                    '���δ����ʷ�ռ���ע��
                    Call RecUpdate(mrsSource, "��֤=1", "������", 0, "����ǰ����", 0, "�����", "��ϵͳ����ʷ�ռ�δ��ZLBakInfo��ע�ᣡ")
                Else
                    Cancel = True
                    Exit Sub
                End If
            End If
        Case HC_Sel
            Call RecUpdate(mrsSource, "ID=" & Val(vsHis.TextMatrix(Row, HC_ID)), "����", IIf(Val(vsHis.TextMatrix(Row, HC_Sel)) = 0, 0, 1))
    End Select
    Call LoadData '���¼�������
    Call RefreshColor(Row)
End Sub

Private Sub vsOptional_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vsOptional
        If Col = PC_Sel Then
            Call RecUpdate(mrsSource, "ID=" & Val(vsOptional.TextMatrix(Row, PC_ID)), "ִ��", IIf(Val(vsOptional.TextMatrix(Row, PC_Sel)) = 0, 0, 1))
            Call RefreshColor(Row)
        End If
    End With
End Sub

Private Sub vsOptional_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> PC_Sel Then Cancel = True
End Sub

Private Sub vsReport_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vsReport
        If Col = RC_AllImp Then
            If Val(.TextMatrix(Row, Col)) <> 0 Then
                .TextMatrix(Row, RC_SourceImp) = 0  '����������Դ
            End If
        ElseIf Col = RC_SourceImp Then
            If Val(.TextMatrix(Row, Col)) <> 0 Then
                .TextMatrix(Row, RC_AllImp) = 0  '�����嵼��
            End If
        End If
        Call RecUpdate(mrsSource, "ID=" & Val(.TextMatrix(Row, RC_ID)), "��������", IIf(Val(.TextMatrix(Row, RC_SourceImp)) <> 0, 2, IIf(Val(.TextMatrix(Row, RC_AllImp)), 1, 0)))
        Call RefreshColor(Row)
    End With
End Sub

Private Sub vsReport_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    If vsReport.ColWidth(Col) < 500 Then vsReport.ColWidth(Col) = 500
End Sub

Private Sub vsReport_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = RC_AllImp Or Col = RC_SourceImp Then
        Cancel = True
    End If
End Sub

Private Sub vsReport_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Not (Col = RC_AllImp Or Col = RC_SourceImp) Then
        Cancel = True
    End If
End Sub

'====================================================================
'==����
'====================================================================
Private Sub SetSelBeach(Optional ByVal blnSel As Boolean)
'���ܣ���������ѡ��
'������blnSel=True������ѡ��False:����ȡ��
    Dim intSel As Integer, lngCol As Long, lngOtherCol As Long
    Dim i As Long
    Dim vsTmp As VSFlexGrid
    
    intSel = IIf(blnSel, 1, 0): lngCol = -1
    If mastSelType = AST_Report Then
        If vsReport.Col = RC_AllImp Or vsReport.Col = RC_SourceImp Then
            lngCol = vsReport.Col
            lngOtherCol = IIf(lngCol = RC_AllImp, RC_SourceImp, RC_AllImp)
        End If
        Set vsTmp = vsReport
        Call RecUpdate(mrsSource, "", "��������", IIf(blnSel And lngCol = RC_SourceImp, 2, IIf(blnSel And lngCol = RC_AllImp, 1, 0)))
    ElseIf mastSelType = AST_His Then
        Set vsTmp = vsHis
        lngCol = HC_Sel
        Call RecUpdate(mrsSource, "������ = 1 And ����ǰ���� = 1" & IIf(Not blnSel, " And ��ǰ<>1", ""), "����", IIf(blnSel, 1, 0))
    Else
        Set vsTmp = vsOptional
        lngCol = PC_Sel
        Call RecUpdate(mrsSource, "", "ִ��", IIf(blnSel, 1, 0))
    End If
    If lngCol = -1 Then Exit Sub
    With vsTmp
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpData, i, 0) = 1 Then
                .TextMatrix(i, lngCol) = intSel
                '������ֻ��ѡ��һ�ֵ��뷽����ѡ����һ�֣���ȡ����һ��
                If intSel = 1 And mastSelType = AST_Report Then
                    .TextMatrix(i, lngOtherCol) = 0
                End If
            End If
        Next
    End With
    Call RefreshColor
End Sub

Private Sub LoadData()
    Dim vsTmp As VSFlexGrid
    Dim strPre As String, strPrePOther As String
    
    mrsSource.Filter = ""
    If mastSelType = AST_His Then
        Set vsTmp = vsHis
        mrsSource.Sort = "ϵͳ���,��ǰ,���,ID"
    ElseIf mastSelType = AST_OptProc Then
        Set vsTmp = vsOptional
        mrsSource.Sort = "ϵͳ���, ��ʷ��, ִ����,ID"
    Else
        Set vsTmp = vsReport
        mrsSource.Sort = "ϵͳ���,���,ID"
    End If
    With vsTmp
        .Rows = .FixedRows
        .MergeCompare = flexMCTrimNoCase
        .MergeCells = flexMergeRestrictColumns
        Select Case mastSelType
            Case AST_His
                Do While Not mrsSource.EOF
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, HC_ID) = mrsSource!Id
                    .Cell(flexcpData, .Rows - 1, HC_ID) = IIf(Not mblnExecBef And mrsSource!������ = 1 Or mblnExecBef And mrsSource!����ǰ���� = 1, 1, 0)
                    If .Cell(flexcpData, .Rows - 1, HC_ID) = 1 And Val(mrsSource!��ǰ & "") = 1 Then .Cell(flexcpData, .Rows - 1, HC_ID) = -1
                    .TextMatrix(.Rows - 1, HC_ϵͳ) = mrsSource!ϵͳ���� & ""
                    .TextMatrix(.Rows - 1, HC_HisDB) = mrsSource!���� & ""
                    .Cell(flexcpData, .Rows - 1, HC_HisDB) = mrsSource!������ & ""
                    .TextMatrix(.Rows - 1, HC_IsCur) = IIf(Val(mrsSource!��ǰ & "") = 1, "��", "")
                    .TextMatrix(.Rows - 1, HC_CurVer) = mrsSource!��ǰ�汾 & ""
                    .TextMatrix(.Rows - 1, HC_AimVer) = IIf(mblnExecBef, mrsSource!��ǰĿ��汾 & "", mrsSource!Ŀ��汾 & "")
                    .Cell(flexcpData, .Rows - 1, HC_PWD) = mrsSource!���� & ""
                    .TextMatrix(.Rows - 1, HC_PWD) = String(Len(mrsSource!���� & ""), "*")
                    .TextMatrix(.Rows - 1, HC_Sel) = Val(mrsSource!���� & "")
                    .TextMatrix(.Rows - 1, HC_Server) = mrsSource!������ & ""
                    .Cell(flexcpData, .Rows - 1, HC_Server) = mrsSource!DB���� & ""
                    .TextMatrix(.Rows - 1, HC_WarnInfo) = IIf(mrsSource!����� & "" = "", mrsSource!��ǰ����� & "", mrsSource!�����)
                    .RowData(.Rows - 1) = 0
                    mrsSource.MoveNext
                Loop
            Case AST_OptProc
                Do While Not mrsSource.EOF
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, PC_ID) = mrsSource!Id
                    .Cell(flexcpData, .Rows - 1, PC_ID) = 1
                    .TextMatrix(.Rows - 1, PC_ϵͳ) = mrsSource!ϵͳ���� & ""
                    .TextMatrix(.Rows - 1, PC_ProcExector) = mrsSource!ִ���� & ""
                    .TextMatrix(.Rows - 1, PC_ProcInfo) = mrsSource!���� & vbNewLine & mrsSource!ע��
                    .TextMatrix(.Rows - 1, PC_ProcVer) = mrsSource!SPVer & ""
                    .TextMatrix(.Rows - 1, PC_Sel) = Val(mrsSource!ִ�� & "")
                     .RowData(.Rows - 1) = Val(mrsSource!ִ�� & "")
                    mrsSource.MoveNext
                Loop
                Call vsTmp.AutoSize(PC_ProcInfo)
            Case Else
                Do While Not mrsSource.EOF
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, RC_ID) = mrsSource!Id
                    .Cell(flexcpData, .Rows - 1, RC_ID) = 1
                    .TextMatrix(.Rows - 1, RC_ϵͳ) = mrsSource!ϵͳ���� & ""
                    .TextMatrix(.Rows - 1, RC_RptNo) = mrsSource!��� & ""
                    .TextMatrix(.Rows - 1, RC_RptName) = mrsSource!���� & ""
                    .TextMatrix(.Rows - 1, RC_AllImp) = IIf(Val(mrsSource!�������� & "") = 1, 1, 0)
                    .TextMatrix(.Rows - 1, RC_SourceImp) = IIf(Val(mrsSource!�������� & "") = 2, 1, 0)
                     .RowData(.Rows - 1) = Val(mrsSource!�������� & "")
                    mrsSource.MoveNext
                Loop
        End Select
        
        If mastSelType = AST_His Then
            .MergeCol(HC_ϵͳ) = True
            .MergeCol(HC_HisDB) = True
        ElseIf mastSelType = AST_OptProc Then
            .MergeCol(PC_ϵͳ) = True
            .MergeCol(PC_ProcExector) = True
        Else
            .MergeCol(RC_ϵͳ) = True
        End If
    End With
    Call RefreshColor
End Sub

Private Sub RefreshColor(Optional ByVal lngRow As Long)
    Dim i As Long
    
    If mastSelType = AST_His Then
        With vsHis
            If lngRow < .FixedRows Then
                For i = .FixedRows To .Rows - 1
                    If Val(.Cell(flexcpData, i, HC_ID)) = 0 Then
                        .Cell(flexcpForeColor, i, HC_ϵͳ, i, .Cols - 1) = &H2222B2 '��ש��
                    Else
                        .Cell(flexcpForeColor, i, HC_ϵͳ, i, .Cols - 1) = .ForeColor
                    End If
                Next
            Else
                If Val(.Cell(flexcpData, lngRow, HC_ID)) = 0 Then
                    .Cell(flexcpForeColor, lngRow, HC_ϵͳ, lngRow, .Cols - 1) = &H2222B2 '��ש��
                Else
                    .Cell(flexcpForeColor, lngRow, HC_ϵͳ, lngRow, .Cols - 1) = .ForeColor
                End If
            End If
        End With
    ElseIf mastSelType = AST_OptProc Then
    
    Else
        With vsReport
            If lngRow < .FixedRows Then
                For i = .FixedRows To .Rows - 1
                    If Val(.TextMatrix(i, RC_AllImp)) <> 0 Then
                        .Cell(flexcpForeColor, i, HC_ϵͳ, i, .Cols - 1) = .ForeColor
                    ElseIf Val(.TextMatrix(i, RC_SourceImp)) <> 0 Then
                        .Cell(flexcpForeColor, i, HC_ϵͳ, i, .Cols - 1) = vbBlue
                    Else
                        .Cell(flexcpForeColor, i, HC_ϵͳ, i, .Cols - 1) = &H808080   '��ɫ
                    End If
                Next
            Else
                If Val(.TextMatrix(lngRow, RC_AllImp)) <> 0 Then
                    .Cell(flexcpForeColor, lngRow, HC_ϵͳ, lngRow, .Cols - 1) = .ForeColor
                ElseIf Val(.TextMatrix(lngRow, RC_SourceImp)) <> 0 Then
                    .Cell(flexcpForeColor, lngRow, HC_ϵͳ, lngRow, .Cols - 1) = vbBlue
                Else
                    .Cell(flexcpForeColor, lngRow, HC_ϵͳ, lngRow, .Cols - 1) = &H808080  '��ɫ
                End If
            End If
        End With
    End If
End Sub

