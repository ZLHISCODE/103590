VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmGrantRevoke 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11295
   Icon            =   "frmGrantRevoke.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   11295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin MSComctlLib.ImageList img16 
      Left            =   2880
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGrantRevoke.frx":058A
            Key             =   "SYS"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGrantRevoke.frx":0B24
            Key             =   "MODULE"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGrantRevoke.frx":10BE
            Key             =   "REPORT"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGrantRevoke.frx":1458
            Key             =   "NODE"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGrantRevoke.frx":17F2
            Key             =   "NEW"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwReports 
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   11040
      _ExtentX        =   19473
      _ExtentY        =   1508
      View            =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      AllowReorder    =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483633
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.CommandButton cmdSingleClear 
      Caption         =   "<"
      Height          =   360
      Left            =   4080
      TabIndex        =   9
      Top             =   3720
      Width           =   255
   End
   Begin VB.CommandButton cmdSingleSelect 
      Caption         =   ">"
      Height          =   360
      Left            =   4080
      TabIndex        =   8
      Top             =   3120
      Width           =   255
   End
   Begin VB.Frame fraSelectted 
      Caption         =   "����λ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4845
      Left            =   4440
      TabIndex        =   10
      Top             =   1320
      Width           =   6735
      Begin VB.CommandButton cmdClear 
         Caption         =   "���(&L)"
         Height          =   360
         Left            =   120
         TabIndex        =   12
         Top             =   4395
         Width           =   1110
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "����(&S)"
         Height          =   360
         Left            =   4320
         TabIndex        =   13
         Top             =   4395
         Width           =   1110
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "ȡ��(&C)"
         Height          =   360
         Left            =   5520
         TabIndex        =   14
         Top             =   4395
         Width           =   1110
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfSelected 
         Height          =   3975
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   6495
         _cx             =   11456
         _cy             =   7011
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
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
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
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   1
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
   Begin VB.Frame fraSelecting 
      Caption         =   "��ѡλ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4845
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   3855
      Begin VB.TextBox txtFind 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   840
         TabIndex        =   7
         ToolTipText     =   "Enter�������ң�F3������������"
         Top             =   4440
         Width           =   2895
      End
      Begin MSComctlLib.TreeView tvwSelecting 
         Height          =   3615
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   6376
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   529
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin VB.ComboBox cboMenuGroup 
         Appearance      =   0  'Flat
         Height          =   300
         ItemData        =   "frmGrantRevoke.frx":1B8C
         Left            =   960
         List            =   "frmGrantRevoke.frx":1B8E
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   330
         Width           =   2775
      End
      Begin VB.Label lblFind 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����(&F)"
         Height          =   180
         Left            =   120
         TabIndex        =   6
         Top             =   4470
         Width           =   630
      End
      Begin VB.Label lblMenuGroup 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�˵����"
         Height          =   180
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   720
      End
   End
   Begin VB.Label lblReportName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "������"
      Height          =   180
      Left            =   1320
      TabIndex        =   16
      Top             =   120
      Width           =   540
   End
   Begin VB.Label lblPos 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   180
      Left            =   4200
      TabIndex        =   15
      Top             =   6200
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label lblReports 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1170
   End
End
Attribute VB_Name = "frmGrantRevoke"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Enum enuMode
    ����̨ = 0
    ģ��
End Enum

Private Const MSTR_NAVIGATION As String = _
    "��,,3,300|���,,3,800|ϵͳ,,3,1500|�˵�,,3,1500|����,,3,2500|PID,,0,0,n|MenuID,,0,0,n|ReportID,,0,0,n|" & _
    "����ID,,0,0,n|SysNo,,0,0,n"

Private WithEvents mobjSelected As clsVSFlexGridEx
Attribute mobjSelected.VB_VarHelpID = -1
Private mbytMode As Byte
Private mblnResult As Boolean
Private mblnGroup As Boolean
Private mcolRevoke As Collection
Private mlngPrevious As Long
Private mblnChanged As Boolean

Property Get Mode_() As enuMode
    Mode_ = mbytMode
End Property
Property Let Mode_(ByVal bytMode As enuMode)
    mbytMode = bytMode
End Property

Public Function ShowMe(ByVal frmOwner As Form, ByVal vsfSelect As VSFlexGrid) As Boolean
    Dim lngCount As Long
    
    '��ʼ��
    mblnGroup = UCase$(vsfSelect.name) = "VSFGROUP"
    If Mode_ = ģ�� And mblnGroup Then
        MsgBox "�鱨������ģ����ķ�������", vbInformation, App.Title
        Exit Function
    End If
    
    Call InitReportList(vsfSelect, lngCount)
    If lngCount <= 0 Then
        MsgBox "δѡ�񱨱�", vbInformation, App.Title
        Exit Function
    End If
    Call InitMenuGroup
    Call InitSelecting
    Call InitSelected
    
    If Mode_ = ����̨ Then
        Me.Caption = "��������-����̨"
    Else
        Me.Caption = "��������-ģ��"
    End If
    
    '��������
    Call RefreshSelected
    Call RefreshMenuGroup
    
    '����
    Me.Show vbModal, frmOwner
    ShowMe = mblnResult
    If mblnResult Then
        Unload Me
    End If
End Function

Private Sub cboMenuGroup_Click()
    If Me.Visible = False Then Exit Sub
    
    Call RefreshMenuGroup
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdClear_Click()
    Dim lngRow As Long
    
    If mblnChanged = False Then mblnChanged = True
    
    With vsfSelected
        .Redraw = False
        For lngRow = .Rows - 1 To 1 Step -1
            '��¼�ѷ����Ĳ˵������Ϣ
            If .Cell(flexcpData, lngRow, .ColIndex("��")) = Val("0-�ѷ���") Then
                Call RevokeAdd(lngRow)
            End If
            .RemoveItem lngRow
        Next
        .Redraw = True
    End With
    
    With tvwSelecting
        For lngRow = 1 To .Nodes.count
            .Nodes(lngRow).Bold = False
        Next
    End With
End Sub

Private Sub DBObjectPrivs(ByVal strFunc As String, ByVal lngReportID As Long, ByVal lngProgID As Long _
    , ByVal lngSys As Long, ByRef colSQL As Collection)
    
    Dim strObject As String, strOwner As String, strSQL As String, strTmp As String
    Dim i As Long
    Dim arrTmp() As String
    
    strObject = GetReportObjects(lngReportID)
    If strObject <> "" Then
        strObject = Mid$(strObject, 2)
        arrTmp = Split(strObject, ",")
        For i = LBound(arrTmp) To UBound(arrTmp)
            strOwner = Left$(arrTmp(i), InStr(arrTmp(i), ".") - 1)
            If InStr(";SYS;SYSTEM;ZLTOOLS;", ";" & strOwner & ";") <= 0 Then
                strTmp = Mid$(arrTmp(i), InStr(arrTmp(i), ".") + 1)
                strSQL = GetInsertProgPrivs(lngSys, lngProgID, strFunc, strTmp, strOwner, "SELECT")
                Call AddArray(colSQL, strSQL)
            End If
        Next
    End If
End Sub

Private Function ExistGrantData(ByVal blnGroup As Boolean, ByVal lngID As Long) As Boolean
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo hErr
    
    If Mode_ = ����̨ Then
        If blnGroup Then
            strSQL = _
                "Select Count(1) Rec " & vbCr & _
                "From zlMenus A, zlRPTGroups B " & vbCr & _
                "Where a.ģ�� = b.����id And a.ϵͳ Is Null And b.ID = [1]"
        Else
            strSQL = _
                "Select Count(1) Rec " & vbCr & _
                "From zlMenus A, zlReports B " & vbCr & _
                "Where a.ģ�� = b.����id And a.ϵͳ Is Null And b.ID = [1]"
        End If
        Set rsTemp = mdlPublic.OpenSQLRecord(strSQL, "��ȡָ��������������̨�ļ�¼", lngID)
    Else
        strSQL = "Select Count(1) Rec From zlRPTPuts Where ����ID = [1]"
        Set rsTemp = mdlPublic.OpenSQLRecord(strSQL, "��ȡָ����������ģ��ļ�¼", lngID)
    End If
    ExistGrantData = rsTemp!Rec > 1
    rsTemp.Close
    
    Exit Function
    
hErr:
    If mdlPublic.ErrCenter = 1 Then Resume
End Function

Private Sub cmdSave_Click()
'˵����
'  1.���������ӱ�����������̨��
'    a.�״η�����Ҫ����zlReports��zlMenus��zlPrograms��zlProgFuncs��zlProgPrivs�������ݣ�
'    b.��η���ֻ��Ҫ����zlReports��zlMenus�������ݣ�
'  2.�鱨����������̨��
'    a.�״η�����Ҫ����zlRPTGroups��zlRPTSubs��zlMenus��zlPrograms��zlProgFuncs��zlProgPrivs�������ݣ�
'    b.��η���ֻ��Ҫ����zlReports��zlMenus�������ݣ�
'  3.���������ӱ�������ģ�顣
'    ������Ҫ����zlReport��zlRPTPuts��zlProgFuncs��zlProgPrivs�������ݣ�
'  4.
'ע�⣺
'  1.ϵͳ����������������ֻ�����Զ��屨����������
'  2.�鱨����������ģ�飻

    Dim colSQL As Collection, colReportGroup As Collection
    Dim l As Long, k As Long, lngSys As Long
    Dim lngGroupID As Long, lngMenuID As Long, lngReportID As Long, lngProgID As Long, lngPid As Long
    Dim strVerifiedRID As String, strSQL As String, strTmp As String, strMenuGroup As String
    Dim rsTmp As ADODB.Recordset
    Dim arrID() As String
    Dim blnTrans As Boolean
    Dim ppbItem As PropertyBag

    With vsfSelected
        '����λ�ù������ѣ�������ֹ
        For l = 1 To .Rows - 1
            If .Cell(flexcpData, l, .ColIndex("��")) = Val("1-����") Then
                k = k + 1
            End If
        Next
        If k > 5 Then
            If MsgBox("���ѣ�" & vbCr & "��������Ŀ���࣬ȷ�ϼ���������", vbQuestion + vbYesNo + vbDefaultButton2 _
                , App.Title) = vbNo Then
                Exit Sub
            End If
        End If
        
        Screen.MousePointer = vbHourglass
        
        On Error GoTo hErr
    
        '��鱨���������顢�ӣ�Ȩ�޺�����
        Set colReportGroup = New Collection
        For l = 1 To .Rows - 1
            If .Cell(flexcpData, l, .ColIndex("��")) = Val("1-����") Then
                If mblnGroup Then
                    '�鱨��
                    lngGroupID = Val(.TextMatrix(l, .ColIndex("ReportID")))
                    If InStr(strVerifiedRID & ",", "," & lngGroupID & ",") <= 0 Then
                        strTmp = ""
                        strSQL = _
                            "Select Distinct a.Id, a.���� " & vbCr & _
                            "From zlReports A, zlRPTSubs B " & vbCr & _
                            "Where a.Id = b.����id And b.��id = [1] "
                        Set rsTmp = mdlPublic.OpenSQLRecord(strSQL, "��ȡ������ı���ID", lngGroupID)
                        Do While rsTmp.EOF = False
                            If ReportVerify(rsTmp!id, rsTmp!����) = False Then
                                Screen.MousePointer = vbDefault
                                Exit Sub
                            End If
                            strTmp = strTmp & "," & rsTmp!id & ";" & rsTmp!����
                            rsTmp.MoveNext
                        Loop
                        '��¼�鱨��ID
                        strVerifiedRID = strVerifiedRID & "," & lngGroupID
                        '��¼�鱨����ӱ���ID���ӱ�������
                        If strTmp <> "" Then
                            colReportGroup.Add Mid$(strTmp, 2), "_" & lngGroupID
                        End If
                        rsTmp.Close
                    End If
                Else
                    '����������ӱ���
                    lngReportID = Val(.TextMatrix(l, .ColIndex("ReportID")))
                    If InStr(strVerifiedRID & ",", "," & lngReportID & ",") <= 0 Then
                        If ReportVerify(lngReportID, .TextMatrix(l, .ColIndex("����"))) = False Then
                            Screen.MousePointer = vbDefault
                            Exit Sub
                        End If
                        '��¼����ID
                        strVerifiedRID = strVerifiedRID & "," & lngReportID
                    End If
                End If
            End If
        Next
        
        If Mode_ = ����̨ Then
            '1.����̨
                        
            If mblnGroup Then
                '�鱨��
                
                '����
                For l = 1 To .Rows - 1
                    If .Cell(flexcpData, l, .ColIndex("��")) = Val("1-����") Then
                        Set colSQL = New Collection
                        strMenuGroup = Trim$(.TextMatrix(l, .ColIndex("���")))
                        lngPid = Val(.TextMatrix(l, .ColIndex("PID")))
                        lngMenuID = Val(.TextMatrix(l, .ColIndex("MenuID")))
                        lngGroupID = Val(.TextMatrix(l, .ColIndex("ReportID")))
                        lngProgID = Val(.TextMatrix(l, .ColIndex("����ID")))
                        If lngProgID <= 0 Then
                            lngProgID = GetProgID(lngGroupID, True)     '����ĳ���ID�Ƿ��Ѵ��ڣ�ʵʱ��
                            If lngProgID <= 0 Then
                                lngProgID = mdlPublic.GetNewProgID()    '�����µĳ���ID
                            
                                strSQL = _
                                    "Update zlRPTSubs A Set ���� = (Select ���� From zlReports Where ID = A.����ID) " & vbCr & _
                                    "Where ��ID = " & lngGroupID
                                Call AddArray(colSQL, strSQL)
                                
                                strSQL = _
                                    "Update zlRPTGroups " & vbCr & _
                                    "Set ����ID = " & lngProgID & ", ����ʱ�� = Sysdate " & vbCr & _
                                    "Where ID = " & lngGroupID
                                Call AddArray(colSQL, strSQL)
                                
                                strSQL = _
                                    "Insert Into zlPrograms(���,����,˵��,ϵͳ,����) " & vbCr & _
                                    "Select " & lngProgID & vbCr & _
                                    "" & _
                                    ", ����, ˵��" & _
                                    ", " & IIF(glngSys <= 0, "Null", glngSys) & _
                                    ", 'zl9Report' " & vbCr & _
                                    "From zlRPTGroups Where ID = " & lngGroupID
                                Call AddArray(colSQL, strSQL)
                                
                                strSQL = _
                                    "Insert Into zlProgFuncs(ϵͳ,���,����,˵��)" & vbCr & _
                                    "Select " & IIF(glngSys <= 0, "Null", glngSys) & _
                                    ", " & lngProgID & _
                                    ", ����, ˵�� " & vbCr & _
                                    "From zlReports " & vbCr & _
                                    "Where ID In (Select ����ID From zlRPTSubs Where ��ID = " & lngGroupID & ")"
                                Call AddArray(colSQL, strSQL)
                                
                                '�鱨���������ӱ�������Դ�����ݱ�������Ȩ��
                                If CollectionFind(colReportGroup, "_" & lngGroupID) Then
                                    arrID = Split(colReportGroup("_" & lngGroupID), ",")
                                    For k = LBound(arrID) To UBound(arrID)
                                        strTmp = arrID(k)
                                        lngReportID = Val(strTmp)
                                        strTmp = Mid$(strTmp, InStr(strTmp, ";") + 1)
                                        Call DBObjectPrivs(strTmp, lngReportID, lngProgID, glngSys, colSQL)
                                    Next
                                End If
                            Else
                                strSQL = "Update zlRPTGroups Set ����ʱ�� = Sysdate Where ID = " & lngGroupID
                                Call AddArray(colSQL, strSQL)
                            End If
                        Else
                            strSQL = "Update zlRPTGroups Set ����ʱ�� = Sysdate Where ID = " & lngGroupID
                            Call AddArray(colSQL, strSQL)
                        End If
                        
                        strSQL = _
                            "Insert Into zlMenus(���,ID,�ϼ�ID,����,���,˵��,ϵͳ,ģ��,�̱���,ͼ��) " & vbCr & _
                            "Select '" & strMenuGroup & "'" & _
                            ", zlMenus_ID.Nextval" & _
                            ", " & lngPid & _
                            ", ����, Null, ˵��" & _
                            ", " & IIF(glngSys <= 0, "Null", glngSys) & _
                            ", " & lngProgID & _
                            ", ����, 105 " & vbCr & _
                            "From zlRPTGroups Where ID = " & lngGroupID
                        Call AddArray(colSQL, strSQL)
                        
                        'ִ�з�����DML
                        gcnOracle.BeginTrans: blnTrans = True
                        For k = 1 To colSQL.count
                            'Debug.Print colSQL(k) & ";"
                            gcnOracle.Execute colSQL(k)
                        Next
                        gcnOracle.CommitTrans: blnTrans = False
                    End If
                Next
            Else
                '���������ӱ���
                
                '����
                For l = 1 To .Rows - 1
                    If .Cell(flexcpData, l, .ColIndex("��")) = Val("1-��") Then
                        Set colSQL = New Collection
                        strMenuGroup = Trim$(.TextMatrix(l, .ColIndex("���")))
                        lngPid = Val(.TextMatrix(l, .ColIndex("PID")))
                        lngMenuID = Val(.TextMatrix(l, .ColIndex("MenuID")))
                        lngReportID = Val(.TextMatrix(l, .ColIndex("ReportID")))
                        lngProgID = Val(.TextMatrix(l, .ColIndex("����ID")))
                        If lngProgID <= 0 Then
                            lngProgID = GetProgID(lngReportID)          '����ĳ���ID�Ƿ��Ѵ��ڣ�ʵʱ��
                            If lngProgID <= 0 Then
                                lngProgID = mdlPublic.GetNewProgID()    '�����µĳ���ID
                            
                                strSQL = _
                                    "Update zlReports " & vbCr & _
                                    "Set ���� = '����', ����ID = " & lngProgID & ", ����ʱ�� = Sysdate " & vbCr & _
                                    "Where ID = " & lngReportID
                                Call AddArray(colSQL, strSQL)
                                
                                strSQL = _
                                    "Insert Into zlPrograms(���,����,˵��,ϵͳ,����) " & vbCr & _
                                    "Select " & lngProgID & _
                                    ", ����, ˵��" & _
                                    ", " & IIF(glngSys <= 0, "Null", glngSys) & _
                                    ", 'zl9Report' " & vbCr & _
                                    "From zlReports Where ID = " & lngReportID
                                Call AddArray(colSQL, strSQL)
                                
                                strSQL = _
                                    "Insert Into zlProgFuncs(ϵͳ,���,����) " & vbCr & _
                                    "Values (" & IIF(glngSys <= 0, "Null", glngSys) & _
                                    ", " & lngProgID & _
                                    ", '����')"
                                Call AddArray(colSQL, strSQL)
                                
                                '�鱨���������ӱ�������Դ�����ݱ�������Ȩ��
                                Call DBObjectPrivs("����", lngReportID, lngProgID, glngSys, colSQL)
                            Else
                                strSQL = "Update zlReports Set ����ʱ�� = Sysdate Where ID = " & lngReportID
                                Call AddArray(colSQL, strSQL)
                            End If
                        Else
                            strSQL = "Update zlReports Set ����ʱ�� = Sysdate Where ID = " & lngReportID
                            Call AddArray(colSQL, strSQL)
                        End If
                        
                        strSQL = _
                            "Insert Into zlMenus(���,ID,�ϼ�ID,����,���,˵��,ϵͳ,ģ��,�̱���,ͼ��) " & vbCr & _
                            "Select '" & strMenuGroup & "'" & _
                            ", zlMenus_ID.Nextval" & _
                            ", " & lngPid & _
                            ", ����, Null, ˵��" & _
                            ", " & IIF(glngSys <= 0, "Null", glngSys) & _
                            ", " & lngProgID & _
                            ", ����, 105 " & vbCr & _
                            "From zlReports Where ID = " & lngReportID
                        Call AddArray(colSQL, strSQL)
                        
                        'ִ�з�����DML
                        gcnOracle.BeginTrans: blnTrans = True
                        For k = 1 To colSQL.count
                            'Debug.Print colSQL(k) & ";"
                            gcnOracle.Execute colSQL(k)
                        Next
                        gcnOracle.CommitTrans: blnTrans = False
                    End If
                Next
            End If
            
            '��������
            Set colSQL = New Collection
            If Not mcolRevoke Is Nothing Then
                For l = 1 To mcolRevoke.count
                    Set ppbItem = mcolRevoke(l)
                    lngProgID = ppbItem.ReadProperty("ProgID")
                    lngMenuID = ppbItem.ReadProperty("MenuID")
                    lngReportID = ppbItem.ReadProperty("ReportID")
                    
                    '�жϱ����Ƿ���ڷ������ݣ�ʵʱ��
                    If ExistGrantData(mblnGroup, lngReportID) Then
                        If mblnGroup Then
                            '�鱨�����ڷ���������̨������
                            strSQL = "Update zlRPTGroups Set ����ʱ�� = Sysdate Where ID = " & lngReportID & " And ����ʱ�� <> Sysdate "
                            Call AddArray(colSQL, strSQL)
                        Else
                            '�������ڷ���������̨������
                            strSQL = _
                                "Update zlReports Set ����ʱ�� = Sysdate " & vbCr & _
                                "Where ID = " & lngReportID & " And ����ʱ�� <> Sysdate "
                            Call AddArray(colSQL, strSQL)
                        End If
                        strSQL = "Delete From zlMenus Where ID = " & lngMenuID & " And Nvl(ϵͳ, 0) = " & glngSys
                        Call AddArray(colSQL, strSQL)
                    Else
                        If mblnGroup Then
                            '�鱨��δ���ڷ���������̨������
                            strSQL = _
                                "Update zlRPTGroups Set ����ID = Null, ����ʱ�� = Null, �Ƿ�ͣ�� = Null " & vbCr & _
                                "Where ID = " & lngReportID
                            Call AddArray(colSQL, strSQL)
                            
                            strSQL = _
                                "Update zlRPTSubs Set ���� = Null " & vbCr & _
                                "Where ��ID = " & lngReportID
                            Call AddArray(colSQL, strSQL)
                        Else
                            '����δ���ڷ���������̨������
                            strSQL = _
                                "Update zlReports " & vbCr & _
                                "Set ���� = Null, ����ID = Null, �Ƿ�ͣ�� = Null" & vbCr & _
                                "  , ����ʱ�� = Case When Exists(Select 1 From zlRPTPuts Where ����Id = " & lngReportID & ") Then " & vbCr & _
                                "                      Sysdate" & vbCr & _
                                "              Else Null End" & vbCr & _
                                "Where ID = " & lngReportID
                            Call AddArray(colSQL, strSQL)
                        End If
                        strSQL = "Delete From zlMenus Where ģ�� = " & lngProgID & " And Nvl(ϵͳ,0) = " & glngSys
                        Call AddArray(colSQL, strSQL)
                        
                        strSQL = "Delete From zlProgPrivs Where ��� = " & lngProgID & " And Nvl(ϵͳ,0) = " & glngSys
                        Call AddArray(colSQL, strSQL)
                        
                        strSQL = "Delete From zlProgFuncs Where ��� = " & lngProgID & " And Nvl(ϵͳ,0) = " & glngSys
                        Call AddArray(colSQL, strSQL)
                        
                        strSQL = "Delete From zlPrograms Where ��� = " & lngProgID & " And Nvl(ϵͳ,0) = " & glngSys
                        Call AddArray(colSQL, strSQL)
                        
                        'ͬʱ���ոñ�����ݣ������ߵĽ�ɫ��Ȩ���������ݣ�
                        strSQL = "Delete From zlRoleGrant Where ��� = " & lngProgID & " And Nvl(ϵͳ,0) = " & glngSys
                        Call AddArray(colSQL, strSQL)
                    End If
                    
                    'ִ�г���������DML
                    gcnOracle.BeginTrans: blnTrans = True
                    For k = 1 To colSQL.count
                        'Debug.Print colSQL(k) & ";"
                        gcnOracle.Execute colSQL(k)
                    Next
                    gcnOracle.CommitTrans: blnTrans = False
                Next
            End If
        Else
            '2.ģ��
            
            '����
            Set colSQL = New Collection
            For l = 1 To .Rows - 1
                If .Cell(flexcpData, l, .ColIndex("��")) = Val("1-��") Then
                    lngReportID = Val(.TextMatrix(l, .ColIndex("ReportID")))
                    lngProgID = Val(.TextMatrix(l, .ColIndex("PID")))
                    lngSys = Val(.TextMatrix(l, .ColIndex("SysNo")))
                
                    strSQL = "Update zlReports Set ����ʱ�� = Sysdate Where ID = " & lngReportID
                    Call AddArray(colSQL, strSQL)
                    
                    strSQL = _
                        "Insert Into zlRPTPuts(����ID, ϵͳ, ����ID, ����) " & vbCr & _
                        "Select " & lngReportID & _
                        ", " & IIF(lngSys <= 0, "Null", lngSys) & _
                        ", " & lngProgID & _
                        ", ���� " & vbCr & _
                        "From zlReports Where ID = " & lngReportID
                    Call AddArray(colSQL, strSQL)
                
                    strSQL = _
                        "Insert Into zlProgFuncs(ϵͳ, ���, ����, ˵��) " & vbCrLf & _
                        "Select " & IIF(lngSys <= 0, "Null", lngSys) & _
                        ", " & lngProgID & _
                        ", ����, ˵�� " & _
                        "From zlReports Where ID = " & lngReportID
                    Call AddArray(colSQL, strSQL)
                    
                    '���������ӱ�������Դ�����ݱ�������Ȩ��
                    Call DBObjectPrivs(Trim$(.TextMatrix(l, .ColIndex("����"))), lngReportID, lngProgID, lngSys, colSQL)
                End If
            Next
            
            'ִ�з�����DML
            If colSQL.count > 0 Then
                gcnOracle.BeginTrans: blnTrans = True
                For k = 1 To colSQL.count
                    'Debug.Print colSQL(k) & ";"
                    gcnOracle.Execute colSQL(k)
                Next
                gcnOracle.CommitTrans: blnTrans = False
            End If
            
            '��������
            Set colSQL = New Collection
            If Not mcolRevoke Is Nothing Then
                For l = 1 To mcolRevoke.count
                    Set ppbItem = mcolRevoke(l)
                    lngProgID = ppbItem.ReadProperty("ProgID")
                    lngReportID = ppbItem.ReadProperty("ReportID")
                    lngSys = ppbItem.ReadProperty("SysNo")
                    
                    strSQL = _
                        "Delete From zlRPTPuts " & vbCr & _
                        "Where ����ID = " & lngReportID & " And ϵͳ = " & lngSys & " And ����ID = " & lngProgID
                    Call AddArray(colSQL, strSQL)
                    
                    strSQL = _
                        "Delete From zlProgPrivs " & vbCr & _
                        "Where ϵͳ = " & lngSys & " And ��� = " & lngProgID & vbCr & _
                        "  And ���� = (Select ���� From zlReports Where ID = " & lngReportID & ")"
                    Call AddArray(colSQL, strSQL)
                    
                    strSQL = _
                        "Delete From zlProgFuncs Where ϵͳ = " & lngSys & " And ���=" & lngProgID & vbCr & _
                        "  And ���� = (Select ���� From zlReports Where ID = " & lngReportID & ")"
                    Call AddArray(colSQL, strSQL)
                    
                    strSQL = _
                        "Delete From zlRoleGrant Where ϵͳ = " & lngSys & " And ���=" & lngProgID & vbCr & _
                        "  And ���� = (Select ���� From zlReports Where ID = " & lngReportID & ")"
                    Call AddArray(colSQL, strSQL)
                                        
                    '�жϱ����Ƿ���ڷ������ݣ�ʵʱ��
                    If ExistGrantData(False, lngReportID) Then
                        '���ڣ��������ݴ���
                    Else
                        '������
                        strSQL = _
                            "Update zlReports Set ����ʱ�� = NULL, �Ƿ�ͣ�� = NULL " & vbCr & _
                            "Where ����ID Is Null And ID = " & lngReportID
                        Call AddArray(colSQL, strSQL)
                    End If
                    
                    'ִ�г���������DML
                    gcnOracle.BeginTrans: blnTrans = True
                    For k = 1 To colSQL.count
                        'Debug.Print colSQL(k) & ";"
                        gcnOracle.Execute colSQL(k)
                    Next
                    gcnOracle.CommitTrans: blnTrans = False
                Next
            End If
        End If
    End With
    
    mblnResult = True
    Me.Hide
    Screen.MousePointer = vbDefault
    Exit Sub

hErr:
    Screen.MousePointer = vbDefault
    If blnTrans Then
        gcnOracle.RollbackTrans
    End If
    Call mdlPublic.ErrCenter
End Sub

Private Sub RevokeAdd(ByVal lngRow As Long)
    Dim strKey As String
    Dim ppbItem As PropertyBag
    
    Set ppbItem = New PropertyBag
    With vsfSelected
        strKey = Trim$(.TextMatrix(lngRow, .ColIndex("PID"))) & "_" & Trim$(.TextMatrix(lngRow, .ColIndex("ReportID")))
        Call ppbItem.WriteProperty("PID", Val(.TextMatrix(lngRow, .ColIndex("PID"))))
        Call ppbItem.WriteProperty("MenuID", Val(.TextMatrix(lngRow, .ColIndex("MenuID"))))
        Call ppbItem.WriteProperty("ReportID", Val(.TextMatrix(lngRow, .ColIndex("ReportID"))))
        Call ppbItem.WriteProperty("ProgID", Val(.TextMatrix(lngRow, .ColIndex("����ID"))))
        Call ppbItem.WriteProperty("SysNo", Val(.TextMatrix(lngRow, .ColIndex("SysNo"))))
        Call CollectionAdd(mcolRevoke, strKey, ppbItem)
    End With
End Sub

Private Sub cmdSingleClear_Click()
    Dim lngRow As Long, lngID As Long
    
    If vsfSelected.Rows <= 1 Then Exit Sub
    If vsfSelected.SelectedRows <= 0 Then Exit Sub
    If vsfSelected.Row <= 0 Then Exit Sub
    
    If mblnChanged = False Then mblnChanged = True
    
    With vsfSelected
        lngRow = .Row
        lngID = Val(.TextMatrix(lngRow, .ColIndex("PID")))
        
        '��¼�ѷ����Ĳ˵������Ϣ
        If .Cell(flexcpData, lngRow, .ColIndex("��")) = Val("0-�ѷ���") Then
            Call RevokeAdd(lngRow)
        End If
        
        'ɾ��
        .Redraw = False
        .RemoveItem .SelectedRow(0)
        If lngRow < .Rows - 1 Then
            .Row = lngRow
        Else
            .Row = .Rows - 1
        End If
        .Redraw = True
        .SetFocus
    End With
    
    '���¿�ѡλ��
    With tvwSelecting
        For lngRow = 1 To .Nodes.count
            If Val(.Nodes(lngRow).Tag) = lngID Then
                Call CheckGranted(lngID, .Nodes(lngRow))
                Exit For
            End If
        Next
    End With
End Sub

Private Sub cmdSingleSelect_Click()
'ע����ĩ�ҽ��ѡ�з������ý����ӽ��������⡣��ˣ���֧��ѡ�к��������
    Dim objItem As ListItem
    Dim i As Long, lngRow As Long
    Dim blnFound As Boolean
    Dim strKey As String, strProgID As String, strText As String
    
    If cmdSingleSelect.Enabled = False Then Exit Sub
    
    If mblnChanged = False Then mblnChanged = True
    
    With vsfSelected
        For Each objItem In lvwReports.ListItems
            blnFound = False
            For i = 1 To .Rows - 1
                '�ų����桰����λ�á��б��λ�ã�����ID �� �����Ĳ˵�λ��ID��
                If Val(.TextMatrix(i, .ColIndex("ReportID"))) = Val(objItem.Tag) _
                    And Val(.TextMatrix(i, .ColIndex("PID"))) = Val(tvwSelecting.SelectedItem.Tag) Then
                    blnFound = True
                    Exit For
                End If
            Next
            
            If blnFound = False Then
                '����
                .Redraw = False
                .Rows = .Rows + 1
                lngRow = .Rows - 1
                
                '��ѡ��ʶ
                tvwSelecting.SelectedItem.Bold = True
                
                '�ų��ѷ�������Ŀ�����ܱ������У�
                strKey = CStr(Val(tvwSelecting.SelectedItem.Tag)) & "_" & Val(objItem.Tag)    '�˵�ID_����ID/��ID
                If CollectionFind(mcolRevoke, strKey) Then
                    .TextMatrix(lngRow, .ColIndex("����ID")) = mcolRevoke(strKey).ReadProperty("ProgID")
                    .TextMatrix(lngRow, .ColIndex("PID")) = mcolRevoke(strKey).ReadProperty("PID")
                    .TextMatrix(lngRow, .ColIndex("SysNo")) = mcolRevoke(strKey).ReadProperty("SysNo")
                    Call CollectionDelete(mcolRevoke, strKey)
                Else
                    If Mode_ = ����̨ Then
                        strProgID = Trim$(Mid$(objItem.Tag, InStr(objItem.Tag, "_") + 1))
                    Else
                        strProgID = Val(tvwSelecting.SelectedItem.Tag)
                    End If
                    .TextMatrix(lngRow, .ColIndex("����ID")) = strProgID
                    .TextMatrix(lngRow, .ColIndex("PID")) = Val(tvwSelecting.SelectedItem.Tag)
                    .TextMatrix(lngRow, .ColIndex("SysNo")) = Abs(Val(GetRootNode(tvwSelecting.SelectedItem).Tag))
                    .Cell(flexcpData, lngRow, .ColIndex("��")) = Val("1-����")
                    .Cell(flexcpPicture, lngRow, .ColIndex("��")) = img16.ListImages("NEW").Picture
                    .Cell(flexcpPictureAlignment, lngRow, .ColIndex("��")) = flexPicAlignCenterCenter
                End If
                strText = GetRootNode(tvwSelecting.SelectedItem).Text
                If InStr(strText, "]") > 0 Then
                    .TextMatrix(lngRow, .ColIndex("ϵͳ")) = Mid$(strText, InStr(strText, "]") + 1)
                Else
                    .TextMatrix(lngRow, .ColIndex("ϵͳ")) = strText
                End If
                .TextMatrix(lngRow, .ColIndex("���")) = cboMenuGroup.Text
                .TextMatrix(lngRow, .ColIndex("�˵�")) = tvwSelecting.SelectedItem.Text
                .TextMatrix(lngRow, .ColIndex("����")) = objItem.Text
                .TextMatrix(lngRow, .ColIndex("MenuID")) = ""
                .TextMatrix(lngRow, .ColIndex("ReportID")) = CStr(Val(objItem.Tag))
                .Row = lngRow
                .TopRow = .BottomRow
                .Redraw = True
            End If
        Next
    End With
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        Call FindItem(False)
    End If
End Sub

Private Sub Form_Load()
    mblnResult = False
    mblnChanged = False
    Me.KeyPreview = True
End Sub

Private Sub RefreshSelected()
    Dim strSQL As String, strReportID As String
    Dim rsTemp As ADODB.Recordset

    On Error GoTo hErr
    
    strReportID = GetReportIDs()
    
    If Mode_ = ����̨ Then
        '�����ѷ�����λ��
        If mblnGroup Then
            strSQL = _
                "Select a2.���, a2.ID PID, a2.���� �˵�, b.���� ϵͳ, a1.ID MenuID, d.���� ����, d.ID ReportID" & vbCr & _
                "  , d.����id " & vbCr & _
                "From zlMenus A1, zlMenus A2, zlSystems B, zlPrograms C, zlRPTGroups D, Table(f_Num2list([1], ',')) E " & vbCr & _
                "Where a1.ģ�� = c.��� And a1.ģ�� = d.����id And a2.ϵͳ = b.��� " & vbCr & _
                "  And a1.�ϼ�id = a2.ID(+) and d.ID = e.Column_Value " & vbCr & _
                "  And Upper(c.����) = 'ZL9REPORT' " & vbCr & _
                "  And a1.ϵͳ is Null And c.ϵͳ is Null And d.ϵͳ is Null "
        Else
            strSQL = _
                "Select a2.���, a2.ID PID, a2.���� �˵�, b.���� ϵͳ, a1.ID MenuID, d.���� ����, d.ID ReportID" & vbCr & _
                "  , d.����id " & vbCr & _
                "From zlMenus A1, zlMenus A2, zlSystems B, zlPrograms C, zlReports D, Table(f_Num2list([1], ',')) E " & vbCr & _
                "Where a1.ģ�� = c.��� And a1.ģ�� = d.����id And a2.ϵͳ = b.��� " & vbCr & _
                "  And a1.�ϼ�id = a2.ID(+) And d.ID = e.Column_Value " & vbCr & _
                "  And Upper(c.����) = 'ZL9REPORT' " & vbCr & _
                "  And a1.ϵͳ is Null And c.ϵͳ is Null And d.ϵͳ is Null "
        End If
        Set rsTemp = mdlPublic.OpenSQLRecord(strSQL, "��ȡ�����ѷ����Ĳ˵�λ��", strReportID)
    Else
        strSQL = _
            "Select a.���, a.ģ�� Pid, a.���� �˵�, b.���� ϵͳ, a.Id Menuid, e.���� ����, e.Id ReportID" & vbCr & _
            "  , a.ģ�� ����id, a.ϵͳ SysNo " & vbCr & _
            "From zlMenus A, zlSystems B, zlPrograms C, zlRPTPuts D, zlReports E " & vbCr & _
            "  , (Select ϵͳ * 100 ϵͳ, ��� From zlRegFunc Group By ϵͳ, ���) F " & vbCr & _
            "  , Table(Cast(f_Num2list([1], ',') As t_Numlist)) G " & vbCr & _
            "Where a.ϵͳ = b.��� " & vbCr & _
            "  And a.ģ�� = c.��� " & vbCr & _
            "  And c.ϵͳ = d.ϵͳ And c.��� = d.����id " & vbCr & _
            "  And c.ϵͳ = f.ϵͳ And c.��� = f.��� " & vbCr & _
            "  And d.����id = e.Id " & vbCr & _
            "  And e.Id = g.Column_Value " & vbCr & _
            "  And Upper(c.����) <> 'ZL9REPORT' "
        Set rsTemp = mdlPublic.OpenSQLRecord(strSQL, "��ȡ�����ѷ�����ģ��λ��", strReportID)
    End If
    
    mobjSelected.Recordset = rsTemp
    Call mobjSelected.Repaint(RT_Rows)
    rsTemp.Close
    
    Exit Sub
    
hErr:
    If mdlPublic.ErrCenter = 1 Then Resume
End Sub

Private Sub RefreshMenuGroup()
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim objNode As Node
    Dim bytStep As Byte

    On Error GoTo hErr
    
    bytStep = Val("0-����")
    
    If Mode_ = ����̨ Then
        strSQL = _
            "Select * From (" & vbCr & _
            "  Select ��� As Scol, 0 As Flag, -��� As ID, -null As �ϼ�id, '[' || ��� || ']' || ���� As ���� " & vbCr & _
            "  From zlSystems A " & vbCr & _
            "  Where Exists(Select 1 From zlMenus B Where b.ϵͳ = a.��� And b.��� = [1]) " & vbCr & _
            "  Union All " & vbCr & _
            "  Select 99999 As Scol, Level As Flag, ID, Nvl(�ϼ�id, -ϵͳ) As �ϼ�id, ���� " & vbCr & _
            "  From zlMenus A " & vbCr & _
            "  Where ��� = [1] And ģ�� Is Null " & vbCr & _
            "    And Exists(Select 1 From zlSystems B Where b.��� = a.ϵͳ) " & vbCr & _
            "  Start With �ϼ�id Is Null And ��� = [1] " & vbCr & _
            "  Connect By Prior ID = �ϼ�id And ��� = [1] " & vbCr & _
            ") Order By Scol, Flag, ID"
        Set rsTemp = mdlPublic.OpenSQLRecord(strSQL, "��ȡ����������̨�Ĳ˵���", cboMenuGroup.Text)
    Else
        strSQL = _
            "Select * From (" & vbCr & _
            "  Select 0 Flag, -null �ϼ�id, -��� ID, '[' || ��� || ']' || ���� ���� " & vbCr & _
            "  From zlSystems A " & vbCr & _
            "  Where Exists(Select 1 From zlMenus B Where b.ϵͳ = a.��� And b.��� = [1]) " & vbCr & _
            "  Union All " & vbCr & _
            "  Select Distinct 1 Flag, -b.ϵͳ �ϼ�id, b.��� ID, b.���� " & vbCr & _
            "  From zlMenus A, zlPrograms B " & vbCr & _
            "    , (Select ϵͳ * 100 ϵͳ, ��� From zlRegFunc Group By ϵͳ, ���) C " & vbCr & _
            "  Where a.ϵͳ = b.ϵͳ And a.ģ�� = b.��� " & vbCr & _
            "    And b.ϵͳ = c.ϵͳ And b.��� = c.��� " & vbCr & _
            "    And Upper(b.����) <> 'ZL9REPORT' And a.��� = [1] " & vbCr & _
            "    And Exists(Select 1 From zlSystems X Where x.��� = a.ϵͳ) " & vbCr & _
            ") Order By Flag, Abs(�ϼ�ID), Abs(id) "
        Set rsTemp = mdlPublic.OpenSQLRecord(strSQL, "��ȡ������ģ��Ĳ˵���", cboMenuGroup.Text)
    End If

    bytStep = Val("1-����")
    
    With tvwSelecting
        .Nodes.Clear
        Do While rsTemp.EOF = False
            If Nvl(rsTemp!Flag, 0) = 0 Then
                Set objNode = .Nodes.Add(, , "_" & rsTemp!id, rsTemp!����, "SYS")
            Else
                If Mode_ = ����̨ Then
                    Set objNode = .Nodes.Add("_" & rsTemp!�ϼ�ID, tvwChild _
                        , "_" & rsTemp!id, rsTemp!����, "NODE")
                Else
                    Set objNode = .Nodes.Add("_" & rsTemp!�ϼ�ID, tvwChild _
                        , "_" & rsTemp!id & "_" & Nvl(rsTemp!�ϼ�ID), rsTemp!����, "MODULE")
                End If
            End If
            objNode.Tag = rsTemp!id
            objNode.Expanded = True
            Call CheckGranted(rsTemp!id, objNode)
            rsTemp.MoveNext
        Loop
        rsTemp.Close
    End With
    
    Exit Sub
    
hErr:
    Select Case bytStep
    Case 1
        MsgBox "���ѣ�" & vbCr & "�˵��ṹ���ܴ����쳣�������̨���ݣ�", vbInformation, App.Title
    Case Else
        If mdlPublic.ErrCenter() = 1 Then Resume
    End Select
End Sub

Private Sub CheckGranted(ByVal lngID As Long, ByRef objNode As Node)
    Dim l As Long
    
    With vsfSelected
        For l = 1 To .Rows - 1
            If lngID = Val(.TextMatrix(l, .ColIndex("PID"))) Then
                objNode.Bold = True
                Exit Sub
            End If
        Next
        objNode.Bold = False
    End With
End Sub

Private Sub InitSelected()
    Set mobjSelected = New clsVSFlexGridEx
    With mobjSelected
        .AppTemplate EM_Display, vsfSelected, MSTR_NAVIGATION, "", True
        .Init False
        .Binding.ExplorerBar = flexExNone
    End With
    
    If lblReportName.Visible Then
        With vsfSelected
            .ColHidden(.ColIndex("����")) = True
            .ColWidth(.ColIndex("�˵�")) = .ColWidth(.ColIndex("�˵�")) + .ColWidth(.ColIndex("����"))
        End With
    End If
End Sub

Private Sub InitSelecting()
    With tvwSelecting
        .Appearance = ccFlat
        .BorderStyle = ccFixedSingle
        .FullRowSelect = True
        .Indentation = 300
        .LineStyle = tvwRootLines
        Set .ImageList = img16
    End With
End Sub

Private Sub InitMenuGroup()
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset

    On Error GoTo hErr
    
    With cboMenuGroup
        .Appearance = 0
        .Clear
    End With
    
    strSQL = "Select Distinct ��� From zlMenus Where ��� Is Not Null "
    Set rsTemp = mdlPublic.OpenSQLRecord(strSQL, "��ȡ�˵����")
    Do While rsTemp.EOF = False
        cboMenuGroup.AddItem rsTemp!���
        If rsTemp!��� = "ȱʡ" Then
            cboMenuGroup.ListIndex = cboMenuGroup.NewIndex
        End If
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    Exit Sub
    
hErr:
    If mdlPublic.ErrCenter() = 1 Then Resume
End Sub

Private Sub InitReportList(ByRef vsfSelect As VSFlexGrid, ByRef lngCount As Long)
    Dim lngRow As Long, lngSelect As Long
    Dim objItem As ListItem

    lngCount = 0
    With lvwReports
        Set .Icons = img16
        Set .SmallIcons = img16
        .AllowColumnReorder = True
        .Appearance = ccFlat
        .BackColor = &H8000000F
        .View = lvwList
    End With
    
    With lvwReports.ColumnHeaders
        .Clear
        .Add , "ID"
        .Add , "Name"
    End With
    
    lngSelect = 0
    lvwReports.ListItems.Clear
    For lngRow = 1 To vsfSelect.Rows - 1
        If vsfSelect.SelectedRow(lngSelect) = lngRow Then
            If LCase$(vsfSelect.name) = "vsfgroup" Then
                Set objItem = lvwReports.ListItems.Add( _
                    , "_" & vsfSelect.TextMatrix(lngRow, vsfSelect.ColIndex("ID")) _
                    , vsfSelect.TextMatrix(lngRow, vsfSelect.ColIndex("����")) _
                    , "REPORT", "REPORT")
            Else
                Set objItem = lvwReports.ListItems.Add( _
                    , "_" & vsfSelect.TextMatrix(lngRow, vsfSelect.ColIndex("ID")) _
                    , vsfSelect.TextMatrix(lngRow, vsfSelect.ColIndex("����")) _
                    , "REPORT", "REPORT")
            End If
            objItem.Tag = vsfSelect.TextMatrix(lngRow, vsfSelect.ColIndex("ID")) & _
                          "_" & _
                          vsfSelect.TextMatrix(lngRow, vsfSelect.ColIndex("����ID"))
            lngSelect = lngSelect + 1
        End If
    Next
    lngCount = lngSelect
    
    If lngCount = 1 Then
        lvwReports.Visible = False
        lblReportName.Visible = True
        lblReportName.Caption = lvwReports.ListItems(1).Text
    Else
        lvwReports.Visible = True
        lblReportName.Visible = False
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If mblnChanged And mblnResult = False Then
        If MsgBox("�Ƿ�ȷ������������", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
            Cancel = 1
            Exit Sub
        End If
        Unload Me
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    If lvwReports.Visible = False Then
        With fraSelecting
            .Top = lblReportName.Top + lblReportName.Height + 150
            .Height = lblPos.Top - .Top - 30
        End With
        
        tvwSelecting.Height = fraSelecting.Height - tvwSelecting.Top - txtFind.Height - 210
        txtFind.Top = fraSelecting.Height - txtFind.Height - 120
        lblFind.Top = txtFind.Top + 30
                
        With fraSelectted
            .Top = fraSelecting.Top
            .Height = fraSelecting.Height
        End With
        
        vsfSelected.Height = fraSelectted.Height - vsfSelected.Top - txtFind.Height - 210
        cmdClear.Top = vsfSelected.Top + vsfSelected.Height + 45
        cmdSave.Top = cmdClear.Top
        cmdCancel.Top = cmdClear.Top
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mcolRevoke = Nothing
    Set mobjSelected = Nothing
End Sub

Private Sub lvwReports_BeforeLabelEdit(Cancel As Integer)
    Cancel = 1      '��ֹ�޸�
End Sub

Private Function GetReportIDs()
    Dim i As Integer
    Dim strResult As String
    
    For i = 1 To lvwReports.ListItems.count
        strResult = strResult & "," & Mid$(lvwReports.ListItems(i).Key, 2)
    Next
    If strResult <> "" Then
        GetReportIDs = Mid$(strResult, 2)
    End If
End Function

Private Sub tvwSelecting_Click()
    If tvwSelecting.SelectedItem Is Nothing Then
        cmdSingleSelect.Enabled = False
    Else
        cmdSingleSelect.Enabled = Not tvwSelecting.SelectedItem.Parent Is Nothing
    End If
End Sub

Private Sub tvwSelecting_DblClick()
    If tvwSelecting.SelectedItem.Parent Is Nothing Then
        cmdSingleSelect.Enabled = False
    Else
        cmdSingleSelect.Enabled = True
        Call cmdSingleSelect_Click
        If tvwSelecting.SelectedItem.Expanded = False Then
            tvwSelecting.SelectedItem.Expanded = True
        End If
    End If
End Sub

Private Sub txtFind_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        '����
        KeyCode = 0
        Call FindItem(True)
    End If
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    If InStr("`~!@#$%^&*()+=[{]}\|;:'"",<.>/?", Chr$(KeyAscii)) > 0 Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub vsfSelected_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    cmdSingleClear.Enabled = vsfSelected.Rows > 1
End Sub

Private Sub vsfSelected_DblClick()
    If vsfSelected.Row < 1 Then Exit Sub
    Call cmdSingleClear_Click
End Sub

Private Function GetRootNode(ByVal objNode As Node) As Node
    If objNode.Parent Is Nothing Then
        Set GetRootNode = objNode
    Else
        Set GetRootNode = GetRootNode(objNode.Parent)
    End If
End Function

Private Sub CollectionAdd(ByRef colVal As Collection, ByVal strKey As String, ByVal varVar As Variant)
    If colVal Is Nothing Then
        Set colVal = New Collection
    End If
    If CollectionFind(colVal, strKey) = False Then
        mcolRevoke.Add varVar, strKey
    End If
End Sub

Private Sub CollectionDelete(ByVal colVal As Collection, ByVal strKey As String)
    If CollectionFind(colVal, strKey) Then
        mcolRevoke.Remove strKey
    End If
End Sub

Private Function CollectionFind(ByVal colVal As Collection, ByVal strKey As String) As Boolean
    On Error Resume Next
    If IsObject(colVal.Item(strKey)) Then
        CollectionFind = Not colVal.Item(strKey) Is Nothing
    Else
        CollectionFind = colVal.Item(strKey) <> ""
    End If
    On Error GoTo 0
End Function

Private Function ReportVerify(ByVal lngReportID As Long, ByVal strName As String) As Boolean
    ReportVerify = False
    
    '��֤����
    If CheckPass(lngReportID) = False Then
        MsgBox mdlPublic.FormatString("��[1]��������֤����ͨ�����ܾ�������", strName) _
            , vbInformation, App.Title
        Exit Function
    End If
    
    'Ȩ��
    If CheckReportPriv(lngReportID) = False Then
        MsgBox mdlPublic.FormatString("��û�С�[1]������������Դ�漰���ݿ����Ĳ�ѯȨ�ޣ����飡", strName) _
            , vbInformation, App.Title
        Exit Function
    End If
    
    ReportVerify = True
End Function

Private Function GetProgID(ByVal lngReportID As Long, Optional blnGroup As Boolean = False) As Long
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo hErr
    
    GetProgID = 0
    If blnGroup Then
        strSQL = "Select ����ID from zlRPTGroups Where ID = [1]"
        Set rsTemp = mdlPublic.OpenSQLRecord(strSQL, "��ȡ�鱨��ĳ���ID", lngReportID)
    Else
        strSQL = "Select ����ID from zlReports Where ID = [1]"
        Set rsTemp = mdlPublic.OpenSQLRecord(strSQL, "��ȡ�������ӱ���ĳ���ID", lngReportID)
    End If
    If rsTemp.RecordCount > 0 Then
        GetProgID = mdlPublic.Nvl(rsTemp!����id, 0)
    End If
    rsTemp.Close
    Exit Function
     
hErr:
    If mdlPublic.ErrCenter = 1 Then Resume
End Function

Private Function GetNodeNext(ByVal objNode As Node) As Node
    If objNode Is Nothing Then
        Exit Function
    Else
        If Not objNode.Next Is Nothing Then
            Set GetNodeNext = objNode.Next
        Else
            Set GetNodeNext = GetNodeNext(objNode.Parent)
        End If
    End If
End Function

Private Function FindItemRecursive(ByVal strFind As String, ByVal objNode As Node) As Node
    If objNode Is Nothing Then
        Exit Function
    End If

    If UCase$(objNode.Text) Like "*" & UCase$(Trim$(strFind)) & "*" And mlngPrevious < objNode.Index Then
        Set FindItemRecursive = objNode
        mlngPrevious = objNode.Index
    Else
        If Not objNode.Child Is Nothing Then
            Set FindItemRecursive = FindItemRecursive(strFind, objNode.Child)
        ElseIf Not objNode.Next Is Nothing Then
            Set FindItemRecursive = FindItemRecursive(strFind, objNode.Next)
        Else
            Set FindItemRecursive = FindItemRecursive(strFind, GetNodeNext(objNode.Parent))
        End If
    End If
End Function

Private Sub FindItem(ByVal blnFirst As Boolean)
    Dim objFind As Node

    If blnFirst Then
        '�״β���
        mlngPrevious = 1
    End If
    
    Set objFind = FindItemRecursive(txtFind.Text, tvwSelecting.Nodes(mlngPrevious))
    If Not objFind Is Nothing Then
        If Not tvwSelecting.SelectedItem Is Nothing Then
            If tvwSelecting.SelectedItem.Index = objFind.Index Then
                tvwSelecting.Nodes(1).Selected = True
            End If
        End If
        objFind.Selected = True
    Else
        If MsgBox("δ���ҵ�ƥ��Ľ�㣬�Ƿ��ͷ��ʼ���ң�", vbQuestion + vbYesNo + vbDefaultButton1, App.Title) = vbYes Then
            Call FindItem(True)
        End If
    End If
End Sub
