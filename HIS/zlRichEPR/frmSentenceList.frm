VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.Unicode.9600.ocx"
Begin VB.Form frmSentenceList 
   BorderStyle     =   0  'None
   Caption         =   "�ʾ�ʾ���б�"
   ClientHeight    =   7590
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6045
   LinkTopic       =   "Form1"
   ScaleHeight     =   7590
   ScaleWidth      =   6045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin XtremeReportControl.ReportControl rptList 
      Height          =   2085
      Left            =   165
      TabIndex        =   1
      Top             =   2385
      Width           =   3000
      _Version        =   589884
      _ExtentX        =   5292
      _ExtentY        =   3678
      _StockProps     =   0
      BorderStyle     =   2
      MultipleSelection=   0   'False
      EditOnClick     =   0   'False
   End
   Begin MSComctlLib.TreeView TreeList 
      Height          =   1470
      Left            =   315
      TabIndex        =   4
      Top             =   705
      Width           =   2310
      _ExtentX        =   4075
      _ExtentY        =   2593
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "imgClass"
      Appearance      =   0
   End
   Begin VB.PictureBox picTerm 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   4050
      Left            =   3330
      ScaleHeight     =   4050
      ScaleWidth      =   2445
      TabIndex        =   2
      Top             =   435
      Visible         =   0   'False
      Width           =   2445
      Begin VSFlex8Ctl.VSFlexGrid vfgTerm 
         Height          =   3690
         Left            =   60
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   120
         Width           =   2340
         _cx             =   4128
         _cy             =   6509
         Appearance      =   2
         BorderStyle     =   0
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
         BackColorFixed  =   16761024
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   16777215
         GridColor       =   -2147483643
         GridColorFixed  =   -2147483643
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   2
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   4
         Cols            =   1
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
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
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   750
      Top             =   4320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSentenceList.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSentenceList.frx":059A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSentenceList.frx":0B34
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSentenceList.frx":10CE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox rtbText 
      Height          =   1755
      Left            =   180
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   4695
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   3096
      _Version        =   393217
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmSentenceList.frx":19A8
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
   Begin MSComctlLib.ImageList imgClass 
      Left            =   2025
      Top             =   300
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSentenceList.frx":1A45
            Key             =   "close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSentenceList.frx":1FDF
            Key             =   "expend"
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgListTmp 
      Height          =   420
      Left            =   3960
      TabIndex        =   5
      Top             =   5775
      Visible         =   0   'False
      Width           =   585
      _cx             =   1032
      _cy             =   741
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
      BackColorFixed  =   15790320
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
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
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
      AutoResize      =   0   'False
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
   Begin XtremeDockingPane.DockingPane dkpMan 
      Left            =   855
      Top             =   105
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   195
      Top             =   30
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
End
Attribute VB_Name = "frmSentenceList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'######################################################################################################################

'��������
'------------------------------------------------------------------------------------------------------------------
Private Enum mCol
    ͼ�� = 0: ID: ����id: ����: ���: ����: ����: ��Ա
End Enum

Private Const con_UnDefine = -999
Private Const conPane_Tree = 400
Private Const conPane_List = 401
Private Const conPane_Term = 403
Private Const conPane_Text = 404


'�������
'------------------------------------------------------------------------------------------------------------------
Private mfrmParent As Form          '��ǰ������ϼ�����
Private mstrPrivs As String         '��ǰʹ����Ȩ�޴�
Private mlngWordId As Long          '��ǰ�ʾ�id
Private mblnCompend As Boolean      '������оٴʾ䣺ά��������򰴷����о٣������༭�а�����о�
Private mlngParentId As Long        '��id����Ϊ������ʱ��Ϊ����Idֱ��ƥ��ķ���id����Ϊ�����ʱ��Ϊ�����ļ��ṹ�����id
Private mlngClassId As Long         '�ʾ�Ĭ�Ϸ���id,���������о�ʱ����mlngParentId��ͬ����������о�ʱ��Ϊ��Ӧ��һ������id
Private mlngPatient As Long          '����id���ڲ��˲����༭ʱ������ȷ�������ʾ��Ƿ�����
Private mlngVisit As Long           '��ҳid��Һŵ�ID
Private mlngAdvice As Long          'ҽ��ID
Private mstrSecondLimit As String   '���ж��ι��˵Ĵʾ�ID�����Զ��ŷָ�

Private mintPower As Integer        '�ʾ����Ȩ��Χ
'    mintPower=con_UnDefine��δ����;
'    mintPower=-1�����߱��ʾ����Ȩ;
'    mintPower=0��ȫԺ����ʱ��ʾ���е�ʾ����Ҳ���Ը���;
'    mintPower=1�����ң���ʱ��ʾȫԺͨ��ʾ��(����id is null)�����ڿ��ҹ��л�������Ա˽�е�ʾ���������ܸ���ȫԺͨ��ʾ��;
'    mintPower=2�����ˣ���ʱ��ʾȫԺͨ��ʾ��(����id is null)�����ڿ���ͨ��ʾ��(��Աid is null)�͸���ʾ����������ʾ���ɸ���
Public Event RowDblClick(ByVal lngSentenceID As Long)    '˫��һ�л������ϰ��س�


'����Ϊ�ⲿ��������
'######################################################################################################################
Public Function zlRefFromClass(ByVal frmParent As Form, ByVal lngClassId As Long) As Long
    '******************************************************************************************************************
    '���ܣ�����ָ�����࣬ˢ���б�����ά������ӿ�
    '������ָ���Ĵʾ�ʾ������id
    '******************************************************************************************************************
    Set mfrmParent = frmParent
    mblnCompend = False
    TreeList.Visible = mblnCompend
    picTerm.Visible = Not mblnCompend
    If Not mblnCompend Then
        dkpMan.FindPane(conPane_Tree).Close
        dkpMan.FindPane(conPane_Term).Closed = False
    End If
    If mlngParentId = lngClassId Then zlRefFromClass = Me.rptList.Rows.Count: Exit Function
    mlngParentId = lngClassId
    mlngClassId = lngClassId
    
    rptList.Columns(mCol.����).Visible = True
    rptList.Columns(mCol.��Ա).Visible = True
    
    zlRefFromClass = zlSubRefList(mlngWordId)
End Function

Public Function zlRefFromCompend(ByVal frmParent As Form, _
                                ByVal lngCompendID As Long, _
                                Optional lngPatient As Long = 0, _
                                Optional lngVisit As Long = 0, _
                                Optional lngAdvice As Long = 0, _
                                Optional blnForce As Boolean, _
                                Optional strSecondLimit As String) As Long
    '******************************************************************************************************************
    '���ܣ� ����ָ����٣�ˢ���б����ڲ����༭�ӿ�
    '������ ָ�����ļ��������id
    '       lngPatient������id
    '       lngVisit�����˾���ID�����ﲡ��Ϊ�Һ�ID��סԺ����Ϊ��ҳid
    '       lngAdvice��ҽ��ID
    '       lngCompendID�����id
    '******************************************************************************************************************
    
    Dim rsTemp As New ADODB.Recordset
    Dim panThis As Pane
    
    If blnForce = False And mlngParentId = lngCompendID And _
        (mstrSecondLimit = strSecondLimit) Then zlRefFromCompend = Me.rptList.Rows.Count: Exit Function
    Set mfrmParent = frmParent
    mblnCompend = True
    TreeList.Visible = mblnCompend
    picTerm.Visible = Not mblnCompend
    mlngParentId = lngCompendID
    mlngPatient = lngPatient
    mlngVisit = lngVisit
    mlngAdvice = lngAdvice
    mstrSecondLimit = strSecondLimit
    
    rptList.Columns(mCol.����).Visible = False
    rptList.Columns(mCol.��Ա).Visible = False
    Set panThis = dkpMan.FindPane(conPane_Term)
    panThis.Close
    Set panThis = dkpMan.FindPane(conPane_Tree)
    panThis.Closed = False
    
    gstrSQL = "Select �ʾ����id From ������ٴʾ� Where ���id = [1]"
    Err = 0: On Error GoTo errHand
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngParentId)
    mlngClassId = 0
    If rsTemp.RecordCount > 0 Then mlngClassId = rsTemp.Fields(0).Value
    
    Call zlSubRefClass
    
    zlRefFromCompend = zlSubRefList(mlngWordId, 0)
    
    Exit Function
    
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    
    Call SaveErrLog
    zlRefFromCompend = 0
End Function

Public Sub zlAddFromEditor()
    '******************************************************************************************************************
    '���ܣ�ִ��ָ���������ؼ������ڲ����༭�ӿ�
    '��������ǰ�ı༭������
    '******************************************************************************************************************
Dim lngRetuId As Long
Dim cbrControl As CommandBarControl
    
    If mlngClassId = 0 Then
        MsgBox "��ǰ���û�����ôʾ�ʾ�������Ӧ������ϵ����Ա��ʼ���������ݣ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If mintPower < 0 Then
        MsgBox "�㲻�߱��ʾ�ʾ�������Ȩ�ޣ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    Set cbrControl = Me.cbsThis.ActiveMenuBar.FindControl(xtpControlButton, conMenu_Edit_NewItem, True, True)
    If cbrControl Is Nothing Then Exit Sub
    If cbrControl.Visible = False Or cbrControl.Enabled = False Then Exit Sub
    
    lngRetuId = frmSentenceEdit.ShowMe(mfrmParent, True, mintPower, mlngClassId, , True)
    If lngRetuId = 0 Then Exit Sub
    
    Call zlSubRefList(lngRetuId)
End Sub

Public Sub zlExecuteControl(ByVal Control As XtremeCommandBars.ICommandBarControl)
    '******************************************************************************************************************
    '���ܣ�ִ��ָ���������ؼ�������ά������ӿ�
    '������ָ�����ҵĿؼ�
    '******************************************************************************************************************
    Call cbsThis_Execute(Control)
End Sub

Public Sub zlUpdateControl(ByVal Control As XtremeCommandBars.ICommandBarControl)
    '******************************************************************************************************************
    '���ܣ�ˢ���������ؼ�״̬������ά������ӿ�
    '������ָ�����ҵĿؼ�
    '******************************************************************************************************************
    Call cbsThis_Update(Control)
End Sub


'����Ϊ�ڲ���������
'######################################################################################################################
Private Function zlGetPower() As Integer
    '******************************************************************************************************************
    '���ܣ���õ�ǰ�û��Ĵʾ�����Ȩ��
    '���أ��ʾ����Ȩ����ֵ
    '******************************************************************************************************************
    If mintPower = con_UnDefine Then
        If InStr(1, gstrPrivsEpr, "ȫԺ�����ʾ�") <> 0 Then
            mintPower = 0
        ElseIf InStr(1, gstrPrivsEpr, "���Ҳ����ʾ�") <> 0 Then
            mintPower = 1
        ElseIf InStr(1, gstrPrivsEpr, "���˲����ʾ�") <> 0 Then
            mintPower = 2
        Else
            mintPower = -1
        End If
    End If
    zlGetPower = mintPower
End Function

Private Function zlSubRefClass() As Boolean
    '******************************************************************************************************************
    '���ܣ�ˢ�·���
    '������
    '���أ�
    '******************************************************************************************************************
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    
    If mblnCompend = False Then Exit Function
    
    gstrSQL = "Select /*+ rule*/ Id,�ϼ�id,����,���� From �����ʾ���� Start With Id In ("
    
    
    '------------------------------------------------------------------------------------------------------------------
    strSQL = "Select L.����id " & vbNewLine & _
            "From �����ʾ�ʾ�� L, ������ٴʾ� A," & vbNewLine & _
            "     Table(Cast(f_Sentence_Usable([1], [2], [3], [4]) As " & gstrDbOwner & ".t_Dic_Rowset)) U" & vbNewLine & _
            "Where L.����id = A.�ʾ����id  And A.���id = [1] And L.ID = To_Number(U.����)"
            
    Select Case mintPower
    Case 0
    Case 1
        strSQL = strSQL & " And" & vbNewLine & _
                "      (Nvl(L.ͨ�ü�, 0) = 0 Or" & vbNewLine & _
                "      L.ͨ�ü� In (1, 2) And" & vbNewLine & _
                "      L.����id In (Select R.����id From ������Ա R, �ϻ���Ա�� U Where R.��Աid = U.��Աid And U.�û��� = User))"

    Case Else
        strSQL = strSQL & " And" & vbNewLine & _
                "      (Nvl(L.ͨ�ü�, 0) = 0 Or" & vbNewLine & _
                "      L.ͨ�ü� = 1 And" & vbNewLine & _
                "      L.����id In (Select R.����id From ������Ա R, �ϻ���Ա�� U Where R.��Աid = U.��Աid And U.�û��� = User) Or" & vbNewLine & _
                "      L.ͨ�ü� = 2 And L.��Աid In (Select U.��Աid From �ϻ���Ա�� U Where U.�û��� = User))"
    End Select
    
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = gstrSQL & strSQL
    gstrSQL = gstrSQL & ") Connect By Prior �ϼ�id=Id  Order By ����"
    
    Dim objNode As Node
    
    TreeList.Nodes.Clear
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngParentId, mlngPatient, mlngVisit, mlngAdvice)
    If rsTemp.BOF = False Then
                
        Set objNode = TreeList.Nodes.Add(, , "K0", "���дʾ�", "close", "expend")
        objNode.Expanded = False
        Do While Not rsTemp.EOF
            
            Set objNode = Nothing
            
            On Error Resume Next
            Set objNode = TreeList.Nodes("K" & rsTemp("ID").Value)
            On Error GoTo errHand
            
            If objNode Is Nothing Then
                Set objNode = TreeList.Nodes.Add("K" & zlCommFun.NVL(rsTemp("�ϼ�id").Value, 0), tvwChild, "K" & rsTemp("ID").Value, rsTemp("����").Value, "close", "expend")
                objNode.Expanded = False
            End If
            rsTemp.MoveNext
        Loop
    End If
    If TreeList.Nodes.Count > 0 Then
        TreeList.Nodes(1).Selected = True
    End If
    
    zlSubRefClass = True
    
    Exit Function
errHand:
    
End Function

Private Function zlSubRefList(Optional lngID As Long, Optional ByVal lng����id As Long) As Long
    '******************************************************************************************************************
    '���ܣ�ˢ��װ���嵥������λ��ָ���ļ�¼��
    '������
    '���أ�
    '******************************************************************************************************************
Dim rsTemp As New ADODB.Recordset
Dim strClassIds As String, strKinds As String, strText As String, blnAdd As Boolean
Dim rptRcd As ReportRecord
Dim rptItem As ReportRecordItem
Dim rptRow As ReportRow
    
    '------------------------------------------------------------------------------------------------------------------
    If mblnCompend = False Then
        '��������ʾ������ά�����������
        
        gstrSQL = "Select /*+ rule*/ L.ID, L.����id, C.���� || '-' || C.���� As ����, L.���, L.����, L.ͨ�ü�, D.���� As ����, P.���� As ��Ա" & vbNewLine & _
                "From �����ʾ���� C, �����ʾ�ʾ�� L, ���ű� D, ��Ա�� P" & vbNewLine & _
                "Where C.ID = L.����id And L.����id = D.ID And L.��Աid = P.ID And L.����id = [1] "
    Else
        '�������ʾ�����ڲ����༭������
        gstrSQL = "Select /*+ rule*/ L.ID, L.����id, C.���� || '-' || C.���� As ����, L.���, L.����, L.ͨ�ü�, D.���� As ����, P.���� As ��Ա" & vbNewLine & _
                "From �����ʾ���� C, �����ʾ�ʾ�� L, ������ٴʾ� A, ���ű� D, ��Ա�� P," & vbNewLine & _
                "     Table(Cast(f_Sentence_Usable([1], [2], [3], [4]) As " & gstrDbOwner & ".t_Dic_Rowset)) U" & vbNewLine & _
                "Where C.ID = L.����id And L.����id = A.�ʾ����id And L.����id = D.ID And L.��Աid = P.ID And A.���id = [1] And" & vbNewLine & _
                "      L.ID = To_Number(U.����)  "
        If lng����id > 0 Then gstrSQL = gstrSQL & "  And L.����id=[5] "
    End If
    
    '------------------------------------------------------------------------------------------------------------------
    Select Case mintPower
    Case 0
    Case 1
        gstrSQL = gstrSQL & " And" & vbNewLine & _
                "      (Nvl(L.ͨ�ü�, 0) = 0 Or" & vbNewLine & _
                "      L.ͨ�ü� In (1, 2) And" & vbNewLine & _
                "      L.����id In (Select R.����id From ������Ա R, �ϻ���Ա�� U Where R.��Աid = U.��Աid And U.�û��� = User))"

    Case Else
        gstrSQL = gstrSQL & " And" & vbNewLine & _
                "      (Nvl(L.ͨ�ü�, 0) = 0 Or" & vbNewLine & _
                "      L.ͨ�ü� = 1 And" & vbNewLine & _
                "      L.����id In (Select R.����id From ������Ա R, �ϻ���Ա�� U Where R.��Աid = U.��Աid And U.�û��� = User) Or" & vbNewLine & _
                "      L.ͨ�ü� = 2 And L.��Աid In (Select U.��Աid From �ϻ���Ա�� U Where U.�û��� = User))"
    End Select
    
    Err = 0: On Error GoTo errHand
    If mblnCompend = False Then
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngParentId)
    Else
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngParentId, mlngPatient, mlngVisit, mlngAdvice, lng����id)
    End If
    
    '------------------------------------------------------------------------------------------------------------------
    rptList.GroupsOrder.DeleteAll
    rptList.Records.DeleteAll
    strClassIds = ","
    With rsTemp
        Do While Not .EOF
            blnAdd = True
            If mstrSecondLimit <> "" Then '���ι���
                If InStr(mstrSecondLimit, "," & !ID & ",") = 0 Then blnAdd = False  '���ڶ��ι��˷�Χ���򲻼����б�
            End If
            
            If blnAdd Then
                If InStr(1, strClassIds, "," & !����id & ",") = 0 Then strClassIds = strClassIds & !����id & ","
                Set rptRcd = Me.rptList.Records.Add()
                Set rptItem = rptRcd.AddItem(CInt(Val("" & !ͨ�ü�))): rptItem.Icon = rptItem.Value
                Select Case rptItem.Value
                Case 0: rptItem.GroupCaption = "1-ȫԺ"
                Case 1: rptItem.GroupCaption = "2-����"
                Case Else: rptItem.GroupCaption = "3-����"
                End Select
                rptRcd.AddItem CStr(!ID)
                rptRcd.AddItem CStr("" & !����id)
                rptRcd.AddItem CStr("" & !����)
                rptRcd.AddItem CStr("" & !���)
                rptRcd.AddItem CStr("" & !����)
                rptRcd.AddItem CStr("" & !����)
                rptRcd.AddItem CStr("" & !��Ա)
            End If
            .MoveNext
        Loop
    End With
    
    If mblnCompend = True And UBound(Split(strClassIds, ",")) > 2 Then Me.rptList.GroupsOrder.Add Me.rptList.Columns(mCol.����)
    Me.rptList.Populate
    
    If lngID <> 0 Then
        For Each rptRow In Me.rptList.Rows
            If rptRow.GroupRow = False Then
                If Val(rptRow.Record(mCol.ID).Value) = lngID Then
                    Set Me.rptList.FocusedRow = rptRow: Exit For
                End If
            End If
        Next
    Else
        For Each rptRow In Me.rptList.Rows
            If rptRow.GroupRow Then rptRow.Expanded = False
        Next
    End If
    If Me.rptList.Rows.Count > 0 And (Me.rptList.FocusedRow Is Nothing) Then
        Set Me.rptList.FocusedRow = Me.rptList.Rows(0)
    End If
    
    Call rptList_SelectionChanged
    zlSubRefList = Me.rptList.Records.Count
    Exit Function

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlSubRefList = Me.rptList.Records.Count
End Function

'����Ϊ�ؼ��¼�����
'######################################################################################################################
Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
Dim lngRetuId As Long, strTemp As String
    
    Err = 0: On Error GoTo errHand
    '------------------------------------
    Select Case Control.ID
    Case conMenu_Edit_NewItem
        If Not (Me.rptList.FocusedRow Is Nothing) Then mlngClassId = Me.rptList.FocusedRow.Record(mCol.����id).Value
        lngRetuId = frmSentenceEdit.ShowMe(mfrmParent, True, mintPower, mlngClassId)
        If lngRetuId = 0 Then Exit Sub
        Call zlSubRefList(lngRetuId)
    Case conMenu_Edit_Modify
        mlngClassId = Me.rptList.FocusedRow.Record(mCol.����id).Value
        lngRetuId = frmSentenceEdit.ShowMe(mfrmParent, False, mintPower, mlngClassId, mlngWordId)
        If lngRetuId = 0 Then Exit Sub
        Call zlSubRefList(lngRetuId)
    Case conMenu_Edit_Delete
        strTemp = "���ɾ���ôʾ���" & vbCrLf & "����" & Me.rptList.FocusedRow.Record(mCol.���).Value & "-" & Me.rptList.FocusedRow.Record(mCol.����).Value
        If MsgBox(strTemp, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        gstrSQL = "Zl_�����ʾ�ʾ��_Edit(3," & mlngWordId & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        With Me.rptList
            mlngWordId = 0: lngRetuId = .FocusedRow.Index
            If .Rows.Count > lngRetuId + 1 Then
                mlngWordId = .Rows(lngRetuId + 1).Record(mCol.ID).Value
            ElseIf lngRetuId > 0 Then
                mlngWordId = .Rows(lngRetuId - 1).Record(mCol.ID).Value
            End If
        End With
        Call zlSubRefList(mlngWordId)
    Case conMenu_Edit_Request
        If frmSentenceRequest.ShowMe(mfrmParent, mlngWordId) = True Then Call zlSubRefList(mlngWordId)
    Case conMenu_File_Preview: Call zlRptPrint(0)
    Case conMenu_File_Print: Call zlRptPrint(1)
    Case conMenu_File_Excel: Call zlRptPrint(3)
    End Select
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Exit Sub
End Sub



Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Err = 0: On Error Resume Next
    Select Case Control.ID
    Case conMenu_Edit_NewItem
        Control.Visible = (mintPower >= 0 And mlngClassId <> 0)
    Case conMenu_Edit_Modify, conMenu_Edit_Delete, conMenu_Edit_Request
        Control.Visible = (mintPower >= 0 And mlngClassId <> 0)
        Control.Enabled = (mlngWordId <> 0)
        If Control.Enabled Then Control.Enabled = (Me.rptList.FocusedRow.Record(mCol.ͼ��).Value >= mintPower)
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        Control.Enabled = (Me.rptList.Records.Count <> 0)
    End Select
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    On Error Resume Next
    Select Case Item.ID
    Case conPane_Tree
        Item.Handle = TreeList.hwnd
    Case conPane_List
        Item.Handle = rptList.hwnd
    Case conPane_Term
        Item.Handle = Me.picTerm.hwnd
    Case conPane_Text
        Item.Handle = Me.rtbText.hwnd
    End Select
End Sub

Private Sub Form_Load()
    '-----------------------------------------------------
    'Ȩ�����ƴ����ƣ�����ͬʱ��������ģ�������gmstrPrivs�仯�����¿�����Ч
Dim cbrControl As CommandBarControl
Dim cbrMenuBar As CommandBarPopup
Dim rptCol As ReportColumn
    mintPower = con_UnDefine
    mintPower = zlGetPower
    mlngWordId = 0
    
    Call zlCommFun.SetWindowsInTaskBar(Me.hwnd, False)
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbsThis.VisualTheme = xtpThemeOffice2003
    Set Me.cbsThis.Icons = zlCommFun.GetPubIcons
    With Me.cbsThis.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    Me.cbsThis.EnableCustomization False
    
    '-----------------------------------------------------
    '�˵�����
    Me.cbsThis.ActiveMenuBar.Title = "�˵�": Me.cbsThis.ActiveMenuBar.Visible = False
    Me.cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)

    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "�༭(&E)", -1, False)
    cbrMenuBar.ID = conMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "����(&A)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸�(&M)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ��(&D)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Request, "��������(&Q)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��(&V)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ(&P)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Excel, "�����&Excel��")
    End With
    
    '-----------------------------------------------------
    '���ôʾ���ʾͣ������
    Dim panThis As Pane
    
    Set panThis = dkpMan.CreatePane(conPane_Tree, 600, 300, DockTopOf, Nothing)
    panThis.Title = "���ͽṹ"
    panThis.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
    Set panThis = dkpMan.CreatePane(conPane_List, 600, 450, DockBottomOf, panThis)
    panThis.Title = "�����б�"
    panThis.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
    Set panThis = dkpMan.CreatePane(conPane_Text, 600, 800, DockBottomOf, panThis)
    panThis.Title = "ʾ������"
    panThis.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
    Set panThis = dkpMan.CreatePane(conPane_Term, 200, 800, DockRightOf, Nothing)
    panThis.Title = "ʾ������"
    panThis.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    panThis.Close 'Ĭ����������������ǲ���ʾ��
    
    Me.dkpMan.SetCommandBars Me.cbsThis
    Me.dkpMan.Options.ThemedFloatingFrames = True
    Me.dkpMan.Options.HideClient = True
    dkpMan.LoadStateFromString GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name & "\" & TypeName(dkpMan), dkpMan.Name, "")
    '-----------------------------------------------------
    With Me.rptList
        Set rptCol = .Columns.Add(mCol.ͼ��, "", 18, False): rptCol.Editable = False: rptCol.Groupable = False
        rptCol.Sortable = False: rptCol.Alignment = xtpAlignmentCenter
        Set rptCol = .Columns.Add(mCol.ID, "ID", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.����id, "����id", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.����, "����", 200, True): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.���, "���", 50, False): rptCol.Editable = False: rptCol.Groupable = False: .SortOrder.Add rptCol
        Set rptCol = .Columns.Add(mCol.����, "����", 120, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.����, "����", 70, False): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.��Ա, "��Ա", 56, False): rptCol.Editable = False: rptCol.Groupable = False
        .SetImageList Me.imgList
        .AllowColumnRemove = False
        .MultipleSelection = False
        .ShowItemsInGroups = False
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .GridLineColor = RGB(225, 225, 225)
            .NoGroupByText = "�϶��б��⵽����,�����з���..."
            .NoItemsText = "û�п���ʾ����Ŀ..."
            .VerticalGridStyle = xtpGridSolid
        End With
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Set mfrmParent = Nothing
    Call SaveSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name & "\" & TypeName(dkpMan), dkpMan.Name, dkpMan.SaveStateToString)
End Sub

Private Sub picTerm_Resize()
    Err = 0: On Error Resume Next
    With Me.vfgTerm
        .Left = 0: .Width = Me.picTerm.ScaleWidth
        .Top = 0: .Height = Me.picTerm.ScaleHeight
        .AutoSize 0
    End With
End Sub

Private Sub rptList_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    With Me.rptList
        If .Visible = False Then Exit Sub
        If .FocusedRow Is Nothing Then Exit Sub
        If .FocusedRow.GroupRow Then Exit Sub
        Call rptList_RowDblClick(.FocusedRow, .FocusedRow.Record.Item(mCol.ID))
    End With
End Sub

Private Sub rptList_MouseUp(Button As Integer, Shift As Integer, x As Long, y As Long)
Dim cbrPopupBar As CommandBar
Dim cbrPopupItem As CommandBarControl
Dim cbrControl As CommandBarControl
Dim cbrMenuBar As CommandBarPopup
    
    If Button <> vbRightButton Then Exit Sub
     
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.FindControl(xtpControlPopup, conMenu_EditPopup)
    If cbrMenuBar Is Nothing Then Exit Sub
    If cbrMenuBar.Visible = False Then Exit Sub
    
    Set cbrPopupBar = Me.cbsThis.Add("�����˵�", xtpBarPopup)
    For Each cbrControl In cbrMenuBar.CommandBar.Controls
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, cbrControl.ID, cbrControl.Caption)
        cbrPopupItem.BeginGroup = cbrControl.BeginGroup
    Next
    cbrPopupBar.ShowPopup
End Sub

Private Sub rptList_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
Dim cbrControl As CommandBarControl
    If Me.rptList.FocusedRow Is Nothing Then
        mlngWordId = 0
    ElseIf Me.rptList.FocusedRow.GroupRow = True Then
        mlngWordId = 0
    Else
        mlngWordId = Me.rptList.FocusedRow.Record.Item(mCol.ID).Value
    End If
    If mlngWordId = 0 Then Exit Sub
    
    If mblnCompend = False Then
        If rptList.FocusedRow.Record(mCol.ͼ��).Value >= mintPower Then
            Set cbrControl = Me.cbsThis.ActiveMenuBar.FindControl(xtpControlButton, conMenu_Edit_Modify, True, True)
            If cbrControl Is Nothing Then Exit Sub
            If cbrControl.Visible = False Or cbrControl.Enabled = False Then Exit Sub
            Call cbsThis_Execute(cbrControl)
        End If
    Else
        RaiseEvent RowDblClick(mlngWordId)
    End If
End Sub

Private Sub rptList_SelectionChanged()
    Dim rsTemp As New ADODB.Recordset
    Dim lngStart As Long, strText As String
    
    If Me.rptList.FocusedRow Is Nothing Then
        mlngWordId = 0
    ElseIf Me.rptList.FocusedRow.GroupRow = True Then
        mlngWordId = 0
    Else
        mlngWordId = Me.rptList.FocusedRow.Record.Item(mCol.ID).Value
    End If
    If Me.Visible = False Then Exit Sub

    'ˢ�´ʾ�����
    '------------------------------------------------------------------------------------------------------------------
    Me.rtbText.Text = ""
    Err = 0: On Error GoTo errHand
    gstrSQL = "Select ��������, �����ı�, Ҫ������, Ҫ�ص�λ From �����ʾ���� Where �ʾ�id = [1] Order By ���д���"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngWordId)
    With rsTemp
        Do While Not .EOF
            lngStart = Len(Me.rtbText.Text)
            Me.rtbText.SelStart = lngStart
            Me.rtbText.SelLength = 0
            Select Case !��������
            Case 0 '��������
                strText = IIf(IsNull(!�����ı�), " ", !�����ı�)
                With Me.rtbText
                    .SelText = strText: .SelStart = lngStart: .SelLength = Len(strText)
                    .SelUnderline = False
                End With
            Case 1, 2 '1-��ʱ����Ҫ��,2-�̶�����Ҫ��
                strText = IIf(IsNull(!�����ı�), "{" & !Ҫ������ & "}" & !Ҫ�ص�λ, "{" & !�����ı� & "}")
                With Me.rtbText
                    .SelText = strText: .SelStart = lngStart: .SelLength = Len(strText)
                    .SelUnderline = True
                End With
            End Select
            .MoveNext
        Loop
        Me.rtbText.SelStart = 0
    End With
    
    'ˢ�´ʾ�����
    Dim panThis As Pane
    Set panThis = Me.dkpMan.FindPane(conPane_Term)
    If panThis Is Nothing Then Exit Sub
    If panThis.Closed Then Exit Sub
    
    Me.vfgTerm.Clear: Me.vfgTerm.Rows = Me.vfgTerm.FixedRows
    Set Me.vfgTerm.Cell(flexcpPicture, Me.vfgTerm.FixedRows - 1, 0) = Me.imgList.ListImages(4).Picture
    gstrSQL = "Select ���� As ������, ���� As ����ֵ" & vbNewLine & _
            "From Table(Cast(f_Sentence_������([1]) As " & gstrDbOwner & ".t_Dic_Rowset))" & vbNewLine & _
            "Where ���� Is Not Null"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngWordId)
    With rsTemp
        If .RecordCount <= 0 Then
            Me.vfgTerm.TextMatrix(Me.vfgTerm.FixedRows - 1, 0) = "��ʹ������������"
        Else
            Me.vfgTerm.TextMatrix(Me.vfgTerm.FixedRows - 1, 0) = "��������������ʱ����ʹ�ã�"
        End If
        Do While Not .EOF
            Me.vfgTerm.Rows = Me.vfgTerm.Rows + 1
            Me.vfgTerm.TextMatrix(Me.vfgTerm.Rows - 1, 0) = Space(2) & Me.vfgTerm.Rows - 1 & ")" & !������ & "Ϊ'" & Replace(!����ֵ, vbTab, "'��'") & "'"
            .MoveNext
        Loop
    End With
    Me.vfgTerm.AutoSize 0
    
    Exit Sub
    
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub TreeList_NodeClick(ByVal Node As MSComctlLib.Node)
    Call zlSubRefList(mlngWordId, Val(Mid(Node.Key, 2)))
End Sub
Private Sub zlRptPrint(ByVal bytMode As Byte)
    '******************************************************************************************************************
    '����:�����ݸ��Ƶ��ɴ�ӡ�Ķ��󣬵��ô�ӡ
    '����:  bytMode��1-��ӡ;2-Ԥ��;3-�����EXCEL
    '******************************************************************************************************************
    
    Dim objPrint As New zlPrint1Grd
    Dim objAppRow As zlTabAppRow
    
    If Me.rptList.Records.Count = 0 Then Exit Sub
    
    '-------------------------------------------------
    '�������ݱ��
    If zlReportToVSFlexGrid(Me.vfgListTmp, Me.rptList) = False Then Exit Sub
    
    '-------------------------------------------------
    '���ô�ӡ��������

    
    Set objPrint.Body = Me.vfgListTmp
    objPrint.Title.Text = "�ʾ�ʾ���嵥"
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("")
    Call objAppRow.Add("��ӡʱ��:" & Now())
    Call objPrint.BelowAppRows.Add(objAppRow)
    
    If bytMode = 1 Then
        bytMode = zlPrintAsk(objPrint)
        If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
End Sub
