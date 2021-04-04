VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{853AAF94-E49C-11D0-A303-0040C711066C}#4.3#0"; "DicomObjects.ocx"
Object = "*\Azl9PacsControl\zl9PacsControl.vbp"
Begin VB.UserControl ucReportHistory 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   9720
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6300
   ScaleHeight     =   9720
   ScaleWidth      =   6300
   Begin VB.PictureBox picBack 
      BorderStyle     =   0  'None
      Height          =   9015
      Left            =   120
      ScaleHeight     =   9015
      ScaleWidth      =   6015
      TabIndex        =   0
      Top             =   120
      Width           =   6015
      Begin zl9PacsControl.ucSplitter ucSplitter1 
         Height          =   135
         Left            =   0
         TabIndex        =   1
         Top             =   2895
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   238
         MousePointer    =   7
         SplitType       =   0
         SplitLevel      =   3
         Con1MinSize     =   1000
         Con2MinSize     =   2000
         Control1Name    =   "vsfStudy"
         Control2Name    =   "picContext"
      End
      Begin VB.PictureBox picContext 
         BorderStyle     =   0  'None
         Height          =   5985
         Left            =   0
         ScaleHeight     =   5985
         ScaleWidth      =   6015
         TabIndex        =   3
         Top             =   3030
         Width           =   6015
         Begin VB.CheckBox chkLinkView 
            BackColor       =   &H00FFFFFF&
            Caption         =   "�� �����鿴"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Left            =   4680
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   30
            Width           =   1335
         End
         Begin VB.CommandButton cmdWrite 
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   495
            Left            =   480
            Picture         =   "ucReportHistory.ctx":0000
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "д�뱨������"
            Top             =   0
            Width           =   495
         End
         Begin DicomObjects.DicomViewer dcmReportImg 
            Height          =   3495
            Left            =   240
            TabIndex        =   6
            Top             =   720
            Visible         =   0   'False
            Width           =   4455
            _Version        =   262147
            _ExtentX        =   7858
            _ExtentY        =   6165
            _StockProps     =   35
            BackColor       =   0
         End
         Begin VB.CheckBox chkReportType 
            BackColor       =   &H00FFFFFF&
            DownPicture     =   "ucReportHistory.ctx":1A72
            Height          =   495
            Left            =   0
            Picture         =   "ucReportHistory.ctx":34E4
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "�ı�����"
            Top             =   0
            Width           =   495
         End
         Begin RichTextLib.RichTextBox rtxtReport 
            Height          =   5505
            Left            =   0
            TabIndex        =   4
            Top             =   480
            Width           =   6015
            _ExtentX        =   10610
            _ExtentY        =   9710
            _Version        =   393217
            ScrollBars      =   2
            Appearance      =   0
            AutoVerbMenu    =   -1  'True
            TextRTF         =   $"ucReportHistory.ctx":4F56
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfStudy 
         Height          =   2895
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   6015
         _cx             =   10610
         _cy             =   5106
         Appearance      =   1
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
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   14737632
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16761024
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
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
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
         ExplorerBar     =   3
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
End
Attribute VB_Name = "ucReportHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const M_STR_LISTVIEWKEY_DESCRIBE As String = "describe" '����޼��������
Private Const M_STR_LISTVIEWKET_PROCESS As String = "process"   '������̱�����


Private Const M_STR_COLNAME = "���|ҽ��ID|����|����|���|��Ŀ|��λ|������|��ǰ����|���ʱ��|ҽ������|�������|�������|�ؼ�ID|ת��״̬"


Private Type TPListCfg
    strSortPro As String
    strColPros As String
End Type


Public Event OnSend()
Public Event OnClick()
Public Event OnDbClick()
Public Event OnMouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event OnLinkView(ByVal lngAdviceId As Long, ByVal blnMoved As Boolean, ByVal blnIsDBClick As Boolean)

Private mblnIsInit As Boolean

Private mobjOwner As Object
Private mlngCurModule As Long
Private mlngCurDeptId As Long
Private mstrPrivs As String

Private mstrGrantDeptIds As String
Private mblnAllDepts As Boolean
Private mlngAdviceId As Long
 

Private mTPListCfg As TPListCfg

Private mstrDescriptionName As String   '���������������
Private mstrAdviseName As String    '��Ͻ����������
Private mstrOpinionName As String   '��������������
Private mstrPatholMaterialInfo As String 'ȡ����ʾ��Ŀ����

Private mdtBegin As Date
Private mdtEnd As Date
Private mblnCustom As Boolean

Private mlngPatientId As Long
Private mlngPatientFrom As Long
Private mlngBabyNum As Long
Private mlngLinkID As Long

Private mstrDateRange As String
Private mblnIsThisTime As Boolean   '�Ƿ񱾴����
Private mblnIsOtherDept As Boolean  '�Ƿ����Ƽ��
Private mblnIsAutoLine As Boolean   '�Ƿ��Զ�����

Private mblnHistoryMoved As Boolean     '��ʷ��¼�Ƿ��н���ת��
Private mblnAdviceMoved As Boolean      '��ǰҽ���Ƿ������ת��

Private mlngSelImgIndex As Long
Private mftpConTag As TFtpConTag
Private mblnAllowWrite As Boolean


'�Ƿ�����д��
Property Get AllowWrite() As Boolean
    AllowWrite = mblnAllowWrite
End Property

Property Let AllowWrite(ByVal value As Boolean)
    mblnAllowWrite = value
    
    If mblnAllowWrite = False Then
        cmdWrite.Visible = False
    Else
        If vsfStudy.Rows > 1 Then cmdWrite.Visible = True
    End If
End Property

Property Get AllowLinkViewer() As Boolean
    AllowLinkViewer = chkLinkView.Visible
End Property

Property Let AllowLinkViewer(ByVal value As Boolean)
    chkLinkView.Visible = value
End Property

'�Ƿ������鿴ģʽ
Property Get LinkViewed() As Boolean
    LinkViewed = IIf(chkLinkView.value <> 0, True, False)
End Property

Property Let LinkViewed(ByVal value As Boolean)
    chkLinkView.value = IIf(value, 1, 0)
End Property

'��ʷ����
Property Get HistoryCount()
    HistoryCount = vsfStudy.Rows - 1
End Property


Property Get hwnd() As Long
    hwnd = UserControl.hwnd
End Property

'���ڷ�Χ
Property Get DataRange() As String
    DataRange = mstrDateRange
End Property

'�������
Property Get IsThisTime() As Boolean
    IsThisTime = mblnIsThisTime
End Property

Property Let IsThisTime(ByVal value As Boolean)
    mblnIsThisTime = value
End Property


'���Ƽ��
Property Get IsOtherDept() As Boolean
    IsOtherDept = mblnIsOtherDept
End Property

Property Let IsOtherDept(ByVal value As Boolean)
    mblnIsOtherDept = value
End Property


'�Զ�����
Property Get IsAutoLine() As Boolean
    IsAutoLine = mblnIsAutoLine
End Property

Property Let IsAutoLine(ByVal value As Boolean)
    mblnIsAutoLine = value
End Property


'��Ȩ����IDs
Property Get GrantDeptIds() As String
    GrantDeptIds = mstrGrantDeptIds
End Property

Property Let GrantDeptIds(ByVal value As String)
    mstrGrantDeptIds = value
End Property

'ѡ����
Property Get SelRow() As Long
    SelRow = vsfStudy.Row
End Property

'ѡ��ҽ��ID
Property Get SelAdviceId() As Long
    Dim intCol As Integer
    
    SelAdviceId = 0
    
    intCol = vsfStudy.ColIndex("ҽ��ID")
    If intCol = -1 Then Exit Property
    
    SelAdviceId = Val(vsfStudy.TextMatrix(vsfStudy.Row, intCol))
End Property


Property Get SelMoved() As Boolean
    Dim intCol As Integer
    
    SelMoved = False
    
    intCol = vsfStudy.ColIndex("ת��״̬")
    If intCol = -1 Then Exit Property
    
    SelMoved = IIf(Val(vsfStudy.TextMatrix(vsfStudy.Row, intCol)) <> 0, True, False)
End Property

'ѡ��ı����ı�
Property Get SelReportText()
    SelReportText = ""
    If rtxtReport.Visible = False Then Exit Property
    
    If rtxtReport.SelLength > 0 Then
        SelReportText = rtxtReport.SelText
    Else
        SelReportText = rtxtReport.Text
    End If
End Property


Public Function IsSelected(Optional ByVal lngIndex As Long = 0) As Boolean
    Dim i As Long
    
    IsSelected = False
    If lngIndex = 0 Then
        For i = 1 To dcmReportImg.Images.Count
            If dcmReportImg.Images(i).BorderColour <> IMG_BACK_BORDER_COLOR Then
                IsSelected = True
                Exit Function
            End If
        Next
    Else
        IsSelected = IIf(dcmReportImg.Images(lngIndex).BorderColour <> IMG_BACK_BORDER_COLOR, True, False)
    End If
End Function


Public Function GetSelects() As Long()
'��ȡѡ�е�ͼ������
'������1��ʼ
    Dim i As Long
    Dim lngBound As Long
    Dim arySelIndex() As Long
    
    ReDim arySelIndex(0)
    
    If dcmReportImg.Visible = False Then Exit Function
    
    For i = 1 To dcmReportImg.Images.Count
        
        If IsSelected(i) Then
            '����Ƿ�͸����ɫ,˵���Ǳ�ѡ�е�ͼ��
            lngBound = UBound(arySelIndex) + 1
            ReDim Preserve arySelIndex(lngBound)
            
            arySelIndex(lngBound) = i
        End If
    Next
    
    GetSelects = arySelIndex
End Function
 
Public Function GetImage(ByVal lngIndex As Long) As DicomImage
'��ȡͼ��
    Dim objSelImg As DicomImage
    
    Set GetImage = Nothing
    
    If lngIndex <= 0 Or lngIndex > dcmReportImg.Images.Count Then Exit Function
    
    Set objSelImg = dcmReportImg.Images(lngIndex)
    
    Set GetImage = objSelImg.SubImage(0, 0, objSelImg.SizeX, objSelImg.SizeY, 1, 1)
End Function


Public Sub SetDateRange(ByVal strDataRange As String)
    mstrDateRange = strDataRange
End Sub




Public Sub Init(ByVal lngModuleNo As Long, ByVal lngDeptId As Long, ByVal strPrivs As String, _
    Optional ByVal blnIsForce As Boolean = False)
On Error GoTo errhandle
    If mblnIsInit And blnIsForce = False Then Exit Sub
     
    
    mlngCurModule = lngModuleNo
    mlngCurDeptId = lngDeptId
    mstrPrivs = strPrivs

    mstrDescriptionName = nvl(GetDeptPara(mlngCurDeptId, "�����������", "�������"))
    mstrAdviseName = nvl(GetDeptPara(mlngCurDeptId, "��������", "��Ͻ���"))
    mstrOpinionName = nvl(GetDeptPara(mlngCurDeptId, "����������", "������"))
    mstrPatholMaterialInfo = ""
    
    If mlngCurModule = G_LNG_PATHSTATION_MODULE Then
        mstrPatholMaterialInfo = zlDatabase.GetPara("ȡ����������", glngSys, mlngCurModule, "1,1,1,1,1,1,1,1,1,1")
    End If
    
    If mblnIsInit = False Then
        Call LoadControlFace
    
        Call SetFontSize(gbytFontSize)
    End If
    
    mblnIsInit = True
Exit Sub
errhandle:
    mblnIsInit = False
End Sub
 

Private Function loadPatholReportList(ByVal lngAdviceId As Long) As Integer
'����ҽ��ID ���ز�����̱������ݵ���ʷ�����б���
'����  0 �쳣   ����ֵ: ��һ���������
'��������lvHistoryList.ListItems.Add��ӵĹؼ��ַ�Ϊprocess�����̱���  describe���޼�����
    Dim objItem As ListItem
    Dim intCount As Integer '�Ѿ��ù������
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    On Error GoTo errH
    
    loadPatholReportList = 0
    
    'ע��������ر�û�в����������ת��
    
    '���ؾ޼������б���
    strSQL = "select  a.ҽ��ID, a.����ҽ��ID as �ؼ�ID, '---' as ����, '---' as ����, '' as Ӱ�����, '�޼�����' as ����, b.ȡ��ʱ�� as ���ʱ��, " & _
                    " '' as ҽ������, '' as �������, '' as �������, 6 as ִ�й���, 'describe' as �������, 0 as ת��״̬ " & _
                  "from ��������Ϣ a,����ȡ����Ϣ b " & _
                  "where a.����ҽ��id=b.����ҽ��id " & _
                  "and b.���= (select min(c.���) from ����ȡ����Ϣ c where c.����ҽ��id=a.����ҽ��id and a.ҽ��id=[1]) " & _
                  "and a.ҽ��id=[1]"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "���ؾ޼������б���", lngAdviceId)
    
    Call LoadList(rsTemp)
    
    '���ع��̱����б���
    strSQL = "select a.ҽ��ID, b.�걾����,b.Id as �ؼ�ID, '---' as ����, '---' as ����, '' as Ӱ�����, b.�������� as ����,b.�������� as ���ʱ��, " & _
                    " ':' || b.�걾���� as ҽ������, '' as �������, '' as �������, 6 as ִ�й���, 'process' as �������, 0 as ת��״̬  " & _
                  "from ��������Ϣ a ,������̱��� b " & _
                  "where a.����ҽ��id=b.����ҽ��id and a.ҽ��id=[1] " & _
                  "order by b.�������� "
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "���ع��̱����б���", lngAdviceId)
    
    Call LoadList(rsTemp)
    
    loadPatholReportList = vsfStudy.Rows - 1
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
End Function

Private Function getReportType(ByVal strType As String) As String
'��þ��屨������  ���������ݿ��е�����
    getReportType = ""
    
    Select Case strType
        Case "0"
            getReportType = "��������"
        Case "1"
            getReportType = "���߱���"
        Case "2"
            getReportType = "���ӱ���"
        Case "3"
            getReportType = "��Ⱦ����"
        Case Else
            getReportType = strType
    End Select
End Function


Private Sub InitPatientInfo()
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
 
    
    strSQL = "select a.����ID,a.��ҳID, a.������Դ,a.�Ա�,a.Ӥ��,b.����id, 0 as ת�� from ����ҽ����¼ a, Ӱ�����¼ b Where a.id=b.ҽ��id(+) and a.id=[1] " & _
        "Union All " & _
        "select a.����ID, a.��ҳID, a.������Դ,a.�Ա�,a.Ӥ��,b.����id, 1 as ת�� from H����ҽ����¼ a, HӰ�����¼ b Where a.id=b.ҽ��id(+) and a.id=[1] "
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "ҽ����Ϣ��ѯ", mlngAdviceId)
    If rsTemp.RecordCount <= 0 Then Exit Sub
    
    mlngPatientId = Val(nvl(rsTemp!����ID))
    mlngPatientFrom = Val(nvl(rsTemp!������Դ))
    mlngBabyNum = Val(nvl(rsTemp!Ӥ��))
    mlngLinkID = Val(nvl(rsTemp!����ID))
    mblnAdviceMoved = IIf(Val(nvl(rsTemp!ת��)) = 1, True, False)
End Sub

Private Sub LoadNormalReportList(ByVal lngAdviceId As Long, _
    Optional ByVal dtBegin As Date = 0, _
    Optional ByVal dtEnd As Date = 0)
    
    Dim strSQL As String
    Dim strTime As String
    Dim strTemp As String
    Dim rsTemp As ADODB.Recordset
    
    If mlngCurModule <> G_LNG_PATHOLSYS_NUM Then
        strSQL = "Select A.ID ҽ��ID,A.ID as �ؼ�ID,A.����ʱ��  ����ʱ��,A.ҽ������, B.ִ�й���, B.�������,C.����, " & _
               " C.Ӱ�����,C.�������,C.����,C.�������� ���ʱ��,E.����,E.�걾��λ,'' as �������,0 as ת��״̬ " & _
               " From ����ҽ����¼ A,����ҽ������ B,Ӱ�����¼ C,������Ϣ D,������ĿĿ¼ E" & _
               " Where A.����id = [1] And A.���id Is Null And B.ҽ��ID=A.ID and a.����id = d.����id  " & _
               " AND A.ID=C.ҽ��ID AND A.������ĿID = E.ID AND b.ִ�й��� >= 2 "
    Else
        strSQL = "Select A.ID ҽ��ID,A.ID as �ؼ�ID,A.����ʱ��  ����ʱ��,A.ҽ������, B.ִ�й���, B.�������,C.����� ����," & _
               " F.Ӱ�����,F.�������,F.����,C.����ʱ�� ���ʱ��,E.����,E.�걾��λ,'' as �������,0 as ת��״̬ " & _
               " From ����ҽ����¼ A,����ҽ������ B,Ӱ�����¼ F,��������Ϣ C,������Ϣ D,������ĿĿ¼ E" & _
               " Where A.����id = [1] And A.���id Is Null And B.ҽ��ID=A.ID and a.����id = d.����id " & _
               " AND A.ID=C.ҽ��ID(+) AND A.������ĿID = E.ID and a.id=F.ҽ��ID AND b.ִ�й��� >= 2 "
    End If
    
    If dtBegin <> 0 And dtEnd <> 0 Then
        strSQL = strSQL & " AND B.����ʱ�� between [6] and [7]"
    End If
    
    '���μ��
    If mblnIsThisTime And mlngPatientFrom = 2 Then
        strSQL = strSQL & " And (A.������Դ=2 And A.��ҳID=D.��ҳID)"
    End If
    
    '���Ƽ��
    If mblnAllDepts = False Then
        If Not mblnIsOtherDept Then
            strSQL = strSQL & " And A.ִ�п���id+0 =[2] "
        Else
            strSQL = strSQL & " And  (A.ִ�п���id+0 <>[2] and B.ִ�й��� >= 5 or A.ִ�п���id+0 =[2]) "
        End If
    Else
        strSQL = strSQL & " And (Instr( [3],',' || A.ִ�п���id || ',' ) >0)"
    End If
    
    'Ӥ��
    strSQL = strSQL & " And NVL(A.Ӥ��,0) = [8]"
    
    '���ù������ˣ��Ų�ѯ����ID
    If mlngLinkID <> 0 Then
        If mlngCurModule <> G_LNG_PATHOLSYS_NUM Then
            strSQL = strSQL & " union select A.ID ҽ��ID,A.ID as �ؼ�ID,A.����ʱ��  ����ʱ��,A.ҽ������, B.ִ�й���, " & _
                " B.�������,C.����, C.Ӱ�����,C.�������,C.����,C.�������� ���ʱ��,E.����,E.�걾��λ,'' as �������, 0 as ת��״̬ " & _
                " From ����ҽ����¼ A ,����ҽ������ B,Ӱ�����¼ C,������Ϣ D,������ĿĿ¼ E" & _
                " Where B.ҽ��ID=A.ID AND A.ID=C.ҽ��ID and a.����id = d.����id AND A.������ĿID = E.ID AND b.ִ�й��� >= 2 AND A.id in (Select ҽ��ID from Ӱ�����¼ Where ����ID =[4]) "
        Else
            strSQL = strSQL & " union select A.ID ҽ��ID,A.ID as �ؼ�ID,A.����ʱ��  ����ʱ��,A.ҽ������, B.ִ�й���, " & _
                " B.�������,C.����� ����,F.Ӱ�����,F.�������,F.����,C.����ʱ�� ���ʱ��, E.����,E.�걾��λ,'' as �������, 0 as ת��״̬ " & _
                " From ����ҽ����¼ A,����ҽ������ B,Ӱ�����¼ F,��������Ϣ C,������Ϣ D,������ĿĿ¼ E" & _
                " Where A.id in (Select ҽ��ID from Ӱ�����¼ Where ����ID =[4]) And B.ҽ��ID=A.ID and a.id=C.ҽ��ID(+) and a.����id = d.����id AND A.������ĿID = E.ID and a.id=F.ҽ��ID and b.ִ�й��� >= 2 "
        End If
        
        If dtBegin <> 0 And dtEnd <> 0 Then
            strSQL = strSQL & " AND B.����ʱ�� between [6] and [7]"
        End If
        
        '���μ��
        If mblnIsThisTime And mlngPatientFrom = 2 Then
            strSQL = strSQL & " And (A.������Դ=2 And A.��ҳID=D.��ҳID)"
        End If
        
'        '���Ƽ��
'        If chkOtherDeptReport.Value <> 1 Then
'            strSql = strSql & " And c.ִ�п���id+0 in(select  ����id  from ������Ա where ��Աid = [5] union all select to_Number([2]) from dual) "
'        End If
        '���Ƽ��
        If mblnAllDepts = False Then
            If Not mblnIsOtherDept Then
                strSQL = strSQL & " And A.ִ�п���id+0 =[2] "
            Else
                strSQL = strSQL & " And  (A.ִ�п���id+0 <>[2] and B.ִ�й��� >= 5 or A.ִ�п���id+0 =[2]) "
            End If
        Else
            strSQL = strSQL & " And (Instr( [3],',' || A.ִ�п���id || ',' ) >0)"
        End If
        
        strSQL = strSQL & " And NVL(A.Ӥ��,0) = [8]"
    End If
    
    If mblnHistoryMoved Then
        strTemp = Replace(strSQL, "0 as ת��״̬", "1 as ת��״̬")
        strTemp = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
        strTemp = Replace(strTemp, "����ҽ������", "H����ҽ������")
        strTemp = Replace(strTemp, "Ӱ�����¼", "HӰ�����¼")
        strTemp = Replace(strTemp, "���˼����Ϣ", "H���˼����Ϣ")
        strSQL = strSQL & vbNewLine & " Union ALL " & vbNewLine & strTemp
    End If
    
    strSQL = "Select * From (" & vbNewLine & strSQL & vbNewLine & ") Order By ����ʱ�� Asc"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "", mlngPatientId, _
            mlngCurDeptId, "," & mstrGrantDeptIds & ",", mlngLinkID, UserInfo.ID, dtBegin, dtEnd, mlngBabyNum)
    
    If rsTemp.RecordCount > 0 Then

        rsTemp.Filter = "ҽ��id <> " & lngAdviceId
        
        Call LoadList(rsTemp)
        
        If mblnIsAutoLine Then
            vsfStudy.WordWrap = True
            vsfStudy.AutoSize 1, vsfStudy.Cols - 1
        Else
            vsfStudy.WordWrap = False
            vsfStudy.AutoSize 1, vsfStudy.Cols - 1
        End If
    
    End If
End Sub

Private Sub LoadList(rsTemp As ADODB.Recordset)
    Dim lngAdviceIdIndex As Long    'ҽ��ID������
    Dim lngSeriesNoIndex As Long    '���������
    Dim lngStudyNoIndex As Long     '����������
    Dim lngAgeIndex As Long         '����������
    Dim lngKindIndex As Long        '���������
    Dim lngItemIndex As Long        '��Ŀ������
    Dim lngProcedureIndex As Long   '��ǰ����������
    Dim lngMasculineIndex As Long   '������������
    Dim lngFollowUpIndex As Long    '���������
    Dim lngAdviceContextIndex As Long 'ҽ������
    Dim lngCheckPointIndex As Long  '��鲿λ
    Dim lngLoadTypeIndex As Long
    Dim lngKyeIdIndex As Long
    Dim lngMovedStateIndex As Long
    Dim lngRow As Long
    
    If rsTemp.RecordCount <= 0 Then Exit Sub
    
    With vsfStudy
        lngAdviceIdIndex = .ColIndex("ҽ��ID")
        lngSeriesNoIndex = .ColIndex("���")
        lngStudyNoIndex = .ColIndex("����")
        lngAgeIndex = .ColIndex("����")
        lngKindIndex = .ColIndex("���")
        lngItemIndex = .ColIndex("��Ŀ")
        lngProcedureIndex = .ColIndex("��ǰ����")
        lngMasculineIndex = .ColIndex("������")
        lngFollowUpIndex = .ColIndex("�������")
        lngAdviceContextIndex = .ColIndex("ҽ������")
        lngCheckPointIndex = .ColIndex("��λ")
        lngLoadTypeIndex = .ColIndex("�������")
        lngKyeIdIndex = .ColIndex("�ؼ�ID")
        lngMovedStateIndex = .ColIndex("ת��״̬")
        
        If mlngCurModule = G_LNG_PATHOLSYS_NUM Then .ColHidden(lngKindIndex) = True
        
        Do While Not rsTemp.EOF
            .Rows = .Rows + 1
'            iCount = .Rows
            lngRow = .Rows - 1
            
            .TextMatrix(lngRow, lngAdviceIdIndex) = Val(nvl(rsTemp!ҽ��ID))
            .TextMatrix(lngRow, lngKyeIdIndex) = Val(nvl(rsTemp!�ؼ�ID))
            .TextMatrix(lngRow, lngLoadTypeIndex) = nvl(rsTemp!�������)
            .TextMatrix(lngRow, lngMovedStateIndex) = nvl(rsTemp!ת��״̬)

            .TextMatrix(lngRow, lngSeriesNoIndex) = lngRow 'iCount
            .TextMatrix(lngRow, lngStudyNoIndex) = nvl(rsTemp!����)
            .TextMatrix(lngRow, lngAgeIndex) = nvl(rsTemp!����)
            
            If mlngCurModule <> G_LNG_PATHOLSYS_NUM Then
                .TextMatrix(lngRow, lngKindIndex) = nvl(rsTemp!Ӱ�����)
                .TextMatrix(lngRow, lngItemIndex) = nvl(rsTemp!����)
            Else
                .TextMatrix(lngRow, lngItemIndex) = getReportType(nvl(rsTemp!����))
            End If
            
            
            
            .TextMatrix(lngRow, lngProcedureIndex) = Decode(Val(nvl(rsTemp!ִ�й���, 0)), -1, "�Ѳ���", 0, "�ѵǼ�", 1, "�ѵǼ�", _
                                                2, "�ѱ���", 3, "�Ѽ��", 4, "�ѱ���", 5, "�����", "�����")
            .Cell(flexcpData, lngRow, lngProcedureIndex) = Val(nvl(rsTemp!ִ�й���, 0))
            
            .TextMatrix(lngRow, lngMasculineIndex) = IIf(Val(nvl(rsTemp!�������)) = 1, "��", "")
            
            .TextMatrix(lngRow, lngFollowUpIndex) = nvl(rsTemp!�������)
            .TextMatrix(lngRow, .ColIndex("���ʱ��")) = Format(rsTemp!���ʱ��, "yyyy-MM-dd hh:mm")
            
            
            If UBound(Split(nvl(rsTemp!ҽ������), ":")) > 0 Then
                .TextMatrix(lngRow, lngAdviceContextIndex) = Split(nvl(rsTemp!ҽ������), ":")(0)
                .TextMatrix(lngRow, lngCheckPointIndex) = Split(nvl(rsTemp!ҽ������), ":")(1)
            Else
                .TextMatrix(lngRow, lngAdviceContextIndex) = nvl(rsTemp!ҽ������)
                .TextMatrix(lngRow, lngCheckPointIndex) = ""
            End If
            
            rsTemp.MoveNext
'                   If .Rows > 1 Then .Row = 1
        Loop
    End With
        
End Sub

Private Sub GetDateRange(ByRef dtBegin As Date, ByRef dtEnd As Date)
    Dim blnNoTime As Boolean
    
    blnNoTime = False
    
    '��ȡʱ�䷶Χ
    Select Case mstrDateRange
        Case "һ����"
            dtBegin = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd 00:00:00")) - 30
            dtEnd = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd 23:59:59"))
        Case "������"
            dtBegin = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd 00:00:00")) - 60
            dtEnd = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd 23:59:59"))
        Case "������"
            dtBegin = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd 00:00:00")) - 90
            dtEnd = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd 23:59:59"))
        Case "����"
            dtBegin = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd 00:00:00")) - 180
            dtEnd = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd 23:59:59"))
        Case "һ��"
            dtBegin = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd 00:00:00")) - 365
            dtEnd = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd 23:59:59"))
        Case "����"
            dtBegin = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd 00:00:00")) - 730
            dtEnd = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd 23:59:59"))
        Case "����"
            dtBegin = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd 00:00:00")) - 1095
            dtEnd = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd 23:59:59"))
        Case "����"
            blnNoTime = True
        Case "�Զ���"
            dtBegin = mdtBegin
            dtEnd = mdtEnd
    End Select
    
    '����ʱ��ʱ����dtBegin + 1��1899/12/31��ȥ�ж��Ƿ�ת�棬������dtBegin��ֱ��blnMoved = true
    If blnNoTime Then
        mblnHistoryMoved = MovedByDate(dtBegin + 1)
    Else
        mblnHistoryMoved = MovedByDate(dtBegin)
    End If
End Sub


Private Sub ResetFace()
    vsfStudy.Rows = 1
    rtxtReport.Text = ""
    rtxtReport.Visible = True
    dcmReportImg.Images.Clear
    dcmReportImg.Visible = False
    chkReportType.value = Unchecked
    
    mftpConTag.Ip = ""
End Sub


Private Function GetRootHwnd() As Long
    GetRootHwnd = GetAncestor(hwnd, GA_ROOT)
End Function

Public Function Refresh(ByVal lngAdviceId As Long, Optional ByVal blnForce As Boolean) As Boolean
'ˢ����ʷ����б�
    Dim dtBegin As Date
    Dim dtEnd As Date
    
    On Error GoTo errhandle

    If lngAdviceId <= 0 Then
        vsfStudy.Rows = 1
        chkReportType.Enabled = False
        cmdWrite.Enabled = False
        
        Exit Function
    Else
        chkReportType.Enabled = True
'        cmdWrite.Enabled = True
    End If
    
    If mlngAdviceId = lngAdviceId And Not blnForce Then Exit Function
    
    Call ResetFace
  
    mlngAdviceId = lngAdviceId
    
    Call InitPatientInfo
     
    Call GetDateRange(dtBegin, dtEnd)
    
    If mlngCurModule = G_LNG_PATHSTATION_MODULE Then
        Call loadPatholReportList(lngAdviceId)
    End If
    
    Call LoadNormalReportList(lngAdviceId, dtBegin, dtEnd)
    

'    If mTPListCfg.strList <> "" Then
'        Call DoLoadListCfg(mTPListCfg.strList)
'    End If
    
    If mTPListCfg.strSortPro <> "" Then
        Call DoLoadListSort(mTPListCfg.strSortPro)
    End If
    
    chkReportType.Enabled = vsfStudy.Rows > 1
    cmdWrite.Enabled = vsfStudy.Row > 1
    cmdWrite.Visible = mblnAllowWrite
    
    Refresh = True
    Exit Function
errhandle:
    MsgboxH GetRootHwnd, err.Description, vbOKOnly, "��ʾ"
    err.Clear
End Function

Public Sub ShowDateConfig()
    Dim objSetTime As New frmSetTime
    
    Call objSetTime.ShowSetTime(mdtBegin, mdtEnd, Me)
End Sub


Private Sub LoadControlFace()
On Error GoTo errhandle
    Dim strValue As String
    Dim objControl As CommandBarControl
           
 
    strValue = GetSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\" & App.ProductName & "\ReportHistory", "�б�����", "")
        
    If InStr(strValue, ";") > 0 Then
        mTPListCfg.strColPros = Split(strValue, ";")(1)
        mTPListCfg.strSortPro = Split(strValue, ";")(0)
    Else
        mTPListCfg.strColPros = M_STR_COLNAME
        mTPListCfg.strSortPro = ""
    End If
    
    '�ж��������Ƿ�һ�£������һ�£���ʹ��Ĭ��������...

    Call GridInit(mTPListCfg.strColPros)
   
    
    mdtBegin = CDate(Format(zlDatabase.Currentdate - 365, "yyyy-mm-dd 00:00:00"))
    mdtEnd = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd 23:59:59"))
    
    
    Exit Sub
errhandle:
    MsgboxH GetRootHwnd, err.Description, vbExclamation, "��ʾ"
    err.Clear
End Sub

Private Sub GridInit(strColName As String)
On Error GoTo errH
    '��ʼ�������б�
    Dim i As Integer
    Dim lngCount As Long
    Dim arrData() As String
    
    Dim strColPros As String
    Dim aryColPro() As String
    
    arrData = Split(strColName, "|")
    lngCount = UBound(arrData) + 1

    With vsfStudy
    
        .Cols = lngCount
        .FixedRows = 1
        .FixedCols = 0
        .RowHeightMin = 320
'        .Cell(flexcpAlignment, 0, 0, 0, lngCount - 1) = flexAlignCenterCenter

        '���һ���Զ�������б�
        .AllowUserResizing = flexResizeColumns
        .ExtendLastCol = True
        .AutoResize = True
        .ExplorerBar = 7 '������ͷ�϶�������
        .AutoSizeMode = flexAutoSizeRowHeight
    
        .WordWrap = True
        .AutoSizeMouse = True
        .SelectionMode = flexSelectionByRow
        .ScrollTrack = True
        
        For i = 0 To lngCount - 1
            strColPros = arrData(i) & ",,,,,"
            aryColPro = Split(strColPros, ",")
            
            .TextMatrix(0, i) = aryColPro(0)
            .ColKey(i) = aryColPro(0)
            
            If Val(aryColPro(1)) > 0 Then
                .ColWidth(i) = Val(aryColPro(1))
            End If
            
            If CBool(Val(aryColPro(2))) = True Then
                .ColHidden(i) = True
            Else
                .ColHidden(i) = False
            End If
        Next
        
        .Rows = 1
        If .Rows > 1 Then .RowSel = 1
        
        .ColHidden(.ColIndex("ҽ��ID")) = True '����ҽ��ID
        .ColHidden(.ColIndex("�������")) = True '���ؼ������
        .ColHidden(.ColIndex("�ؼ�ID")) = True '���عؼ�ID
        .ColHidden(.ColIndex("ת��״̬")) = True '���عؼ�ID
    End With
    Exit Sub
errH:
    MsgboxH GetRootHwnd, err.Description, vbExclamation, "��ʾ"
End Sub

Public Sub SetFontSize(ByVal bytFontSize As Byte)
    Dim CtlFont As StdFont
    
    Set CtlFont = New StdFont
    CtlFont.Size = bytFontSize
    
    UserControl.FontSize = bytFontSize

    Call SetColWithd(bytFontSize)
    
    vsfStudy.FontSize = bytFontSize
     
End Sub

Private Sub SetColWithd(ByVal bytSize As Long)
    With vsfStudy
        Select Case bytSize
            Case 9
                .ColWidth(.ColIndex("���")) = 500
                .ColWidth(.ColIndex("��λ")) = 1000
                .ColWidth(.ColIndex("��ǰ����")) = 900
                .ColWidth(.ColIndex("����")) = 700
                .ColWidth(.ColIndex("���")) = 700
                .ColWidth(.ColIndex("����")) = 700
                .ColWidth(.ColIndex("���ʱ��")) = 1600
                .ColWidth(.ColIndex("��Ŀ")) = 1200
                .ColWidth(.ColIndex("������")) = 800
            Case 12
                .ColWidth(.ColIndex("���")) = 600
                .ColWidth(.ColIndex("��λ")) = 1250
                .ColWidth(.ColIndex("��ǰ����")) = 1100
                .ColWidth(.ColIndex("����")) = 900
                .ColWidth(.ColIndex("���")) = 900
                .ColWidth(.ColIndex("����")) = 900
                .ColWidth(.ColIndex("���ʱ��")) = 2200
                .ColWidth(.ColIndex("��Ŀ")) = 1450
                .ColWidth(.ColIndex("������")) = 1000
            Case 15
                .ColWidth(.ColIndex("���")) = 700
                .ColWidth(.ColIndex("��λ")) = 1500
                .ColWidth(.ColIndex("��ǰ����")) = 1300
                .ColWidth(.ColIndex("����")) = 1100
                .ColWidth(.ColIndex("���")) = 1100
                .ColWidth(.ColIndex("����")) = 1100
                .ColWidth(.ColIndex("���ʱ��")) = 2800
                .ColWidth(.ColIndex("��Ŀ")) = 1700
                .ColWidth(.ColIndex("������")) = 1200
        End Select
    End With
End Sub

 

Private Sub chkLinkView_Click()
On Error GoTo errhandle
    chkLinkView.Caption = IIf(chkLinkView.value <> 0, "�� �����鿴", "�� �����鿴")
    chkLinkView.ForeColor = IIf(chkLinkView.value <> 0, vbBlack, &H808080)
    
    If chkLinkView.value = 0 Then
        RaiseEvent OnLinkView(0, False, False)
        Exit Sub
    End If
    
    Call DoLinkView(False)
    
Exit Sub
errhandle:
    MsgboxH hwnd, err.Description, vbOKOnly, "��ʾ"
End Sub

Public Sub CloseLinkViewer()
On Error GoTo errhandle
    RaiseEvent OnLinkView(0, False, False)
Exit Sub
errhandle:
    MsgboxH hwnd, err.Description, vbOKOnly, "��ʾ"
End Sub

Private Sub DoLinkView(ByVal blnIsDBClick As Boolean)
    Dim intCol As Long
    Dim lngAdviceId As Long
    Dim blnMoved As Boolean
On Error GoTo errhandle
    
    If vsfStudy.Rows <= 1 Then Exit Sub
    
    intCol = vsfStudy.ColIndex("ҽ��ID")
    If intCol = -1 Then Exit Sub
    
    lngAdviceId = Val(vsfStudy.TextMatrix(vsfStudy.Row, intCol))
    blnMoved = IIf(Val(vsfStudy.TextMatrix(vsfStudy.Row, vsfStudy.ColIndex("ת��״̬"))) = 0, False, True)
    
    RaiseEvent OnLinkView(lngAdviceId, blnMoved, blnIsDBClick)
    
Exit Sub
errhandle:
    MsgboxH hwnd, err.Description, vbOKOnly, "��ʾ"
End Sub

Private Sub chkReportType_Click()
    
On Error GoTo errhandle
    cmdWrite.Enabled = False
    
    If chkReportType.value = 1 Then
 
        '���뱨��ͼ��
        Call ViewReportImage
        
        chkReportType.ToolTipText = "����ͼ��"
    Else
        '���뱨������
        Call ViewReportContext
        
        chkReportType.ToolTipText = "�ı�����"
    End If
Exit Sub
errhandle:
    MsgboxH hwnd, err.Description, vbOKOnly, "��ʾ"
End Sub


Public Sub ViewReportImage()
'�鿴����ͼ
    Dim strLoadType As String
    Dim blnMoved As Boolean
    
    If SelAdviceId <= 0 Then
        MsgboxH hwnd, "��ѡ����Ҫ�鿴����ʷ��¼��", vbOKOnly, "��ʾ"
        Exit Sub
    End If
    
    blnMoved = IIf(Val(vsfStudy.TextMatrix(vsfStudy.Row, vsfStudy.ColIndex("ת��״̬"))) = 0, False, True)
    strLoadType = vsfStudy.TextMatrix(vsfStudy.Row, vsfStudy.ColIndex("�������"))
    
    If Len(strLoadType) > 0 Then
        MsgboxH hwnd, "�ñ�������²�����ʾ����ͼ��", vbOKOnly, "��ʾ"
        Exit Sub
    End If
         
    If SelAdviceId > 0 And SelAdviceId <> dcmReportImg.tag Then

        Call LoadReportImage(SelAdviceId, blnMoved)
        dcmReportImg.tag = SelAdviceId
    End If
    
    dcmReportImg.Visible = True
    rtxtReport.Visible = False
    
    chkReportType.value = Checked
End Sub

Public Sub ViewReportContext()
'���ر����ı�
    Dim strLoadType As String
    Dim blnMoved As Boolean
    
    If SelAdviceId <= 0 Then
        MsgboxH hwnd, "��ѡ����Ҫ�鿴����ʷ��¼��", vbOKOnly, "��ʾ"
        Exit Sub
    End If
    
    blnMoved = IIf(Val(vsfStudy.TextMatrix(vsfStudy.Row, vsfStudy.ColIndex("ת��״̬"))) = 0, False, True)
    strLoadType = vsfStudy.TextMatrix(vsfStudy.Row, vsfStudy.ColIndex("�������"))
    
    If SelAdviceId > 0 And SelAdviceId <> rtxtReport.tag Then
        Call LoadReport(SelAdviceId, _
            strLoadType, _
            vsfStudy.TextMatrix(vsfStudy.Row, vsfStudy.ColIndex("�ؼ�ID")), _
            blnMoved)
            
        rtxtReport.tag = SelAdviceId
    End If
    
    rtxtReport.Visible = True
    dcmReportImg.Visible = False
    
    chkReportType.value = Unchecked
End Sub

Private Sub LoadReportImage(ByVal lngAdviceId As Long, ByVal blnMoveState As Boolean)
'����ͼ��ѯ...
    Dim strSQL As String
    Dim strPicSql As String
    Dim rsData As ADODB.Recordset
    Dim lngFileId As Long
    
    strSQL = "select ����ID from ����ҽ������ where ҽ��ID=[1] "
    If blnMoveState Then
        strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
    End If
    
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "��ʷҽ������ID��ѯ", lngAdviceId)
    
    dcmReportImg.Images.Clear
    If rsData.RecordCount <= 0 Then Exit Sub
    
    
    lngFileId = Val(nvl(rsData!����Id))
    
    '�ӵ��Ӳ��������в�ѯ����
    strSQL = "Select  Id As ���Id From ���Ӳ�������" & _
                " Where �ļ�id = [1] And �������� = 3 And Substr(��������, Instr(��������, ';', 1, 18) + 1, 1) = '2' " & _
                " Order By �������"
                
    strPicSql = "select ID,�ļ�ID,��ID,��ʼ��,������,��������,�����д� from ���Ӳ������� where  �ļ�ID=[1] and ��ID=[2] and ��������=5 order by ������"
        
    '�Ƿ�ת������
    If blnMoveState Then
        strSQL = Replace(strSQL, "���Ӳ�������", "H���Ӳ�������")
        strPicSql = Replace(strPicSql, "���Ӳ�������", "H���Ӳ�������")
    End If
    
    
    '��ȡ����ͼ��Ϣ****************************************
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "��ѯ����ͼ��", lngFileId)
    
    If rsData.RecordCount > 0 Then
        '��ȡ���ͼ������ͼ
        'ͼ������ѯ
        dcmReportImg.tag = Val(nvl(rsData!���ID))
        
        Set rsData = zlDatabase.OpenSQLRecord(strPicSql, "��ѯ����ͼƬ", lngFileId, Val(nvl(rsData!���ID)))
        If rsData.RecordCount > 0 Then
            
            Call ParshReportImgData(lngAdviceId, rsData, blnMoveState)
        End If
    End If
    
  
End Sub


Private Sub ParshReportImgData(ByVal lngAdviceId As Long, rsData As ADODB.Recordset, ByVal blnMoveState As Boolean)
'��������ͼ������
    Dim aryImgPro() As String
    Dim reportImgTag As TReportImgTag
    Dim result As ftpResult
    Dim blnIsAbort As Boolean
    Dim objDcmImg As DicomImage
    
 
    If rsData Is Nothing Then Exit Sub
    
    rsData.MoveFirst
    blnIsAbort = False
    
    While Not rsData.EOF
        '��һ������˵����0��ͨͼ��1���ͼ��2����ͼ��
        aryImgPro = Split(nvl(rsData!��������) & ";;;;;;;;;;;;;;;;;;;;", ";")
        
        reportImgTag.lngFileId = Val(rsData!�ļ�ID)
        reportImgTag.lngTableId = Val(rsData!��ID)
        reportImgTag.strObjectTag = Val(rsData!������)
        reportImgTag.strPros = nvl(rsData!��������)
        reportImgTag.lngStartVer = Val(rsData!��ʼ��)
        reportImgTag.strKey = Val(rsData!ID)
        reportImgTag.strImgMarks = ""
        
        If Val(aryImgPro(0)) = 2 Then '����ͼ
            reportImgTag.lngImgType = ritReport
            
            If blnIsAbort = False Then
                result = ReadReportImage(lngAdviceId, dcmReportImg.Images, reportImgTag, blnMoveState)
            Else
                '�����滻ͼ��
                Set objDcmImg = dcmReportImg.Images.AddNew
                
                dcmReportImg.Images(dcmReportImg.Images.Count).tag = reportImgTag
                
                Call DrawBorder(objDcmImg, 0)
                Call DrawErrorText(objDcmImg, "�ѱ���ֹ")
                
            End If
            
            Call CalcImgView
            
            If result = frAbort Then
                '��������쳣����ѡ����ֹ���أ����˳�ͼ����ش���
                blnIsAbort = True
            End If
        End If
        
        Call rsData.MoveNext
    Wend
End Sub


Private Sub CalcImgView()
    Dim iCols As Integer, iRows As Integer
    
    If dcmReportImg.Images.Count = 1 Then Exit Sub
    
On Error Resume Next
      
    '����ͼ����ʾ����
    ResizeRegion dcmReportImg.Images.Count, dcmReportImg.Width, dcmReportImg.Height, iRows, iCols

    dcmReportImg.MultiColumns = iCols
    dcmReportImg.MultiRows = iRows
    
    If dcmReportImg.Images.Count > 0 Then
        dcmReportImg.CurrentIndex = 1
    Else
        dcmReportImg.CurrentIndex = 0
    End If
End Sub


Private Function ReadReportImage(ByVal lngAdviceId As Long, _
    objImages As DicomImages, reportImgTag As TReportImgTag, ByVal blnMoveState As Boolean) As ftpResult
'��ȡ����ͼ
    Dim strFile As String
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim objPicMarks As clsPicMarks
    Dim dblMarkZoom As Double
    Dim strError As String
    Dim objDcmImg As DicomImage
    Dim strFileName As String
    Dim blnImgReadState As Boolean
    Dim strReportImgPath As String
    Dim lngImgAdviceId As Long
    
    
    ReadReportImage = frNormal
    blnImgReadState = True
    
    strFileName = GetReportImagePro(reportImgTag.strPros, "PicName")
    lngImgAdviceId = Val(GetReportImagePro(reportImgTag.strPros, "ADVICEID"))
    
    strReportImgPath = GetReportImgPath(lngAdviceId, blnMoveState)
    
    If DirExists(strReportImgPath) = False Then Call MkLocalDir(strReportImgPath)
    
    If Len(strFileName) > 0 Then
        
        strFile = FormatFilePath(strReportImgPath & "\" & strFileName)
        
        '��ftp����ͼ��
        If FileExists(strFile) = False Then
            If lngImgAdviceId = 0 Then lngImgAdviceId = lngAdviceId
            
            ReadReportImage = DownLoadFtpFile(lngImgAdviceId, strFileName, strFile, blnMoveState)
            If ReadReportImage <> frNormal Then
                blnImgReadState = False
            End If
        End If
    Else
        strFile = FormatFilePath(strReportImgPath & "\����ͼ_" & reportImgTag.strKey & ".JPG")
        
        '�����ݿ��ȡͼ��
        If FileExists(strFile) = False Then
            Call Sys.ReadLob(glngSys, 6, reportImgTag.strKey, strFile)
        End If
    End If
    
    If FileExists(strFile) = False Then
        If Len(strError) <= 0 Then strError = "δ�ҵ���Ӧ�ı���ͼ���ļ� [" & strFile & "]"
        blnImgReadState = False
    End If
    
    If blnImgReadState Then
        'ͼ���ȡ�ɹ��Ĵ���
        Set objDcmImg = ReadDicomFile(strFile, strError)
        
        If Not objDcmImg Is Nothing Then
            reportImgTag.strImgFile = strFileName
            
            objDcmImg.tag = reportImgTag
            
            Call objImages.Add(objDcmImg)
            Call DrawBorder(objDcmImg, 0)
        Else
            blnImgReadState = False
        End If
    End If
    
    If blnImgReadState = False Then
        '����ʧ�ܵ�ͼ��
        
        Set objDcmImg = objImages.AddNew
        
        objImages(objImages.Count).tag = reportImgTag
        
        Call DrawBorder(objDcmImg, 0)
        Call DrawErrorText(objDcmImg, strError)
        
        If ReadReportImage = frNormal Then Call MsgboxH(hwnd, "ͼ���ȡʧ�ܡ�" & vbCrLf & strError, vbOKOnly, "��ʾ")
    End If
End Function

Public Function GetLayoutStr() As String
'���ظ�ʽ�ַ���[Key=TESTNAME@picturebox1.width:20;picturebox1.height:30;]
    GetLayoutStr = "[KEY=HISTORY@" & _
                                        GetProFmt("VSFSTUDY.HEIGHT", vsfStudy.Height) & _
                                        GetProFmt("PICCONTEXT.HEIGHT", picContext.Height) & _
                                 "]"
End Function

Public Sub SetLayout(ByVal strLayout As String)
    Dim strPros As String
    Dim strPro As String
    
    If Len(strLayout) <= 0 Then Exit Sub
    
    strPros = GetPros(strLayout, "HISTORY")
    
    strPro = GetProValue(strPros, "VSFSTUDY.HEIGHT")
    If Val(strPro) > 0 Then vsfStudy.Height = Val(strPro)
    
    strPro = GetProValue(strPro, "PICCONTEXT.HEIGHT")
    If Val(strPro) > 0 Then picContext.Height = Val(strPro)
    

End Sub


Private Function DownLoadFtpFile(ByVal lngAdviceId As Long, _
    ByVal strFtpFile As String, ByVal strLocalFile As String, ByVal blnMoveState As Boolean) As ftpResult
'����ftp�ļ�
    DownLoadFtpFile = frNormal
    If Len(mftpConTag.Ip) <= 0 Or Val(mftpConTag.tag) <> lngAdviceId Then
        mftpConTag = GetReportDevice(lngAdviceId, blnMoveState)
        mftpConTag.tag = lngAdviceId
        
        If Len(mftpConTag.Ip) <= 0 Then
            DownLoadFtpFile = frAbort
            Exit Function
        End If
    End If
    
    DownLoadFtpFile = FtpDownload(mftpConTag, strFtpFile, strLocalFile)
End Function


Private Function GetReportDevice(ByVal lngAdviceId As Long, ByVal blnMoveState As Boolean) As TFtpConTag
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    
On Error GoTo errhandle
    strSQL = "select NVl(���ID, ID) as ID from ����ҽ����¼ where ID=[1]"
    
    If blnMoveState Then strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "��ѯ��ҽ��ID", lngAdviceId)
    
    If rsData.RecordCount <= 0 Then
        Call MsgboxH(hwnd, "ҽ������У��ʧ�ܣ�δ�ҵ��������ҽ����Ϣ��", vbOKOnly, "��ʾ")
        Exit Function
    End If
    
    strSQL = " Select Decode(A.��������,Null,'',to_Char(A.��������,'YYYYMMDD')||'/') ||A.���UID||'/' As URL," & _
            " B.�豸�� as �豸��1, B.�豸�� As �豸��1, B.FTP�û��� As User1,B.FTP���� As Pwd1, B.IP��ַ As Host1, " & _
                    " decode(B.FtpĿ¼, null, '/', '/'||B.FtpĿ¼||'/') As Root1,B.����Ŀ¼ as ����Ŀ¼1,B.����Ŀ¼�û��� as ����Ŀ¼�û���1,B.����Ŀ¼���� as ����Ŀ¼����1 " & _
            " From  Ӱ�����¼ A,Ӱ���豸Ŀ¼ B " & _
            " Where A.ҽ��ID=[1] And nvl(A.λ��һ, A.λ�ö�)=B.�豸��(+)  "
            
    If blnMoveState Then strSQL = Replace(strSQL, "Ӱ�����¼", "HӰ�����¼")
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "��ѯ����ͼ��洢", Val(rsData!ID))
            
    If rsData.RecordCount <= 0 Then
        Call MsgboxH(hwnd, "δ�ҵ�����ͼ��Ӧ�Ĵ洢�豸�����������Ƿ���ȷ��", vbOKOnly, "��ʾ")
        Exit Function
    End If
    
    If nvl(rsData!Host1) <> "" Then
        GetReportDevice = FtpTagInstance(rsData!Host1, rsData!User1, rsData!Pwd1, rsData!Root1 & rsData!Url)
    End If
Exit Function
errhandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub SelectedAll()
'ȫѡ
    Dim i As Long
    
    For i = 1 To dcmReportImg.Images.Count
        Call DrawBorder(dcmReportImg.Images(i), ColorConstants.vbRed, True)
    Next i
End Sub

Private Sub cmdWrite_Click()
On Error GoTo errhandle
    Call WriteReport
Exit Sub
errhandle:
    Call MsgboxH(hwnd, err.Description, vbOKOnly, "��ʾ")
End Sub

Public Sub WriteReport()
'д�뱨��
    If vsfStudy.Rows <= 0 Then Exit Sub
    
    RaiseEvent OnSend
End Sub

Private Sub dcmReportImg_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo errhandle
    Dim lngSelectIndex As Long
    Dim i As Long
    
    If dcmReportImg.Images.Count <= 0 Then Exit Sub
    
    Select Case KeyCode
        Case 37     '�������
            lngSelectIndex = mlngSelImgIndex - 1
            If lngSelectIndex <= 0 Then Exit Sub
        Case 38    '�Ϲ���
            lngSelectIndex = mlngSelImgIndex - dcmReportImg.MultiColumns
            If lngSelectIndex <= 0 Then Exit Sub
        Case 39      '�ҹ���
            lngSelectIndex = mlngSelImgIndex + 1
            If lngSelectIndex > dcmReportImg.Images.Count Then Exit Sub
        Case 40      '�¹���
            lngSelectIndex = mlngSelImgIndex + dcmReportImg.MultiColumns
            If lngSelectIndex > dcmReportImg.Images.Count Then Exit Sub
        Case 65
            If Shift = 2 Then
                Call SelectedAll  '����ȫѡ
                lngSelectIndex = 0
                Exit Sub
            End If
            
        Case Else
            Exit Sub
    End Select
    
    For i = 1 To dcmReportImg.Images.Count
        Call DrawBorder(dcmReportImg.Images(i), 0)
    Next
        
    If lngSelectIndex > 0 Then
        Call DrawBorder(dcmReportImg.Images(lngSelectIndex), ColorConstants.vbRed, True)
    End If
    
    mlngSelImgIndex = lngSelectIndex
     
Exit Sub
errhandle:
    Call MsgboxH(hwnd, err.Description, vbOKOnly, "��ʾ")
End Sub

Private Sub dcmReportImg_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
On Error GoTo errhandle
    Dim i As Long
    
    If Button = 2 Then
        '����Ҽ�
    Else
        mlngSelImgIndex = dcmReportImg.ImageIndex(X, Y)
        
        If mlngSelImgIndex <= 0 Or mlngSelImgIndex > dcmReportImg.Images.Count Then Exit Sub
        
        If Shift <> 2 Then
            For i = 1 To dcmReportImg.Images.Count
                Call DrawBorder(dcmReportImg.Images(i), 0)
            Next
        End If
            
        Call DrawBorder(dcmReportImg.Images(mlngSelImgIndex), ColorConstants.vbRed, True)
    End If
    
    RaiseEvent OnMouseUp(Button, Shift, CSng(X), CSng(Y))
Exit Sub
errhandle:
    MsgboxH hwnd, err.Description, vbOKOnly, "��ʾ"
End Sub

Private Sub picContext_Resize()
On Error GoTo errhandle
    chkReportType.Left = 0
    chkReportType.Top = 0
    
    cmdWrite.Top = 0
    cmdWrite.Left = chkReportType.Left + chkReportType.Width ' picContext.ScaleWidth - cmdWrite.Width
    
    chkLinkView.Left = picContext.Width - chkLinkView.Width
    chkLinkView.Top = 45
    
    rtxtReport.Left = 0
    rtxtReport.Top = chkReportType.Height
    rtxtReport.Width = picContext.ScaleWidth
    rtxtReport.Height = picContext.ScaleHeight - chkReportType.Height
    
    dcmReportImg.Left = 0
    dcmReportImg.Top = chkReportType.Height
    dcmReportImg.Width = picContext.ScaleWidth
    dcmReportImg.Height = picContext.ScaleHeight - chkReportType.Height
Exit Sub
errhandle:

End Sub

Private Sub rtxtReport_SelChange()
On Error GoTo errhandle
    cmdWrite.Enabled = IIf(rtxtReport.SelLength > 0, True, False)
Exit Sub
errhandle:

End Sub

Private Sub UserControl_Initialize()
    mstrDateRange = "һ��"
    mblnAllowWrite = True
End Sub

Private Sub UserControl_Resize()
On Error GoTo errhandle
    picBack.Move 0, 0, ScaleWidth, ScaleHeight
    
    Call ucSplitter1.RePaint(False)
Exit Sub
errhandle:

End Sub
 
Private Sub UserControl_Terminate()
    Call Destory
    
    Set mobjOwner = Nothing
End Sub

'Private Sub UserControl_Show()
'On Error GoTo errHandle
'    If UserControl.Ambient.UserMode Then
'        Call Init(mlngCurModule, mlngCurDeptID)
'    End If
'Exit Sub
'errHandle:
'
'End Sub



Private Sub vsfStudy_AfterMoveColumn(ByVal Col As Long, Position As Long)
On Error GoTo errhandle
    mTPListCfg.strColPros = GetListHeadString
    
    SaveSetting "ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\" & App.ProductName & "\ReportHistory", "�б�����", mTPListCfg.strSortPro & ";" & mTPListCfg.strColPros
Exit Sub
errhandle:

End Sub

Private Sub vsfStudy_AfterSort(ByVal Col As Long, Order As Integer)
    Dim strName As String
    Dim i As Integer
    
On Error GoTo errhandle
    For i = 1 To vsfStudy.Rows - 1
        vsfStudy.TextMatrix(i, vsfStudy.ColIndex("���")) = i
    Next
    
    strName = vsfStudy.TextMatrix(0, Col)
    mTPListCfg.strSortPro = strName & "," & Order
    
    mTPListCfg.strColPros = GetListHeadString
    
    SaveSetting "ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\" & App.ProductName & "\ReportHistory", "�б�����", mTPListCfg.strSortPro & ";" & mTPListCfg.strColPros
Exit Sub
errhandle:

End Sub

Private Sub vsfStudy_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
On Error GoTo errhandle
    mTPListCfg.strColPros = GetListHeadString
    
    SaveSetting "ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\" & App.ProductName & "\ReportHistory", "�б�����", mTPListCfg.strSortPro & ";" & mTPListCfg.strColPros
Exit Sub
errhandle:
End Sub


Private Sub vsfStudy_Click()
On Error GoTo errhandle
     
    RaiseEvent OnClick
    
    Exit Sub
errhandle:
    MsgboxH hwnd, err.Description, vbExclamation, "��ʾ"
    err.Clear
End Sub

Private Sub vsfStudy_DblClick()
On Error GoTo errhandle

    If chkLinkView.Visible Then Call DoLinkView(True)
    
    '��ͨ��˫�����������α�����ϸ��ʾ���ڣ��������μ���¼��ʾ������������ʾ��ҽ����ʾ��������ʾ�ȣ�
    RaiseEvent OnDbClick
    
    Exit Sub
errhandle:
    MsgboxH hwnd, err.Description, vbExclamation, "��ʾ"
    err.Clear
End Sub


Private Sub vsfStudy_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo errH
     
    RaiseEvent OnMouseUp(Button, Shift, X, Y)
    
errH:
End Sub

Public Sub ClearData()
'�������
    vsfStudy.Rows = 1
    
    mlngAdviceId = 0
End Sub


Public Function IsImageEnable(ByVal lngAdvice As Long) As Boolean
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    If lngAdvice <= 0 Then
        IsImageEnable = False
        Exit Function
    End If
    
    strSQL = "select ���UID from Ӱ�����¼ where  ҽ��id = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ѯ���UID", lngAdvice)
    
    If rsTemp.RecordCount <= 0 Then Exit Function
    
    IsImageEnable = Len(nvl(rsTemp!���UID)) > 0
End Function

Public Function IsReportEnable(ByVal lngAdvice As Long) As Boolean
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    If lngAdvice <= 0 Then
        IsReportEnable = False
        Exit Function
    End If
    
    strSQL = "select count(1) ���� from ����ҽ������ where  ҽ��id = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ѯ����", lngAdvice)
    
    If rsTemp.RecordCount <= 0 Then Exit Function
    
    IsReportEnable = Val(nvl(rsTemp!����)) > 0
End Function

Public Sub Destory()
On Error GoTo errhandle

    ucSplitter1.Destory
    
Exit Sub
errhandle:
    Debug.Print "ucReportHistory_Destory Err:" & err.Description
End Sub

Private Sub vsfStudy_SelChange()
On Error GoTo errhandle
    Dim intCol As Integer
    Dim lngAdviceId As Long
    Dim blnMoved As Boolean
    
    If vsfStudy.Rows <= 1 Then Exit Sub
    If vsfStudy.Row <= 0 Then Exit Sub
    
    intCol = vsfStudy.ColIndex("ҽ��ID")
    If intCol = -1 Then Exit Sub
    
    lngAdviceId = Val(vsfStudy.TextMatrix(vsfStudy.Row, intCol))
    mftpConTag.Ip = ""
    
    blnMoved = IIf(Val(vsfStudy.TextMatrix(vsfStudy.Row, vsfStudy.ColIndex("ת��״̬"))) = 0, False, True)
    '��ʾ��������...
    Call LoadReport(lngAdviceId, _
        vsfStudy.TextMatrix(vsfStudy.Row, vsfStudy.ColIndex("�������")), _
        vsfStudy.TextMatrix(vsfStudy.Row, vsfStudy.ColIndex("�ؼ�ID")), _
        blnMoved)
        
    rtxtReport.tag = lngAdviceId
    
    rtxtReport.Visible = True
    dcmReportImg.Visible = False
    dcmReportImg.Images.Clear
    dcmReportImg.tag = 0
    
    chkReportType.value = Unchecked
    
    '�����鿴
    If chkLinkView.Visible And chkLinkView.value <> 0 Then Call DoLinkView(False)
    
Exit Sub
errhandle:
    MsgboxH GetRootHwnd, err.Description, vbOKOnly, "��ʾ"
End Sub


Private Sub LoadProcessReport(ByVal lngKey As Long, ByVal strFormatHead As String, ByVal strFontSize As String, _
    ByVal blnMovedState As Boolean)
'���벡����̱�������
'lngReportId ���̱���ID
'strFormatHead ��ʽͷ
'strFontSize �����С

    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim strFormatContext  As String
    Dim strtext As String
    Dim strTitle As String
    Dim strAttachInfo As String
    
    On Error GoTo errH
    strFormatContext = strFormatHead
    
    
    '��ʾ��������
    strAttachInfo = LoadAttachInfo(mlngAdviceId, strFontSize, blnMovedState)
    strFormatContext = strFormatContext & strAttachInfo
    
    '��ѯ���̱�������
    strSQL = "select �����,������ from ������̱��� where id=[1]"

    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ʷ������̱����ѯ", lngKey)
                
    If rsTemp.RecordCount > 0 Then
    
        If Trim(strAttachInfo) <> "" Then
            strFormatContext = strFormatContext & "\b\cf0\fs" & strFontSize & "==**************************==" & "\par"
        End If
        
        strTitle = "�����" & "��"
        strtext = nvl(rsTemp!�����) & vbCrLf
        strFormatContext = strFormatContext & "\b\cf2\fs24 " & strTitle & " \par\b0\cf0\fs" & strFontSize & " " & Replace(strtext, vbCrLf, " \par\cf0\fs" & strFontSize & " ") & "\par"
        
        strTitle = "������" & "��"
        strtext = nvl(rsTemp!������) & vbCrLf
        strFormatContext = strFormatContext & "\b\cf2\fs24 " & strTitle & " \par\b0\cf0\fs" & strFontSize & " " & Replace(strtext, vbCrLf, " \par\cf0\fs" & strFontSize & " ") & "\par"
            
        strFormatContext = strFormatContext & "}"
        rtxtReport.SelRTF = strFormatContext
        rtxtReport.SelStart = 0
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub LoadDescription(ByVal lngKey As Long, ByVal strFormatHead As String, ByVal strFontSize As String, _
    ByVal blnMovedState As Boolean)
'����޼���������  mstrPatholMaterialInfo �걾����,ȡ��λ��,��״,������,��Ƭ��,��ȡҽʦ,ȡ��ʱ��,����,��ɫ,�걾��
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim strFormatContext  As String
    Dim strtext As String
    Dim strTitle As String
    Dim str�޼����� As String
    Dim blnIsCell As Boolean '�Ŀ��Ƿ���ϸ������ ϸ��������2
    Dim strAttachInfo As String
    
    On Error GoTo errH
    
    strFormatContext = strFormatHead
    
    '��ʾ��������
    strAttachInfo = LoadAttachInfo(mlngAdviceId, strFontSize, blnMovedState)
    strFormatContext = strFormatContext & strAttachInfo
    
    '��ѯ�޼�����
    strSQL = "select  a.�޼�����,a.�������,b.���,b.�걾����, b.��״,b.ȡ��λ��,b.������,b.��ȡҽʦ,b.����,b.��ɫ,b.�걾��,b.ȡ��ʱ��, b.�걾����, c.��Ƭ�� " & _
                      "from ��������Ϣ a ,����ȡ����Ϣ b ,������Ƭ��Ϣ c " & _
                      "where b.�Ŀ�id=c.�Ŀ�id and a.����ҽ��id=c.����ҽ��id and a.����ҽ��id=b.����ҽ��id and a.����ҽ��id=[1] order by b.��� "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ʷ�޼�������ѯ", lngKey)
        
    If rsTemp.RecordCount > 0 Then
        str�޼����� = nvl(rsTemp!�޼�����)
        blnIsCell = (Val(nvl(rsTemp!�������)) = 2)   'ϸ��������2
        
        If Trim(strAttachInfo) <> "" Then
            strFormatContext = strFormatContext & "\b\cf0\fs" & strFontSize & "==**************************==" & "\par"
        End If
    End If
    
    If UBound(Split(mstrPatholMaterialInfo, ",")) <> 9 Then mstrPatholMaterialInfo = "1,1,1,1,1,1,1,1,1,1"
                
    While Not rsTemp.EOF
    
        strTitle = "����" & nvl(rsTemp!���) & "��"
        strtext = ""
        
        If Split(mstrPatholMaterialInfo, ",")(0) = 1 And Trim(nvl(rsTemp!�걾����)) <> "" Then strtext = "�걾���ƣ�" & nvl(rsTemp!�걾����)
        If Split(mstrPatholMaterialInfo, ",")(1) = 1 And Trim(nvl(rsTemp!ȡ��λ��)) <> "" Then strtext = IIf(strtext <> "", strtext & "��" & "ȡ��λ�ã�" & nvl(rsTemp!ȡ��λ��), "ȡ��λ�ã�" & nvl(rsTemp!ȡ��λ��))
        If Split(mstrPatholMaterialInfo, ",")(2) = 1 And Trim(nvl(rsTemp!��״)) <> "" Then strtext = IIf(strtext <> "", strtext & "��" & "��״��" & nvl(rsTemp!��״), "��״��" & nvl(rsTemp!��״))
        
        If blnIsCell Then
            If Split(mstrPatholMaterialInfo, ",")(7) = 1 And Trim(nvl(rsTemp!����)) <> "" Then strtext = IIf(strtext <> "", strtext & "��" & "���ʣ�" & nvl(rsTemp!����), "���ʣ�" & nvl(rsTemp!����))
            If Split(mstrPatholMaterialInfo, ",")(8) = 1 And Trim(nvl(rsTemp!��ɫ)) <> "" Then strtext = IIf(strtext <> "", strtext & "��" & "��ɫ��" & nvl(rsTemp!��ɫ), "��ɫ��" & nvl(rsTemp!��ɫ))
            If Split(mstrPatholMaterialInfo, ",")(9) = 1 And Trim(nvl(rsTemp!�걾��)) <> "" Then strtext = IIf(strtext <> "", strtext & "��" & "�걾����" & nvl(rsTemp!�걾��), "�걾����" & nvl(rsTemp!�걾��))
            If Split(mstrPatholMaterialInfo, ",")(3) = 1 And Trim(nvl(rsTemp!������)) <> "" Then strtext = IIf(strtext <> "", strtext & "��" & "ϸ��������" & nvl(rsTemp!������), "ϸ��������" & nvl(rsTemp!������))
        Else
            If Split(mstrPatholMaterialInfo, ",")(3) = 1 And Trim(nvl(rsTemp!������)) <> "" Then strtext = IIf(strtext <> "", strtext & "��" & "�Ŀ�����" & nvl(rsTemp!������), "�Ŀ�����" & nvl(rsTemp!������))
        End If
        
        If Split(mstrPatholMaterialInfo, ",")(4) = 1 And Trim(nvl(rsTemp!��Ƭ��)) <> "" Then strtext = IIf(strtext <> "", strtext & "��" & "��Ƭ����" & nvl(rsTemp!��Ƭ��), "��Ƭ����" & nvl(rsTemp!��Ƭ��))
        If Split(mstrPatholMaterialInfo, ",")(5) = 1 And Trim(nvl(rsTemp!��ȡҽʦ)) <> "" Then strtext = IIf(strtext <> "", strtext & "��" & "��ȡҽʦ��" & nvl(rsTemp!��ȡҽʦ), "��ȡҽʦ��" & nvl(rsTemp!��ȡҽʦ))
        If Split(mstrPatholMaterialInfo, ",")(6) = 1 And Trim(nvl(rsTemp!ȡ��ʱ��)) <> "" Then strtext = IIf(strtext <> "", strtext & "��" & "ȡ��ʱ�䣺" & nvl(rsTemp!ȡ��ʱ��), "ȡ��ʱ�䣺" & nvl(rsTemp!ȡ��ʱ��))
        
        If strtext <> "" Then
            strtext = strtext & vbCrLf
            strFormatContext = strFormatContext & "\b\cf2\fs24 " & strTitle & " \par\b0\cf0\fs" & strFontSize & " " & Replace(strtext, vbCrLf, " \par\cf0\fs" & strFontSize & " ") & "\par"
        End If
        
        rsTemp.MoveNext
    Wend
    
    If Trim(str�޼�����) <> "" Then strFormatContext = strFormatContext & "\b\cf2\fs24 " & "�޼�����:" & " \par\b0\cf0\fs" & strFontSize & " " & Replace(str�޼�����, vbCrLf, " \par\cf0\fs" & strFontSize & " ") & "\par"
        
    strFormatContext = strFormatContext & "}"
    rtxtReport.SelRTF = strFormatContext
    rtxtReport.SelStart = 0
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
End Sub

Private Function LoadAttachInfo(ByVal lngAdviceId As Long, ByVal strFontSize As String, ByVal blnMovedState As Boolean) As String
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim strResult As String
    
    strResult = ""
    
    '��ʾ��������
    strSQL = "Select ��Ŀ,���� From ����ҽ������ Where ҽ��ID=[1] Order By ����"
    If blnMovedState Then
        strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
    End If

    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ���˸���", mlngAdviceId)
    Do Until rsTemp.EOF
        If nvl(rsTemp!��Ŀ) <> "" And nvl(rsTemp!����) <> "" Then
            strResult = strResult & "\b\cf0\fs" & strFontSize & " " & rsTemp!��Ŀ & ":" & " \b0\cf0\fs" & strFontSize & " " & Replace(nvl(rsTemp!����), vbCrLf, " \par\cf0\fs24 ") & "\par"
        End If
        rsTemp.MoveNext
    Loop

    strSQL = "select ��Ϣ��,��Ϣֵ from ������Ϣ�ӱ� where ����ID=[1] and ����id=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ��Ժ������Ϣ", mlngPatientId, mlngAdviceId)
    Do Until rsTemp.EOF
        If nvl(rsTemp!��Ϣ��) <> "" And nvl(rsTemp!��Ϣֵ) <> "" Then
            strResult = strResult & "\b\cf0\fs" & strFontSize & " " & rsTemp!��Ϣ�� & ":" & " \b0\cf0\fs" & strFontSize & " " & Replace(nvl(rsTemp!��Ϣֵ), vbCrLf, " \par\cf0\fs24 ") & "\par"
        End If
        rsTemp.MoveNext
    Loop
    
    LoadAttachInfo = strResult
End Function

Private Sub LoadReportContent(ByVal lngKey As Long, _
    ByVal strFormatHead As String, _
    ByVal strFontSize As String, _
    ByVal blnMovedState As Boolean)
'���뱨������
'lngReportId ����ID
'strFormatHead ��ʽͷ
'strFontSize �����С

    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim blnShow As Boolean
    Dim strFormatContext  As String
    Dim strtext As String
    Dim strTitle As String
    Dim strAttachInfo As String
    
    On Error GoTo errH
    
    strFormatContext = strFormatHead
    
    '��ʾ��������
    strAttachInfo = LoadAttachInfo(mlngAdviceId, strFontSize, blnMovedState)
    strFormatContext = strFormatContext & strAttachInfo

    '��ȡ���������
    strSQL = "Select a.�����ı� As ����, b.��������, b.�����ı� As ����,b.��ʼ�� as �汾 From ���Ӳ������� a,���Ӳ������� b " & _
             " Where a.�ļ�id = [1] And a.�������� = 3 And a.Id = b.��ID And b.�������� = 2 and b.��ֹ��=0  "
    
    If blnMovedState Then
        strSQL = Replace(strSQL, "���Ӳ�������", "H���Ӳ�������")
    End If
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ʷ�������ݲ�ѯ", lngKey)
    
    If rsTemp.RecordCount > 0 And Trim(strAttachInfo) <> "" Then
        strFormatContext = strFormatContext & "\par\b\cf0\fs" & strFontSize & "==**************************==" & "\par"
    End If
    
    If rsTemp.RecordCount <= 0 Then
        If Len(strAttachInfo) <= 0 Then
            strFormatContext = strFormatContext & "\b\cf1\fs" & strFontSize & "�����ޱ���..." & "\par"
        Else
            strFormatContext = strFormatContext & "\par\b\cf1\fs" & strFontSize & "�����ޱ���..." & "\par"
        End If
    End If
                
    While Not rsTemp.EOF
        blnShow = False
        Select Case rsTemp!����
            Case "�������"
                strTitle = mstrDescriptionName
                strtext = nvl(rsTemp!����) & vbCrLf
                blnShow = True
            Case "������"
                strTitle = mstrOpinionName
                strtext = nvl(rsTemp!����) & vbCrLf
                blnShow = True
            Case "����"
                strTitle = mstrAdviseName
                strtext = nvl(rsTemp!����) & vbCrLf
                blnShow = True
        End Select
        
        If blnShow = True Then strFormatContext = strFormatContext & "\b\cf2\fs24 " & strTitle & " \par\b0\cf0\fs" & strFontSize & " " & Replace(strtext, vbCrLf, " \par\cf0\fs" & strFontSize & " ") & "\par"
        rsTemp.MoveNext
    Wend
    
    strFormatContext = strFormatContext & "}"
    rtxtReport.SelRTF = strFormatContext
    rtxtReport.SelStart = 0
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub LoadReport(ByVal lngAdviceId As Long, _
    ByVal strLoadType As String, ByVal lngKeyId As Long, _
    ByVal blnMoveState As Boolean)
'��������lvHistoryList.ListItems�ؼ��ַ�Ϊ process�����̱��� ��describe���޼�������������������������������ݣ�ԭ��ʹ�õ�K��
On Error GoTo err
    Dim strSQL As String
    Dim strtext As String
    Dim strFormatContext As String
    Dim strSize As String
    Dim rsTemp As ADODB.Recordset
    
    rtxtReport.Text = ""
    
    strSize = FontSize
    strSize = 2 * Round(Val(strSize))
    

    strFormatContext = "{\rtf1\ansi\ansicpg936\deff0\deflang1033\deflangfe2052{\fonttbl{\f0\fnil\fcharset134 \'cb\'ce\'cc\'e5;}}" & _
                       "{\colortbl ;\red255\green104\blue104;\red19\green164\blue251;}" & _
                       "{\*\generator Msftedit 5.41.21.2509;}\viewkind4\uc1\sl276\slmult1\lang2052\b\f0\fs24 "
         
    If InStr(strLoadType, M_STR_LISTVIEWKET_PROCESS) > 0 Then
        Call LoadProcessReport(lngKeyId, strFormatContext, strSize, blnMoveState)
    ElseIf InStr(strLoadType, M_STR_LISTVIEWKEY_DESCRIBE) > 0 Then
        Call LoadDescription(lngKeyId, strFormatContext, strSize, blnMoveState)
    Else
        '����ҽ��id����ѯ��Ӧ����ID
        strSQL = "select ����ID from ����ҽ������ where ҽ��ID=[1] "
        If blnMoveState Then
            strSQL = Replace(strSize, "����ҽ������", "H����ҽ������")
        End If
        
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ʷҽ������ID��ѯ", lngAdviceId)
        
        If rsTemp.RecordCount > 0 Then
            Call LoadReportContent(Val(nvl(rsTemp!����Id)), strFormatContext, strSize, blnMoveState)
        Else
            Call LoadReportContent(0, strFormatContext, strSize, blnMoveState)
        End If
        
    End If

    Exit Sub
err:
    If ErrCenter = 1 Then
        Resume
    Else
        Call MsgboxH(GetRootHwnd, err.Description, vbOKOnly, "��ʾ")
    End If
End Sub


Private Function GetListHeadString() As String
'�õ���������: ����,���,�Ƿ���ʾ  ����  "���,1000,1|ִ�й���,2000,0|"
On Error GoTo errH
    Dim i As Integer
    Dim strTemp As String
    Dim strName As String
    Dim lngWidth As Long
    Dim blnIsHide As Boolean
    
    For i = 0 To vsfStudy.Cols - 1
        
        strName = vsfStudy.TextMatrix(0, i)
        
        lngWidth = vsfStudy.ColWidth(i)
        blnIsHide = vsfStudy.ColHidden(i)
        
        If Len(strTemp) > 0 Then
            strTemp = strTemp & "|"
        End If
        
        strTemp = strTemp & strName & "," & lngWidth & "," & blnIsHide
    Next

    GetListHeadString = strTemp
    
    Exit Function
errH:
    err.Raise -1, "��ʷ���", "[��ȡ��ͷ����]" & vbCrLf & err.Description
    Resume
End Function

Private Sub DoLoadListSort(ByVal strcfg As String)
'�ָ�����
On Error GoTo errH
    Dim strName As String
    Dim intWay As Integer
    Dim intPos As Integer
    Dim intCol As Integer
    Dim i As Integer
    
    intPos = InStr(strcfg, ",")
    If intPos = 0 Then Exit Sub
    
    strName = Split(strcfg, ",")(0)
    intWay = Val(Split(strcfg, ",")(1))
    
    With vsfStudy
        For i = 1 To .Cols - 1
            If strName = .TextMatrix(0, i) Then
                intCol = i
                Exit For
            End If
        Next
         
        .Col = intCol
        .Sort = intWay
        
        For i = 1 To .Rows - 1
            .TextMatrix(i, vsfStudy.ColIndex("���")) = i
        Next
    End With
    
    Exit Sub
errH:
    err.Raise -1, "�б���Ի�����", "[DoLoadListSort]" & vbCrLf & err.Description
    Resume
End Sub





