VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "*\Azl9PacsControl\zl9PacsControl.vbp"
Begin VB.UserControl ucReportSegment 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6585
   BeginProperty Font 
      Name            =   "����"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LockControls    =   -1  'True
   ScaleHeight     =   9000
   ScaleWidth      =   6585
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1440
      Top             =   0
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
            Picture         =   "ucReportSegment.ctx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucReportSegment.ctx":06FA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picBack 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8415
      Left            =   120
      ScaleHeight     =   8415
      ScaleWidth      =   6255
      TabIndex        =   0
      Top             =   120
      Width           =   6255
      Begin zl9PacsControl.ucSplitter ucSplitter1 
         Height          =   135
         Left            =   0
         TabIndex        =   1
         Top             =   3015
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   238
         MousePointer    =   7
         SplitType       =   0
         SplitLevel      =   3
         Con1MinSize     =   1000
         Con2MinSize     =   2000
         Control1Name    =   "trvWordTree"
         Control2Name    =   "vsWordContext"
      End
      Begin VSFlex8Ctl.VSFlexGrid vsWordContext 
         Height          =   5265
         Left            =   0
         TabIndex        =   2
         Top             =   3150
         Width           =   6255
         _cx             =   11033
         _cy             =   9287
         Appearance      =   0
         BorderStyle     =   0
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   16777215
         BackColorAlternate=   16777215
         GridColor       =   14737632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   1
         HighLight       =   0
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   4
         SelectionMode   =   0
         GridLines       =   4
         GridLinesFixed  =   0
         GridLineWidth   =   0
         Rows            =   0
         Cols            =   2
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   0
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
         FillStyle       =   1
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   2
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
         AllowUserFreezing=   3
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
         Begin RichTextLib.RichTextBox txtWordEdit 
            Height          =   1935
            Left            =   1320
            TabIndex        =   4
            Top             =   2040
            Visible         =   0   'False
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   3413
            _Version        =   393217
            ScrollBars      =   2
            Appearance      =   0
            AutoVerbMenu    =   -1  'True
            TextRTF         =   $"ucReportSegment.ctx":0DF4
         End
      End
      Begin MSComctlLib.TreeView trvWordTree 
         Height          =   3015
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   5318
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   176
         LineStyle       =   1
         Style           =   7
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Image imgAdvi 
      Height          =   360
      Left            =   5400
      Picture         =   "ucReportSegment.ctx":0E91
      Top             =   8520
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgOpin 
      Height          =   360
      Left            =   4800
      Picture         =   "ucReportSegment.ctx":1593
      Top             =   8520
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgDesc 
      Height          =   360
      Left            =   4200
      Picture         =   "ucReportSegment.ctx":1C95
      Top             =   8520
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   3720
      Picture         =   "ucReportSegment.ctx":2397
      Top             =   8520
      Visible         =   0   'False
      Width           =   360
   End
End
Attribute VB_Name = "ucReportSegment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const MOUSEEVENTF_RIGHTDOWN = &H8 '  right button down
Private Const MOUSEEVENTF_RIGHTUP = &H10 '  right button up


Private Const MOUSEEVENTF_LEFTDOWN = &H2 '  left button down
Private Const MOUSEEVENTF_LEFTUP = &H4 '  left button up
                
Private Const NODE_BACKCOLOR_DISABLE As Long = &HF1F1F1
Private Const NODE_FORCECOLOR_DISABLE As Long = &HC0C0C0    '&H808080


Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)



Private Const LVW_KEY_WORD As String = "L"  ' ���οؼ���Ŀ
Private Const LVW_KEY_NODE As String = "T"   ' ���οؼ�Ŀ¼
  
    
Private mrsClass As ADODB.Recordset     '�ʾ����
Private mrsWords As ADODB.Recordset     '�ʾ���Ŀ

Private mFileID As Long                 '����ID
Private mstrOutLineKey As String        '�ʾ�ʾ����������,��ٹؼ��֣� ������������ϡ� ��������������顱��
Private mlngOutlineId As Long
Private mlngAdviceId As Long            'ҽ��ID
Private mlngFileID As Long

Private mstrDBOwner As String              '���ݿ�������
Private mintWordDblClickMode As Integer     '�ʾ�˫����Ĳ�����0--ֱ��д�뱨�棻1--�򿪴ʾ�༭����
Private mintWordPower As Integer        '�ʾ����Ȩ��Χ

Private mlngWordTreeH As Long               '�ʿ�ģ�����ĸ߶�
Private mlngWordShowH As Long               '�ʿ�ģ�����ݵĸ߶�


Private mlngCurModule As Long
Private mlngCurDeptId As Long

Private mlngPatientId As Long
Private mlngPageID As Long
Private mblnAdviceMoved As Boolean

Private mblnIsInit As Boolean           '�Ƿ��ʼ��
Private mlngExpandLevel As Long         '�Զ�չ���㼶,Ĭ��Ϊ1
Private mblnIsWordValid As Boolean      '�Ƿ�Դʾ��������������ж�
Private mblnAutoRemove As Boolean       '�Ƿ��Զ��Ƴ������ôʾ估����
Private mblnIsSyncWordFragment As Boolean

Public Event OnRequestState(ByRef lngOutlineType As TOutlineType, _
                            ByRef str�������� As String, ByRef str������� As String, ByRef str�������� As String)
    
Public Event OnSendContext(ByVal strFreeText As String, _
                            ByVal str�������� As String, ByVal str������� As String, ByVal str�������� As String)
                            
Public Event OnMouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Property Get IsSyncWordFragment() As Boolean
    IsSyncWordFragment = mblnIsSyncWordFragment = True
End Property
                            
                            
Property Get hwnd() As Long
    hwnd = UserControl.hwnd
End Property

Property Get DblWrite() As Boolean
    DblWrite = IIf(mintWordDblClickMode = 0, True, False)
End Property

Property Let DblWrite(value As Boolean)
    mintWordDblClickMode = IIf(value, 0, 1)
End Property
 

'�ڵ�����
Property Get NodeCount() As Long
    NodeCount = trvWordTree.Nodes.Count
End Property

'ѡ��ڵ�����
Property Get SelNodeType() As Long
    SelNodeType = 0
    
    If trvWordTree.SelectedItem Is Nothing Then Exit Property
    
    If Left(trvWordTree.SelectedItem.Key, 1) = LVW_KEY_WORD Then
        SelNodeType = 2
    Else
        SelNodeType = 1
    End If
End Property

'չ������
Property Get ExpandLevel() As Long
    ExpandLevel = mlngExpandLevel
End Property

Property Let ExpandLevel(value As Long)
    mlngExpandLevel = value
    
    Call AutoExpand
End Property

'�Զ�����
Property Get AutoHide() As Boolean
    AutoHide = mblnAutoRemove
End Property

Property Let AutoHide(value As Boolean)
    mblnAutoRemove = value
    
    If mFileID <> 0 Then
        Call LoadWordClass(mFileID, mstrOutLineKey, True)
    End If
End Property


Public Sub Init(ByVal lngModuleNo As Long, ByVal lngDeptId As Long, _
    Optional ByVal blnIsForce As Boolean = False)
On Error GoTo errhandle
    If mblnIsInit And blnIsForce = False Then Exit Sub
    
    mlngCurModule = lngModuleNo
    mlngCurDeptId = lngDeptId
    
'    intWordPower=-1�����߱��ʾ����Ȩ;
'    intWordPower=0��ȫԺ����ʱ��ʾ���е�ʾ����Ҳ���Ը���;
'    intWordPower=1�����ң���ʱ��ʾȫԺͨ��ʾ��(����id is null)�����ڿ��ҹ��л�������Ա˽�е�ʾ���������ܸ���ȫԺͨ��ʾ��;
'    intWordPower=2�����ˣ���ʱ��ʾȫԺͨ��ʾ��(����id is null)�����ڿ���ͨ��ʾ��(��Աid is null)�͸���ʾ����������ʾ���ɸ���
    
    mintWordPower = zlGetWordPower
    
    Call InitDbOwner(glngSys)
    
    trvWordTree.ImageList = ImageList1
    
    mstrOutLineKey = ""
    
    Call InitLoaclParas
    
    mblnIsInit = True
Exit Sub
errhandle:
    mblnIsInit = False
End Sub


Public Sub SetFontSize(ByVal bytFontSize As Byte)
    FontSize = bytFontSize
    
    picBack.FontSize = bytFontSize
    vsWordContext.FontSize = bytFontSize
    Set txtWordEdit.Font = Font
    
    Set trvWordTree.Font = Font
End Sub

Private Sub InitPatientInfo(ByVal lngAdviceId As Long)
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
 
    
    strSQL = "select a.����ID,a.��ҳID, a.������Դ,a.�Ա�,a.Ӥ��,b.����id, 0 as ת�� from ����ҽ����¼ a, Ӱ�����¼ b Where a.id=b.ҽ��id(+) and a.id=[1] " & _
        "Union All " & _
        "select a.����ID, a.��ҳID, a.������Դ,a.�Ա�,a.Ӥ��,b.����id, 1 as ת�� from H����ҽ����¼ a, HӰ�����¼ b Where a.id=b.ҽ��id(+) and a.id=[1] "
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "ҽ����Ϣ��ѯ", lngAdviceId)
    If rsTemp.RecordCount <= 0 Then Exit Sub
    
    mlngPatientId = Val(nvl(rsTemp!����ID))
    mlngPageID = Val(nvl(rsTemp!��ҳID))
'    mlngPatientFrom = Val(nvl(rsTemp!������Դ))
'    mlngBabyNum = Val(nvl(rsTemp!Ӥ��))
    mblnAdviceMoved = IIf(Val(nvl(rsTemp!ת��)) = 1, True, False)
End Sub


Private Sub LoadWordClass(FileID As Long, strOutlineKey As String, Optional blnForceRefresh As Boolean = False)
    Dim strSQL As String
    Dim rsCurClass As ADODB.Recordset
    Dim rsCurWords As ADODB.Recordset
    
    Dim rsTemp As ADODB.Recordset
    Dim objNode As Node
    Dim objPnode As Node
    Dim strKey As String
    Dim blnIsOnlyRefreshOutline As Boolean
    Dim i As Long
    Dim strUserInfo As String
    Dim strWith As String
    Dim lngIndex As Long
    Dim aryOutlineId(3) As Long
    
    
    blnIsOnlyRefreshOutline = False
    
    If FileID = mFileID And trvWordTree.Nodes.Count > 0 And blnForceRefresh = False Then
        Set rsCurClass = mrsClass
'        Set rsCurWords = mrsWords
        
        If strOutlineKey <> mstrOutLineKey Then
            '��ٲ�ͬ�������ٽ��д���
            blnIsOnlyRefreshOutline = True
        Else
            '���ҳ��ͬʱ�����˳�
            Exit Sub
        End If
    Else
        Set rsCurClass = Nothing
'        Set rsCurWords = Nothing
    End If
    
    mFileID = FileID
    mstrOutLineKey = strOutlineKey
    
    strSQL = "Select nvl(a.��id,0) as ���ID  From �����ļ��ṹ a" & _
             " Where a.�ļ�ID=[1] and a.�����ı� like '%' || [2] || '%' And a.��������=3 And Rownum =1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ѯ�������", FileID, mstrOutLineKey)
    If rsTemp.RecordCount <= 0 Then
        trvWordTree.Nodes.Clear
        vsWordContext.Rows = 0
        txtWordEdit.Text = ""
        
        Exit Sub
    End If
    
    mlngOutlineId = Val(nvl(rsTemp!���id))
    
    '���ģ������
    vsWordContext.Rows = 0
    
    If mblnAutoRemove = False Then '�����Ҫ�Զ����أ�����Ҫ������ڵ����¼��أ�����������أ���ֱ�����ýڵ�״̬
        If blnIsOnlyRefreshOutline Then
            Call HideOutlineNode(mlngOutlineId)
            Exit Sub
        End If
    End If
    
    '��������API�����Ҳ�������ѭ��ɾ��TreeView�ķ�������������ٶȸ���
    Call TrvwClear
            
    If rsCurClass Is Nothing Then
        '��ѯ�ʾ����
'        strSQL = "Select * from (with OutLinesTab as (" & _
'                             " Select nvl(��id,0) as ���ID " & _
'                             " From �����ļ��ṹ " & _
'                             " Where �ļ�ID=[1]  And ��������=3   ) " & _
'                        " select a.ID, a.�ϼ�ID,a.����,a.����,b.���ID " & _
'                        " from  �����ʾ���� a, OutLinesTab b " & _
'                        " where  a.Id In ( " & _
'                        "                 select id " & _
'                        "                 from �����ʾ���� x " & _
'                        "                 start with x.id in( " & _
'                        "                                   select �ʾ����id " & _
'                        "                                   from ������ٴʾ� a " & _
'                        "                                   Where a.���ID = b.���ID ) " & _
'                        "                 Connect By Prior �ϼ�id=Id " & _
'                        " ) and substr(a.��Χ,7,1)='1') order by Id"
        strSQL = "select  distinct a.���ID from ������ٴʾ� a, �����ļ��ṹ b where a.���ID=nvl(��id,0) and b.�ļ�ID=[1] and ��������=3"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ѯ�����ʾ����", mlngFileID)
        
        If rsTemp.RecordCount <= 0 Then
            
            Exit Sub
        End If
        
        lngIndex = 1
        strWith = ""
        strSQL = ""
        
        
'with OutlineReleation as (select ���ID, �ʾ����ID from ������ٴʾ� where ���ID=122),
'     OutlineReleation1 as (select ���ID, �ʾ����ID from ������ٴʾ� where ���ID=782),
'     outlineWord as (select id,�ϼ�ID,����,����,��Χ from �����ʾ���� where  substr(��Χ,7,1)='1' )
'select a.id, a.�ϼ�ID,a.����, ������, b.���ID as ������� from (
'       select id,�ϼ�ID,����,122 as ������ from outlineWord
'       start with  id in(select  �ʾ����ID from OutlineReleation)
'       Connect By Prior �ϼ�id=Id
') a, OutlineReleation b
'where a.id=b.�ʾ����ID(+)
'
'Union All
'
'select a.id, a.�ϼ�ID,a.����,������, b.���ID as ������� from (
'       select id,�ϼ�ID,���� , 782 as ������ from outlineWord
'       start with  id in(select  �ʾ����ID from OutlineReleation1)
'       Connect By Prior �ϼ�id=Id
') a, OutlineReleation1 b
'where a.id=b.�ʾ����ID(+)
'order by ID
    

        While Not rsTemp.EOF
            If strWith <> "" Then strWith = strWith & "," & vbCrLf
            strWith = strWith & "OutlineReleation" & lngIndex & " as (select ���ID, �ʾ����ID from ������ٴʾ� where ���ID=[" & lngIndex & "])"
            
            If strSQL <> "" Then strSQL = strSQL & vbCrLf & "Union All " & vbCrLf
            strSQL = strSQL & "select a.id, a.�ϼ�ID,a.����, a.����, ������, b.���ID as ������� from (" & vbCrLf & _
                            "    select id,�ϼ�ID,����,����,[" & lngIndex & "] as ������ from outlineWord" & vbCrLf & _
                            "    start with  id in(select  �ʾ����ID from OutlineReleation" & lngIndex & ") " & vbCrLf & _
                            "    Connect By Prior �ϼ�id=Id ) a, OutlineReleation" & lngIndex & " b where a.id=b.�ʾ����ID(+)"
            
            aryOutlineId(lngIndex) = Val(nvl(rsTemp!���id))
            lngIndex = lngIndex + 1
            Call rsTemp.MoveNext
        Wend
        
        strSQL = "select * from (with " & strWith & "," & vbCrLf & _
                        " OutlineWord as (select id,�ϼ�ID,����,����,��Χ from �����ʾ���� where  substr(��Χ,7,1)='1' )" & vbCrLf & _
                        strSQL & vbCrLf & _
                        ") Order by ������, ID"
        
        
                        
        Set mrsClass = zlDatabase.OpenSQLRecord(strSQL, "��ѯ�ʾ����", aryOutlineId(1), aryOutlineId(2), aryOutlineId(3))
'        Set mrsClass = zlDatabase.CopyNewRec(mrsClass)
        
        If mrsClass.RecordCount <= 0 Then Exit Sub
        
        Set rsCurClass = mrsClass
    End If
    
        '��ѯ�ʾ�
'        strSQL = "select /*+ RULE*/ b.ID,b.����ID,b.���� " & _
'                   " from  �����ʾ���� a,   �����ʾ�ʾ�� b" & IIf(mblnAutoRemove, ", Table(Cast(f_Sentence_Usable([2], [3], [4], [5]) as zlhis.t_Dic_Rowset )) C ", "") & _
'                   " where a.id=b.����ID " & IIf(mblnAutoRemove, " and b.Id=c.���� ", "") & " and a.Id In ( " & _
'                   "        select id " & _
'                   "        from �����ʾ���� x " & _
'                   "        start with x.id in( " & _
'                   "                select �ʾ����id " & _
'                   "                from ������ٴʾ� a " & _
'                   "                where a.���ID in ( " & _
'                   "                       Select nvl(��id,0) as ���ID " & _
'                   "                       From �����ļ��ṹ " & _
'                   "                       Where �ļ�ID=[1]  And ��������=3 ) " & _
'                   "                     ) " & _
'                   "        Connect By Prior �ϼ�id=Id " & _
'                   "        ) and substr(a.��Χ,7,1)='1' order by Id "
        strSQL = "select /*+ RULE*/ b.ID,b.����ID,b.���� " & _
                   " from  �����ʾ���� a,   �����ʾ�ʾ�� b" & IIf(mblnAutoRemove, ", Table(Cast(f_Sentence_Usable([2], [3], [4], [5]) as zlhis.t_Dic_Rowset )) C ", "") & _
                   " where a.id=b.����ID " & IIf(mblnAutoRemove, " and b.Id=c.���� ", "") & " and a.Id In ( " & _
                   "                select �ʾ����id " & _
                   "                from ������ٴʾ� a " & _
                   "                where a.���ID in ( " & _
                   "                       Select nvl(��id,0) as ���ID " & _
                   "                       From �����ļ��ṹ " & _
                   "                       Where �ļ�ID=[1]  And ��������=3 ) " & _
                   "                      " & _
                   "        ) and substr(a.��Χ,7,1)='1' order by Id "
        Set mrsWords = zlDatabase.OpenSQLRecord(strSQL, "��ѯ�ʾ���Ŀ", FileID, mlngOutlineId, mlngPatientId, mlngPageID, mlngAdviceId)
'        Set mrsWords = zlDatabase.CopyNewRec(mrsWords)
        
        If mrsWords.RecordCount <= 0 Then Exit Sub
        
        Set rsCurWords = mrsWords
         
     
    If mblnAutoRemove Then
        rsCurClass.Filter = "������=" & mlngOutlineId
    Else
        rsCurClass.Filter = ""
    End If
    
    rsCurClass.Sort = "����"
    
    strUserInfo = "[" & UserInfo.�û��� & "]"
    '�������з���
    Do While Not rsCurClass.EOF
        
        Set objNode = Nothing
        
        On Error Resume Next
        Set objNode = trvWordTree.Nodes("T-" & rsCurClass("ID").value)
        
        If err.Number <> 0 Then
            Set objNode = Nothing
            err.Clear
        End If
        
        If zlCommFun.nvl(rsCurClass("�ϼ�id").value, 0) <> 0 Then
            Set objPnode = trvWordTree.Nodes("T-" & rsCurClass("�ϼ�id").value)
            
            If err.Number <> 0 Then
                Set objPnode = Nothing
                err.Clear
            End If
        Else
            Set objPnode = Nothing
        End If
        
        On Error GoTo errhandle
        
        If objNode Is Nothing Then
            If objPnode Is Nothing Then
                Set objNode = trvWordTree.Nodes.Add(, , "T-" & rsCurClass("ID").value, Replace(rsCurClass("����").value, strUserInfo, ""), 2)
            Else
                Set objNode = trvWordTree.Nodes.Add("T-" & zlCommFun.nvl(rsCurClass("�ϼ�id").value, 0), tvwChild, "T-" & rsCurClass("ID").value, Replace(rsCurClass("����").value, strUserInfo, ""), 2)
            End If
             
            objNode.tag = 0  '��ʾ��δ���شʾ�ڵ�
        End If
    
        rsCurClass.MoveNext
    Loop
    
    '���ز����ڱ���ٵĴʾ����ڵ�,
    If mblnAutoRemove = False Then Call HideOutlineNode(mlngOutlineId)
    
    '����չ���㼶�Զ�չ��
    Call AutoExpand
    
    Exit Sub
errhandle:
    If err.Number <> 35602 Then
        If ErrCenter() = 1 Then Resume Next
        Call SaveErrLog
    End If
End Sub

Private Sub AutoExpand()
    Dim objNode As Node
    Dim i As Long
    Dim objSelNode As Node
    
    LockWindowUpdate trvWordTree.hwnd
On Error GoTo errhandle
    Set objSelNode = trvWordTree.SelectedItem
    '����չ���㼶�Զ�չ��
    For i = 1 To trvWordTree.Nodes.Count
        Set objNode = trvWordTree.Nodes(i)
        
        If Left(objNode.Key, 1) = LVW_KEY_NODE Then
            objNode.Expanded = False
            
            If GetNodeDepth(objNode) < IIf(mlngExpandLevel = 0, 999, mlngExpandLevel) Then
                objNode.Expanded = True
                
                If Val(objNode.tag) <> 1 Then
                    Call LoadWordItem(objNode)
                    objNode.tag = 1 '��ʾ�Ѿ������˴ʾ�ڵ�
                End If
                
                If objNode.BackColor = NODE_BACKCOLOR_DISABLE Then objNode.Expanded = False
            End If
        End If
    Next
    
    '�ָ�ѡ��Ľڵ�
    If Not objSelNode Is Nothing Then
        While GetNodeDepth(objSelNode) > IIf(mlngExpandLevel = 0, 999, mlngExpandLevel)
            Set objSelNode = objSelNode.Parent
        Wend
        
        objSelNode.Selected = True
    End If
    
    LockWindowUpdate 0
Exit Sub
errhandle:
    LockWindowUpdate 0
End Sub

Private Sub LoadWordItem(objNode As Node)
    Dim lngWordClassId As Long
    Dim objSubNode As Node
    Dim objSubClassNode As Node
    Dim i As Long
    
    lngWordClassId = Split(objNode.Key, "-")(1)
    
    mrsWords.Filter = "����ID=" & lngWordClassId
    If mrsWords.RecordCount > 0 Then
        '���ص�ǰ�ڵ��µĴʾ�
        Do While Not mrsWords.EOF
            On Error Resume Next
            Set objSubNode = trvWordTree.Nodes("L-" & mrsWords("ID").value)
            
            If err.Number <> 0 Then
                Set objSubNode = Nothing
                err.Clear
            End If
            
            If objSubNode Is Nothing Then
                Set objSubNode = trvWordTree.Nodes.Add(objNode, tvwChild, "L-" & mrsWords("ID").value, mrsWords("����").value, 1)
                objSubNode.tag = -1 '��ʾû�н����������ж�
            End If
            
            '�жϸôʾ��Ƿ�Ըñ���ģ������
            Call mrsWords.MoveNext
        Loop
    End If
    
    Set objSubClassNode = objNode.Child
    '�����ӽڵ��µĵ�һ���ʾ�
    While Not objSubClassNode Is Nothing
        
        lngWordClassId = Split(objSubClassNode.Key, "-")(1)
        mrsWords.Filter = "����ID=" & lngWordClassId
        If mrsWords.RecordCount > 0 Then
            On Error Resume Next
            Set objSubNode = trvWordTree.Nodes("L-" & mrsWords("ID").value)
            
            If err.Number <> 0 Then
                Set objSubNode = Nothing
                err.Clear
            End If
        
            If objSubNode Is Nothing Then
                Set objSubNode = trvWordTree.Nodes.Add(objSubClassNode, tvwChild, "L-" & mrsWords("ID").value, mrsWords("����").value, 1)
                objSubNode.tag = -1 '��ʾû�н����������ж�
            End If
        End If
        
        Set objSubClassNode = objSubClassNode.Next
    Wend
End Sub

Private Sub HideOutlineNode(ByVal lngOutlineId As Long)
    Dim rsOutlineClass As ADODB.Recordset
    Dim i As Long
    Dim lngClassID As Long
    Dim objNode As Node
    Dim objSubNode As Node

    mrsClass.Filter = ""
    Set rsOutlineClass = mrsClass.Clone

    For i = trvWordTree.Nodes.Count To 1 Step -1
        Set objNode = trvWordTree.Nodes(i)
        
        If Left(objNode.Key, 1) = LVW_KEY_NODE Then
            lngClassID = Val(Split(objNode.Key & "-", "-")(1))
            
            rsOutlineClass.Filter = "�������=" & lngOutlineId & " and ID=" & lngClassID
    
            'Node.tag:0_1_�ı����� ��Ӧ˵�� ҽ��ID_����״̬_�ı�����
            If rsOutlineClass.RecordCount > 0 Then
                '��ٴ��ڶ�Ӧ����
                objNode.BackColor = vbWhite
                objNode.ForeColor = vbBlack
            Else
                objNode.BackColor = NODE_BACKCOLOR_DISABLE
                objNode.ForeColor = NODE_FORCECOLOR_DISABLE
            End If
            
            
            If objNode.Children > 0 Then
                 Set objSubNode = objNode.Child
                 
                 While Not objSubNode Is Nothing
                     If Left(objSubNode.Key, 1) = LVW_KEY_WORD Then
                         objSubNode.BackColor = objNode.BackColor
                         objSubNode.ForeColor = objNode.ForeColor
                     End If
                     
                     Set objSubNode = objSubNode.Next
                 Wend
             End If
        End If
    Next
End Sub


Private Function GetNodeDepth(objNode As Object) As Long
'��ȡ�ڵ����
    GetNodeDepth = UBound(Split(objNode.FullPath, trvWordTree.PathSeparator))
End Function


Private Sub Form_Unload(Cancel As Integer)
'    Dim strRegPath As String
'
'
'    strRegPath = "����ģ��\" & App.ProductName & "\frmReportWord"
'
'    '����ʾ�ʾ������ĸ߶�
'    '285��Pane�ı���߶ȣ�ʹ���˱��⣬����Ҫ�ӻ�����߶�
'    If Not (picWordTree.Height = 0 And picWordShow.Height = 0 And picPrivateWord.Height = 0) Then
'      SaveSetting "ZLSOFT", strRegPath, "WordTreeH", picWordTree.Height
'      SaveSetting "ZLSOFT", strRegPath, "WordShowH", picWordShow.Height
'      SaveSetting "ZLSOFT", strRegPath, "PrivateWordH", picPrivateWord.Height ' + 285
'    End If
'    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\frmReportWord", "ֱ�ӱ༭", CLng(chkֱ�ӱ༭.value)
'
'    If mblnShowWord = False Then    'ͨ��˫���򿪣�����ʾȷ����ȡ����ť,��¼����߶�
'        SaveSetting "ZLSOFT", strRegPath, "ButtonH", picCommandButton.Height
'    End If
'
'    '����ʾ�ʾ������Ŀ��
'    If mblnSingleWindow = True Then
'        strRegPath = "����ģ��\" & App.ProductName & "\frmReport\SingleWindow"
'    Else
'        strRegPath = "����ģ��\" & App.ProductName & "\frmReport"
'    End If
'    SaveSetting "ZLSOFT", strRegPath, "CX1", picWordTree.Width
'
'    '����ģʽ,��ģʽ�¼�¼����λ��
'    If mblnShowWord = False Then
'        Call SaveWinState(Me, App.ProductName)
'    End If
End Sub
 



 

'Private Sub menuAutoHide_Click()
'On Error GoTo errHandle
'    menuAutoHide.Checked = Not menuAutoHide.Checked
'    mblnAutoRemove = menuAutoHide.Checked
'
'    SaveSetting "ZLSOFT", mstrRegPrivatePath, "�Զ�����", mblnAutoRemove
'
'    Call LoadWordClass(mFileID, mstrOutLineName, True)
'Exit Sub
'errHandle:
'    MsgBoxH hWnd, err.Description, vbOKOnly, "��ʾ"
'End Sub

Private Function GetRootHwnd() As Long
    GetRootHwnd = GetAncestor(hwnd, GA_ROOT)
End Function

Public Sub DirectWrite()
'ֱ��д��ʾ�
On Error GoTo errhandle
    Dim objSelNode As Node

    Set objSelNode = trvWordTree.SelectedItem
    
    If Not objSelNode Is Nothing Then
        If Left(objSelNode.Key, 1) = LVW_KEY_WORD Then
            '�ʾ�˫���󣬴򿪴ʾ�༭����
            Call WriteWordDirect
        End If
    End If
Exit Sub
errhandle:
    MsgboxH GetRootHwnd, err.Description, vbOKOnly, "��ʾ"
End Sub


Public Sub EditWrite()
On Error GoTo errhandle
    Dim objSelNode As Node

    Set objSelNode = trvWordTree.SelectedItem
    
    If Not objSelNode Is Nothing Then
        If Left(objSelNode.Key, 1) = LVW_KEY_WORD Then
            '�ʾ�˫���󣬴򿪴ʾ�༭����
            WriteWordEdit Val(Split(objSelNode.Key & "-", "-")(1))
        End If
    End If
Exit Sub
errhandle:
    MsgboxH GetRootHwnd, err.Description, vbOKOnly, "��ʾ"
End Sub


Private Sub WriteWordDirect()
'ֱ��д��
    Dim i As Long
    Dim objNode As Node
    Dim lngApply As Long
    
    Set objNode = trvWordTree.SelectedItem
    
    If objNode Is Nothing Then Exit Sub
    If vsWordContext.Rows <= 0 Then Exit Sub
    
    lngApply = Val(Split(objNode.tag & "__", "_")(1))
    
    If lngApply = 0 Then
        If MsgboxH(GetRootHwnd, "�ôʾ䲻�����ڵ�ǰ��٣��Ƿ������", vbYesNo + vbDefaultButton2, "��ʾ") = vbNo Then Exit Sub
    End If
    
    For i = 0 To vsWordContext.Rows - 1
        If vsWordContext.RowData(i) <> "WARING" Then
            Call DoWritWord(i, False)
        End If
    Next
End Sub
 

 

'��ʱ���ṩ������ش�����
'Private Sub menuNewClass_Click()
'On Error GoTo errHandle
'    Call NewClass
'Exit Sub
'errHandle:
'    MsgBoxH GetRootHwnd,  err.Description, vbOKOnly, "��ʾ"
'End Sub

'Private Sub NewClass()
''��������
'    Dim objPNode As Node
'    Dim objSubNode As Node
'    Dim strSql As String
'    Dim rsData As ADODB.Recordset
'    Dim lngPId As Long
'    Dim strPCode As String
'    Dim rsClass As ADODB.Recordset
'    Dim lngCurClassId As Long
'    Dim strCurClassCode As String
'    Dim strCurClassName As String
'    Dim i As Long
'On Error GoTo errHandle
'    Set objPNode = trvWordTree.SelectedItem
'
'    If objPNode Is Nothing Then Exit Sub
'    If Left(objPNode.Key, 1) = LVW_KEY_WORD Then Exit Sub
'
'    lngPId = Split(objPNode.Key, "-")(1)
'
'    Set rsClass = mrsClass.Clone
'    rsClass.Filter = "ID=" & lngPId
'
'    strPCode = ""
'    If rsClass.RecordCount > 0 Then strPCode = nvl(rsClass!����)
'
'    strSql = "select �����ʾ����_ID.NEXTVAL as �ʾ�ID from dual"
'    Set rsData = zlDatabase.OpenSQLRecord(strSql, "��ѯ�ʾ����ID")
'    If rsData.RecordCount <= 0 Then
'        MsgBoxH GetRootHwnd,  "���ܻ�ȡ�ʾ����ID.", vbOKOnly, "��ʾ"
'        Exit Sub
'    End If
'
'    lngCurClassId = Val(nvl(rsData!�ʾ�Id))
'
'
'    strSql = "select nvl(max(����), 0) as ���� from �����ʾ���� where �ϼ�ID=[1]"
'    Set rsData = zlDatabase.OpenSQLRecord(strSql, "��ѯ�����ʾ�������", lngPId)
'    If rsData.RecordCount <= 0 Then
'        MsgBoxH GetRootHwnd,  "���ܻ�ȡ�ʾ�������.", vbOKOnly, "��ʾ"
'        Exit Sub
'    End If
'
'    strCurClassName = "�·���1"
'    If Val(nvl(rsData!����)) = 0 Then
'        strCurClassCode = strPCode & "01"
'    Else
'        strCurClassCode = Val(nvl(rsData!����)) + 1
'
'        rsClass.Filter = "�ϼ�ID=" & lngPId & " and ����='" & strCurClassName & "[" & UserInfo.�û��� & "]" & strCurClassName & "'"
'
'        i = 1
'        While rsClass.RecordCount > 0
'            i = i + 1
'            strCurClassName = "�·���" & i
'            rsClass.Filter = "�ϼ�ID=" & lngPId & " and ����='" & "[" & UserInfo.�û��� & "]" & strCurClassName & "'"
'        Wend
'    End If
'
'    If Len(strCurClassCode) > 8 Then
'        MsgBoxH GetRootHwnd,  "����㼶�ѳ������ƣ����ܼ��������ӷ��ࡣ", vbOKOnly, "��ʾ"
'        Exit Sub
'    End If
'
'    strSql = "Zl_�����ʾ����_Edit(1," & lngCurClassId & "," & lngPId & ",'" & strCurClassCode & "','" & _
'                                "[" & UserInfo.�û��� & "]" & strCurClassName & "','','00000010')"
'    Call zlDatabase.ExecuteProcedure(strSql, "�����ʾ����")
'
'    mrsClass.AddNew
'    mrsClass!ID = lngCurClassId
'    mrsClass!�ϼ�ID = lngPId
'    mrsClass!���� = strCurClassCode
'    mrsClass!���� = "[" & UserInfo.�û��� & "]" & strCurClassName
'    mrsClass!���ID = 0
'
'    mrsClass.Update
'
'    Set objSubNode = trvWordTree.Nodes.Add(objPNode, tvwChild, "T-" & lngCurClassId, strCurClassName, 2)
'
'    objSubNode.Selected = True
'    trvWordTree.StartLabelEdit
'Exit Sub
'errHandle:
'    If ErrCenter() = 1 Then Resume
'    Call SaveErrLog
'End Sub
 
Public Function WordNew() As Boolean
    Dim strErr As String
On Error GoTo errhandle
    Dim lngCurOutline As TOutlineType
    
    Dim str���� As String
    Dim str��� As String
    Dim str���� As String

'    str���� = "������������"
'    lngCurOutline = otDesc

    RaiseEvent OnRequestState(lngCurOutline, str����, str���, str����)

    Select Case lngCurOutline
        Case otDesc '����
            str��� = ""
            str���� = ""
        Case otOpin '���
            str���� = ""
            str���� = ""
        Case otAdvi '����
            str���� = ""
            str��� = ""
    End Select

    WordNew = WordInsert(str����, str���, str����)
Exit Function
errhandle:
    strErr = err.Description
    MsgboxH GetRootHwnd, strErr, vbOKOnly, "��ʾ"
End Function

Public Function FullSave() As Boolean
'ȫ�״���
    Dim strErr As String
On Error GoTo errhandle
    Dim lngCurOutline As Long
    Dim str���� As String
    Dim str��� As String
    Dim str���� As String
    
'    str���� = "������������"
'    str��� = "�����������"
'    str���� = "���Խ�������"
    
    RaiseEvent OnRequestState(lngCurOutline, str����, str���, str����)
    
    FullSave = WordInsert(str����, str���, str����)
Exit Function
errhandle:
    strErr = err.Description
    MsgboxH GetRootHwnd, strErr, vbOKOnly, "��ʾ"
End Function


Public Function WordInsert(ByVal str���� As String, ByVal str��� As String, ByVal str���� As String) As Boolean
'����ʾ�
    Dim objReportWordList As New frmReportWordList
    Dim strWordContext As String
    Dim objNode As Node
    Dim objWordNode As Node
    Dim lngClassID As Long
    Dim strClassName As String
    Dim lngNewWordId As Long
    Dim strNewWordName As String
    
    
    WordInsert = False
    strWordContext = ""
    
    If Trim(str����) <> "" Then
        strWordContext = strWordContext & "<<����>>" & str����
    End If
    
    If Trim(str���) <> "" Then
        If Trim(strWordContext) <> "" Then strWordContext = strWordContext & vbCrLf
        strWordContext = strWordContext & "<<���>>" & str���
    End If
    
    If Trim(str����) <> "" Then
        If Trim(strWordContext) <> "" Then strWordContext = strWordContext & vbCrLf
        strWordContext = strWordContext & "<<����>>" & str����
    End If
                    
    Set objNode = trvWordTree.SelectedItem
    
    If Left(objNode.Key, 1) = LVW_KEY_WORD Then Set objNode = objNode.Parent
    
    lngClassID = Split(objNode.Key, "-")(1)
    strClassName = objNode.Text
    
    Call objReportWordList.ZlShowMe(Me, strWordContext, mintWordPower, _
                                    lngClassID, strClassName, _
                                    mlngCurDeptId, lngNewWordId, strNewWordName)
    If lngNewWordId <= 0 Then Exit Function
    
    Set objWordNode = trvWordTree.Nodes.Add(objNode, tvwChild, "L-" & lngNewWordId, strNewWordName, 1)
    objWordNode.tag = -1 '��ʾû�н����������ж�
    
    mblnIsSyncWordFragment = True
    WordInsert = True
End Function

Public Function WordDelete() As Boolean
'ɾ��ѡ��ʾ�
'ɾ���ʾ�ʾ��
On Error GoTo errH
    Dim objWordNode As Node
    Dim lngWordID As Long
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    WordDelete = False
    
    Set objWordNode = trvWordTree.SelectedItem
    If Left(objWordNode.Key, 1) <> LVW_KEY_WORD Then Exit Function
    
    If MsgboxH(GetRootHwnd, "ȷ��Ҫɾ����ǰѡ��Ĵʾ���", vbYesNo + vbDefaultButton2, "��ʾ") = vbNo Then Exit Function
    
    lngWordID = Val(Split(objWordNode.Key, "-")(1))
    
    '����ʾ�Ĵ�����ID ���ǵ�ǰ�û�ID ������ɾ������ʾ�
    strSQL = " SELECT 1 FROM  �����ʾ�ʾ�� WHERE ID=[1] AND ��ԱID=[2] "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "�жϴʾ䴴����", lngWordID, UserInfo.ID)
    
    If rsTemp.RecordCount > 0 Then
        strSQL = "zl_�����ʾ�ʾ��_delete(" & lngWordID & ")"
        
        Call zlDatabase.ExecuteProcedure(strSQL, "ɾ���ʾ�")
    Else
        MsgboxH GetRootHwnd, "����ɾ���Ĵʾ䲻�ǵ�ǰ�û������ģ�������ɾ����", vbOKOnly, "��ʾ"
        Exit Function
    End If
    
    Call trvWordTree.Nodes.Remove(objWordNode.Index)
    
    mblnIsSyncWordFragment = True
    WordDelete = True
    
    Exit Function
errH:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Public Function WordModify() As Boolean
'�޸�ѡ��ʾ�
    Dim objReportWordList As New frmReportWordList
    Dim strWordContext As String
    Dim objPnode As Node
    Dim objWordNode As Node
    Dim lngClassID As Long
    Dim strClassName As String
    Dim lngWordID As Long
    Dim strWordName As String
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    WordModify = False
    
    Set objWordNode = trvWordTree.SelectedItem
    If Left(objWordNode.Key, 1) <> LVW_KEY_WORD Then Exit Function
    
    lngWordID = Val(Split(objWordNode.Key, "-")(1))
    strWordName = objWordNode.Text
    
    '����ʾ�Ĵ�����ID ���ǵ�ǰ�û�ID ������ɾ������ʾ�
    strSQL = " SELECT 1 FROM  �����ʾ�ʾ�� WHERE ID=[1] AND ��ԱID=[2] "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "�жϴʾ䴴����", lngWordID, UserInfo.ID)
    
    If rsTemp.RecordCount = 0 Then
        MsgboxH GetRootHwnd, "�����޸ĵĴʾ䲻�ǵ�ǰ�û������ģ��������޸ġ�", vbOKOnly, "��ʾ"
        Exit Function
    End If
         
    strWordContext = Split(objWordNode.tag & "__", "_")(2)
                    
    Set objPnode = objWordNode.Parent
    
    lngClassID = Val(Split(objPnode.Key, "-")(1))
    strClassName = objPnode.Text
    
    Call objReportWordList.ZlShowMe(Me, strWordContext, mintWordPower, _
                                    lngClassID, strClassName, _
                                    mlngCurDeptId, lngWordID, strWordName)
    If lngWordID <= 0 Then Exit Function
    
    objWordNode.tag = -1 '��ʾû�н����������ж�

    Call trvWordTree_NodeClick(objWordNode)
    
    mblnIsSyncWordFragment = True
    WordModify = True
    
End Function


Private Sub trvWordTree_DblClick()
    Dim i As Integer
    Dim objSelNode As Node
    Dim strErr As String
On Error GoTo errhandle
    Set objSelNode = trvWordTree.SelectedItem
    
    If Not objSelNode Is Nothing Then
        If Left(objSelNode.Key, 1) = LVW_KEY_WORD Then
 
            If mintWordDblClickMode = 1 Then
                '�ʾ�˫���󣬴򿪴ʾ�༭����
                WriteWordEdit Val(Split(objSelNode.Key & "-", "-")(1))
            Else
                Call WriteWordDirect
            End If
        End If
    End If
Exit Sub
errhandle:
    strErr = err.Description
    Call MsgboxH(GetRootHwnd, strErr, vbOKOnly, "��ʾ")
End Sub

Private Sub LoadWordData(Node As Node)
    If Left(Node.Key, 1) = LVW_KEY_NODE Then
        If Val(Node.tag) <> 1 Then
            '����ʾ���Ŀ
            Call LoadWordItem(Node)
            Node.tag = 1
            
            If mblnAutoRemove = False Then
                Call HideOutlineNode(mlngOutlineId)
            End If
        End If
    ElseIf Left(Node.Key, 1) = LVW_KEY_WORD Then
        '����ʾ�����
        Call LoadWordContext(Node)
    End If
End Sub

Private Sub LoadWordContext(Node As Node)
    Dim lngWordID As Long
    Dim str�����ı� As String
    Dim aryWordLines() As TWordLine
    Dim aryPro() As String
    Dim blnIsApply As Boolean
    
On Error GoTo errhandle
    '���ԭ�пؼ�
    vsWordContext.Rows = 0
    
    Call LevalEdit(False)
    
    If Left(Node.Key, 1) <> LVW_KEY_WORD Then Exit Sub
        
    ReDim aryWordLines(0)
    
    'Node.tag:0_1_�ı����� ��Ӧ˵�� ҽ��ID_����״̬_�ı�����
    aryPro = Split(Node.tag & "__", "_")
    lngWordID = Right(Node.Key, Len(Node.Key) - 2)
    
    If Val(aryPro(0)) < 0 Then
        '��ȡ�ʾ�����
        str�����ı� = GetWordContext(lngWordID)
    Else
        str�����ı� = aryPro(2)
    End If
    
    '�����ʾ�����
    Call FormatWords(lngWordID, str�����ı�, aryWordLines())
    
    '�жϴʾ���������
    If mblnIsWordValid Then
        If Val(aryPro(0)) <> mlngAdviceId Then
            '�����ݿ��жϴʾ��Ƿ����øü�黼�߱���
            blnIsApply = WordApplyState(lngWordID)
            
            If blnIsApply = False Then
                Node.BackColor = NODE_BACKCOLOR_DISABLE
                Node.ForeColor = NODE_FORCECOLOR_DISABLE
            Else
                Node.BackColor = vbWhite
                Node.ForeColor = vbBlack
            End If
        Else
            blnIsApply = IIf(Val(aryPro(1)) = 1, True, False)
        End If
        
        Node.tag = mlngAdviceId & "_" & IIf(blnIsApply, 1, 0) & "_" & str�����ı�
    Else
        blnIsApply = IIf(Node.Parent.BackColor = NODE_BACKCOLOR_DISABLE, False, True)
        
        Node.tag = "0_1_" & str�����ı�
    End If
    
    If blnIsApply = False Then
        '�����ô���
        vsWordContext.Rows = 1
        vsWordContext.Cell(flexcpText, 0, 1) = "ע:�ôʾ䲻���ô˱������..."
        vsWordContext.RowData(0) = "WARING"
        
        vsWordContext.BackColor = &HE0E0E0
        vsWordContext.BackColorBkg = &HE0E0E0
        
        vsWordContext.Cell(flexcpBackColor, 0, 0, 0, 1) = vbYellow
        vsWordContext.Cell(flexcpData, 0, 1) = 0
    Else
        vsWordContext.BackColor = vbWhite
        vsWordContext.BackColorBkg = vbWhite
    End If
    
    '��ʾ�ʾ�����
    Call ShowWordContext(aryWordLines(), blnIsApply)
    
    Exit Sub
errhandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function WordApplyState(ByVal lngWordID As Long) As Boolean
    '�жϴʾ�����״̬
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
On Error GoTo errhandle
    'mblnAutoRemove���Ϊtrue��˵�������õĴʾ��Ѿ����Զ��Ƴ�������Ҫ�����ж�
    If mblnAutoRemove Then
        WordApplyState = True
        Exit Function
    End If
    
    strSQL = "Select ���� " & _
                " From Table(Cast(f_Sentence_Usable([1], [2], [3], [4]) as zlhis.t_Dic_Rowset )) U " & _
                " Where ����=[5]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ѯ�ʾ�����״̬", mlngOutlineId, mlngPatientId, mlngPageID, mlngAdviceId, lngWordID)
    
    WordApplyState = IIf(rsTemp.RecordCount <= 0, False, True)
Exit Function
errhandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetWordContext(ByVal lngWordID As Long) As String
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim str�����ı� As String
    Dim strLineText As String
    
On Error GoTo errhandle
    GetWordContext = ""
    
    strSQL = "Select �ʾ�id,���д���,��������,�����ı�,����Ҫ��ID,�滻��,Ҫ������,Ҫ������,Ҫ�س���,Ҫ��С��," & _
             " Ҫ�ص�λ,Ҫ�ر�ʾ,Ҫ��ֵ��,������̬ From �����ʾ���� Where �ʾ�ID=[1] order by ���д��� "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ѯ�ʾ����", lngWordID)
    
    If rsTemp.RecordCount <= 0 Then Exit Function
    
    str�����ı� = ""
    '�����ݿ��ж�ȡ�ʾ�����з�������ʾ
    While rsTemp.EOF = False
        '�ȰѼ�¼�еĴʾ����ݶ�ȡ��str�����ı���
        strLineText = nvl(rsTemp!�����ı�)

        If rsTemp!�������� = 0 Then     '�������ı���ֱ�Ӽ�������
            If Trim(strLineText) <> "" Then  '�����ı���Ϊ�գ����������ʾ�����ı�
                str�����ı� = str�����ı� & strLineText
            End If
        Else        'rsTemp!��������<>0 ,��Ҫ�أ���Ҫ����
            Select Case Val(nvl(rsTemp!Ҫ�ر�ʾ))
                Case 0 ''�ı�Ҫ�ؽ����ɿա� ��
                    str�����ı� = str�����ı� & "  " & nvl(rsTemp!Ҫ�ص�λ)
                
                Case 1 '����
                'Ŀǰû��ʹ�������ʽ
                
                Case 2 '��ѡ
                    str�����ı� = str�����ı� & "{{" & nvl(rsTemp!Ҫ��ֵ��) & "}}" & nvl(rsTemp!Ҫ�ص�λ)
                
                Case 3 '��ѡ
                    str�����ı� = str�����ı� & "{<" & nvl(rsTemp!Ҫ��ֵ��) & ">}" & nvl(rsTemp!Ҫ�ص�λ)
            
            End Select
        End If
      
        rsTemp.MoveNext
    Wend
    
    GetWordContext = str�����ı�
Exit Function
errhandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetLayoutStr() As String
'���ظ�ʽ�ַ���[Key=picturebox1.width:20;picturebox1.height:30;]
    GetLayoutStr = "[KEY=SEGMENT@" & _
                                        GetProFmt("TRVWORDTREE.HEIGHT", trvWordTree.Height) & _
                                        GetProFmt("VSWORDCONTEXT.HEIGHT", vsWordContext.Height) & _
                                 "]"
End Function

Public Sub SetLayout(ByVal strLayout As String)
    Dim strPros As String
    Dim lngKeyIndex As String
    Dim strPro As String
    
    If Len(strLayout) <= 0 Then Exit Sub
    
    strPros = GetPros(strLayout, "SEGMENT")
    
    strPro = GetProValue(strPros, "TRVWORDTREE.HEIGHT")
    If Val(strPro) > 0 Then trvWordTree.Height = Val(strPro)
    
    strPro = GetProValue(strPro, "VSWORDCONTEXT.HEIGHT")
    If Val(strPro) > 0 Then vsWordContext.Height = Val(strPro)
    

End Sub


Private Sub ShowWordContext(aryWordLines() As TWordLine, ByVal blnIsApply As Boolean)
    Dim i As Long
    Dim lngBaseRow As Long
    Dim lngFillRow As Long
    Dim strWordOutline As String
    
    lngBaseRow = vsWordContext.Rows
    
    vsWordContext.Rows = lngBaseRow + UBound(aryWordLines)
    For i = 1 To UBound(aryWordLines)
        If Trim(aryWordLines(i).strContext) <> "" Then
            lngFillRow = lngBaseRow + (i - 1)
            strWordOutline = aryWordLines(i).strOutlineName
            
            vsWordContext.Cell(flexcpText, lngFillRow, 1) = aryWordLines(i).strContext
            vsWordContext.RowData(lngFillRow) = aryWordLines(i).strOutlineName
            
            If Len(strWordOutline) > 0 Then
                If InStr(strWordOutline, "����") >= 1 Then
                    Set vsWordContext.Cell(flexcpPicture, lngFillRow, 0) = imgDesc.Picture
                ElseIf InStr(strWordOutline, "���") >= 1 Or InStr(strWordOutline, "���") >= 1 Or (InStr(strWordOutline, "���") >= 1 And InStr(strWordOutline, "����") <= 0) Then
                    'ƥ�����,���,��Ͻ�����ų���Ͻ�����ش�
                    Set vsWordContext.Cell(flexcpPicture, lngFillRow, 0) = imgOpin.Picture
                ElseIf InStr(strWordOutline, "����") >= 1 Then
                    Set vsWordContext.Cell(flexcpPicture, lngFillRow, 0) = imgAdvi.Picture
                Else
                    Set vsWordContext.Cell(flexcpPicture, lngFillRow, 0) = Image1.Picture
                End If
            Else
                Set vsWordContext.Cell(flexcpPicture, lngFillRow, 0) = Image1.Picture
            End If
            
            If blnIsApply Then
                vsWordContext.Cell(flexcpData, lngFillRow, 1) = 1
            End If
        End If
    Next
    
    vsWordContext.ColWidth(0) = 450
    If lngBaseRow <> 0 Then
        vsWordContext.Cell(flexcpAlignment, lngBaseRow, 1, vsWordContext.Rows - 1, 1) = flexAlignLeftTop
    Else
        vsWordContext.ColAlignment(1) = flexAlignLeftTop
    End If
    
    Call vsWordContext.AutoSize(0, 1)
End Sub

Private Sub trvWordTree_Expand(ByVal Node As MSComctlLib.Node)
    Dim strErr As String
On Error GoTo errhandle
'    If Node.BackColor = NODE_BACKCOLOR_DISABLE Then
'        Node.Expanded = False
'    End If
    
    Call LoadWordData(Node)
Exit Sub
errhandle:
    strErr = err.Description
    MsgboxH GetRootHwnd, strErr, vbOKOnly, "��ʾ"
End Sub

Private Sub trvWordTree_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strErr As String
On Error GoTo errhandle
    '�����Ҽ������˵����ж��Ƿ��Ҽ�
    RaiseEvent OnMouseUp(Button, Shift, X, Y)
Exit Sub
errhandle:
    strErr = err.Description
    MsgboxH GetRootHwnd, strErr, vbOKOnly, "��ʾ"
End Sub

Private Sub trvWordTree_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim strErr As String
On Error GoTo errhandle
    Call LoadWordData(Node)
Exit Sub
errhandle:
    strErr = err.Description
    MsgboxH GetRootHwnd, strErr, vbOKOnly, "��ʾ"
End Sub


Private Sub TrvwClear()
     Dim X As Integer
     
     With trvWordTree
        SendMessage .hwnd, WM_SETREDRAW, 0, 0
        
        For X = .Nodes.Count To 1 Step -1
            .Nodes.Remove X
        Next X
        
        SendMessage .hwnd, WM_SETREDRAW, 1, 0
     End With
End Sub
 
 
Public Sub SyncOutline(ByVal strOutlineKey As String)
'ͬ�����
    Dim i As Long
    Dim strWordOutline As String
    
    
    
    If mstrOutLineKey = strOutlineKey Then Exit Sub
    
    txtWordEdit.Visible = False
'    mstrOutLineKey = strOutlineKey
        
'    If mblnAutoRemove Then
        Call LoadWordClass(mlngFileID, strOutlineKey, False)
'    End If
    
    mstrOutLineKey = strOutlineKey
    
'    If vsWordContext.Rows <= 0 Then Exit Sub
'    If Len(mstrOutLineName) <= 0 Then Exit Sub
'
'    For i = 1 To vsWordContext.Rows - 1
'        strWordOutline = vsWordContext.RowData(i)
'
'        If Len(strWordOutline) > 0 Then
'            If InStr(strWordOutline, mstrOutLineName) <= 0 Then
'                '��ƥ�䵱ǰѡ�����
'                vsWordContext.Cell(flexcpBackColor, 0, 0, 0, 1) = &HE0E0E0
'            Else
'                vsWordContext.Cell(flexcpBackColor, 0, 0, 0, 1) = vbWhite
'            End If
'        End If
'    Next
    
End Sub

Public Sub Refresh(ByVal lngAdviceId As Long, ByVal lngFileId As Long, _
    Optional ByVal strOutlineName As String = "����", _
    Optional blnForceRefresh As Boolean)

    mblnIsSyncWordFragment = False
    
    If lngAdviceId <> mlngAdviceId Then
        Call InitPatientInfo(lngAdviceId)
    End If
    
    mlngAdviceId = lngAdviceId
    mlngFileID = lngFileId
    
    If lngFileId <= 0 Then
        trvWordTree.Nodes.Clear
        vsWordContext.Rows = 0
        txtWordEdit.Text = ""
        Exit Sub
    End If

    Call LoadWordClass(lngFileId, strOutlineName, blnForceRefresh)
     
End Sub

Private Sub InitDbOwner(ByVal lngSys As Long)
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL  As String
On Error GoTo errHand
    If mstrDBOwner <> "" Then Exit Sub

    strSQL = "Select ������ From Zlsystems Where ��� = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ���ݿ�������", lngSys)
    
    If rsTemp.RecordCount <> 0 Then mstrDBOwner = "" & rsTemp!������
    rsTemp.Close
    
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub InitLoaclParas()
'    Dim strSQL As String
'    Dim rsTemp As ADODB.Recordset
'
'    On Error GoTo err
'
     
    
'    mintWordDblClickMode = Val(GetDeptPara(mlngCurDeptId, "����ʾ�˫������", 0))
''
'
'    mlngWordTreeH = GetSetting("ZLSOFT", strRegPath, "WordTreeH", 200)
'    mlngWordShowH = GetSetting("ZLSOFT", strRegPath, "WordShowH", 300) - 15
'    mlngPrivateWordH = GetSetting("ZLSOFT", strRegPath, "PrivateWordH", 200) + 355
'    mlngButtonH = GetSetting("ZLSOFT", strRegPath, "ButtonH", 500) + 325
''    chkֱ�ӱ༭.value = IIf(CBool(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\frmReportWord", "ֱ�ӱ༭", False)), 1, 0)
''    ChkAutoExpand.value = IIf(CBool(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\frmReportWord", "�Զ�չ��", False)), 1, 0)
''
'
'    Exit Sub
'err:
'    If ErrCenter() = 1 Then Resume Next
'    Call SaveErrLog
End Sub

Private Sub WriteWordEdit(lngWordID As Long)
    Dim intReportViewType As TOutlineType
    Dim str���� As String
    Dim str��� As String
    Dim str���� As String
    Dim objNode As Node
    Dim objWordEdit As New frmReportWordEdit
    
    '��ȡ��ǰ������������
    RaiseEvent OnRequestState(intReportViewType, str����, str���, str����)

    Set objNode = trvWordTree.SelectedItem
    If objNode Is Nothing Then Exit Sub

    objWordEdit.zlShowMeEx Me, mlngCurDeptId, lngWordID, Split(objNode.tag & "__", "_")(2), intReportViewType, str����, str���, str����

    RaiseEvent OnSendContext("", str����, str���, str����)
End Sub

Private Sub EnterEdit(ByVal lngRow As Long, ByVal lngCol As Long)
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngWidth As Long
    Dim lngHeight As Long
    Dim lngScrollWidth As Long
    Dim lngStartSel As Long
    Dim strSelContext As String

    txtWordEdit.Visible = False
    txtWordEdit.Text = ""
    
    vsWordContext.Row = lngRow
    vsWordContext.Col = lngCol
    
    vsWordContext.EditCell
    vsWordContext.EditSelStart = 0
    vsWordContext.EditSelLength = 0
    
    mouse_event MOUSEEVENTF_LEFTDOWN Or MOUSEEVENTF_LEFTUP, 0&, 0&, 0&, 0&
    
    DoEvents    'ֻ��ִ���˸þ��EditSelStart�Ż���Ӧ

    lngStartSel = vsWordContext.EditSelStart
    
    vsWordContext.EditSelStart = 0
    vsWordContext.EditSelLength = lngStartSel
    
    strSelContext = vsWordContext.EditSelText
    
    vsWordContext.EditSelLength = 0
    
    lngStartSel = Len(strSelContext)
    

    lngLeft = vsWordContext.ColPos(lngCol)
    lngTop = vsWordContext.RowPos(lngRow)
    lngWidth = vsWordContext.CellWidth
    lngHeight = vsWordContext.CellHeight
    
    txtWordEdit.Left = lngLeft
    txtWordEdit.Top = lngTop
    
    txtWordEdit.Width = lngWidth
    txtWordEdit.Height = lngHeight
    
    txtWordEdit.Text = vsWordContext.TextMatrix(lngRow, lngCol)
    Call SetWordStyle(txtWordEdit, vsWordContext.FontSize)
    
    txtWordEdit.Visible = True
    
    txtWordEdit.tag = -1 & "#" & lngRow & "#" & lngCol
    
    txtWordEdit.SetFocus
    txtWordEdit.SelStart = lngStartSel
    
'    mouse_event MOUSEEVENTF_LEFTDOWN Or MOUSEEVENTF_LEFTUP, 0&, 0&, 0&, 0&
End Sub

Private Sub txtWordEdit_DblClick()
    Dim strErr As String
On Error GoTo errhandle
    Call richTextBoxShowElements(txtWordEdit)
Exit Sub
errhandle:
    strErr = err.Description
    MsgboxH GetRootHwnd, strErr, vbOKOnly, "��ʾ"
End Sub

Private Sub UserControl_Initialize()
    mlngExpandLevel = 1
    mblnIsWordValid = True
    mblnAutoRemove = False
End Sub


Private Sub UserControl_Resize()
On Error Resume Next

    picBack.Move 0, 0, ScaleWidth, ScaleHeight
'    picBack.Left = 0
'    picBack.Top = 0
'    picBack.Width = UserControl.ScaleWidth
'    picBack.Height = UserControl.ScaleHeight

    Call ucSplitter1.RePaint(False)
End Sub

Public Sub Destory()
    ucSplitter1.Destory
    
    Set mrsClass = Nothing
    Set mrsWords = Nothing
End Sub

Private Sub UserControl_Terminate()
    Call Destory
End Sub

Private Sub vsWordContext_BeforeScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long, Cancel As Boolean)
    Dim strErr As String
On Error GoTo errhandle
    If vsWordContext.Rows <= 0 Then Exit Sub
    
    If txtWordEdit.Visible Then Call LevalEdit
Exit Sub
errhandle:
    strErr = err.Description
    MsgboxH GetRootHwnd, strErr, vbOKOnly, "��ʾ"
End Sub

Private Sub vsWordContext_Click()
    Dim strErr As String
On Error GoTo errhandle
    If vsWordContext.Rows <= 0 Then Exit Sub
    If vsWordContext.MouseRow < 0 Then
        Call LevalEdit
        Exit Sub
    End If
    
    If vsWordContext.Row < 0 Then Exit Sub
    
    If vsWordContext.Row <> vsWordContext.MouseRow Then vsWordContext.Row = vsWordContext.MouseRow
    
    If vsWordContext.Col > 0 Then
        '����ʾ�༭
        If txtWordEdit.Visible Then
            Call LevalEdit
        End If
 
        If vsWordContext.Cell(flexcpBackColor, vsWordContext.Row) = vbYellow Then Exit Sub
        Call EnterEdit(vsWordContext.Row, vsWordContext.Col)
   
        Exit Sub
    ElseIf vsWordContext.Col = 0 Then
        'д��ѡ��Ĵʾ䵽����
        If txtWordEdit.Visible Then
            Call LevalEdit
        End If
        
        Call DoWritWord(vsWordContext.Row)
    End If
Exit Sub
errhandle:
    strErr = err.Description
    MsgboxH GetRootHwnd, strErr, vbOKOnly, "��ʾ"
End Sub

Private Sub LevalEdit(Optional ByVal blnUpdateEdit As Boolean = True)
    Dim aryEditPro() As String
    
    If txtWordEdit.Visible = False Then Exit Sub
    
    If Val(txtWordEdit.tag) = -1 And blnUpdateEdit Then
        aryEditPro = Split(txtWordEdit.tag, "#")
        vsWordContext.Cell(flexcpText, Val(aryEditPro(1)), Val(aryEditPro(2))) = txtWordEdit.Text
        
        Call vsWordContext.AutoSize(0, 1)
    End If
    
    txtWordEdit.tag = ""
    txtWordEdit.Visible = False
End Sub


Private Sub vsWordContext_DblClick()
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    Dim strApplys As String
    
    If vsWordContext.Row <> 0 Then Exit Sub
    
    If vsWordContext.RowData(0) = "WARING" Then
        '��ȡ������������
        strSQL = "select �ʾ�ID,������,����ֵ from �����ʾ����� Where �ʾ�ID=[1]"
        Set rsData = zlDatabase.OpenSQLRecord(strSQL, "��ѯ�ʾ���������", Val(Split(trvWordTree.SelectedItem.Key, "-")(1)))
        
        If rsData.RecordCount <= 0 Then
            MsgboxH GetRootHwnd, "��ٴʾ�δ����������", vbOKOnly, "��������"
            Exit Sub
        End If
        
        strApplys = ""
        While Not rsData.EOF
            strApplys = nvl(rsData!������) & ":" & nvl(rsData!����ֵ) & vbCrLf & strApplys
            rsData.MoveNext
        Wend
        
        MsgboxH GetRootHwnd, strApplys & vbCrLf & "��������ԭ����ٴʾ�δ��������", vbOKOnly, "��������"
    End If
End Sub

Private Sub vsWordContext_KeyPress(KeyAscii As Integer)
    Dim strErr As String
On Error GoTo errhandle
    
    If KeyAscii <> 13 Then Exit Sub
    
    If vsWordContext.Rows <= 0 Then Exit Sub
    If vsWordContext.Col <> 0 Then Exit Sub
    If vsWordContext.Row < 0 Then Exit Sub
    
    Call DoWritWord(vsWordContext.Row)
Exit Sub
errhandle:
    strErr = err.Description
    MsgboxH GetRootHwnd, strErr, vbOKOnly, "��ʾ"
End Sub

Private Sub DoWritWord(ByVal lngRow As Long, Optional ByVal blnApplyHint As Boolean = True)
    Dim strOutline As String
    Dim str���� As String
    Dim str��� As String
    Dim str���� As String
    Dim strFree As String
    
    If vsWordContext.RowData(lngRow) = "WARING" Then Exit Sub
    
    If Val(vsWordContext.Cell(flexcpData, lngRow, 1)) <> 1 And blnApplyHint Then
        If MsgboxH(GetRootHwnd, "�ôʾ䲻�����ڵ�ǰ��٣��Ƿ������", vbYesNo + vbDefaultButton2, "��ʾ") = vbNo Then Exit Sub
    End If
    
    strOutline = vsWordContext.RowData(lngRow)
    
    If strOutline = "" Then
        strFree = vsWordContext.Cell(flexcpText, lngRow, 1)
    Else
        Select Case strOutline
            Case "<<����>>"
                str���� = vsWordContext.Cell(flexcpText, lngRow, 1)
            Case "<<���>>"
                str��� = vsWordContext.Cell(flexcpText, lngRow, 1)
            Case "<<����>>"
                str���� = vsWordContext.Cell(flexcpText, lngRow, 1)
        End Select
    End If
    
    RaiseEvent OnSendContext(strFree, str����, str���, str����)
End Sub

Private Sub vsWordContext_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo errhandle
    If Button = 2 Then Exit Sub
    If vsWordContext.Rows <= 0 Then Exit Sub
    
    If vsWordContext.MouseCol <> 0 Then Exit Sub
    If vsWordContext.MouseRow < 0 Then Exit Sub
    
    If vsWordContext.Cell(flexcpPicture, vsWordContext.MouseRow, 0) Is Nothing Then Exit Sub
    If Val(vsWordContext.Cell(flexcpData, vsWordContext.MouseRow, 1)) <> 1 Then Exit Sub
    
    vsWordContext.Cell(flexcpBackColor, vsWordContext.MouseRow, 0, vsWordContext.MouseRow, 1) = &HC0FFFF
Exit Sub
errhandle:

End Sub

Private Sub vsWordContext_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo errhandle
    If Button = 2 Then Exit Sub
    If vsWordContext.Rows <= 0 Then Exit Sub
    
    If vsWordContext.Col <> 0 Then Exit Sub
    If vsWordContext.Row < 0 Then Exit Sub
    
    If vsWordContext.Cell(flexcpPicture, vsWordContext.Row, 0) Is Nothing Then Exit Sub
    If Val(vsWordContext.Cell(flexcpData, vsWordContext.Row, 1)) <> 1 Then Exit Sub
    
    vsWordContext.Cell(flexcpBackColor, vsWordContext.Row, 0, vsWordContext.Row, 1) = vbWhite
Exit Sub
errhandle:

End Sub
