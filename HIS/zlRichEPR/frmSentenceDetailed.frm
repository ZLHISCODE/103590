VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Begin VB.Form frmSentenceDetailed 
   BorderStyle     =   0  'None
   Caption         =   "�ʾ����"
   ClientHeight    =   7290
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6345
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7290
   ScaleWidth      =   6345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox piclist 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1935
      Left            =   600
      ScaleHeight     =   1935
      ScaleWidth      =   2655
      TabIndex        =   4
      Top             =   4680
      Width           =   2655
      Begin VSFlex8Ctl.VSFlexGrid vsfList 
         Height          =   1695
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   1695
         _cx             =   2990
         _cy             =   2990
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
         BackColorBkg    =   16777215
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
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
   Begin VB.PictureBox picfind 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   3240
      ScaleHeight     =   1335
      ScaleWidth      =   2655
      TabIndex        =   0
      Top             =   1080
      Width           =   2655
      Begin VB.TextBox txtFind 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   840
         TabIndex        =   6
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label lblFilter 
         BackColor       =   &H00FFFFFF&
         Caption         =   "���ˣ�"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   615
      End
      Begin VB.Shape shpFind 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00400040&
         Height          =   255
         Left            =   0
         Top             =   120
         Width           =   1215
      End
   End
   Begin MSComctlLib.ImageList imgClass 
      Left            =   1800
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
            Picture         =   "frmSentenceDetailed.frx":0000
            Key             =   "close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSentenceDetailed.frx":059A
            Key             =   "expend"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView TreeList 
      Height          =   1470
      Left            =   240
      TabIndex        =   2
      Top             =   720
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
   Begin MSComctlLib.ImageList imgList 
      Left            =   720
      Top             =   2280
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
            Picture         =   "frmSentenceDetailed.frx":0B34
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSentenceDetailed.frx":10CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSentenceDetailed.frx":1668
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSentenceDetailed.frx":1C02
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfFind 
      Height          =   615
      Left            =   600
      TabIndex        =   3
      Top             =   3720
      Visible         =   0   'False
      Width           =   2175
      _cx             =   3836
      _cy             =   1085
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
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   16777215
      BackColorAlternate=   16777215
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
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   360
      Top             =   240
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Bindings        =   "frmSentenceDetailed.frx":24DC
      Left            =   840
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmSentenceDetailed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'######################################################################################################################

'��������
'------------------------------------------------------------------------------------------------------------------
Private Enum mCol
    ID = 0: sortid: Sort: Num: pName: Range: Depart: personnel: Pinyin: Wubi
End Enum


Private Const conPane_Tree = 400
Private Const conPane_List = 401
Private Const conPane_Text = 404
Private cbrControl As CommandBarControl
Private cbrMenuBar As CommandBarPopup
Private cbrToolBar As CommandBar

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
Private mfrmTipInfo As New frmTipInfo
Private mlngId As Long
Private mrsTmp As ADODB.Recordset
Private mintPrompt As Integer '�ж��Ƿ�ˢ����ʾ����Ϣ
Private mstrContent As String
Private mintPower As Integer
Private mLeftRight As Integer
Public Event RowDblClick(ByVal lngSentenceID As Long)    '˫��һ�л������ϰ��س�
Public Event ShiftFocus()           '�ı佹��


'����Ϊ�ⲿ��������
'######################################################################################################################

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
    mlngClassId = lngPatient
    Set mfrmParent = frmParent
    If blnForce = False And mlngParentId = lngCompendID And _
        (mstrSecondLimit = strSecondLimit) Then zlRefFromCompend = Me.vsfList.Rows: Exit Function
    mblnCompend = True
    TreeList.Visible = mblnCompend
    mlngParentId = lngCompendID
    mlngPatient = lngPatient
    mlngVisit = lngVisit
    mlngAdvice = lngAdvice
    mstrSecondLimit = strSecondLimit
    
    Set panThis = dkpMan.FindPane(conPane_Tree)
    panThis.Closed = False
    
    gstrSQL = "Select �ʾ����id From ������ٴʾ� Where ���id = [1]"
    Err = 0: On Error GoTo errHand
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "frmSentenceDetailed", mlngParentId)
  
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



'����Ϊ�ڲ���������
'######################################################################################################################

Public Function zlSubRefList(Optional lngID As Long, Optional ByVal lng����id As Long) As Long
    '******************************************************************************************************************
    '���ܣ�ˢ��װ���嵥������λ��ָ���ļ�¼��
    '������
    '���أ�
    '******************************************************************************************************************
Dim rsTemp As New ADODB.Recordset
Dim strClassIds As String, strKinds As String, strText As String, blnAdd As Boolean
Dim i As Integer


    
    '------------------------------------------------------------------------------------------------------------------
        '�������ʾ�����ڲ����༭������
        gstrSQL = "Select /*+ rule*/ L.ID, L.����id, C.���� || '-' || C.���� As ����, L.���, L.����, L.ͨ�ü� as ��Χ, D.���� As ����, P.���� As ��Ա,zlspellcode(L.����) as ƴ��,zlwbcode(L.����) as ���" & vbNewLine & _
                "From �����ʾ���� C, �����ʾ�ʾ�� L, ������ٴʾ� A, ���ű� D, ��Ա�� P," & vbNewLine & _
                "     Table(Cast(f_Sentence_Usable([1], [2], [3], [4]) As " & gstrDbOwner & ".t_Dic_Rowset)) U" & vbNewLine & _
                "Where C.ID = L.����id And L.����id = A.�ʾ����id And L.����id = D.ID And L.��Աid = P.ID And A.���id = [1] And" & vbNewLine & _
                "      L.ID = To_Number(U.����)  "
        If lng����id > 0 Then gstrSQL = gstrSQL & "  And L.����id=[5] "

    '------------------------------------------------------------------------------------------------------------------
    If InStr(1, gstrPrivsEpr, "���Ҳ����ʾ�") <> 0 Then
        gstrSQL = gstrSQL & " And" & vbNewLine & _
                "      (Nvl(L.ͨ�ü�, 0) = 0 Or" & vbNewLine & _
                "      L.ͨ�ü� In (1, 2) And" & vbNewLine & _
                "      L.����id In (Select R.����id From ������Ա R, �ϻ���Ա�� U Where R.��Աid = U.��Աid And U.�û��� = User))"

     Else
        gstrSQL = gstrSQL & " And" & vbNewLine & _
                "      (Nvl(L.ͨ�ü�, 0) = 0 Or" & vbNewLine & _
                "      L.ͨ�ü� = 1 And" & vbNewLine & _
                "      L.����id In (Select R.����id From ������Ա R, �ϻ���Ա�� U Where R.��Աid = U.��Աid And U.�û��� = User) Or" & vbNewLine & _
                "      L.ͨ�ü� = 2 And L.��Աid In (Select U.��Աid From �ϻ���Ա�� U Where U.�û��� = User))"
    End If
    
    Err = 0: On Error GoTo errHand
    gstrSQL = gstrSQL & "Order by L.ͨ�ü� Desc, Lpad(L.���,13,'0')"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "frmSentenceDetailed", mlngParentId, mlngPatient, mlngVisit, mlngAdvice, lng����id)
     Set mrsTmp = rsTemp
    '------------------------------------------------------------------------------------------------------------------

    strClassIds = ","

    If Not rsTemp.EOF Then
        With Me.vsfList
            Set .DataSource = rsTemp
            .ColWidth(mCol.ID) = 0
            .ColWidth(mCol.sortid) = 0
            .ColWidth(mCol.Sort) = 0
            .ColWidth(mCol.Num) = Me.picList.Width / 6 + 200
            .ColWidth(mCol.pName) = Me.picList.Width / 3 * 2 - 200
            .ColWidth(mCol.Range) = Me.picList.Width / 6 - 50
            .ColWidth(mCol.Depart) = 0
            .ColWidth(mCol.personnel) = 0
            .ColWidth(mCol.Pinyin) = 0
            .ColWidth(mCol.Wubi) = 0
            For i = 1 To .Rows - 1
                Select Case .TextMatrix(i, mCol.Range)
                    Case 0:
                        .TextMatrix(i, 5) = ""
                        .Cell(flexcpPicture, i, mCol.Range) = Me.imgList.ListImages(1).Picture '"1-ȫԺ"
                    Case 1:
                        .TextMatrix(i, 5) = ""
                        .Cell(flexcpPicture, i, mCol.Range) = Me.imgList.ListImages(2).Picture '"2-����"
                    Case 2:
                        .TextMatrix(i, 5) = ""
                        .Cell(flexcpPicture, i, mCol.Range) = Me.imgList.ListImages(3).Picture '"3-����"
                    Case Else:
                        .TextMatrix(i, 5) = ""
                        .Cell(flexcpPicture, i, mCol.Range) = Me.imgList.ListImages(1).Picture '"1-ȫԺ"
                End Select
            Next
        End With
    Else
        Me.vsfList.Rows = 1
    End If

    If mlngId <> mlngParentId Or rsTemp.RecordCount = 0 Then
        Me.vsfFind.Visible = False
        Me.txtFind.Text = ""
        mlngId = mlngParentId
    End If
    If vsfList.Rows > 1 Then
        vsfList.Row = 1
    End If
    Exit Function
    
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlSubRefList = Me.vsfList.Rows
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
            
    If InStr(1, gstrPrivsEpr, "���Ҳ����ʾ�") <> 0 Then
        strSQL = strSQL & " And" & vbNewLine & _
                "      (Nvl(L.ͨ�ü�, 0) = 0 Or" & vbNewLine & _
                "      L.ͨ�ü� In (1, 2) And" & vbNewLine & _
                "      L.����id In (Select R.����id From ������Ա R, �ϻ���Ա�� U Where R.��Աid = U.��Աid And U.�û��� = User))"

    Else
        strSQL = strSQL & " And" & vbNewLine & _
                "      (Nvl(L.ͨ�ü�, 0) = 0 Or" & vbNewLine & _
                "      L.ͨ�ü� = 1 And" & vbNewLine & _
                "      L.����id In (Select R.����id From ������Ա R, �ϻ���Ա�� U Where R.��Աid = U.��Աid And U.�û��� = User) Or" & vbNewLine & _
                "      L.ͨ�ü� = 2 And L.��Աid In (Select U.��Աid From �ϻ���Ա�� U Where U.�û��� = User))"
    End If
    
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = gstrSQL & strSQL
    gstrSQL = gstrSQL & ") Connect By Prior �ϼ�id=Id  Order By ����"
    
    Dim objNode As node
    
    TreeList.Nodes.Clear
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "frmSentenceDetailed", mlngParentId, mlngPatient, mlngVisit, mlngAdvice)
    If rsTemp.BOF = False Then
                
        Set objNode = TreeList.Nodes.Add(, , "K0", "���дʾ�", "close", "expend")
        objNode.Expanded = True
        Do While Not rsTemp.EOF
            
            Set objNode = Nothing
            
            On Error Resume Next
            Set objNode = TreeList.Nodes("K" & rsTemp("ID").Value)
            objNode.Expanded = True
            On Error GoTo errHand
            
            If objNode Is Nothing Then
                Set objNode = TreeList.Nodes.Add("K" & zlCommFun.NVL(rsTemp("�ϼ�id").Value, 0), tvwChild, "K" & rsTemp("ID").Value, rsTemp("����").Value, "close", "expend")
                objNode.Expanded = True
            End If
            rsTemp.MoveNext
        Loop
    End If
    If TreeList.Nodes.Count > 0 Then
        TreeList.Nodes(1).Selected = True
    Else
        mlngClassId = 0
    End If
    zlSubRefClass = True
    
    Exit Function
errHand:
    
End Function

'����Ϊ�ؼ��¼�����
'######################################################################################################################
Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    On Error Resume Next
    Select Case Item.ID
    Case conPane_Tree
        Item.Handle = TreeList.hwnd
    Case conPane_List
        Item.Handle = picList.hwnd
    Case conPane_Text
        Item.Handle = Me.picfind.hwnd
    End Select
End Sub


Private Sub Form_Load()
    '-----------------------------------------------------
    'Ȩ�����ƴ����ƣ�����ͬʱ��������ģ�������gmstrPrivs�仯�����¿�����Ч
Dim rptCol As ReportColumn
     mlngWordId = 0
     
     Call zlCommFun.SetWindowsInTaskBar(Me.hwnd, False)
    
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
     
     Set panThis = dkpMan.CreatePane(conPane_Text, 600, 50, DockTopOf, Nothing)
     panThis.Title = "��ݲ�ѯ"
     panThis.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
     panThis.MaxTrackSize.Height = 23
     panThis.MinTrackSize.Height = 23
     
     Set panThis = dkpMan.CreatePane(conPane_Tree, 600, 300, DockBottomOf, panThis)
     panThis.Title = "���ͽṹ"
     panThis.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
     
     Set panThis = dkpMan.CreatePane(conPane_List, 600, 450, DockBottomOf, panThis)
     panThis.Title = "�����б�"
     panThis.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
     panThis.Selected = False
    
    
     Me.dkpMan.Options.ThemedFloatingFrames = True
     Me.dkpMan.Options.HideClient = True
     dkpMan.LoadStateFromString GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & "frmSentenceDetailed" & "\" & TypeName(dkpMan), dkpMan.Name, "")
     '-----------------------------------------------------
    With Me.vsfFind
         .FixedCols = 0
         .SelectionMode = flexSelectionByRow
    End With
    '��vsflist���г�ʼ��
    With Me.vsfList
         .SelectionMode = flexSelectionByRow
         .FixedCols = 0
         .ExplorerBar = flexExSortShow
         .AddItem ""
         .Rows = 1
         .Cols = 10
         .TextMatrix(0, mCol.ID) = ""
         .TextMatrix(0, mCol.Num) = "���"
         .TextMatrix(0, mCol.pName) = "����"
         .TextMatrix(0, mCol.Range) = "��Χ"
    End With
    If InStr(1, gstrPrivsEpr, "ȫԺ�����ʾ�") <> 0 Then
        mintPower = 0
    ElseIf InStr(1, gstrPrivsEpr, "���Ҳ����ʾ�") <> 0 Then
        mintPower = 1
    ElseIf InStr(1, gstrPrivsEpr, "���˲����ʾ�") <> 0 Then
        mintPower = 2
    Else
        mintPower = -1
    End If
End Sub




Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Set cbrControl = Nothing
    Set cbrMenuBar = Nothing
    Set cbrToolBar = Nothing
    Set mfrmParent = Nothing
    Unload mfrmTipInfo
    Set mrsTmp = Nothing
    imgClass.ListImages.Clear
    imgList.ListImages.Clear
    ImageList_Destroy imgClass.hImageList
    ImageList_Destroy imgList.hImageList
    Call SaveSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & "frmSentenceDetailed" & "\" & TypeName(dkpMan), dkpMan.Name, dkpMan.SaveStateToString)
End Sub

Private Sub picfind_Resize()
    On Error Resume Next
    Me.picfind.BackColor = RGB(216, 231, 252)
    Me.lblFilter.BackColor = RGB(216, 231, 252)
    Me.lblFilter.Move 0, 80, Me.picfind.Width / 5, 220
    Me.txtFind.Move Me.lblFilter.Width + Screen.TwipsPerPixelX, 80, Me.picfind.Width / 5 * 4 - 2 * Screen.TwipsPerPixelX, 220
    Me.shpFind.Move Me.lblFilter.Width, 80 - Screen.TwipsPerPixelY, Me.txtFind.Width + 2 * Screen.TwipsPerPixelX, Me.txtFind.Height + 2 * Screen.TwipsPerPixelY
End Sub



Private Function getSentenceContent(lid As Long) As String
    Dim rsTemp As New ADODB.Recordset
    Dim strContent As String
    Dim lngStart As Long
        mlngWordId = lid
    If Me.Visible = False Then Exit Function

    'ˢ�´ʾ�����
    '------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo errHand
    gstrSQL = "Select ��������, �����ı�, Ҫ������, Ҫ�ص�λ From �����ʾ���� Where �ʾ�id = [1] Order By ���д���"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "frmSentenceDetailed", mlngWordId)
    With rsTemp
       Do While Not .EOF
            Select Case !��������
            Case 0 '��������
                strContent = strContent & IIf(IsNull(!�����ı�), " ", !�����ı�)
            Case 1, 2 '1-��ʱ����Ҫ��,2-�̶�����Ҫ��
                strContent = strContent & IIf(IsNull(!�����ı�), "{" & !Ҫ������ & "}" & !Ҫ�ص�λ, "{" & !�����ı� & "}")
            End Select
            .MoveNext
        Loop
    getSentenceContent = strContent
    End With
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Sub piclist_Resize()
    On Error Resume Next
    Me.vsfList.Move 0, 0, Me.picList.Width, Me.picList.Height
    With Me.vsfList
        .ColWidth(mCol.ID) = 0
        .ColWidth(mCol.sortid) = 0
        .ColWidth(mCol.Sort) = 0
        .ColWidth(mCol.Num) = Me.picList.Width / 6 + 200
        .ColWidth(mCol.pName) = Me.picList.Width / 3 * 2 - 200
        .ColWidth(mCol.Range) = Me.picList.Width / 6 - 50
        .ColWidth(mCol.Depart) = 0
        .ColWidth(mCol.personnel) = 0
        .ColWidth(mCol.Pinyin) = 0
        .ColWidth(mCol.Wubi) = 0
    End With
End Sub

Private Sub TreeList_NodeClick(ByVal node As MSComctlLib.node)
    If Val(Mid(node.Key, 2)) <> 0 Then mlngClassId = Val(Mid(node.Key, 2))
    Call zlSubRefList(mlngWordId, Val(Mid(node.Key, 2)))
End Sub

Private Sub txtFind_Change()
'�����ǰ�ʾ��б�û�����ݾͲ�����
 If Me.vsfList.Rows < 2 Then Exit Sub
    If Me.txtFind.Text <> "" Then
        mrsTmp.Filter = ""
        On Error GoTo aa
        mrsTmp.Filter = "���� like '*" & Me.txtFind.Text & "*' or ��� like '*" & Me.txtFind.Text & "*' or ƴ�� like '*" & Me.txtFind.Text & "*' or ��� like '*" & Me.txtFind.Text & "*'"
        If mrsTmp.RecordCount < 1 Then
            Me.vsfFind.Visible = False
            Exit Sub
        End If
        Set vsfFind.DataSource = mrsTmp
        Me.vsfFind.Move Me.txtFind.Left, Me.picfind.Height, Me.txtFind.Width, (mrsTmp.RecordCount + 2) * Me.vsfFind.ROWHEIGHT(1)
        Me.vsfFind.SheetBorder = RGB(216, 231, 252)
        Me.vsfFind.ZOrder 0
        Me.vsfFind.Visible = True
    Else
        Me.vsfFind.ZOrder 1
        Me.vsfFind.Visible = False
    End If
    With Me.vsfFind
       .ColWidth(0) = 0
       .ColWidth(1) = 0
       .ColWidth(2) = 0
       .ColWidth(3) = Me.vsfFind.Width / 3 - 50
       .ColWidth(4) = Me.vsfFind.Width / 3 * 2
       .ColWidth(5) = 0
       .ColWidth(6) = 0
       .ColWidth(7) = 0
       .ColWidth(8) = 0
       .ColWidth(9) = 0
    End With
    Exit Sub
aa:
    MsgBox "����������ݲ��Ϸ�����ֻ�������ַ��������Լ����ı����ţ�", vbInformation, gstrSysName
End Sub

Private Sub txtFind_GotFocus()
    RaiseEvent ShiftFocus
End Sub

Private Sub txtFind_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer

    '���˳���ť
    If KeyCode = vbKeyEscape Then
        Me.vsfFind.Visible = False
        Me.txtFind.Text = ""
        Exit Sub
    End If
    '�����»س���ʱ��
    If KeyCode = vbKeyReturn Then
        If mrsTmp Is Nothing Then Exit Sub
        '�����¼��û�����ݣ���ô���²�ѯ����
        If mrsTmp.RecordCount < 1 Then
            Call zlSubRefList(mlngWordId, 0)
            Call txtFind_Change
        Else
            With Me.vsfList
                If .Rows < 2 Then Exit Sub
                If .TextMatrix(.Row, 0) = "" Then Exit Sub
                'ѡ��ѡ�������
                For i = 1 To .Rows - 1
                    If .TextMatrix(i, 0) = Me.vsfFind.TextMatrix(Me.vsfFind.Row, 0) Then
                        .Row = i
                    End If
                Next
                '�Ѵʾ���ӵ��ļ���
                RaiseEvent RowDblClick(Me.vsfFind.TextMatrix(Me.vsfFind.Row, 0))
                Me.txtFind.Text = ""
            End With
            Me.vsfFind.Visible = False
        End If
    End If

    
    '�������¼���ʱ��ı���Ӧ��vsfѡ����
    If KeyCode = vbKeyDown And Me.vsfFind.Row < Me.vsfFind.Rows - 1 Then
        Me.vsfFind.Row = Me.vsfFind.Row + 1
    End If
    
    '�������ϼ���ʱ��ı���Ӧ��vsfѡ����
    If KeyCode = vbKeyUp And Me.vsfFind.Row > 1 Then
        Me.vsfFind.Row = Me.vsfFind.Row - 1
    End If
End Sub

Private Sub vsfFind_DblClick()
   Call txtFind_KeyDown(vbKeyReturn, -1)
End Sub
Private Sub vsfFind_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
         Me.picfind.Visible = False
    End If
    If KeyCode = vbKeyReturn Then
         Call vsfFind_DblClick
    End If
End Sub

Private Sub vsfList_DblClick()
Dim introw As Integer
    introw = Me.vsfList.MouseRow
    If introw < 1 Then Exit Sub
    mlngWordId = Val(Me.vsfList.TextMatrix(Me.vsfList.Row, mCol.ID))
   
    If mlngWordId = 0 Then Exit Sub
    RaiseEvent RowDblClick(mlngWordId)

End Sub

Private Sub vsfList_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    With Me.vsfList
        If .Rows < 2 Then Exit Sub
        Call vsfList_DblClick
    End With
End Sub


Private Sub vsfList_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim intpro As Integer
    intpro = y \ Me.vsfList.ROWHEIGHT(0)
    If intpro > vsfList.Rows - 1 Then
        Call mfrmTipInfo.ShowTipInfo(vsfList.hwnd, "", True)
        Exit Sub
    End If
    If Me.Width / 6 * 5 < x And x < Me.Width And intpro <> 0 Then
        
        '�����ͬһ�оͲ�����ˢ�´ʾ�
        If mintPrompt <> intpro Then
            mstrContent = getSentenceContent(Me.vsfList.TextMatrix(intpro, mCol.ID))
            mintPrompt = intpro
        End If
        Call mfrmTipInfo.ShowTipInfo(vsfList.hwnd, mstrContent, True)
    Else
        Call mfrmTipInfo.ShowTipInfo(vsfList.hwnd, "", True)
   End If
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
Dim lngRetuId As Long, strTemp As String
    
    Err = 0: On Error GoTo errHand
    '------------------------------------
    Select Case Control.ID
    Case conMenu_Edit_NewItem
        If Me.vsfList.Rows > 1 Then mlngClassId = Me.vsfList.TextMatrix(vsfList.Row, mCol.sortid)
        lngRetuId = frmSentenceEdit.ShowMe(mfrmParent, True, mintPower, mlngClassId)
        If lngRetuId = 0 Then Exit Sub
        Call zlSubRefList(lngRetuId)
    Case conMenu_Edit_Modify
        lngRetuId = frmSentenceEdit.ShowMe(mfrmParent, False, mintPower, Me.vsfList.TextMatrix(vsfList.Row, mCol.sortid), Me.vsfList.TextMatrix(vsfList.Row, mCol.ID))
        If lngRetuId = 0 Then Exit Sub
        Call zlSubRefList(lngRetuId)
    Case conMenu_Edit_Delete
        strTemp = "���ɾ���ôʾ���" & vbCrLf & "����" & Me.vsfList.TextMatrix(vsfList.Row, mCol.Num) & "-" & Me.vsfList.TextMatrix(vsfList.Row, mCol.pName)
        If MsgBox(strTemp, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        gstrSQL = "Zl_�����ʾ�ʾ��_Edit(3," & vsfList.TextMatrix(vsfList.Row, mCol.ID) & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "frmSentenceDetailed")
        Call zlSubRefList(mlngWordId)
    Case conMenu_Edit_Request
        If frmSentenceRequest.ShowMe(mfrmParent, Me.vsfList.TextMatrix(vsfList.Row, mCol.ID)) = True Then Call zlSubRefList(Me.vsfList.TextMatrix(vsfList.Row, mCol.ID))
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
    Dim lngEnable As Long
    Select Case Control.ID
    Case conMenu_Edit_NewItem
        Control.Visible = (mintPower >= 0 And mlngClassId <> 0)
    Case conMenu_Edit_Modify, conMenu_Edit_Delete, conMenu_Edit_Request
        Control.Visible = (mintPower >= 0 And mlngClassId <> 0)
        If vsfList.Rows > 1 Then
            Control.Enabled = (Me.vsfList.TextMatrix(vsfList.Row, mCol.ID) <> 0)
        Else
            Control.Enabled = False
        End If
        With Me.vsfList
              Select Case .Cell(flexcpPicture, vsfList.Row, mCol.Range)
                  Case Me.imgList.ListImages(1).Picture: '"1-ȫԺ"
                       lngEnable = 0
                  Case Me.imgList.ListImages(2).Picture: '"2-����"
                       lngEnable = 1
                  Case Me.imgList.ListImages(3).Picture: '"3-����"
                       lngEnable = 2
                  Case Else:
                      lngEnable = 0
              End Select
          End With
        If Control.Enabled Then Control.Enabled = (lngEnable >= mintPower)
        If mintPower = -1 Then Control.Enabled = False
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        Control.Enabled = (Me.vsfList.Rows > 1)
    End Select
End Sub
Private Sub zlRptPrint(ByVal bytMode As Byte)
    '******************************************************************************************************************
    '����:�����ݸ��Ƶ��ɴ�ӡ�Ķ��󣬵��ô�ӡ
    '����:  bytMode��1-��ӡ;2-Ԥ��;3-�����EXCEL
    '******************************************************************************************************************
    
    Dim objPrint As New zlPrint1Grd
    Dim objAppRow As zlTabAppRow
    
    If Me.vsfList.Rows < 1 Then Exit Sub

    
    '-------------------------------------------------
    '���ô�ӡ��������

    Set objPrint.Body = Me.vsfList
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

Private Sub vsfList_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim cbrPopupBar As CommandBar
    Dim cbrPopupItem As CommandBarControl

    If Button <> vbRightButton Then Exit Sub

    Set cbrMenuBar = Nothing
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
